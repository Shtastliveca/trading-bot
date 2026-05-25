"""
Hyperliquid Bridge (FIXED VERSION)
===================================
This bridge connects your Bybit trading bot to Hyperliquid exchange.
It provides data to the Bybit bot and executes trades on Hyperliquid.

Fixed issues:
- Added get_server_time() for connectivity checks
- Added get_api_key_information()
- Added get_executions()
- Added get_closed_pnl()

Usage:
    from hyperliquid_bridge import HyperliquidBridge
    bridge = HyperliquidBridge()
"""

import json
import logging
import time
import math
import threading
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Any, Tuple
from decimal import Decimal, ROUND_DOWN
import configparser
import os

# Hyperliquid SDK
from hyperliquid.info import Info
from hyperliquid.exchange import Exchange
from hyperliquid.utils import constants
import eth_account
from eth_account.signers.local import LocalAccount

from dotenv import load_dotenv
load_dotenv()

# Set up logging
logger = logging.getLogger("hyperliquid_bridge")

class HyperliquidBridge:
    """
    Bridge between Bybit trading bot and Hyperliquid exchange.
    """
    
    def __init__(self, config_path: str = 'config.ini'):
        """Initialize the Hyperliquid bridge"""
        self.config = configparser.ConfigParser()
        self.config_path = config_path
        self._load_config()
        
        # Connection instances
        self._info: Optional[Info] = None
        self._exchange: Optional[Exchange] = None
        self._account: Optional[LocalAccount] = None
        
        # Cache for symbol info
        self._symbol_info_cache: Dict[str, Dict] = {}
        self._symbol_info_cache_time: Dict[str, datetime] = {}
        self._cache_duration = timedelta(minutes=5)
        
        # Symbol mapping: Bybit format -> Hyperliquid format
        self._symbol_map: Dict[str, str] = {}
        self._reverse_symbol_map: Dict[str, str] = {}
        
        # Initialize connection
        self._initialize_connection()
        
        # Build symbol mapping
        self._build_symbol_mapping()
        
        logger.info("Hyperliquid Bridge initialized successfully")
    
    def _normalize_order_id(self, order_id) -> str:
        """Ensure order ID is consistently a string for comparison."""
        if order_id is None:
            return ''
        return str(order_id).strip()
    
    def _load_config(self):
        """Load configuration from file"""
        if os.path.exists(self.config_path):
            self.config.read(self.config_path)
        
        if 'HYPERLIQUID' not in self.config:
            self.config['HYPERLIQUID'] = {
                'private_key': '',
                'wallet_address': '',
                'testnet': 'False',
                'trading_fee_percent': '0.035',
                'maker_fee_percent': '0.01',
            }
            with open(self.config_path, 'w') as f:
                self.config.write(f)
    
    def _initialize_connection(self):
        """Initialize connection to Hyperliquid"""
        try:
            # Get API secret / private key
            private_key = self.config.get('HYPERLIQUID', 'private_key', fallback='')
            if not private_key:
                private_key = os.getenv('HYPERLIQUID_PRIVATE_KEY', '')
            
            if not private_key:
                raise ValueError("Hyperliquid private key not configured")
            
            if not private_key.startswith('0x'):
                private_key = '0x' + private_key
            
            # Create account from private key (for signing)
            self._account = eth_account.Account.from_key(private_key)
            
            # IMPORTANT: Get the actual wallet address
            # If using API keys, wallet_address must be set to your ACTUAL wallet
            # (not the API key ID which is derived from the API secret)
            wallet_address = self.config.get('HYPERLIQUID', 'wallet_address', fallback='')
            if not wallet_address:
                wallet_address = os.getenv('HYPERLIQUID_WALLET_ADDRESS', '')
            
            # If wallet_address is provided, use it; otherwise derive from private key
            if wallet_address:
                if not wallet_address.startswith('0x'):
                    wallet_address = '0x' + wallet_address
                self._wallet_address = wallet_address
                logger.info(f"Using specified wallet address: {self._wallet_address}")
                logger.info(f"API key address (for signing): {self._account.address}")
            else:
                # No separate wallet address - assume using direct wallet private key
                self._wallet_address = self._account.address
                logger.info(f"Using derived wallet address: {self._wallet_address}")
            
            use_testnet = self.config.getboolean('HYPERLIQUID', 'testnet', fallback=False)
            base_url = constants.TESTNET_API_URL if use_testnet else constants.MAINNET_API_URL
            
            self._info = Info(base_url, skip_ws=True)
            
            # For Exchange, we need to specify the vault/account address if using API keys
            self._exchange = Exchange(
                self._account,
                base_url,
                account_address=self._wallet_address  # Use actual wallet for trading
            )
            
            logger.info(f"Connected to Hyperliquid {'Testnet' if use_testnet else 'Mainnet'}")
            
        except Exception as e:
            logger.error(f"Failed to initialize Hyperliquid connection: {e}")
            raise
    
    def _build_symbol_mapping(self):
        """Build mapping between Bybit symbols and Hyperliquid symbols"""
        try:
            meta = self._info.meta()
            universe = meta.get('universe', [])
            
            for asset in universe:
                hl_symbol = asset.get('name', '')
                bybit_symbol = f"{hl_symbol}USDT"
                self._symbol_map[bybit_symbol] = hl_symbol
                self._reverse_symbol_map[hl_symbol] = bybit_symbol
            
            logger.info(f"Built symbol mapping for {len(self._symbol_map)} assets")
            
        except Exception as e:
            logger.error(f"Error building symbol mapping: {e}")
            defaults = ['BTC', 'ETH', 'SOL', 'DOGE', 'WIF', 'PEPE', 'SUI', 
                       'AVAX', 'LINK', 'ARB', 'OP', 'INJ', 'MATIC', 'APT', 'NEAR']
            for sym in defaults:
                self._symbol_map[f"{sym}USDT"] = sym
                self._reverse_symbol_map[sym] = f"{sym}USDT"
    
    def normalize_symbol(self, symbol: str) -> str:
        """Convert Bybit symbol format to Hyperliquid format"""
        symbol = symbol.upper()
        if symbol in self._reverse_symbol_map:
            return symbol
        if symbol in self._symbol_map:
            return self._symbol_map[symbol]
        for suffix in ['USDT', 'USD', 'PERP']:
            if symbol.endswith(suffix):
                return symbol[:-len(suffix)]
        return symbol
    
    def denormalize_symbol(self, symbol: str) -> str:
        """Convert Hyperliquid symbol format to Bybit format"""
        symbol = symbol.upper()
        if symbol in self._symbol_map:
            return symbol
        if symbol in self._reverse_symbol_map:
            return self._reverse_symbol_map[symbol]
        return f"{symbol}USDT"
    
    # ========== Server/Connection Methods ==========
    
    def get_server_time(self, **kwargs) -> Dict:
        """
        Get server time - Bybit-specific method.
        Hyperliquid doesn't have this, so we return current time.
        """
        try:
            current_time_ms = int(time.time() * 1000)
            return {
                'retCode': 0,
                'retMsg': 'OK',
                'result': {
                    'timeSecond': str(int(time.time())),
                    'timeNano': str(current_time_ms * 1000000)
                },
                'time': current_time_ms
            }
        except Exception as e:
            logger.error(f"Error in get_server_time: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def get_api_key_information(self, **kwargs) -> Dict:
        """Get API key info - returns wallet info for Hyperliquid"""
        try:
            return {
                'retCode': 0,
                'retMsg': 'OK',
                'result': {
                    'id': self._wallet_address,
                    'note': 'Hyperliquid Wallet',
                    'apiKey': self._wallet_address,
                    'signingKey': self._account.address,
                    'readOnly': 0,
                    'permissions': {'ContractTrade': ['Order', 'Position']}
                }
            }
        except Exception as e:
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    # ========== Wallet Balance Methods ==========
    
    def get_wallet_balance(self, **kwargs) -> Dict:
        """Get wallet balance in Bybit response format."""
        try:
            # Try perps user_state first (standard/legacy mode)
            user_state = self._info.user_state(self._wallet_address)
            margin_summary = user_state.get('marginSummary', {})
            account_value = float(margin_summary.get('accountValue', 0))

            # If zero, try spot clearinghouse (unified account mode)
            if account_value == 0:
                try:
                    spot_state = self._info.spot_user_state(self._wallet_address)
                    balances = spot_state.get('balances', [])
                    for b in balances:
                        if b.get('coin', '').upper() in ('USDC', 'USDH'):
                            account_value += float(b.get('total', 0))
                    logger.info(f"Using unified account balance: ${account_value}")
                except Exception as e:
                    logger.error(f"Error getting spot state: {e}")

            total_margin_used = float(margin_summary.get('totalMarginUsed', 0))
            available_balance = account_value - total_margin_used

            return {
                'retCode': 0,
                'retMsg': 'OK',
                'result': {
                    'list': [{
                        'totalEquity': str(account_value),
                        'availableBalance': str(available_balance),
                        'totalMarginBalance': str(account_value),
                        'totalWalletBalance': str(account_value),
                        'coin': [{
                            'coin': 'USDC',
                            'equity': str(account_value),
                            'availableToWithdraw': str(available_balance),
                            'walletBalance': str(account_value)
                        }]
                    }]
                }
            }
        except Exception as e:
            logger.error(f"Error getting wallet balance: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def get_balance(self) -> float:
        """Get account balance as a simple float"""
        try:
            result = self.get_wallet_balance()
            if result.get('retCode') == 0:
                return float(result['result']['list'][0]['totalEquity'])
            return 0.0
        except Exception as e:
            logger.error(f"Error getting balance: {e}")
            return 0.0
    
    # ========== Market Data Methods ==========
    
    def get_tickers(self, category: str = "linear", symbol: str = None) -> Dict:
        """Get ticker information in Bybit response format."""
        try:
            all_mids = self._info.all_mids()
            
            if symbol:
                hl_symbol = self.normalize_symbol(symbol)
                if hl_symbol in all_mids:
                    price = float(all_mids[hl_symbol])
                    return {
                        'retCode': 0,
                        'retMsg': 'OK',
                        'result': {
                            'list': [{
                                'symbol': symbol,
                                'lastPrice': str(price),
                                'markPrice': str(price),
                                'indexPrice': str(price),
                                'bid1Price': str(price * 0.9999),
                                'ask1Price': str(price * 1.0001),
                            }]
                        }
                    }
                else:
                    return {'retCode': -1, 'retMsg': f'Symbol {symbol} not found', 'result': {}}
            else:
                tickers = []
                for hl_sym, price in all_mids.items():
                    bybit_sym = self.denormalize_symbol(hl_sym)
                    tickers.append({
                        'symbol': bybit_sym,
                        'lastPrice': str(price),
                        'markPrice': str(price),
                    })
                return {'retCode': 0, 'retMsg': 'OK', 'result': {'list': tickers}}
        except Exception as e:
            logger.error(f"Error getting tickers: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def get_current_price(self, symbol: str) -> Optional[float]:
        """Get current price for a symbol"""
        try:
            hl_symbol = self.normalize_symbol(symbol)
            all_mids = self._info.all_mids()
            if hl_symbol in all_mids:
                return float(all_mids[hl_symbol])
            return None
        except Exception as e:
            logger.error(f"Error getting price for {symbol}: {e}")
            return None
    
    # ========== Symbol Info Methods ==========
    
    def get_instruments_info(self, category: str = "linear", symbol: str = None) -> Dict:
        """Get instrument/symbol information in Bybit response format."""
        try:
            meta = self._info.meta()
            universe = meta.get('universe', [])
            
            # Normalize input symbol for comparison
            search_symbol = None
            if symbol:
                search_symbol = self.normalize_symbol(symbol)  # BTCUSDT -> BTC
                logger.debug(f"Searching for symbol: {symbol} -> {search_symbol}")
            
            instruments = []
            for idx, asset in enumerate(universe):
                hl_symbol = asset.get('name', '')
                
                # Skip if searching for specific symbol and this isn't it
                if search_symbol and hl_symbol != search_symbol:
                    continue
                
                bybit_symbol = self.denormalize_symbol(hl_symbol)
                
                # Get szDecimals with proper handling
                sz_decimals = asset.get('szDecimals')
                
                # Log what we got for debugging
                if search_symbol and hl_symbol == search_symbol:
                    logger.info(f"Found {hl_symbol} (index {idx}): szDecimals={sz_decimals}, maxLeverage={asset.get('maxLeverage')}")
                
                # Handle None or invalid szDecimals
                if sz_decimals is None or sz_decimals < 0:
                    if hl_symbol in ['BTC', 'ETH']:
                        sz_decimals = 4
                    elif hl_symbol in ['SOL', 'AVAX', 'LINK', 'DOT']:
                        sz_decimals = 2
                    else:
                        sz_decimals = 3
                    logger.warning(f"Using default szDecimals={sz_decimals} for {hl_symbol}")
                
                # Ensure szDecimals is at least 1
                if sz_decimals == 0:
                    logger.warning(f"szDecimals is 0 for {hl_symbol}, using 4 as minimum")
                    sz_decimals = 4
                
                qty_step = 10 ** (-sz_decimals)
                min_qty = qty_step
                max_leverage = asset.get('maxLeverage', 50)
                
                # Format qty_step as a proper decimal string
                qty_step_str = f"{qty_step:.{sz_decimals}f}"
                
                # Determine tick size based on asset price range
                # Hyperliquid uses different tick sizes for different price ranges
                # BTC (~$70k+): tick size = 1.0
                # ETH (~$3k): tick size = 0.1
                # Lower priced assets: smaller tick sizes
                if hl_symbol == 'BTC':
                    tick_size = 1.0
                elif hl_symbol == 'ETH':
                    tick_size = 0.10
                elif hl_symbol in ['SOL', 'AVAX', 'LINK', 'DOT', 'NEAR', 'APT', 'ARB', 'OP', 'INJ', 'SUI']:
                    tick_size = 0.01
                elif hl_symbol in ['DOGE', 'MATIC', 'XRP', 'ADA']:
                    tick_size = 0.0001
                elif hl_symbol in ['PEPE', 'SHIB', 'WIF', 'BONK']:
                    tick_size = 0.000001
                else:
                    # Default: try to get a sensible tick size
                    # For unknown assets, use 0.01 as a safe default
                    tick_size = 0.01
                
                # Format tick size properly
                if tick_size >= 1:
                    tick_size_str = str(int(tick_size))
                else:
                    tick_decimals = max(0, -int(math.floor(math.log10(tick_size))))
                    tick_size_str = f"{tick_size:.{tick_decimals}f}"
                
                instruments.append({
                    'symbol': bybit_symbol,
                    'baseCoin': hl_symbol,
                    'quoteCoin': 'USDC',
                    'status': 'Trading',
                    'lotSizeFilter': {
                        'basePrecision': qty_step_str,
                        'quotePrecision': '0.00001',
                        'minOrderQty': qty_step_str,
                        'maxOrderQty': '1000000',
                        'qtyStep': qty_step_str,
                    },
                    'priceFilter': {
                        'tickSize': tick_size_str,
                        'minPrice': '0.000001',
                        'maxPrice': '99999999',
                    },
                    'leverageFilter': {
                        'minLeverage': '1',
                        'maxLeverage': str(max_leverage),
                        'leverageStep': '1',
                    },
                    '_szDecimals': sz_decimals,
                    '_tickSize': tick_size,
                })
                
                if search_symbol:
                    break
            
            # If we searched but found nothing, return default
            if search_symbol and not instruments:
                logger.error(f"Symbol {symbol} ({search_symbol}) not found in Hyperliquid universe!")
                instruments.append({
                    'symbol': symbol,
                    'baseCoin': search_symbol,
                    'quoteCoin': 'USDC',
                    'status': 'Trading',
                    'lotSizeFilter': {
                        'basePrecision': '0.0001',
                        'quotePrecision': '0.00001',
                        'minOrderQty': '0.0001',
                        'maxOrderQty': '1000000',
                        'qtyStep': '0.0001',
                    },
                    'priceFilter': {
                        'tickSize': '1',
                        'minPrice': '0.000001',
                        'maxPrice': '99999999',
                    },
                    'leverageFilter': {
                        'minLeverage': '1',
                        'maxLeverage': '50',
                        'leverageStep': '1',
                    },
                    '_szDecimals': 4,
                    '_tickSize': 1.0,
                })
            
            return {'retCode': 0, 'retMsg': 'OK', 'result': {'list': instruments}}
        except Exception as e:
            logger.error(f"Error getting instruments info: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def get_symbol_info(self, symbol: str) -> Optional[Dict]:
        """Get symbol info with caching"""
        if symbol in self._symbol_info_cache:
            cache_time = self._symbol_info_cache_time.get(symbol)
            if cache_time and datetime.now() - cache_time < self._cache_duration:
                return self._symbol_info_cache[symbol]
        
        response = self.get_instruments_info(symbol=symbol)
        if response.get('retCode') == 0:
            instruments = response['result'].get('list', [])
            if instruments:
                self._symbol_info_cache[symbol] = instruments[0]
                self._symbol_info_cache_time[symbol] = datetime.now()
                return instruments[0]
        return None
    
    # ========== Position Methods ==========
    
    def get_positions(self, category: str = "linear", symbol: str = None, settleCoin: str = None) -> Dict:
        """Get position information in Bybit response format."""
        try:
            user_state = self._info.user_state(self._wallet_address)
            asset_positions = user_state.get('assetPositions', [])
            
            positions = []
            for asset_pos in asset_positions:
                position = asset_pos.get('position', {})
                szi = float(position.get('szi', 0))
                
                if abs(szi) < 0.0000001:
                    continue
                
                hl_symbol = position.get('coin', '')
                bybit_symbol = self.denormalize_symbol(hl_symbol)
                
                if symbol and bybit_symbol != symbol.upper():
                    continue
                
                entry_price = float(position.get('entryPx', 0))
                unrealized_pnl = float(position.get('unrealizedPnl', 0))
                leverage = float(position.get('leverage', {}).get('value', 1))
                
                positions.append({
                    'symbol': bybit_symbol,
                    'side': 'Buy' if szi > 0 else 'Sell',
                    'size': str(abs(szi)),
                    'avgPrice': str(entry_price),
                    'entryPrice': str(entry_price),
                    'markPrice': str(entry_price),
                    'unrealisedPnl': str(unrealized_pnl),
                    'leverage': str(leverage),
                    'positionValue': str(abs(szi) * entry_price),
                    'positionStatus': 'Normal',
                    'positionIdx': 0,
                })
            
            return {'retCode': 0, 'retMsg': 'OK', 'result': {'list': positions}}
        except Exception as e:
            logger.error(f"Error getting positions: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def get_position_size(self, symbol: str) -> Tuple[float, str]:
        """Get position size and direction for a symbol"""
        try:
            hl_symbol = self.normalize_symbol(symbol)
            user_state = self._info.user_state(self._wallet_address)
            
            for asset_pos in user_state.get('assetPositions', []):
                position = asset_pos.get('position', {})
                if position.get('coin') == hl_symbol:
                    szi = float(position.get('szi', 0))
                    if abs(szi) > 0.0000001:
                        direction = 'long' if szi > 0 else 'short'
                        return abs(szi), direction
            return 0.0, ''
        except Exception as e:
            logger.error(f"Error getting position size: {e}")
            return 0.0, ''
    
    # ========== Order Methods ==========
    
    def place_order(self, category: str = "linear", symbol: str = None, side: str = None,
                   orderType: str = "Market", qty: str = None, price: str = None,
                   timeInForce: str = "GTC", reduceOnly: bool = False,
                   stopLoss: str = None, takeProfit: str = None, **kwargs) -> Dict:
        """Place an order in Bybit API format."""
        try:
            hl_symbol = self.normalize_symbol(symbol)
            is_buy = side.upper() == 'BUY'
            quantity = float(qty)
            
            # CRITICAL: Validate quantity is not zero
            if quantity <= 0:
                logger.error(f"Cannot place order with zero or negative quantity: {quantity}")
                return {
                    'retCode': -1, 
                    'retMsg': f'Invalid quantity: {quantity}. Quantity must be greater than 0.', 
                    'result': {}
                }
            
            # Get symbol info and format quantity
            symbol_info = self.get_symbol_info(symbol)
            if symbol_info:
                sz_decimals = symbol_info.get('_szDecimals', self._get_size_decimals(hl_symbol))
                quantity = self._round_size(quantity, sz_decimals)
                logger.info(f"Order quantity after rounding: {quantity} ({sz_decimals} decimals)")
                
                # Re-check after rounding
                if quantity <= 0:
                    logger.error(f"Quantity rounded to zero. Original: {qty}, szDecimals: {sz_decimals}")
                    return {
                        'retCode': -1, 
                        'retMsg': f'Quantity too small. Minimum for {symbol} requires more precision.', 
                        'result': {}
                    }
            
            # Format price for limit orders
            limit_price = None
            if orderType.upper() != 'MARKET' and price:
                limit_price = float(price)
                
                # Round price to tick size
                if symbol_info:
                    tick_size = symbol_info.get('_tickSize', 1.0)
                else:
                    # Default tick sizes based on asset
                    if hl_symbol == 'BTC':
                        tick_size = 1.0
                    elif hl_symbol == 'ETH':
                        tick_size = 0.1
                    else:
                        tick_size = 0.01
                
                # Round to nearest tick
                limit_price = round(limit_price / tick_size) * tick_size
                
                # Format based on tick size
                if tick_size >= 1:
                    limit_price = int(limit_price)
                else:
                    tick_decimals = max(0, -int(math.floor(math.log10(tick_size))))
                    limit_price = round(limit_price, tick_decimals)
                
                logger.info(f"Price after tick rounding: {limit_price} (tick size: {tick_size})")
            
            logger.info(f"Placing {orderType} order: {side} {quantity} {hl_symbol} @ {limit_price if limit_price else 'market'}")
            
            if orderType.upper() == 'MARKET':
                if reduceOnly:
                    order_result = self._exchange.market_close(hl_symbol, quantity, slippage=0.05)
                else:
                    order_result = self._exchange.market_open(hl_symbol, is_buy, quantity, slippage=0.05)
            else:
                if limit_price is None:
                    return {'retCode': -1, 'retMsg': 'Price required for limit orders', 'result': {}}
                
                order_result = self._exchange.order(
                    hl_symbol, is_buy, quantity, limit_price,
                    {"limit": {"tif": "Gtc"}}, reduce_only=reduceOnly
                )
            
            logger.info(f"Hyperliquid order response: {order_result}")
            
            if order_result.get('status') == 'ok':
                statuses = order_result.get('response', {}).get('data', {}).get('statuses', [])
                order_id = None
                
                if statuses:
                    status_entry = statuses[0]
                    
                    if 'resting' in status_entry:
                        order_id = status_entry['resting'].get('oid', None)
                    elif 'filled' in status_entry:
                        order_id = status_entry['filled'].get('oid', None)
                    elif 'error' in status_entry:
                        error_msg = status_entry['error']
                        logger.error(f"Order rejected by Hyperliquid: {error_msg}")
                        return {'retCode': -1, 'retMsg': error_msg, 'result': {}}
                    else:
                        # Unknown status type - log it so we can learn what it is
                        logger.warning(f"Unknown order status format from Hyperliquid: {status_entry}")
                        # Try to extract oid from any nested dict in the status
                        for key, value in status_entry.items():
                            if isinstance(value, dict) and 'oid' in value:
                                order_id = value['oid']
                                logger.info(f"Extracted order ID from '{key}' status: {order_id}")
                                break
                
                # If we still don't have an order ID, try to find it from open orders
                if order_id is None:
                    logger.warning(f"Could not extract order ID from response: {order_result}")
                    logger.warning("Attempting to find order ID from open orders...")
                    
                    try:
                        # Query Hyperliquid for open orders on this symbol
                        open_orders = self._info.open_orders(self._wallet_address)
                        
                        if open_orders:
                            # Look for the most recent order matching our parameters
                            for open_order in reversed(open_orders):
                                if open_order.get('coin') == hl_symbol:
                                    order_id = open_order.get('oid')
                                    logger.info(f"Found order ID from open orders query: {order_id}")
                                    break
                    except Exception as lookup_error:
                        logger.error(f"Failed to look up order ID from open orders: {lookup_error}")
                
                # Final check: if we still have no order ID, treat this as a failure
                if order_id is None:
                    logger.error(f"CRITICAL: Order appears to have been accepted by Hyperliquid "
                                f"but we could not extract the order ID. "
                                f"Full response: {order_result}")
                    logger.error("This order may exist on Hyperliquid without being tracked by the bot. "
                                "Check your Hyperliquid positions manually!")
                    return {
                        'retCode': -1, 
                        'retMsg': 'Order may have been placed but order ID could not be extracted. '
                                  'Check Hyperliquid manually.', 
                        'result': {}
                    }
                
                # Convert to string for consistency with Bybit format
                order_id = str(order_id)
                
                return {
                    'retCode': 0, 
                    'retMsg': 'OK', 
                    'result': {
                        'orderId': order_id, 
                        'orderLinkId': ''
                    }
                }
            else:
                error_msg = order_result.get('response', str(order_result))
                logger.error(f"Order failed: {error_msg}")
                return {'retCode': -1, 'retMsg': str(error_msg), 'result': {}}
                
        except Exception as e:
            logger.error(f"Error placing order: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def cancel_order(self, category: str = "linear", symbol: str = None, 
                    orderId: str = None, **kwargs) -> Dict:
        """Cancel an order in Bybit API format."""
        try:
            hl_symbol = self.normalize_symbol(symbol)
            
            if orderId:
                # IMPORTANT: Hyperliquid requires integer order IDs
                try:
                    order_id_int = int(orderId)
                except (ValueError, TypeError):
                    logger.error(f"Invalid order ID format: {orderId}")
                    return {'retCode': -1, 'retMsg': f'Invalid order ID: {orderId}', 'result': {}}
                
                logger.info(f"Cancelling order {order_id_int} for {hl_symbol}")
                
                # Try cancelling as a regular limit order first
                try:
                    result = self._exchange.cancel(hl_symbol, order_id_int)
                    logger.info(f"Cancel result: {result}")
                    
                    if result.get('status') == 'ok':
                        return {'retCode': 0, 'retMsg': 'OK', 'result': {'orderId': str(orderId)}}
                except Exception as e:
                    logger.warning(f"Regular cancel failed, trying trigger order cancel: {e}")
                
                # If regular cancel failed, try cancelling as a trigger order (SL/TP)
                try:
                    # For trigger orders, we need to use a different approach
                    # The SDK's cancel method should work for triggers too, but let's try bulk cancel
                    result = self._exchange.cancel_by_cloid(hl_symbol, order_id_int)
                    if result and result.get('status') == 'ok':
                        return {'retCode': 0, 'retMsg': 'OK', 'result': {'orderId': str(orderId)}}
                except Exception as e2:
                    logger.warning(f"Trigger cancel also failed: {e2}")
                
                # Final fallback: cancel all orders for this symbol and let the bot re-place what's needed
                # This is aggressive but ensures the order gets cancelled
                logger.warning(f"Individual cancel failed, order may already be cancelled or filled")
                return {'retCode': -1, 'retMsg': 'Cancel failed - order may be filled/cancelled', 'result': {}}
            else:
                return {'retCode': -1, 'retMsg': 'orderId required', 'result': {}}
                
        except Exception as e:
            logger.error(f"Error cancelling order: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def cancel_all_orders(self, category: str = "linear", symbol: str = None, **kwargs) -> Dict:
        """Cancel all open orders for a symbol"""
        try:
            hl_symbol = self.normalize_symbol(symbol) if symbol else None
            result = self._exchange.cancel_all_orders(hl_symbol)
            
            if result.get('status') == 'ok':
                return {'retCode': 0, 'retMsg': 'OK', 'result': {}}
            else:
                return {'retCode': -1, 'retMsg': str(result), 'result': {}}
        except Exception as e:
            logger.error(f"Error cancelling all orders: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def get_open_orders(self, category: str = "linear", symbol: str = None, **kwargs) -> Dict:
        """Get open orders in Bybit API format - includes both limit and trigger orders."""
        try:
            all_orders = self._info.frontend_open_orders(self._wallet_address)
            
            orders = []
            for order in all_orders:
                hl_symbol = order.get('coin', '')
                bybit_symbol = self.denormalize_symbol(hl_symbol)
                
                if symbol and bybit_symbol != symbol.upper():
                    continue
                
                # Get order ID as string for consistent comparison
                order_id = str(order.get('oid', ''))
                
                # Get Hyperliquid's order type directly
                hl_order_type = str(order.get('orderType', 'Limit'))
                is_trigger = order.get('isTrigger', False)
                trigger_price = float(order.get('triggerPx', 0) or 0)
                # tpsl field: 'sl' for stop loss, 'tp' for take profit, '' for regular
                tpsl = str(order.get('tpsl', '') or '')

                # Determine order type for Bybit format
                order_type = 'Limit'
                stop_order_type = ''
                price = float(order.get('limitPx', 0))

                # Check if this is a stop loss trigger order
                if 'stop' in hl_order_type.lower() or (is_trigger and tpsl == 'sl'):
                    order_type = 'Market'
                    stop_order_type = 'StopLoss'
                    if trigger_price > 0:
                        price = trigger_price
                # Check if this is a take profit trigger order
                elif 'take profit' in hl_order_type.lower() or (is_trigger and tpsl == 'tp'):
                    order_type = 'Market'
                    stop_order_type = 'TakeProfit'
                    if trigger_price > 0:
                        price = trigger_price
                # Fallback: any other trigger with no tpsl label — treat as unknown trigger
                elif is_trigger and trigger_price > 0:
                    order_type = 'Market'
                    stop_order_type = 'StopLoss'  # conservative fallback
                    price = trigger_price
                
                orders.append({
                    'orderId': order_id,
                    'symbol': bybit_symbol,
                    'side': 'Buy' if order.get('side') == 'B' else 'Sell',
                    'orderType': order_type,
                    'stopOrderType': stop_order_type,
                    'triggerPrice': str(trigger_price) if trigger_price else '0',
                    'price': str(price),
                    'qty': str(order.get('sz', 0)),
                    'cumExecQty': str(order.get('filled', 0) if 'filled' in order else 0),
                    'orderStatus': 'New',
                    'createdTime': str(order.get('timestamp', int(time.time() * 1000))),
                    'reduceOnly': order.get('reduceOnly', False),
                })
            
            return {'retCode': 0, 'retMsg': 'OK', 'result': {'list': orders}}
            
        except Exception as e:
            logger.error(f"Error getting open orders: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return {'retCode': -1, 'retMsg': str(e), 'result': {'list': []}}
    
    def get_order_history(self, category: str = "linear", symbol: str = None, 
                         limit: int = 50, **kwargs) -> Dict:
        """Get order history in Bybit API format"""
        try:
            fills = self._info.user_fills(self._wallet_address)
            
            orders = []
            for fill in fills[:limit]:
                hl_symbol = fill.get('coin', '')
                bybit_symbol = self.denormalize_symbol(hl_symbol)
                
                if symbol and bybit_symbol != symbol.upper():
                    continue
                
                orders.append({
                    'orderId': str(fill.get('oid', '')),
                    'symbol': bybit_symbol,
                    'side': 'Buy' if fill.get('side') == 'B' else 'Sell',
                    'orderType': 'Market' if fill.get('crossed') else 'Limit',
                    'price': str(fill.get('px', 0)),
                    'qty': str(fill.get('sz', 0)),
                    'cumExecQty': str(fill.get('sz', 0)),
                    'avgPrice': str(fill.get('px', 0)),
                    'orderStatus': 'Filled',
                    'createdTime': str(fill.get('time', 0)),
                })
            
            return {'retCode': 0, 'retMsg': 'OK', 'result': {'list': orders}}
        except Exception as e:
            logger.error(f"Error getting order history: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def get_executions(self, category: str = "linear", symbol: str = None, 
                      limit: int = 50, **kwargs) -> Dict:
        """Get trade executions/fills in Bybit format."""
        try:
            fills = self._info.user_fills(self._wallet_address)
            
            executions = []
            for fill in fills[:limit]:
                hl_symbol = fill.get('coin', '')
                bybit_symbol = self.denormalize_symbol(hl_symbol)
                
                if symbol and bybit_symbol != symbol.upper():
                    continue
                
                executions.append({
                    'symbol': bybit_symbol,
                    'orderId': str(fill.get('oid', '')),
                    'orderLinkId': '',
                    'side': 'Buy' if fill.get('side') == 'B' else 'Sell',
                    'orderPrice': str(fill.get('px', 0)),
                    'orderQty': str(fill.get('sz', 0)),
                    'execPrice': str(fill.get('px', 0)),
                    'execQty': str(fill.get('sz', 0)),
                    'execFee': str(fill.get('fee', 0)),
                    'execType': 'Trade',
                    'execTime': str(fill.get('time', 0)),
                    'isMaker': not fill.get('crossed', True),
                })
            
            return {'retCode': 0, 'retMsg': 'OK', 'result': {'list': executions}}
        except Exception as e:
            logger.error(f"Error getting executions: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    def get_closed_pnl(self, category: str = "linear", symbol: str = None,
                      limit: int = 50, **kwargs) -> Dict:
        """Get closed PnL records in Bybit format."""
        try:
            return {'retCode': 0, 'retMsg': 'OK', 'result': {'list': []}}
        except Exception as e:
            logger.error(f"Error getting closed PnL: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    # ========== Leverage Methods ==========
    
    def set_leverage(self, category: str = "linear", symbol: str = None,
                    buyLeverage: str = None, sellLeverage: str = None, **kwargs) -> Dict:
        """Set leverage for a symbol in Bybit API format."""
        try:
            hl_symbol = self.normalize_symbol(symbol)
            leverage = int(float(buyLeverage or sellLeverage or 10))
            
            result = self._exchange.update_leverage(leverage, hl_symbol, is_cross=True)
            
            if result.get('status') == 'ok':
                return {'retCode': 0, 'retMsg': 'OK', 'result': {}}
            else:
                error_str = str(result)
                if 'not modified' in error_str.lower():
                    return {'retCode': 10001, 'retMsg': 'leverage not modified', 'result': {}}
                return {'retCode': -1, 'retMsg': str(result), 'result': {}}
        except Exception as e:
            error_str = str(e)
            if 'not modified' in error_str.lower():
                return {'retCode': 10001, 'retMsg': 'leverage not modified', 'result': {}}
            logger.error(f"Error setting leverage: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    # ========== TP/SL Methods ==========
    
    def set_trading_stop(self, category: str = "linear", symbol: str = None,
                        stopLoss: str = None, takeProfit: str = None, **kwargs) -> Dict:
        """Set stop loss and take profit in Bybit API format."""
        try:
            hl_symbol = self.normalize_symbol(symbol)
            pos_size, direction = self.get_position_size(symbol)
            
            if pos_size == 0:
                return {'retCode': -1, 'retMsg': 'No position found', 'result': {}}
            
            is_long = direction == 'long'
            sl_order_id = None
            tp_order_id = None
            # Get tick size for price rounding
            symbol_info = self.get_symbol_info(symbol)
            if symbol_info:
                tick_size = symbol_info.get('_tickSize', 1.0)
            else:
                if hl_symbol == 'BTC':
                    tick_size = 1.0
                elif hl_symbol == 'ETH':
                    tick_size = 0.1
                else:
                    tick_size = 0.01

            if stopLoss:
                sl_price = round(float(stopLoss) / tick_size) * tick_size
                if tick_size >= 1:
                    sl_price = int(sl_price)
                else:
                    tick_decimals = max(0, -int(math.floor(math.log10(tick_size))))
                    sl_price = round(sl_price, tick_decimals)
                logger.info(f"SL price after tick rounding: {sl_price} (original: {stopLoss}, tick size: {tick_size})")
                sl_result = self._exchange.order(
                    hl_symbol, not is_long, pos_size, sl_price,
                    {"trigger": {"triggerPx": sl_price, "isMarket": True, "tpsl": "sl"}},
                    reduce_only=True
                )
                logger.info(f"SL order raw response: {sl_result}")
                
                if sl_result.get('status') != 'ok':
                    logger.error(f"Failed to set stop loss: {sl_result}")
                    return {'retCode': -1, 'retMsg': f'Failed to set stop loss: {sl_result}', 'result': {}}
                
                # Check inner statuses for errors
                sl_statuses = sl_result.get('response', {}).get('data', {}).get('statuses', [])
                if sl_statuses:
                    sl_status_entry = sl_statuses[0]
                    if 'error' in sl_status_entry:
                        logger.error(f"SL order rejected by Hyperliquid: {sl_status_entry['error']}")
                        return {'retCode': -1, 'retMsg': f"SL rejected: {sl_status_entry['error']}", 'result': {}}
                    elif 'resting' in sl_status_entry:
                        sl_order_id = sl_status_entry['resting'].get('oid')
                        logger.info(f"SL trigger order resting with ID: {sl_order_id}")
                    elif 'filled' in sl_status_entry:
                        sl_order_id = sl_status_entry['filled'].get('oid')
                        logger.info(f"SL trigger order filled immediately with ID: {sl_order_id}")
            
            if takeProfit:
                tp_price = round(float(takeProfit) / tick_size) * tick_size
                if tick_size >= 1:
                    tp_price = int(tp_price)
                else:
                    tick_decimals = max(0, -int(math.floor(math.log10(tick_size))))
                    tp_price = round(tp_price, tick_decimals)
                logger.info(f"TP price after tick rounding: {tp_price} (original: {takeProfit}, tick size: {tick_size})")
                tp_result = self._exchange.order(
                    hl_symbol, not is_long, pos_size, tp_price,
                    {"trigger": {"triggerPx": tp_price, "isMarket": True, "tpsl": "tp"}},
                    reduce_only=True
                )
                logger.info(f"TP order raw response: {tp_result}")
                
                if tp_result.get('status') != 'ok':
                    logger.error(f"Failed to set take profit: {tp_result}")
                    return {'retCode': -1, 'retMsg': f'Failed to set take profit: {tp_result}', 'result': {}}
                
                # Check inner statuses for errors
                tp_statuses = tp_result.get('response', {}).get('data', {}).get('statuses', [])
                if tp_statuses:
                    tp_status_entry = tp_statuses[0]
                    if 'error' in tp_status_entry:
                        logger.error(f"TP order rejected by Hyperliquid: {tp_status_entry['error']}")
                        return {'retCode': -1, 'retMsg': f"TP rejected: {tp_status_entry['error']}", 'result': {}}
                    elif 'resting' in tp_status_entry:
                        tp_order_id = tp_status_entry['resting'].get('oid')
                        logger.info(f"TP trigger order resting with ID: {tp_order_id}")
                    elif 'filled' in tp_status_entry:
                        tp_order_id = tp_status_entry['filled'].get('oid')
                        logger.info(f"TP trigger order filled immediately with ID: {tp_order_id}")
                
            result = {}
            if sl_order_id:
                result['sl_order_id'] = str(sl_order_id)
            if tp_order_id:
                result['tp_order_id'] = str(tp_order_id)
            return {'retCode': 0, 'retMsg': 'OK', 'result': result}
            
        except Exception as e:
            logger.error(f"Error setting trading stop: {e}")
            return {'retCode': -1, 'retMsg': str(e), 'result': {}}
    
    # ========== Helper Methods ==========
    
    def _get_size_decimals(self, hl_symbol: str) -> int:
        """Get size decimals for a symbol"""
        try:
            meta = self._info.meta()
            for asset in meta.get('universe', []):
                if asset.get('name') == hl_symbol:
                    return asset.get('szDecimals', 3)
            return 3
        except:
            return 3
    
    def _round_size(self, size: float, decimals: int) -> float:
        """Round size to appropriate decimals"""
        factor = 10 ** decimals
        return int(size * factor) / factor
    
    def test_connection(self) -> bool:
        """Test if connection to Hyperliquid is working"""
        try:
            balance = self.get_balance()
            if balance >= 0:
                logger.info(f"Connection test successful. Balance: ${balance:.2f}")
                return True
            return False
        except Exception as e:
            logger.error(f"Connection test failed: {e}")
            return False


# ============================================================
# Wrapper class that provides a pybit-compatible interface
# ============================================================

class HyperliquidClient:
    """Wrapper that makes HyperliquidBridge fully compatible with pybit interface."""
    
    def __init__(self, bridge: HyperliquidBridge = None, **kwargs):
        self.bridge = bridge or HyperliquidBridge()
    
    def get_wallet_balance(self, **kwargs) -> Dict:
        return self.bridge.get_wallet_balance(**kwargs)
    
    def get_tickers(self, **kwargs) -> Dict:
        return self.bridge.get_tickers(**kwargs)
    
    def get_instruments_info(self, **kwargs) -> Dict:
        return self.bridge.get_instruments_info(**kwargs)
    
    def get_positions(self, **kwargs) -> Dict:
        return self.bridge.get_positions(**kwargs)
    
    def place_order(self, **kwargs) -> Dict:
        return self.bridge.place_order(**kwargs)
    
    def cancel_order(self, **kwargs) -> Dict:
        return self.bridge.cancel_order(**kwargs)
    
    def cancel_all_orders(self, **kwargs) -> Dict:
        return self.bridge.cancel_all_orders(**kwargs)
    
    def get_open_orders(self, **kwargs) -> Dict:
        return self.bridge.get_open_orders(**kwargs)
    
    def get_order_history(self, **kwargs) -> Dict:
        return self.bridge.get_order_history(**kwargs)
    
    def set_leverage(self, **kwargs) -> Dict:
        return self.bridge.set_leverage(**kwargs)
    
    def set_trading_stop(self, **kwargs) -> Dict:
        return self.bridge.set_trading_stop(**kwargs)
    
    def get_server_time(self, **kwargs) -> Dict:
        return self.bridge.get_server_time(**kwargs)
    
    def get_api_key_information(self, **kwargs) -> Dict:
        return self.bridge.get_api_key_information(**kwargs)
    
    def get_executions(self, **kwargs) -> Dict:
        return self.bridge.get_executions(**kwargs)
    
    def get_closed_pnl(self, **kwargs) -> Dict:
        return self.bridge.get_closed_pnl(**kwargs)


# ============================================================
# Standalone functions
# ============================================================

_global_bridge: Optional[HyperliquidBridge] = None

def get_bridge() -> HyperliquidBridge:
    """Get or create global bridge instance"""
    global _global_bridge
    if _global_bridge is None:
        _global_bridge = HyperliquidBridge()
    return _global_bridge

def get_hyperliquid_client(**kwargs) -> HyperliquidClient:
    """Get a pybit-compatible client backed by Hyperliquid"""
    return HyperliquidClient(bridge=get_bridge())


if __name__ == "__main__":
    print("Testing Hyperliquid Bridge...")
    
    bridge = HyperliquidBridge()
    
    if bridge.test_connection():
        print("✓ Connection successful")
        
        # Test get_server_time (the method that was missing)
        server_time = bridge.get_server_time()
        print(f"✓ Server time: {server_time}")
        
        btc_price = bridge.get_current_price("BTCUSDT")
        print(f"✓ BTC Price: ${btc_price}")
        
        positions = bridge.get_positions()
        print(f"✓ Positions: {positions}")
    else:
        print("✗ Connection failed")