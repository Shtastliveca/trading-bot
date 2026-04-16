# db_queue.py - Simple database queue system for trading bot
import sqlite3
import threading
import queue
import logging
import time
from datetime import datetime
#from bybit_trading_bot import get_current_capital


# Set up logger
logger = logging.getLogger("db_queue")

# Global variables
db_queue = queue.Queue()
worker_thread = None
is_running = False

# Connection settings
DB_PATH = 'trading_bot.db'
CONNECTION_TIMEOUT = 30.0  # seconds

def get_optimized_connection():
    """Get a SQLite connection with optimized settings"""
    conn = sqlite3.connect(DB_PATH, timeout=CONNECTION_TIMEOUT)
    
    # Enable WAL mode for better concurrency
    conn.execute('PRAGMA journal_mode = WAL')
    
    # Set busy timeout (in milliseconds)
    conn.execute('PRAGMA busy_timeout = 10000')
    
    # Set synchronous mode to reduce waiting for disk writes
    conn.execute('PRAGMA synchronous = NORMAL')
    
    return conn

def worker_function():
    """Background worker that processes database operations sequentially"""
    global is_running
    
    logger.info("Database worker thread started")
    
    while is_running:
        try:
            # Get the next operation from the queue with a timeout
            # This allows the thread to check the is_running flag periodically
            try:
                operation, args, callback = db_queue.get(timeout=3.0)
            except queue.Empty:
                # No items in the queue, continue the loop
                continue
                
            # Execute the database operation
            try:
                result = operation(*args)
                # Call the callback with the result if one was provided
                if callback:
                    callback(True, result)
            except Exception as e:
                logger.error(f"Error executing database operation: {str(e)}")
                # Call the callback with the error if one was provided
                if callback:
                    callback(False, str(e))
            finally:
                # Mark the task as done in the queue
                db_queue.task_done()
                
        except Exception as e:
            logger.error(f"Error in database worker: {str(e)}")
            # Wait a bit before continuing to avoid busy loops on persistent errors
            time.sleep(1.0)
    
    logger.info("Database worker thread stopped")

def start_worker():
    """Start the database worker thread if it's not already running"""
    global worker_thread, is_running
    
    if not is_running:
        is_running = True
        worker_thread = threading.Thread(target=worker_function, daemon=True)
        worker_thread.start()
        logger.info("Database worker started")

def stop_worker():
    """Stop the database worker thread gracefully"""
    global is_running
    
    if is_running:
        logger.info("Stopping database worker...")
        is_running = False
        
        # Wait for the thread to finish, but don't block indefinitely
        if worker_thread and worker_thread.is_alive():
            worker_thread.join(timeout=5.0)
            
        logger.info("Database worker stopped")

def queue_operation(operation, args=(), callback=None):
    """Add an operation to the database queue"""
    # Start the worker if not already running
    if not is_running:
        start_worker()
        
    # Add the operation to the queue
    db_queue.put((operation, args, callback))
    
    # Return a queue size for information
    return db_queue.qsize()

# Database operations
def save_position(symbol, position_data):
    """Save position to database"""
    conn = get_optimized_connection()
    cursor = conn.cursor()
    
    try:
        # Check if all required columns exist
        cursor.execute("PRAGMA table_info(positions)")
        columns = [col[1] for col in cursor.fetchall()]
        
        # Start with symbol as it's our primary key
        query_columns = ['symbol']
        query_values = [symbol]
        
        # Map position_data keys to database columns
        column_mappings = {
            'direction': 'direction',
            'entry_price': 'entry_price',
            'position_size': 'position_size',
            'stop_loss': 'stop_loss',
            'order_id': 'order_id',
            'entry_time': 'entry_time',
            'sl_order_id': 'sl_order_id',
            'last_check_time': 'last_check_time',
            'take_profit': 'take_profit',
            'account_balance': 'account_balance',
            'order_type': 'order_type',
            'tp_order_type': 'tp_order_type',
            'tp_order_id': 'tp_order_id'
        }
        
        # Add columns and values that exist in both the database and position_data
        for data_key, db_column in column_mappings.items():
            if db_column in columns and data_key in position_data:
                query_columns.append(db_column)
                if data_key in position_data:
                    query_values.append(position_data[data_key])
                else:
                    # Handle missing keys with defaults
                    if data_key == 'sl_order_id' or data_key == 'tp_order_id':
                        query_values.append('')
                    elif data_key == 'stop_loss' or data_key == 'take_profit':
                        query_values.append(0)
                    elif data_key == 'last_check_time':
                        query_values.append(datetime.now().isoformat())
                    elif data_key == 'account_balance':
                        # This will need the current capital function to be passed as an argument
                        # or imported, but for now we'll just use None
                        query_values.append(None)
                    elif data_key == 'order_type' or data_key == 'tp_order_type':
                        query_values.append('Market')
                    else:
                        query_values.append(None)
        
        # Build the dynamic query
        placeholders = ','.join(['?'] * len(query_values))
        column_str = ','.join(query_columns)
        
        query = f'''
        INSERT OR REPLACE INTO positions ({column_str})
        VALUES ({placeholders})
        '''
        
        # Execute the query with all values
        cursor.execute(query, query_values)
        
        conn.commit()
        return True
    except Exception as e:
        logger.error(f"Error in save_position operation: {str(e)}")
        return False
    finally:
        conn.close()

def save_pending_order(order_data):
    """Save pending order to database"""
    conn = get_optimized_connection()
    cursor = conn.cursor()
    
    try:
        # Get all column names from pending_orders table
        cursor.execute("PRAGMA table_info(pending_orders)")
        columns = [col[1] for col in cursor.fetchall()]
        
        # Build dynamic insert query based on available columns
        column_names = []
        placeholders = []
        values = []
        
        # Process all available fields in order_data
        for field, value in order_data.items():
            if field in columns:
                column_names.append(field)
                placeholders.append('?')
                values.append(value)
        
        # Create dynamic query
        query = f'''
        INSERT OR REPLACE INTO pending_orders 
        ({', '.join(column_names)})
        VALUES ({', '.join(placeholders)})
        '''
        
        cursor.execute(query, values)
        conn.commit()
        return True
    except Exception as e:
        logger.error(f"Error in save_pending_order operation: {str(e)}")
        return False
    finally:
        conn.close()

def remove_pending_order(order_id):
    """Remove pending order from database"""
    conn = get_optimized_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute('DELETE FROM pending_orders WHERE order_id = ?', (order_id,))
        conn.commit()
        return True
    except Exception as e:
        logger.error(f"Error in remove_pending_order operation: {str(e)}")
        return False
    finally:
        conn.close()

def remove_position(symbol):
    """Remove position from database"""
    conn = get_optimized_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute('DELETE FROM positions WHERE symbol = ?', (symbol,))
        conn.commit()
        return True
    except Exception as e:
        logger.error(f"Error in remove_position operation: {str(e)}")
        return False
    finally:
        conn.close()

def log_reconciliation_event(event_type, symbol, details, position_data=None):
    """Log reconciliation event to database"""
    conn = get_optimized_connection()
    cursor = conn.cursor()
    
    try:
        position_json = None
        if position_data:
            import json
            position_json = json.dumps(position_data)
            
        cursor.execute('''
        INSERT INTO reconciliation_log (timestamp, event_type, symbol, details, position_data)
        VALUES (?, ?, ?, ?, ?)
        ''', (
            datetime.now().isoformat(),
            event_type,
            symbol,
            details,
            position_json
        ))
        
        conn.commit()
        return True
    except Exception as e:
        logger.error(f"Error in log_reconciliation_event operation: {str(e)}")
        return False
    finally:
        conn.close()

def update_position_check_time(symbol):
    """Update the last_check_time for a position"""
    conn = get_optimized_connection()
    cursor = conn.cursor()
    
    try:
        # Check if last_check_time column exists
        cursor.execute("PRAGMA table_info(positions)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'last_check_time' in columns:
            cursor.execute('''
            UPDATE positions SET last_check_time = ? WHERE symbol = ?
            ''', (datetime.now().isoformat(), symbol))
            
            conn.commit()
            return True
        return False
    except Exception as e:
        logger.error(f"Error in update_position_check_time operation: {str(e)}")
        return False
    finally:
        conn.close()