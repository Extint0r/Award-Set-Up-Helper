import sqlite3
import os
from datetime import datetime

DB_NAME = "pipeline_tracker.db"

def initialize_db():
    """
    Initializes a local SQLite database to track execution states 
    and maintain checkpoint memory across massive batch operations.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Create a table to track the state of every single file in your loop
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS file_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT UNIQUE NOT NULL,
            file_path TEXT NOT NULL,
            status TEXT NOT NULL, -- 'Pending', 'Success', 'Failed'
            parent_proposal_number TEXT, -- The unique key from Protocol 1.71
            error_log TEXT,
            processed_at TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

def register_directory_files(file_list):
    """
    Scans your physical folder and registers any newly discovered files 
    into the database as 'Pending' without erasing history.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    new_count = 0
    for filepath in file_list:
        filename = os.path.basename(filepath)
        try:
            # INSERT OR IGNORE skips the file if it was already recorded in a previous run
            cursor.execute("""
                INSERT OR IGNORE INTO file_log (file_name, file_path, status)
                VALUES (?, ?, 'Pending')
            """, (filename, filepath))
            if cursor.rowcount > 0:
                new_count += 1
        except sqlite3.Error as e:
            print(f"⚠️ Database registration error for {filename}: {e}")
            
    conn.commit()
    conn.close()
    return new_count

def get_next_pending_file():
    """
    Fetches the single next 'Pending' file from the queue. 
    Enables true sequential stream processing.
    """
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row  # Access columns by name like a dictionary
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT file_name, file_path 
        FROM file_log 
        WHERE status = 'Pending' 
        LIMIT 1
    """)
    row = cursor.fetchone()
    conn.close()
    return dict(row) if row else None

def update_file_status(filename, status, proposal_number=None, error_log=None):
    """
    Updates a file's state boundary. Logs unique proposal anchor keys
    or tracking exceptions on failure.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cursor.execute("""
        UPDATE file_log 
        SET status = ?, 
            parent_proposal_number = ?, 
            error_log = ?, 
            processed_at = ?
        WHERE file_name = ?
    """, (status, proposal_number, error_log, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), filename))
    
    conn.commit()
    conn.close()

def get_pipeline_metrics():
    """
    Aggregates database totals to feed real-time metric counters 
    on your Streamlit dashboard dashboard view.
    """
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cursor.execute("SELECT COUNT(*) FROM file_log")
    total = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM file_log WHERE status = 'Pending'")
    pending = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM file_log WHERE status = 'Success'")
    success = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM file_log WHERE status = 'Failed'")
    failed = cursor.fetchone()[0]
    
    conn.close()
    return {"total": total, "pending": pending, "success": success, "failed": failed}

if __name__ == "__main__":
    initialize_db()
    print("✨ SQLite tracking database initialized successfully.")
    print(f"📊 Current Metrics: {get_pipeline_metrics()}")