import sqlite3
import os

db_path = 'instance/adenticket.db'
if not os.path.exists(db_path):
    # Try current directory if not in instance
    db_path = 'adenticket.db'

if os.path.exists(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    try:
        print(f"Updating database at {db_path}...")
        cursor.execute("ALTER TABLE concert ADD COLUMN date VARCHAR(200) DEFAULT ''")
        print("Added column 'date'")
    except sqlite3.OperationalError as e:
        print(f"Column 'date' might already exist: {e}")
        
    try:
        cursor.execute("ALTER TABLE concert ADD COLUMN location VARCHAR(500) DEFAULT ''")
        print("Added column 'location'")
    except sqlite3.OperationalError as e:
        print(f"Column 'location' might already exist: {e}")
        
    conn.commit()
    conn.close()
    print("Database update complete.")
else:
    print(f"Database not found at {db_path}")
