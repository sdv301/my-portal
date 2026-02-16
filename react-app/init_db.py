import sqlite3
import os

def init_db():
    conn = sqlite3.connect('data.sqlite')
    cursor = conn.cursor()
    
    with open('schema.sql', 'r', encoding='utf-8') as f:
        cursor.executescript(f.read())
        
    with open('insert_data.sql', 'r', encoding='utf-8') as f:
        cursor.executescript(f.read())
        
    conn.commit()
    conn.close()
    print("Database initialized successfully.")

if __name__ == "__main__":
    init_db()
