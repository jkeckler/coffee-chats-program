# captain_history.py
import sqlite3
from datetime import datetime
import os
from typing import Dict, List, Optional, Tuple

class CaptainHistoryDB:
    """Handles persistent storage of captain assignments and meeting participation"""
    
    def __init__(self, db_path: str = "data/captain_history.db"):
        """Initialize database connection and create tables if they don't exist"""
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        self.db_path = db_path
        self._create_tables()

    def _create_tables(self):
        """Create necessary database tables if they don't exist"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Create captain history table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS captain_history (
                    employee_id TEXT PRIMARY KEY,
                    employee_name TEXT NOT NULL,
                    department TEXT NOT NULL,
                    captain_count INTEGER DEFAULT 0,
                    last_captain_date TEXT,
                    total_meetings_attended INTEGER DEFAULT 0,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Create meetings table for detailed history
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS meetings (
                    meeting_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    captain_id TEXT NOT NULL,
                    group_members TEXT NOT NULL,  -- Stored as comma-separated employee IDs
                    meeting_time TEXT NOT NULL,
                    FOREIGN KEY (captain_id) REFERENCES captain_history (employee_id)
                )
            """)
            
            conn.commit()

    def get_captain_stats(self, employee_id: str) -> Optional[Tuple[int, str]]:
        """Get captain count and last captain date for an employee"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT captain_count, last_captain_date
                FROM captain_history
                WHERE employee_id = ?
            """, (employee_id,))
            
            result = cursor.fetchone()
            if result:
                return result[0], result[1]
            return None

    def update_or_create_employee(self, employee_data: Dict) -> None:
        """Update or create employee record in captain history"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            cursor.execute("""
                INSERT INTO captain_history (
                    employee_id, employee_name, department, captain_count,
                    total_meetings_attended, updated_at
                ) VALUES (?, ?, ?, 0, 0, CURRENT_TIMESTAMP)
                ON CONFLICT(employee_id) DO UPDATE SET
                    employee_name = excluded.employee_name,
                    department = excluded.department,
                    updated_at = CURRENT_TIMESTAMP
            """, (
                employee_data['id'],
                employee_data['name'],
                employee_data['department']
            ))
            
            conn.commit()

    def record_meeting(self, captain_id: str, group_members: List[str], meeting_time: str) -> None:
        """Record a new meeting with captain and participants"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Update captain's record
            cursor.execute("""
                UPDATE captain_history 
                SET captain_count = captain_count + 1,
                    last_captain_date = ?,
                    total_meetings_attended = total_meetings_attended + 1,
                    updated_at = CURRENT_TIMESTAMP
                WHERE employee_id = ?
            """, (datetime.now().isoformat(), captain_id))
            
            # Update other members' meeting counts
            member_ids = [m for m in group_members if m != captain_id]
            for member_id in member_ids:
                cursor.execute("""
                    UPDATE captain_history 
                    SET total_meetings_attended = total_meetings_attended + 1,
                        updated_at = CURRENT_TIMESTAMP
                    WHERE employee_id = ?
                """, (member_id,))
            
            # Record meeting details
            cursor.execute("""
                INSERT INTO meetings (date, captain_id, group_members, meeting_time)
                VALUES (?, ?, ?, ?)
            """, (
                datetime.now().isoformat(),
                captain_id,
                ','.join(group_members),
                meeting_time
            ))
            
            conn.commit()