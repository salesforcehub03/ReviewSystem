import sqlite3
from datetime import datetime
import json

class Database:
    def __init__(self, db_path='user_review_system.db'):
        self.db_path = db_path
        self.init_db()
    
    def get_connection(self):
        """Get database connection"""
        conn = sqlite3.connect(self.db_path, timeout=5.0) # 5 second timeout
        conn.row_factory = sqlite3.Row
        
        # Enable WAL mode for concurrency
        conn.execute('PRAGMA journal_mode=WAL;')
        conn.execute('PRAGMA synchronous=NORMAL;')
        
        return conn
    
    def init_db(self):
        """Initialize database tables"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Owners table
        cursor.execute('''

            CREATE TABLE IF NOT EXISTS owners (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name TEXT NOT NULL,
                email TEXT NOT NULL UNIQUE,
                owner_type TEXT NOT NULL,
                cc_email TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # MIGRATION: Check if cc_email exists (for existing databases)
        cursor.execute("PRAGMA table_info(owners)")
        columns = [column[1] for column in cursor.fetchall()]
        if 'cc_email' not in columns:
            cursor.execute("ALTER TABLE owners ADD COLUMN cc_email TEXT")
        
        # Users table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_name TEXT NOT NULL,
                email TEXT NOT NULL UNIQUE,
                last_login TEXT,
                roles TEXT,
                groups TEXT,
                department TEXT NOT NULL,
                owner_email TEXT,
                status TEXT DEFAULT 'active',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Tickets table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS tickets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id TEXT NOT NULL UNIQUE,
                department TEXT NOT NULL,
                owner_email TEXT NOT NULL,
                cc_emails TEXT,
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status TEXT DEFAULT 'pending'
            )
        ''')
        
        # Email responses table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS email_responses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id TEXT NOT NULL,
                from_email TEXT NOT NULL,
                subject TEXT,
                body TEXT,
                has_attachment BOOLEAN DEFAULT 0,
                attachment_data TEXT,
                received_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                processed BOOLEAN DEFAULT 0,
                FOREIGN KEY (ticket_id) REFERENCES tickets(ticket_id)
            )
        ''')
        
        # Change logs table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS change_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ticket_id TEXT NOT NULL,
                action_type TEXT NOT NULL,
                user_email TEXT,
                old_value TEXT,
                new_value TEXT,
                description TEXT,
                performed_by TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (ticket_id) REFERENCES tickets(ticket_id)
            )
        ''')
        
        # Configuration table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS configuration (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                key TEXT NOT NULL UNIQUE,
                value TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        conn.commit()
        conn.close()
    
    # Owners operations
    def add_owner(self, user_name, email, owner_type, cc_email=None):
        """Add owner to database"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                'INSERT INTO owners (user_name, email, owner_type, cc_email) VALUES (?, ?, ?, ?)',
                (user_name, email, owner_type, cc_email)
            )
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            # Owner already exists, update it
            cursor.execute(
                'UPDATE owners SET user_name = ?, owner_type = ?, cc_email = ? WHERE email = ?',
                (user_name, owner_type, cc_email, email)
            )
            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()
    
    def get_owners(self):
        """Get all owners"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM owners ORDER BY owner_type, user_name')
        owners = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return owners
    
    def clear_owners(self):
        """Clear all owners"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM owners')
        conn.commit()
        conn.close()
    
    # Users operations
    def add_user(self, user_name, email, last_login, roles, groups, department, owner_email):
        """Add user to database"""
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                '''INSERT INTO users (user_name, email, last_login, roles, groups, department, owner_email)
                   VALUES (?, ?, ?, ?, ?, ?, ?)''',
                (user_name, email, last_login, roles, groups, department, owner_email)
            )
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            # User already exists, update it
            cursor.execute(
                '''UPDATE users SET user_name = ?, last_login = ?, roles = ?, groups = ?, 
                   department = ?, owner_email = ?, updated_at = CURRENT_TIMESTAMP 
                   WHERE email = ?''',
                (user_name, last_login, roles, groups, department, owner_email, email)
            )
            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()
    
    def get_users(self, department=None, status='active'):
        """Get users, optionally filtered by department and status"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        query = "SELECT * FROM users"
        params = []
        conditions = []
        
        if department:
            conditions.append("department = ?")
            params.append(department)
            
        if status != 'all':
            conditions.append("LOWER(status) = LOWER(?)")
            params.append(status)
            
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
            
        query += " ORDER BY user_name"
        
        cursor.execute(query, tuple(params))
        users = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return users
    
    def update_user_status(self, email, status):
        """Update user status"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE users SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE email = ?',
            (status, email)
        )
        conn.commit()
        conn.close()
    
    def update_user_role(self, email, new_role):
        """Update user role"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE users SET roles = ?, updated_at = CURRENT_TIMESTAMP WHERE email = ?',
            (new_role, email)
        )
        conn.commit()
        conn.close()
    
    def delete_user(self, email):
        """Delete user (soft delete by setting status to deleted)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE users SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE email = ?',
            ('deleted', email)
        )
        conn.commit()
        conn.close()
    
    def clear_users(self):
        """Clear all users"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM users')
        conn.commit()
        conn.close()
    
    # Tickets operations
    def add_ticket(self, ticket_id, department, owner_email, cc_emails=None):
        """Add ticket to database"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'INSERT INTO tickets (ticket_id, department, owner_email, cc_emails) VALUES (?, ?, ?, ?)',
            (ticket_id, department, owner_email, cc_emails)
        )
        conn.commit()
        conn.close()
    
    def get_tickets(self):
        """Get all tickets"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM tickets ORDER BY sent_at DESC')
        tickets = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return tickets
    
    def update_ticket_status(self, ticket_id, status):
        """Update ticket status"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE tickets SET status = ? WHERE ticket_id = ?',
            (status, ticket_id)
        )
        conn.commit()
        conn.close()
    
    # Email responses operations
    def add_email_response(self, ticket_id, from_email, subject, body, has_attachment=False, attachment_data=None):
        """Add email response to database"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            '''INSERT INTO email_responses 
               (ticket_id, from_email, subject, body, has_attachment, attachment_data)
               VALUES (?, ?, ?, ?, ?, ?)''',
            (ticket_id, from_email, subject, body, has_attachment, attachment_data)
        )
        conn.commit()
        conn.close()
    
    def get_email_responses(self, ticket_id=None, processed=None):
        """Get email responses"""
        conn = self.get_connection()
        cursor = conn.cursor()
        if ticket_id:
            if processed is not None:
                cursor.execute(
                    'SELECT * FROM email_responses WHERE ticket_id = ? AND processed = ? ORDER BY received_at DESC',
                    (ticket_id, processed)
                )
            else:
                cursor.execute(
                    'SELECT * FROM email_responses WHERE ticket_id = ? ORDER BY received_at DESC',
                    (ticket_id,)
                )
        else:
            if processed is not None:
                cursor.execute(
                    'SELECT * FROM email_responses WHERE processed = ? ORDER BY received_at DESC',
                    (processed,)
                )
            else:
                cursor.execute('SELECT * FROM email_responses ORDER BY received_at DESC')
        responses = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return responses
    
    def mark_response_processed(self, response_id):
        """Mark email response as processed"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'UPDATE email_responses SET processed = 1 WHERE id = ?',
            (response_id,)
        )
        conn.commit()
        conn.close()
    
    # Change logs operations
    def add_change_log(self, ticket_id, action_type, user_email, old_value, new_value, description, performed_by):
        """Add change log entry"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            '''INSERT INTO change_logs 
               (ticket_id, action_type, user_email, old_value, new_value, description, performed_by)
               VALUES (?, ?, ?, ?, ?, ?, ?)''',
            (ticket_id, action_type, user_email, old_value, new_value, description, performed_by)
        )
        conn.commit()
        conn.close()
    
    def get_change_logs(self, ticket_id=None):
        """Get change logs"""
        conn = self.get_connection()
        cursor = conn.cursor()
        if ticket_id:
            cursor.execute(
                'SELECT * FROM change_logs WHERE ticket_id = ? ORDER BY created_at DESC',
                (ticket_id,)
            )
        else:
            cursor.execute('SELECT * FROM change_logs ORDER BY created_at DESC')
        logs = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return logs
    
    # Configuration operations
    def set_config(self, key, value):
        """Set configuration value"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            '''INSERT OR REPLACE INTO configuration (key, value, updated_at)
               VALUES (?, ?, CURRENT_TIMESTAMP)''',
            (key, value)
        )
        conn.commit()
        conn.close()
    
    def get_config(self, key):
        """Get configuration value"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT value FROM configuration WHERE key = ?', (key,))
        row = cursor.fetchone()
        conn.close()
        return row['value'] if row else None
    def reset_system_data(self):
        """
        Reset transactional data (Change Logs, Email Responses)
        Does NOT delete Users or Owners or Tickets (configuration)
        """
        conn = self.get_connection()
        conn.execute('DELETE FROM change_logs')
        conn.execute('DELETE FROM email_responses')
        # We also reset ticket status to 'pending' so they can be reused if needed?
        # Actually, let's just clear the logs/responses as requested.
        conn.commit()
        conn.close()
        return True

    def clear_process_data(self):
        """
        Clear all process data (tickets, responses, logs)
        Keeps users, owners, and configuration.
        """
        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            # Order likely doesn't matter for SQLite unless FK constraints are strictly enforced
            # But safer to delete children first
            cursor.execute("DELETE FROM change_logs")
            cursor.execute("DELETE FROM email_responses")
            cursor.execute("DELETE FROM tickets")
            conn.commit()
            return True, "Process data cleared successfully"
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()
