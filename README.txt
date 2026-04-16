AI-POWERED USER REVIEW SYSTEM
=============================

SETUP
-----
1. Install Python 3.9+
2. Install dependencies:
   pip install -r requirements.txt
3. Run the application:
   python app.py
4. Open browser at http://localhost:5000

FILES
-----
1. app.py           - Main application, configuration, and Excel handling
2. database.py      - Database management (SQLite)
3. gemini.py        - AI logic for parsing emails
4. email_system.py  - Email sending and receiving logic

USAGE
-----
1. Step 1: Upload 'owners.xlsx' (defines who manages which department)
2. Step 2: Upload 'users.xlsx' (list of users to review)
3. Step 3: 'Generate Tickets' sends emails to owners.
   - Owners reply to the email with instructions (e.g., "Delete John", "Change Role of Jane").
   - Click 'Fetch Responses' to download emails.
   - Click 'Process Responses' to let AI execute changes.
4. Step 4: View 'Final User List' and 'Change Logs'.

Note: The system resets data on page reload for security execution in this demo environment.
