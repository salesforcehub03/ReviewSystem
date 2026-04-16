import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import random
import string
import imaplib
import email
from email.header import decode_header
import re
import os
import tempfile
import email.utils
from email.mime.base import MIMEBase
from email import encoders

class EmailAutomation:
    """Handle email automation for ticket generation and distribution"""
    
    def __init__(self, smtp_server, smtp_port, email, password):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email = email
        self.password = password
    
    def generate_ticket_id(self, owner_type='IT'):
        """
        Generate ticket ID: [TYPE]-YYYYDDMM-XX####
        Example: IT-20251712-AB1234
        """
        prefix = 'BUSINESS' if owner_type and owner_type.strip().upper() == 'BUSINESS' else 'IT'
        
        now = datetime.now()
        # Format: Year Date Month (YYYYDDMM) per user request
        date_part = now.strftime("%Y%d%m")
        
        # 2 random letters + 4 random numbers
        random_alpha = ''.join(random.choices(string.ascii_uppercase, k=2))
        random_digits = ''.join(random.choices(string.digits, k=4))
        
        ticket_id = f"{prefix}-{date_part}-{random_alpha}{random_digits}"
        return ticket_id
    
    def create_review_email(self, ticket_id, department, users, owner_name):
        """
        Create email content for user review
        """
        # Create user table
        user_rows = ""
        for user in users:
            user_rows += f"""
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">{user['user_name']}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{user['email']}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{user.get('roles', '')}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{user.get('groups', '')}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{user.get('last_login', '')}</td>
            </tr>
            """
        
        html_content = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                .container {{ max-width: 800px; margin: 0 auto; padding: 20px; }}
                .header {{ background-color: #4CAF50; color: white; padding: 20px; text-align: center; }}
                .content {{ padding: 20px; background-color: #f9f9f9; }}
                table {{ width: 100%; border-collapse: collapse; margin: 20px 0; background-color: white; }}
                th {{ background-color: #4CAF50; color: white; padding: 10px; text-align: left; border: 1px solid #ddd; }}
                td {{ padding: 8px; border: 1px solid #ddd; }}
                .instructions {{ background-color: #fff3cd; padding: 15px; margin: 20px 0; border-left: 4px solid #ffc107; }}
                .footer {{ margin-top: 20px; padding: 15px; background-color: #e9ecef; text-align: center; font-size: 12px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>User Review Request</h1>
                    <p>Ticket ID: {ticket_id}</p>
                </div>
                
                <div class="content">
                    <p>Dear {owner_name},</p>
                    
                    <p>This is a periodic review request for users under the <strong>{department}</strong> department. 
                    Please review the following users and respond with any necessary actions.</p>
                    
                    <h3>Users for Review ({len(users)} total):</h3>
                    <table>
                        <thead>
                            <tr>
                                <th>User Name</th>
                                <th>Email</th>
                                <th>Roles</th>
                                <th>Groups</th>
                                <th>Last Login</th>
                            </tr>
                        </thead>
                        <tbody>
                            {user_rows}
                        </tbody>
                    </table>
                    
                    <div class="instructions">
                        <h3>📝 How to Respond:</h3>
                        <p><strong>Option 1: Reply to this email with text instructions</strong></p>
                        <ul>
                            <li>To delete a user: "Delete [user name or email]"</li>
                            <li>To update a role: "Update [user name] to [new role]" or "Change [user name] role to [new role]"</li>
                            <li>To keep users: "Keep all" or "No changes"</li>
                        </ul>
                        
                        <p><strong>Option 2: Reply with an Excel attachment</strong></p>
                        <ul>
                            <li>Create a sheet named "Actions"</li>
                            <li>Include columns: Action, User Email, Details</li>
                            <li>Example: Action="Delete", User Email="user@example.com"</li>
                        </ul>
                        
                        <p><strong>Examples:</strong></p>
                        <ul>
                            <li>"Delete Rohit Sharma"</li>
                            <li>"Update Lavi Singh to Project Manager"</li>
                            <li>"Change role of virat@gmail.com to Senior Developer"</li>
                            <li>"Remove Gaurav Di and Kiran Singh"</li>
                        </ul>
                    </div>
                    
                    <p><strong>Important:</strong> Please respond within 7 days. Your response will be automatically processed by our AI system.</p>
                </div>
                
                <div class="footer">
                    <p>This is an automated email from the User Review System.</p>
                    <p>Ticket ID: {ticket_id} | Department: {department} | Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
                </div>
            </div>
        </body>
        </html>
        """
        return html_content
    
    def send_email(self, to_email, cc_emails, subject, html_content, attachment_path=None, internal_error=None):
        """Send email via SMTP"""
        try:
            # Create root message (multipart/mixed)
            msg = MIMEMultipart('mixed')
            msg['From'] = self.email
            msg['To'] = to_email
            msg['Subject'] = subject
            
            if cc_emails:
                if isinstance(cc_emails, str):
                    cc_emails = [email.strip() for email in cc_emails.split(',')]
                msg['Cc'] = ', '.join(cc_emails)
            
            # Append any internal errors to the email body for debugging
            if internal_error:
                html_content += f'<div style="background-color: #fee; border: 1px solid #f00; padding: 10px; margin: 10px 0;"><h3>⚠️ System Warning</h3><p>{internal_error}</p></div>'

            # Check attachment validity BEFORE creating email body
            if attachment_path:
                if not os.path.exists(attachment_path):
                    html_content += f'<div style="background-color: #fee; border: 1px solid #f00; padding: 10px;"><h3>⚠️ Attachment Missing</h3><p>System could not find file at: {attachment_path}</p></div>'
                    attachment_path = None # Prevent further errors

            # Create body part (multipart/alternative) to hold text/html
            body_part = MIMEMultipart('alternative')
            html_part = MIMEText(html_content, 'html')
            body_part.attach(html_part)
            
            # Attach body to root
            msg.attach(body_part)
            
            # Attach file if valid
            if attachment_path:
                try:
                    filename = os.path.basename(attachment_path)
                    
                    # Open file in binary mode
                    with open(attachment_path, "rb") as attachment:
                        # Use correct MIME type for Excel
                        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        part.set_payload(attachment.read())
                    
                    encoders.encode_base64(part)
                    
                    # Add header with QUOTED filename and NO space after equals
                    part.add_header(
                        "Content-Disposition",
                        f'attachment; filename="{filename}"'
                    )
                    
                    msg.attach(part)
                except Exception as e:
                    print(f"CRITICAL ERROR attaching file: {e}") 
            
            # Connect and send
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email, self.password)
                
                recipients = [to_email]
                if cc_emails:
                    recipients.extend(cc_emails)
                
                server.send_message(msg)
            
            return True, "Email sent successfully"
            
        except Exception as e:
            return False, f"Failed to send email: {str(e)}"
    
    def send_review_emails(self, tickets_data):
        """Send review emails to all owners"""
        results = []
        
        for ticket in tickets_data:
            subject = f"User Review Request - {ticket['ticket_id']} - {ticket['department']} Department"
            
            html_content = self.create_review_email(
                ticket['ticket_id'],
                ticket['department'],
                ticket['users'],
                ticket['owner_name']
            )
            
            success, message = self.send_email(
                ticket['owner_email'],
                ticket.get('cc_emails'),
                subject,
                html_content,
                ticket.get('attachment_path'),
                ticket.get('error_msg')
            )
            
            results.append({
                'ticket_id': ticket['ticket_id'],
                'owner_email': ticket['owner_email'],
                'success': success,
                'message': message
            })
        
        return results
    
    def test_connection(self):
        """Test SMTP connection"""
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email, self.password)
            return True, "SMTP connection successful"
        except Exception as e:
            return False, f"SMTP connection failed: {str(e)}"

class EmailTracker:
    """Track and monitor email responses via IMAP"""
    
    def __init__(self, imap_server, imap_port, email_address, password):
        self.imap_server = imap_server
        self.imap_port = imap_port
        self.email_address = email_address
        self.password = password
    
    def connect(self):
        """Connect to IMAP server"""
        try:
            mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            mail.login(self.email_address, self.password)
            return mail, None
        except Exception as e:
            return None, f"IMAP connection failed: {str(e)}"
    
    def extract_ticket_id(self, subject):
        """Extract ticket ID from email subject"""
        # Look for pattern IT-YYYYMMDD-XXXXXX or BUSINESS-...
        match = re.search(r'(IT|BUSINESS)-\d{8}-[A-Z0-9]{6}', subject, re.IGNORECASE)
        if match:
            return match.group(0).upper()
        
        # Fallback for legacy format
        match_legacy = re.search(r'IT-\d{4}-\d{2}-\d{2}-\d{4}', subject)
        if match_legacy:
             return match_legacy.group(0)
             
        return None
    
    def decode_email_subject(self, subject):
        """Decode email subject"""
        try:
            decoded_parts = decode_header(subject)
            decoded_subject = ""
            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    if encoding == 'unknown-8bit' or not encoding:
                        encoding = 'utf-8'
                    try:
                        decoded_subject += part.decode(encoding, errors='replace')
                    except LookupError:
                        decoded_subject += part.decode('latin-1', errors='replace')
                else:
                    decoded_subject += part
            return decoded_subject
        except Exception as e:
            return str(subject)
    
    def get_email_body(self, msg):
        """Extract email body from message"""
        body = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                
                # Get text content
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    try:
                        body += part.get_payload(decode=True).decode(errors='ignore')
                    except:
                        pass
                elif content_type == "text/html" and "attachment" not in content_disposition and not body:
                    try:
                        html_body = part.get_payload(decode=True).decode(errors='ignore')
                        # Simple HTML to text conversion
                        body += re.sub('<[^<]+?>', '', html_body)
                    except:
                        pass
        else:
            try:
                body = msg.get_payload(decode=True).decode(errors='ignore')
            except:
                pass
        
        return self.clean_email_body(body.strip())

    def clean_email_body(self, body_text):
        """
        Clean email body by removing quoted text/replies
        """
        if not body_text:
            return ""

        # Pattern 1: On ... wrote:
        pattern_on_wrote = r'On\s+.*,\s+.*\s+wrote:.*'
        
        # Pattern 2: Typical separators
        pattern_separator = r'(?m)^[-_]*\s*Original Message\s*[-_]*$'
        # pattern_from = r'(?m)^From:\s.*$' # Unused variable
        
        # Check "On ... wrote:"
        match = re.search(pattern_on_wrote, body_text, re.DOTALL | re.IGNORECASE)
        if match:
            body_text = body_text[:match.start()]
            
        # Check "Original Message"
        match = re.search(pattern_separator, body_text, re.IGNORECASE)
        if match:
            body_text = body_text[:match.start()]
            
        # Check for simple divider lines
        if '__________' in body_text:
             body_text = body_text.split('__________')[0]
             
        return body_text.strip()
    
    def get_attachments(self, msg):
        """Extract attachments from email"""
        attachments = []
        
        if msg.is_multipart():
            for part in msg.walk():
                content_disposition = str(part.get("Content-Disposition"))
                
                if "attachment" in content_disposition:
                    filename = part.get_filename()
                    if filename:
                        # Decode filename
                        try:
                            decoded_parts = decode_header(filename)
                            filename = ""
                            for part_data, encoding in decoded_parts:
                                if isinstance(part_data, bytes):
                                    if encoding == 'unknown-8bit' or not encoding:
                                        encoding = 'utf-8'
                                    try:
                                        filename += part_data.decode(encoding, errors='replace')
                                    except LookupError:
                                        filename += part_data.decode('latin-1', errors='replace')
                                else:
                                    filename += part_data
                        except Exception:
                            filename = "unknown_attachment"
                        
                        # Save attachment to temp file
                        if filename.endswith(('.xlsx', '.xls')):
                            temp_dir = tempfile.gettempdir()
                            filepath = os.path.join(temp_dir, filename)
                            
                            with open(filepath, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                            
                            attachments.append({
                                'filename': filename,
                                'filepath': filepath
                            })
        
        return attachments
    
    def fetch_responses(self, ticket_ids=None, since_date=None):
        """
        Fetch email responses
        """
        mail, error = self.connect()
        if error:
            return [], error
        
        try:
            # Select inbox
            mail.select('INBOX')
            
            # Build search criteria
            search_criteria = 'ALL'
            if since_date:
                date_str = since_date.strftime("%d-%b-%Y")
                search_criteria = f'(SINCE {date_str})'
            
            # Search for emails
            status, messages = mail.search(None, search_criteria)
            
            if status != 'OK':
                return [], "Failed to search emails"
            
            email_ids = messages[0].split()
            responses = []
            
            # Process emails (latest first)
            # Limit to last 30 emails to improve performance
            for email_id in reversed(email_ids[-30:]):
                try:
                    status, msg_data = mail.fetch(email_id, '(RFC822)')
                    
                    if status != 'OK':
                        continue
                    
                    # Parse email
                    raw_email = msg_data[0][1]
                    msg = email.message_from_bytes(raw_email)
                    
                    # Get subject
                    subject = self.decode_email_subject(msg.get('Subject', ''))
                    
                    # Extract ticket ID
                    ticket_id = self.extract_ticket_id(subject)
                    
                    # Skip if no ticket ID found or not in requested tickets
                    if not ticket_id:
                        continue
                    
                    if ticket_ids and ticket_id not in ticket_ids:
                        continue
                    
                    # Get sender
                    from_header = msg.get('From', '')
                    # Robust extraction
                    from_name, from_email = email.utils.parseaddr(from_header)
                    
                    if not from_email:
                         email_match = re.search(r'[\w\.-]+@[\w\.-]+', from_header)
                         if email_match:
                             from_email = email_match.group(0)
                         else:
                             from_email = from_header
                    
                    # Get date
                    date_str = msg.get('Date', '')
                    
                    # Get body
                    body = self.get_email_body(msg)
                    
                    # Get attachments
                    attachments = self.get_attachments(msg)
                    
                    response = {
                        'ticket_id': ticket_id,
                        'from_email': from_email.lower(),
                        'subject': subject,
                        'body': body,
                        'date': date_str,
                        'has_attachment': len(attachments) > 0,
                        'attachments': attachments
                    }
                    
                    responses.append(response)
                    
                except Exception as e:
                    print(f"Error processing email {email_id}: {str(e)}")
                    continue
            
            mail.close()
            mail.logout()
            
            return responses, None
            
        except Exception as e:
            try:
                mail.close()
                mail.logout()
            except:
                pass
            return [], f"Error fetching emails: {str(e)}"
    
    def test_connection(self):
        """Test IMAP connection"""
        mail, error = self.connect()
        if error:
            return False, error
        
        try:
            mail.select('INBOX')
            mail.close()
            mail.logout()
            return True, "IMAP connection successful"
        except Exception as e:
            return False, f"IMAP test failed: {str(e)}"
