from flask import Flask, render_template, request, jsonify, session, send_file
from werkzeug.utils import secure_filename
import os
import secrets
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
from dotenv import load_dotenv

load_dotenv()

# Database imports
try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    DB_AVAILABLE = True
except ImportError:
    print("⚠️ psycopg2 not available - database features disabled")
    DB_AVAILABLE = False

# ==================== EMAIL SENDER CLASS WITH CC/BCC ====================
class PHOCONFastEmailSender:
    def __init__(self, excel_file_path, conference_image_path, abstract_image_path, creative_image_path):
        self.excel_file_path = excel_file_path
        self.conference_image_path = conference_image_path
        self.abstract_image_path = abstract_image_path
        self.creative_image_path = creative_image_path
        
        # Gmail SMTP Configuration from environment variables
        self.smtp_config = {
            'smtp_server': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
            'smtp_port': int(os.getenv('SMTP_PORT', 587)),
            'sender_email': os.getenv('SENDER_EMAIL'),
            'sender_name': os.getenv('SENDER_NAME', 'PHOCON 2025 Team'),
            'username': os.getenv('SMTP_USERNAME'),
            'password': os.getenv('SMTP_PASSWORD'),
            'service': 'gmail',
            'security': 'starttls',
            'encryption': 'TLS',
            'use_tls': True
        }
        
        # Email templates
        self.email_templates = {
            '1': {
                'name': 'Conference Invitation Email',
                'description': 'Main conference invitation with workshop details'
            },
            '2': {
                'name': 'Abstract Submission Reminder',
                'description': 'Last call for abstract submission (10 days left)'
            },
            '3': {
                'name': 'Final Abstract Submission Reminder',
                'description': 'Final reminder for abstract submission (3 days left)'
            }
        }
        
        # Thread-safe counters
        self.successful_emails = queue.Queue()
        self.failed_emails = queue.Queue()
        self.skipped_emails = queue.Queue()
        self.selected_template = None
        
        # Performance settings
        self.max_workers = 5
        self.batch_size = 50
        self.delay_between_emails = 0.1
    
    def validate_email(self, email):
        """Email format validate karta hai"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def extract_emails_from_cell(self, cell_value):
        """Cell se multiple emails extract karta hai (comma/semicolon/newline separated)"""
        if pd.isna(cell_value) or str(cell_value).strip() == '':
            return []
        
        cell_str = str(cell_value).strip()
        # Split by comma, semicolon, or newline
        emails = re.split(r'[,;\n]+', cell_str)
        
        valid_emails = []
        for email in emails:
            email = email.strip()
            if email and self.validate_email(email):
                valid_emails.append(email)
        
        return valid_emails

    def create_conference_invitation_email(self, doctor_name):
        """Template 1: Conference invitation email content"""
        subject = "PHOCON 2025 | Meet our Esteemed International Faculty"
    
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
    
        <p style="font-size: 16px;"><strong>Dear {doctor_name}</strong></p>
    
        <p style="font-size: 14px;">Join us at the <strong>28th Annual Pediatric Hematology Oncology Conference</strong> as <strong>Dr. Michele P Lambert</strong> shares insights on <strong>Immune Thrombocytopenia (ITP)</strong>.</p>
    
        <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #007bff; margin: 20px 0;">
        <p style="margin: 0; font-size: 14px;"><strong>📅 Date:</strong> 28th – 30th November 2025</p>
        <p style="margin: 0; font-size: 14px;"><strong>📍 Venue:</strong> Dr TMA Pai Halls, KMC, Manipal</p>
        </div>
    
        <div style="text-align: center; margin: 25px 0;">
        <a href="https://followmyevent.com/phocon-2025/" style="background-color: #007bff; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-size: 16px; font-weight: bold;">
        👉 Secure your spot today!
        </a>
        </div>
    
        <p style="font-size: 14px;"><strong>For Queries:</strong> +91 63646 90353</p>
    
        <div style="text-align: center; margin: 20px 0;">
        <img src="cid:phocon_conference_image" style="max-width: 100%; height: auto; border-radius: 8px;" alt="PHOCON Conference Invitation">
        </div>
    
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
        <p style="font-size: 14px; margin: 0;">Warm Regards,</p>
        <p style="font-size: 14px; margin: 0;"><strong>Team PHOCON 2025</strong></p>
        </div>
    
        </div>
        </body>
        </html>
        """
        return subject, body
    
    def create_mahanavami_offer_email(self, doctor_name):
        """Template 2: Mahanavami special offer email content"""
        subject = "Special Mahanavami Offer – Exclusive Discounts on PHOCON 2025 Workshops!"

        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px;">

        <p style="font-size: 16px;"><strong>Dear {doctor_name}</strong></p>

        <div style="background-color: #ff6b6b; color: white; padding: 15px; text-align: center; border-radius: 8px; margin: 20px 0;">
        <h2 style="margin: 0; font-size: 24px;">🎉 Celebrate Mahanavami!</h2>
        <p style="margin: 5px 0 0 0; font-size: 16px;">Exclusive Discounted Rates on PHOCON 2025 Workshops</p>
        </div>

        <div style="background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0;">
        <p style="margin: 0; font-size: 14px;"><strong>⏰ Offer Valid:</strong> Only on 1st & 2nd October</p>
        <p style="margin: 5px 0 0 0; font-size: 14px; color: #856404;"><strong>Don't miss it!</strong></p>
        </div>

        <div style="text-align: center; margin: 30px 0;">
        <a href="https://followmyevent.com/phocon-2025/" style="background-color: #28a745; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-size: 18px; font-weight: bold; display: inline-block;">
        🚀 REGISTER NOW
        </a>
        </div>

        <div style="text-align: center; margin: 20px 0;">
        <img src="cid:phocon_abstract_image" style="max-width: 100%; height: auto; border-radius: 8px;" alt="PHOCON Mahanavami Offer">
        </div>

        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
        <p style="font-size: 14px; margin: 0;">Warm Regards,</p>
        <p style="font-size: 14px; margin: 0;"><strong>Team PHOCON 2025</strong></p>
        </div>

        </div>
        </body>
        </html>
        """

        return subject, body

    def create_final_abstract_reminder_email(self, doctor_name):
        """Template 3: Final reminder"""
        subject = "⏳ Final Reminder: Abstract Submission Closes 14th Sept!"
        
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
        
        <p style="font-size: 16px;"><strong>Dear {doctor_name},</strong></p>
        
        <div style="background-color: #dc3545; color: white; padding: 15px; text-align: center; border-radius: 8px; margin: 20px 0;">
        <h2 style="margin: 0; font-size: 24px;">🚨 Final Reminder! 🚨</h2>
        </div>
        
        <p style="font-size: 14px;">📅 Deadline: 14th Sept 2025 (Midnight)</p>
        
        <div style="text-align: center; margin: 30px 0;">
        <a href="https://phocon-conference-system.vercel.app/" style="background-color: #007bff; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-size: 18px; font-weight: bold; display: inline-block;">
        🚀 REGISTER NOW
        </a>
        </div>
        
        <div style="text-align: center; margin: 20px 0;">
        <img src="cid:phocon_creative_image" style="max-width: 100%; height: auto; border-radius: 8px;" alt="PHOCON Creative">
        </div>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
        <p style="font-size: 14px; margin: 0;">Warm regards,</p>
        <p style="font-size: 14px; margin: 0;"><strong>Team PHOCON 2025</strong></p>
        </div>
        
        </div>
        </body>
        </html>
        """
        
        return subject, body
    
    def create_email_content(self, doctor_name):
        """Selected template ke basis pe email content create karta hai"""
        if self.selected_template == '1':
            return self.create_conference_invitation_email(doctor_name)
        elif self.selected_template == '2':
            return self.create_mahanavami_offer_email(doctor_name)
        elif self.selected_template == '3':
            return self.create_final_abstract_reminder_email(doctor_name)
        else:
            raise Exception("❌ No template selected!")
    
    def create_smtp_connection(self):
        """New SMTP connection create karta hai (thread-safe)"""
        try:
            server = smtplib.SMTP(self.smtp_config['smtp_server'], self.smtp_config['smtp_port'], timeout=30)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(self.smtp_config['username'], self.smtp_config['password'])
            return server
        except Exception as e:
            raise Exception(f"SMTP connection failed: {str(e)}")
    
    def create_message_with_cc_bcc(self, recipient_email, doctor_name, cc_emails=None, bcc_emails=None):
        """
        Email message create karta hai with CC and BCC support
        
        Args:
            recipient_email: Primary recipient (TO)
            doctor_name: Name for personalization
            cc_emails: List of CC email addresses (visible to all)
            bcc_emails: List of BCC email addresses (hidden from recipients)
        """
        msg = MIMEMultipart('related')
        
        # From header
        msg['From'] = f"{self.smtp_config['sender_name']} <{self.smtp_config['sender_email']}>"
        
        # To header (primary recipient)
        msg['To'] = recipient_email
        
        # CC header (visible to all recipients)
        if cc_emails and len(cc_emails) > 0:
            msg['Cc'] = ', '.join(cc_emails)
        
        # NOTE: BCC is NOT added to headers (that's the point - it's blind/hidden)
        # BCC addresses will be included in sendmail() recipient list only
        
        # Get email subject and body
        subject, body = self.create_email_content(doctor_name)
        msg['Subject'] = subject
        
        # Attach HTML body
        msg.attach(MIMEText(body, 'html'))
        
        # Attach template-specific image
        self._attach_template_image(msg)
        
        return msg
    
    def _attach_template_image(self, msg):
        """Template ke basis pe appropriate image attach karta hai"""
        if self.selected_template == '1' and os.path.exists(self.conference_image_path):
            try:
                with open(self.conference_image_path, 'rb') as f:
                    img = MIMEImage(f.read())
                    img.add_header('Content-ID', '<phocon_conference_image>')
                    img.add_header('Content-Disposition', 'inline', 
                                 filename=os.path.basename(self.conference_image_path))
                    msg.attach(img)
            except Exception:
                pass  # Continue without image if error
        
        elif self.selected_template == '2' and os.path.exists(self.abstract_image_path):
            try:
                with open(self.abstract_image_path, 'rb') as f:
                    img = MIMEImage(f.read())
                    img.add_header('Content-ID', '<phocon_abstract_image>')
                    img.add_header('Content-Disposition', 'inline', 
                                 filename=os.path.basename(self.abstract_image_path))
                    msg.attach(img)
            except Exception:
                pass
        
        elif self.selected_template == '3' and os.path.exists(self.creative_image_path):
            try:
                with open(self.creative_image_path, 'rb') as f:
                    img = MIMEImage(f.read())
                    img.add_header('Content-ID', '<phocon_creative_image>')
                    img.add_header('Content-Disposition', 'inline', 
                                 filename=os.path.basename(self.creative_image_path))
                    msg.attach(img)
            except Exception:
                pass
    
    def send_single_email_with_cc_bcc(self, email_data):
        """
        Single email send karta hai with CC/BCC support (thread-safe)
        
        Args:
            email_data: Tuple of (to_email, name, cc_list, bcc_list, thread_id)
        """
        recipient_email, doctor_name, cc_emails, bcc_emails, thread_id = email_data
        
        try:
            # Create SMTP connection
            server = self.create_smtp_connection()
            
            # Create message with CC/BCC
            msg = self.create_message_with_cc_bcc(recipient_email, doctor_name, cc_emails, bcc_emails)
            
            # Build complete recipient list for SMTP delivery
            # SMTP needs ALL recipients (TO + CC + BCC) in the sendmail() call
            all_recipients = [recipient_email]
            if cc_emails:
                all_recipients.extend(cc_emails)
            if bcc_emails:
                all_recipients.extend(bcc_emails)
            
            # Send email to ALL recipients
            # IMPORTANT: Only TO and CC appear in email headers
            # BCC recipients get the email but are hidden from others
            server.sendmail(
                self.smtp_config['sender_email'],
                all_recipients,  # TO + CC + BCC
                msg.as_string()
            )
            server.quit()
            
            # Log success with CC/BCC info
            success_data = {
                'name': doctor_name,
                'email': recipient_email,
                'cc': ', '.join(cc_emails) if cc_emails else '',
                'bcc': ', '.join(bcc_emails) if bcc_emails else '',
                'template': self.email_templates[self.selected_template]['name'],
                'thread_id': thread_id
            }
            self.successful_emails.put(success_data)
            
            # Build log message
            cc_info = f" + CC({len(cc_emails)})" if cc_emails else ""
            bcc_info = f" + BCC({len(bcc_emails)})" if bcc_emails else ""
            return True, f"✅ [Thread-{thread_id}] Email sent to {doctor_name}{cc_info}{bcc_info}"
            
        except Exception as e:
            # Log failure with CC/BCC info
            error_data = {
                'name': doctor_name,
                'email': recipient_email,
                'cc': ', '.join(cc_emails) if cc_emails else '',
                'bcc': ', '.join(bcc_emails) if bcc_emails else '',
                'reason': str(e),
                'template': self.email_templates[self.selected_template]['name'],
                'thread_id': thread_id
            }
            self.failed_emails.put(error_data)
            
            return False, f"❌ [Thread-{thread_id}] Failed: {doctor_name} - {str(e)}"
    
    def process_excel_and_send_emails_fast(self):
        """Excel file process karta hai with CC/BCC support aur emails send karta hai"""
        try:
            print(f"📁 Reading Excel file: {self.excel_file_path}")
            df = pd.read_excel(self.excel_file_path)
            
            # Normalize column names (lowercase, trim spaces)
            df.columns = df.columns.str.lower().str.strip()
            
            # Find required columns
            name_col = None
            email_col = None
            cc_col = None
            bcc_col = None
            
            for col in df.columns:
                col_lower = col.lower()
                if 'name' in col_lower:
                    name_col = col
                # Email column (but not CC or BCC)
                if ('email' in col_lower or 'mail' in col_lower) and 'cc' not in col_lower and 'bcc' not in col_lower:
                    email_col = col
                # CC column
                if 'cc' in col_lower and 'bcc' not in col_lower:
                    cc_col = col
                # BCC column
                if 'bcc' in col_lower:
                    bcc_col = col
            
            if name_col is None or email_col is None:
                raise Exception("❌ Name or Email column not found in Excel file")
            
            print(f"✅ Found {len(df)} records")
            print(f"📝 Columns detected:")
            print(f"   Name: {name_col}")
            print(f"   Email (TO): {email_col}")
            if cc_col:
                print(f"   CC: {cc_col}")
            if bcc_col:
                print(f"   BCC: {bcc_col}")
            
            template_name = self.email_templates[self.selected_template]['name']
            print(f"📧 Using Template: {template_name}")
            print(f"⚡ Performance: {self.max_workers} concurrent threads")
            print("-" * 60)
            
            # Prepare email tasks
            email_tasks = []
            thread_counter = 0
            
            for index, row in df.iterrows():
                # Extract name
                doctor_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else f"Doctor_{index+1}"
                
                # Extract TO email(s)
                to_emails = self.extract_emails_from_cell(row[email_col])
                
                # Extract CC email(s)
                cc_emails = []
                if cc_col and cc_col in row:
                    cc_emails = self.extract_emails_from_cell(row[cc_col])
                
                # Extract BCC email(s)
                bcc_emails = []
                if bcc_col and bcc_col in row:
                    bcc_emails = self.extract_emails_from_cell(row[bcc_col])
                
                # Skip if no valid TO email
                if not to_emails:
                    self.skipped_emails.put({
                        'name': doctor_name,
                        'email': str(row[email_col]),
                        'reason': 'No valid TO email found'
                    })
                    continue
                
                # Create task for each TO email
                # (CC and BCC are shared across all TO emails from same row)
                for to_email in to_emails:
                    thread_counter += 1
                    email_tasks.append((to_email, doctor_name, cc_emails, bcc_emails, thread_counter))
            
            total_emails = len(email_tasks)
            print(f"🚀 Ready to send {total_emails} emails using {self.max_workers} threads...")
            
            # Process emails with ThreadPoolExecutor
            completed = 0
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                future_to_email = {
                    executor.submit(self.send_single_email_with_cc_bcc, task): task 
                    for task in email_tasks
                }
                
                for future in as_completed(future_to_email):
                    completed += 1
                    try:
                        success, message = future.result()
                        print(f"[{completed}/{total_emails}] {message}")
                        
                        # Small delay to avoid overwhelming SMTP server
                        if self.delay_between_emails > 0:
                            time.sleep(self.delay_between_emails)
                    except Exception as e:
                        print(f"[{completed}/{total_emails}] ❌ Exception: {str(e)}")
                    
                    # Progress update every 10 emails
                    if completed % 10 == 0:
                        progress = (completed/total_emails)*100
                        print(f"📊 Progress: {progress:.1f}% ({completed}/{total_emails})")
            
            print(f"✅ All {total_emails} email tasks completed!")
            return True
            
        except Exception as e:
            print(f"❌ Error processing Excel file: {str(e)}")
            return False

# ==================== FLASK APP ====================
app = Flask(__name__, template_folder='../templates')
app.secret_key = os.getenv('SECRET_KEY', secrets.token_hex(16))

UPLOAD_FOLDER = '/tmp/uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

DATABASE_URL = os.getenv('DATABASE_URL')

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
ALLOWED_IMAGES = {'jpg', 'jpeg', 'png'}

def get_db_connection():
    """Database connection banata hai"""
    if not DATABASE_URL or not DB_AVAILABLE:
        return None
    try:
        conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
        return conn
    except Exception as e:
        print(f"Database connection error: {e}")
        return None

def log_to_database(campaign_id, recipient_name, recipient_email, template, status, 
                    error_msg=None, thread_id=None, cc_recipients=None, bcc_recipients=None):
    """
    Email status database mein log karta hai with CC/BCC support
    
    Args:
        campaign_id: Campaign ID
        recipient_name: Recipient name
        recipient_email: Primary email (TO)
        template: Template ID used
        status: 'sent', 'failed', or 'skipped'
        error_msg: Error message if failed
        thread_id: Thread ID for tracking
        cc_recipients: Comma-separated CC emails
        bcc_recipients: Comma-separated BCC emails
    """
    if not DB_AVAILABLE:
        return
    
    conn = get_db_connection()
    if not conn:
        return
    
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO email_logs 
            (campaign_id, recipient_name, recipient_email, template_used, status, 
             error_message, thread_id, cc_recipients, bcc_recipients, sent_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, CURRENT_TIMESTAMP)
        """, (campaign_id, recipient_name, recipient_email, template, status, 
              error_msg, thread_id, cc_recipients, bcc_recipients))
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Database logging error: {e}")

def create_campaign(campaign_name, template_id, performance_mode, total_recipients, excel_filename):
    """Naya campaign database mein create karta hai"""
    if not DB_AVAILABLE:
        return None
    
    conn = get_db_connection()
    if not conn:
        return None
    
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO campaigns 
            (campaign_name, template_id, performance_mode, total_recipients, excel_filename, status)
            VALUES (%s, %s, %s, %s, %s, 'running')
            RETURNING id
        """, (campaign_name, template_id, performance_mode, total_recipients, excel_filename))
        campaign_id = cursor.fetchone()['id']
        conn.commit()
        cursor.close()
        conn.close()
        return campaign_id
    except Exception as e:
        print(f"Campaign creation error: {e}")
        return None

def update_campaign_status(campaign_id, emails_sent, emails_failed, status='completed'):
    """Campaign status update karta hai"""
    if not DB_AVAILABLE:
        return
    
    conn = get_db_connection()
    if not conn or not campaign_id:
        return
    
    try:
        total = emails_sent + emails_failed
        success_rate = (emails_sent / total * 100) if total > 0 else 0
        
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE campaigns 
            SET emails_sent = %s, 
                emails_failed = %s, 
                success_rate = %s,
                status = %s,
                completed_at = CURRENT_TIMESTAMP
            WHERE id = %s
        """, (emails_sent, emails_failed, success_rate, status, campaign_id))
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Campaign update error: {e}")

def log_file_upload(filename, file_type, file_path, session_id):
    """File upload log karta hai"""
    if not DB_AVAILABLE:
        return
    
    conn = get_db_connection()
    if not conn:
        return
    
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO uploaded_files (filename, file_type, file_path, session_id)
            VALUES (%s, %s, %s, %s)
        """, (filename, file_type, file_path, session_id))
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"File logging error: {e}")

def allowed_file(filename, allowed_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/health')
def health():
    """Health check with database status"""
    db_status = "not_configured"
    if DATABASE_URL and DB_AVAILABLE:
        conn = get_db_connection()
        if conn:
            db_status = "connected"
            conn.close()
        else:
            db_status = "error"
    
    return jsonify({
        'status': 'healthy',
        'platform': 'vercel',
        'database': db_status,
        'psycopg2_available': DB_AVAILABLE,
        'timestamp': datetime.now().isoformat()
    })

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': 'Excel file is required'}), 400
        
        excel_file = request.files['excel_file']
        conference_img = request.files.get('conference_image')
        abstract_img = request.files.get('abstract_image')
        creative_img = request.files.get('creative_image')
        
        if excel_file.filename == '':
            return jsonify({'error': 'No Excel file selected'}), 400
        
        if not allowed_file(excel_file.filename, ALLOWED_EXTENSIONS):
            return jsonify({'error': 'Invalid Excel file format'}), 400
        
        # Save files
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
        excel_file.save(excel_path)
        
        conference_path = None
        abstract_path = None
        creative_path = None
        
        if conference_img and allowed_file(conference_img.filename, ALLOWED_IMAGES):
            conference_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(conference_img.filename))
            conference_img.save(conference_path)
        
        if abstract_img and allowed_file(abstract_img.filename, ALLOWED_IMAGES):
            abstract_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(abstract_img.filename))
            abstract_img.save(abstract_path)
        
        if creative_img and allowed_file(creative_img.filename, ALLOWED_IMAGES):
            creative_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(creative_img.filename))
            creative_img.save(creative_path)
        
        # Store in session
        session['excel_path'] = excel_path
        session['excel_filename'] = excel_file.filename
        session['conference_path'] = conference_path or ''
        session['abstract_path'] = abstract_path or ''
        session['creative_path'] = creative_path or ''
        
        # Log to database
        session_id = session.get('session_id', secrets.token_hex(8))
        session['session_id'] = session_id
        
        log_file_upload(excel_file.filename, 'excel', excel_path, session_id)
        if conference_path:
            log_file_upload(os.path.basename(conference_path), 'image', conference_path, session_id)
        
        return jsonify({
            'success': True,
            'message': 'Files uploaded and logged successfully'
        })
    
    except Exception as e:
        print(f"Upload error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/send_emails', methods=['POST'])
def send_emails():
    campaign_id = None
    try:
        data = request.json
        template = data.get('template')
        performance_mode = data.get('performance_mode')
        
        if not template or not performance_mode:
            return jsonify({'error': 'Template and performance mode required'}), 400
        
        excel_path = session.get('excel_path')
        excel_filename = session.get('excel_filename', 'unknown.xlsx')
        conference_path = session.get('conference_path', '')
        abstract_path = session.get('abstract_path', '')
        creative_path = session.get('creative_path', '')
        
        if not excel_path or not os.path.exists(excel_path):
            return jsonify({'error': 'Excel file not found'}), 400
        
        # Count recipients
        df = pd.read_excel(excel_path)
        total_recipients = len(df)
        
        # Create campaign in database
        campaign_name = f"Campaign_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        campaign_id = create_campaign(
            campaign_name, 
            template, 
            performance_mode, 
            total_recipients,
            excel_filename
        )
        
        # Create email sender
        email_sender = PHOCONFastEmailSender(
            excel_path,
            conference_path,
            abstract_path,
            creative_path
        )
        
        email_sender.selected_template = template
        
        # Performance settings
        performance_settings = {
            '1': {'workers': 1, 'delay': 0.5},
            '2': {'workers': 5, 'delay': 0.1},
            '3': {'workers': 8, 'delay': 0.05},
            '4': {'workers': 10, 'delay': 0.02}
        }
        
        settings = performance_settings.get(performance_mode)
        email_sender.max_workers = settings['workers']
        email_sender.delay_between_emails = settings['delay']
        
        # Send emails
        success = email_sender.process_excel_and_send_emails_fast()
        
        # Collect results with CC/BCC logging
        successful_list = []
        failed_list = []
        skipped_list = []
        
        while not email_sender.successful_emails.empty():
            email_data = email_sender.successful_emails.get()
            successful_list.append(email_data)
            log_to_database(
                campaign_id,
                email_data.get('name'),
                email_data.get('email'),
                template,
                'sent',
                thread_id=email_data.get('thread_id'),
                cc_recipients=email_data.get('cc'),  # NEW: Log CC
                bcc_recipients=email_data.get('bcc')  # NEW: Log BCC
            )
        
        while not email_sender.failed_emails.empty():
            email_data = email_sender.failed_emails.get()
            failed_list.append(email_data)
            log_to_database(
                campaign_id,
                email_data.get('name'),
                email_data.get('email'),
                template,
                'failed',
                error_msg=email_data.get('reason'),
                thread_id=email_data.get('thread_id'),
                cc_recipients=email_data.get('cc'),  # NEW: Log CC
                bcc_recipients=email_data.get('bcc')  # NEW: Log BCC
            )
            
        while not email_sender.skipped_emails.empty():
            email_data = email_sender.skipped_emails.get()
            skipped_list.append(email_data)
            log_to_database(
                campaign_id,
                email_data.get('name'),
                email_data.get('email'),
                template,
                'skipped',
                error_msg=email_data.get('reason')
            )
        
        # Update campaign status
        update_campaign_status(
            campaign_id,
            len(successful_list),
            len(failed_list) + len(skipped_list)
        )
        
        # Generate reports
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_files = []
        
        if successful_list:
            success_df = pd.DataFrame(successful_list)
            success_file = f"successful_emails_template{template}_{timestamp}.xlsx"
            success_path = os.path.join(app.config['UPLOAD_FOLDER'], success_file)
            success_df.to_excel(success_path, index=False)
            report_files.append({
                'type': 'success',
                'filename': success_file,
                'count': len(successful_list)
            })
        
        if failed_list or skipped_list:
            failed_df = pd.DataFrame(failed_list + skipped_list)
            failed_file = f"failed_emails_template{template}_{timestamp}.xlsx"
            failed_path = os.path.join(app.config['UPLOAD_FOLDER'], failed_file)
            failed_df.to_excel(failed_path, index=False)
            report_files.append({
                'type': 'failed',
                'filename': failed_file,
                'count': len(failed_list) + len(skipped_list)
            })
        
        total_attempts = len(successful_list) + len(failed_list)
        success_rate = (len(successful_list) / total_attempts * 100) if total_attempts > 0 else 0
        
        return jsonify({
            'success': success,
            'total_sent': len(successful_list),
            'total_failed': len(failed_list) + len(skipped_list),
            'success_rate': success_rate,
            'reports': report_files,
            'campaign_id': campaign_id
        })
    
    except Exception as e:
        print(f"Error in send_emails: {str(e)}")
        import traceback
        traceback.print_exc()
        
        if campaign_id:
            update_campaign_status(campaign_id, 0, 0, status='failed')
        
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_report(filename):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        print(f"Download error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/download_template')
def download_template():
    """Sample Excel template with CC/BCC columns download karta hai"""
    try:
        # Create sample DataFrame with CC/BCC examples
        sample_data = {
            'Name': [
                'Dr. Azad Singh', 
                'Dr. Aman Kumar', 
                'Dr. Priya Shah',
                'Dr. Rajesh Verma'
            ],
            'Email': [
                'azad@hospital.com', 
                'aman@kmc.edu', 
                'priya@aiims.in',
                'rajesh@sgpgi.ac.in'
            ],
            'CC': [
                'secretary@hospital.com', 
                'head@kmc.edu; dept@kmc.edu', 
                '',
                'team@sgpgi.ac.in'
            ],
            'BCC': [
                'admin@phocon2025.com', 
                '', 
                'tracking@phocon2025.com',
                'admin@phocon2025.com; analytics@phocon2025.com'
            ]
        }
        
        df = pd.DataFrame(sample_data)
        
        # Add instructions as a second sheet
        instructions_data = {
            'Column': ['Name', 'Email', 'CC', 'BCC'],
            'Required': ['Yes', 'Yes', 'No', 'No'],
            'Description': [
                'Recipient name for personalization',
                'Primary recipient email (TO field)',
                'Carbon copy - visible to all recipients (separate multiple with semicolon)',
                'Blind carbon copy - hidden from other recipients (separate multiple with semicolon)'
            ],
            'Example': [
                'Dr. John Doe',
                'john@hospital.com',
                'secretary@hospital.com; assistant@hospital.com',
                'admin@phocon2025.com'
            ]
        }
        
        instructions_df = pd.DataFrame(instructions_data)
        
        # Create Excel file with multiple sheets
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], 'PHOCON_2025_Template.xlsx')
        
        with pd.ExcelWriter(template_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Recipients', index=False)
            instructions_df.to_excel(writer, sheet_name='Instructions', index=False)
        
        return send_file(
            template_path,
            as_attachment=True,
            download_name='PHOCON_2025_Recipients_Template.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        print(f"Template download error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/campaigns')
def get_campaigns():
    """Saari campaigns list karta hai"""
    if not DB_AVAILABLE:
        return jsonify({'error': 'Database not available'}), 500
    
    conn = get_db_connection()
    if not conn:
        return jsonify({'error': 'Database not configured'}), 500
    
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, campaign_name, template_id, performance_mode, 
                   status, total_recipients, emails_sent, emails_failed, 
                   success_rate, created_at, completed_at
            FROM campaigns
            ORDER BY created_at DESC
            LIMIT 50
        """)
        
        campaigns = cursor.fetchall()
        cursor.close()
        conn.close()
        
        return jsonify({'campaigns': campaigns})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Vercel ke liye
app = app