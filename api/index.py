from flask import Flask, render_template, request, jsonify, session, send_file
from werkzeug.utils import secure_filename
import os
import secrets
import pandas as pd
from datetime import datetime
import sys
import psycopg2
from psycopg2.extras import RealDictCursor

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from phocon_email_sender import PHOCONFastEmailSender

app = Flask(__name__, template_folder='../templates')
app.secret_key = os.getenv('SECRET_KEY', secrets.token_hex(16))

# Use /tmp for Vercel
UPLOAD_FOLDER = '/tmp/uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

# Database configuration
DATABASE_URL = os.getenv('DATABASE_URL')

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
ALLOWED_IMAGES = {'jpg', 'jpeg', 'png'}

def get_db_connection():
    """Database connection banata hai"""
    if not DATABASE_URL:
        return None
    try:
        conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
        return conn
    except Exception as e:
        print(f"Database connection error: {e}")
        return None

def log_to_database(campaign_id, recipient_name, recipient_email, template, status, error_msg=None, thread_id=None):
    """Email status database mein log karta hai"""
    conn = get_db_connection()
    if not conn:
        return
    
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO email_logs 
            (campaign_id, recipient_name, recipient_email, template_used, status, error_message, thread_id, sent_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, CURRENT_TIMESTAMP)
        """, (campaign_id, recipient_name, recipient_email, template, status, error_msg, thread_id))
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Database logging error: {e}")

def create_campaign(campaign_name, template_id, performance_mode, total_recipients, excel_filename):
    """Naya campaign database mein create karta hai"""
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
    if DATABASE_URL:
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
        
        # Collect results
        successful_list = []
        failed_list = []
        skipped_list = []
        
        while not email_sender.successful_emails.empty():
            email_data = email_sender.successful_emails.get()
            successful_list.append(email_data)
            # Log to database
            log_to_database(
                campaign_id,
                email_data.get('name'),
                email_data.get('email'),
                template,
                'sent',
                thread_id=email_data.get('thread_id')
            )
        
        while not email_sender.failed_emails.empty():
            email_data = email_sender.failed_emails.get()
            failed_list.append(email_data)
            # Log to database
            log_to_database(
                campaign_id,
                email_data.get('name'),
                email_data.get('email'),
                template,
                'failed',
                error_msg=email_data.get('reason'),
                thread_id=email_data.get('thread_id')
            )
            
        while not email_sender.skipped_emails.empty():
            email_data = email_sender.skipped_emails.get()
            skipped_list.append(email_data)
            # Log to database
            log_to_database(
                campaign_id,
                email_data.get('name'),
                email_data.get('email'),
                template,
                'skipped',
                error_msg=email_data.get('reason')
            )
        
        # Update campaign status in database
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

@app.route('/campaigns')
def get_campaigns():
    """Saari campaigns list karta hai"""
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