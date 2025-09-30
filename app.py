from flask import Flask, render_template, request, jsonify, session, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
from phocon_email_sender import PHOCONFastEmailSender
from datetime import datetime
import secrets
import pandas as pd

app = Flask(__name__)
CORS(app)  # Enable CORS
app.secret_key = secrets.token_hex(16)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create uploads folder
os.makedirs('uploads', exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
ALLOWED_IMAGES = {'jpg', 'jpeg', 'png'}

def allowed_file(filename, allowed_set):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_set

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        # Check if files are present
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
        
        # Save images if provided
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
        
        # Store paths in session
        session['excel_path'] = excel_path
        session['conference_path'] = conference_path or ''
        session['abstract_path'] = abstract_path or ''
        session['creative_path'] = creative_path or ''
        
        return jsonify({
            'success': True,
            'message': 'Files uploaded successfully'
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/send_emails', methods=['POST'])
def send_emails():
    try:
        data = request.json
        template = data.get('template')
        performance_mode = data.get('performance_mode')
        
        if not template or not performance_mode:
            return jsonify({'error': 'Template and performance mode required'}), 400
        
        # Get file paths from session
        excel_path = session.get('excel_path')
        conference_path = session.get('conference_path', '')
        abstract_path = session.get('abstract_path', '')
        creative_path = session.get('creative_path', '')
        
        if not excel_path or not os.path.exists(excel_path):
            return jsonify({'error': 'Excel file not found'}), 400
        
        # Create email sender
        email_sender = PHOCONFastEmailSender(
            excel_path,
            conference_path,
            abstract_path,
            creative_path
        )
        
        # Set template and performance mode
        email_sender.selected_template = template
        
        performance_settings = {
            '1': {'workers': 1, 'delay': 0.5},
            '2': {'workers': 5, 'delay': 0.1},
            '3': {'workers': 10, 'delay': 0.05},
            '4': {'workers': 15, 'delay': 0.02}
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
            successful_list.append(email_sender.successful_emails.get())
        
        while not email_sender.failed_emails.empty():
            failed_list.append(email_sender.failed_emails.get())
            
        while not email_sender.skipped_emails.empty():
            skipped_list.append(email_sender.skipped_emails.get())
        
        # Generate Excel reports with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_files = []
        
        # Save successful emails report
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
            print(f"✅ Success report created: {success_file}")
        
        # Save failed emails report
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
            print(f"❌ Failed report created: {failed_file}")
        
        total_attempts = len(successful_list) + len(failed_list)
        success_rate = (len(successful_list) / total_attempts * 100) if total_attempts > 0 else 0
        
        print(f"📊 Reports generated: {len(report_files)} files")
        for report in report_files:
            print(f"   - {report['filename']} ({report['count']} records)")
        
        return jsonify({
            'success': success,
            'total_sent': len(successful_list),
            'total_failed': len(failed_list) + len(skipped_list),
            'success_rate': success_rate,
            'reports': report_files
        })
    
    except Exception as e:
        print(f"❌ Error in send_emails: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_report(filename):
    """Download generated report files"""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Security check
        if not os.path.exists(file_path):
            print(f"❌ File not found: {file_path}")
            return jsonify({'error': 'File not found'}), 404
        
        # Check if file is in allowed directory
        abs_upload_folder = os.path.abspath(app.config['UPLOAD_FOLDER'])
        abs_file_path = os.path.abspath(file_path)
        
        if not abs_file_path.startswith(abs_upload_folder):
            print("⚠️ Security: Attempted access outside upload folder")  
            return jsonify({'error': 'Invalid file path'}), 403
        
        print(f"📥 Downloading: {filename}")
        
        # Use send_from_directory for better security and compatibility
        return send_from_directory(
            app.config['UPLOAD_FOLDER'],
            filename,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        print(f"❌ Download error: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    print("=" * 70)
    print("🚀 PHOCON 2025 Campaign Control Center")
    print("=" * 70)
    print("📁 Upload folder:", os.path.abspath(app.config['UPLOAD_FOLDER']))
    print("🌐 Server: http://localhost:5000")
    print("=" * 70)
    app.run(debug=True, port=5000)