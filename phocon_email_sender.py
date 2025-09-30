import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import re
import os
from datetime import datetime
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue
from dotenv import load_dotenv  # Only import once from correct module

load_dotenv()  # Load environment variables

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
        self.max_workers = 5  # Concurrent threads (Gmail limit safe)
        self.batch_size = 50  # Process in batches
        self.delay_between_emails = 0.1  # Reduced delay (0.1 seconds)
        
    def display_performance_options(self):
        """Performance options display karta hai"""
        print("\n⚡ PERFORMANCE OPTIONS:")
        print("="*50)
        print("1. 🐌 SAFE MODE     - 1 email at a time (slow but safe)")
        print("2. 🚀 FAST MODE     - 5 concurrent threads (5x faster)")
        print("3. 🏎  TURBO MODE    - 10 concurrent threads (10x faster)")
        print("4. 🔥 BEAST MODE    - 15 concurrent threads (15x faster)")
        print("-"*50)
        print("⚠  NOTE: Higher modes may hit Gmail rate limits")
        
    def select_performance_mode(self):
        """Performance mode select karne ke liye"""
        while True:
            try:
                self.display_performance_options()
                choice = input("\n🎯 Select performance mode (1-4): ").strip()
                
                if choice == '1':
                    self.max_workers = 1
                    self.delay_between_emails = 0.5
                    print("✅ Selected: SAFE MODE (1 thread, 0.5s delay)")
                    return True
                elif choice == '2':
                    self.max_workers = 5
                    self.delay_between_emails = 0.1
                    print("✅ Selected: FAST MODE (5 threads, 0.1s delay)")
                    return True
                elif choice == '3':
                    self.max_workers = 10
                    self.delay_between_emails = 0.05
                    print("✅ Selected: TURBO MODE (10 threads, 0.05s delay)")
                    return True
                elif choice == '4':
                    self.max_workers = 15
                    self.delay_between_emails = 0.02
                    print("✅ Selected: BEAST MODE (15 threads, 0.02s delay)")
                    print("⚠  Warning: May hit Gmail limits!")
                    return True
                else:
                    print("❌ Invalid choice! Please select 1-4.")
                    
            except KeyboardInterrupt:
                print("\n❌ Selection cancelled.")
                return False
    
    def display_email_templates(self):
        """Available email templates display karta hai"""
        print("\n📧 AVAILABLE EMAIL TEMPLATES:")
        print("="*60)
        for key, template in self.email_templates.items():
            print(f"{key}. {template['name']}")
            print(f"   📝 {template['description']}")
            print("-"*40)
        
    def select_email_template(self):
        """User se email template select karane ke liye"""
        while True:
            try:
                self.display_email_templates()
                choice = input("\n🎯 Select email template (1, 2, or 3): ").strip()
                
                if choice in self.email_templates:
                    self.selected_template = choice
                    selected_name = self.email_templates[choice]['name']
                    print(f"✅ Selected: {selected_name}")
                    return True
                else:
                    print("❌ Invalid choice! Please select 1, 2, or 3.")
                    
            except KeyboardInterrupt:
                print("\n❌ Selection cancelled.")
                return False
    
    def validate_email(self, email):
        """Email format validate karta hai"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def extract_emails_from_cell(self, cell_value):
        """Cell se multiple emails extract karta hai"""
        if pd.isna(cell_value) or str(cell_value).strip() == '':
            return []
        
        cell_str = str(cell_value).strip()
        emails = re.split(r'[,;\s\n]+', cell_str)
        
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
    
    def create_abstract_submission_email(self, doctor_name):
        """Template 2: Abstract submission reminder email content"""
        subject = "PHOCON 2025 Abstracts – Last Call, Hurry Up...!"
        
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
        
        <p style="font-size: 16px;"><strong>Dear {doctor_name}</strong></p>
        
        <div style="background-color: #ff6b6b; color: white; padding: 15px; text-align: center; border-radius: 8px; margin: 20px 0;">
        <h2 style="margin: 0; font-size: 24px;">⏰ ONLY 10 DAYS LEFT!</h2>
        <p style="margin: 5px 0 0 0; font-size: 16px;">Don't miss this opportunity!</p>
        </div>
        
        <p style="font-size: 14px;">Submit your abstract for <strong>PHOCON 2025</strong> (28–30 November, Manipal).</p>
        
        <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #28a745; margin: 20px 0;">
        <p style="margin: 0; font-size: 14px;"><strong>🏆 Showcase your research & gain recognition!</strong></p>
        <p style="margin: 5px 0 0 0; font-size: 14px;">Join leading experts in Pediatric Hematology & Oncology</p>
        </div>
        
        <div style="background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0;">
        <p style="margin: 0; font-size: 14px;"><strong>📅 Conference Dates:</strong> 28–30 November 2025</p>
        <p style="margin: 5px 0 0 0; font-size: 14px;"><strong>📍 Venue:</strong> Kasturba Medical College, Manipal</p>
        </div>
        
        <div style="text-align: center; margin: 30px 0;">
        <a href="https://phocon-conference-system.vercel.app/" style="background-color: #28a745; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-size: 18px; font-weight: bold; display: inline-block;">
        🚀 SUBMIT NOW
        </a>
        </div>
        
        <div style="background-color: #d1ecf1; padding: 15px; border-left: 4px solid #17a2b8; margin: 20px 0;">
        <p style="margin: 0; font-size: 14px;"><strong>🎯 Why Submit Your Abstract?</strong></p>
        <ul style="font-size: 13px; margin: 10px 0 0 0; padding-left: 20px;">
        <li>Present to leading pediatric specialists</li>
        <li>Network with experts in your field</li>
        <li>Get published in conference proceedings</li>
        <li>Win recognition awards</li>
        </ul>
        </div>
        
        <div style="text-align: center; margin: 20px 0;">
        <img src="cid:phocon_abstract_image" style="max-width: 100%; height: auto; border-radius: 8px;" alt="PHOCON Abstract Submission">
        </div>
        
        <div style="text-align: center; background-color: #f8d7da; padding: 15px; border-radius: 8px; margin: 20px 0;">
        <p style="margin: 0; font-size: 16px; color: #721c24;"><strong>⚠ DEADLINE APPROACHING FAST!</strong></p>
        <p style="margin: 5px 0 0 0; font-size: 14px; color: #721c24;">Submit before it's too late!</p>
        </div>
        
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee;">
        <p style="font-size: 14px; margin: 0;">Warm Regards,</p>
        <p style="font-size: 14px; margin: 0;"><strong>Team PHOCON 2025</strong></p>
        <p style="font-size: 12px; color: #666; margin: 10px 0 0 0;">Kasturba Medical College, Manipal</p>
        </div>
        
        </div>
        </body>
        </html>
        """
        
        return subject, body

    def create_final_abstract_reminder_email(self, doctor_name):
        """Template 3: Early Bird Ends Soon"""
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
            return self.create_abstract_submission_email(doctor_name)
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
    
    def send_single_email(self, email_data):
        """Single email send karta hai (thread-safe)"""
        recipient_email, doctor_name, thread_id = email_data
        
        try:
            # Create new SMTP connection for this thread
            server = self.create_smtp_connection()
            
            # Create message
            msg = self.create_message(recipient_email, doctor_name)
            
            # Send email
            text = msg.as_string()
            server.sendmail(self.smtp_config['sender_email'], recipient_email, text)
            server.quit()
            
            # Thread-safe logging
            success_data = {
                'name': doctor_name,
                'email': recipient_email,
                'template': self.email_templates[self.selected_template]['name'],
                'thread_id': thread_id
            }
            self.successful_emails.put(success_data)
            
            return True, f"✅ [Thread-{thread_id}] Email sent to {doctor_name} ({recipient_email})"
            
        except Exception as e:
            # Thread-safe error logging
            error_data = {
                'name': doctor_name,
                'email': recipient_email,
                'reason': str(e),
                'template': self.email_templates[self.selected_template]['name'],
                'thread_id': thread_id
            }
            self.failed_emails.put(error_data)
            
            return False, f"❌ [Thread-{thread_id}] Failed to send to {doctor_name} ({recipient_email}): {str(e)}"
    
    def create_message(self, recipient_email, doctor_name):
        """Email message create karta hai"""
        msg = MIMEMultipart('related')
        msg['From'] = f"{self.smtp_config['sender_name']} <{self.smtp_config['sender_email']}>"
        msg['To'] = recipient_email
        
        subject, body = self.create_email_content(doctor_name)
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'html'))
        
        # Template ke basis pe different images attach karte hain
        if self.selected_template == '1':
            if os.path.exists(self.conference_image_path):
                try:
                    with open(self.conference_image_path, 'rb') as attachment:
                        img = MIMEImage(attachment.read())
                        img.add_header('Content-ID', '<phocon_conference_image>')
                        img.add_header('Content-Disposition', 'inline', 
                                     filename=os.path.basename(self.conference_image_path))
                        msg.attach(img)
                except Exception:
                    pass  # Continue without image if error
        
        elif self.selected_template == '2':
            if os.path.exists(self.abstract_image_path):
                try:
                    with open(self.abstract_image_path, 'rb') as attachment:
                        img = MIMEImage(attachment.read())
                        img.add_header('Content-ID', '<phocon_abstract_image>')
                        img.add_header('Content-Disposition', 'inline', 
                                     filename=os.path.basename(self.abstract_image_path))
                        msg.attach(img)
                except Exception:
                    pass  # Continue without image if error
        
        elif self.selected_template == '3':
            if os.path.exists(self.creative_image_path):
                try:
                    with open(self.creative_image_path, 'rb') as attachment:
                        img = MIMEImage(attachment.read())
                        img.add_header('Content-ID', '<phocon_creative_image>')
                        img.add_header('Content-Disposition', 'inline', 
                                     filename=os.path.basename(self.creative_image_path))
                        msg.attach(img)
                except Exception:
                    pass  # Continue without image if error
        
        return msg
    
    def test_smtp_connection(self):
        """Working SMTP configuration test karta hai"""
        print("🧪 Testing Gmail SMTP connection...")
        print(f"   📧 Email: {self.smtp_config['sender_email']}")
        print(f"   🖥  Server: {self.smtp_config['smtp_server']}:{self.smtp_config['smtp_port']}")
        
        try:
            server = self.create_smtp_connection()
            server.quit()
            print("   ✅ Connection test: PASSED")
            return True
        except Exception as e:
            print(f"   ❌ Connection test: FAILED - {str(e)}")
            return False
    
    def process_excel_and_send_emails_fast(self):
        """Excel file process karta hai aur emails send karta hai (FAST VERSION)"""
        try:
            print(f"📁 Reading Excel file: {self.excel_file_path}")
            df = pd.read_excel(self.excel_file_path)
            
            df.columns = df.columns.str.lower().str.strip()
            
            name_col = None
            email_col = None
            
            for col in df.columns:
                if 'name' in col:
                    name_col = col
                if 'email' in col or 'mail' in col:
                    email_col = col
            
            if name_col is None or email_col is None:
                raise Exception("❌ Name or Email column not found in Excel file")
            
            print(f"✅ Found {len(df)} records")
            print(f"📝 Name column: {name_col}")
            print(f"📧 Email column: {email_col}")
            
            template_name = self.email_templates[self.selected_template]['name']
            print(f"📧 Using Template: {template_name}")
            print(f"⚡ Performance: {self.max_workers} concurrent threads")
            print("-" * 60)
            
            # Prepare email list
            email_tasks = []
            thread_counter = 0
            
            for index, row in df.iterrows():
                doctor_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else f"Doctor_{index+1}"
                email_cell = row[email_col]
                
                emails = self.extract_emails_from_cell(email_cell)
                
                if not emails:
                    skip_data = {
                        'name': doctor_name,
                        'email': str(email_cell),
                        'reason': 'No valid email found'
                    }
                    self.skipped_emails.put(skip_data)
                    continue
                
                for email in emails:
                    thread_counter += 1
                    email_tasks.append((email, doctor_name, thread_counter))
            
            total_emails = len(email_tasks)
            print(f"🚀 Ready to send {total_emails} emails using {self.max_workers} threads...")
            
            # Process emails with ThreadPoolExecutor
            completed = 0
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # Submit all tasks
                future_to_email = {executor.submit(self.send_single_email, task): task for task in email_tasks}
                
                # Process completed tasks
                for future in as_completed(future_to_email):
                    completed += 1
                    email_data = future_to_email[future]
                    recipient_email, doctor_name, thread_id = email_data
                    
                    try:
                        success, message = future.result()
                        print(f"[{completed}/{total_emails}] {message}")
                        
                        # Small delay to avoid overwhelming Gmail
                        if self.delay_between_emails > 0:
                            time.sleep(self.delay_between_emails)
                            
                    except Exception as e:
                        print(f"[{completed}/{total_emails}] ❌ [Thread-{thread_id}] Exception: {str(e)}")
                    
                    # Progress update every 10 emails
                    if completed % 10 == 0:
                        progress = (completed/total_emails)*100
                        print(f"📊 Progress: {progress:.1f}% ({completed}/{total_emails})")
            
            print(f"✅ All {total_emails} email tasks completed!")
            
        except Exception as e:
            print(f"❌ Error processing Excel file: {str(e)}")
            return False
        
        return True
    
    def generate_report(self):
        """Complete report generate karta hai (thread-safe version)"""
        print("\n" + "="*70)
        print("📊 PHOCON 2025 FAST EMAIL CAMPAIGN REPORT")
        print("="*70)
        
        # Convert queues to lists for counting
        successful_list = []
        failed_list = []
        skipped_list = []
        
        while not self.successful_emails.empty():
            successful_list.append(self.successful_emails.get())
        
        while not self.failed_emails.empty():
            failed_list.append(self.failed_emails.get())
            
        while not self.skipped_emails.empty():
            skipped_list.append(self.skipped_emails.get())
        
        total_attempts = len(successful_list) + len(failed_list)
        template_name = self.email_templates[self.selected_template]['name']
        
        print(f"📧 Template Used: {template_name}")
        print(f"⚡ Performance Mode: {self.max_workers} concurrent threads")
        print(f"📈 Total Email Attempts: {total_attempts}")
        print(f"✅ Successful Emails: {len(successful_list)}")
        print(f"❌ Failed Emails: {len(failed_list)}")
        print(f"⏭  Skipped Records: {len(skipped_list)}")
        
        if total_attempts > 0:
            success_rate = (len(successful_list)/total_attempts*100)
            print(f"🎯 Success Rate: {success_rate:.2f}%")
        else:
            print("🎯 Success Rate: 0%")
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if failed_list or skipped_list:
            failed_df = pd.DataFrame(failed_list + skipped_list)
            failed_file = f"failed_emails_fast_template{self.selected_template}_{timestamp}.xlsx"
            failed_df.to_excel(failed_file, index=False)
            print(f"\n💾 Failed emails saved to: {failed_file}")
        
        if successful_list:
            success_df = pd.DataFrame(successful_list)
            success_file = f"successful_emails_fast_template{self.selected_template}_{timestamp}.xlsx"
            success_df.to_excel(success_file, index=False)
            print(f"💾 Successful emails saved to: {success_file}")
        
        # Return both lists for use by app.py
        return successful_list, failed_list

# ==================== COMMENTED OUT FOR WEB DEPLOYMENT ====================
# The main() function below is commented out because it's designed for
# command-line usage with hardcoded Windows file paths. When deploying to
# Render or other cloud platforms, app.py handles all the file uploads and
# email sending through the web interface instead.
# ==========================================================================

# def main():
#     print("="*70)
#     print("🚀 PHOCON 2025 FAST BULK EMAIL SENDER")
#     print("⚡ Multi-threaded • High Performance • Gmail Safe")
#     print("📧 Email: admin@phocon2025.com")
#     print("="*70)
#     
#     # File paths
#     excel_path = r"C:\Users\Azad Singh\OneDrive\Desktop\pythone\PHO_List-with email id.xlsx"
#     conference_image_path = r"C:\Users\Azad Singh\OneDrive\Desktop\pythone\PHOCON  Workshop Creative.jpeg"
#     abstract_image_path = r"C:\Users\Azad Singh\Downloads\PHOCON Abstract Submission.jpeg"
#     creative_image_path = r"C:\Users\Azad Singh\OneDrive\Desktop\pythone\Creative.jpeg"
#     print("\n📁 Step 1: File validation...")
#     if not os.path.exists(excel_path):
#         print(f"❌ Excel file not found: {excel_path}")
#         return
#     
#     print(f"✅ Excel file found: {excel_path}")
#     
#     if os.path.exists(conference_image_path):
#         print(f"✅ Conference image found: {os.path.basename(conference_image_path)}")
#     else:
#         print(f"⚠  Conference image not found: {os.path.basename(conference_image_path)}")
#     
#     if os.path.exists(abstract_image_path):
#         print(f"✅ Abstract image found: {os.path.basename(abstract_image_path)}")
#     else:
#         print(f"⚠  Abstract image not found: {os.path.basename(abstract_image_path)}")
#     
#     if os.path.exists(creative_image_path):
#         print(f"✅ Creative image found: {os.path.basename(creative_image_path)}")
#     else:
#         print(f"⚠  Creative image not found: {os.path.basename(creative_image_path)}")
#     
#     ## Create email sender instance
#     email_sender = PHOCONFastEmailSender(excel_path, conference_image_path, abstract_image_path, creative_image_path)
#     
#     print("\n📧 Step 2: Select Email Template")
#     if not email_sender.select_email_template():
#         print("❌ Template selection cancelled.")
#         return
#     
#     print("\n⚡ Step 3: Select Performance Mode")
#     if not email_sender.select_performance_mode():
#         print("❌ Performance mode selection cancelled.")
#         return
#     
#     print("\n🔍 Step 4: Testing SMTP connection...")
#     if not email_sender.test_smtp_connection():
#         print("❌ SMTP connection test failed!")
#         return
#     
#     print(f"\n⚙  Step 5: Configuration Summary")
#     selected_template = email_sender.email_templates[email_sender.selected_template]
#     print(f"📧 Template: {selected_template['name']}")
#     print(f"⚡ Performance: {email_sender.max_workers} concurrent threads")
#     print(f"⏱  Delay: {email_sender.delay_between_emails}s between emails")
#     
#     if email_sender.selected_template == '1':
#         print(f"🖼  Image: {os.path.basename(conference_image_path)} (Conference)")
#     elif email_sender.selected_template == '2':
#         print(f"🖼  Image: {os.path.basename(abstract_image_path)} (Abstract)")
#     else:
#         print(f"🖼  Image: {os.path.basename(creative_image_path)} (Creative)")
#     
#     print("📧 Sender: admin@phocon2025.com (PHOCON Official)")
#     
#     confirm = input("\n🚀 Start FAST email campaign? (y/n): ").strip().lower()
#     if confirm != 'y':
#         print("❌ Email campaign cancelled.")
#         return
#     
#     print("\n" + "="*70)
#     print("🚀 STARTING FAST PHOCON 2025 EMAIL CAMPAIGN...")
#     print(f"📧 Template: {selected_template['name']}")
#     print(f"⚡ Mode: {email_sender.max_workers} concurrent threads")
#     print("="*70)
#     
#     start_time = datetime.now()
#     print(f"⏰ Started: {start_time.strftime('%H:%M:%S')}")
#     
#     success = email_sender.process_excel_and_send_emails_fast()
#     
#     end_time = datetime.now()
#     duration = end_time - start_time
#     
#     email_sender.generate_report()
#     
#     print(f"\n⏰ Completed: {end_time.strftime('%H:%M:%S')}")
#     print(f"⏱  Total Duration: {duration}")
#     
#     if success:
#         print("\n🎉 FAST PHOCON 2025 EMAIL CAMPAIGN COMPLETED!")
#         print("⚡ Multi-threaded bulk email sending successful!")
#         print("📧 All emails sent from official PHOCON account!")
#     else:
#         print("\n⚠  Campaign completed with some issues.")
#         print("📋 Check the generated reports for details.")

# if __name__ == "__main__":
#     main()