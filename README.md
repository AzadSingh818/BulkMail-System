# PHOCON 2025 Bulk Email System

## Overview
This project is a high-performance bulk email sender and campaign control center for the PHOCON 2025 conference. It enables organizers to send personalized, multi-threaded email campaigns to large lists of recipients, with support for custom templates, image attachments, and detailed reporting. The system is built with Flask and provides both a web interface and a command-line utility.

## Features
- **Multi-threaded Email Sending:** Fast, concurrent email delivery with Gmail-safe limits.
- **Customizable Email Templates:** Three professional templates for invitations and reminders.
- **Image Attachments:** Easily include conference creatives and abstracts in emails.
- **Excel Integration:** Upload recipient lists in Excel format (.xlsx, .xls).
- **Performance Modes:** Choose between Safe, Fast, Turbo, and Beast modes for optimal speed.
- **Detailed Reporting:** Downloadable Excel reports for successful and failed email deliveries.
- **Modern Web UI:** Neon-themed dashboard for campaign management and file uploads.
- **Secure File Handling:** All uploads and downloads are sandboxed to the `uploads/` directory.

## Folder Structure
```
app.py                  # Flask web server and API endpoints
phocon_email_sender.py  # Bulk email sender logic and templates
requirements.txt        # Python dependencies
.env                    # SMTP credentials and configuration
render.yaml             # (Optional) Deployment configuration
uploads/                # Uploaded files and generated reports
templates/index.html    # Web dashboard UI
Creative.jpeg           # Sample creative image
PHOCON  Workshop Creative.jpeg  # Sample workshop image
PHOCON Abstract Submission.jpeg # Sample abstract image
```

## Setup Instructions
1. **Clone the repository** and navigate to the project folder.
2. **Install dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```
3. **Configure SMTP credentials:**
   - Edit the `.env` file with your Gmail SMTP details:
     ```env
     SMTP_SERVER=smtp.gmail.com
     SMTP_PORT=587
     SENDER_EMAIL=your_email@gmail.com
     SENDER_NAME=PHOCON 2025 Team
     SMTP_USERNAME=your_email@gmail.com
     SMTP_PASSWORD=your_app_password
     ```
   - Use an [App Password](https://support.google.com/accounts/answer/185833) for Gmail accounts with 2FA.
4. **Run the server:**
   ```powershell
   python app.py
   ```
   - The dashboard will be available at [http://localhost:5000](http://localhost:5000).

## Usage
### Web Dashboard
1. **Upload Files:**
   - Excel file with recipient emails and optional images (JPEG/PNG).
2. **Select Email Template:**
   - Choose from invitation or reminder templates.
3. **Choose Performance Mode:**
   - Select speed based on your Gmail rate limits.
4. **Start Campaign:**
   - Monitor progress and download reports for sent/failed emails.

### Command-Line
- Run `phocon_email_sender.py` directly for CLI-based campaigns.

## Security Notes
- All uploads and downloads are restricted to the `uploads/` folder.
- Sensitive SMTP credentials are stored in `.env` (never commit this file).
- The app uses Flask sessions for file path management.

## Troubleshooting
- **Gmail Rate Limits:** Use Safe/Fast mode to avoid account lockouts.
- **App Password Required:** Regular Gmail passwords may not work; use an App Password.
- **Excel Format:** Ensure your Excel file contains valid email addresses.
- **File Size Limit:** Maximum upload size is 16MB per file.

## License
This project is for educational and conference use. Please contact the PHOCON 2025 team for commercial licensing.

## Contact
- Email: admin@phocon2025.com
- Conference Website: [https://followmyevent.com/phocon-2025/](https://followmyevent.com/phocon-2025/)

---
*Developed for PHOCON 2025 by Kasturba Medical College, Manipal.*
