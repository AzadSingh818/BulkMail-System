-- PHOCON 2025 Database Schema
-- Run this in Neon SQL Editor

-- 1. Campaigns Table (Track har campaign)
CREATE TABLE IF NOT EXISTS campaigns (
    id SERIAL PRIMARY KEY,
    campaign_name VARCHAR(255) NOT NULL,
    template_id VARCHAR(10) NOT NULL,
    performance_mode VARCHAR(10) NOT NULL,
    status VARCHAR(50) DEFAULT 'pending',
    total_recipients INTEGER DEFAULT 0,
    emails_sent INTEGER DEFAULT 0,
    emails_failed INTEGER DEFAULT 0,
    success_rate DECIMAL(5,2) DEFAULT 0.00,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    completed_at TIMESTAMP,
    excel_filename VARCHAR(255)
);

-- 2. Email Logs Table (Har email ki details)
CREATE TABLE IF NOT EXISTS email_logs (
    id SERIAL PRIMARY KEY,
    campaign_id INTEGER REFERENCES campaigns(id) ON DELETE CASCADE,
    recipient_name VARCHAR(255),
    recipient_email VARCHAR(255) NOT NULL,
    template_used VARCHAR(10),
    status VARCHAR(50) DEFAULT 'pending',
    error_message TEXT,
    sent_at TIMESTAMP,
    thread_id INTEGER,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 3. Uploaded Files Table (Track files)
CREATE TABLE IF NOT EXISTS uploaded_files (
    id SERIAL PRIMARY KEY,
    filename VARCHAR(255) NOT NULL,
    file_type VARCHAR(50) NOT NULL,
    file_path TEXT NOT NULL,
    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    session_id VARCHAR(255)
);

-- Indexes (Fast queries ke liye)
CREATE INDEX IF NOT EXISTS idx_campaigns_status ON campaigns(status);
CREATE INDEX IF NOT EXISTS idx_campaigns_created ON campaigns(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_email_logs_campaign ON email_logs(campaign_id);
CREATE INDEX IF NOT EXISTS idx_email_logs_status ON email_logs(status);
CREATE INDEX IF NOT EXISTS idx_email_logs_email ON email_logs(recipient_email);

-- View (Summary ke liye)
CREATE OR REPLACE VIEW campaign_summary AS
SELECT 
    c.id,
    c.campaign_name,
    c.template_id,
    c.performance_mode,
    c.status,
    c.total_recipients,
    c.emails_sent,
    c.emails_failed,
    c.success_rate,
    c.created_at,
    c.completed_at,
    COUNT(el.id) as total_logs
FROM campaigns c
LEFT JOIN email_logs el ON c.id = el.campaign_id
GROUP BY c.id;

-- Success message
SELECT 'Database setup complete! ✅' as message;