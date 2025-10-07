-- PHOCON 2025 Complete Database Schema
-- Run this in Neon SQL Editor: https://console.neon.tech

-- Drop existing tables if you want to start fresh (CAREFUL!)
-- DROP TABLE IF EXISTS email_logs CASCADE;
-- DROP TABLE IF EXISTS campaigns CASCADE;
-- DROP TABLE IF EXISTS uploaded_files CASCADE;

-- 1. Campaigns Table
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

-- 2. Email Logs Table
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
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    cc_recipients TEXT,
    bcc_recipients TEXT
);

-- 3. Uploaded Files Table
CREATE TABLE IF NOT EXISTS uploaded_files (
    id SERIAL PRIMARY KEY,
    filename VARCHAR(255) NOT NULL,
    file_type VARCHAR(50) NOT NULL,
    file_path TEXT NOT NULL,
    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    session_id VARCHAR(255)
);

-- Indexes for better performance
CREATE INDEX IF NOT EXISTS idx_campaigns_status ON campaigns(status);
CREATE INDEX IF NOT EXISTS idx_campaigns_created ON campaigns(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_email_logs_campaign ON email_logs(campaign_id);
CREATE INDEX IF NOT EXISTS idx_email_logs_status ON email_logs(status);
CREATE INDEX IF NOT EXISTS idx_email_logs_email ON email_logs(recipient_email);
CREATE INDEX IF NOT EXISTS idx_uploaded_files_session ON uploaded_files(session_id);

-- View for campaign summary
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

-- Verify tables created
SELECT 'Database setup complete! ✅' as message;
SELECT table_name FROM information_schema.tables WHERE table_schema = 'public';

-- PHOCON 2025 Database Schema Update for CC/BCC Support
-- Run this in Neon SQL Editor: https://console.neon.tech

-- Add CC/BCC columns to email_logs table
ALTER TABLE email_logs 
ADD COLUMN IF NOT EXISTS cc_recipients TEXT,
ADD COLUMN IF NOT EXISTS bcc_recipients TEXT;

-- Create index for CC/BCC searches
CREATE INDEX IF NOT EXISTS idx_email_logs_cc ON email_logs(cc_recipients);
CREATE INDEX IF NOT EXISTS idx_email_logs_bcc ON email_logs(bcc_recipients);

-- Update the campaign_summary view to include CC/BCC stats
DROP VIEW IF EXISTS campaign_summary;

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
    COUNT(el.id) as total_logs,
    COUNT(CASE WHEN el.cc_recipients IS NOT NULL AND el.cc_recipients != '' THEN 1 END) as emails_with_cc,
    COUNT(CASE WHEN el.bcc_recipients IS NOT NULL AND el.bcc_recipients != '' THEN 1 END) as emails_with_bcc
FROM campaigns c
LEFT JOIN email_logs el ON c.id = el.campaign_id
GROUP BY c.id;

-- Verify updates
SELECT 'Database schema updated with CC/BCC support! ✅' as message;

-- Check new columns
SELECT column_name, data_type 
FROM information_schema.columns 
WHERE table_name = 'email_logs' 
AND column_name IN ('cc_recipients', 'bcc_recipients');


--repeated code



-- PHOCON 2025 - Custom Email Composer Database Schema
-- Run this in Neon SQL Editor: https://console.neon.tech

-- 1. Add custom email fields to campaigns table
ALTER TABLE campaigns 
ADD COLUMN IF NOT EXISTS custom_subject TEXT,
ADD COLUMN IF NOT EXISTS custom_body TEXT,
ADD COLUMN IF NOT EXISTS is_custom_template BOOLEAN DEFAULT FALSE;

-- 2. Create email templates table for saving custom templates
CREATE TABLE IF NOT EXISTS email_templates (
    id SERIAL PRIMARY KEY,
    template_name VARCHAR(255) NOT NULL,
    subject TEXT NOT NULL,
    body_html TEXT NOT NULL,
    created_by VARCHAR(255),
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    usage_count INTEGER DEFAULT 0
);

-- 3. Create indexes
CREATE INDEX IF NOT EXISTS idx_campaigns_custom ON campaigns(is_custom_template);
CREATE INDEX IF NOT EXISTS idx_templates_active ON email_templates(is_active);
CREATE INDEX IF NOT EXISTS idx_templates_created ON email_templates(created_at DESC);

-- 4. Update campaign_summary view
DROP VIEW IF EXISTS campaign_summary;

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
    c.is_custom_template,
    COUNT(el.id) as total_logs,
    COUNT(CASE WHEN el.cc_recipients IS NOT NULL AND el.cc_recipients != '' THEN 1 END) as emails_with_cc,
    COUNT(CASE WHEN el.bcc_recipients IS NOT NULL AND el.bcc_recipients != '' THEN 1 END) as emails_with_bcc
FROM campaigns c
LEFT JOIN email_logs el ON c.id = el.campaign_id
GROUP BY c.id;

-- Verify updates
SELECT 'Database schema updated with Custom Email Composer support! ✅' as message;

-- Check new columns
SELECT column_name, data_type 
FROM information_schema.columns 
WHERE table_name = 'campaigns' 
AND column_name IN ('custom_subject', 'custom_body', 'is_custom_template');

SELECT table_name 
FROM information_schema.tables 
WHERE table_schema = 'public' 
AND table_name = 'email_templates';