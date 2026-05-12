AWS Transaction Reporting Automation

Automated serverless solution for transactional report generation using AWS Lambda, integrated with MongoDB and Amazon S3.

Features
- Hourly transaction extraction
- Automated XLSX report generation
- Email delivery automation
- Scheduled execution with EventBridge
- MongoDB integration
- Amazon S3 storage
- Secure AWS resource access using IAM roles
- VPC-integrated Lambda execution

Technologies
- Python
- AWS Lambda
- Amazon EventBridge
- Amazon S3
- MongoDB
- Amazon VPC
- IAM Roles
- Security Groups
- OpenPyXL
- SMTP

Arquitectura:
EventBridge → AWS Lambda → MongoDB → OpenPyXL Report Generation → Amazon S3 → Email Delivery