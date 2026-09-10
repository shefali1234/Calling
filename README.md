
# Company HR Contact & Calling Dashboard

This is a working first version built from your uploaded Excel file.

## Included features

- 4 authorized users: 1 Master Admin + 3 calling admins
- Master Admin can assign companies to admins
- Each admin sees their assigned companies
- Call responses are saved separately in the database
- Follow-up dates and reminder views
- HR name, phone, and email can be updated
- Old HR/contact details are preserved in HR History
- Email can be sent through the portal using SMTP configuration
- Master Admin analytics and admin-wise work allocation
- Excel report download with three sheets:
  - Companies
  - Responses
  - HR History

## Default first-run logins

CHANGE THESE PASSWORDS IMMEDIATELY AFTER FIRST LOGIN.

- Master Admin
  - Username: `master`
  - Password: `Master@123`

- Admin 1
  - Username: `admin1`
  - Password: `Admin1@123`

- Admin 2
  - Username: `admin2`
  - Password: `Admin2@123`

- Admin 3
  - Username: `admin3`
  - Password: `Admin3@123`

## Run locally

1. Install Python 3.10 or newer.
2. Open a terminal in this project folder.
3. Install packages:

   `pip install -r requirements.txt`

4. Start the portal:

   `streamlit run app.py`

5. Your browser will open the application.

## Email setup

To enable email:

1. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`
2. Fill in your SMTP details.
3. For Gmail/Google Workspace, use an App Password rather than your normal password.

Example:

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = "placement@example.com"
SMTP_PASSWORD = "your-app-password"
SMTP_SENDER_NAME = "Placement Cell"

## Data storage

The uploaded Excel file is used only as initial seed data.

On first run, the application creates:

`company_dashboard.db`

All assignments, responses, HR updates, reminders and email logs are then stored in this SQLite database.

This avoids multi-user conflicts that would occur if four users edited the same Excel file directly.

## Recommended production deployment

For a small internal deployment, Streamlit + SQLite is sufficient to test the workflow.

For permanent multi-user institutional use, the next production upgrade should be:

- deploy on a secured server/cloud
- switch SQLite to PostgreSQL
- enable HTTPS
- use organization-managed credentials
- configure automated backups
- optionally add actual email/calendar notifications
