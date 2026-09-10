
import streamlit as st
import pandas as pd
import sqlite3
import hashlib
import smtplib
import ssl
from email.message import EmailMessage
from datetime import datetime, date
from pathlib import Path
from io import BytesIO

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "company_dashboard.db"
SEED_FILE = APP_DIR / "data" / "company_contacts.xlsx"

st.set_page_config(
    page_title="Company HR Contact Dashboard",
    page_icon="📞",
    layout="wide"
)

# -----------------------------
# Helpers
# -----------------------------
def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()

def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def execute(sql, params=()):
    with get_conn() as conn:
        conn.execute(sql, params)
        conn.commit()

def query_df(sql, params=()):
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)

def init_db():
    with get_conn() as conn:
        cur = conn.cursor()

        cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            display_name TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('master','admin')),
            active INTEGER NOT NULL DEFAULT 1
        )
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS companies (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company TEXT NOT NULL,
            sector TEXT,
            hr_name TEXT,
            phone TEXT,
            email TEXT,
            assigned_to INTEGER,
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(assigned_to) REFERENCES users(id)
        )
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS hr_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            old_hr_name TEXT,
            old_phone TEXT,
            old_email TEXT,
            changed_by INTEGER,
            changed_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(changed_by) REFERENCES users(id)
        )
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER NOT NULL,
            user_id INTEGER NOT NULL,
            response_status TEXT NOT NULL,
            remarks TEXT,
            call_date TEXT NOT NULL,
            followup_date TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS email_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_id INTEGER,
            user_id INTEGER NOT NULL,
            recipient TEXT NOT NULL,
            subject TEXT NOT NULL,
            body TEXT,
            sent_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            success INTEGER NOT NULL DEFAULT 0,
            error_message TEXT,
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
        """)

        # Create initial 4 users only if database is empty.
        count = cur.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if count == 0:
            # CHANGE THESE PASSWORDS IN THE ADMIN PANEL AFTER FIRST LOGIN.
            users = [
                ("master", "Master Admin", hash_password("Master@123"), "master"),
                ("admin1", "Admin 1", hash_password("Admin1@123"), "admin"),
                ("admin2", "Admin 2", hash_password("Admin2@123"), "admin"),
                ("admin3", "Admin 3", hash_password("Admin3@123"), "admin"),
            ]
            cur.executemany(
                "INSERT INTO users(username, display_name, password_hash, role) VALUES(?,?,?,?)",
                users
            )

        # Import Excel only if no company rows exist.
        ccount = cur.execute("SELECT COUNT(*) FROM companies").fetchone()[0]
        if ccount == 0 and SEED_FILE.exists():
            df = pd.read_excel(SEED_FILE)
            df = df.rename(columns={
                "Company": "company",
                "Sector": "sector",
                "HR": "hr_name",
                "Phno": "phone",
                "Email": "email"
            })
            for col in ["company", "sector", "hr_name", "phone", "email"]:
                if col not in df.columns:
                    df[col] = None
            df = df[["company", "sector", "hr_name", "phone", "email"]].copy()
            df["company"] = df["company"].fillna("").astype(str).str.strip()
            df = df[df["company"] != ""]
            for _, row in df.iterrows():
                cur.execute(
                    """INSERT INTO companies(company, sector, hr_name, phone, email)
                       VALUES(?,?,?,?,?)""",
                    (
                        row["company"],
                        None if pd.isna(row["sector"]) else str(row["sector"]),
                        None if pd.isna(row["hr_name"]) else str(row["hr_name"]),
                        None if pd.isna(row["phone"]) else str(row["phone"]),
                        None if pd.isna(row["email"]) else str(row["email"]),
                    )
                )

        conn.commit()

def authenticate(username, password):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE username=? AND active=1",
            (username,)
        ).fetchone()
    if row and row["password_hash"] == hash_password(password):
        return dict(row)
    return None

def get_users(active_only=True):
    sql = "SELECT id, username, display_name, role, active FROM users"
    if active_only:
        sql += " WHERE active=1"
    sql += " ORDER BY role DESC, display_name"
    return query_df(sql)

def get_company(company_id):
    df = query_df("SELECT * FROM companies WHERE id=?", (company_id,))
    return None if df.empty else df.iloc[0].to_dict()

def export_excel():
    companies = query_df("""
        SELECT c.id, c.company AS Company, c.sector AS Sector,
               c.hr_name AS HR, c.phone AS Phone, c.email AS Email,
               u.display_name AS Assigned_To, c.active AS Active
        FROM companies c
        LEFT JOIN users u ON c.assigned_to=u.id
        ORDER BY c.company
    """)

    responses = query_df("""
        SELECT r.id AS Response_ID, c.company AS Company,
               c.hr_name AS Current_HR, u.display_name AS Called_By,
               r.response_status AS Response, r.remarks AS Remarks,
               r.call_date AS Call_Date, r.followup_date AS Follow_Up_Date,
               r.created_at AS Created_At
        FROM responses r
        JOIN companies c ON r.company_id=c.id
        JOIN users u ON r.user_id=u.id
        ORDER BY r.created_at DESC
    """)

    hr_history = query_df("""
        SELECT h.id, c.company AS Company, h.old_hr_name AS Previous_HR,
               h.old_phone AS Previous_Phone, h.old_email AS Previous_Email,
               u.display_name AS Changed_By, h.changed_at AS Changed_At
        FROM hr_history h
        JOIN companies c ON h.company_id=c.id
        LEFT JOIN users u ON h.changed_by=u.id
        ORDER BY h.changed_at DESC
    """)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        companies.to_excel(writer, sheet_name="Companies", index=False)
        responses.to_excel(writer, sheet_name="Responses", index=False)
        hr_history.to_excel(writer, sheet_name="HR History", index=False)
    out.seek(0)
    return out.getvalue()

def send_email(recipient, subject, body):
    smtp_host = st.secrets.get("SMTP_HOST", "")
    smtp_port = int(st.secrets.get("SMTP_PORT", 587))
    smtp_user = st.secrets.get("SMTP_USER", "")
    smtp_password = st.secrets.get("SMTP_PASSWORD", "")
    sender_name = st.secrets.get("SMTP_SENDER_NAME", "Placement Cell")

    if not all([smtp_host, smtp_user, smtp_password]):
        raise RuntimeError("SMTP settings are not configured in .streamlit/secrets.toml")

    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{smtp_user}>"
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    context = ssl.create_default_context()
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls(context=context)
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

# -----------------------------
# App setup
# -----------------------------
init_db()

if "user" not in st.session_state:
    st.session_state.user = None

# -----------------------------
# Login
# -----------------------------
if st.session_state.user is None:
    st.title("📞 Company HR Contact Dashboard")
    st.caption("Authorized users only")

    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login", use_container_width=True)

    if submitted:
        user = authenticate(username.strip(), password)
        if user:
            st.session_state.user = user
            st.rerun()
        else:
            st.error("Invalid username or password.")

    st.info(
        "First-run demo accounts are listed in README.md. "
        "Change all passwords immediately after your first login."
    )
    st.stop()

user = st.session_state.user
is_master = user["role"] == "master"

# -----------------------------
# Sidebar
# -----------------------------
st.sidebar.title("📞 HR Calling Portal")
st.sidebar.write(f"Logged in as **{user['display_name']}**")
st.sidebar.caption("Master Admin" if is_master else "Calling Admin")

pages = ["Dashboard", "My Companies", "Follow-ups", "Call History", "Send Email", "Change Password"]
if is_master:
    pages = ["Dashboard", "Company Database", "Assignments", "Follow-ups", "Call History", "Reports", "User Management", "Send Email", "Change Password"]

page = st.sidebar.radio("Navigate", pages)

if st.sidebar.button("Logout", use_container_width=True):
    st.session_state.user = None
    st.rerun()

# -----------------------------
# Dashboard
# -----------------------------
if page == "Dashboard":
    st.title("Dashboard")

    if is_master:
        total = query_df("SELECT COUNT(*) n FROM companies WHERE active=1").iloc[0]["n"]
        assigned = query_df("SELECT COUNT(*) n FROM companies WHERE active=1 AND assigned_to IS NOT NULL").iloc[0]["n"]
        calls = query_df("SELECT COUNT(*) n FROM responses").iloc[0]["n"]
        due = query_df(
            "SELECT COUNT(*) n FROM responses WHERE followup_date=?",
            (date.today().isoformat(),)
        ).iloc[0]["n"]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Active Companies", int(total))
        c2.metric("Assigned", int(assigned))
        c3.metric("Total Responses", int(calls))
        c4.metric("Follow-ups Today", int(due))

        st.subheader("Admin-wise Workload")
        workload = query_df("""
            SELECT u.display_name AS Admin,
                   COUNT(c.id) AS Assigned,
                   SUM(CASE WHEN EXISTS(
                       SELECT 1 FROM responses r WHERE r.company_id=c.id
                   ) THEN 1 ELSE 0 END) AS Contacted
            FROM users u
            LEFT JOIN companies c ON c.assigned_to=u.id AND c.active=1
            WHERE u.active=1 AND u.role='admin'
            GROUP BY u.id, u.display_name
            ORDER BY u.display_name
        """)
        if not workload.empty:
            workload["Pending"] = workload["Assigned"] - workload["Contacted"]
            st.dataframe(workload, use_container_width=True, hide_index=True)
            st.bar_chart(workload.set_index("Admin")[["Assigned", "Contacted", "Pending"]])
    else:
        uid = user["id"]
        assigned = query_df(
            "SELECT COUNT(*) n FROM companies WHERE assigned_to=? AND active=1", (uid,)
        ).iloc[0]["n"]
        called = query_df(
            """SELECT COUNT(DISTINCT company_id) n FROM responses WHERE user_id=?""", (uid,)
        ).iloc[0]["n"]
        due = query_df(
            "SELECT COUNT(*) n FROM responses WHERE user_id=? AND followup_date=?",
            (uid, date.today().isoformat())
        ).iloc[0]["n"]
        overdue = query_df(
            "SELECT COUNT(*) n FROM responses WHERE user_id=? AND followup_date < ?",
            (uid, date.today().isoformat())
        ).iloc[0]["n"]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("My Assigned Companies", int(assigned))
        c2.metric("Companies Contacted", int(called))
        c3.metric("Follow-ups Today", int(due))
        c4.metric("Overdue Follow-ups", int(overdue))

        st.subheader("Today's Follow-ups")
        today_df = query_df("""
            SELECT c.company AS Company, c.hr_name AS HR, c.phone AS Phone,
                   r.response_status AS Last_Response, r.remarks AS Remarks
            FROM responses r
            JOIN companies c ON c.id=r.company_id
            WHERE r.user_id=? AND r.followup_date=?
            ORDER BY c.company
        """, (uid, date.today().isoformat()))
        st.dataframe(today_df, use_container_width=True, hide_index=True)

# -----------------------------
# Company Database
# -----------------------------
elif page == "Company Database" and is_master:
    st.title("🏢 Company Database")

    search = st.text_input("Search company / sector / HR / email")
    params = []
    where = "WHERE c.active=1"
    if search.strip():
        q = f"%{search.strip()}%"
        where += " AND (c.company LIKE ? OR c.sector LIKE ? OR c.hr_name LIKE ? OR c.email LIKE ? OR c.phone LIKE ?)"
        params.extend([q, q, q, q, q])

    companies = query_df(f"""
        SELECT c.id, c.company AS Company, c.sector AS Sector,
               c.hr_name AS HR, c.phone AS Phone, c.email AS Email,
               u.display_name AS Assigned_To
        FROM companies c
        LEFT JOIN users u ON c.assigned_to=u.id
        {where}
        ORDER BY c.company
    """, tuple(params))
    st.dataframe(companies, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Edit HR / Contact Details")
    if not companies.empty:
        options = {
            f"{row.Company} — {row.HR if pd.notna(row.HR) else 'No HR'}": int(row.id)
            for _, row in companies.iterrows()
        }
        selected_label = st.selectbox("Select company", list(options.keys()))
        cid = options[selected_label]
        current = get_company(cid)

        with st.form("edit_company"):
            company = st.text_input("Company", value=current.get("company") or "")
            sector = st.text_input("Sector", value=current.get("sector") or "")
            hr = st.text_input("Current HR", value=current.get("hr_name") or "")
            phone = st.text_input("Phone", value=current.get("phone") or "")
            email = st.text_input("Email", value=current.get("email") or "")
            save = st.form_submit_button("Save Changes")

        if save:
            changed_contact = (
                (current.get("hr_name") or "") != hr or
                (current.get("phone") or "") != phone or
                (current.get("email") or "") != email
            )
            with get_conn() as conn:
                if changed_contact:
                    conn.execute("""
                        INSERT INTO hr_history(company_id, old_hr_name, old_phone, old_email, changed_by)
                        VALUES(?,?,?,?,?)
                    """, (cid, current.get("hr_name"), current.get("phone"), current.get("email"), user["id"]))
                conn.execute("""
                    UPDATE companies
                    SET company=?, sector=?, hr_name=?, phone=?, email=?, updated_at=CURRENT_TIMESTAMP
                    WHERE id=?
                """, (company, sector, hr, phone, email, cid))
                conn.commit()
            st.success("Company/contact details updated.")
            st.rerun()

    st.divider()
    st.subheader("Add New Company")
    with st.form("add_company"):
        company = st.text_input("Company name")
        sector = st.text_input("Sector")
        hr = st.text_input("HR name")
        phone = st.text_input("Phone")
        email = st.text_input("Email")
        add = st.form_submit_button("Add Company")
    if add:
        if not company.strip():
            st.error("Company name is required.")
        else:
            execute("""
                INSERT INTO companies(company, sector, hr_name, phone, email)
                VALUES(?,?,?,?,?)
            """, (company.strip(), sector.strip(), hr.strip(), phone.strip(), email.strip()))
            st.success("Company added.")
            st.rerun()

# -----------------------------
# Assignments
# -----------------------------
elif page == "Assignments" and is_master:
    st.title("👑 Assign Companies")

    admins = get_users()
    admins = admins[admins["role"] == "admin"]
    admin_map = {row.display_name: int(row.id) for _, row in admins.iterrows()}

    filter_mode = st.radio("Show", ["Unassigned", "All"], horizontal=True)
    if filter_mode == "Unassigned":
        companies = query_df("""
            SELECT id, company, sector, hr_name
            FROM companies
            WHERE active=1 AND assigned_to IS NULL
            ORDER BY company
        """)
    else:
        companies = query_df("""
            SELECT id, company, sector, hr_name
            FROM companies
            WHERE active=1
            ORDER BY company
        """)

    if companies.empty:
        st.info("No companies in this view.")
    else:
        labels = {
            f"{row.company} | {row.sector if pd.notna(row.sector) else ''} | HR: {row.hr_name if pd.notna(row.hr_name) else ''}": int(row.id)
            for _, row in companies.iterrows()
        }

        selected = st.multiselect("Select one or more companies", list(labels.keys()))
        admin_name = st.selectbox("Assign to", list(admin_map.keys()))

        if st.button("Assign Selected", type="primary"):
            if not selected:
                st.warning("Select at least one company.")
            else:
                ids = [labels[x] for x in selected]
                with get_conn() as conn:
                    conn.executemany(
                        "UPDATE companies SET assigned_to=?, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                        [(admin_map[admin_name], cid) for cid in ids]
                    )
                    conn.commit()
                st.success(f"Assigned {len(ids)} company/companies to {admin_name}.")
                st.rerun()

    st.subheader("Current Assignments")
    current = query_df("""
        SELECT c.company AS Company, c.sector AS Sector,
               c.hr_name AS HR, u.display_name AS Assigned_To
        FROM companies c
        LEFT JOIN users u ON c.assigned_to=u.id
        WHERE c.active=1
        ORDER BY u.display_name, c.company
    """)
    st.dataframe(current, use_container_width=True, hide_index=True)

# -----------------------------
# My Companies
# -----------------------------
elif page == "My Companies" and not is_master:
    st.title("📞 My Assigned Companies")
    uid = user["id"]

    search = st.text_input("Search my companies")
    q = f"%{search.strip()}%"
    companies = query_df("""
        SELECT c.id, c.company AS Company, c.sector AS Sector, c.hr_name AS HR,
               c.phone AS Phone, c.email AS Email
        FROM companies c
        WHERE c.active=1 AND c.assigned_to=?
          AND (?='' OR c.company LIKE ? OR c.hr_name LIKE ? OR c.sector LIKE ?)
        ORDER BY c.company
    """, (uid, search.strip(), q, q, q))

    st.dataframe(companies, use_container_width=True, hide_index=True)

    if not companies.empty:
        options = {f"{row.Company} — {row.HR if pd.notna(row.HR) else 'No HR'}": int(row.id)
                   for _, row in companies.iterrows()}
        label = st.selectbox("Select company to record call", list(options.keys()))
        cid = options[label]
        company = get_company(cid)

        st.write(f"**Phone:** {company.get('phone') or '-'}")
        st.write(f"**Email:** {company.get('email') or '-'}")

        statuses = [
            "Interested",
            "Not Interested",
            "Call Later",
            "No Response",
            "Wrong Number",
            "HR Changed",
            "Email Sent",
            "Meeting Scheduled",
            "Follow-up Required",
            "Other"
        ]

        with st.form("response_form"):
            status = st.selectbox("Response", statuses)
            remarks = st.text_area("Remarks")
            call_date = st.date_input("Call date", value=date.today())
            needs_followup = st.checkbox("Set follow-up reminder")
            followup_date = st.date_input("Next follow-up date", value=date.today()) if needs_followup else None
            save = st.form_submit_button("Save Response", type="primary")

        if save:
            execute("""
                INSERT INTO responses(company_id, user_id, response_status, remarks, call_date, followup_date)
                VALUES(?,?,?,?,?,?)
            """, (
                cid, uid, status, remarks.strip(),
                call_date.isoformat(),
                followup_date.isoformat() if followup_date else None
            ))
            st.success("Response saved to Call History.")
            st.rerun()

        st.subheader("Update HR / phone if contact has changed")
        with st.form("admin_hr_update"):
            new_hr = st.text_input("HR name", value=company.get("hr_name") or "")
            new_phone = st.text_input("Phone", value=company.get("phone") or "")
            new_email = st.text_input("Email", value=company.get("email") or "")
            update = st.form_submit_button("Update Contact")

        if update:
            with get_conn() as conn:
                conn.execute("""
                    INSERT INTO hr_history(company_id, old_hr_name, old_phone, old_email, changed_by)
                    VALUES(?,?,?,?,?)
                """, (cid, company.get("hr_name"), company.get("phone"), company.get("email"), uid))
                conn.execute("""
                    UPDATE companies SET hr_name=?, phone=?, email=?, updated_at=CURRENT_TIMESTAMP WHERE id=?
                """, (new_hr, new_phone, new_email, cid))
                conn.commit()
            st.success("Contact updated; previous details were preserved in HR History.")
            st.rerun()

# -----------------------------
# Follow-ups
# -----------------------------
elif page == "Follow-ups":
    st.title("⏰ Follow-ups")

    if is_master:
        data = query_df("""
            SELECT r.id, c.company AS Company, c.hr_name AS HR,
                   c.phone AS Phone, u.display_name AS Admin,
                   r.response_status AS Last_Response,
                   r.remarks AS Remarks, r.followup_date AS Follow_Up
            FROM responses r
            JOIN companies c ON r.company_id=c.id
            JOIN users u ON r.user_id=u.id
            WHERE r.followup_date IS NOT NULL
            ORDER BY r.followup_date
        """)
    else:
        data = query_df("""
            SELECT r.id, c.company AS Company, c.hr_name AS HR,
                   c.phone AS Phone, r.response_status AS Last_Response,
                   r.remarks AS Remarks, r.followup_date AS Follow_Up
            FROM responses r
            JOIN companies c ON r.company_id=c.id
            WHERE r.user_id=? AND r.followup_date IS NOT NULL
            ORDER BY r.followup_date
        """, (user["id"],))

    if data.empty:
        st.info("No follow-ups have been scheduled.")
    else:
        data["Follow_Up"] = pd.to_datetime(data["Follow_Up"], errors="coerce").dt.date
        today = date.today()
        data["Status"] = data["Follow_Up"].apply(
            lambda d: "🔴 Overdue" if d and d < today else ("🟠 Due Today" if d == today else "🟡 Upcoming")
        )
        st.dataframe(data, use_container_width=True, hide_index=True)

# -----------------------------
# Call History
# -----------------------------
elif page == "Call History":
    st.title("📝 Call Response History")

    if is_master:
        history = query_df("""
            SELECT r.id, c.company AS Company, c.hr_name AS Current_HR,
                   u.display_name AS Called_By, r.response_status AS Response,
                   r.remarks AS Remarks, r.call_date AS Call_Date,
                   r.followup_date AS Follow_Up, r.created_at AS Logged_At
            FROM responses r
            JOIN companies c ON r.company_id=c.id
            JOIN users u ON r.user_id=u.id
            ORDER BY r.created_at DESC
        """)
    else:
        history = query_df("""
            SELECT r.id, c.company AS Company, c.hr_name AS Current_HR,
                   r.response_status AS Response, r.remarks AS Remarks,
                   r.call_date AS Call_Date, r.followup_date AS Follow_Up,
                   r.created_at AS Logged_At
            FROM responses r
            JOIN companies c ON r.company_id=c.id
            WHERE r.user_id=?
            ORDER BY r.created_at DESC
        """, (user["id"],))

    st.dataframe(history, use_container_width=True, hide_index=True)

# -----------------------------
# Reports
# -----------------------------
elif page == "Reports" and is_master:
    st.title("📊 Reports")

    status_counts = query_df("""
        SELECT response_status AS Response, COUNT(*) AS Count
        FROM responses
        GROUP BY response_status
        ORDER BY Count DESC
    """)
    if not status_counts.empty:
        st.subheader("Response Distribution")
        st.bar_chart(status_counts.set_index("Response"))

    st.subheader("Admin Performance")
    admin_perf = query_df("""
        SELECT u.display_name AS Admin,
               COUNT(r.id) AS Calls_Logged,
               COUNT(DISTINCT r.company_id) AS Companies_Contacted,
               SUM(CASE WHEN r.response_status='Interested' THEN 1 ELSE 0 END) AS Interested
        FROM users u
        LEFT JOIN responses r ON r.user_id=u.id
        WHERE u.role='admin'
        GROUP BY u.id, u.display_name
        ORDER BY u.display_name
    """)
    st.dataframe(admin_perf, use_container_width=True, hide_index=True)

    st.subheader("Download Excel Report")
    st.download_button(
        "⬇️ Download Companies + Responses + HR History",
        data=export_excel(),
        file_name=f"company_calling_report_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------
# User Management
# -----------------------------
elif page == "User Management" and is_master:
    st.title("👥 User Management")
    users = get_users(active_only=False)
    st.dataframe(users, use_container_width=True, hide_index=True)

    st.info("This version is intended for 4 authorized accounts: 1 Master Admin + 3 calling admins.")

    with st.form("reset_password"):
        username = st.selectbox("User", users["username"].tolist())
        new_password = st.text_input("New password", type="password")
        reset = st.form_submit_button("Reset Password")
    if reset:
        if len(new_password) < 8:
            st.error("Use at least 8 characters.")
        else:
            execute(
                "UPDATE users SET password_hash=? WHERE username=?",
                (hash_password(new_password), username)
            )
            st.success("Password changed.")

# -----------------------------
# Send Email
# -----------------------------
elif page == "Send Email":
    st.title("📧 Send Email Through Portal")

    if is_master:
        companies = query_df("SELECT id, company, hr_name, email FROM companies WHERE active=1 ORDER BY company")
    else:
        companies = query_df(
            "SELECT id, company, hr_name, email FROM companies WHERE active=1 AND assigned_to=? ORDER BY company",
            (user["id"],)
        )

    if companies.empty:
        st.info("No companies available.")
    else:
        options = {f"{row.company} — {row.hr_name if pd.notna(row.hr_name) else 'No HR'}": int(row.id)
                   for _, row in companies.iterrows()}
        label = st.selectbox("Company", list(options.keys()))
        cid = options[label]
        company = get_company(cid)

        with st.form("email_form"):
            recipient = st.text_input("To", value=company.get("email") or "")
            subject = st.text_input("Subject")
            body = st.text_area("Message", height=250)
            send = st.form_submit_button("Send Email", type="primary")

        if send:
            success = 0
            err = None
            try:
                send_email(recipient.strip(), subject.strip(), body)
                success = 1
                st.success("Email sent.")
            except Exception as e:
                err = str(e)
                st.error(f"Email could not be sent: {err}")

            execute("""
                INSERT INTO email_log(company_id, user_id, recipient, subject, body, success, error_message)
                VALUES(?,?,?,?,?,?,?)
            """, (cid, user["id"], recipient.strip(), subject.strip(), body, success, err))

# -----------------------------
# Change Password
# -----------------------------
elif page == "Change Password":
    st.title("🔐 Change Password")
    with st.form("change_password"):
        old_pw = st.text_input("Current password", type="password")
        new_pw = st.text_input("New password", type="password")
        confirm = st.text_input("Confirm new password", type="password")
        change = st.form_submit_button("Change Password")

    if change:
        if authenticate(user["username"], old_pw) is None:
            st.error("Current password is incorrect.")
        elif len(new_pw) < 8:
            st.error("New password must be at least 8 characters.")
        elif new_pw != confirm:
            st.error("Passwords do not match.")
        else:
            execute(
                "UPDATE users SET password_hash=? WHERE id=?",
                (hash_password(new_pw), user["id"])
            )
            st.success("Password changed successfully.")
