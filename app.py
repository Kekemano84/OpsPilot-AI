import os
import sqlite3
from datetime import datetime, date
from io import BytesIO
from functools import wraps

from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import cm

load_dotenv()

APP_NAME = "OpsPilot AI"
DB_PATH = os.path.join(os.path.dirname(__file__), "opspilot.db")
FREE_DAILY_LIMIT = 5

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-secret-change-me")


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                company TEXT DEFAULT '',
                role TEXT DEFAULT '',
                plan TEXT DEFAULT 'free',
                is_admin INTEGER DEFAULT 0,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                report_type TEXT NOT NULL,
                input_text TEXT NOT NULL,
                output_text TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY(user_id) REFERENCES users(id)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS usage_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                usage_date TEXT NOT NULL,
                count INTEGER DEFAULT 0,
                UNIQUE(user_id, usage_date),
                FOREIGN KEY(user_id) REFERENCES users(id)
            )
        """)
        admin = conn.execute("SELECT id FROM users WHERE email=?", ("admin@opspilot.local",)).fetchone()
        if not admin:
            conn.execute(
                "INSERT INTO users (name,email,password_hash,company,role,plan,is_admin,created_at) VALUES (?,?,?,?,?,?,?,?)",
                (
                    "OpsPilot Admin",
                    "admin@opspilot.local",
                    generate_password_hash("admin123"),
                    "OpsPilot AI",
                    "Admin",
                    "business",
                    1,
                    datetime.utcnow().isoformat(),
                ),
            )


def current_user():
    uid = session.get("user_id")
    if not uid:
        return None
    with db() as conn:
        return conn.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()


@app.context_processor
def inject_globals():
    return {"app_name": APP_NAME, "user": current_user()}


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            flash("Please log in first.", "warning")
            return redirect(url_for("login"))
        return fn(*args, **kwargs)

    return wrapper


def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        user = current_user()
        if not user or not user["is_admin"]:
            flash("Admin access required.", "danger")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)

    return wrapper


def get_usage_today(user_id):
    today = date.today().isoformat()
    with db() as conn:
        row = conn.execute(
            "SELECT count FROM usage_logs WHERE user_id=? AND usage_date=?",
            (user_id, today),
        ).fetchone()
        return row["count"] if row else 0


def increment_usage(user_id):
    today = date.today().isoformat()
    with db() as conn:
        row = conn.execute(
            "SELECT count FROM usage_logs WHERE user_id=? AND usage_date=?",
            (user_id, today),
        ).fetchone()
        if row:
            conn.execute(
                "UPDATE usage_logs SET count=count+1 WHERE user_id=? AND usage_date=?",
                (user_id, today),
            )
        else:
            conn.execute(
                "INSERT INTO usage_logs (user_id, usage_date, count) VALUES (?,?,1)",
                (user_id, today),
            )


def can_generate(user):
    if user["plan"] in ("pro", "business"):
        return True
    return get_usage_today(user["id"]) < FREE_DAILY_LIMIT


TEMPLATES = {
    "daily_brief": {
        "title": "Daily Brief Generator",
        "label": "Quick notes for today's shift",
        "placeholder": (
            "Example: High volume, planned HC 24, actual HC 21, 3 agency, "
            "2 delayed trailers, focus on loading standards, safety and housekeeping."
        ),
        "system": (
            "You are OpsPilot AI, a professional assistant for UK warehouse leaders. "
            "Create a short, clear and practical start-of-shift daily brief for warehouse staff. "
            "Use simple English. Do not invent facts. Focus on safety, quality, service, people and communication. "
            "Make it suitable for a Team Leader, First Line Manager, Shift Manager or Operations Manager to read aloud."
        ),
    },
    "handover": {
        "title": "Shift Handover Generator",
        "label": "Quick notes for the next shift",
        "placeholder": (
            "Example: 1 delayed trailer, 2 damaged pallets isolated, missing safety pin reported, "
            "chamber 2 clean, lane HBS2 still needs checking."
        ),
        "system": (
            "You are OpsPilot AI, a professional assistant for UK warehouse leaders. "
            "Create a structured and professional shift handover. "
            "Include completed work, outstanding issues, safety/quality concerns, risks and next actions. "
            "Use simple English and do not invent facts."
        ),
    },
    "email": {
        "title": "Warehouse Email Writer",
        "label": "What should the email be about?",
        "placeholder": (
            "Example: Write an email to night shift about poor loading standards and damaged pallets found on handover."
        ),
        "system": (
            "You write professional, simple and polite UK workplace emails for warehouse and logistics managers. "
            "Use a clear subject, short paragraphs and a helpful tone."
        ),
    },
    "incident": {
        "title": "Incident Report Generator",
        "label": "Describe the incident or issue",
        "placeholder": (
            "Example: Damaged pallet found on lane HBS2, stock leaning, isolated and reported to FLM. No injury."
        ),
        "system": (
            "You create clear warehouse incident reports. Include what happened, immediate action, risk, corrective action "
            "and follow-up. Do not invent facts. Keep wording factual and professional."
        ),
    },
}


def normalise_notes(text):
    lines = []
    for raw in text.replace(";", "\n").split("\n"):
        item = raw.strip(" -•\t")
        if item:
            lines.append(item)
    if not lines and text.strip():
        lines.append(text.strip())
    return lines


def bullet_list(lines):
    if not lines:
        return "- No additional notes provided."
    return "\n".join(f"- {line}" for line in lines)


def template_ai(report_type, text, user):
    clean = text.strip()
    lines = normalise_notes(clean)
    today = date.today().strftime("%d/%m/%Y")
    name = user["name"]
    company = user["company"] or "N/A"
    role = user["role"] or "Warehouse Leader"

    if report_type == "daily_brief":
        return f"""GOOD MORNING TEAM

Date: {today}
Briefed by: {name}
Role: {role}
Company/Site: {company}

SHIFT PRIORITIES

{bullet_list(lines)}

SAFETY FOCUS

- Please keep safety as the first priority throughout the shift.
- Report any hazards, damages or near misses immediately.
- Keep walkways, fire exits and working areas clear.

OPERATIONAL FOCUS

- Follow the agreed loading and warehouse standards.
- Escalate delays, damages or stock issues early.
- Communicate clearly between team leaders, loaders, runners and managers.
- Keep housekeeping standards high throughout the shift.

TEAM MESSAGE

Let's support each other, stay focused, and aim to leave the operation in a strong position for the next shift.

Thank you for your cooperation.

{name}"""

    if report_type == "handover":
        return f"""SHIFT HANDOVER

Date: {today}
Completed by: {name}
Role: {role}
Company/Site: {company}

SUMMARY

{bullet_list(lines)}

COMPLETED / UPDATED

- Key operational points have been reviewed.
- Relevant issues have been communicated where required.
- Any immediate safety or quality concerns should be escalated.

OUTSTANDING ISSUES

- Review the summary points above at the start of the next shift.
- Confirm ownership for any open actions.
- Monitor for repeat issues or further operational impact.

SAFETY / QUALITY RISKS

- Check whether any issues affect safety, service, quality or customer experience.
- Ensure damaged stock, unsafe pallets or racking concerns are isolated and reported.

NEXT ACTIONS

- Prioritise outstanding trailers, safety concerns and customer-critical work.
- Communicate any delays or blockers early.
- Update the relevant manager once actions are complete.

Handover completed by: {name}"""

    if report_type == "email":
        return f"""Subject: Warehouse Operations Update

Hi team,

I wanted to make you aware of the following operational point:

{clean}

Please review this and follow up where required. The main focus is to maintain safe working standards, clear communication and consistent operational quality.

If there are any issues or blockers, please escalate them as early as possible.

Kind regards,
{name}"""

    return f"""INCIDENT / ISSUE REPORT

Reported by: {name}
Role: {role}
Company/Site: {company}
Date: {today}

INCIDENT / ISSUE DESCRIPTION

{clean}

IMMEDIATE ACTION TAKEN

- The issue was identified and made safe where required.
- Relevant colleagues or management should be informed.
- Any affected area, pallet, stock or equipment should be isolated if needed.

RISK / IMPACT

- Potential impact on safety, quality, service or customer experience if not addressed.

CORRECTIVE ACTION

- Review the root cause.
- Communicate the expected standard to the team.
- Confirm ownership for follow-up actions.
- Monitor for repeat issues.

FOLLOW-UP REQUIRED

- Confirm that corrective actions have been completed.
- Record any further findings or escalation required."""


def generate_ai(report_type, text, user):
    api_key = os.getenv("OPENAI_API_KEY")
    meta = TEMPLATES[report_type]

    if not api_key:
        return template_ai(report_type, text, user), "Demo AI Template"

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": meta["system"]},
                {
                    "role": "user",
                    "content": (
                        f"User name: {user['name']}\n"
                        f"User role: {user['role'] or 'Warehouse Leader'}\n"
                        f"Company/site: {user['company'] or 'Warehouse Operation'}\n"
                        f"Report type: {report_type}\n"
                        f"Input notes:\n{text}\n\n"
                        "Return only the finished document. Do not explain your process."
                    ),
                },
            ],
            temperature=0.35,
            max_tokens=900,
        )
        return resp.choices[0].message.content.strip(), "OpenAI"
    except Exception:
        return (
            template_ai(report_type, text, user)
            + "\n\n[Note: OpenAI was unavailable, so OpsPilot used the built-in demo template.]",
            "Demo fallback",
        )


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        company = request.form.get("company", "").strip()
        role = request.form.get("role", "").strip()

        if not name or not email or not password:
            flash("Name, email and password are required.", "danger")
            return redirect(url_for("register"))

        try:
            with db() as conn:
                conn.execute(
                    "INSERT INTO users (name,email,password_hash,company,role,plan,created_at) VALUES (?,?,?,?,?,?,?)",
                    (
                        name,
                        email,
                        generate_password_hash(password),
                        company,
                        role,
                        "free",
                        datetime.utcnow().isoformat(),
                    ),
                )
            flash("Account created. Please log in.", "success")
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            flash("An account with this email already exists.", "danger")

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        with db() as conn:
            user = conn.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()

        if user and check_password_hash(user["password_hash"], password):
            session["user_id"] = user["id"]
            flash("Welcome back.", "success")
            return redirect(url_for("dashboard"))

        flash("Invalid email or password.", "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("Logged out.", "info")
    return redirect(url_for("index"))


@app.route("/dashboard")
@login_required
def dashboard():
    user = current_user()
    with db() as conn:
        reports = conn.execute(
            "SELECT * FROM reports WHERE user_id=? ORDER BY id DESC LIMIT 5",
            (user["id"],),
        ).fetchall()
    return render_template(
        "dashboard.html",
        usage=get_usage_today(user["id"]),
        limit=FREE_DAILY_LIMIT,
        reports=reports,
    )


@app.route("/generate/<report_type>", methods=["GET", "POST"])
@login_required
def generator(report_type):
    if report_type not in TEMPLATES:
        flash("Unknown generator.", "danger")
        return redirect(url_for("dashboard"))

    user = current_user()
    result = None
    source = None

    if request.method == "POST":
        input_text = request.form.get("input_text", "").strip()

        if not input_text:
            flash("Please enter some notes first.", "warning")
        elif not can_generate(user):
            flash("Free daily limit reached. Upgrade to PRO for fair-use unlimited generations.", "warning")
            return redirect(url_for("pricing"))
        else:
            result, source = generate_ai(report_type, input_text, user)
            increment_usage(user["id"])

            with db() as conn:
                conn.execute(
                    "INSERT INTO reports (user_id, report_type, input_text, output_text, created_at) VALUES (?,?,?,?,?)",
                    (
                        user["id"],
                        report_type,
                        input_text,
                        result,
                        datetime.utcnow().isoformat(),
                    ),
                )

            flash(f"Generated using {source}.", "success")

    return render_template(
        "generator.html",
        report_type=report_type,
        meta=TEMPLATES[report_type],
        result=result,
        usage=get_usage_today(user["id"]),
        limit=FREE_DAILY_LIMIT,
    )


@app.route("/history")
@login_required
def history():
    user = current_user()
    with db() as conn:
        reports = conn.execute(
            "SELECT * FROM reports WHERE user_id=? ORDER BY id DESC",
            (user["id"],),
        ).fetchall()
    return render_template("history.html", reports=reports)


@app.route("/report/<int:report_id>/pdf")
@login_required
def report_pdf(report_id):
    user = current_user()
    with db() as conn:
        report = conn.execute(
            "SELECT * FROM reports WHERE id=? AND user_id=?",
            (report_id, user["id"]),
        ).fetchone()

    if not report:
        flash("Report not found.", "danger")
        return redirect(url_for("history"))

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )
    styles = getSampleStyleSheet()
    body = ParagraphStyle("Body", parent=styles["BodyText"], leading=16, fontSize=10)
    story = [Paragraph("OpsPilot AI Report", styles["Title"]), Spacer(1, 12)]
    story.append(Paragraph(f"Type: {report['report_type'].replace('_', ' ').title()}", styles["Heading2"]))
    story.append(Paragraph(f"Created: {report['created_at']}", body))
    story.append(Spacer(1, 12))

    for line in report["output_text"].split("\n"):
        safe_line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") or " "
        story.append(Paragraph(safe_line, body))

    doc.build(story)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"opspilot_report_{report_id}.pdf",
        mimetype="application/pdf",
    )


@app.route("/pricing")
def pricing():
    return render_template("pricing.html")


@app.route("/upgrade/<plan>")
@login_required
def upgrade(plan):
    if plan not in ("free", "pro", "business"):
        flash("Invalid plan.", "danger")
        return redirect(url_for("pricing"))

    user = current_user()
    with db() as conn:
        conn.execute("UPDATE users SET plan=? WHERE id=?", (plan, user["id"]))

    flash(f"Demo upgrade complete: {plan.upper()} plan activated. Stripe can be connected later.", "success")
    return redirect(url_for("dashboard"))


@app.route("/admin")
@login_required
@admin_required
def admin():
    with db() as conn:
        users = conn.execute("SELECT * FROM users ORDER BY id DESC").fetchall()
        reports_count = conn.execute("SELECT COUNT(*) AS c FROM reports").fetchone()["c"]

    return render_template("admin.html", users=users, reports_count=reports_count)


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    user = current_user()

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        company = request.form.get("company", "").strip()
        role = request.form.get("role", "").strip()

        with db() as conn:
            conn.execute(
                "UPDATE users SET name=?, company=?, role=? WHERE id=?",
                (name, company, role, user["id"]),
            )

        flash("Profile updated.", "success")
        return redirect(url_for("profile"))

    return render_template("profile.html")


if __name__ == "__main__":
    init_db()
    app.run(debug=True)
