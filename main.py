import os
import tempfile
from flask import Flask, request, render_template_string, redirect, url_for, flash
import win32com.client
from werkzeug.utils import secure_filename
import pythoncom

app = Flask(__name__)
app.secret_key = "replace-with-a-secure-key"  # לשימוש ב־flash (במקום זה בחרו מפתח תקין)

# כל סוג קובץ שמותר - אפשר לצמצם לפי הצורך
ALLOWED_EXTENSIONS = {"pdf", "doc", "docx", "rtf", "txt"}

HTML_FORM = """
<!doctype html>
<html dir="rtl" lang="he">
<head>
    <meta charset="UTF-8">
    <title>פתח טיוטות ב-Outlook</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 50px auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        h2 {
            color: #0078d4;
        }
        form {
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        label {
            font-weight: bold;
            color: #333;
        }
        input[type="text"], textarea {
            width: 100%;
            padding: 8px;
            margin: 5px 0 15px 0;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-sizing: border-box;
        }
        input[type="file"] {
            margin: 5px 0 15px 0;
        }
        button {
            background-color: #0078d4;
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }
        button:hover {
            background-color: #005a9e;
        }
        .messages {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
            padding: 12px;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .error {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }
    </style>
</head>
<body>
    <h2>טופס פתיחת טיוטות ב-Outlook</h2>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        <div class="messages">
        {% for msg in messages %}
          <div>{{ msg }}</div>
        {% endfor %}
        </div>
      {% endif %}
    {% endwith %}

    <form method="post" enctype="multipart/form-data" action="{{ url_for('send') }}">
      <label>נושא המייל:</label><br>
      <input type="text" name="subject" required><br>

      <label>נמענים (מופרדים בפסיק):</label><br>
      <input type="text" name="recipients" placeholder="user1@example.com, user2@example.com" required><br>

      <label>גוף ההודעה:</label><br>
      <textarea name="body" rows="10"></textarea><br>

      <label>קובץ קורות חיים (מצורף):</label><br>
      <input type="file" name="cv" required><br>

      <button type="submit">פתח טיוטות ב-Outlook</button>
    </form>
</body>
</html>
"""


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def create_outlook_draft(to_email: str, subject: str, body: str, attachment_path: str):
    """
    יוצר טיוטת מייל ב-Outlook ופותח אותה לעריכה
    """
    print("Trying to connect to Outlook...")

    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("Connected to Outlook successfully.")
    except Exception as e:
        print("Failed to connect to Outlook:", e)
        pythoncom.CoUninitialize()
        raise

    try:
        # יצירת מייל חדש
        mail = outlook.CreateItem(0)
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body or ""

        # הוספת קובץ מצורף
        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)

        # פתיחת חלון המייל לעריכה
        mail.Display(False)
        print(f"Draft opened for {to_email}")

    except Exception as e:
        print(f"Error creating draft for {to_email}:", e)
        pythoncom.CoUninitialize()
        raise

    pythoncom.CoUninitialize()
    return True


@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML_FORM)


@app.route("/send", methods=["POST"])
def send():
    subject = request.form.get("subject", "").strip()
    recipients_raw = request.form.get("recipients", "").strip()
    body = request.form.get("body", "").strip()
    file = request.files.get("cv")

    # בדיקת שדות חובה
    if not subject or not recipients_raw or not file:
        flash("חסרים שדות חובה.")
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    if not filename or not allowed_file(filename):
        flash("סוג הקובץ לא נתמך. מותר: " + ", ".join(sorted(ALLOWED_EXTENSIONS)))
        return redirect(url_for("index"))

    # פירוק רשימת נמענים - מפרידים על ידי פסיק או ; או שורה חדשה
    raw_list = [r.strip() for r in recipients_raw.replace(";", ",").replace("\n", ",").split(",")]
    recipients = [r for r in raw_list if r]

    if not recipients:
        flash("לא נמצאו כתובות דוא\"ל תקינות בשדה הנמענים.")
        return redirect(url_for("index"))

    # שמירת הקובץ זמנית בתיקייה בטוחה
    tmp_dir = tempfile.mkdtemp(prefix="outlook_drafts_")
    saved_path = os.path.join(tmp_dir, filename)
    file.save(saved_path)

    success = []
    failed = []

    # יצירת טיוטה לכל נמען
    for r in recipients:
        try:
            create_outlook_draft(r, subject, body, saved_path)
            success.append(r)
        except Exception as e:
            failed.append((r, str(e)))

    # הודעת סיכום
    if success:
        msg = f"✅ טיוטות נפתחו בהצלחה עבור: {', '.join(success)}"
        flash(msg)

    if failed:
        error_msg = "❌ נכשל ליצור טיוטה עבור: " + ", ".join([f"{email} ({error})" for email, error in failed])
        flash(error_msg)

    return redirect(url_for("index"))


if __name__ == "__main__":
    print("🚀 Starting Flask application...")
    print("📧 Open your browser and go to: http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)