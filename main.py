from flask import Flask, request, render_template, redirect, url_for, flash
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

app = Flask(__name__)
app.secret_key = 's\xf8\x10 ?\xc7\xa49$\xe7J\x9dj\xdfJ\x0bi\x0e\xe7\xb2F\xb9\xf8\n'  # Required for flash messages

SMTP_SERVER = 'vservit.icewarpcloud.in'
SMTP_PORT = 587
SMTP_USER = 'shubham.singh@vservit.com'
SMTP_PASSWORD = 'Mail@2023'

EMAILS = {
    "Abdusslam": "Abdusslam@vservit.com",
    "Abhijeet": "abhijit.sarkar@vservit.com",
    "Adesh": "Adesh@vservit.com",
    "Anwesha": "anwesha.palit@vservit.com",
    "Dharmender": "d.singh@vservit.com",
    "Kanika": "kanika.mathur@vservit.com",
    "Khushtar": "khushtar@vservit.com",
    "Komal": "komal.gautam@vservit.com",
    "Mahesh": "ho.ops@vservit.com",
    "Monika": "monika.jajoria@vservit.com",
    "Neha": "chd.cc@vservit.com",
    "Nishant": "nishant.kumar@vservit.com",
    "Pawan": "pawan.tiwari@vservit.com",
    "Prahlad": "prahlad.prasad@vservit.com",
    "Pramod": "Pramod@vservit.com",
    "Ravinder": "ravinder.singh@vservit.com",
    "Sanjay": "sanjay.mishra@vservit.com",
    "Sashikant": "shashikant@vservit.com",
    "Shahwar Yar Khan": "shahwar.khan@vservit.com",
    "Shivam": "shivam@vservit.com",
    "Surbhi": "surbhi.singh@vservit.com",
    "Swapnil": "swapnil.shrivastava@vservit.com",
    "Danish":"danish.naseem@vservit.com"
}

@app.route("/", methods=["GET", "POST"])
def index():
    descriptions = []
    temp_file_path = None
    message = None
    
    if request.method == "POST":
        if 'file' in request.files:
            file = request.files['file']
            df = pd.read_excel(file)
            descriptions = df['DISCRIPTION'].unique().tolist()
            temp_file_path = 'temp_file.xlsx'
            df.to_excel(temp_file_path, index=False)
        
        elif 'description' in request.form:
            selected_email = request.form['email']
            selected_date = request.form['date']
            selected_description = request.form.getlist('description')
            temp_file_path = request.form['temp_file_path']
            
            df = pd.read_excel(temp_file_path)
            success = True
            try:
                for description in selected_description:
                    person_data = df[df['DISCRIPTION'] == description]
                    file_name = 'Paid_data.xlsx'
                    person_data.drop(columns=['Email'], inplace=True, errors='ignore')
                    person_data.to_excel(file_name, index=False)

                    subject = f"Data for {description}"
                    body = f"Hello {description},\n\nPlease find the attached document : Date: {selected_date}\n\nBest regards,\nShubham Singh \n "
                    send_email(selected_email, subject, body, file_name)

                    os.remove(file_name)

                flash('Emails sent successfully!', 'success')
            except Exception as e:
                success = False
                flash(f'Failed to send emails: {e}', 'error')
            
            os.remove(temp_file_path)
            return redirect(url_for('index'))

    return render_template("index.html", emails=EMAILS, descriptions=descriptions, temp_file_path=temp_file_path, message=message)


def send_email(to_email, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    if attachment_path:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
            msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, to_email, msg.as_string())
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
        raise

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
