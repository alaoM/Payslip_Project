import os
import sys
#File paths
sys.path.append(os.path.join(os.path.dirname(__file__), 'libs'))
from flask import Flask, request, render_template, send_file, redirect, url_for, flash
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import logging


# Import email configuration
import email_config

logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'uploads'
EXTRACTED_FOLDER = 'extracted'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXTRACTED_FOLDER'] = EXTRACTED_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(EXTRACTED_FOLDER):
    os.makedirs(EXTRACTED_FOLDER)

# Read employee details from a CSV
employees_df = pd.read_csv('employees.csv')

def extract_pages(input_pdf, output_pdf, start_page, end_page):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for i in range(start_page - 1, end_page):
        try:
            writer.add_page(reader.pages[i])
        except IndexError:
            print(f"Error: Tried to access page {i} which does not exist in the PDF.")
            break

    with open(output_pdf, 'wb') as output_file:
        writer.write(output_file)

def send_email(receiver_email, subject, body, attachment):
    sender_email = email_config.SENDER_EMAIL
    sender_password = email_config.SENDER_PASSWORD
    smtp_server = email_config.SMTP_SERVER
    smtp_port = email_config.SMTP_PORT

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(attachment, 'rb').read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(attachment))
    msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
        server.quit()
        print('Email sent successfully.')
    except Exception as e:
        print(f'Failed to send email. Error: {e}')

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    if file:
        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)
        flash(f'File successfully uploaded: {file.filename}')
        return redirect(url_for('process_payslips', filename=file.filename))

@app.route('/process')
def process_payslips():
    filename = request.args.get('filename')
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    # For simplicity, assume PDF file for now. Extend this to handle Excel files.
    if filename.endswith('.pdf'):
        reader = PdfReader(file_path)
        number_of_pages = len(reader.pages)
        
        # Assuming each employee's payslip is 1 page, adjust this as necessary
        for i in range(number_of_pages):
            output_pdf = os.path.join(app.config['EXTRACTED_FOLDER'], f"payslip_{i+1}.pdf")
            extract_pages(file_path, output_pdf, i+1, i+2)

        # Get extracted files for display
        extracted_files = os.listdir(app.config['EXTRACTED_FOLDER'])
        return render_template('process.html', files=extracted_files)
    else:
        flash('Unsupported file format')
        return redirect(url_for('index'))

@app.route('/preview/<filename>')
def preview_file(filename):
    return send_file(os.path.join(app.config['EXTRACTED_FOLDER'], filename))

@app.route('/send_email', methods=['POST'])
def send_payslip_email():
    filename = request.form['filename']
    employee_email = request.form['employee_email']
    
    # Get employee email
    employee = employees_df.loc[employees_df['email'] == (employee_email)]
    if not employee.empty:
        email = employee['email'].values[0]
        subject = "Your Payslip"
        body = f"Dear {employee['name'].values[0]},\n\nPlease find attached your payslip.\n\nBest regards,\nYour Company"
        send_email(email, subject, body, os.path.join(app.config['EXTRACTED_FOLDER'], filename))
        flash(f'Payslip sent to {email}')
    else:
        flash('Employee not found')

    return redirect(url_for('process_payslips', filename=request.form['upload_filename']))

if __name__ == "__main__":
    app.run(debug=True)
