import sys
import smtplib
import re
import os
import json
import warnings
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QLabel, QLineEdit, QTextEdit, QPushButton, QTabWidget, 
                             QMessageBox, QListWidget, QFileDialog, QHBoxLayout)
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Suppress the DeprecationWarning related to sip in PyQt
warnings.filterwarnings("ignore", category=DeprecationWarning)

class EmailApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set the window properties
        self.setWindowTitle("Email Sending App")
        self.setGeometry(300, 100, 800, 600)

        # Create the main layout
        self.main_layout = QVBoxLayout()

        # Create Tabs (Compose Email, Sent Logs, and Settings)
        self.tabs = QTabWidget()

        # Tab 1 - Email Composition
        self.compose_tab = QWidget()
        self.compose_layout = QVBoxLayout()

        # Fields for 'From Name', 'To', 'Subject', and Message Body
        self.from_name_label = QLabel("From Name:")
        self.from_name_field = QLineEdit()
        self.to_label = QLabel("To:")
        self.to_field = QLineEdit()
        self.subject_label = QLabel("Subject:")
        self.subject_field = QLineEdit()
        self.message_label = QLabel("Message:")
        self.message_body = QTextEdit()

        # Attach File Button and Preview
        self.attach_button = QPushButton("Attach File")
        self.attach_button.clicked.connect(self.attach_file)
        self.file_preview_label = QLabel("No file attached")

        # Send Button
        self.send_button = QPushButton("Send Email")
        self.send_button.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px; padding: 10px;")
        self.send_button.clicked.connect(self.validate_and_send_email)  # Link the button to the validation function

        # Add to Compose Layout
        self.compose_layout.addWidget(self.from_name_label)
        self.compose_layout.addWidget(self.from_name_field)
        self.compose_layout.addWidget(self.to_label)
        self.compose_layout.addWidget(self.to_field)
        self.compose_layout.addWidget(self.subject_label)
        self.compose_layout.addWidget(self.subject_field)
        self.compose_layout.addWidget(self.message_label)
        self.compose_layout.addWidget(self.message_body)

        # Layout for Attach button and file preview
        attach_layout = QHBoxLayout()
        attach_layout.addWidget(self.attach_button)
        attach_layout.addWidget(self.file_preview_label)
        self.compose_layout.addLayout(attach_layout)

        self.compose_layout.addWidget(self.send_button)
        self.compose_tab.setLayout(self.compose_layout)
        self.tabs.addTab(self.compose_tab, "Compose Email")

        # Tab 2 - Sent Logs
        self.sent_logs_tab = QWidget()
        self.sent_logs_layout = QVBoxLayout()
        self.sent_logs_label = QLabel("Sent Emails:")
        self.sent_logs_list = QListWidget()  # List widget to display sent emails

        # Button to Clear Sent Logs
        self.clear_logs_button = QPushButton("Clear Logs")
        self.clear_logs_button.clicked.connect(self.clear_sent_logs)

        # Add to Sent Logs Layout
        self.sent_logs_layout.addWidget(self.sent_logs_label)
        self.sent_logs_layout.addWidget(self.sent_logs_list)
        self.sent_logs_layout.addWidget(self.clear_logs_button)
        self.sent_logs_tab.setLayout(self.sent_logs_layout)
        self.tabs.addTab(self.sent_logs_tab, "Sent Logs")

        # Tab 3 - Settings (SMTP Configurations)
        self.settings_tab = QWidget()
        self.settings_layout = QVBoxLayout()
        self.smtp_label = QLabel("SMTP Server Settings:")
        self.smtp_host_label = QLabel("SMTP Host:")
        self.smtp_host_field = QLineEdit()
        self.smtp_port_label = QLabel("SMTP Port:")
        self.smtp_port_field = QLineEdit()
        self.smtp_user_label = QLabel("Email Address:")
        self.smtp_user_field = QLineEdit()
        self.smtp_pass_label = QLabel("Password:")
        self.smtp_pass_field = QLineEdit()
        self.smtp_pass_field.setEchoMode(QLineEdit.Password)  # Hide password input

        # Save SMTP Settings Button
        self.save_smtp_button = QPushButton("Save SMTP Settings")
        self.save_smtp_button.clicked.connect(self.save_smtp_settings)

        # Add to Settings Layout
        self.settings_layout.addWidget(self.smtp_label)
        self.settings_layout.addWidget(self.smtp_host_label)
        self.settings_layout.addWidget(self.smtp_host_field)
        self.settings_layout.addWidget(self.smtp_port_label)
        self.settings_layout.addWidget(self.smtp_port_field)
        self.settings_layout.addWidget(self.smtp_user_label)
        self.settings_layout.addWidget(self.smtp_user_field)
        self.settings_layout.addWidget(self.smtp_pass_label)
        self.settings_layout.addWidget(self.smtp_pass_field)
        self.settings_layout.addWidget(self.save_smtp_button)

        self.settings_tab.setLayout(self.settings_layout)
        self.tabs.addTab(self.settings_tab, "Settings")

        # Add Tabs to Main Layout
        self.main_layout.addWidget(self.tabs)

        # Set Main Widget
        container = QWidget()
        container.setLayout(self.main_layout)
        self.setCentralWidget(container)

        # Initialize sent email logs (in-memory)
        self.sent_emails = []

        # Variable to store the attached file path
        self.attached_file_path = None

        # Load saved SMTP settings
        self.load_smtp_settings()

    def attach_file(self):
        # Open file dialog for attachment
        file_name, _ = QFileDialog.getOpenFileName(self, "Attach File", "", "All Files (*)")
        if file_name:
            self.attached_file_path = file_name
            self.file_preview_label.setText(f"Attached: {os.path.basename(file_name)}")

    def validate_and_send_email(self):
        # Validate fields before sending
        from_name = self.from_name_field.text()
        recipient = self.to_field.text()
        subject = self.subject_field.text()
        message_body = self.message_body.toPlainText()

        # Validate email format
        if not self.is_valid_email(recipient):
            QMessageBox.warning(self, "Validation Error", "Invalid email format!")
            return

        # Check for empty fields
        if not from_name or not recipient or not subject or not message_body:
            QMessageBox.warning(self, "Validation Error", "All fields (From Name, To, Subject, and Message) are required!")
            return

        # Proceed to send email
        self.send_email(from_name, recipient, subject, message_body)

    def is_valid_email(self, email):
        # Basic regex pattern for validating an email
        email_regex = r"[^@]+@[^@]+\.[^@]+"
        return re.match(email_regex, email)

    def send_email(self, from_name, recipient, subject, message_body):
        # Get SMTP settings
        smtp_host = self.smtp_host_field.text()
        smtp_port = int(self.smtp_port_field.text())
        smtp_user = self.smtp_user_field.text()
        smtp_pass = self.smtp_pass_field.text()

        # Create message
        msg = MIMEMultipart()
        msg['From'] = f"{from_name} <{smtp_user}>"
        msg['To'] = recipient
        msg['Subject'] = subject

        # Attach the message body
        html_text = f"<html><body>{message_body}</body></html>"
        msg.attach(MIMEText(html_text, 'html'))

        # Attach file if one is selected
        if self.attached_file_path:
            attachment = open(self.attached_file_path, "rb")
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(self.attached_file_path)}")
            msg.attach(part)
            attachment.close()

        try:
            # Connect to SMTP server and send the email
            server = smtplib.SMTP(smtp_host, smtp_port)
            server.starttls()  # Secure the connection
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, recipient, msg.as_string())
            server.quit()

            # Notify user of success
            QMessageBox.information(self, "Success", "Email sent successfully!")

            # Log sent email details in memory
            self.log_sent_email(recipient, subject, message_body)

        except Exception as e:
            # Handle errors and notify the user
            QMessageBox.critical(self, "Error", f"Failed to send email: {str(e)}")

    def log_sent_email(self, recipient, subject, message_body):
        # Add log entry to in-memory list
        log_entry = f"To: {recipient}, Subject: {subject}, Snippet: {message_body[:50]}..."
        self.sent_emails.append(log_entry)
        self.sent_logs_list.addItem(log_entry)  # Update the sent logs display

    def clear_sent_logs(self):
        # Clear the sent emails log
        self.sent_emails.clear()
        self.sent_logs_list.clear()

    def save_smtp_settings(self):
        # Save SMTP settings to local file (JSON)
        smtp_settings = {
            "smtp_host": self.smtp_host_field.text(),
            "smtp_port": self.smtp_port_field.text(),
            "smtp_user": self.smtp_user_field.text(),
            "smtp_pass": self.smtp_pass_field.text()
        }
        with open("smtp_settings.json", "w") as file:
            json.dump(smtp_settings, file)

        QMessageBox.information(self, "Success", "SMTP settings saved successfully!")

    def load_smtp_settings(self):
        # Load SMTP settings from local file (if exists)
        if os.path.exists("smtp_settings.json"):
            with open("smtp_settings.json", "r") as file:
                smtp_settings = json.load(file)
                self.smtp_host_field.setText(smtp_settings.get("smtp_host", ""))
                self.smtp_port_field.setText(smtp_settings.get("smtp_port", ""))
                self.smtp_user_field.setText(smtp_settings.get("smtp_user", ""))
                self.smtp_pass_field.setText(smtp_settings.get("smtp_pass", ""))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = EmailApp()
    window.show()
    sys.exit(app.exec_())
