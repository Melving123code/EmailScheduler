import smtplib
import schedule
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import configparser
import logging
from tkinter import Tk, Label, Entry, Button, StringVar, messagebox, OptionMenu, Text
from threading import Thread

# Set up logging
logging.basicConfig(filename='email_scheduler.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

config = configparser.ConfigParser()

# Function to save configuration
def save_config():
    config['email'] = {
        'address': email_var.get(),
        'password': password_var.get()
    }
    config['smtp'] = {
        'server': "smtp.gmail.com",
        'port': 587
    }
    config['recipient'] = {
        'emails': recipient_var.get("1.0", "end-1c"),
        'subject': subject_var.get(),
        'message': message_var.get("1.0", "end-1c"),
        'time_to_send': f"{time_var.get()} {am_pm_var.get()}"
    }
    with open('config.ini', 'w') as configfile:
        config.write(configfile)

# Function to load configuration
def load_config():
    config.read('config.ini')
    if 'email' in config:
        email_var.set(config['email'].get('address', ''))
        password_var.set(config['email'].get('password', ''))
    if 'recipient' in config:
        recipient_var.delete("1.0", "end")
        recipient_var.insert("1.0", config['recipient'].get('emails', ''))
        subject_var.set(config['recipient'].get('subject', ''))
        message_var.delete("1.0", "end")
        message_var.insert("1.0", config['recipient'].get('message', ''))
        time_to_send = config['recipient'].get('time_to_send', '12:00 AM')
        try:
            time, am_pm = time_to_send.split()
        except ValueError:
            time, am_pm = '12:00', 'AM'
        time_var.set(time)
        am_pm_var.set(am_pm)

# Function to convert 12-hour format to 24-hour format
def convert_to_24_hour_format(time_str, am_pm):
    hour, minute = map(int, time_str.split(':'))
    if am_pm == "PM" and hour != 12:
        hour += 12
    if am_pm == "AM" and hour == 12:
        hour = 0
    return f"{hour:02}:{minute:02}"

# Function to send email
def send_email(email_address, email_password, smtp_server, smtp_port, recipient_emails, subject, message):
    msg = MIMEMultipart()
    msg['From'] = email_address
    msg['To'] = ", ".join(recipient_emails)
    msg['Subject'] = subject

    msg.attach(MIMEText(message, 'plain'))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email_address, email_password)
        text = msg.as_string()
        server.sendmail(email_address, recipient_emails, text)
        server.quit()
        logging.info(f"Email sent successfully to {', '.join(recipient_emails)}")
        messagebox.showinfo("Success", f"Email sent successfully to {', '.join(recipient_emails)}")
    except smtplib.SMTPAuthenticationError:
        logging.error("Failed to send email: Authentication error.")
        messagebox.showerror("Error", "Failed to send email: Authentication error. Please check your email and password.")
    except Exception as e:
        logging.error(f"Failed to send email to {', '.join(recipient_emails)}: {e}")
        messagebox.showerror("Error", f"Failed to send email: {e}")

# Function to schedule email
def schedule_email():
    email_address = email_var.get()
    email_password = password_var.get()
    recipient_emails = [email.strip() for email in recipient_var.get("1.0", "end-1c").split(",")]
    subject = subject_var.get()
    message = message_var.get("1.0", "end-1c")
    time_to_send = convert_to_24_hour_format(time_var.get(), am_pm_var.get())

    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    if not email_address or not email_password or not recipient_emails or not subject or not message or not time_to_send:
        messagebox.showerror("Error", "All fields must be filled out.")
        return

    schedule.every().day.at(time_to_send).do(send_email, email_address, email_password, smtp_server, smtp_port, recipient_emails, subject, message)
    logging.info(f"Email scheduled to be sent every day at {time_to_send} to {', '.join(recipient_emails)}")

    while True:
        schedule.run_pending()
        time.sleep(1)

# Function to start the scheduler in a new thread
def start_scheduler():
    save_config()
    thread = Thread(target=schedule_email)
    thread.daemon = True
    thread.start()

# Set up the GUI
root = Tk()
root.title("Email Scheduler")

Label(root, text="Email Address").grid(row=0, column=0, sticky="e")
email_var = StringVar()
Entry(root, textvariable=email_var, width=50).grid(row=0, column=1, padx=5, pady=5)

Label(root, text="Email Password").grid(row=1, column=0, sticky="e")
password_var = StringVar()
Entry(root, textvariable=password_var, show="*", width=50).grid(row=1, column=1, padx=5, pady=5)

Label(root, text="Recipient Emails (comma-separated)").grid(row=2, column=0, sticky="e")
recipient_var = Text(root, width=50, height=5)
recipient_var.grid(row=2, column=1, padx=5, pady=5)

Label(root, text="Subject").grid(row=3, column=0, sticky="e")
subject_var = StringVar()
Entry(root, textvariable=subject_var, width=50).grid(row=3, column=1, padx=5, pady=5)

Label(root, text="Message").grid(row=4, column=0, sticky="e")
message_var = Text(root, width=50, height=10)
message_var.grid(row=4, column=1, padx=5, pady=5)

Label(root, text="Time to Send").grid(row=5, column=0, sticky="e")
time_var = StringVar()
Entry(root, textvariable=time_var, width=10).grid(row=5, column=1, padx=5, pady=5, sticky="w")

Label(root, text="AM/PM").grid(row=5, column=1)
am_pm_var = StringVar(value="AM")
OptionMenu(root, am_pm_var, "AM", "PM").grid(row=5, column=1, padx=5, pady=5, sticky="e")

Button(root, text="Schedule Email", command=start_scheduler).grid(row=6, column=0, columnspan=2, pady=10)

load_config()

root.mainloop()