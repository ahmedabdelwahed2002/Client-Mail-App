import customtkinter as ctk
import re
from tkinter import messagebox
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import imaplib
import email



# Function to send email
def send_email(to_email, subject, body, sender_email, password):
    # Create a MIME object
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = to_email
    message['Subject'] = subject

    # Attach the body to the MIME object
    message.attach(MIMEText(body, 'plain'))

    # Connect to outlook's SMTP server
    with smtplib.SMTP('smtp.office365.com', 587) as server:
        server.starttls()  # Secure the connection
        server.login(sender_email, password)
        server.sendmail(sender_email, to_email, message.as_string())

# Send button action
def send():
    to_email = to_email_entry.get()
    subject = subject_entry.get()
    body = body_textbox.get("1.0", "end-1c")

    # Check if all fields are filled
    if not to_email or not subject or not body:
        messagebox.showwarning("Send Email", "All fields are required.")
        return

    # Sender email and password
    sender_email = email_entry.get()
    password = password_entry.get()

    try:
        send_email(to_email, subject, body, sender_email, password)
        messagebox.showinfo("Send Email", "Email sent successfully.")
        clear_fields()
    except Exception as e:
        messagebox.showerror("Send Email", f"An error occurred: {e}")

# Email validation function
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None

# Show email fields
def show_email_components():
    to_email_label.pack(pady=(10, 0))
    to_email_entry.pack()
    subject_label.pack(pady=(10, 0))
    subject_entry.pack()
    body_label.pack(pady=(10, 0))
    body_textbox.pack()
    send_button.pack(pady=20)

# Clear all fields
def clear_fields():
    to_email_entry.delete(0, 'end')
    subject_entry.delete(0, 'end')
    body_textbox.delete('1.0', 'end')

# Login button action

# Hide all components
def hide_all_components():
    to_email_label.pack_forget()
    to_email_entry.pack_forget()
    subject_label.pack_forget()
    subject_entry.pack_forget()
    body_label.pack_forget()
    body_textbox.pack_forget()
    send_button.pack_forget()
    last_email_textbox.pack_forget()


def show_last_email_components():
    hide_all_components()

    # Set up IMAP server connection details
    imap_server = "imap-mail.outlook.com"
    email_address = email_entry.get()
    password = password_entry.get()

    try:
        # Connect to the IMAP server
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(email_address, password)
        imap.select("Inbox")

        # Fetch the latest email
        result, data = imap.search(None, "ALL")
        if result == "OK":
            # Get the last email ID
            latest_email_id = data[0].split()[-1]
            result, data = imap.fetch(latest_email_id, "(RFC822)")

            if result == "OK":
                # Parse the email content
                raw_email = data[0][1]
                message = email.message_from_bytes(raw_email)

                # Extract email details
                from_header = message["From"]
                to_header = message["To"]
                subject_header = message["Subject"]

                # Initialize email body
                email_body = ""
                if message.is_multipart():
                    for part in message.walk():
                        ctype = part.get_content_type()
                        cdispo = str(part.get('Content-Disposition'))

                        if ctype == 'text/plain' and 'attachment' not in cdispo:
                            email_body = part.get_payload(decode=True).decode('utf-8')  # Decode byte string
                            break
                else:
                    email_body = message.get_payload(decode=True).decode('utf-8')

                # Clear any previous content
                last_email_textbox.delete('1.0', 'end')

                # Format and insert new content
                formatted_email = f"From: {from_header}\nTo: {to_header}\nSubject: {subject_header}\n\n{email_body}"
                last_email_textbox.insert('1.0', formatted_email)

                last_email_textbox.pack(pady=20)
            else:
                raise Exception("Failed to fetch the latest email.")
        else:
            raise Exception("Failed to search emails in the inbox.")

    except Exception as e:
        messagebox.showerror("Fetch Email", f"An error occurred: {e}")
    finally:
        # Logout from IMAP server
        imap.logout()

    hide_all_components()
    last_email_textbox.pack(pady=20)

# Show email send and receive options after login
def show_email_options():
    hide_all_components()
    send_email_button = ctk.CTkButton(master, text="Send Email", command=show_send_email_components)
    receive_email_button = ctk.CTkButton(master, text="Show Last Email Received", command=show_last_email_components)
    send_email_button.pack(pady=(10, 20))
    receive_email_button.pack(pady=(10, 20))
def login_action():
    email = email_entry.get()
    password = password_entry.get()
    if not email or not password:
        messagebox.showwarning("Login Failed", "All fields are required")
        return
    elif not is_valid_email(email):
        messagebox.showwarning("Login Failed", "Please enter a valid email.")
        return
    try:
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(email, password)
            show_email_options()
    except smtplib.SMTPAuthenticationError:
        messagebox.showwarning("Login Failed", "Incorrect email or password.")
    except Exception as e:
        messagebox.showerror("Login Failed", f"An error occurred: {e}")

def show_send_email_components():
    hide_all_components()
    to_email_label.pack(pady=(10, 0))
    to_email_entry.pack()
    subject_label.pack(pady=(10, 0))
    subject_entry.pack()
    body_label.pack(pady=(10, 0))
    body_textbox.pack()
    send_button.pack(pady=20)
# GUI setup
master = ctk.CTk()
master.geometry("800x600")
master.title("Client Mail App")
ctk.set_appearance_mode("dark")

# Email entry
email_label = ctk.CTkLabel(master, text="Email:")
email_label.pack(pady=(10, 0))
email_entry = ctk.CTkEntry(master, width=350, placeholder_text="Enter your email")
email_entry.pack()

# Password entry
password_label = ctk.CTkLabel(master, text="Password:")
password_label.pack(pady=(10, 0))
password_entry = ctk.CTkEntry(master, width=350, placeholder_text="Enter your password", show="*")
password_entry.pack()

# Login button
login_button = ctk.CTkButton(master, text="Login", command=login_action)
login_button.pack(pady=20)

# Recipient (To) entry
to_email_label = ctk.CTkLabel(master, text="To:")
to_email_entry = ctk.CTkEntry(master, width=350, placeholder_text="Recipient's email")

# Subject entry
subject_label = ctk.CTkLabel(master, text="Subject:")
subject_entry = ctk.CTkEntry(master, width=350, placeholder_text="Email subject")

# Email body text box
body_label = ctk.CTkLabel(master, text="Body:")
body_textbox = ctk.CTkTextbox(master, width=350, height=100)

# Send button
send_button = ctk.CTkButton(master, text="Send Email", command=send)

# Textbox for displaying the last received email
last_email_textbox = ctk.CTkTextbox(master, width=350, height=200)


# Initially, do not display the email components
hide_all_components()

# Run the GUI loop
master.mainloop()
