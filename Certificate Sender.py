import os
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage

def main():
    # Get user input
    certificates_folder = input("Enter the path to the certificates folder: ")
    csv_file = input("Enter the path to the CSV file: ")
    event_name = input("Enter the name of the event: ")
    sender_email = input("Enter the sender's email address: ")
    sender_password = input("Enter the sender's email password: ")

    # Email configuration
    smtp_server = "smtp.office365.com"
    smtp_port = 587

    # Define the path to the MLSA logo
    logo_filename = r"C:\Users\LENOVO\Documents\MLSA Certificate\mlsa_logo_2.png"  # Update this path as needed

    with open(csv_file, mode="r") as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            recipient_email = row["Email"]
            name = row["Name"]

            # Certificate file path
            certificate_filename = os.path.join(certificates_folder, f"{name}_Certificate.pdf")
            certificate_filename = os.path.abspath(certificate_filename)

            # Check if the certificate file exists
            if os.path.exists(certificate_filename):
                # Create the email
                msg = MIMEMultipart("related")
                msg["To"] = recipient_email
                msg["From"] = sender_email
                msg["Subject"] = f"Completion Certificate for {event_name}"

                # Email body with LinkedIn URL using HTML
                linkedin_url = "https://www.linkedin.com/in/sreya-e-p-79b915214/"
                body = f"""
                <html>
                <head></head>
                <body>
                    <div style="text-align: center; padding: 20px;">
                        <img src="cid:mlsa_logo" alt="MLSA Logo" style="width: 150px; height: auto;">
                    </div>
                    <p>Hello {name},</p>
                    <p>Thank you for participating in the <strong>{event_name}</strong> Event. Kindly find your completion certificate attached to this email.</p>
                    <p>Continue to explore, innovate, and push the boundaries of what's possible!</p>
                    <p>Best regards,</p>
                    <p><strong>Sreya E P</strong><br>Microsoft Learn Student Ambassador</p>
                    <p>LinkedIn: <a href="{linkedin_url}">{linkedin_url}</a></p>
                </body>
                </html>
                """
                msg.attach(MIMEText(body, "html"))

                # Attach the MLSA logo to the email
                with open(logo_filename, "rb") as logo_file:
                    image_part = MIMEImage(logo_file.read(), Name=os.path.basename(logo_filename))
                    image_part.add_header("Content-ID", "<mlsa_logo>")
                    image_part.add_header("Content-Disposition", "inline", filename=os.path.basename(logo_filename))
                    msg.attach(image_part)

                # Attach the existing certificate file
                with open(certificate_filename, "rb") as cert_file:
                    attachment = MIMEApplication(cert_file.read(), Name=os.path.basename(certificate_filename))
                    attachment["Content-Disposition"] = f"attachment; filename={os.path.basename(certificate_filename)}"
                    msg.attach(attachment)

                # Send the email
                try:
                    with smtplib.SMTP(smtp_server, smtp_port) as server:
                        server.starttls()
                        server.login(sender_email, sender_password)
                        server.send_message(msg)
                    print(f"Email sent successfully to {recipient_email}.")
                except Exception as e:
                    print(f"Error sending email to {recipient_email}: {e}")
            else:
                print(f"Certificate not found for {name}.")

if __name__ == "__main__":
    main()
