import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import shutil
import os
import time
from datetime import datetime, timedelta
import logging

def send_email_with_attachments_by_date(smtp_server, port, sender_email, password, subject, body, recipient_emails, source_folder):
    try:
        # Record start time
        #changes made
        start_time = time.time()
        timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        sent_files_folder = 'sent_files'  # Folder to store sent files
        timestamp_folder = f"sent_files_{timestamp}"  # Subfolder with timestamp

        # Construct absolute paths
        base_folder = os.path.abspath(os.path.dirname(__file__))  # Directory where the script is located
        sent_files_folder_path = os.path.join(base_folder, sent_files_folder)
        timestamp_folder_path = os.path.join(sent_files_folder_path, timestamp_folder)
        files_to_send = []  # variable to append the files
        log_folder_path = os.path.join(source_folder, 'log_folder')
        email_log_file = 'email_log.log'
        log_file_path = os.path.join(log_folder_path, email_log_file)  # Use a fixed log file name
        email_date_file = os.path.join(log_folder_path, 'emailDate.txt')

        # Set up logging with append mode
        logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filemode='a')
        logging.info("Python script is being processed")

        if not os.path.exists(source_folder):
            logging.error(f"Source folder '{source_folder}' does not exist.")

        # Check the date in the emailDate file
        if os.path.exists(email_date_file):
            with open(email_date_file, 'r') as date_file:
                last_execution_date_str = date_file.read().strip()
                last_execution_date = datetime.strptime(last_execution_date_str, "%Y-%m-%d")
        else:
            # If emailDate file doesn't exist, create it with the current date
            last_execution_date = datetime.now()
            with open(email_date_file, 'w') as date_file:
                date_file.write(last_execution_date.strftime("%Y-%m-%d"))

        # Calculate the difference in days
        date_difference = (datetime.now() - last_execution_date).days
        
        if date_difference > 30:
            logging.info("Deleting existing log file and updating emailDate due to date difference.")

            # Close the existing log file to release any locks
            logging.shutdown()

            # Delete the existing log file
            if os.path.exists(log_file_path):
                try:
                    with open(log_file_path, 'w'):  # Open the file in write mode to clear its contents
                        pass
                except Exception as clear_error:
                    logging.error(f"Error clearing log file: {clear_error}")

            # Update the last execution date in the emailDate file
            last_execution_date = datetime.now()
            with open(email_date_file, 'w') as date_file:
                date_file.write(last_execution_date.strftime("%Y-%m-%d"))

            # Configure logging to create a new log file
            log_folder_path = os.path.join(source_folder, 'log_folder')
            os.makedirs(log_folder_path, exist_ok=True)
            email_log_file = 'email_log.log'
            log_file_path = os.path.join(log_folder_path, email_log_file)

            logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filemode='w')

        # Create a folder for sent files if it doesn't exist
        os.makedirs(sent_files_folder_path, exist_ok=True)

        # Create a folder with a timestamp for sent files
        os.makedirs(timestamp_folder_path, exist_ok=True)

        previous_date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")

        # Attach files based on creation date
        for root, dirs, files in os.walk(source_folder):
            for file in files:
                file_path = os.path.join(root, file)
                creation_date = datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d")

                if creation_date == previous_date:  # Compare with the date part of the timestamp
                    files_to_send.append(file_path)

        # Log the count of files after collecting them
        logging.info(f"Files to be attached count: {len(files_to_send)-1}")

        if not files_to_send:
            logging.info("No files found to send.")
            return

        # Create a new MIME object for all recipients
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = ', '.join(recipient_emails)  # Join recipients into a comma-separated string

        # Attach text body
        msg.attach(MIMEText(body, 'plain'))

        # Attach files to the message
        for file_path in files_to_send:
            if file_path != log_file_path and file_path !=  email_date_file : 
                try:
                    with open(file_path, 'rb') as attachment:
                        part = MIMEBase('application', 'octet-stream')  # for processing the email with attachments
                        part.set_payload(attachment.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
                        msg.attach(part)

                except Exception as e:
                    logging.error(f"Error attaching file {file_path}: {e}")

        try:
            # Connect to the SMTP server
            with smtplib.SMTP(smtp_server, port) as server:
                server.starttls()
                server.login(sender_email, password)

                # Send the email
                server.sendmail(sender_email, recipient_emails, msg.as_string())

                logging.info(f'Email sent successfully to {", ".join(recipient_emails)}')
                
                # Move sent files to the timestamp folder after sending to all recipients
                for file_path in files_to_send:
                    if file_path != log_file_path and file_path!=email_date_file :
                        try:
                            base_name, extension = os.path.splitext(os.path.basename(file_path))
                            destination_file_path = os.path.join(timestamp_folder_path,f"{base_name}_{extension}")

                            # Append a unique identifier to the filename to avoid overwriting
                            unique_identifier = 1
                            while os.path.exists(destination_file_path):
                                new_file_path = os.path.join(timestamp_folder_path,f"{base_name}_{unique_identifier}{extension}")
                                destination_file_path = new_file_path
                                unique_identifier += 1

                            shutil.move(file_path, destination_file_path)

                        except Exception as main_error:
                            logging.error(f"Error in main process: {main_error}")
            print("Mail sent successfully!")
            

        except Exception as e:
            logging.error(f"Error sending email to {', '.join(recipient_emails)}: {e}")
            logging.info("Pausing for 5 seconds before retrying...")
            time.sleep(5)

        finally:
            # Clear attachments from the message to avoid duplicates
            msg = MIMEMultipart()

    except Exception as main_error:
        logging.error(f"Error in main process: {main_error}")

    finally:
        # Remove timestamp_folder if it's empty after processing all recipients
        if timestamp_folder_path and os.path.exists(timestamp_folder_path) and not os.listdir(
                timestamp_folder_path):
            try:
                os.rmdir(timestamp_folder_path)
                logging.info(f'Folder "{timestamp_folder_path}" removed successfully because it was empty.')
            except OSError as e:
                logging.error(f"Error removing folder '{timestamp_folder_path}': {e}")

        # Record end time and calculate duration
        end_time = time.time()
        duration = end_time - start_time
        logging.info("End of process")
        logging.info(f"Duration for completing the process: {duration} seconds\n\n")


# Call the function with your parameters
smtp_server = 'smtp.gmail.com'
port = 587
sender_email = 'mmahalakshmi@student.tce.edu'
password = 'K8N7T7A6'

subject = 'MIME Test Email'
body = 'This is a test email sent from Python.'
recipient_emails = ['muniraja@kaspl.net', 'kavi@kaspl.net']
source_folder = r"C:\Users\mahal\files to send"

send_email_with_attachments_by_date(smtp_server, port, sender_email, password, subject, body, recipient_emails, source_folder)
