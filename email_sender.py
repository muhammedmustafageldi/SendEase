import os
from queue import Queue
import smtplib
import mimetypes
from email.message import EmailMessage


def send_email(sender_mail, sender_password, receiver_list, subject, body, attachments, queue: Queue):
    success_list = []
    fail_list = []
    try:
        # Connect server
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(user=sender_mail, password=sender_password)
            for receiver in receiver_list:
                # Create a new message for each recipient
                msg = EmailMessage()
                msg["From"] = sender_mail
                msg["To"] = receiver
                msg["Subject"] = subject
                msg.set_content(body)

                # Add attachments
                if attachments:
                    for file_path in attachments:
                        if not os.path.isfile(file_path):
                            print(f"File is not found: {file_path}")
                            continue

                        mime_type, _ = mimetypes.guess_type(file_path)  # Predict MIME type
                        if mime_type is None:
                            mime_type = "application/octet-stream"  # For unknown file types

                        main_type, sub_type = mime_type.split("/", 1)

                        with open(file_path, "rb") as file:
                            file_data = file.read()
                            file_name = os.path.basename(file_path)

                        # Add attachment
                        msg.add_attachment(file_data, maintype=main_type, subtype=sub_type, filename=file_name)

                try:
                    # Send mail ->
                    server.send_message(msg)
                    # Success
                    success_list.append(receiver)
                    queue.put(f"Email sent successfully: {receiver}")
                except Exception as e:
                    fail_list.append(receiver)
                    queue.put(f"Failed to send email to {receiver}: {e}")

    except smtplib.SMTPAuthenticationError:
        result = create_result(state='Fail', title='Authentication Error',
                               desc="Your email provider may require an App Password.\nGo to your email account settings, generate an App Password, and use it instead of your regular password.",
                               success_list=success_list, fail_list=fail_list)
        return result
    except Exception as e:
        result = create_result(
            state='Fail', title='Error', desc=f"An error occurred\n{e}",
            success_list=success_list, fail_list=fail_list)
        return result
    # Return result to app ->
    result = create_result(state='Success', title='Successfully completed.',
                           desc=f"The mail was successfully sent to {len(success_list)} recipients.",
                           success_list=success_list, fail_list=fail_list)
    return result


def create_result(state, title, desc, success_list, fail_list) -> dict:
    result = dict.fromkeys(['state', 'title', 'desc', 'success_list', 'fail_list'], None)

    result['state'] = state
    result['title'] = title
    result['desc'] = desc
    result['success_list'] = success_list
    result['fail_list'] = fail_list
    return result
