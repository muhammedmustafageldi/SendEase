import re
from tkinter.filedialog import askopenfilename
import pandas as pd
from tkinter import messagebox

def import_emails_from_txt():
    file_path = askopenfilename(filetypes=[("Text files", "*.txt")])

    if not file_path:
        return  # Cancel the operation if the user has not selected a file

    with open(file_path, "r", encoding="utf-8") as file:
        lines = file.readlines()


    # Clean empty lines and check email format
    emails = [line.strip() for line in lines if line.strip()]
    invalid_emails = [email for email in emails if not re.match(r"[^@]+@[^@]+\.[^@]+", email)]

    if invalid_emails:
        show_warning(title="Warning",
                     desc=f"Invalid email addresses found:\n{', '.join(invalid_emails[:5])}\n\nThey will be ignored.")

    # E-mail list
    valid_emails = [email for email in emails if email not in invalid_emails]

    if valid_emails:
        show_info_success(title="Success", desc=f"Successfully imported {len(valid_emails)} emails!")
    else:
        show_error(title="Error", desc="No valid emails found in the file.")

    return valid_emails


def import_emails_from_xls():
    file_path = askopenfilename(filetypes=[("Excel files", "*.xls"), ("Excel files", "*.xlsx")])

    if not file_path:
        return  # Cancel if the user has not selected a file

    try:
        df = pd.read_excel(file_path)  # Read Excel file
    except Exception as e:
        show_error(title="Error", desc=f"Could not read the Excel file.\n{e}")
        return

    # Require the user to select the email column
    column_names = df.columns.tolist()

    if not column_names:
        show_error(title="Error", desc="The Excel file does not contain any columns.")
        return

    # Find the column with the first available email
    email_column = None
    for col in column_names:
        if df[col].astype(str).str.contains(r"[^@]+@[^@]+\.[^@]+").any():
            email_column = col
            break

    if not email_column:
        show_error(title="Error", desc="No email column found in the file.")
        return

    # Get valid values in the Email column
    emails = df[email_column].dropna().astype(str).tolist()
    valid_emails = [email.strip() for email in emails if re.match(r"[^@]+@[^@]+\.[^@]+", email)]

    if valid_emails:
        show_info_success(title="Success", desc=f"Successfully imported {len(valid_emails)} emails!")
    else:
        show_error(title="Error", desc="No valid emails found in the file.")

    return valid_emails


def show_warning(title: str, desc: str):
    messagebox.showwarning(title, message=desc)
    #CTkMessagebox(title=title, message=desc, icon="warning", button_color="#f2b23f", button_hover_color="#f2b23f")


def show_info_success(title: str, desc: str):
    messagebox.showinfo(title, message=desc)
    #CTkMessagebox(title=title, message=desc, icon="check", button_color="#f2b23f", button_hover_color="#f2b23f")


def show_error(title: str, desc: str):
    messagebox.showerror(title, message=desc)
    #CTkMessagebox(title=title, message=desc, icon="cancel", button_color="#f2b23f", button_hover_color="#f2b23f")
