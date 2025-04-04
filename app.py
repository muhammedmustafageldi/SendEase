import threading
import queue
from customtkinter import *
from email_sender import send_email
from PIL import Image, ImageTk
from tkinter import messagebox
from customtkinter import ThemeManager
from custom_items import AttachmentsScrollableFrame, ReceiverScrollableFrame
from file_transactions import import_emails_from_txt, import_emails_from_xls


class App(CTk):

    def __init__(self):
        super().__init__()

        # Global variables
        self.sender_email_entry = None
        self.sender_password_entry = None
        self.subject_entry = None
        self.body_text_box = None
        self.to_list_text_box = None
        self.send_button = None
        self.progress_bar = None
        self.select_file_button = None
        self.import_button = None
        self.receiver_label = None

        self.import_queue = queue.Queue()
        self.receiver_queue = None
        self.attachments_list = None
        self.to_list_title = None
        self.import_dialog = None
        self.result_dialog = None
        self.selected_files = []  # List of selected files

        # Title
        self.title("SendEase")

        # App icon
        self.iconbitmap('assets/app_icon.ico')

        ico = Image.open('assets/app_icon.png')
        photo = ImageTk.PhotoImage(ico)
        self.wm_iconphoto(True, photo)

        # Set window size
        window_width = 800
        window_height = 800

        # Block resizable
        #self.resizable(False, False)

        # Get window position as center
        #position_x, position_y = self.calculate_window_position(window_width, window_height)
        #self.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

        # Set window full size
        self.geometry(f"{self.winfo_screenwidth()}x{self.winfo_screenheight()}+0+0")

        # Set theme
        ThemeManager.load_theme("green")
        set_appearance_mode("dark")

        self.create_frames()


    # Method that ensures the window opens in the center of the screen
    def calculate_window_position(self, window_width, window_height):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        position_x = int((screen_width / 2) - (window_width / 2))
        position_y = int((screen_height / 2) - (window_height / 2))

        return position_x, position_y

    # Method that creates and places upper and lower frames
    def create_frames(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=8)
        self.grid_columnconfigure(0, weight=1)

        self.upper_frame()
        self.center_frame()
        self.lower_frame()

    # Brand and logo
    def upper_frame(self):
        upper_frame = CTkFrame(self)
        upper_frame.grid(row=0, column=0, sticky="nsew")

        logo_image = Image.open("assets/logo.png")
        logo = CTkImage(logo_image, size=(90, 90))

        CTkLabel(upper_frame, image=logo, text="").pack(pady=5)

    def center_frame(self):
        center_frame = CTkFrame(self)
        center_frame.grid(row=1, column=0, sticky="nsew")

        CTkLabel(center_frame, text="Login Information", font=("Arial", 20, "bold"), anchor="center",
                 text_color="#f9cd59", pady=5).pack(pady=5, padx=20)

        CTkLabel(center_frame, text="Email and (App Password):", font=("Arial", 14)).pack()

        self.sender_email_entry = CTkEntry(center_frame, placeholder_text="Enter your email", width=180)
        self.sender_email_entry.pack(pady=5, padx=10)

        self.sender_password_entry = CTkEntry(center_frame, placeholder_text="Enter your app password", show="*",
                                              width=180)
        self.sender_password_entry.pack(pady=5, padx=10)

    def lower_frame(self):
        lower_frame = CTkFrame(self)
        lower_frame.grid(row=2, column=0, sticky="nsew")

        CTkLabel(lower_frame, text="Mail Information", font=("Arial", 20, "bold"), anchor="center",
                 text_color="#f9cd59", pady=5).pack(pady=5, padx=20)

        # Subject
        subject_title = CTkLabel(lower_frame, text="Subject:", font=("Arial", 14)).pack()

        self.subject_entry = CTkEntry(lower_frame, placeholder_text="Enter the subject of the mail...", width=200)
        self.subject_entry.pack(pady=10)

        # Body frame
        body_frame = CTkFrame(lower_frame, fg_color="transparent")
        body_frame.pack(fill="x")

        body_frame_left = CTkFrame(body_frame)
        body_frame_left.pack(side="left", expand=True, fill="both")

        body_frame_right = CTkFrame(body_frame)
        body_frame_right.pack(side="right", expand=True, fill="both")

        # Mail body
        CTkLabel(body_frame_left, text="Mail body:", font=("Arial", 14), anchor="w").pack(pady=5, fill="x", padx=25)

        self.body_text_box = CTkTextbox(body_frame_left, corner_radius=16, border_color="#FFCC70", border_width=2,
                                        height=120,
                                        width=300)
        self.body_text_box.pack(fill="x", padx=20)

        # Send to
        right_upper_frame = CTkFrame(body_frame_right, fg_color="transparent")
        right_upper_frame.pack(padx=25, fill="x")

        self.to_list_title = CTkLabel(right_upper_frame, text="Send to: (0)", font=("Arial", 14))
        self.to_list_title.pack(pady=5, side="left")

        # Load import img
        import_file = Image.open("assets/import.png")
        import_icon = CTkImage(import_file, size=(20, 20))

        self.import_button = CTkButton(right_upper_frame, text="", image=import_icon, compound="left",
                                       fg_color="#FFCC70",
                                       width=20,
                                       height=20, hover_color="#f2a30f", command=self.show_import_dialog)
        self.import_button.pack(side="right", padx=5)

        self.to_list_text_box = CTkTextbox(body_frame_right, corner_radius=16, border_color="#FFCC70", border_width=2,
                                           height=120, width=300)
        self.to_list_text_box.pack(fill="x", padx=20)

        self.to_list_text_box.bind("<FocusOut>", self.update_line_count)

        CTkLabel(body_frame_right, text="⚠️ Each line will be considered as a mail address.").pack()

        # Attachments

        # Attachments icon
        attachments_file = Image.open("assets/attached_file.png")
        attachments_icon = CTkImage(attachments_file, size=(20, 20))

        # Title
        CTkLabel(lower_frame, text="Attachments (Optional) :", image=attachments_icon,
                 compound="left", font=("Arial", 14), padx=10).pack(pady=5)

        self.attachments_list = AttachmentsScrollableFrame(lower_frame, width=300, height=75)
        self.attachments_list.pack()

        self.select_file_button = CTkButton(lower_frame, text="Select file", fg_color="#f2b23f", text_color="white",
                                            border_width=2, border_color="#f5a002", hover_color="#f2a30f",
                                            command=self.select_file_listener)
        self.select_file_button.pack(padx=20, pady=10)

        self.send_button = CTkButton(lower_frame, height=35, text="Send", fg_color="#f2b23f", text_color="black",
                                     border_width=2,
                                     border_color="black", hover_color="#f2a30f", font=("Arial", 17, "bold"),
                                     command=self.send_mail)
        self.send_button.pack(padx=20, pady=5, fill="x", expand=True)

        self.receiver_label = CTkLabel(lower_frame, text="")

        # Progress bar
        self.progress_bar = CTkProgressBar(lower_frame, mode="indeterminate", progress_color="#f2b23f",
                                           border_color="black", border_width=2, height=25)

    def select_file_listener(self):
        file = filedialog.askopenfilename(title="Select File", filetypes=[("All Files", "*.*")])

        if file:
            self.selected_files.append(file)
            self.update_attachments(file)

    def update_line_count(self, event=None):
        # Get all rows and filter non-blank ones
        lines = self.to_list_text_box.get("1.0", "end-1c").split("\n")
        non_empty_lines = [line for line in lines if line.strip() != ""]

        # Count non-empty lines and update the header
        line_count = len(non_empty_lines)
        self.to_list_title.configure(text=f"Send to: ({line_count})")

    def update_attachments(self, file_path):
        if self.selected_files:
            self.attachments_list.add_item(item=file_path)

    def show_import_dialog(self):
        if self.import_dialog is None or not self.import_dialog.winfo_exists():
            self.import_dialog = CTkToplevel(self)

            window_width = 550
            window_height = 450

            # Center the dialog on the screen
            x, y = self.calculate_window_position(window_width=window_width, window_height=window_height)
            self.import_dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
            self.import_dialog.title("Import emails")

            # Make the dialog modal (stay on top)
            self.import_dialog.grab_set()

            # Window icon
            self.import_dialog.after(200, lambda: self.import_dialog.iconbitmap('assets/app_icon.ico'))

            icon_file = Image.open("assets/import.png")
            import_icon = CTkImage(icon_file, size=(100, 100))

            CTkLabel(self.import_dialog, image=import_icon, text="").pack(padx=10, pady=10)

            # Description
            CTkLabel(self.import_dialog,
                     text="You can import your email recipients from a file. File types supported for now:",
                     justify="left", anchor="w", wraplength=460, font=("Arial", 15), pady=5).pack(padx=10, pady=10)

            # Supported Files ->
            supported_files_frame = CTkFrame(self.import_dialog, fg_color="transparent")
            supported_files_frame.pack()

            selected_import_file_type = StringVar(value="txt")

            def update_warning_type():
                selected_type = selected_import_file_type.get()

                if selected_type == "txt":
                    warning_label.configure(
                        text="⚠️ TXT file must contain one email per line.\nEmpty lines will be ignored.")
                elif selected_type == "xls":
                    warning_label.configure(
                        text="⚠️ Please ensure that your file contains a column with email addresses. The app will automatically detect the column with email addresses and extract valid ones. Clear and recognizable column headers will help ensure accurate email extraction.")

            # Txt ->
            txt_frame = CTkFrame(supported_files_frame, fg_color="transparent")
            txt_frame.pack(side="left", padx=5)

            txt_image = Image.open("assets/txt.png")
            txt_icon = CTkImage(txt_image, size=(75, 75))

            txt_radio_button = CTkRadioButton(txt_frame, text="", hover_color="#f2b23f", fg_color="#f2b23f",
                                              value="txt",
                                              variable=selected_import_file_type, command=update_warning_type)
            txt_radio_button.pack()

            CTkLabel(txt_frame, text="", image=txt_icon).pack()

            # Excel ->
            excel_frame = CTkFrame(supported_files_frame, fg_color="transparent")
            excel_frame.pack(side="right", padx=5)

            excel_image = Image.open("assets/xls.png")
            excel_icon = CTkImage(excel_image, size=(75, 75))

            excel_radio_button = CTkRadioButton(excel_frame, text="", hover_color="#f2b23f", fg_color="#f2b23f",
                                                value="xls", variable=selected_import_file_type,
                                                command=update_warning_type)
            excel_radio_button.pack()

            CTkLabel(excel_frame, text="", image=excel_icon).pack()

            warning_label = CTkLabel(self.import_dialog,
                                     text="⚠️ TXT file must contain one email per line.\nEmpty lines will be ignored.",
                                     text_color="#FFCC70", font=("Arial", 12), pady=5, wraplength=480)
            warning_label.pack(pady=20)

            # Select file button
            CTkButton(self.import_dialog, text="Select file", fg_color="#f2b23f", text_color="white",
                      border_width=2, border_color="#f5a002", hover_color="#f2a30f",
                      command=lambda: self.import_emails_threaded(selected_import_file_type.get())).pack(padx=20,
                                                                                                         pady=10)
        else:
            self.import_dialog.focus()

    def import_emails_threaded(self, file_type):
        thread = threading.Thread(target=self.import_emails, args=(file_type,), daemon=True)
        thread.start()
        self.check_import_queue()

    def import_emails(self, file_type):
        email_list = []
        if file_type == "txt":
            email_list = import_emails_from_txt()
        elif file_type == "xls":
            email_list = import_emails_from_xls()

        self.import_queue.put(email_list)

    def send_mail(self):
        sender_email = self.sender_email_entry.get().strip()
        sender_password = self.sender_password_entry.get()
        subject = self.subject_entry.get().strip()
        mail_body = self.body_text_box.get("1.0", "end-1c").strip()
        to_list = self.to_list_text_box.get("1.0", "end-1c").strip().strip().split("\n")

        missing_fields = []

        if not sender_email:
            missing_fields.append("Sender Email")
        if not sender_password:
            missing_fields.append("Sender Password")
        if not subject:
            missing_fields.append("Subject")
        if not mail_body:
            missing_fields.append("Mail Body")
        if not to_list:
            missing_fields.append("Recipient List")

        if missing_fields:
            missing_text = "\n".join(missing_fields)
            messagebox.showerror(title="Error", message=f"The following fields are required:\n{missing_text}")
            return

        self.loading(loading=True)

        self.receiver_queue = queue.Queue()

        thread = threading.Thread(target=self.send_mail_threaded, args=(
            sender_email, sender_password, to_list, subject, mail_body
        ), daemon=True)
        thread.start()

        self.check_receiver_queue()

    def send_mail_threaded(self, sender_email, sender_password, to_list, subject, mail_body):
        self.receiver_queue = queue.Queue()
        result = send_email(sender_email, sender_password, to_list, subject, mail_body, self.selected_files,
                            self.receiver_queue)

        self.handle_response(result=result)

        self.loading(loading=False)

    def check_import_queue(self):
        try:
            # Get data from queue
            email_list = self.import_queue.get_nowait()
            if email_list and len(email_list) != 0:
                self.to_list_text_box.insert("end", "\n".join(email_list))
                self.to_list_text_box.insert("end", "\n")
                self.import_dialog.destroy()
                self.update_line_count()
            else:
                messagebox.showerror(title="Error", message="An unknown error occurred. Please try again.")
        except queue.Empty:
            self.after(100, self.check_import_queue)

    def check_receiver_queue(self):
        try:
            while not self.receiver_queue.empty():
                message = self.receiver_queue.get_nowait()
                self.receiver_label.configure(text=message)
            self.after(100, self.check_receiver_queue)
        except queue.Empty:
            pass

    def handle_response(self, result):
        # Show result dialog
        if self.result_dialog is None or not self.result_dialog.winfo_exists():
            self.result_dialog = CTkToplevel(self)

            if result:
                # Open result pack
                state = result.get('state')
                title = result.get('title')
                desc = result.get('desc')
                success_list = result.get('success_list')
                fail_list = result.get('fail_list')

                # Set dialog title
                self.result_dialog.title(state)

                # Make the dialog modal (stay on top)
                self.result_dialog.grab_set()

                # Window icon
                self.result_dialog.after(200, lambda: self.import_dialog.iconbitmap('assets/app_icon.ico'))

                CTkLabel(self.result_dialog, text=title, font=("Arial", 16)).pack(padx=10, pady=10)

                # Set result img
                result_img_file = 'assets/successful.png' if state == 'Success' else 'assets/error.png'
                result_img = Image.open(result_img_file)
                result_icon = CTkImage(result_img, size=(100, 100))

                CTkLabel(self.result_dialog, image=result_icon, text="").pack(padx=10, pady=10)

                # Set description
                CTkLabel(self.result_dialog, text=desc, font=("Arial", 15), wraplength=450, pady=5).pack(pady=5)

                # Create lists if state is successful
                if state == "Success":
                    # List frame ->
                    list_frame = CTkFrame(self.result_dialog, fg_color='transparent')
                    list_frame.pack()

                    # Success list frame
                    success_frame = CTkFrame(list_frame, fg_color='transparent')
                    success_frame.pack(side='left')

                    CTkLabel(success_frame, text='Successful recipient list:', text_color='#00ce8e').pack(fill='x')

                    # Create custom scroll view
                    ReceiverScrollableFrame(success_frame, width=200, height=100, items=success_list).pack()

                    # Fail list frame
                    fail_frame = CTkFrame(list_frame, fg_color='transparent')
                    fail_frame.pack(side='right')

                    CTkLabel(fail_frame, text='Fail recipient list:', text_color='red').pack(fill='x')

                    ReceiverScrollableFrame(fail_frame, width=200, height=100, items=fail_list).pack()

                CTkButton(self.result_dialog, text='Okay', fg_color="#f2b23f", text_color="white",
                          border_width=2, border_color="#f5a002", hover_color="#f2a30f",
                          command=lambda: self.result_dialog.destroy()).pack(fill='x', pady=10, padx=10)

                def update_result_dialog_size():
                    self.result_dialog.update_idletasks()
                    width = 500
                    height = self.result_dialog.winfo_reqheight()

                    # Center the dialog on the screen
                    x, y = self.calculate_window_position(window_width=width, window_height=height)
                    self.result_dialog.geometry(f'{width}x{height}+{x}+{y}')

                self.result_dialog.after(100, update_result_dialog_size)

    def loading(self, loading):
        if loading:
            # Hide button and show progress
            self.send_button.pack_forget()
            self.receiver_label.pack(padx=20, pady=5)
            self.progress_bar.pack(fill="x", padx=20, pady=5)
            self.progress_bar.start()
            self.import_button.configure(state="disabled")
            self.select_file_button.configure(state="disabled")
        else:
            # Hide progress and show button
            self.send_button.pack(padx=20, pady=10, fill="x", expand=True)
            self.receiver_label.pack_forget()
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.import_button.configure(state="normal")
            self.select_file_button.configure(state="normal")