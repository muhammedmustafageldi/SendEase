from customtkinter import *
from PIL import Image


class AttachmentsScrollableFrame(CTkScrollableFrame):
    def __init__(self, master, command=None, width=200, height=200, **kwargs):
        super().__init__(master, width=width, height=height, scrollbar_button_color="#f5a002",
                         scrollbar_button_hover_color="black", **kwargs)

        self.command = command
        self.item_frames = []
        self.path_values = []

        self._scrollbar.configure(height=0)

        item_file = Image.open("assets/remove.png")
        self.item_icon = CTkImage(item_file, size=(20, 20))

    def add_item(self, item):
        item_frame = CTkFrame(self, fg_color="transparent")
        item_frame.pack(fill="x", padx=5, pady=2)

        item_label = CTkLabel(item_frame, text=os.path.basename(item), font=("Arial", 16, "bold"), pady=5)
        item_label.pack(side="left")

        remove_button = CTkButton(item_frame, image=self.item_icon, text="", width=20, height=20, fg_color="#f2b23f",
                                  border_width=2, border_color="#f5a002", hover_color="#f2a30f",
                                  command=lambda: self.remove_item(item_frame, item))
        remove_button.pack(side="right")

        self.item_frames.append(item_frame)
        self.path_values.append(item)

    def remove_item(self, item_frame, item):
        item_frame.destroy()
        self.item_frames.remove(item_frame)
        self.path_values.remove(item)

    def get_path_values(self):
        return self.path_values

class ReceiverScrollableFrame(CTkScrollableFrame):

    def __init__(self, master, width, height, items):
        super().__init__(master, width=width, height=height, scrollbar_button_color="#f5a002",
                         scrollbar_button_hover_color="black")

        self._scrollbar.configure(height=0)

        # Add items
        for item in items:
            CTkLabel(self, text=item, font=('Arial', 10, 'bold'), pady=3, fg_color='transparent', wraplength=width).pack(padx=5)