import tkinter as tk
import customtkinter
from tkinter import filedialog
import pandas as pd
import excel_edit


class Gui(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)
        customtkinter.set_appearance_mode("dark")
        customtkinter.set_default_color_theme("dark-blue")
        self.frame = customtkinter.CTkFrame(master=root)
        self.frame.pack(pady=20, padx=60, fill="both", expand=True)
        self.label = customtkinter.CTkLabel(
            master=self.frame, text="Wybierz raport",)
        self.label.pack(pady=12, padx=10)

        self.button1 = customtkinter.CTkButton(
            master=self.frame, text="Raport GPS", command=self.get_input_path)
        self.button1.pack(pady=12, padx=10)

        self.button2 = customtkinter.CTkButton(
            master=self.frame, text="Folder docelowy", command=self.get_output_path, state="disabled")
        self.button2.pack(pady=12, padx=10)

        self.button3 = customtkinter.CTkButton(
            master=self.frame, text="Uruchom", command=self.transform_xlsx, state="disabled")
        self.button3.pack(pady=12, padx=10)

    def get_input_path(self):
        self.input_file_path = filedialog.askopenfilename(
            filetypes=[("Excel file", ".xlsx")])
        if self.input_file_path:
            self.label.configure(text="Wybierz folder docelowy")
            self.button2.configure(state="normal")

    def get_output_path(self):
        self.output_file_path = filedialog.askdirectory()
        if self.input_file_path and self.output_file_path:
            self.label.configure(text="Kliknij uruchom")
            self.button3.configure(state="normal")

    def transform_xlsx(self):
        try:
            excel_edit.data_clean(self.input_file_path, self.output_file_path)
            self.label.configure(text="Gotowe")
        except KeyError:
            self.label.configure(
                text="Plik nie jest raportem GPS\nlub zmieniono jego strukturę")


# Main screen
root = customtkinter.CTk()
root.geometry("300x300")
root.minsize(300, 300)
root.maxsize(300, 300)
root.title("DriversExcel")
Gui(root).pack()
root.mainloop()
