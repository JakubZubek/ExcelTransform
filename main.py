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
        frame = customtkinter.CTkFrame(master=root)
        frame.pack(pady=20, padx=60, fill="both", expand=True)
        label = customtkinter.CTkLabel(master=frame, text="Wybierz raport",)
        label.pack(pady=12, padx=10)

        button1 = customtkinter.CTkButton(
            master=frame, text="Raport GPS", command=self.get_input_path)
        button1.pack(pady=12, padx=10)

        button2 = customtkinter.CTkButton(
            master=frame, text="Folder docelowy", command=self.get_output_path)
        button2.pack(pady=12, padx=10)

        button3 = customtkinter.CTkButton(
            master=frame, text="Uruchom", command=self.transform_xlsx)
        button3.pack(pady=12, padx=10)

        button4 = customtkinter.CTkButton(
            master=frame, text="Wyjście", command=self.quit)
        button4.pack(pady=12, padx=10)

    def get_input_path(self):
        self.input_file_path = filedialog.askopenfilename()

    def get_output_path(self):
        self.output_file_path = filedialog.askdirectory()

    def transform_xlsx(self):
        excel_edit.data_clean(self.input_file_path, self.output_file_path)

    def quit(self):
        exit()


# Main screen
root = customtkinter.CTk()
root.geometry("400x300")
Gui(root).pack()
root.mainloop()
