import customtkinter as ctk
from openpyxl import Workbook
import pyperclip
import time


class Application(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Cadastro de dados")
        self.geometry("300x400")

        self.date_label = ctk.CTkLabel(self, text="Data:")
        self.date_label.pack()

        self.date_entry = ctk.CTkEntry(self)
        self.date_entry.pack()

        self.plate_label = ctk.CTkLabel(self, text="Placa:")
        self.plate_label.pack()

        self.plate_entry = ctk.CTkEntry(self)
        self.plate_entry.pack()

        self.state_label = ctk.CTkLabel(self, text="Estado:")
        self.state_label.pack()

        self.state_entry = ctk.CTkEntry(self)
        self.state_entry.pack()

        self.weight_label = ctk.CTkLabel(self, text="Peso:")
        self.weight_label.pack()

        self.weight_entry = ctk.CTkEntry(self)
        self.weight_entry.pack()

        self.save_button = ctk.CTkButton(
            self, text="Salvar", command=self.save_data)
        self.save_button.pack()

        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.append(["Data", "Placa", "Estado", "Peso"])

    def save_data(self):
        date = self.date_entry.get()
        plate = self.plate_entry.get()
        state = self.state_entry.get()
        weight = self.weight_entry.get()

        self.ws.append([date, plate, state, weight])
        self.wb.save("dados.xlsx")

        pyperclip.copy(f"{plate}")
        time.sleep(3)
        pyperclip.copy(f"{state}")
        time.sleep(3)
        pyperclip.copy(f"{weight}")


app = Application()
app.mainloop()
