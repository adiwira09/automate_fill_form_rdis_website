import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook

class DataRegistration:
    def __init__(self, file_path):
        self.file_path = file_path
        self.options = webdriver.ChromeOptions()
        self.options.add_experimental_option("detach", True)

    def load_data(self):
        wb = load_workbook(filename=self.file_path)
        sheet_range = wb['Sheet1']
        return sheet_range

    def register_data(self, data):
        self.driver = webdriver.Chrome(options=self.options)
        self.driver.get('https://rdis.idx.co.id/id/register-visit-onsite/390')

        name = data['name']
        identity_number = '-'
        npwp_number = '-'
        email = data['email']
        phone_number = data['phone_number']
        birth_year = '2004'
        address = 'Depok'
        gender = data['gender']

        self.driver.find_element(By.NAME, 'identity_number').send_keys(identity_number)
        self.driver.find_element(By.NAME, 'name').send_keys(name)
        self.driver.find_element(By.NAME, 'npwp_number').send_keys(npwp_number)
        self.driver.find_element(By.NAME, 'email').send_keys(email)
        self.driver.find_element(By.NAME, 'phone_number').send_keys(phone_number)
        select = Select(self.driver.find_element(By.NAME, 'birth_year'))
        select.select_by_visible_text(birth_year)
        self.driver.find_element(By.NAME, 'address').send_keys(address)
        if gender == 'Pria':
            self.driver.find_element(By.ID, 'rbOption1').click()
        elif gender == 'Wanita':
            self.driver.find_element(By.ID, 'rbOption2').click()

class DataRegistrationGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Registrasi Data")
        self.master.geometry("300x210")

        self.label1 = tk.Label(master, text="Nama File:")
        self.label1.pack()

        self.file_entry = tk.Entry(master)
        self.file_entry.pack()

        self.browse_button = tk.Button(master, text="Pilih File", command=self.browse_file)
        self.browse_button.pack()

        self.row_label = tk.Label(master, text="Baris ke:")
        self.row_label.pack()

        self.row_entry = tk.Entry(master)
        self.row_entry.pack()

        self.data_label = tk.Label(master, text="Banyak Data:")
        self.data_label.pack()

        self.data_entry = tk.Entry(master)
        self.data_entry.pack()

        self.submit_button = tk.Button(master, text="Mulai Registrasi", command=self.start_registration)
        self.submit_button.pack()

        self.add_creator_label()

    def add_creator_label(self):
        creator_label = tk.Label(self.master, text="by Nugroho Adi")
        creator_label.pack(anchor='nw', side='bottom')

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, file_path)

    def start_registration(self):
        file_path = self.file_entry.get()
        row = int(self.row_entry.get())
        banyak_data = int(self.data_entry.get())

        registration = DataRegistration(file_path)
        sheet_range = registration.load_data()

        for i in range(row+1, banyak_data+row+1):
            data = {
                'name': sheet_range['E'+str(i)].value,
                'identity_number' : '-',
                'npwp_number' : '-',
                'email': sheet_range['H'+str(i)].value,
                'phone_number': f"0{sheet_range['I'+str(i)].value}",
                'birth_year' :'2003',
                'address' : 'Depok',
                'gender': 'Pria' if i % 2 == 0 else 'Wanita'
            }
            registration.register_data(data)

def main():
    root = tk.Tk()
    app = DataRegistrationGUI(root)
    root.mainloop()

if __name__ == '__main__':
    main()
