from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook
# from openpyxl.utils.exceptions import InvalidFileException

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

def main():
    # nama_file = input('Nama File: ')
    nama_file = '6 juni 2023.xlsx'
    file_path = f"C:/Users/HP/Desktop/Program/Selenium/{nama_file}"
    row = int(input('Baris ke: '))
    banyak_data = int(input('Banyak Data: '))

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

if __name__ == '__main__':
    main()
# 3 row done