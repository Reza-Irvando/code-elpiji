import sys
sys.path.append('/path/to/openpyxl')
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import time
from openpyxl  import Workbook, load_workbook

delay1 = 1
delay2 = 2
delay3 = 5
delay4 = 7

# List Akun
allakun = [
    ['email', 'pw'],
    ]

# List Sheet
allsheet = [''] ##nama sheet

wb = load_workbook(filename='fileexcel') #file excel 

# Display Menu 
print('Pilih akun berikut: ')
print('1. Siswati')
print('2. Edi')
akunbenar = False
while akunbenar == False:
    nomorakun = input('Masukkan angka dari akun :')
    if nomorakun == '1':
        selectedakun = allakun[0]
    elif nomorakun == '2':
        selectedakun = allakun[1]
    email = selectedakun[0]
    password = selectedakun[1]
    print('Email : ', email)
    print('Password : ', password)
    if nomorakun == '1' or nomorakun == '2':
        print('Pilih sheet berikut: ')
        print('1. Siswati - Selasa')
        print('2. Siswati - Kamis')
        print('3. Siswati - Sabtu')
        print('4. Edi - Jumat')
        print('5. NIK BARU')
        print('6. Supply Fadli')
        print('0. Cadangan Fix')
        sheetbenar = False
        while sheetbenar == False:
            nomorsheet = input('Masukkan angka dari sheet : ')
            if nomorsheet == '1':
                namasheet = allsheet[0]
                sheetbenar = True
            elif nomorsheet == '2':
                namasheet = allsheet[1]
                sheetbenar = True
            elif nomorsheet == '3':
                namasheet = allsheet[2]
                sheetbenar = True
            elif nomorsheet == '4':
                namasheet = allsheet[3]
                sheetbenar = True
            elif nomorsheet == '5':
                namasheet = allsheet[4]
                sheetbenar = True
            elif nomorsheet == '6':
                namasheet = allsheet[5]
                sheetbenar = True
            elif nomorsheet == '0':
                namasheet = allsheet[6]
                sheetbenar = True
            else :
                print('Pilihan tidak ada')
        akunbenar = True
    else:
        print('Akun tidak ada')
        
sheetRange = wb[namasheet]
print('Sheet : ', namasheet)

minggubenar = False
while minggubenar == False:
    inputminggu = input('Minggu ke (1 - 5) : ')
    if inputminggu == '1':
        minggu = sheetRange['G']
        minggubenar = True
    elif inputminggu == '2':
        minggu = sheetRange['H']
        minggubenar = True
    elif inputminggu == '3':
        minggu = sheetRange['I']
        minggubenar = True
    elif inputminggu == '4':
        minggu = sheetRange['J']
        minggubenar = True
    elif inputminggu == '5':
        minggu = sheetRange['K']
        minggubenar = True
    else :
        minggubenar = False
inputawal = input('Nomor mulai : ')

allnomor = sheetRange['A']
allnik = sheetRange['C']
allbeli = sheetRange['D']
allrole = sheetRange['E']

service = Service(executable_path="chromedriver.exe")

driver = webdriver.Chrome(service=service)

driver.get('https://subsiditepatlpg.mypertamina.id/merchant/auth/login')

WebDriverWait(driver, 60).until( 
                EC.visibility_of_any_elements_located((By.XPATH, '/html/body/div[5]/div/div[1]/iframe')))[0]

time.sleep(delay2)
# Proses Login Akun
inputemail = driver.find_element(By.XPATH, '//*[@id="mantine-r0"]')
inputemail.send_keys(email)
inputpassword = driver.find_element(By.XPATH, '//*[@id="mantine-r1"]')
inputpassword.send_keys(password)
driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div[1]/form/div[4]/button').click()
time.sleep(delay4)
driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div[3]/div[2]/div/div[1]/div/div/span').click()
time.sleep(delay2)

awal = int(inputawal)
i = awal
jumlah = 0
j = 1

# Terjual sebelumnya
while j < i:
    if minggu[j].value == 1:
        jumlah = jumlah + allbeli[j].value
    j = j + 1
print('Terjual : ', jumlah)

# Perulangan input data gas
# while i <= len(allnik)-2:
while i <= len(allnik)-2:
    nomor = allnomor[i].value
    nik = allnik[i].value
    roleError = False
    ready = False
    success = False
    if minggu[i].value == 1:
        beli = allbeli[i].value
        role = allrole[i].value
        time.sleep(delay2)
        # Input NIK
        driver.find_element(By.CLASS_NAME, 'mantine-3trqqh').send_keys(nik)
        driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div[2]/div/div[1]/form/div[2]/button').click()
        time.sleep(delay2)

        # Role Conditioning
        match role:
            case 'RT':
                ready = True
            case 'ERT':
                try:
                    driver.find_element(By.CSS_SELECTOR, "label:nth-child(8)").click()
                    time.sleep(delay1)
                    driver.find_element(By.CSS_SELECTOR, 'div:nth-child(11) > button').click()
                    ready = True
                    time.sleep(delay2)
                except:
                    print('nomor: ', nik, ', nik: ', 'Tidak klik role')
            case 'URT':
                try:
                    driver.find_element(By.CSS_SELECTOR, "label:nth-child(8)").click()
                    time.sleep(delay1)
                    driver.find_element(By.CSS_SELECTOR, 'div:nth-child(11) > button').click()
                    ready = True
                    time.sleep(delay2)
                except:
                    print('nomor: ', nik, ', nik: ', 'Tidak klik role')

            case 'EURT':
                try:
                    driver.find_element(By.CSS_SELECTOR, "label:nth-child(8)").click()
                    time.sleep(delay1)
                    driver.find_element(By.CSS_SELECTOR, 'div:nth-child(12) > button').click()
                    ready = True
                    time.sleep(delay2)
                except:
                    ready = False
            case 'U':
                try: 
                    try:
                        driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div/div/form/div[2]/div[2]/button[2]').click()
                        ready = True
                        time.sleep(delay1)
                    except:
                        ready = False
                except:
                    ready = False
            case _:
                ready = False
                roleError = True
                
        if ready == True :
            try:
                driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div/div/form/div[4]/div/button').click()
                time.sleep(delay2)
                driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div/div/div[2]/div/div[4]/div[1]/button').click()
                success = True
                jumlah += beli
                print('nomor: ', nomor, ' , beli: ', beli, ' , Succes. Total: ', jumlah)
                time.sleep(delay2)
            except:
                try:
                    driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div/div/div[2]/div[1]/div[1]/div[2]/button').click()
                    print('nomor: ', nomor, ' , nik: ', nik, ' , Error : NIK Over')
                except:
                    print('nomor: ', nomor, ' , nik: ', nik, ' , Error: Request Time Out')
                    driver.get('https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik')
        else:
            if roleError == True:
                print('nomor: ', nomor, ', nik: ', nik, ' , Error: Role tidak sesuai')
                driver.get('https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik')
            else:
                try:
                    driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div/div/div[2]/div[1]/div[1]/div[2]/button').click()
                    print('nomor: ', nomor, ' , nik: ', nik, ' , Error : NIK Over')
                except:
                    print('nomor: ', nomor, ' , nik: ', nik, ' , Error: Request Time Out')
                    driver.get('https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik')

        if (success == True):
            try:
                driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/div/main/div/div/div/div/div[2]/div[3]/span/a').click()
            except:
                print('Request Time Out. Reconnecting...')
                driver.get('https://subsiditepatlpg.mypertamina.id/merchant/app/verification-nik')
    else:
        print('nomor: ', nomor, 'nik: ', nik, ' . skip')
    i = i + 1
print('done')
time.sleep(delay2)
driver.find_element(By.XPATH, '//*[@id="__next"]/div[1]/footer/div/div[2]').click()
time.sleep(1000)
driver.quit()