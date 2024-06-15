from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time

from tkinter import *
from tkinter import filedialog
import xlrd

wb = xlrd.open_workbook('kontak.xls')
sheet= wb.sheet_by_index(0)
nama= sheet.cell_value(1,0)
jumlahRow = sheet.nrows
jumlahCol = sheet.ncols


driver = webdriver.Chrome()
driver.maximize_window()
driver.get('https://web.whatsapp.com')

filepath= ''

#names,nama=['me'],0
#pesan = input("Masukan Pesan:")

window = Tk()
window.title('chose picture')
window.columnconfigure(0, weight=1)
window.geometry('600x300')

def openfile():
    global filepath
    filepath = filedialog.askopenfilename()
    if len(filepath) > 1:
        label_notfound.config(text=filepath,fg='green')
    else:
        label_notfound.config(text="not selected",fg='red')
def sendWA():
    data = []
    for a in range(1,jumlahRow):
        data.append(sheet.cell_value(a,0))
    image = filepath
    count = len(data)
    for a in range(0, count):
        user = driver.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p')
        user.click()
        user.clear()
        user.send_keys(data[a])
        user.send_keys(Keys.ENTER)
        time.sleep(5)

        attachbox = driver.find_element(By .XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div')
        attachbox.click()
        time.sleep(5)

        photo = driver.find_element(By .XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/input')
        photo.send_keys(image)
        time.sleep(5)

        send = driver.find_element(By .XPATH,'//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div')
        send.click()
        time.sleep(5)
    window.destroy()
choose = Label(window,text='Choose File', fg='blue', font=('jost', 15, 'bold'))
choose.grid()

button_save = Button(window, width=20, text='search pictures', bg='lightblue', command=openfile)
button_save.grid()

label_notfound = Label(window,text='file not selected', fg='black', font=('jost', 10))
label_notfound.grid()

button_send = Button(window, width=10, text='send', bg='lightgreen', command=sendWA)
button_send.grid()


window.mainloop()