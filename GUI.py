import time
import pyautogui
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_driver_path = './chromedriver.exe'

driver = None

def open_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    entry_file.delete(0, tk.END)
    entry_file.insert(tk.END, filepath)

def open_directory():
    directory = filedialog.askdirectory()
    entry_directory.delete(0, tk.END)
    entry_directory.insert(tk.END, directory)

def send_message():
    filepath = entry_file.get()
    message = entry_message.get("1.0", tk.END)

    df = pd.read_excel(filepath, sheet_name='Sheet1')
    names = set(df['Names'])
    
    print(names)
    
    for name in names:
        find_person(name)
        msg(message)
        time.sleep(1)

    driver.quit()

def send_message_and_images():
    filepath = entry_file.get()
    message = entry_message.get("1.0", tk.END)

    df = pd.read_excel(filepath, sheet_name='Sheet1')
    names = set(df['Names'])
    
    print(names)

    for name in names:
        find_person(name)
        msg(message)
        send_image()
        time.sleep(1)

def open_whatsapp():
    global driver
    driver = webdriver.Chrome(service=Service(executable_path=chrome_driver_path), options=chrome_options)
    driver.get('https://web.whatsapp.com/')


def find_person(name):
    search_box = driver.find_element(By.XPATH, "(//p[@class='selectable-text copyable-text iq0m558w'])[1]")
    search_box.click()
    search_box.send_keys(name)

    time.sleep(2.5)
    search_box.send_keys('\ue004\ue004')
    search_box.send_keys('\ue007')

    time.sleep(2)

def msg(message):
    text_box = driver.find_element(By.XPATH, "//div[@title='Type a message']//p[@class='selectable-text copyable-text iq0m558w']")
    text_box.send_keys(message)
    
    time.sleep(1)

def send_image():
    image_path = entry_directory.get()

    attachment_icon = driver.find_element(By.XPATH, "//span[@data-testid='clip']")
    attachment_icon.click()

    image_button = driver.find_element(By.XPATH, "//span[@data-testid='attach-image']")
    image_button.click()

    time.sleep(1)

    pyautogui.click(421,74)
    time.sleep(0.5)
    pyautogui.write(image_path)
    pyautogui.press('enter')
    pyautogui.click(214,195)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')

    time.sleep(2)

window = tk.Tk()
window.title("WhatsApp Message Sender")
window.iconbitmap("./Whatsapp.ico")
window.geometry("600x650")

label_file = tk.Label(window, text="Excel File Path:")
label_file.pack(pady=10)

entry_file = tk.Entry(window, width=50)
entry_file.pack(pady=10)

button_browse = tk.Button(window, text="Browse", command=open_file, fg='white', bg='black')
button_browse.pack(pady=10)

label_directory = tk.Label(window, text="Image Path:")
label_directory.pack(pady=10)

entry_directory = tk.Entry(window, width=50)
entry_directory.pack(pady=10)

button_browse_directory = tk.Button(window, text="Browse", command=open_directory, fg='white', bg='black')
button_browse_directory.pack(pady=10)

label_message = tk.Label(window, text="Message:")
label_message.pack(pady=10)

entry_message = tk.Text(window, height=5, width=50)
entry_message.pack(pady=10)

button_open_whatsapp = tk.Button(window, text="Open WhatsApp", command=open_whatsapp, fg='white', bg='black')
button_open_whatsapp.pack(pady=10)

button_send_message = tk.Button(window, text="Send Message", command=send_message, fg='white', bg='black')
button_send_message.pack(pady=10)

button_send_message_and_image = tk.Button(window, text="Send Message and Images", command=send_message_and_images, fg='white', bg='black')
button_send_message_and_image.pack(pady=10)

window.mainloop()

