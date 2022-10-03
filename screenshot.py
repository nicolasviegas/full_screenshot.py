from selenium import webdriver
from selenium.webdriver.common.by import By
#from Screenshot import Screenshot  
#from selenium.webdriver.chrome.options import Options
#from Screenshot import Screenshot_Clipping
import time
import os

import cv2
import numpy as np

from PIL import Image


###################################################

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument("--start-maximized")
#chromedriver = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'chromedriver.exe')
chrome = webdriver.Chrome(executable_path=r"xxxxx", options=chrome_options)


url = 'xxxx'

chrome.maximize_window()


chrome.get(url)

time.sleep(2)

username = "xxx"
password = "xxx"

chrome.find_element(By.ID, "login-form-username").send_keys(username)

chrome.find_element(By.ID, "login-form-password").send_keys(password)

chrome.find_element(By.ID, "login-form-submit").click()


time.sleep(5)
#########################

chrome.find_element(By.ID, "all-tabpanel").click()

time.sleep(1)

chrome.find_element(By.ID, "key-val").click()


#CLICKEA EN TODO, PERO HAY QUE HACER QUE EL CLICK VUELVA AL PRINCIPIO DE TODO

total_height = chrome.execute_script("return document.body.parentNode.scrollHeight") + 50000


chrome.set_window_size(1920, total_height)


time.sleep(5)

path = "xxxx"



chrome.save_screenshot(path)


########################## ENCONTRAR COORDENADAS DE DONDE TERMINA LA IMAGEN/// AMARILLO: RGB: (255, 211,  81)

im = cv2.imread(path)
yellow = [81, 211, 255]

Y, X = np.where(np.all(im==yellow,axis=2))

print(X,Y)

array_length = len(Y)
last_element = Y[array_length - 1]

#####

im2 =  Image.open(path)
#EL PRIMERO Y EL SEGUNDO TIENEN QUE SER 0, el tercero 1920,  el cuarto es la diferencia de el ultimo amarillo mas 2000 ejemplo
im_crop = im2.crop((0, 0, 1920, last_element + 1000))
#im_crop.save(path, quality=95)
im_crop.save(path, quality=95)




chrome.quit()