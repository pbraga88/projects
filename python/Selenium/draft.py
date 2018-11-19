# from selenium import webdriver
# browser = webdriver.Firefox()
# browser.get('http://192.168.0.1')

import time
import unittest
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from PIL import Image
import pytesseract


# class PythonOrgSearch(unittest.TestCase):
#
#     def setUp(self):
#         self.driver = webdriver.Firefox()
#
#     def test_search_in_python_org(self):
#         driver = self.driver
#         driver.get("http://192.168.0.1")
#         # self.assertIn("Python", driver.title)
#         elem = driver.find_element_by_id("Password")
#         elem.send_keys("password")
#         elem.send_keys(Keys.RETURN)
#         assert "No results found." not in driver.page_source
#
#
#     def tearDown(self):
#         self.driver.close()
#
# if __name__ == "__main__":
#     unittest.main()



driver = webdriver.Firefox()



driver = driver
driver.get("http://192.168.0.1") #WebDriver will wait until the page has fully loaded
elem = driver.find_element_by_id("login")
elem.clear()
elem.send_keys("NET_0B4470")
elem = driver.find_element_by_id("senha")
elem.clear()
elem.send_keys("A811FC0B4470")
elem.send_keys(Keys.RETURN)

WebDriverWait(driver, 10).until(lambda driver: driver.find_element_by_class_name("icon-signal"))
time.sleep(10)
driver.find_element_by_class_name("icon-signal").click()
driver.save_screenshot('screenshot.png')

'''OCR'''
img = Image.open("screenshot.png")
crop_Tx = img.crop((985, 48, 1011, 62))
crop_Tx.save("crop_Tx.png")
crop_Rx = img.crop((985, 62, 1011, 75))
crop_Rx.save("crop_Rx.png")

valorTx = pytesseract.image_to_string(Image.open('crop_Tx.png'))
print("TX:", valorTx)

# valorRx = pytesseract.image_to_string(Image.open('crop_Rx.png'))
# print("RX:", valorRx)