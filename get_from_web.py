from selenium import webdriver
from time import sleep

# Using Chrome to access web
driver = webdriver.Chrome()
driver.get(r'https://www.neryisrael.co.uk/minyanbookingsignup.html')
q = driver.find_elements_by_tag_name('input')
email_box = driver.find_elements_by_name('email')[1]
email_box.send_keys('moshesteiny@gmail.com')
password_box = driver.find_elements_by_name('password')[1]
password_box.send_keys('___')
password_box.submit()


driver.get("https://www.neryisrael.co.uk/admin/forms.php?action=viewSubmissions&id=61322")

driver.find_element_by_id('export').click()
sleep(5)
driver.quit()