from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

driver = webdriver.Chrome()
driver.set_page_load_timeout(300)

# Open and login
driver.get("http://wrts.lge.com/wrts_main_page.php")
driver.maximize_window()
time.sleep(5)
driver.find_element(By.ID, "username").send_keys("nithyanantham.s")
driver.find_element(By.ID, "password").send_keys("LGsoft@1234567")
driver.find_element(By.XPATH, "//input[@type='submit']").click()

# Handle popup if present
try:
    close_btn = WebDriverWait(driver, 600).until(
        EC.element_to_be_clickable((By.ID, "id_shut_down_checkbox"))
    )
    close_btn.click()
except:
    pass

region_info = ("EU","89881")
TC= [ "24Y_DEV_A001" , "24Y_DEV_A002" , "24Y_DEV_A003" , "24Y_DEV_A004" , "24Y_DEV_A005" ]

try:
    webID = WebDriverWait(driver, 300).until(
        EC.element_to_be_clickable((By.XPATH, f"//td[@title='89881']"))
    )
    webID.click()
    print("Clicked wID")
except Exception as e:
    print(e)
    # continue
try:
    elements = WebDriverWait(driver, 10).until(
         EC.presence_of_all_elements_located((By.XPATH, "//td[@title='Basic.DeviceControl']"))
    )

    target_td = elements[0]
    # Optional: scroll into view if needed
    driver.execute_script("arguments[0].scrollIntoView(true);", target_td)

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(target_td))

    # Click the <td>
    target_td.click()

except Exception as e:
    print("Failed to click the element:", e)

try:
    webTC = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH,"//li[@rel='tab2']"))
    )
    webTC.click()
    print("Clicked TC Tab")
except Exception as e:
    print(e,"TC Tab not clicked")
    time.sleep(5)
    print("TC Tab Opened")

for id in TC :
    try :
        TCID = WebDriverWait(driver,20).until(
            EC.element_to_be_clickable((By.XPATH, f"//td[@title='{id}']"))
        )
        TCID.click()
    except Exception as e:
        print( e ,"TC not found")
    
    try :
        TCComment = WebDriverWait(driver,20).until(
            EC.visibility_of_element_located((By.XPATH,"//textarea[@id='comment_view']"))
        )
        TCComment.click()
        TCComment.send_keys("Passed")
        try :
            TCsave = WebDriverWait(driver,20).until(
            EC.element_to_be_clickable((By.XPATH, f"//td[@title={id}]"))
            )
            TCsave.click()
        except Exception as e:
            print( e ,"Comment not saved")
    except Exception as e:
        print(e,"Not commented")

    try :
        TCPass = WebDriverWait(driver,20).until(
            EC.visibility_of_element_located((By.XPATH,"//input[@id='wrts_testcase_status_radio1']"))
        )
        TCPass.click()
    except Exception as e:
        print(e,"Not Passed")

time.sleep(5)
driver.quit()