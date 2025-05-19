# # ============ Get User Input FIRST ============
# region_info = []

# num_regions_input = input("Enter number of regions: ")
# try:
#     num_regions = int(num_regions_input)
# except ValueError:
#     print("Invalid number. Exiting.")
#     exit()

# for i in range(num_regions):
#     region = input(f"Enter Region Name #{i+1}: ").strip()
#     wrts_id = input(f"Enter WRTS ID for {region}: ").strip()
#     region_info.append((region, wrts_id))

# print("\nProceeding with the following regions:")
# for region, wrts_id in region_info:
#     print(f"  - {region}: {wrts_id}")


from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl

driver = webdriver.Chrome()

# Open and login
driver.get("http://wrts.lge.com/wrts_main_page.php")
driver.maximize_window()
time.sleep(5)
driver.find_element(By.ID, "username").send_keys("nithyanantham.s")
driver.find_element(By.ID, "password").send_keys("LGsoft@1234567")
driver.find_element(By.XPATH, "//input[@type='submit']").click()

# Handle popup if present
try:
    close_btn = WebDriverWait(driver, 30).until(
        EC.element_to_be_clickable((By.ID, "id_shut_down_checkbox"))
    )
    close_btn.click()
except:
    pass

# Excel setup
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Test Results"
headers = ['S.No', 'Region', 'WRTS ID', 'Total', 'Total NA', 'Pass', 'Fail', 'NA', 'NE', 'NI', 'Blocked']
ws.append(headers)

region_info = [
    ("EU", "89881"),
    ("KR", "89882"),
    ("US", "89885"),
    ("BR", "89886"),
    ("TW", "89887"),
    ("PH", "89888"),
    ("AJ", "89889"),
    ("JP", "89890"),
    ("CN", "89891"),
]

sno = 1
for region, wrts_id in region_info:
    # Click on WRTS ID
    try:
        webID = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, f"//td[@title='{wrts_id}']"))
        )
        webID.click()
        print(f"{region} - {wrts_id} Clicked")
    except Exception as e:
        print(f"{region} - {wrts_id} Not Found", e)
        continue

    # Click on the model status button
    driver.find_element(By.ID, "btn_main_view_model_status").click()

    # Wait and switch to new window
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    windows = driver.window_handles
    driver.switch_to.window(windows[1])

    # Wait for table to load
    try:
        WebDriverWait(driver, 20).until(
            EC.visibility_of_element_located((By.XPATH, "//table[@class='ui-jqgrid-ftable']"))
        )
    except:
        print("Table not found for", wrts_id)
        driver.close()
        driver.switch_to.window(windows[0])
        continue

    time.sleep(2)
    driver.execute_script("document.querySelector('.ui-dialog-titlebar-close').click();")

    rows = driver.find_elements(By.CSS_SELECTOR, "table.ui-jqgrid-ftable tbody tr")

    for row in rows:
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) >= 12:
            values = []
            for i in range(4, 12):
                text = cells[i].text.strip().replace(',', '')
                if text.isdigit():
                    values.append(int(text))
                else:
                    try:
                        values.append(float(text))
                    except ValueError:
                        values.append(text)  # keep as string if not numeric

            data = [sno, region, wrts_id] + values
            ws.append(data)
            sno += 1

    driver.close()
    driver.switch_to.window(windows[0])
    time.sleep(1)

# Save file
wb.save("wrts_all_regions.xlsx")
print("Excel file saved.")

# Initialize totals
total_columns = [0] * 8  # Only the last 8 columns are numeric (Total, Total NA, Pass, Fail, NA, NE, NI, Blocked)

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=4, max_col=11):
    for i, cell in enumerate(row):
        try:
            value = int(str(cell.value).replace(',', ''))
            total_columns[i] += value
        except:
            continue  # Skip non-numeric

# Add the totals to the final row
grand_total_row = ["", "", "Grand Total"] + total_columns
ws.append(grand_total_row)

# Save Excel
wb.save("wrts_all_regions_with_total.xlsx")
print("Excel file with Grand Total saved.")

driver.quit()
