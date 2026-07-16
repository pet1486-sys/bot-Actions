import os
import sys
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from google.cloud import bigquery

# สั่งให้ Python พิมพ์ข้อความเรียงตามบรรทัดจริงบน GitHub Actions ป้องกัน Logs สลับกัน
sys.stdout.reconfigure(line_buffering=True)

# ==========================================================
# CONFIGURATION & SETTINGS
# ==========================================================
USERNAME = "pet1486@gmail.com"
PASSWORD = "htz32151"

DOWNLOAD_DIR = "stock_data" 
SCREENSHOT_DIR = "/home/runner/work_screenshots"

for folder in [DOWNLOAD_DIR, SCREENSHOT_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# เปลี่ยนชื่อเป็น Sales.xlsx
file_path = os.path.join(DOWNLOAD_DIR, "Sales.xlsx")

# ตั้งค่า Chrome Options สำหรับทำงานบน GitHub Actions
chrome_options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": os.path.abspath(DOWNLOAD_DIR), 
    "download.prompt_for_download": False,        
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.password_manager_leak_detection": False,
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False
}
chrome_options.add_experimental_option("prefs", prefs)

chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

print("กำลังสั่งเปิด Chrome (Headless) บน GitHub Actions...")
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# ==========================================================
# ส่วนที่ 1: ดาวน์โหลดไฟล์ด้วย Selenium
# ==========================================================
try:
    print("กำลังเปิดหน้าเว็บไซต์ Silom POS...")
    driver.get("https://dashboard.silompos.com/login")
    
    print("กำลังกรอกข้อมูลเข้าสู่ระบบ...")
    username_input = wait.until(EC.presence_of_element_located((
        By.XPATH, "//input[@type='text' or @type='email' or @autocomplete='username']"
    )))
    username_input.clear()
    username_input.send_keys(USERNAME)
    
    password_input = driver.find_element(By.XPATH, "//input[@type='password']")
    password_input.clear()
    password_input.send_keys(PASSWORD)
    
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Sign In')]")))
    login_button.click()
    
    print("กำลังรอโหลดหน้า Dashboard...")
    time.sleep(8)
    
    try:
        driver.execute_script("""
            var modals = document.querySelectorAll('.v-modal, .el-dialog__wrapper, .modal-backdrop, [role="dialog"]');
            modals.forEach(function(el) { el.remove(); });
            document.body.style.overflow = 'auto';
            var chats = document.querySelectorAll('#crisp-chat-box, .crisp-client, [class^="cc-"]');
            chats.forEach(function(el) { el.remove(); });
        """)
        print("ล้างสิ่งกีดขวางหน้าจอเรียบร้อย")
    except Exception:
        pass

    # คลิกไปที่เมนูการขาย > ยอดขายตามรายละเอียดบิล
    print("กำลังคลิกหัวข้อหลัก 'การขาย'...")
    menu_sales = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//*[contains(@class, 'sidebar') or contains(@class, 'menu')]//*[contains(text(), 'การขาย')]"
    )))
    menu_sales.click()
    time.sleep(2)
    
    print("กำลังคลิกเมนูย่อย 'ยอดขายตามรายละเอียดบิล'...")
    submenu_sales_detail = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//*[contains(@class, 'sidebar') or contains(@class, 'menu')]//*[contains(text(), 'ยอดขายตามรายละเอียดบิล')]"
    )))
    submenu_sales_detail.click()
    
    # ส่งออกไฟล์
    print("กำลังดักรอปุ่ม 'ส่งออกไฟล์' ปรากฏ...")
    export_button = wait.until(EC.presence_of_element_located((
        By.XPATH, "//*[contains(text(), 'ส่งออกไฟล์') or contains(@id, 'Export')]"
    )))
    time.sleep(5)
    
    driver.save_screenshot(os.path.join(SCREENSHOT_DIR, "1_sales_before_click.png"))
    print("📸 บันทึกภาพหน้าจอก่อนกดปุ่มส่งออกไฟล์เรียบร้อย")

    print("กำลังใช้ JavaScript สั่งกดส่งออกไฟล์ Excel...")
    driver.execute_script("arguments[0].click();", export_button)
    
    print("⏱️ รอระบบบันทึกไฟล์ลงดิสก์บนเซิร์ฟเวอร์ 20 วินาทีเพื่อความชัวร์...")
    time.sleep(20)
    
    driver.save_screenshot(os.path.join(SCREENSHOT_DIR, "2_sales_after_click.png"))
    print("📸 บันทึกภาพหน้าจอหลังกดปุ่มส่งออกไฟล์เรียบร้อย")
    
    # ตรวจสอบไฟล์ที่ดาวน์โหลดมา
    files = os.listdir(DOWNLOAD_DIR)
    print(f"ไฟล์ที่พบในโฟลเดอร์ดาวน์โหลด: {files}")
    
    # เปลี่ยนชื่อไฟล์ล่าสุดให้กลายเป็น Sales.xlsx
    if files:
        latest_file = max([os.path.join(DOWNLOAD_DIR, f) for f in files], key=os.path.getctime)
        if latest_file != file_path:
            os.rename(latest_file, file_path)
        print(f"เตรียมอัปโหลดไฟล์เสร็จสมบูรณ์ที่: {file_path}")
    else:
        raise FileNotFoundError("บอทหาไฟล์ Excel ที่ดาวน์โหลดไม่เจอในโฟลเดอร์!")

except Exception as e:
    print(f"เกิดข้อผิดพลาดในการทำงาน: {str(e)}")
    try:
        driver.save_screenshot(os.path.join(SCREENSHOT_DIR, "sales_error_screenshot.png"))
        print("📸 บันทึกภาพหน้าจอขณะเกิดข้อผิดพลาดเรียบร้อย")
    except Exception:
        pass
    raise e
finally:
    driver.quit()


# ==========================================================
# ส่วนที่ 2: โค้ดส่งไฟล์จริงเข้า BigQuery ด้วย Pandas
# ==========================================================
print("\n--- เริ่มกระบวนการส่งข้อมูลเข้า Google Cloud BigQuery ด้วย Pandas ---")

# ใช้ Pandas เปิดอ่านไฟล์ Sales.xlsx ข้ามหัว 7 แถวแรก
print(f"กำลังเปิดอ่านข้อมูลภายในไฟล์ Excel: {file_path}")
df = pd.read_excel(file_path, header=7)

# 🌟 [เพิ่มใหม่] สั่งลบแถวข้อความสรุป/หมายเหตุท้ายตารางออกก่อนส่งขึ้น BigQuery
print("กำลังตรวจสอบและลบแถวข้อความหมายเหตุท้ายตาราง...")
df = df[~df.iloc[:, 0].astype(str).str.contains('\*\*\*\*|รวมสุทธิ', na=False)]
df = df.dropna(how='all')

# จัดการลบอักขระพิเศษออกจากชื่อคอลัมน์
df.columns = (
    df.columns.astype(str)
    .str.replace(' ', '_')
    .str.replace('.', '_', regex=False)
    .str.replace('/', '_')
    .str.replace('(', '')
    .str.replace(')', '')
)

# เพิ่มคอลัมน์ "วันเวลาที่รันบอท" (Run_Date) อ้างอิงเวลาประเทศไทย GMT+7
th_time = datetime.utcnow() + timedelta(hours=7)
df['Run_Date'] = th_time.strftime('%Y-%m-%d %H:%M:%S')

# เปลี่ยนเทเบิ้ลที่จะเอาไปต่อเป็น sales_list
table_id = "stock_data.sales_list"
project_id = "northern-eon-470602-a2"
full_table_path = f"{project_id}.{table_id}"

print(f"กำลังส่งข้อมูลจำนวน {len(df)} แถว เข้าสู่ BigQuery ตาราง {full_table_path}...")

# สั่งอัปโหลดข้อมูลแบบต่อท้าย (append)
df.to_gbq(
    destination_table=table_id,
    project_id=project_id,
    if_exists='append',
    progress_bar=False
)

print(f"🎉 🎉 🎉 อัปโหลดสำเร็จ 100%! ข้อมูลถูกเพิ่มเข้าตาราง {full_table_path} เรียบร้อยแล้วครับ")


# ==========================================================
# ส่วนที่ 3: ดึงเวลาอัปเดตล่าสุดจาก BigQuery และเซฟลงไฟล์ข้อความ
# ==========================================================
print("\n--- เริ่มกระบวนการดึงเวลาแก้ไขล่าสุดจาก BigQuery เพื่อส่งให้หน้าเว็บ ---")
try:
    client = bigquery.Client(project=project_id)
    # คิวรีดึง Metadata ตาราง sales_list
    query = f"""
        SELECT TIMESTAMP_MILLIS(last_modified_time) AS last_updated
        FROM `{project_id}.stock_data.__TABLES__`
        WHERE table_id = 'sales_list';
    """
    query_job = client.query(query)
    results = query_job.result()
    
    for row in results:
        # แปลงเวลาจาก UTC เป็นเวลาไทย (+7 ชั่วโมง)
        utc_time = row.last_updated
        thai_time = utc_time + timedelta(hours=7)
        time_str = thai_time.strftime('%Y-%m-%d %H:%M:%S')
        
        # เพิ่มไฟล์ใหม่ชื่อ sales_last_update.txt
        with open("sales_last_update.txt", "w", encoding="utf-8") as f:
            f.write(time_str)
        print(f"🎯 ดึงข้อมูลสำเร็จ! เวลาแก้ไขจริงใน BigQuery คือ: {time_str} น. (บันทึกลงไฟล์แล้ว)")

except Exception as e:
    print(f"⚠️ เกิดข้อผิดพลาดในการเขียนไฟล์เวลาอัปเดต: {str(e)}")