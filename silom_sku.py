import os
import sys
import time
from datetime import datetime, timedelta
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

file_path = os.path.join(DOWNLOAD_DIR, "SKU.xlsx")

# 🌟 เช็กสถานะการสั่งงานว่าเป็นการ Manual Upload หรือไม่
IS_MANUAL = os.environ.get('IS_MANUAL_UPLOAD', '').lower() == 'true'
MANUAL_FILE_NAME = os.environ.get('MANUAL_FILE_NAME', '')

print(f"🔍 ตรวจสอบโหมดการทำงาน SKU: IS_MANUAL_UPLOAD = {IS_MANUAL}")

# ==========================================================
# ส่วนที่ 1: การเตรียมไฟล์ข้อมูล (อัตโนมัติ vs Manual)
# ==========================================================
if IS_MANUAL:
    print("\n📁 --- โหมด Manual Upload: ข้ามกระบวนการ Selenium/Silom POS ---")
    
    # ค้นหาไฟล์ SKU ที่ผู้ใช้อัปโหลดเข้ามาผ่านหน้าเว็บ
    candidate_files = [
        MANUAL_FILE_NAME,
        "manual_sku_data.csv",
        "manual_sku_data.xlsx",
        "manual_sku_data.xls"
    ]
    
    found_file = None
    for cf in candidate_files:
        if cf and os.path.exists(cf):
            found_file = cf
            break
            
    if not found_file and os.path.exists(file_path):
        found_file = file_path

    if not found_file:
        raise FileNotFoundError(f"🔴 ไม่พบไฟล์ SKU ที่อัปโหลดแมนนวลเข้ามาในระบบ! (ไฟล์ที่ค้นหา: {candidate_files})")

    print(f"✅ พบไฟล์ SKU สำหรับนำเข้า BigQuery: {found_file}")
    target_data_file = found_file

else:
    print("\n🤖 --- โหมดดึงข้อมูลอัตโนมัติ: กำลังเปิด Chrome สแครปข้อมูลจาก Silom POS ---")
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

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
    chrome_options.add_argument("--window-size=1440,900")

    print("กำลังสั่งเปิด Chrome (Headless) บน GitHub Actions...")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

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
        
        print("🧼 เริ่มกระบวนการเคลียร์หน้าจอและสิ่งกีดขวาง...")
        try:
            driver.execute_script("""
                var modals = document.querySelectorAll('.v-modal, .el-dialog__wrapper, .modal-backdrop, [role="dialog"], .el-drawer__wrapper, .el-drawer');
                modals.forEach(function(el) { el.remove(); });
                document.body.style.overflow = 'auto';
                var chats = document.querySelectorAll('#crisp-chat-box, .crisp-client, [class^="cc-"], [class^="crisp-"]');
                chats.forEach(function(el) { el.remove(); });
            """)
            print("-> ล้างสิ่งกีดขวางหน้าจอเรียบร้อย")
            time.sleep(1)
        except Exception:
            pass

        print("⚡ [JS] กำลังคลิกหัวข้อหลัก 'สินค้าคงคลัง'...")
        menu_inventory = wait.until(EC.presence_of_element_located((
            By.XPATH, "//*[contains(@class, 'sidebar') or contains(@class, 'menu')]//*[contains(text(), 'สินค้าคงคลัง')]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu_inventory)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", menu_inventory)
        time.sleep(3)
        
        print("⚡ [JS] กำลังคลิกเมนูย่อย 'สินค้าคงเหลือตาม SKU'...")
        submenu_sku = wait.until(EC.presence_of_element_located((
            By.XPATH, "//*[contains(@class, 'sidebar') or contains(@class, 'menu')]//*[contains(text(), 'สินค้าคงเหลือตาม SKU')]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", submenu_sku)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", submenu_sku)
        
        print("🎯 ถึงหน้าตารางรายงาน SKU แล้ว รอระบบดึงข้อมูล...")
        time.sleep(8)
        
        print("กำลังดักรอปุ่ม 'ส่งออกไฟล์' ปรากฏ...")
        export_button = wait.until(EC.presence_of_element_located((By.ID, "SKUInventoryExportButton")))
        time.sleep(2)
        
        try:
            driver.save_screenshot(os.path.join(SCREENSHOT_DIR, "1_before_click.png"))
        except Exception:
            pass

        print("กำลังใช้ JavaScript สั่งกดส่งออกไฟล์ Excel...")
        driver.execute_script("arguments[0].click();", export_button)
        time.sleep(3)

        try:
            alert = driver.switch_to.alert
            alert_text = alert.text
            print(f"⚠️ เจอแจ้งเตือนจากหน้าเว็บ SKU: {alert_text}")
            alert.accept()
        except Exception:
            pass
        
        print("⏱️ วนลูปรอตรวจสอบไฟล์รายงาน SKU ตกลงมาในโฟลเดอร์...")
        downloaded = False
        for i in range(12):
            time.sleep(5)
            files = os.listdir(DOWNLOAD_DIR)
            valid_files = [f for f in files if not f.endswith('.crdownload') and f != '']
            if valid_files:
                print(f"พบไฟล์ดาวน์โหลดในรอบที่ {i+1}: {valid_files}")
                downloaded = True
                break
            print(f"รอบที่ {i+1}: ยังไม่พบไฟล์ SKU กำลังรอต่อ...")
            
        try:
            driver.save_screenshot(os.path.join(SCREENSHOT_DIR, "2_after_click.png"))
        except Exception:
            pass
        
        files = os.listdir(DOWNLOAD_DIR)
        if downloaded and files:
            latest_file = max([os.path.join(DOWNLOAD_DIR, f) for f in files], key=os.path.getctime)
            if latest_file != file_path:
                os.rename(latest_file, file_path)
            print(f"เตรียมอัปโหลดไฟล์เสร็จสมบูรณ์ที่: {file_path}")
            target_data_file = file_path
        else:
            raise FileNotFoundError("บอทหาไฟล์ Excel ที่ดาวน์โหลดไม่เจอในโฟลเดอร์!")

    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการทำงาน: {str(e)}")
        try:
            driver.save_screenshot(os.path.join(SCREENSHOT_DIR, "error_screenshot.png"))
        except Exception:
            pass
        raise e
    finally:
        driver.quit()


# ==========================================================
# ส่วนที่ 2: โค้ดส่งไฟล์จริงเข้า BigQuery (เวอร์ชันเคลียร์วันซ้ำ)
# ==========================================================
print("\n--- เริ่มกระบวนการส่งข้อมูลเข้า Google Cloud BigQuery ด้วย Pandas ---")

print(f"กำลังเปิดอ่านข้อมูลภายในไฟล์: {target_data_file}")

# รองรับทั้งไฟล์ CSV และ Excel (.xlsx/.xls)
if target_data_file.endswith('.csv'):
    try:
        df = pd.read_csv(target_data_file, header=7)
    except Exception:
        df = pd.read_csv(target_data_file)
else:
    try:
        df = pd.read_excel(target_data_file, header=7)
    except Exception:
        df = pd.read_excel(target_data_file)

df.columns = (
    df.columns.astype(str)
    .str.replace(' ', '_')
    .str.replace('.', '_', regex=False)
    .str.replace('/', '_')
    .str.replace('(', '')
    .str.replace(')', '')
)

# คำนวณเวลาไทยปัจจุบันสำหรับระบุวันเก็บข้อมูล
th_time = datetime.utcnow() + timedelta(hours=7)
df['Run_Date'] = th_time.strftime('%Y-%m-%d %H:%M:%S')

table_id = "stock_data.sku_list"
project_id = "northern-eon-470602-a2"
full_table_path = f"{project_id}.{table_id}"

# สั่งลบข้อมูลเก่าของ "วันนี้" ออกก่อนเพื่อป้องกันการบันทึกซ้ำ
try:
    client = bigquery.Client(project=project_id)
    today_str = th_time.strftime('%Y-%m-%d')
    
    delete_query = f"""
        DELETE FROM `{full_table_path}`
        WHERE DATE(Run_Date) = '{today_str}';
    """
    print(f"🧹 กำลังเคลียร์ข้อมูลเก่าของวันที่ {today_str} ใน BigQuery (sku_list) เพื่อป้องกันการบันทึกซ้ำ...")
    query_job = client.query(delete_query)
    query_job.result()
    print("-> ลบข้อมูล SKU รอบเดิมของวันนี้เรียบร้อยแล้ว!")
except Exception as err:
    print(f"⚠️ ไม่สามารถลบข้อมูลเก่าได้: {str(err)}")

# ส่งข้อมูลชุดล่าสุดเข้าไปเติมลงใน BigQuery
print(f"กำลังส่งข้อมูลจำนวน {len(df)} แถว เข้าสู่ BigQuery ตาราง {full_table_path}...")
df.to_gbq(
    destination_table=table_id,
    project_id=project_id,
    if_exists='append',
    progress_bar=False
)

print(f"🎉 🎉 🎉 อัปโหลดสำเร็จ 100%! ข้อมูล SKU ถูกรีเฟรชเป็นรุ่นล่าสุดเรียบร้อยครับ")


# ==========================================================
# ส่วนที่ 3: ดึงเวลาอัปเดตล่าสุดจาก BigQuery และเซฟลงไฟล์ข้อความ
# ==========================================================
print("\n--- เริ่มกระบวนการดึงเวลาแก้ไขล่าสุดจาก BigQuery เพื่อส่งให้หน้าเว็บ ---")
try:
    client = bigquery.Client(project=project_id)
    query = f"""
        SELECT TIMESTAMP_MILLIS(last_modified_time) AS last_updated
        FROM `{project_id}.stock_data.__TABLES__`
        WHERE table_id = 'sku_list';
    """
    query_job = client.query(query)
    results = query_job.result()
    
    for row in results:
        utc_time = row.last_updated
        thai_time = utc_time + timedelta(hours=7)
        time_str = thai_time.strftime('%Y-%m-%d %H:%M:%S')
        
        with open("sku_last_update.txt", "w", encoding="utf-8") as f:
            f.write(time_str)
        print(f"🎯 ดึงข้อมูลสำเร็จ! เวลาแก้ไขจริงใน BigQuery คือ: {time_str} น. (บันทึกลงไฟล์แล้ว)")

except Exception as e:
    # หากดึงจาก BigQuery ไม่สำเร็จ ให้เขียนเวลาปัจจุบันลงไฟล์สำรอง เพื่อให้จุดเขียวแสดงผล
    time_str = th_time.strftime('%Y-%m-%d %H:%M:%S')
    with open("sku_last_update.txt", "w", encoding="utf-8") as f:
        f.write(time_str)
    print(f"⚠️ เขียนเวลาปัจจุบันลงไฟล์สำรอง: {time_str} น.")
