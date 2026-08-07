import os
import sys
import time
from datetime import datetime, timedelta
import pandas as pd
from google.cloud import bigquery

# สั่งให้ Python พิมพ์ข้อความเรียงตามบรรทัดจริงบน GitHub Actions
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

file_path = os.path.join(DOWNLOAD_DIR, "Sales.xlsx")

# เช็กสถานะการสั่งงานว่าเป็นการ Manual Upload หรือไม่
IS_MANUAL = os.environ.get('IS_MANUAL_UPLOAD', '').lower() == 'true'
MANUAL_FILE_NAME = os.environ.get('MANUAL_FILE_NAME', '')

print(f"🔍 ตรวจสอบโหมดการทำงาน: IS_MANUAL_UPLOAD = {IS_MANUAL}")

# ==========================================================
# ส่วนที่ 1: การเตรียมไฟล์ข้อมูล (อัตโนมัติ vs Manual)
# ==========================================================
if IS_MANUAL:
    print("\n📁 --- โหมด Manual Upload: ข้ามกระบวนการ Selenium/Silom POS ---")
    
    candidate_files = [
        MANUAL_FILE_NAME,
        "manual_sales_data.csv",
        "manual_sales_data.xlsx",
        "manual_sales_data.xls"
    ]
    
    found_file = None
    for cf in candidate_files:
        if cf and os.path.exists(cf):
            found_file = cf
            break
            
    if not found_file and os.path.exists(file_path):
        found_file = file_path

    if not found_file:
        raise FileNotFoundError(f"🔴 ไม่พบไฟล์ที่อัปโหลดแมนนวลเข้ามาในระบบ! (ไฟล์ที่ค้นหา: {candidate_files})")

    print(f"✅ พบไฟล์สำหรับนำเข้า BigQuery: {found_file}")
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

        print("กำลังค้นหาและคลิกหัวข้อหลัก 'การขาย'...")
        menu_sales = wait.until(EC.presence_of_element_located((
            By.XPATH, "//*[contains(text(), 'การขาย') or contains(@class, 'menu')]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu_sales)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", menu_sales)
        
        print("รอเมนูกางออก 4 วินาที...")
        time.sleep(4)
        
        print("กำลังค้นหาและคลิกเมนูย่อย 'ยอดขายตามรายละเอียดบิล'...")
        submenu_sales_detail = wait.until(EC.presence_of_element_located((
            By.XPATH, "//*[contains(text(), 'ยอดขายตามรายละเอียดบิล')]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", submenu_sales_detail)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", submenu_sales_detail)
        
        print("🎯 กำลังโหลดหน้ารายงานยอดขายตามรายละเอียดบิล...")
        time.sleep(8)
        
        try:
            driver.execute_script("""
                var crispElements = document.querySelectorAll(
                    '#crisp-chat-box, .crisp-client, [class^="crisp-"], [id^="crisp-"], ' +
                    'iframe[title*="chat" i], iframe[src*="crisp"], iframe[id*="crisp"]'
                );
                crispElements.forEach(function(el) { el.remove(); });
            """)
            time.sleep(1)
        except Exception:
            pass

        try:
            driver.execute_script("""
                document.querySelectorAll('.el-drawer__wrapper, .el-drawer, .v-modal').forEach(el => el.remove());
                document.body.style.overflow = 'auto';
            """)
            time.sleep(2)
        except Exception:
            pass

        print("กำลังดักรอปุ่ม 'ส่งออกไฟล์' ปรากฏ...")
        export_button = wait.until(EC.presence_of_element_located((
            By.XPATH, "//*[contains(text(), 'ส่งออกไฟล์')]"
        )))
        
        print("กำลังส่งคำสั่งคลิกปุ่มส่งออกไฟล์...")
        driver.execute_script("arguments[0].click();", export_button)
        
        try:
            time.sleep(3) 
            alert = driver.switch_to.alert
            alert_text = alert.text
            alert.accept() 
            
            if "ลองใหม่อีกครั้ง" in alert_text or "เตรียมไฟล์" in alert_text:
                time.sleep(60)
                driver.execute_script("""
                    document.querySelectorAll('.el-drawer__wrapper, .el-drawer, .v-modal').forEach(el => el.remove());
                    document.body.style.overflow = 'auto';
                """)
                time.sleep(2)
                export_button = driver.find_element(By.XPATH, "//*[contains(text(), 'ส่งออกไฟล์')]")
                driver.execute_script("arguments[0].click();", export_button)
                try:
                    time.sleep(2)
                    alert2 = driver.switch_to.alert
                    alert2.accept()
                except:
                    pass
        except Exception:
            pass

        print("⏱ กำลังตรวจสอบโฟลเดอร์และรอไฟล์ดาวน์โหลดเข้าดิสก์...")
        downloaded = False
        for i in range(18):
            time.sleep(5)
            files = os.listdir(DOWNLOAD_DIR)
            valid_files = [f for f in files if not f.endswith('.crdownload') and f != '']
            if valid_files:
                downloaded = True
                break
                
            if i == 5 and not downloaded:
                try:
                    export_button = driver.find_element(By.XPATH, "//*[contains(text(), 'ส่งออกไฟล์')]")
                    driver.execute_script("arguments[0].click();", export_button)
                except:
                    pass
        
        files = os.listdir(DOWNLOAD_DIR)
        if downloaded and files:
            latest_file = max([os.path.join(DOWNLOAD_DIR, f) for f in files], key=os.path.getctime)
            if latest_file != file_path:
                os.rename(latest_file, file_path)
            target_data_file = file_path
        else:
            raise FileNotFoundError("บอทหาไฟล์ Excel ที่ดาวน์โหลดไม่เจอ!")

    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการทำงาน: {str(e)}")
        try:
            driver.save_screenshot(os.path.join(SCREENSHOT_DIR, "sales_error_screenshot.png"))
        except Exception:
            pass
        raise e
    finally:
        driver.quit()


# ==========================================================
# ส่วนที่ 2: ดักจับวันที่และเวลาจริงจากภายในไฟล์ Excel
# ==========================================================
print("\n--- เริ่มกระบวนการส่งข้อมูลเข้า Google Cloud BigQuery ด้วย Pandas ---")

print(f"กำลังเปิดอ่านข้อมูลภายในไฟล์: {target_data_file}")

# 1. อ่านส่วน Header ของ Excel เพื่อดึงวันที่ในบรรทัดที่ 5 (From: 06 August 2026 ...)
extracted_date_str = None
try:
    df_raw = pd.read_excel(target_data_file, header=None, nrows=7)
    for row_idx in range(len(df_raw)):
        row_str = " ".join(df_raw.iloc[row_idx].dropna().astype(str))
        if "From" in row_str or "00:00:00" in row_str:
            import re
            date_match = re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})', row_str)
            if date_match:
                extracted_date_str = date_match.group(1)
                print(f"🎯 ดักจับวันที่จาก Header บรรทัดที่ 5 สำเร็จ: {extracted_date_str}")
                break
except Exception as e:
    print(f"⚠️ ไม่สามารถแกะวันที่จาก Header ได้: {e}")

# 2. อ่านข้อมูลตารางจริงตั้งแต่บรรทัดที่ 8 ขึ้นไป
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

# ลบแถวหมายเหตุท้ายตาราง
df = df[~df.iloc[:, 0].astype(str).str.contains('\*\*\*\*|รวมสุทธิ', na=False)]
df = df.dropna(how='all')

df.columns = (
    df.columns.astype(str)
    .str.replace(' ', '_')
    .str.replace('.', '_', regex=False)
    .str.replace('/', '_')
    .str.replace('(', '')
    .str.replace(')', '')
)

# 3. สร้างคอลัมน์ Run_Date โดยอิงวันที่จริงในไฟล์
target_date_for_bq = None

if 'วันที่' in df.columns and len(df) > 0:
    try:
        if 'เวลา' in df.columns:
            df['Run_Date'] = pd.to_datetime(df['วันที่'].astype(str) + ' ' + df['เวลา'].astype(str), errors='coerce')
        else:
            df['Run_Date'] = pd.to_datetime(df['วันที่'], errors='coerce')
            
        first_valid_date = df['Run_Date'].dropna().iloc[0]
        target_date_for_bq = first_valid_date.strftime('%Y-%m-%d')
        print(f"📅 ดักจับวันที่จากคอลัมน์ตารางสำเร็จ: {target_date_for_bq}")
    except Exception as e:
        print(f"⚠️ ไม่สามารถแปลงวันที่จากตารางได้: {e}")

if not target_date_for_bq and extracted_date_str:
    try:
        parsed_dt = pd.to_datetime(extracted_date_str)
        target_date_for_bq = parsed_dt.strftime('%Y-%m-%d')
        df['Run_Date'] = parsed_dt.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        pass

if not target_date_for_bq:
    th_time = datetime.utcnow() + timedelta(hours=7)
    target_date_for_bq = th_time.strftime('%Y-%m-%d')
    df['Run_Date'] = th_time.strftime('%Y-%m-%d %H:%M:%S')

df['Run_Date'] = pd.to_datetime(df['Run_Date']).dt.strftime('%Y-%m-%d %H:%M:%S')


# ==========================================================
# ส่วนที่ 3: ส่งข้อมูลเข้า BigQuery (สั่งลบเฉพาะวันที่ตรงกับในไฟล์)
# ==========================================================
table_id = "stock_data.sales_list"
project_id = "northern-eon-470602-a2"
full_table_path = f"{project_id}.{table_id}"

try:
    client = bigquery.Client(project=project_id)
    
    delete_query = f"""
        DELETE FROM `{full_table_path}`
        WHERE DATE(Run_Date) = '{target_date_for_bq}';
    """
    print(f"🧹 กำลังเคลียร์ข้อมูลเก่าเฉพาะของวันที่ {target_date_for_bq} ใน BigQuery...")
    query_job = client.query(delete_query)
    query_job.result()
    print(f"-> เคลียร์ข้อมูลเดิมของวันที่ {target_date_for_bq} เรียบร้อย!")
except Exception as err:
    print(f"⚠️ ไม่สามารถลบข้อมูลเก่าได้: {str(err)}")

print(f"กำลังส่งข้อมูลจำนวน {len(df)} แถว ของวันที่ {target_date_for_bq} เข้าสู่ BigQuery...")
df.to_gbq(
    destination_table=table_id, project_id=project_id, if_exists="append", progress_bar=False
)

print(f"🎉 🎉 🎉 อัปโหลดสำเร็จ 100%! ข้อมูลของวันที่ {target_date_for_bq} ถูกบันทึกเรียบร้อยครับ")


# ==========================================================
# ส่วนที่ 4: บันทึกวันเวลาลงไฟล์ sales_last_update.txt (เพื่อให้จุดเขียวขึ้นตรงวัน)
# ==========================================================
print("\n--- บันทึกเวลาอัปเดตลงไฟล์สำหรับหน้าเว็บ ---")
try:
    if target_date_for_bq:
        time_str = f"{target_date_for_bq} 23:59:59"
    else:
        th_time = datetime.utcnow() + timedelta(hours=7)
        time_str = th_time.strftime('%Y-%m-%d %H:%M:%S')

    with open("sales_last_update.txt", "w", encoding="utf-8") as f:
        f.write(time_str)
        
    print(f"🎯 บันทึกเวลาสำเร็จ! เวลาที่จะส่งให้หน้าเว็บปักจุดเขียวคือ: {time_str}")

except Exception as e:
    print(f"⚠️ เกิดข้อผิดพลาดในการเขียนไฟล์เวลา: {str(e)}")
