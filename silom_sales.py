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

# 🌟 เช็กสถานะการสั่งงานว่าเป็นการ Manual Upload หรือไม่
IS_MANUAL = os.environ.get('IS_MANUAL_UPLOAD', '').lower() == 'true'
MANUAL_FILE_NAME = os.environ.get('MANUAL_FILE_NAME', '')

print(f"🔍 ตรวจสอบโหมดการทำงาน: IS_MANUAL_UPLOAD = {IS_MANUAL}")

# ==========================================================
# ส่วนที่ 1: การเตรียมไฟล์ข้อมูล (อัตโนมัติ vs Manual)
# ==========================================================
if IS_MANUAL:
    print("\n📁 --- โหมด Manual Upload: ข้ามกระบวนการ Selenium/Silom POS ---")
    
    # ค้นหาไฟล์ที่ผู้ใช้อัปโหลดเข้ามาผ่านหน้าเว็บ
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
        
        print("🧼 เริ่มกระบวนการเคลียร์หน้าจอ (ลบกล่องแชท Crisp)...")
        try:
            driver.execute_script("""
                var crispElements = document.querySelectorAll(
                    '#crisp-chat-box, .crisp-client, [class^="crisp-"], [id^="crisp-"], ' +
                    'iframe[title*="chat" i], iframe[src*="crisp"], iframe[id*="crisp"]'
                );
                crispElements.forEach(function(el) { el.remove(); });
            """)
            print("-> ลบกล่องแชท Crisp เรียบร้อย")
            time.sleep(1)
        except Exception:
            pass

        print("🧼 เริ่มกระบวนการเคลียร์ป๊อปอัป (ซ่อนแถบคู่มือฝั่งขวา)...")
        try:
            driver.execute_script("""
                document.querySelectorAll('.el-drawer__wrapper, .el-drawer, .v-modal').forEach(el => el.remove());
                document.body.style.overflow = 'auto';
            """)
            print("-> บังคับลบแถบคู่มือฝั่งขวาเรียบร้อย!")
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
            print(f"⚠️ เจอ Alert ของระบบ: {alert_text}")
            alert.accept() 
            
            if "ลองใหม่อีกครั้ง" in alert_text or "เตรียมไฟล์" in alert_text:
                waiting_time = 60  
                print(f"⏳ ระบบหลังบ้านยังไม่พร้อม... บอทจะหยุดรอ {waiting_time} วินาที...")
                time.sleep(waiting_time)
                
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
                    print(f"⚠️ เจอ Alert รอบสอง: {alert2.text}")
                    alert2.accept()
                except:
                    print("✅ รอบสองไม่เจอ Alert เพิ่มเติม")
                    
        except Exception:
            print("✅ ยิงคำสั่งดาวน์โหลดสำเร็จในรอบแรก!")

        print("⏱ กำลังตรวจสอบโฟลเดอร์และรอไฟล์ดาวน์โหลดเข้าดิสก์...")
        downloaded = False
        for i in range(18):
            time.sleep(5)
            files = os.listdir(DOWNLOAD_DIR)
            valid_files = [f for f in files if not f.endswith('.crdownload') and f != '']
            if valid_files:
                print(f"พบไฟล์ดาวน์โหลดในรอบที่ {i+1}: {valid_files}")
                downloaded = True
                break
                
            if i == 5 and not downloaded:
                try:
                    export_button = driver.find_element(By.XPATH, "//*[contains(text(), 'ส่งออกไฟล์')]")
                    driver.execute_script("arguments[0].click();", export_button)
                except:
                    pass
                    
            print(f"รอบที่ {i+1}: ยังไม่พบไฟล์ กำลังรอต่อ...")
        
        files = os.listdir(DOWNLOAD_DIR)
        if downloaded and files:
            latest_file = max([os.path.join(DOWNLOAD_DIR, f) for f in files], key=os.path.getctime)
            if latest_file != file_path:
                os.rename(latest_file, file_path)
            print(f"เตรียมอัปโหลดไฟล์เสร็จสมบูรณ์ที่: {file_path}")
            target_data_file = file_path
        else:
            raise FileNotFoundError(f"บอทหาไฟล์ Excel ที่ดาวน์โหลดไม่เจอ!")

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
# ส่วนที่ 2: โค้ดส่งไฟล์จริงเข้า BigQuery ด้วย Pandas
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

print("กำลังตรวจสอบและลบแถวข้อความหมายเหตุท้ายตาราง...")
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

# คำนวณเวลาไทยปัจจุบันสำหรับระบุวันเก็บข้อมูล
th_time = datetime.utcnow() + timedelta(hours=7)
df['Run_Date'] = th_time.strftime('%Y-%m-%d %H:%M:%S')

table_id = "stock_data.sales_list"
project_id = "northern-eon-470602-a2"
full_table_path = f"{project_id}.{table_id}"

# สั่งลบข้อมูลเก่าของวันนี้ออกจาก BigQuery ก่อนเพื่อเคลียร์ทาง
try:
    client = bigquery.Client(project=project_id)
    today_str = th_time.strftime('%Y-%m-%d')
    
    delete_query = f"""
        DELETE FROM `{full_table_path}`
        WHERE DATE(Run_Date) = '{today_str}';
    """
    print(f"🧹 กำลังเคลียร์ข้อมูลเก่าของวันที่ {today_str} ใน BigQuery เพื่อป้องกันการบันทึกซ้ำ...")
    query_job = client.query(delete_query)
    query_job.result()
    print("-> ลบข้อมูลรอบเดิมของวันนี้เรียบร้อยแล้ว!")
except Exception as err:
    print(f"⚠️ ไม่สามารถลบข้อมูลเก่าได้: {str(err)}")

# โยนข้อมูลชุดล่าสุดเติมลงไป
print(f"กำลังส่งข้อมูลชุดล่าสุดจำนวน {len(df)} แถว เข้าสู่ BigQuery ตาราง {full_table_path}...")
df.to_gbq(
    destination_table=table_id, project_id=project_id, if_exists="append", progress_bar=False
)

print(f"🎉 🎉 🎉 อัปโหลดสำเร็จ 100%! อัปเดตข้อมูลเป็นเวอร์ชันล่าสุดเรียบร้อยแล้วครับ")


# ==========================================================
# ส่วนที่ 3: ดึงเวลาอัปเดตล่าสุดจาก BigQuery และเซฟลงไฟล์ข้อความ
# ==========================================================
print("\n--- เริ่มกระบวนการดึงเวลาแก้ไขล่าสุดจาก BigQuery เพื่อส่งให้หน้าเว็บ ---")
try:
    client = bigquery.Client(project=project_id)
    query = f"""
        SELECT TIMESTAMP_MILLIS(last_modified_time) AS last_updated
        FROM `{project_id}.stock_data.__TABLES__`
        WHERE table_id = 'sales_list';
    """
    query_job = client.query(query)
    results = query_job.result()
    
    for row in results:
        utc_time = row.last_updated
        thai_time = utc_time + timedelta(hours=7)
        time_str = thai_time.strftime('%Y-%m-%d %H:%M:%S')
        
        with open("sales_last_update.txt", "w", encoding="utf-8") as f:
            f.write(time_str)
        print(f"🎯 ดึงข้อมูลสำเร็จ! เวลาแก้ไขจริงใน BigQuery คือ: {time_str} น. (บันทึกลลงไฟล์แล้ว)")

except Exception as e:
    # หากดึงจาก BigQuery ไม่สำเร็จ ให้เขียนเวลาไทยปัจจุบันลงไฟล์โดยตรงเพื่อไม่ให้จุดเขียวพลาด
    time_str = th_time.strftime('%Y-%m-%d %H:%M:%S')
    with open("sales_last_update.txt", "w", encoding="utf-8") as f:
        f.write(time_str)
    print(f"⚠️ เขียนเวลาปัจจุบันลงไฟล์สำรอง: {time_str} น.")
