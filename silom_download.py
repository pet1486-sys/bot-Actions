import os
import requests  # สมมติว่าโค้ดเดิมของคุณใช้ไลบรารีนี้ในการโหลดไฟล์
from google.cloud import bigquery

# ==========================================================
# ส่วนที่ 1: โค้ดดึงไฟล์เดิมของคุณ (คงไว้ตามเดิม)
# ==========================================================
print("กำลังเริ่มดาวน์โหลดไฟล์จากเว็บไซต์...")

# 📁 ตรวจสอบและสร้างโฟลเดอร์สำหรับเก็บไฟล์ชั่วคราว
os.makedirs('stock_data', exist_ok=True)[cite: 1]
file_path = 'stock_data/SKU.xlsx'[cite: 1]

# --- (ใส่โค้ดการดาวน์โหลดของคุณต่อตรงนี้ หรือปรับเปลี่ยนตามการทำงานจริงของคุณ) ---
# ตัวอย่างเช่น:
# url = "http://example.com/data.xlsx"
# response = requests.get(url)
# with open(file_path, 'wb') as f:
#     f.write(response.content)

print(f"ดาวน์โหลดไฟล์สำเร็จและเซฟไว้ที่: {file_path}")[cite: 1]


# ==========================================================
# ส่วนที่ 2: โค้ดส่งไฟล์ไป BigQuery (วิธีที่ 2 วางต่อท้ายได้เลย)
# ==========================================================
# เรียกใช้สิทธิ์การเข้าถึงจากที่ตั้งไว้ใน GitHub Secrets โดยอัตโนมัติ
client = bigquery.Client()

# 🛑 อย่าลืมเปลี่ยนชื่อโปรเจกต์, dataset และชื่อตารางของคุณตรงนี้ให้ตรงกับใน Google Cloud
table_id = "ชื่อ-gcp-project-ของคุณ.ชื่อ_dataset_ของคุณ.ชื่อ_table_ของคุณ"

# ตั้งค่าการโหลดไฟล์
job_config = bigquery.LoadJobConfig(
    source_format=bigquery.SourceFormat.EXCEL, # กำหนดเป็น EXCEL (หรือเปลี่ยนเป็น CSV ถ้าไฟล์จริงเป็น .csv)
    autodetect=True,                           # ให้ BigQuery เดาประเภทข้อมูล (Schema) ให้เอง
    write_disposition="WRITE_APPEND",          # "WRITE_APPEND" = เพิ่มต่อท้ายทุกวัน | "WRITE_TRUNCATE" = ลบของเก่าแล้วเขียนทับใหม่หมด
)

print(f"กำลังอัปโหลดไฟล์ {file_path} เข้า BigQuery...")

# เริ่มส่งไฟล์เข้า BigQuery
with open(file_path, "rb") as source_file:
    job = client.load_table_from_file(source_file, table_id, job_config=job_config)

job.result() # รอให้ระบบทำงานจนเสร็จสมบูรณ์

print(f"🎉 อัปโหลดสำเร็จ! ข้อมูลถูกเพิ่มเข้าตาราง {table_id} เรียบร้อยแล้ว")