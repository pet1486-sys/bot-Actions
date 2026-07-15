import os
import sys
import requests
from google.cloud import bigquery

# สั่งให้ Python พิมพ์ข้อความเรียงตามบรรทัดจริงบน GitHub Actions
sys.stdout.reconfigure(line_buffering=True)

# ==========================================================
# ส่วนที่ 1: โค้ดดึงไฟล์และสร้างโฟลเดอร์
# ==========================================================
print("กำลังเริ่มดาวน์โหลดไฟล์จากเว็บไซต์...")

# สร้างโฟลเดอร์สำหรับเก็บไฟล์ชั่วคราว
os.makedirs('stock_data', exist_ok=True)
file_path = 'stock_data/SKU.xlsx'

# 🛑 ตรวจสอบและใส่ URL โหลดไฟล์จริงของคุณตรงนี้ครับ
url = "https://example.com/your-download-link.xlsx" 
response = requests.get(url)
with open(file_path, 'wb') as f:
    f.write(response.content)

print(f"ดาวน์โหลดไฟล์สำเร็จและเซฟไว้ที่: {file_path}")


# ==========================================================
# ส่วนที่ 2: โค้ดส่งไฟล์ไป BigQuery
# ==========================================================
# เรียกใช้สิทธิ์การเข้าถึงจากที่เราตั้งค่าไว้ใน GitHub Secrets (GCP_SA_KEY) อัตโนมัติ
client = bigquery.Client()

# 🛑 ตรวจสอบให้มั่นใจว่า ชื่อตาราง sku_list ถูกต้องตรงกับบน BigQuery
table_id = "northern-eon-470602-a2.stock_data.sku_list"

# ตั้งค่าการโหลดไฟล์เข้า BigQuery แบบระบุชนิดไฟล์เข้มงวดด้วยการใช้ string "EXCEL" ป้องกันบั๊ก Enum
job_config = bigquery.LoadJobConfig(
    source_format="EXCEL",            # ใช้ข้อความตัวพิมพ์ใหญ่ "EXCEL" แบบนี้ปลอดภัยที่สุดครับ
    autodetect=True,                  
    write_disposition="WRITE_APPEND", 
)

print(f"กำลังอัปโหลดไฟล์ {file_path} เข้า BigQuery...")

# เริ่มส่งไฟล์เข้า BigQuery
with open(file_path, "rb") as source_file:
    job = client.load_table_from_file(
        source_file, 
        table_id, 
        job_config=job_config
    )

job.result() # รอให้ระบบทำงานจนเสร็จสมบูรณ์

print(f"🎉 อัปโหลดสำเร็จ! ข้อมูลถูกเพิ่มเข้าตาราง {table_id} เรียบร้อยแล้ว")