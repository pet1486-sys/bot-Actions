import os
import requests
from google.cloud import bigquery

# ==========================================================
# ส่วนที่ 1: โค้ดดึงไฟล์เดิมของคุณ (คงไว้ตามเดิม)
# ==========================================================
print("กำลังเริ่มดาวน์โหลดไฟล์จากเว็บไซต์...")

# 📁 สร้างโฟลเดอร์สำหรับเก็บไฟล์ชั่วคราวแบบสะอาด
os.makedirs('stock_data', exist_ok=True)
file_path = 'stock_data/SKU.xlsx'

# --- (ใส่โค้ดการดาวน์โหลดของคุณต่อตรงนี้ หรือปรับเปลี่ยนตามการทำงานจริงของคุณ) ---
# ตัวอย่างเช่น:
# url = "http://example.com/data.xlsx"
# response = requests.get(url)
# with open(file_path, 'wb') as f:
#     f.write(response.content)

print(f"ดาวน์โหลดไฟล์สำเร็จและเซฟไว้ที่: {file_path}")


# ==========================================================
# ส่วนที่ 2: โค้ดส่งไฟล์ไป BigQuery (วิธีที่ 2)
# ==========================================================
# เรียกใช้สิทธิ์การเข้าถึงจากที่เราตั้งค่าไว้ใน GitHub Secrets (GCP_SA_KEY) อัตโนมัติ
client = bigquery.Client()

# 🛑 อย่าลืมเปลี่ยนชื่อโปรเจกต์, dataset และชื่อตารางของคุณตรงนี้ให้ตรงกับใน Google Cloud
table_id = "ชื่อ-gcp-project-ของคุณ.ชื่อ_dataset_ของคุณ.ชื่อ_table_ของคุณ"

# ตั้งค่าการโหลดไฟล์
job_config = bigquery.LoadJobConfig(
    source_format="EXCEL",            # กำหนดเป็นสตริง "EXCEL" เพื่อป้องกันการหาแอตทริบิวต์ไม่เจอ
    autodetect=True,                  # ให้ BigQuery เดาประเภทข้อมูล (Schema) ให้เอง
    write_disposition="WRITE_APPEND", # "WRITE_APPEND" = เพิ่มต่อท้ายทุกวัน | "WRITE_TRUNCATE" = ลบของเก่าแล้วเขียนทับใหม่หมด
)

print(f"กำลังอัปโหลดไฟล์ {file_path} เข้า BigQuery...")

# เริ่มส่งไฟล์เข้า BigQuery
with open(file_path, "rb") as source_file:
    job = client.load_table_from_file(source_file, table_id, job_config=job_config)

job.result() # รอให้ระบบทำงานจนเสร็จสมบูรณ์

print(f"🎉 อัปโหลดสำเร็จ! ข้อมูลถูกเพิ่มเข้าตาราง {table_id} เรียบร้อยแล้ว")