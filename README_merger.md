# Dispatch + SIR Merger — คู่มือการใช้งาน

## มีไฟล์อะไรบ้าง

| ไฟล์ | คำอธิบาย |
|------|-----------|
| `dispatch_sir_merger_gui.py` | **แนะนำ** — แอปหน้าต่าง GUI คลิกง่าย |
| `dispatch_sir_merge.py` | Command-line script (แก้ CONFIG แล้วรันได้เลย) |

---

## ติดตั้งครั้งแรก (ทำแค่ครั้งเดียว)

เปิด Command Prompt แล้วพิมพ์:

```
pip install pyxlsb pandas openpyxl
```

---

## วิธีใช้ GUI (แนะนำ)

1. ดับเบิลคลิก `dispatch_sir_merger_gui.py`  
   (หรือ Right-click → "Open with Python")

2. เลือกไฟล์:
   - **Dispatch File** → เลือกไฟล์ `.xlsb` ที่ได้รับมา
   - **SIR Root Folder** → `O:\2026 Official SIR` (default แล้ว)
   - **Output Folder** → โฟลเดอร์ที่จะบันทึกผล

3. **Week No.** → ใส่ `auto` ระบบจะหา folder เลขสูงสุดอัตโนมัติ  
   หรือใส่เลข เช่น `21` เพื่อเจาะจง week

4. กด **▶ รัน Merge**

5. ผลลัพธ์จะบันทึกเป็น `Dispatch_merged_week21_YYYYMMDD_HHMMSS.xlsx`

---

## Logic การทำงาน

```
Dispatch (.xlsb)          SIR (.zip → .xlsx)
  คอลัมน์ A–BE              คอลัมน์ Article, Category,
  Article (col Y)    ──►    SubCategory, Cost, Price,
  เป็น key match             globalDepartmentName,
                             globalCategoryName,
                             globalBrandName
                                    │
                                    ▼
              ไฟล์ผลลัพธ์: คอลัมน์ A–BE + BF–BM (SIR data)
```

---

## Folder SIR จะรันขึ้นทุกวันจันทร์

ระบบจะ scan โฟลเดอร์ `O:\2026 Official SIR\` และเลือก folder ที่มีเลขสูงสุดโดยอัตโนมัติ:

```
O:\2026 Official SIR\
  ├── 21\   ← Official_SIR 26_21 All.zip
  ├── 22\   ← Official_SIR 26_22 All.zip   (สัปดาห์หน้า)
  ├── 23\   ← ...
```

เมื่อ IT เพิ่ม folder 22, 23, ... ระบบจะหยิบ folder ล่าสุดให้เอง ไม่ต้องแก้ code

---

## Match Rate ต่ำ — ทำอย่างไร?

ถ้า log บอก match น้อยกว่า 80% ให้ตรวจสอบ:

1. Format ของ Article ใน SIR — บางครั้งเป็น `4516650` บางครั้งเป็น `'4516650`
2. Sheet name ใน SIR — แก้ `SIR_SHEET = 0` เป็นชื่อ sheet เช่น `SIR_SHEET = "sir"`
3. Column name ใน SIR — script จะ log ชื่อคอลัมน์ที่พบให้ดู

---

## ต้องการความช่วยเหลือเพิ่ม

แจ้ง error message หรือ screenshot มาได้เลยครับ
