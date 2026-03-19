# Tools Other CE V1

โปรแกรมสำหรับจัดการงาน `Other / อื่นๆ ระบุ` จาก rawdata งานวิจัย โดยเน้น workflow 2 ส่วนหลัก:

1. สร้าง `CodeSheet.xlsx` จาก `Rawdata + SPSS Labels`
2. Apply code / แก้ verbatim กลับเข้า `Rawdata_CE Complete.xlsx`

มีแท็บ AI เพิ่มสำหรับช่วยจัดกลุ่ม verbatim เป็น codeframe เบื้องต้น แต่เป็น `Demo`

## ไฟล์สำคัญ

- `app.py` : GUI หลัก
- `core.py` : business logic หลักของโปรแกรม
- `build.bat` : ใช้ build เป็น `.exe`
- `Iconapp.ico` : icon ของโปรแกรม
- `Context.md` : คู่มือใช้งานฉบับละเอียด
- `tests/test_core.py` : ชุดทดสอบ logic หลัก

## การติดตั้ง

```bash
pip install -r requirements.txt
```

## การรันโปรแกรม

```bash
python app.py
```

## โครงหน้าจอ

- `Phase1 - CodeSheet`
- `Phase2 - ลง Code`
- `Phase2 - AI Group Code (Demo)`

## Workflow หลัก

### 1. Phase1 - CodeSheet

ใช้สำหรับอ่าน:

- Rawdata `.xlsx`
- SPSS Labels `.sav`

แล้วสร้าง `CodeSheet.xlsx`

สิ่งที่โปรแกรมทำ:

- หา code ที่เป็น `อื่นๆ ระบุ` จาก value labels ใน `.sav`
- หา open text column จากชื่อที่มี `oth`
- จับคู่ `Other_Code` กับ `Open_Text`
- รองรับทั้ง direct pair และ indexed pair เช่น `93 -> s14_1_93_oth`
- ถ้ามี verbatim แต่ไม่มี other code ก็ยังลิสต์เข้า `CodeSheet` ได้
- ถ้าเป็นเคส verbatim-only โปรแกรมจะใส่ `ตัด` ใน `New_Open_Text` ให้อัตโนมัติ

ชื่อไฟล์ output default:

- `CodeSheet.xlsx`

### 2. Phase2 - ลง Code

ใช้สำหรับอ่าน:

- Rawdata ต้นฉบับ
- `CodeSheet.xlsx` ที่ทีมลง code แล้ว

แล้วสร้าง:

- `Rawdata_CE Complete.xlsx`

สิ่งที่ทำได้:

- เปลี่ยน code ด้วย `New_Code`
- ลบ code ด้วย `New_Code = ตัด`
- แก้ verbatim ด้วย `New_Open_Text`
- ลบ verbatim ด้วย `New_Open_Text = ตัด`
- ถ้า `New_Open_Text = ตัด` และ code ในคำถามตรงกับ `Other_Code` โปรแกรมจะลบ code นั้นออกให้ด้วย

หมายเหตุ:

- ตอนนี้โปรแกรมไม่ save `recode_log.xlsx` แยกออกมาแล้ว
- แต่ preview log ยังดูได้ในหน้าโปรแกรม

### 3. Phase2 - AI Group Code (Demo)

ใช้ `CodeSheet.xlsx` ที่มี `Open_Text` ไปสร้าง `codeframe.xlsx`

output จะมี:

- `Index` sheet
- `Back to Index`
- codeframe แยกตามข้อ

ความหมายคอลัมน์หลัก:

- `Thai Group1` = ข้อความดิบทั้งหมดในกลุ่ม คั่นด้วย `/`
- `Thai Group2` = ชื่อกลุ่มสรุปความหมายเดียวกัน
- `English` = label อังกฤษสั้นๆ
- `Count` = จำนวน

## คอลัมน์สำคัญใน CodeSheet

- `Question`
- `Variable_Label`
- `Sbjnum`
- `Other_Label`
- `Other_Code`
- `New_Code`
- `Open_Text`
- `New_Open_Text`
- `Open_Text_From`
- `Remark`

ความหมาย:

- `Open_Text` คือข้อความเดิม
- `New_Open_Text` คือข้อความหลังแก้
- `New_Code` และ `New_Open_Text` เป็นช่องสีเหลืองสำหรับแก้ไข

## การใช้คำว่า `ตัด`

รองรับ 2 ช่อง:

- `New_Code = ตัด`
  โปรแกรมจะลบ code ของแถวนั้น

- `New_Open_Text = ตัด`
  โปรแกรมจะลบ verbatim ของแถวนั้น
  และถ้า code ในคำถามตรงกับ `Other_Code` ก็จะลบ code นั้นด้วย

## การ merge CodeSheet เดิมกับข้อมูลใหม่

ถ้ามี `CodeSheet.xlsx` เดิมอยู่แล้ว แล้วนำ rawdata ชุดใหม่มา export ซ้ำ:

- แถวเดิมจะอยู่ก่อน
- แถวใหม่จะถูก append ต่อท้าย
- ค่าที่เคยลงใน `New_Code`, `New_Open_Text`, `Remark` จะพยายามคงไว้

## การ detect `oth`

โปรแกรมถือว่าเป็น open text column ถ้าชื่อคอลัมน์มี `oth` อยู่ ไม่ว่าจะเป็น:

- `oth`
- `Oth`
- `OTH`

และรองรับหลาย pattern เช่น:

- `q1_oth`
- `q1_oth93`
- `q1_93_oth`
- `q193oth`

## การทดสอบ

มี test สำหรับ logic หลักใน `core.py`

รันได้ด้วย:

```bash
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

แนะนำให้รัน test เมื่อ:

- แก้ `core.py`
- เปลี่ยน logic `oth`
- เปลี่ยน logic `New_Code` / `New_Open_Text`
- ก่อน build exe

## การ build เป็น EXE

ใช้:

```bat
build.bat
```

output:

```text
dist\Tools Other CE V1.exe
```

โปรแกรมตั้งค่า icon ไว้แล้วทั้งใน app และ exe แต่ถ้า taskbar ยังโชว์ icon เก่า:

1. ปิดโปรแกรมให้หมด
2. unpin ของเก่าออกจาก taskbar
3. build ใหม่
4. เปิด exe จาก `dist` ตรงๆ
5. ค่อย pin ใหม่

## ระบบ Check for Updates

โปรแกรมรองรับการเช็กอัปเดตจาก `GitHub Releases` แล้ว

ไฟล์ที่เกี่ยวข้อง:

- `update_config.example.json`
- `.github/workflows/release.yml`

แนวคิดการใช้งาน:

1. push code ขึ้น GitHub repo
2. สร้าง tag เช่น `v1.0.1`
3. GitHub Actions จะ build `.exe` และอัปโหลดเข้า Release
4. ผู้ใช้กดปุ่ม `Check Update` ในโปรแกรม
5. โปรแกรมจะเช็ก `latest release` ของ GitHub
6. ถ้ามีเวอร์ชันใหม่ โปรแกรมจะให้ดาวน์โหลด `.exe` ตัวล่าสุด

รูปแบบ `update_config.json`

```json
{
  "provider": "github",
  "repo": "Icezy159753/Tools-CE-Other",
  "asset_name": "Tools Other CE V1.exe",
  "updater_asset_name": "Tools Other CE Updater.exe",
  "auto_check": true
}
```

ไฟล์ workflow ที่ใช้ปล่อย release:

- `.github/workflows/release.yml`

release จะมี asset หลัก 2 ตัว:

- `Tools Other CE V1.exe`
- `Tools Other CE Updater.exe`

flow อัปเดตปัจจุบัน:

1. โปรแกรมเปิดมาแล้วเช็ก GitHub Release อัตโนมัติ
2. ถ้ามีเวอร์ชันใหม่ จะเด้งถามให้อัปเดต
3. ถ้าผู้ใช้กดตกลง โปรแกรมจะดาวน์โหลดทั้งตัวโปรแกรมใหม่และ updater
4. โปรแกรมจะปิดตัวเอง
5. `Updater.exe` จะมาแทนที่ไฟล์เดิม
6. จากนั้นเปิดเวอร์ชันใหม่ให้อัตโนมัติ

## หมายเหตุ

- ถ้าต้องการคู่มือแบบละเอียด ให้ดู [Context.md](/abs/path/c:/Users/songklod/Desktop/Test%20Other/Context.md)
- AI Group Code เป็นตัวช่วย draft ไม่ควรใช้แทน human check ทั้งหมด
