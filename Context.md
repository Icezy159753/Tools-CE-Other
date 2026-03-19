# Tools Other CE V1 - คู่มือใช้งานล่าสุด

เอกสารนี้อธิบายวิธีใช้โปรแกรม `Tools Other CE V1` ตามพฤติกรรมล่าสุดของโค้ดในโปรเจกต์นี้ โดยเน้นงาน `Other / อื่นๆ ระบุ`, การแก้ verbatim, การ apply code กลับเข้า rawdata, และการสร้าง codeframe แบบ AI Demo

## 1. ภาพรวมโปรแกรม

โปรแกรมแบ่งเป็น 3 แท็บ:

1. `Phase1 - CodeSheet`
2. `Phase2 - ลง Code`
3. `Phase2 - AI Group Code (Demo)`

หน้าที่ของแต่ละแท็บ:

- `Phase1 - CodeSheet`
  ดึงแถวที่เกี่ยวกับ `Other / อื่นๆ ระบุ` จาก rawdata และสร้าง `CodeSheet.xlsx`

- `Phase2 - ลง Code`
  เอา `CodeSheet.xlsx` ที่แก้เสร็จแล้วไป apply กลับเข้า rawdata แล้ว save เป็น `Rawdata_CE Complete.xlsx`

- `Phase2 - AI Group Code (Demo)`
  ใช้ AI ช่วยจัดกลุ่ม verbatim จาก `CodeSheet.xlsx` เพื่อสร้าง `codeframe.xlsx`

## 2. ไฟล์ในโปรเจกต์

- `app.py`
  GUI หลักของโปรแกรม

- `core.py`
  business logic หลักของระบบ

- `build.bat`
  ใช้ build โปรแกรมเป็น `.exe`

- `Iconapp.ico`
  icon ของโปรแกรม

- `README.md`
  คู่มือแบบย่อ

- `Context.md`
  คู่มือแบบละเอียดไฟล์นี้

- `tests/test_core.py`
  ชุดทดสอบ logic หลักของ `core.py`

## 3. รูปแบบไฟล์ที่ใช้

### Input หลัก

- Rawdata: `.xlsx`
- SPSS Labels: `.sav`
- Coding Sheet: `.xlsx`

### Output หลัก

- `CodeSheet.xlsx`
- `Rawdata_CE Complete.xlsx`
- `codeframe.xlsx`

## 4. หลักการหา Other / Open Text

โปรแกรมจะพยายามหาข้อที่เป็น `Other / อื่นๆ ระบุ` โดยดูจากหลายชั้นของ logic

### 4.1 หา other code จาก SPSS labels

โปรแกรมอ่าน `.sav` แล้วมองหา value labels ที่มีคำประมาณนี้:

- `อื่น`
- `ระบุ`
- `other`
- `specify`
- `others`
- `else`

code ที่มี label ลักษณะนี้จะถูกมองว่าเป็น `Other_Code`

### 4.2 หา open text column จากชื่อคอลัมน์

โปรแกรมจะมองว่าเป็นคอลัมน์ open text ถ้าชื่อมีคำว่า `oth` อยู่ ไม่ว่าจะเป็น:

- `oth`
- `Oth`
- `OTH`
- ตัวพิมพ์ผสมแบบอื่น

ตัวอย่าง pattern ที่รองรับ:

- `q1_oth`
- `q1_oth93`
- `q1_93_oth`
- `q193oth`
- `s14_1_93_oth`
- `q0_95_oth`

### 4.3 การจับคู่กรณีมีหลาย oth column

ถ้าข้อเดียวกันมีหลาย open text columns เช่น:

- `s14_1_93_oth`
- `s14_1_94_oth`
- `s14_1_95_oth`

โปรแกรมจะพยายามจับคู่ตาม `Other_Code` ก่อนเสมอ

ตัวอย่าง:

- `Other_Code = 93` จะจับคู่กับ `s14_1_93_oth`
- `Other_Code = 94` จะจับคู่กับ `s14_1_94_oth`

โปรแกรมจะไม่เอา verbatim ทุกคอลัมน์มาลิสต์รวมกันแบบเดิมแล้ว

### 4.4 กรณี fallback

ถ้าจับ exact pair ไม่ได้:

- ถ้าเหลือ open text ที่ไม่กำกวมแค่ 1 ช่อง อาจ fallback ไปใช้ช่องนั้น
- ถ้ามีหลายช่องพร้อมกัน โปรแกรมจะไม่เดาสุ่ม และจะ skip เคสนั้น

### 4.5 กรณีมี verbatim แต่ไม่มี other code

ถ้าเจอว่า:

- มีข้อความใน open text
- แต่ respondent ไม่ได้เลือก other code

โปรแกรมก็ยังลิสต์แถวนั้นเข้า `CodeSheet.xlsx` ได้

พฤติกรรม:

- `Other_Code` จะว่าง
- `New_Open_Text` จะถูกใส่คำว่า `ตัด` ให้อัตโนมัติ

จุดประสงค์คือกันไม่ให้ verbatim ที่หลุดมาโดยไม่มี code ถูกนำไปใช้ผิดโดยไม่ได้ตั้งใจ

## 5. Phase1 - CodeSheet

### 5.1 หน้าที่

สร้าง `CodeSheet.xlsx` จาก:

- Rawdata `.xlsx`
- SPSS labels `.sav`

### 5.2 วิธีใช้

1. เปิดโปรแกรม
2. ไปที่แท็บ `Phase1 - CodeSheet`
3. เลือกไฟล์ `Rawdata Excel (.xlsx)`
4. เลือกไฟล์ `SPSS Labels (.sav)`
5. กดปุ่ม `CodeSheet`

### 5.3 Output

ไฟล์ output default คือ:

- `CodeSheet.xlsx`

จะถูกสร้างในโฟลเดอร์เดียวกับ rawdata

### 5.4 ถ้ามี CodeSheet เดิมอยู่แล้ว

ถ้าในโฟลเดอร์มี `CodeSheet.xlsx` เดิมอยู่ โปรแกรมจะพยายาม merge ให้

หลักการ:

- รักษาแถวเดิมไว้ก่อน
- append แถวใหม่ต่อท้าย
- พยายามคงค่าที่เคยลงไว้ใน:
  - `New_Code`
  - `New_Open_Text`
  - `Remark`

ดังนั้น workflow ที่ถูกต้องเวลามี rawdata รอบใหม่เข้ามา คือ:

1. ใช้ `CodeSheet.xlsx` เดิม
2. โหลด rawdata ใหม่ใน `Phase1 - CodeSheet`
3. export ซ้ำ
4. โปรแกรมจะเพิ่มเฉพาะแถวใหม่ต่อท้าย

## 6. โครงสร้างคอลัมน์ใน CodeSheet

คอลัมน์หลัก:

- `Question`
  ชื่อคอลัมน์คำถามต้นทาง

- `Variable_Label`
  label ของคำถามจาก SPSS

- `Sbjnum`
  respondent id

- `Other_Label`
  label ของ other code

- `Other_Code`
  code เดิมของ other

- `New_Code`
  ช่องให้ทีมลง code ใหม่

- `Open_Text`
  verbatim เดิม

- `New_Open_Text`
  ช่องให้ทีมแก้ข้อความ verbatim ใหม่

- `Open_Text_From`
  ชื่อคอลัมน์ต้นทางที่ดึง open text มา

- `Remark`
  หมายเหตุเพิ่มเติม

### สีในไฟล์

- `New_Code` เป็นช่องสีเหลือง
- `New_Open_Text` เป็นช่องสีเหลือง
- `Open_Text` คือข้อความก่อนแก้
- `New_Open_Text` คือข้อความหลังแก้

## 7. วิธีลงข้อมูลใน CodeSheet

### 7.1 ถ้าต้องการเปลี่ยน code

กรอกค่าใน `New_Code`

ตัวอย่าง:

- เดิม `Other_Code = 97`
- ต้องการเปลี่ยนเป็น `5`
- ใส่ `5` ใน `New_Code`

### 7.2 ถ้าต้องการแก้ verbatim

กรอกข้อความใหม่ใน `New_Open_Text`

ตัวอย่าง:

- `Open_Text = ชาเขียวว`
- `New_Open_Text = ชาเขียว`

### 7.3 ถ้าต้องการลบ code

ใส่คำว่า:

- `ตัด`

ในคอลัมน์ `New_Code`

ผล:

- โปรแกรมจะลบ code ของแถวนั้นออก

### 7.4 ถ้าต้องการลบ verbatim

ใส่คำว่า:

- `ตัด`

ในคอลัมน์ `New_Open_Text`

ผล:

- โปรแกรมจะลบข้อความ verbatim ของแถวนั้น
- และถ้า code ในคำถามตรงกับ `Other_Code` ของแถวนั้น โปรแกรมจะลบ code นั้นออกด้วย

ตัวอย่าง:

- แถวนี้เป็น `Other_Code = 93`
- ใส่ `ตัด` ใน `New_Open_Text`
- โปรแกรมจะลบ verbatim ของ `93`
- และลบ code `93` ออกจากคำถามนั้น
- code อื่นในแถวเดียวกันจะไม่ถูกลบ

## 8. Phase2 - ลง Code

### 8.1 หน้าที่

เอา `CodeSheet.xlsx` ที่แก้เสร็จแล้ว ไป apply กลับเข้า rawdata

### 8.2 วิธีใช้

1. ไปที่แท็บ `Phase2 - ลง Code`
2. เลือก `Rawdata ต้นฉบับ (.xlsx)`
3. เลือก `Coding Sheet (.xlsx)`
4. ถ้าต้องการเปลี่ยนชื่อ output ให้กด browse
5. ถ้าไม่เปลี่ยน โปรแกรมจะใช้ชื่อ default:
   `Rawdata_CE Complete.xlsx`
6. กดปุ่ม `ลง Code`

### 8.3 Output

ไฟล์ output คือ:

- `Rawdata_CE Complete.xlsx`

หมายเหตุ:

- โปรแกรมไม่ save `recode_log.xlsx` แยกแล้ว
- แต่ preview log ยังแสดงใน GUI

### 8.4 สิ่งที่โปรแกรม apply

- ถ้ามีค่าใน `New_Code`
  โปรแกรมจะใช้ค่านั้นแทน code เดิม

- ถ้า `New_Code = ตัด`
  โปรแกรมจะลบ code เดิมออก

- ถ้ามีค่าใน `New_Open_Text`
  โปรแกรมจะใช้ข้อความนั้นเขียนทับ verbatim เดิม

- ถ้า `New_Open_Text = ตัด`
  โปรแกรมจะลบ verbatim เดิมออก
  และอาจลบ code ที่ match ด้วยตาม logic

## 9. Phase2 - AI Group Code (Demo)

### 9.1 หน้าที่

ใช้ AI ช่วยจัดกลุ่ม `Open_Text` เพื่อสร้าง codeframe แบบ draft

### 9.2 วิธีใช้

1. ไปที่แท็บ `Phase2 - AI Group Code (Demo)`
2. เลือก `Coding Sheet (.xlsx)` ที่มี `Open_Text`
3. ใส่ `OpenRouter API Key`
4. เลือก model หรือใช้ default
5. กด `AI Group Code (Demo)`

### 9.3 Output

จะได้ไฟล์:

- `codeframe.xlsx`

### 9.4 โครงสร้าง codeframe

แต่ละชีตมีลักษณะคล้าย template ที่ตกลงไว้

มี:

- `Index`
- `Back to Index`
- codeframe แยกตามข้อ

คอลัมน์หลัก:

- `Code No.`
- `Thai Group1`
- `Thai Group2`
- `English`
- `Count`

ความหมาย:

- `Thai Group1`
  รวมข้อความดิบทั้งหมดในกลุ่ม คั่นด้วย `/`

- `Thai Group2`
  ชื่อกลุ่มความหมายเดียวกันที่ AI จัดให้

- `English`
  label อังกฤษแบบสั้น

- `Count`
  จำนวนแถวในกลุ่ม

หมายเหตุ:

- AI ส่วนนี้เป็น `Demo`
- ควรใช้เป็น draft / starting point
- ควรมีคนตรวจต่อก่อนใช้งานจริง

## 10. การ build เป็น EXE

### วิธี build

ใช้:

```bat
build.bat
```

### ผลลัพธ์

ไฟล์ exe จะอยู่ที่:

```text
dist\Tools Other CE V1.exe
```

### เรื่อง icon

โปรแกรมตั้งค่า icon ไว้ทั้ง:

- `QApplication`
- `MainWindow`
- Windows AppUserModelID
- Win32 small/big icon
- PyInstaller `--icon`
- bundle ไฟล์ icon เข้า exe

ถ้า taskbar ยังแสดง icon เก่า:

1. ปิดโปรแกรมทั้งหมด
2. unpin ของเก่าออกจาก taskbar
3. build ใหม่
4. เปิด exe จาก `dist` ตรงๆ
5. ค่อย pin ใหม่

## 10.1 ระบบ Check for Updates

โปรแกรมรองรับระบบเช็กอัปเดตแล้ว โดยใช้ `GitHub Releases`

หลักการ:

1. โปรแกรมอ่าน `update_config.json` ที่อยู่ข้างไฟล์ exe
2. ใน config จะมีชื่อ repo GitHub
3. โปรแกรมเรียก `latest release` จาก GitHub API
4. ถ้า tag เวอร์ชันใหม่กว่าเวอร์ชันปัจจุบัน โปรแกรมจะแจ้งผู้ใช้
5. ผู้ใช้กดดาวน์โหลด exe ตัวใหม่ได้จาก release asset

ไฟล์ตัวอย่างในโปรเจกต์:

- `update_config.example.json`
- `.github/workflows/release.yml`

ตัวอย่าง `update_config.json`

```json
{
  "provider": "github",
  "repo": "Icezy159753/Tools-CE-Other",
  "asset_name": "Tools Other CE V1.exe",
  "updater_asset_name": "Tools Other CE Updater.exe",
  "auto_check": true
}
```

หมายเหตุ:

- ถ้าไม่ได้ตั้ง `repo` ปุ่ม `Check Update` จะบอกว่ายังไม่ตั้งค่า
- สามารถเปิด auto check ได้ด้วย `auto_check: true`
- โปรแกรมจะใช้ `Updater.exe` แยกอีกตัวเพื่อแทนไฟล์เดิมให้อัตโนมัติ

### 10.2 GitHub Actions สำหรับ build release

ในโปรเจกต์มีไฟล์:

- `.github/workflows/release.yml`

workflow นี้จะ:

1. รันเมื่อ push tag เช่น `v1.0.1`
2. ติดตั้ง dependencies
3. รัน test
4. build `Tools Other CE V1.exe`
5. build `Tools Other CE Updater.exe`
6. อัปโหลดไฟล์เข้า GitHub Release

ดังนั้น workflow แนะนำคือ:

1. แก้โค้ดในเครื่อง
2. push ขึ้น GitHub
3. สร้าง tag เช่น `v1.0.1`
4. push tag ขึ้น GitHub
5. รอ Action build release
6. ผู้ใช้เปิดโปรแกรมแล้วกด `Check Update` หรือให้โปรแกรมเช็กเอง
7. ถ้ามีเวอร์ชันใหม่ โปรแกรมจะเด้งถามและอัปเดตแทนไฟล์เดิมได้เลย

## 11. การทดสอบ (Tests)

มีชุด test สำหรับ `core.py` แล้ว

ไฟล์:

- `tests/test_core.py`

### วิธีรัน

```bash
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

### ตอนนี้ test ครอบคลุมอะไรบ้าง

- parse ชื่อคอลัมน์ `oth` หลายรูปแบบ
- หา pair `Question -> oth`
- exact mapping `Other_Code -> indexed oth`
- verbatim-only ทั้งแบบธรรมดาและแบบ indexed
- skip เคส ambiguous หลาย `oth`
- merge `CodeSheet` เดิมกับข้อมูลใหม่
- save/read `CodeSheet` หลายชีต
- `New_Code` ปกติ
- `New_Code = ตัด`
- `New_Open_Text` ปกติ
- `New_Open_Text = ตัด`
- path error บางกรณี เช่น `sbjnum not found`
- integration ของ `phase1_export` แบบ mock SPSS

### ควรรัน test ตอนไหน

- หลังแก้ `core.py`
- หลังแก้ logic `oth`
- หลังแก้ logic `New_Code` หรือ `New_Open_Text`
- ก่อน build exe
- ก่อนส่งให้คนอื่นใช้

### ถ้าแก้แค่ GUI

ถ้าแก้เฉพาะ `app.py` เรื่องหน้าตา เช่น label, ปุ่ม, สี, layout
อาจไม่จำเป็นต้องรัน test ทุกครั้ง

แต่ถ้าแตะ `core.py` แนะนำให้รันทันที

## 12. ข้อควรระวัง

### 12.1 อย่าแก้ `Open_Text` ตรงๆ ถ้าต้องการเก็บ before/after

ให้ใช้:

- `Open_Text` เป็น before
- `New_Open_Text` เป็น after

### 12.2 ถ้าจะลบ ให้ใส่ `ตัด` ในช่องที่ถูกต้อง

ใส่ได้เฉพาะ:

- `New_Code`
- `New_Open_Text`

ไม่ควรใส่ใน:

- `Open_Text`

### 12.3 ถ้ามี rawdata รอบใหม่

ให้ใช้ `CodeSheet.xlsx` เดิมและ export ซ้ำ
โปรแกรมจะ merge และ append แถวใหม่ให้

### 12.4 AI Group Code ยังเป็น Demo

อย่าถือเป็น output final โดยไม่ตรวจซ้ำ

## 13. Workflow แนะนำ

### กรณีงานปกติ

1. เปิด `Phase1 - CodeSheet`
2. เลือก rawdata + sav
3. กด `CodeSheet`
4. เปิด `CodeSheet.xlsx`
5. ลง `New_Code` และ/หรือ `New_Open_Text`
6. ถ้าต้องการลบ ให้ใช้คำว่า `ตัด`
7. เปิด `Phase2 - ลง Code`
8. เลือก rawdata + coding sheet
9. กด `ลง Code`
10. ได้ `Rawdata_CE Complete.xlsx`

### กรณีมี rawdata รอบใหม่เข้ามา

1. ใช้ `CodeSheet.xlsx` เดิม
2. เอา rawdata ใหม่เข้า `Phase1 - CodeSheet`
3. export ซ้ำ
4. โปรแกรมจะ merge ของเดิมและ append แถวใหม่
5. ลง code เพิ่มเฉพาะแถวใหม่

### กรณีอยากใช้ AI ช่วย grouping

1. เตรียม `CodeSheet.xlsx` ที่มี `Open_Text`
2. เปิด `Phase2 - AI Group Code (Demo)`
3. ใส่ API key
4. กดรัน
5. ตรวจ `codeframe.xlsx`

## 14. สรุป

จุดเด่นของเวอร์ชันปัจจุบันคือ:

- รองรับ `oth` หลายรูปแบบ
- รองรับ indexed open text
- รองรับ before/after verbatim
- รองรับคำว่า `ตัด`
- merge `CodeSheet` เดิมกับข้อมูลใหม่ได้
- มี test สำหรับ core logic แล้ว

ถ้ามีการเปลี่ยน logic ในอนาคต ควรอัปเดต `README.md` และ `Context.md` ตามทุกครั้ง
