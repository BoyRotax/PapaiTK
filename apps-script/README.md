# Apps Script (ฝั่งหลังบ้าน)

`Code.gs` คือสคริปต์ Google Apps Script ฉบับแก้ไข สำหรับคัดลอกไปวางทับของเดิมในโปรเจกต์

## สิ่งที่แก้จากฉบับเดิม
1. **เพิ่ม `getAttendanceReport`** (และบรรทัดเรียกใน `doGet`) ใช้กับหน้า `attendance_report.html` (ใบเช็คชื่อรายเดือน PDF)
2. **แก้ `saveRoutineData`** ให้หาแถวเดิมเจอ: เดิมเทียบวันที่ที่มี `'` นำหน้ากับค่าในชีตที่ไม่มี `'` จึงไม่เคยตรงและเพิ่มแถวซ้ำทุกครั้ง ตอนนี้ตัด `'` ก่อนเทียบ และเทียบกับค่าที่แสดงจริง (`getDisplayValues`)

ส่วนอื่นเหมือนฉบับเดิมทุกประการ (ไม่ได้แก้เรื่องสิทธิ์/รหัสผ่าน)

## วิธีนำไปใช้
1. เปิดโปรเจกต์ Apps Script → เปิดไฟล์โค้ดเดิม → ลบแล้ววางเนื้อหา `Code.gs` นี้ → บันทึก
2. **Deploy → Manage deployments → ไอคอนดินสอ → Version: New version → Deploy** (ใช้ URL เดิม ถ้าเลือก New deployment URL จะเปลี่ยน)
3. ทดสอบ: เปิด `<WebAppURL>?action=getAttendanceReport` ต้องได้ JSON `{"status":"success","data":[...]}`

## หมายเหตุ
- แถวซ้ำที่เกิดขึ้นแล้วในชีต DailyRoutine ไม่ถูกลบอัตโนมัติ ถ้ามีมาก ให้ลบเองหรือใช้ Data → Data cleanup → Remove duplicates (เลือกคอลัมน์วันที่ + รหัสนักเรียน)
- `DRIVE_FOLDER_ID` ในไฟล์คือค่าเดิมของโปรเจกต์
