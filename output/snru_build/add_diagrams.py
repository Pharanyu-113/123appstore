# Adds Context Diagram, DFD level 1, flowcharts and the per-table data dictionary to chapter 3
# (run after make_content.py and add_eval.py). Images come from diagrams.py; columns follow
# supabase/schema.sql + custom-auth-migration.sql + migrations/* (final state, 2026-10-16).
import io, json

c = json.load(io.open("content.json", encoding="utf-8"))
ch3 = next(x for x in c["chapters"] if x["number"] == 3)["blocks"]

def P(t): return {"type": "p", "text": t}
def H3(t): return {"type": "h3", "text": t}
def FIG(img, cap, w=14.6): return {"type": "figure", "image": f"img/{img}.png", "caption": cap, "source": None, "width_cm": w}
def TABLE(cap, header, rows, widths, align=None):
    # each of these tables fits on one page: keep it whole so the caption never sits alone
    t = {"type": "table", "caption": cap, "header": header, "rows": rows, "widths": widths,
         "keep_together": True, "source": None}
    if align: t["align"] = align
    return t

# figure / table numbers after insertion (chapter 3):
#   ภาพ 3.1 สถาปัตยกรรม, 3.2 บริบท, 3.3 DFD, 3.4 การไหลคำสั่งซื้อ, 3.5-3.8 ผังงาน, 3.9 แบบจำลองข้อมูล, 3.10 สิทธิ์
#   ตาราง 3.1 กระแสข้อมูลแผนภาพบริบท, 3.2 กระบวนการ DFD, 3.3 พจนานุกรมข้อมูลหลัก, 3.4-3.12 รายตาราง, 3.13 สิทธิ์
context = [
    H3("3.1.1 แผนภาพบริบท (Context Diagram)"),
    P("แผนภาพบริบทแสดงขอบเขตของระบบจัดการค้าส่ง 123พาณิชย์ปลีกส่งในภาพรวม โดยมองทั้งระบบเป็นกระบวนการเดียว และแสดงข้อมูลที่ไหลเข้าและออกระหว่างระบบกับผู้ใช้งานภายนอก 4 กลุ่ม ได้แก่ เจ้าของร้าน (OWNER) ผู้ดูแลคลัง (ADMIN) พนักงาน (EMPLOYEE) และผู้ค้าปลีก (RETAILER) ดังภาพที่ 3.2"),
    FIG("ctx", "แผนภาพบริบทของระบบจัดการค้าส่ง 123พาณิชย์ปลีกส่ง"),
    P("จากภาพที่ 3.2 สามารถอธิบายเส้นทางการไหลของข้อมูลระหว่างผู้ใช้งานแต่ละกลุ่มกับระบบได้ดังตารางที่ 3.1"),
    TABLE("คำอธิบายเส้นทางการไหลของข้อมูลในแผนภาพบริบท", ["ผู้ใช้งาน", "ข้อมูลเข้าสู่ระบบ", "ข้อมูลออกจากระบบ"], [
        ["เจ้าของร้าน (OWNER)", "ข้อมูลเข้าสู่ระบบ ข้อมูลสินค้าและประเภทสินค้า และจำนวนปรับสต็อก",
         "ภาพรวมสินค้า สต็อก คำสั่งซื้อ ใบเสร็จรายวัน และตำแหน่งการนำส่ง"],
        ["ผู้ดูแลคลัง (ADMIN)", "ข้อมูลเข้าสู่ระบบ และจำนวนปรับสต็อก", "รายการสินค้าและจำนวนคงเหลือ"],
        ["พนักงาน (EMPLOYEE)", "ข้อมูลเข้าสู่ระบบ สถานะคำสั่งซื้อ และพิกัดตำแหน่งระหว่างนำส่ง",
         "รายการคำสั่งซื้อและข้อมูลร้านค้าผู้รับ"],
        ["ผู้ค้าปลีก (RETAILER)", "ข้อมูลสมัครสมาชิกหรือเข้าสู่ระบบ และรายการสั่งซื้อแบบกล่องหรือแพ็ค",
         "รายการสินค้า สถานะคำสั่งซื้อ ใบเสร็จรายวัน และตำแหน่งการนำส่ง"],
    ], [2006, 3150, 3150], ["left", "left", "left"]),
    H3("3.1.2 แผนภาพกระแสข้อมูลระดับที่ 1 (Data Flow Diagram)"),
    P("แผนภาพกระแสข้อมูลระดับที่ 1 แยกระบบออกเป็น 7 กระบวนการย่อย และแสดงความสัมพันธ์กับแหล่งเก็บข้อมูล 7 แหล่ง (D1–D7) ซึ่งตรงกับตารางในฐานข้อมูล ผู้ใช้งานภายนอกที่ปรากฏมากกว่าหนึ่งครั้งมีเส้นทแยงที่มุมกรอบเพื่อแสดงว่าเป็นผู้ใช้งานเดียวกัน ดังภาพที่ 3.3"),
    FIG("dfd1", "แผนภาพกระแสข้อมูลระดับที่ 1 (Data Flow Diagram)", 13.0),
    P("จากภาพที่ 3.3 สามารถอธิบายการทำงานของแต่ละกระบวนการได้ดังตารางที่ 3.2"),
    TABLE("คำอธิบายกระบวนการในแผนภาพกระแสข้อมูลระดับที่ 1", ["กระบวนการ", "ผู้ใช้งาน", "การทำงานและแหล่งข้อมูล"], [
        ["1.0 ยืนยันตัวตน", "ทุกบทบาท", "รับเบอร์โทรศัพท์และรหัสผ่าน ตรวจสอบกับ D1 profiles และ D2 app_credentials แล้วสร้าง session ตามบทบาท ผู้ค้าปลีกสมัครบัญชีใหม่ได้ที่กระบวนการนี้"],
        ["2.0 จัดการสินค้าและประเภทสินค้า", "OWNER", "เพิ่มหรือลบสินค้า และเพิ่ม เปลี่ยนชื่อ เปิดหรือปิดการใช้งานประเภทสินค้า บันทึกลง D3 product_categories และ D4 products"],
        ["3.0 ปรับจำนวนสต็อก", "OWNER, ADMIN", "เพิ่มหรือลดจำนวนกล่องคงเหลือใน D4 products โดยจำนวนคงเหลือไม่ต่ำกว่า 0"],
        ["4.0 สั่งซื้อและออกใบเสร็จรายวัน", "RETAILER", "ตรวจสต็อกใน D4 สร้างคำสั่งซื้อใน D5 รวมยอดลงใบเสร็จรายวันใน D6 และหักสต็อกใน D4 ภายในธุรกรรมเดียว"],
        ["5.0 จัดการสถานะคำสั่งซื้อ", "EMPLOYEE", "เปลี่ยนสถานะคำสั่งซื้อใน D5 ตามลำดับ PENDING, PACKED, DELIVERING และ COMPLETED"],
        ["6.0 บันทึกตำแหน่งการนำส่ง", "EMPLOYEE", "บันทึกพิกัดล่าสุดของคำสั่งซื้อที่อยู่ระหว่างนำส่งลง D7 delivery_tracking"],
        ["7.0 แสดงข้อมูลแดชบอร์ดตามบทบาท", "ทุกบทบาท", "อ่านข้อมูลจาก D1 และ D3–D7 แล้วส่งเฉพาะข้อมูลที่บทบาทนั้นมีสิทธิ์เห็น"],
    ], [2600, 1400, 4306], ["left", "left", "left"]),
]

flowcharts = [
    H3("3.2.1 ผังงานระบบ (Flowchart)"),
    P("ผังงานระบบแสดงลำดับการทำงานและเงื่อนไขของกระบวนการหลัก 4 กระบวนการ ได้แก่ การเข้าสู่ระบบ การเพิ่มสินค้า การยืนยันคำสั่งซื้อ และการเปลี่ยนสถานะการนำส่ง โดยผังงานการเข้าสู่ระบบแสดงดังภาพที่ 3.5"),
    FIG("fc_login", "ผังงานการเข้าสู่ระบบ", 11.0),
    P("การเพิ่มสินค้าของเจ้าของร้านมีเงื่อนไขเรื่องจำนวนแพ็คต่อกล่อง เมื่อกำหนดราคาแพ็คต้องระบุจำนวนแพ็คต่อกล่องด้วย ดังภาพที่ 3.6"),
    FIG("fc_product", "ผังงานการเพิ่มสินค้า", 11.0),
    P("การยืนยันคำสั่งซื้อของผู้ค้าปลีกตรวจสอบคำขอซ้ำ ตรวจสอบสต็อก สร้างคำสั่งซื้อ รวมยอดลงใบเสร็จรายวัน และหักสต็อกในธุรกรรมเดียว ดังภาพที่ 3.7"),
    FIG("fc_checkout", "ผังงานการยืนยันคำสั่งซื้อและหักสต็อก", 11.0),
    P("การเปลี่ยนสถานะคำสั่งซื้อของพนักงานตรวจสอบผู้รับผิดชอบและลำดับสถานะ และส่งตำแหน่งระหว่างนำส่ง ดังภาพที่ 3.8"),
    FIG("fc_delivery", "ผังงานการเปลี่ยนสถานะและการนำส่ง", 11.0),
]

PK, FK = "PK", "FK"
def dd(name, desc, cols):
    return TABLE(f"{name} ({desc})", ["ชื่อฟิลด์", "ชนิดข้อมูล", "คีย์", "คำอธิบาย"], cols,
                 [2400, 1900, 700, 3306], ["left", "left", "center", "left"])
dictionary = [
    H3("3.3.1 พจนานุกรมข้อมูล (Data Dictionary)"),
    P("พจนานุกรมข้อมูลแสดงรายละเอียดของแต่ละตารางในฐานข้อมูล Supabase PostgreSQL ได้แก่ ชื่อฟิลด์ ชนิดข้อมูล คีย์ และคำอธิบาย โดย PK หมายถึงคีย์หลัก (Primary Key) และ FK หมายถึงคีย์นอก (Foreign Key) ดังตารางที่ 3.4 ถึงตารางที่ 3.12"),
    dd("profiles", "ข้อมูลผู้ใช้และร้านค้า", [
        ["id", "uuid", PK, "รหัสผู้ใช้"],
        ["phone", "text", "", "เบอร์โทรศัพท์ที่ใช้เข้าสู่ระบบ ห้ามซ้ำ"],
        ["shop_name", "text", "", "ชื่อร้านหรือชื่อผู้ใช้"],
        ["role", "app_role", "", "บทบาท OWNER, ADMIN, EMPLOYEE หรือ RETAILER"],
        ["auth_version", "integer", "", "รุ่นของข้อมูลยืนยันตัวตน ใช้ยกเลิก session เดิม"],
        ["created_at", "timestamptz", "", "วันเวลาที่สร้างบัญชี"]]),
    dd("app_credentials", "ข้อมูลรหัสผ่าน", [
        ["profile_id", "uuid", "PK, FK", "รหัสผู้ใช้ อ้างอิง profiles.id"],
        ["password_salt", "text", "", "ค่า salt สำหรับแฮชรหัสผ่าน"],
        ["password_hash", "text", "", "รหัสผ่านที่ผ่านการแฮชแล้ว ไม่เก็บรหัสผ่านจริง"],
        ["created_at", "timestamptz", "", "วันเวลาที่สร้าง"],
        ["updated_at", "timestamptz", "", "วันเวลาที่แก้ไขล่าสุด"]]),
    dd("product_categories", "ข้อมูลประเภทสินค้า", [
        ["id", "uuid", PK, "รหัสประเภทสินค้า"],
        ["name", "text", "", "ชื่อประเภทสินค้า ยาว 1–80 ตัวอักษร"],
        ["normalized_name", "text", "", "ชื่อที่ปรับรูปแบบแล้ว ใช้ป้องกันชื่อซ้ำ"],
        ["is_active", "boolean", "", "สถานะเปิดหรือปิดการใช้งาน"],
        ["created_by", "uuid", FK, "ผู้สร้าง อ้างอิง profiles.id"],
        ["created_at", "timestamptz", "", "วันเวลาที่สร้าง"],
        ["updated_at", "timestamptz", "", "วันเวลาที่แก้ไขล่าสุด"]]),
    dd("products", "ข้อมูลสินค้า", [
        ["id", "uuid", PK, "รหัสสินค้า"],
        ["name", "text", "", "ชื่อสินค้า"],
        ["category", "text", "", "ชื่อประเภทสินค้า ณ เวลาที่บันทึก"],
        ["category_id", "uuid", FK, "รหัสประเภทสินค้า อ้างอิง product_categories.id"],
        ["price_box", "numeric(12,2)", "", "ราคาต่อกล่อง"],
        ["price_pack", "numeric(12,2)", "", "ราคาต่อแพ็ค ว่างได้เมื่อขายยกกล่องอย่างเดียว"],
        ["packs_per_box", "integer", "", "จำนวนแพ็คต่อกล่อง ต้องระบุเมื่อมีราคาแพ็ค"],
        ["stock", "integer", "", "จำนวนกล่องเต็มคงเหลือ ไม่ต่ำกว่า 0"],
        ["stock_pack", "integer", "", "จำนวนแพ็คที่แกะจากกล่องแล้ว น้อยกว่าแพ็คต่อกล่อง"],
        ["created_at", "timestamptz", "", "วันเวลาที่เพิ่มสินค้า"],
        ["created_by", "uuid", FK, "ผู้เพิ่มสินค้า อ้างอิง profiles.id"]]),
    dd("orders", "ข้อมูลคำสั่งซื้อ", [
        ["id", "uuid", PK, "รหัสคำสั่งซื้อ"],
        ["retailer_id", "uuid", FK, "ผู้ค้าปลีกที่สั่งซื้อ อ้างอิง profiles.id"],
        ["receipt_id", "uuid", FK, "ใบเสร็จรายวัน อ้างอิง daily_receipts.id"],
        ["assigned_employee_id", "uuid", FK, "พนักงานผู้รับผิดชอบ อ้างอิง profiles.id"],
        ["status", "order_status", "", "สถานะ PENDING, PACKED, DELIVERING หรือ COMPLETED"],
        ["total_amount", "numeric(12,2)", "", "ยอดรวมของคำสั่งซื้อ"],
        ["stock_deducted_at", "timestamptz", "", "วันเวลาที่หักสต็อก"],
        ["checkout_request_id", "uuid", "", "รหัสคำขอยืนยัน ใช้ป้องกันคำสั่งซื้อซ้ำ"],
        ["created_at", "timestamptz", "", "วันเวลาที่สั่งซื้อ"],
        ["updated_at", "timestamptz", "", "วันเวลาที่แก้ไขล่าสุด"]]),
    dd("order_items", "ข้อมูลรายการสินค้าในคำสั่งซื้อ", [
        ["id", "uuid", PK, "รหัสรายการ"],
        ["order_id", "uuid", FK, "คำสั่งซื้อ อ้างอิง orders.id"],
        ["product_id", "uuid", FK, "สินค้า อ้างอิง products.id ว่างได้เมื่อสินค้าถูกลบ"],
        ["product_name", "text", "", "ชื่อสินค้า ณ เวลาที่สั่งซื้อ"],
        ["quantity_box", "integer", "", "จำนวนกล่องที่สั่ง"],
        ["quantity_pack", "integer", "", "จำนวนแพ็คที่สั่ง"],
        ["unit_price_box", "numeric(12,2)", "", "ราคาต่อกล่อง ณ เวลาที่สั่งซื้อ"],
        ["unit_price_pack", "numeric(12,2)", "", "ราคาต่อแพ็ค ณ เวลาที่สั่งซื้อ"],
        ["line_total", "numeric(12,2)", "", "ยอดรวมของรายการ"]]),
    dd("daily_receipts", "ข้อมูลใบเสร็จรายวัน", [
        ["id", "uuid", PK, "รหัสใบเสร็จ"],
        ["retailer_id", "uuid", FK, "ผู้ค้าปลีก อ้างอิง profiles.id"],
        ["receipt_date", "date", "", "วันที่ของใบเสร็จตามเวลาประเทศไทย ร้านละหนึ่งใบต่อวัน"],
        ["total_amount", "numeric(12,2)", "", "ยอดรวมทั้งวัน"],
        ["created_at", "timestamptz", "", "วันเวลาที่สร้าง"],
        ["updated_at", "timestamptz", "", "วันเวลาที่แก้ไขล่าสุด"]]),
    dd("receipt_items", "ข้อมูลรายการสินค้าในใบเสร็จรายวัน", [
        ["id", "uuid", PK, "รหัสรายการ"],
        ["receipt_id", "uuid", FK, "ใบเสร็จ อ้างอิง daily_receipts.id"],
        ["product_id", "uuid", FK, "สินค้า อ้างอิง products.id ว่างได้เมื่อสินค้าถูกลบ"],
        ["product_name", "text", "", "ชื่อสินค้า ณ เวลาที่สั่งซื้อ"],
        ["category", "text", "", "ประเภทสินค้า ณ เวลาที่สั่งซื้อ"],
        ["quantity_box", "integer", "", "จำนวนกล่องรวม"],
        ["quantity_pack", "integer", "", "จำนวนแพ็ครวม"],
        ["unit_price_box", "numeric(12,2)", "", "ราคาต่อกล่อง"],
        ["unit_price_pack", "numeric(12,2)", "", "ราคาต่อแพ็ค"],
        ["line_total", "numeric(12,2)", "", "ยอดรวมของรายการ"]]),
    dd("delivery_tracking", "ข้อมูลตำแหน่งการนำส่ง", [
        ["order_id", "uuid", "PK, FK", "คำสั่งซื้อ อ้างอิง orders.id"],
        ["employee_id", "uuid", FK, "พนักงานที่ส่งพิกัด อ้างอิง profiles.id"],
        ["latitude", "double precision", "", "ละติจูด อยู่ระหว่าง -90 ถึง 90"],
        ["longitude", "double precision", "", "ลองจิจูด อยู่ระหว่าง -180 ถึง 180"],
        ["updated_at", "timestamptz", "", "วันเวลาที่บันทึกพิกัดล่าสุด"]]),
]

def after(pred, items):
    i = next(k for k, b in enumerate(ch3) if pred(b))
    ch3[i + 1:i + 1] = items

# 3.1: after the paragraph that explains the architecture figure
after(lambda b: b["type"] == "p" and b["text"].startswith("ผู้ใช้เข้าถึงส่วนติดต่อบนเว็บ"), context)
# 3.2: after the order-flow paragraph
after(lambda b: b["type"] == "p" and b["text"].startswith("ผู้ค้าปลีกเลือกสินค้าพร้อมจำนวน"), flowcharts)
# 3.3: after the summary data-dictionary table
after(lambda b: b["type"] == "table" and b["caption"].startswith("พจนานุกรมข้อมูลหลัก"), dictionary)

io.open("content.json", "w", encoding="utf-8").write(json.dumps(c, ensure_ascii=False, indent=1))
print("ok")
