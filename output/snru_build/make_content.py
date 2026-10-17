# Builds content.json for the SNRU builder from the existing report text.
# Body text is copied unchanged from รายงานโครงงาน_123พาณิชย์ปลีกส่ง_ปรับปรุงระบบแพ็คต่อกล่อง.docx.
# Items the old report never had (advisor, chair, exam date, acknowledgement, bio details)
# are left as visible placeholders "(...)" for the author to fill in Word.
import json, re, io

dump = io.open("dump.txt", encoding="utf-8").read().splitlines()

def text_at(line_no):
    """Text of dump line N (1-based), joining a wrapped continuation line if any."""
    line = dump[line_no - 1]
    return line.split("|", 6)[6] if line.startswith("P|") else line

def P(n): return {"type": "p", "text": text_at(n)}
def H2(n): return {"type": "h2", "text": text_at(n)}
def H3(n): return {"type": "h3", "text": text_at(n)}
def numbered(lines):
    return {"type": "list", "style": "number",
            "items": [re.sub(r"^\d+\.\s*", "", text_at(n)) for n in lines]}
def dashed(lines):
    return {"type": "list", "style": "dash",
            "items": [re.sub(r"^-\s*", "", text_at(n)) for n in lines]}
def item(n): return {"type": "item", "label": text_at(n), "text": ""}
def caption(n):
    return re.sub(r"^(ตารางที่|ตาราง|ภาพที่)\s*[\dก-ฮ]+\.\d+\s*", "", text_at(n))
def rows_of(a, b):
    out = []
    for n in range(a, b + 1):
        line = dump[n - 1]
        if line.startswith("  ROW| "):
            out.append([c.strip() for c in line[7:].split(" || ")])
        elif out:  # cell text that wrapped onto its own dump line ("OWNER\nเจ้าของร้าน")
            out[-1][0] += " " + line.split(" || ")[0].strip()
            out[-1][1:] = [c.strip() for c in line.split(" || ")[1:]] or out[-1][1:]
    return out
def table(cap_line, a, b, widths, align=None):
    rows = rows_of(a, b)
    t = {"type": "table", "caption": caption(cap_line), "header": rows[0], "rows": rows[1:],
         "widths": widths, "source": None}
    if align: t["align"] = align
    return t

# Gantt chart: arrows sat in these month columns (1 = ก.พ. ... 9 = ต.ค.)
gantt_months = ["ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค."]
gantt_rows = [
    ("1. ศึกษาและรวบรวมข้อมูล", [1]),
    ("2. วิเคราะห์และออกแบบระบบ", [2, 3]),
    ("3. ออกแบบหน้าจอ", [3]),
    ("4. พัฒนาระบบ", [3, 4]),
    ("5. ทดสอบระบบและแก้ไข", [4, 5]),
    ("6. ประเมินประสิทธิภาพระบบโดยผู้เชี่ยวชาญ", [5, 6]),
    ("7. ทำเอกสาร", [7, 8]),
    ("8. วิเคราะห์ข้อมูลและสรุปผลจัดทำรายงาน", [8, 9]),
]
gantt = {
    "type": "table", "caption": caption(112), "grid": True,
    "header_rows": [["กิจกรรม", "ระยะเวลาในการดำเนินงาน (ปี พ.ศ. 2569)"] + [""] * 8,
                    [""] + gantt_months],
    "rows": [[name] + ["<->" if (m + 1) in months else "" for m in range(9)] for name, months in gantt_rows],
    "merge": [[0, 0, 1, 0], [0, 1, 0, 9]],
    "widths": [2456] + [650] * 9,
    "align": ["left"] + ["center"] * 9,
    "source": None,
}

def fig(img_no, cap_line, width_cm=14.6):
    return {"type": "figure", "image": f"img/fig{img_no:02d}.png", "caption": caption(cap_line),
            "source": None, "width_cm": width_cm}

chapters = [
    {"number": 1, "title": "บทนำ", "blocks": [
        H2(51), P(52), P(53),
        H2(54), numbered(range(55, 59)),
        H2(59), H3(60),
        item(61), dashed(range(62, 66)),
        item(66), dashed(range(67, 71)),
        item(71), dashed(range(72, 74)),
        item(74), dashed(range(75, 78)),
        item(78), dashed(range(79, 82)),
        item(82), dashed(range(83, 87)),
        table(88, 90, 98, [2200, 6106]),
        H2(101), table(102, 104, 108, [1800, 3000, 3506]),
        H2(111), gantt,
        H2(126), numbered(range(127, 131)),
    ]},
    {"number": 2, "title": "เอกสารและงานที่เกี่ยวข้อง", "blocks": [
        P(132),
        H2(133), H3(134), P(135), P(136), H3(137), P(138), P(139), H3(140), P(141),
        H2(142), P(143),
        H3(144), P(145), H3(146), P(147), H3(148), P(149), H3(150), P(151),
        H3(152), P(153), H3(154), P(155), H3(156), P(157),
        table(158, 160, 168, [2800, 5506]),
        H2(170), H3(171), P(172), H3(173), P(174),
        H2(175), H3(176), P(177), H3(178), P(179),
        H2(180), H3(181), P(182), H3(183), P(184), H3(185), P(186), P(187),
        H2(188), H3(189), P(190), H3(191), P(192),
        H2(193), P(194),
    ]},
    {"number": 3, "title": "วิธีดำเนินการวิจัย", "blocks": [
        H2(196), P(197), fig(16, 199), P(200),
        H2(201), fig(17, 203), P(204),
        H2(205), fig(18, 207), table(208, 210, 216, [2800, 5506]),
        H2(219), fig(19, 221), table(222, 224, 230, [5000, 3306]),
        H2(233), P(234),
    ]},
    {"number": 4, "title": "ผลการดำเนินงาน", "blocks": [
        H2(236), P(237),
        H2(238), P(239),
        H2(240), table(241, 243, 250, [1900, 3000, 3406]),
        H2(253), P(254),
        H2(255), P(256),
    ]},
    {"number": 5, "title": "สรุปผลและข้อเสนอแนะ", "blocks": [
        H2(258), P(259), P(260),
        H2(261), P(262),
        H2(263), numbered(range(264, 268)),
    ]},
]

def appendix_role(head_line, first, last):
    return [{"type": "p", "text": f"**{text_at(head_line)}**", "first": 0}, numbered(range(first, last + 1))]

appendices = [
    {"letter": "ก", "title": "คู่มือย่อการใช้งานตามบทบาท", "blocks":
        appendix_role(280, 281, 283) + appendix_role(284, 285, 287) +
        appendix_role(288, 289, 291) + appendix_role(292, 293, 295)},
    {"letter": "ข", "title": "แบบบันทึกผลการทดสอบ", "blocks": [
        P(298), table(299, 301, 305, [2600, 3300, 2406])]},
]

refs = [text_at(n) for n in range(269, 278)]
content = {
    "meta": {
        "title_th": ["การพัฒนาเว็บแอปพลิเคชันระบบจัดการค้าส่ง", "123พาณิชย์ปลีกส่ง"],
        "title_en": ["Development of a Web Application for", "123 Phanit Retail and Wholesale Management"],
        "title_en_upper": ["DEVELOPMENT OF A WEB APPLICATION FOR", "123 PHANIT RETAIL AND WHOLESALE", "MANAGEMENT"],
        "author_prefix_th": "",
        "author_th": "ภรัณยู โยธาพล",
        "author_en": "PHARANYU YOTHAPHON",
        "student_id": "66102122113",
        "degree_th": "วิทยาศาสตรบัณฑิต",
        "program_th": "เทคโนโลยีคอมพิวเตอร์และดิจิทัล",
        "program_en": "COMPUTER AND DIGITAL TECHNOLOGY",
        "university_th": "มหาวิทยาลัยราชภัฏสกลนคร",
        "faculty_th": "คณะวิทยาศาสตร์และเทคโนโลยี",
        "university_en": "SAKON NAKHON RAJABHAT UNIVERSITY",
        "year_be": "2569", "year_ce": "2026",
        "exam_date_line": "เมื่อวันที่      เดือน               พ.ศ. 2569",
        "advisor": "อาจารย์ ดร.ชายแดน มิ่งเมือง",
        "examiners": ["อาจารย์ ดร.สุทิศา ซองเหล็กนอก", "อาจารย์ปิยวรรณ  โถปาสอน"],
        "chair": "ผู้ช่วยศาสตราจารย์ ดร.ธิดารัตน์ เหลืองรุ่งเรือง",
    },
    "abstract": {
        "paragraphs": [text_at(n) for n in (16, 17, 18)],
        "keywords": [k.strip() for k in text_at(19).split(":", 1)[1].split(",")],
    },
    # style 3 chosen by the author on 2026-10-17: opens with the project's own challenges
    "acknowledgement": [
        "การพัฒนาเว็บแอปพลิเคชันระบบจัดการค้าส่ง 123พาณิชย์ปลีกส่ง เป็นโอกาสให้ผู้จัดทำได้นำความรู้ด้านการพัฒนาเว็บและฐานข้อมูลมาแก้ปัญหาการจัดการสินค้า คำสั่งซื้อ และการนำส่งของร้านค้าจริง ระหว่างการพัฒนามีโจทย์ที่ท้าทายหลายประการ เช่น การตัดสต็อกแบบกล่องและแพ็คให้ถูกต้อง และการป้องกันคำสั่งซื้อซ้ำ ซึ่งผ่านมาได้ด้วยคำแนะนำของ อาจารย์ ดร.ชายแดน มิ่งเมือง อาจารย์ที่ปรึกษาโครงงาน ที่ได้ให้คำปรึกษา ตรวจสอบ และแก้ไขข้อบกพร่องต่าง ๆ ตลอดระยะเวลาการดำเนินโครงงาน ผู้จัดทำขอขอบพระคุณเป็นอย่างสูงมา ณ โอกาสนี้",
        "ขอขอบพระคุณ อาจารย์ ดร.สุทิศา ซองเหล็กนอก และอาจารย์ปิยวรรณ โถปาสอน กรรมการสอบโครงงาน รวมถึงผู้ช่วยศาสตราจารย์ ดร.ธิดารัตน์ เหลืองรุ่งเรือง ประธานหลักสูตรเทคโนโลยีคอมพิวเตอร์และดิจิทัล ที่ได้ให้ข้อเสนอแนะอันเป็นประโยชน์ต่อการปรับปรุงโครงงานให้มีความสมบูรณ์ยิ่งขึ้น ขอขอบพระคุณคณาจารย์สาขาวิชาเทคโนโลยีคอมพิวเตอร์และดิจิทัล คณะวิทยาศาสตร์และเทคโนโลยี มหาวิทยาลัยราชภัฏสกลนคร ที่ได้ถ่ายทอดความรู้และประสบการณ์ซึ่งเป็นพื้นฐานสำคัญในการพัฒนาโครงงานครั้งนี้",
        "ขอขอบพระคุณผู้เชี่ยวชาญทั้ง 5 ท่านที่สละเวลาประเมินประสิทธิภาพของเว็บแอปพลิเคชัน และขอขอบคุณร้าน 123พาณิชย์ปลีกส่ง ที่ให้ข้อมูลการดำเนินงานซึ่งใช้ในการวิเคราะห์และออกแบบระบบ",
        "สุดท้ายนี้ ผู้จัดทำขอกราบขอบพระคุณครอบครัว รวมถึงเพื่อน ๆ ที่ให้กำลังใจและการสนับสนุนตลอดมา หากโครงงานฉบับนี้มีข้อผิดพลาดประการใด ผู้จัดทำขอน้อมรับคำแนะนำเพื่อนำไปปรับปรุงต่อไป",
    ],
    "chapters": chapters,
    "references": {"thai": [r for r in refs if re.search(r"[ก-๙]", r.split(".")[0])],
                   "english": [r for r in refs if not re.search(r"[ก-๙]", r.split(".")[0])]},
    "appendices": appendices,
    "bio": {"rows": [["ผู้จัดทำ", "ภรัณยู โยธาพล"],
                     ["วันเดือนปีเกิด", "(กรอกข้อมูล)"],
                     ["ที่อยู่", "(กรอกข้อมูล)"],
                     ["การศึกษา", "วิทยาศาสตรบัณฑิต (เทคโนโลยีคอมพิวเตอร์และดิจิทัล)"],
                     ["", "มหาวิทยาลัยราชภัฏสกลนคร"],
                     ["", "(กรอกประวัติการศึกษาระดับมัธยม)"],
                     ["คติประจำใจ", "(กรอกข้อมูล)"],
                     ["E-mail", "(กรอกข้อมูล)"],
                     ["เบอร์โทรศัพท์", "(กรอกข้อมูล)"]]},
}
# TH SarabunPSK has no arrow glyph and Word substituted Arial, so write the status order the
# way chapter 2 already does ("PENDING, PACKED, DELIVERING และ COMPLETED").
def no_arrows(obj):
    if isinstance(obj, str):
        if "→" in obj:
            parts = [x.strip() for x in obj.split("→")]
            head, first = parts[0].rsplit(" ", 1) if " " in parts[0] else ("", parts[0])
            last_word, tail = (parts[-1].split(" ", 1) + [""])[:2]
            seq = [first] + parts[1:-1] + [last_word]
            obj = (head + " " if head else "") + ", ".join(seq[:-1]) + " และ " + seq[-1] + (" " + tail if tail else "")
        return obj
    if isinstance(obj, list): return [no_arrows(x) for x in obj]
    if isinstance(obj, dict): return {k: no_arrows(v) for k, v in obj.items()}
    return obj
content = no_arrows(content)
io.open("content.json", "w", encoding="utf-8").write(json.dumps(content, ensure_ascii=False, indent=1))
print("ok")
