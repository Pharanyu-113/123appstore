# Draws the Context Diagram, DFD level 1 and flowcharts for chapter 3 in TH SarabunPSK.
# Processes, data stores and flows follow app/api/* and the custom_* RPCs in supabase/.
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "sk", "scripts"))
from thai_fonts import setup_matplotlib
setup_matplotlib(13)
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle, Polygon, Circle

INK = "#1f2933"
LINE = dict(color=INK, lw=1.1)


def canvas(w_in, h_units, w_units=100):
    fig = plt.figure(figsize=(w_in, w_in * h_units / w_units))
    ax = fig.add_axes([0, 0, 1, 1]); ax.set_xlim(0, w_units); ax.set_ylim(0, h_units); ax.axis("off")
    return fig, ax


def text(ax, x, y, s, size=13, bold=False, ha="center", va="center"):
    ax.text(x, y, s, fontsize=size, fontweight="bold" if bold else "normal", ha=ha, va=va,
            color=INK, linespacing=1.0)


def arrow(ax, pts, both=False):
    xs, ys = zip(*pts)
    ax.plot(xs[:-1] + (xs[-1],), ys[:-1] + (ys[-1],), **LINE)
    ax.annotate("", xy=pts[-1], xytext=pts[-2], arrowprops=dict(arrowstyle="-|>", color=INK, lw=1.1,
                                                              mutation_scale=12, shrinkA=0, shrinkB=0))
    if both:
        ax.annotate("", xy=pts[0], xytext=pts[1], arrowprops=dict(arrowstyle="-|>", color=INK, lw=1.1,
                                                                mutation_scale=12, shrinkA=0, shrinkB=0))


def entity(ax, cx, cy, w, h, label, dup=False, size=13):
    ax.add_patch(Rectangle((cx - w / 2, cy - h / 2), w, h, fill=False, ec=INK, lw=1.3))
    if dup:   # duplicated external entity: diagonal mark in the lower-right corner
        ax.plot([cx + w / 2 - 2.2, cx + w / 2], [cy - h / 2, cy - h / 2 + 2.2], **LINE)
    text(ax, cx, cy, label, size, bold=True)


def process(ax, cx, cy, w, h, num, label, size=13):
    ax.add_patch(FancyBboxPatch((cx - w / 2, cy - h / 2), w, h, boxstyle="round,pad=0,rounding_size=2.2",
                                fill=False, ec=INK, lw=1.3))
    ax.plot([cx - w / 2, cx + w / 2], [cy + h / 2 - 3.6, cy + h / 2 - 3.6], **LINE)
    text(ax, cx, cy + h / 2 - 1.8, num, size - 1, bold=True)
    text(ax, cx, cy - 1.6, label, size)


def store(ax, cx, cy, w, h, code, label, size=12):
    x0, x1, y0, y1 = cx - w / 2, cx + w / 2, cy - h / 2, cy + h / 2
    ax.plot([x1, x0, x0, x1], [y1, y1, y0, y0], **LINE)
    ax.plot([x0 + 7, x0 + 7], [y0, y1], **LINE)
    text(ax, x0 + 3.5, cy, code, size, bold=True)
    text(ax, x0 + 7 + (w - 7) / 2, cy, label, size)


def label(ax, x, y, s, size=11, ha="center"):
    ax.text(x, y, s, fontsize=size, ha=ha, va="center", color=INK,
            bbox=dict(fc="white", ec="none", pad=0.6), linespacing=1.0)


def save(fig, name):
    fig.savefig(f"img/{name}.png", dpi=220, facecolor="white", bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)


# ------------------------------------------------------------------ context diagram
def context_diagram():
    fig, ax = canvas(7.2, 78)
    cx, cy = 50, 39
    ax.add_patch(Circle((cx, cy), 13.5, fill=False, ec=INK, lw=1.4))
    text(ax, cx, cy + 6, "0", 14, bold=True)
    text(ax, cx, cy - 1, "ระบบจัดการค้าส่ง\n123พาณิชย์ปลีกส่ง", 14, bold=True)
    boxes = {"OWNER": (12, 66), "ADMIN": (88, 66), "RETAILER": (12, 12), "EMPLOYEE": (88, 12)}
    names = {"OWNER": "เจ้าของร้าน\n(OWNER)", "ADMIN": "ผู้ดูแลคลัง\n(ADMIN)",
             "RETAILER": "ผู้ค้าปลีก\n(RETAILER)", "EMPLOYEE": "พนักงาน\n(EMPLOYEE)"}
    for k, (x, y) in boxes.items():
        entity(ax, x, y, 20, 11, names[k])
    # each entity: one arrow into the system and one arrow out, labelled
    flows = {
        "OWNER": ("ข้อมูลเข้าสู่ระบบ, ข้อมูลสินค้า\nและประเภทสินค้า, จำนวนปรับสต็อก",
                  "ภาพรวมสินค้า สต็อก คำสั่งซื้อ\nใบเสร็จรายวัน และตำแหน่งนำส่ง"),
        "ADMIN": ("ข้อมูลเข้าสู่ระบบ,\nจำนวนปรับสต็อก", "รายการสินค้าและ\nจำนวนคงเหลือ"),
        "RETAILER": ("ข้อมูลสมัครสมาชิก/เข้าสู่ระบบ,\nรายการสั่งซื้อ (กล่อง/แพ็ค)",
                     "รายการสินค้า สถานะคำสั่งซื้อ\nใบเสร็จรายวัน และตำแหน่งนำส่ง"),
        "EMPLOYEE": ("ข้อมูลเข้าสู่ระบบ, สถานะ\nคำสั่งซื้อ, พิกัดตำแหน่ง", "รายการคำสั่งซื้อและ\nข้อมูลร้านค้าผู้รับ"),
    }
    for k, (x, y) in boxes.items():
        sx = 1 if x < cx else -1; sy = 1 if y < cy else -1
        ex, ey = x + sx * 10, y + sy * 0          # box side
        # inbound (entity -> system) runs horizontally then into the circle
        inn = [(ex, y + 2.2 * -sy), (cx - sx * 6.5, y + 2.2 * -sy), (cx - sx * 6.5, cy - sy * 12.0)]
        out = [(cx - sx * 11.5, cy - sy * 6.8), (cx - sx * 11.5, y - 2.2 * -sy), (ex, y - 2.2 * -sy)]
        arrow(ax, inn); arrow(ax, out)
        mid = (ex + cx - sx * 6.5) / 2
        label(ax, mid, y + 2.2 * -sy + (-sy) * 4.0, flows[k][0])
        label(ax, (ex + cx - sx * 11.5) / 2, y - 2.2 * -sy + sy * 6.5, flows[k][1])
    save(fig, "ctx")


# ------------------------------------------------------------------ DFD level 1
def dfd():
    rows = [
        ("1.0", "ยืนยันตัวตน", ["ผู้ใช้งาน\nทุกบทบาท"], "ข้อมูลสมัคร/\nเข้าสู่ระบบ", "session และ\nบทบาท",
         [("D1", "profiles", "rw"), ("D2", "app_credentials", "rw")]),
        ("2.0", "จัดการสินค้าและ\nประเภทสินค้า", ["OWNER"], "ข้อมูลสินค้า\nและประเภท", "ผลการบันทึก",
         [("D3", "product_categories", "w"), ("D4", "products", "w")]),
        ("3.0", "ปรับจำนวนสต็อก", ["OWNER /\nADMIN"], "จำนวนที่ปรับ", "สต็อกคงเหลือ",
         [("D4", "products", "rw")]),
        ("4.0", "สั่งซื้อและ\nออกใบเสร็จรายวัน", ["RETAILER"], "รายการสั่งซื้อ\n(กล่อง/แพ็ค)", "เลขคำสั่งซื้อ\nและใบเสร็จ",
         [("D4", "products", "rw"), ("D5", "orders, order_items", "w"), ("D6", "daily_receipts,\nreceipt_items", "w")]),
        ("5.0", "จัดการสถานะ\nคำสั่งซื้อ", ["EMPLOYEE"], "สถานะใหม่", "คำสั่งซื้อที่\nรับผิดชอบ",
         [("D5", "orders, order_items", "rw")]),
        ("6.0", "บันทึกตำแหน่ง\nการนำส่ง", ["EMPLOYEE"], "พิกัดละติจูด/\nลองจิจูด", "ผลการบันทึก",
         [("D7", "delivery_tracking", "w")]),
        ("7.0", "แสดงข้อมูล\nแดชบอร์ดตามบทบาท", ["ผู้ใช้งาน\nทุกบทบาท"], "คำขอข้อมูล", "ข้อมูลตามสิทธิ์",
         [("D1,\nD3–D7", "ผู้ใช้ สินค้า คำสั่งซื้อ\nใบเสร็จ และตำแหน่ง", "r")]),
    ]
    row_h = 21; top = row_h * len(rows) + 2
    fig, ax = canvas(7.2, top)
    ex, px, sx = 11, 46, 82            # column centres
    for i, (num, name, ents, fin, fout, stores) in enumerate(rows):
        cy = top - 1 - row_h * (i + 0.5)
        entity(ax, ex, cy, 18, 10, ents[0], dup=True, size=12)
        process(ax, px, cy, 22, 15, num, name, size=12)
        arrow(ax, [(ex + 9, cy + 2.4), (px - 11, cy + 2.4)])
        arrow(ax, [(px - 11, cy - 2.4), (ex + 9, cy - 2.4)])
        label(ax, (ex + 9 + px - 11) / 2, cy + 6.4, fin, 10)
        label(ax, (ex + 9 + px - 11) / 2, cy - 6.4, fout, 10)
        n = len(stores); sh = 7.0 if n < 3 else 6.0
        for j, (code, sname, mode) in enumerate(stores):
            sy = cy + (n - 1) * (sh + 0.8) / 2 - j * (sh + 0.8)
            store(ax, sx, sy, 34, sh if "\n" not in code else 9, code, sname, size=10 if n == 3 else 11)
            pts = [(px + 11, cy + (sy - cy) * 0.35), (sx - 17, sy)]
            if mode == "rw":
                arrow(ax, pts, both=True)
            elif mode == "w":
                arrow(ax, pts)
            else:
                arrow(ax, pts[::-1])
    save(fig, "dfd1")


# ------------------------------------------------------------------ flowcharts
def node(ax, kind, x, y, s, w=34, h=7.5, size=12):
    if kind == "term":
        ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h, boxstyle=f"round,pad=0,rounding_size={h / 2}",
                                    fill=False, ec=INK, lw=1.3))
    elif kind == "proc":
        ax.add_patch(Rectangle((x - w / 2, y - h / 2), w, h, fill=False, ec=INK, lw=1.3))
    elif kind == "io":
        k = 2.5
        ax.add_patch(Polygon([(x - w / 2 + k, y + h / 2), (x + w / 2 + k, y + h / 2), (x + w / 2 - k, y - h / 2),
                              (x - w / 2 - k, y - h / 2)], closed=True, fill=False, ec=INK, lw=1.3))
    elif kind == "dec":
        ax.add_patch(Polygon([(x, y + h / 2), (x + w / 2, y), (x, y - h / 2), (x - w / 2, y)],
                             closed=True, fill=False, ec=INK, lw=1.3))
    text(ax, x, y, s, size)


def flowchart(name, height, nodes, edges, notes=()):
    fig, ax = canvas(6.2, height)
    for n in nodes:
        node(ax, *n)
    for pts, lab in edges:
        arrow(ax, pts)
        if lab:
            (x0, y0), (x1, y1) = pts[0], pts[1]
            label(ax, x0 + (2.6 if x1 == x0 else 0), y0 + (0 if x1 == x0 else 1.8), lab, 11,
                  ha="left" if x1 == x0 else "center")
    save(fig, name)


def flowcharts():
    D = 11  # vertical step
    # login ------------------------------------------------------------------
    y = [96 - D * i for i in range(9)]
    flowchart("fc_login", 102, [
        ("term", 50, y[0], "เริ่มต้น", 22),
        ("io", 50, y[1], "กรอกเบอร์โทรศัพท์และรหัสผ่าน", 40),
        ("dec", 50, y[2], "รูปแบบเบอร์\nถูกต้อง?", 30, 9),
        ("proc", 50, y[3], "ค้นหาบัญชีและตรวจสอบรหัสผ่าน\nที่แฮชไว้ใน app_credentials", 44, 9),
        ("dec", 50, y[4], "บัญชีและ\nรหัสผ่านถูกต้อง?", 30, 9),
        ("proc", 50, y[5], "สร้าง session cookie\nแบบ httpOnly ที่ลงลายมือชื่อ", 42, 9),
        ("proc", 50, y[6], "ไปยังแดชบอร์ดตามบทบาท", 40),
        ("term", 50, y[7], "สิ้นสุด", 22),
        ("proc", 86, y[3], "แจ้งข้อผิดพลาด", 22),
    ], [
        ([(50, y[0] - 3.75), (50, y[1] + 3.75)], None),
        ([(50, y[1] - 3.75), (50, y[2] + 4.5)], None),
        ([(50, y[2] - 4.5), (50, y[3] + 4.5)], "ใช่"),
        ([(50, y[3] - 4.5), (50, y[4] + 4.5)], None),
        ([(50, y[4] - 4.5), (50, y[5] + 4.5)], "ใช่"),
        ([(50, y[5] - 4.5), (50, y[6] + 3.75)], None),
        ([(50, y[6] - 3.75), (50, y[7] + 3.75)], None),
        ([(65, y[2]), (86, y[2]), (86, y[3] + 3.75)], "ไม่ใช่"),
        ([(65, y[4]), (86, y[4]), (86, y[3] - 3.75)], "ไม่ใช่"),
        ([(97, y[3]), (99, y[3]), (99, y[1]), (72.5, y[1])], None),
    ])

    # add product -------------------------------------------------------------
    y = [96 - D * i for i in range(9)]
    flowchart("fc_product", 102, [
        ("term", 50, y[0], "เริ่มต้น", 22),
        ("io", 50, y[1], "กรอกชื่อ ประเภท ราคากล่อง ราคาแพ็ค\nแพ็คต่อกล่อง และจำนวนเริ่มต้น", 50, 9),
        ("dec", 50, y[2], "กรอก\nราคาแพ็ค?", 28, 9),
        ("dec", 50, y[3], "กรอกแพ็คต่อ\nกล่องแล้ว?", 28, 9),
        ("proc", 50, y[4], "บันทึกสินค้าลง products\n(stock = กล่อง, stock_pack = 0)", 44, 9),
        ("proc", 50, y[5], "แสดงสินค้าในคลังและหน้าสั่งซื้อ", 44),
        ("term", 50, y[6], "สิ้นสุด", 22),
        ("proc", 14, y[3], "ขายยกกล่อง\nอย่างเดียว", 22, 9),
        ("proc", 86, y[3], "แจ้งให้ระบุ\nแพ็คต่อกล่อง", 22, 9),
    ], [
        ([(50, y[0] - 3.75), (50, y[1] + 4.5)], None),
        ([(50, y[1] - 4.5), (50, y[2] + 4.5)], None),
        ([(50, y[2] - 4.5), (50, y[3] + 4.5)], "ใช่"),
        ([(50, y[3] - 4.5), (50, y[4] + 4.5)], "ใช่"),
        ([(50, y[4] - 4.5), (50, y[5] + 3.75)], None),
        ([(50, y[5] - 3.75), (50, y[6] + 3.75)], None),
        ([(36, y[2]), (14, y[2]), (14, y[3] + 4.5)], "ไม่ใช่"),
        ([(14, y[3] - 4.5), (14, y[4]), (28, y[4])], None),
        ([(64, y[3]), (75, y[3])], "ไม่ใช่"),
        ([(86, y[3] + 4.5), (86, y[1]), (77, y[1])], None),
    ])

    # checkout ----------------------------------------------------------------
    y = [108 - D * i for i in range(10)]
    flowchart("fc_checkout", 114, [
        ("term", 50, y[0], "เริ่มต้น", 22),
        ("io", 50, y[1], "เลือกสินค้าและจำนวนกล่อง/แพ็คลงตะกร้า", 50),
        ("proc", 50, y[2], "ยืนยันคำสั่งซื้อพร้อม request id", 44),
        ("dec", 50, y[3], "request id\nเคยใช้แล้ว?", 28, 9),
        ("proc", 50, y[4], "ล็อกแถวสินค้า แปลงสต็อกและ\nจำนวนที่สั่งเป็นหน่วยแพ็ค", 44, 9),
        ("dec", 50, y[5], "สต็อก\nเพียงพอ?", 28, 9),
        ("proc", 50, y[6], "สร้าง orders และ order_items\nรวมยอดลงใบเสร็จรายวัน", 44, 9),
        ("proc", 50, y[7], "หักสต็อกแล้วแปลงกลับเป็น\nกล่องเต็มและแพ็คที่แกะแล้ว", 44, 9),
        ("term", 50, y[8], "สิ้นสุด", 22),
        ("proc", 86, y[3], "ส่งผลคำสั่งซื้อ\nเดิมกลับ", 22, 9),
        ("proc", 86, y[5], "ยกเลิกทั้งรายการ\nแจ้งสต็อกไม่พอ", 22, 9),
    ], [
        ([(50, y[0] - 3.75), (50, y[1] + 3.75)], None),
        ([(50, y[1] - 3.75), (50, y[2] + 3.75)], None),
        ([(50, y[2] - 3.75), (50, y[3] + 4.5)], None),
        ([(50, y[3] - 4.5), (50, y[4] + 4.5)], "ไม่ใช่"),
        ([(50, y[4] - 4.5), (50, y[5] + 4.5)], None),
        ([(50, y[5] - 4.5), (50, y[6] + 4.5)], "ใช่"),
        ([(50, y[6] - 4.5), (50, y[7] + 4.5)], None),
        ([(50, y[7] - 4.5), (50, y[8] + 3.75)], None),
        ([(64, y[3]), (75, y[3])], "ใช่"),
        ([(64, y[5]), (75, y[5])], "ไม่ใช่"),
        ([(97, y[3]), (99, y[3]), (99, y[8]), (61, y[8])], None),
        ([(97, y[5]), (99, y[5])], None),
    ])

    # delivery status -----------------------------------------------------------
    y = [108 - D * i for i in range(10)]
    flowchart("fc_delivery", 114, [
        ("term", 50, y[0], "เริ่มต้น", 22),
        ("io", 50, y[1], "พนักงานเลือกคำสั่งซื้อและสถานะถัดไป", 50),
        ("dec", 50, y[2], "ยังไม่มีผู้รับ\nหรือเป็นงานของตน?", 32, 9.5),
        ("dec", 50, y[3], "ลำดับสถานะ\nถูกต้อง?", 28, 9),
        ("proc", 50, y[4], "บันทึกสถานะและผู้รับผิดชอบลง orders", 46),
        ("dec", 50, y[5], "สถานะเป็น\nDELIVERING?", 28, 9),
        ("proc", 50, y[6], "ส่งพิกัดเมื่อเคลื่อนที่ 25 เมตร\nหรือครบ 15 วินาที ลง delivery_tracking", 50, 9),
        ("proc", 50, y[7], "ส่งสินค้าสำเร็จ เปลี่ยนเป็น COMPLETED", 46),
        ("term", 50, y[8], "สิ้นสุด", 22),
        ("proc", 87, y[2], "แจ้งว่างานนี้มี\nผู้รับผิดชอบแล้ว", 22, 9),
        ("proc", 87, y[3], "แจ้งว่าข้ามขั้น\nไม่ได้", 22, 9),
    ], [
        ([(50, y[0] - 3.75), (50, y[1] + 3.75)], None),
        ([(50, y[1] - 3.75), (50, y[2] + 4.75)], None),
        ([(50, y[2] - 4.75), (50, y[3] + 4.5)], "ใช่"),
        ([(50, y[3] - 4.5), (50, y[4] + 3.75)], "ใช่"),
        ([(50, y[4] - 3.75), (50, y[5] + 4.5)], None),
        ([(50, y[5] - 4.5), (50, y[6] + 4.5)], "ใช่"),
        ([(50, y[6] - 4.5), (50, y[7] + 3.75)], None),
        ([(50, y[7] - 3.75), (50, y[8] + 3.75)], None),
        ([(66, y[2]), (76, y[2])], "ไม่ใช่"),
        ([(64, y[3]), (76, y[3])], "ไม่ใช่"),
        ([(36, y[5]), (12, y[5]), (12, y[8]), (39, y[8])], "ไม่ใช่"),
        ([(98, y[3]), (99, y[3]), (99, y[8]), (61, y[8])], None),
        ([(98, y[2]), (99, y[2]), (99, y[3])], None),
    ])


if __name__ == "__main__":
    context_diagram(); dfd(); flowcharts()
    print("ok")
