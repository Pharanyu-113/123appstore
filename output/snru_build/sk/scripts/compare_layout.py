#!/usr/bin/env python3
"""Compare the rendered layout of two thesis PDFs role by role.

  python compare_layout.py reference.pdf candidate.pdf

For every element type (cover lines, chapter headings, x.y / x.y.z headings,
body first lines and wrapped lines, captions, TOC lines, bibliography
entries, page numbers ...) it measures position (x, y), font size and bold
in both PDFs and reports the differences. Used to check that a newly
generated thesis looks exactly like an approved one.
"""
import re, sys, statistics as st
import fitz

TOL = 2.5   # points


def lines(page):
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            sp = [s for s in l["spans"] if "".join(c["c"] for c in s["chars"]).strip()]
            if not sp: continue
            chars = [c for s in sp for c in s["chars"] if c["c"].strip()]
            tx = "".join("".join(c["c"] for c in s["chars"]) for s in l["spans"]).strip()
            out.append(dict(x0=min(c["bbox"][0] for c in chars), x1=max(c["bbox"][2] for c in chars),
                            y=round(l["bbox"][1], 1), size=round(st.median(s["size"] for s in sp)),
                            bold=[("Bold" in s["font"]) for s in sp], text=tx))
    return sorted(out, key=lambda r: (r["y"], r["x0"]))


def profile(pdf):
    d = fitz.open(pdf)
    P = {}
    pages = [lines(p) for p in d]
    def put(k, v):
        P.setdefault(k, []).append(v)
    W = d[0].rect.width
    # covers: first 4 pages
    for ci, name in enumerate(("cover_outer", "cover_inner_th", "cover_en", "approval")):
        if ci >= len(pages): break
        img = [r for r in d[ci].get_image_rects(d[ci].get_images()[0][0])] if d[ci].get_images() else []
        if img: put(f"{name}.logo_y", round(img[0].y0, 1)); put(f"{name}.logo_h", round(img[0].height, 1))
        for j, r in enumerate(pages[ci][:22]):
            put(f"{name}.line{j:02d}", (r["y"], round((r["x0"] + r["x1"]) / 2), r["size"], all(r["bold"])))
    for pi, ls in enumerate(pages):
        txt = "\n".join(r["text"] for r in ls)
        # page numbers
        for r in ls:
            if re.fullmatch(r"\d+", r["text"]) and r["y"] < 80: put("pagenum_arabic", (round(r["x1"]), r["y"], r["size"]))
            if re.fullmatch(r"[IVX]+", r["text"]) and r["y"] > 700: put("pagenum_roman", (round((r["x0"] + r["x1"]) / 2), r["y"], r["size"]))
        body = [r for r in ls if not re.fullmatch(r"[0-9IVX]+", r["text"])]
        if not body: continue
        first = body[0]
        # abstract page
        if body[0]["text"].startswith("ชื่อเรื่อง"):
            put("abstract.label_x", body[0]["x0"])
            for r in body:
                if r["text"] == "บทคัดย่อ": put("abstract.title", (r["y"], r["size"], all(r["bold"])))
                if r["text"].startswith("คำสำคัญ"): put("abstract.keywords_bold", r["bold"][0])
        if first["text"] == "กิตติกรรมประกาศ":
            put("ack.title", (first["y"], first["size"], all(first["bold"])))
            put("ack.first_para_x", body[1]["x0"]); put("ack.first_para_y", body[1]["y"])
            put("ack.sign_right", round(body[-1]["x1"]))
        # TOC / lists
        if first["text"] in ("สารบัญ", "สารบัญ (ต่อ)", "สารบัญตาราง", "สารบัญภาพ", "สารบัญภาพ (ต่อ)", "สารบัญตาราง (ต่อ)"):
            key = "toc" if first["text"].startswith("สารบัญ") and not first["text"].startswith(("สารบัญตาราง", "สารบัญภาพ")) else "list"
            put(f"{key}.title", (first["y"], first["size"], all(first["bold"])))
            for r in body[1:]:
                if r["text"].endswith("หน้า") or r["text"] == "หน้า":
                    put(f"{key}.header_y", r["y"]); put(f"{key}.header_right", round(r["x1"]))
                elif r["text"] in ("ตารางที่", "ภาพที่"):
                    put(f"{key}.header_left", r["x0"])
                elif re.match(r"^บทที่ \d", r["text"]):
                    put("toc.chapter_x", r["x0"]); put("toc.chapter_bold", all(r["bold"]))
                    put("toc.chapter_has_dots", "...." in r["text"])
                elif re.match(r"^\d+\.\d+ ", r["text"]):
                    put(f"{key}.item_x", r["x0"]); put(f"{key}.item_right", round(r["x1"]))
                    put(f"{key}.item_bold", any(r["bold"]))
                elif r["text"].startswith(("บทคัดย่อ", "กิตติกรรม", "บรรณานุกรม", "ประวัติ")):
                    put("toc.main_x", r["x0"]); put("toc.main_bold_pattern", tuple(r["bold"]))
            ys = [r["y"] for r in body]
            put(f"{key}.pitch", round(st.median([b - a for a, b in zip(ys, ys[1:]) if 15 < b - a < 25]), 1) if len(ys) > 3 else None)
        # chapter first pages
        if re.fullmatch(r"บทที่ \d+", first["text"]):
            put("chapter.no", (first["y"], round((first["x0"] + first["x1"]) / 2), first["size"], all(first["bold"])))
            t2 = body[1]
            put("chapter.title", (t2["y"], round((t2["x0"] + t2["x1"]) / 2), t2["size"], all(t2["bold"])))
            put("chapter.next_y", body[2]["y"])
        for k, r in enumerate(body):
            t = r["text"]
            nxt = body[k + 1] if k + 1 < len(body) else None
            prv = body[k - 1] if k > 0 else None
            if re.match(r"^\d+\.\d+ \S", t) and all(r["bold"]) and r["size"] == 18:
                put("h2", (r["x0"], r["size"]))
                if prv: put("h2.gap_above", round(r["y"] - prv["y"], 1))
                if nxt: put("h2.gap_below", round(nxt["y"] - r["y"], 1))
                if nxt and not nxt["bold"][0]: put("h2.body_first_x", nxt["x0"])
            elif re.match(r"^\d+\.\d+\.\d+ \S", t) and r["bold"][0] and r["size"] == 16 and len(t) < 90:
                put("h3", (r["x0"], r["size"], all(r["bold"])))
                if nxt and not nxt["bold"][0] and not re.match(r"^\d", nxt["text"]): put("h3.body_first_x", nxt["x0"])
            elif re.match(r"^(ภาพที่|ตารางที่) (\d+|[ก-ฮ])\.\d+ ", t) and r["bold"][0]:
                kind = "fig" if t.startswith("ภาพที่") else "tab"
                put(f"{kind}.caption", (round((r["x0"] + r["x1"]) / 2) if kind == "fig" else r["x0"], r["size"], tuple(r["bold"][:2])))
            elif t.startswith("ที่มา :"):
                put("source." + ("fig" if prv and prv["text"].startswith("ภาพที่") else "tab"), (round((r["x0"] + r["x1"]) / 2), r["size"], any(r["bold"])))
        # body text metrics on chapter pages
        if any(re.match(r"^\d+\.\d+ ", r["text"]) for r in body):
            ys = [r["y"] for r in body if r["size"] == 16]
            d16 = [round(b - a, 1) for a, b in zip(ys, ys[1:]) if 18 < b - a < 23]
            if d16: put("body.pitch", st.median(d16))
            for r in body:
                if r["size"] == 16 and not any(r["bold"]) and abs(r["x0"] - 108) < 1 and r["x1"] > 500:
                    put("body.wrap_x", r["x0"]); put("body.right", round(r["x1"]))
        # bibliography
        if first["text"].startswith("บรรณานุกรม"):
            put("bib.title", (first["y"], first["size"], all(first["bold"])))
            if len(body) > 1: put("bib.first_entry_gap", round(body[1]["y"] - first["y"], 1))
            for r in body[1:]:
                if re.match(r"^[A-Zก-ฮเแโใไ]", r["text"]) and r["x0"] < 112: put("bib.entry_x", r["x0"])
                elif r["x0"] > 140: put("bib.hang_x", r["x0"])
        if first["text"].startswith("ภาคผนวก"):
            put("appendix.title", (first["y"] if len(body) < 3 else "top", first["size"], all(first["bold"])))
        if first["text"] == "ประวัติผู้จัดทำ":
            put("bio.title", (first["y"], first["size"], all(first["bold"])))
            for r in body[1:]:
                if all(r["bold"][:1]) and r["x0"] < 112: put("bio.label_x", r["x0"])
    return P


def summarize(v):
    if not v: return None
    if all(isinstance(x, (int, float)) for x in v):
        return round(st.median(v), 1)
    vals = sorted(set(map(str, v)))
    return vals if len(vals) > 1 else vals[0]


def close(a, b):
    try:
        return abs(float(a) - float(b)) <= TOL
    except Exception:
        pass
    ta = re.findall(r"-?\d+\.?\d*|True|False", str(a)); tb = re.findall(r"-?\d+\.?\d*|True|False", str(b))
    if len(ta) == len(tb) and ta:
        for x, y in zip(ta, tb):
            if x in ("True", "False") or y in ("True", "False"):
                if x != y: return False
            elif abs(float(x) - float(y)) > TOL:
                return False
        return True
    return str(a) == str(b)


def main():
    ref, cand = profile(sys.argv[1]), profile(sys.argv[2])
    keys = sorted(set(ref) | set(cand))
    bad = 0
    for k in keys:
        a, b = summarize(ref.get(k)), summarize(cand.get(k))
        if a is None or b is None:
            if a is None and b is None: continue
            print(f"?    {k:28s} ref={a} new={b}"); continue
        ok = close(a, b)
        bad += not ok
        print(("OK   " if ok else "DIFF ") + f"{k:28s} ref={a}  new={b}")
    print(f"\n{bad} differences")


if __name__ == "__main__":
    main()
