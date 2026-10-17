"""Replace a generated TOC field with the report's explicit style hierarchy."""

from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
FIELD = ' TOC \\h \\z \\t "Chapter Title,1,Heading 1,2,Heading 2,3" '


def main() -> None:
    path = Path("output") / "รายงานโครงงาน_ระบบจัดการร้าน123พาณิชย์ปลีกส่ง_ฉบับปรับปรุง_วิธีการดำเนินงานเดิม.docx"
    temp_dir = Path(tempfile.mkdtemp(prefix="toc_style_patch_"))
    try:
        with zipfile.ZipFile(path) as source:
            source.extractall(temp_dir)

        document_xml = temp_dir / "word" / "document.xml"
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(document_xml), parser)
        fields = tree.xpath("//w:instrText[contains(., 'TOC')]", namespaces=NS)
        if len(fields) != 1:
            raise RuntimeError(f"Expected exactly one TOC field, found {len(fields)}")
        fields[0].text = FIELD
        tree.write(str(document_xml), encoding="UTF-8", xml_declaration=True, standalone=True)

        # Ask Word to refresh dynamic fields (including the page numbers) when
        # the document is opened in a full Word client.
        settings_xml = temp_dir / "word" / "settings.xml"
        settings_tree = etree.parse(str(settings_xml), parser)
        update_fields = settings_tree.xpath("//w:updateFields", namespaces=NS)
        if update_fields:
            update_field = update_fields[0]
        else:
            update_field = etree.SubElement(settings_tree.getroot(), f"{{{W_NS}}}updateFields")
        update_field.set(f"{{{W_NS}}}val", "true")
        settings_tree.write(str(settings_xml), encoding="UTF-8", xml_declaration=True, standalone=True)

        staged = temp_dir / "report_with_toc.docx"
        with zipfile.ZipFile(staged, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for item in temp_dir.rglob("*"):
                if item.is_file() and item != staged:
                    target.write(item, item.relative_to(temp_dir).as_posix())
        shutil.copyfile(staged, path)
        print("TOC field now maps Chapter Title, Heading 1, and Heading 2.")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
