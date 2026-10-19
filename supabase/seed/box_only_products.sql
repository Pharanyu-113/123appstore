-- 123พาณิชย์ปลีกส่ง · ปรับสินค้าขวดและเบียร์ให้ขายยกลังอย่างเดียว
-- ใช้เมื่อรัน 20261016_starter_products.sql ไปแล้วก่อนที่ไฟล์นั้นจะถูกแก้ ถ้ายังไม่เคยรันไฟล์สินค้า ไม่ต้องรันไฟล์นี้
-- รันซ้ำได้: สินค้าที่ปรับแล้วหรือหาชื่อไม่เจอจะไม่ถูกแตะ
-- สินค้าที่มีแพ็คแกะค้างอยู่ (stock_pack > 0) จะถูกข้าม เพื่อไม่ให้แพ็คที่เหลือหายไปจากสต๊อก

begin;

create temporary table box_only_products (
  current_name text not null,
  new_name text not null
) on commit drop;

insert into box_only_products (current_name, new_name) values
  ('น้ำปลาแท้ ทิพรส 700 มล.', 'น้ำปลาแท้ ทิพรส 700 มล.'),
  ('ซีอิ๊วขาว เด็กสมบูรณ์ สูตร 1 700 มล.', 'ซีอิ๊วขาว เด็กสมบูรณ์ สูตร 1 700 มล.'),
  ('ซอสหอยนางรม แม่ครัว 600 มล.', 'ซอสหอยนางรม แม่ครัว 600 มล.'),
  ('ซอสปรุงรส ภูเขาทอง ฝาเขียว 600 มล.', 'ซอสปรุงรส ภูเขาทอง ฝาเขียว 600 มล.'),
  ('ซอสพริก ศรีราชาพานิช 450 มล.', 'ซอสพริก ศรีราชาพานิช 450 มล.'),
  ('ซอสมะเขือเทศ โรซ่า 300 ก.', 'ซอสมะเขือเทศ โรซ่า 300 ก.'),
  ('น้ำมันปาล์ม มรกต 1 ลิตร', 'น้ำมันปาล์ม มรกต 1 ลิตร'),
  ('น้ำมันถั่วเหลือง องุ่น 1 ลิตร', 'น้ำมันถั่วเหลือง องุ่น 1 ลิตร'),
  ('น้ำส้มสายชูกลั่น 5% 700 มล.', 'น้ำส้มสายชูกลั่น 5% 700 มล.'),
  ('เบียร์ลีโอ กระป๋อง 490 มล. (แพ็ค 6)', 'เบียร์ลีโอ กระป๋อง 490 มล. (ลัง 24 กระป๋อง)'),
  ('เบียร์ช้าง คลาสสิค กระป๋อง 490 มล. (แพ็ค 6)', 'เบียร์ช้าง คลาสสิค กระป๋อง 490 มล. (ลัง 24 กระป๋อง)'),
  ('เบียร์สิงห์ กระป๋อง 490 มล. (แพ็ค 6)', 'เบียร์สิงห์ กระป๋อง 490 มล. (ลัง 24 กระป๋อง)');

-- ลบราคาแพ็คออก ปุ่มเลือกแพ็คจะหายจากหน้าสั่งซื้อ และการตัดสต๊อกจะนับเป็นกล่องอย่างเดียว
-- ใบเสร็จและออเดอร์ย้อนหลังเก็บชื่อและราคาเดิมไว้ในตัวเอง จึงไม่เปลี่ยนตาม
update public.products as product
set name = target.new_name,
    price_pack = null,
    packs_per_box = null
from box_only_products as target
where product.name = target.current_name
  and product.stock_pack = 0;

-- แสดงผลแต่ละรายการ: box_only = ปรับแล้ว, skipped_loose_packs = มีแพ็คแกะค้าง, not_found = ไม่พบชื่อนี้
select
  target.new_name,
  case
    when product.id is null then 'not_found'
    when product.price_pack is null then 'box_only'
    else 'skipped_loose_packs'
  end as result,
  product.stock,
  product.stock_pack
from box_only_products as target
left join public.products as product
  on product.name in (target.current_name, target.new_name)
order by target.new_name;

commit;
