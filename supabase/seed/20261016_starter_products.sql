-- 123พาณิชย์ปลีกส่ง · สินค้าเริ่มต้น 47 รายการ
-- รันหลังจาก migrations/20261016_packs_per_box_stock.sql (ต้องมีคอลัมน์ packs_per_box และ stock_pack แล้ว)
--
-- ราคาเป็นราคาส่งโดยประมาณ ปัดเป็นเลขกลม (บาท) ไม่ใช่ราคาตลาด ณ วันที่รัน ตรวจและแก้ก่อนรันได้
-- price_box  = ราคาต่อลัง/กล่องใหญ่
-- price_pack = ราคาต่อหน่วยย่อยที่แกะขาย (แพ็ค / ขวด / ถุง)
-- packs_per_box = จำนวนหน่วยย่อยในลังตามขนาดบรรจุที่พบทั่วไป ควรเทียบกับลังจริงอีกครั้ง
-- stock = จำนวนลังเริ่มต้น (ตั้งไว้ 20 ทุกรายการ)
--
-- รันซ้ำได้: สินค้าที่มีชื่อซ้ำอยู่แล้วจะถูกข้าม
-- ถ้าหาประเภทสินค้าไม่เจอ ไฟล์จะหยุดและไม่เพิ่มอะไรเลย

begin;

create temporary table seed_products (
  category_group text not null,
  name text not null,
  price_box numeric(12, 2) not null,
  price_pack numeric(12, 2),
  packs_per_box integer,
  stock integer not null
) on commit drop;

insert into seed_products (category_group, name, price_box, price_pack, packs_per_box, stock) values
  -- ขนม: ราคาในชื่อคือราคาขายปลีกต่อซอง/ชิ้น, 1 แพ็คมี 12 ซอง
  ('snack', 'เลย์ คลาสสิค รสมันฝรั่งแท้ 5 บาท',          290, 50, 6, 20),
  ('snack', 'เลย์ ร็อค รสโนริสาหร่าย 10 บาท',            590, 100, 6, 20),
  ('snack', 'ทาโร่ ปลาสวรรค์ รสบาร์บีคิว 5 บาท',          290, 50, 6, 20),
  ('snack', 'ตะวัน ข้าวเกรียบกุ้ง 5 บาท',                290, 50, 6, 20),
  ('snack', 'โปเต้ รสสาหร่าย 10 บาท',                   590, 100, 6, 20),
  ('snack', 'ยูโร่ คัสตาร์ดเค้ก 5 บาท',                 570, 50, 12, 20),
  ('snack', 'ฟันโอ คุกกี้แซนวิช รสวานิลลา 5 บาท',         570, 50, 12, 20),
  ('snack', 'เถ้าแก่น้อย สาหร่ายทอด รสคลาสสิค 10 บาท',    590, 100, 6, 20),
  ('snack', 'นมแท่ง รสนม 5 บาท',                       570, 50, 12, 20),
  ('snack', 'นมแท่ง รสช็อกโกแลต 5 บาท',                 570, 50, 12, 20),

  -- เครื่องครัว: 1 แพ็ค = 1 ขวด / ถุง / กล่อง
  ('kitchen', 'น้ำปลาแท้ ทิพรส 700 มล.',                 400, 35, 12, 20),
  ('kitchen', 'ซีอิ๊วขาว เด็กสมบูรณ์ สูตร 1 700 มล.',       520, 45, 12, 20),
  ('kitchen', 'ซอสหอยนางรม แม่ครัว 600 มล.',             580, 50, 12, 20),
  ('kitchen', 'ซอสปรุงรส ภูเขาทอง ฝาเขียว 600 มล.',        460, 40, 12, 20),
  ('kitchen', 'ซอสพริก ศรีราชาพานิช 450 มล.',             580, 50, 12, 20),
  ('kitchen', 'ซอสมะเขือเทศ โรซ่า 300 ก.',               340, 30, 12, 20),
  ('kitchen', 'น้ำมันปาล์ม มรกต 1 ลิตร',                  580, 50, 12, 20),
  ('kitchen', 'น้ำมันถั่วเหลือง องุ่น 1 ลิตร',               700, 60, 12, 20),
  ('kitchen', 'น้ำตาลทรายขาว มิตรผล 1 กก.',              720, 30, 25, 20),
  -- ผงชูรส 4 ขนาด และรสดี 3 ขนาด เรียงจากเล็กสุด ขนาดซองเล็กขายเป็นแพ็ค 12 ซอง
  ('kitchen', 'ผงชูรส อายิโนะโมะโต๊ะ 75 ก. (แพ็ค 12 ซอง)',  1250, 130, 10, 20),
  ('kitchen', 'ผงชูรส อายิโนะโมะโต๊ะ 165 ก.',             1150, 25, 48, 20),
  ('kitchen', 'ผงชูรส อายิโนะโมะโต๊ะ 250 ก.',             1350, 35, 40, 20),
  ('kitchen', 'ผงชูรส อายิโนะโมะโต๊ะ 500 ก.',             1400, 60, 24, 20),
  ('kitchen', 'ผงปรุงรส รสดี รสหมู 70 ก. (แพ็ค 12 ซอง)',   1050, 110, 10, 20),
  ('kitchen', 'ผงปรุงรส รสดี รสหมู 165 ก.',              950, 20, 48, 20),
  ('kitchen', 'ผงปรุงรส รสดี รสหมู 400 ก.',              1050, 45, 24, 20),
  ('kitchen', 'เกลือป่น ปรุงทิพย์ 1 กก.',                  280, 12, 24, 20),
  ('kitchen', 'น้ำส้มสายชูกลั่น 5% 700 มล.',              170, 15, 12, 20),
  ('kitchen', 'กะทิ ชาวเกาะ UHT 250 มล.',               700, 20, 36, 20),
  ('kitchen', 'แป้งทอดกรอบ โกกิ 500 ก.',                680, 35, 20, 20),

  -- เครื่องดื่ม: 1 แพ็ค = แพ็คหดที่ขายในลัง
  ('drink', 'โคคา-โคล่า กระป๋อง 325 มล. (แพ็ค 6)',       290, 75, 4, 20),
  ('drink', 'แฟนต้า น้ำแดง กระป๋อง 325 มล. (แพ็ค 6)',     290, 75, 4, 20),
  ('drink', 'โออิชิ กรีนที น้ำผึ้งผสมมะนาว 380 มล. (แพ็ค 6)', 330, 85, 4, 20),
  ('drink', 'เอ็ม-150 150 มล. (แพ็ค 10)',               470, 95, 5, 20),
  ('drink', 'เบอร์ดี้ กาแฟกระป๋อง โรบัสต้า 180 มล. (แพ็ค 6)', 340, 70, 5, 20),
  ('drink', 'น้ำดื่ม สิงห์ 600 มล. (แพ็ค 12)',            55, null, null, 20),

  -- เบียร์: ลังละ 24 กระป๋อง = 4 แพ็ค แพ็คละ 6 กระป๋อง
  ('beer', 'เบียร์ลีโอ กระป๋อง 490 มล. (แพ็ค 6)',          1050, 270, 4, 20),
  ('beer', 'เบียร์ช้าง คลาสสิค กระป๋อง 490 มล. (แพ็ค 6)',    1050, 270, 4, 20),
  ('beer', 'เบียร์สิงห์ กระป๋อง 490 มล. (แพ็ค 6)',          1180, 300, 4, 20),

  -- นม: ลังละ 48 กล่อง = 12 แพ็ค แพ็คละ 4 กล่อง (แลคตาซอย 60 กล่อง = 10 แพ็ค แพ็คละ 6)
  ('milk', 'นมไทย-เดนมาร์ค UHT รสจืด 200 มล. (แพ็ค 4)',     590, 50, 12, 20),
  ('milk', 'นมไทย-เดนมาร์ค UHT รสช็อกโกแลต 200 มล. (แพ็ค 4)', 590, 50, 12, 20),
  ('milk', 'นมโฟร์โมสต์ UHT รสจืด 180 มล. (แพ็ค 4)',        530, 45, 12, 20),
  ('milk', 'นมโฟร์โมสต์ UHT รสหวาน 180 มล. (แพ็ค 4)',       530, 45, 12, 20),
  ('milk', 'นมโฟร์โมสต์ UHT รสช็อกโกแลต 180 มล. (แพ็ค 4)',   530, 45, 12, 20),
  ('milk', 'นมเมจิ UHT รสจืด 200 มล. (แพ็ค 4)',            590, 50, 12, 20),
  ('milk', 'ดัชมิลล์ นมเปรี้ยว รสสตรอเบอร์รี่ 180 มล. (แพ็ค 4)', 530, 45, 12, 20),
  ('milk', 'แลคตาซอย นมถั่วเหลือง รสหวาน 125 มล. (แพ็ค 6)',  390, 40, 10, 20);

-- จับคู่กลุ่มกับประเภทที่มีอยู่แล้ว เทียบด้วยชื่อที่ normalize (ตัวพิมพ์เล็ก ไม่มีช่องว่างเกิน)
create temporary table seed_category_map on commit drop as
select
  seed_group.category_group,
  (
    select category.id
    from public.product_categories as category
    where category.is_active
      and case seed_group.category_group
        when 'snack' then category.normalized_name like '%ขนม%'
        when 'kitchen' then category.normalized_name like '%เครื่องครัว%'
        when 'drink' then category.normalized_name like '%เครื่องดื่ม%'
          and category.normalized_name not like '%แอลกอฮอล%'
          and category.normalized_name not like '%เบียร์%'
        when 'beer' then category.normalized_name like '%แอลกอฮอล%'
          or category.normalized_name like '%เบียร์%'
        when 'milk' then category.normalized_name like 'นม%'
          or category.normalized_name like '%ผลิตภัณฑ์นม%'
      end
    order by char_length(category.name), category.created_at
    limit 1
  ) as category_id
from (values ('snack'), ('kitchen'), ('drink'), ('beer'), ('milk')) as seed_group(category_group);

do $$
declare
  v_missing text;
begin
  select string_agg(
    case category_group
      when 'snack' then 'ขนม'
      when 'kitchen' then 'เครื่องครัว'
      when 'drink' then 'เครื่องดื่ม'
      when 'beer' then 'เครื่องดื่มแอลกอฮอล์/เบียร์'
      when 'milk' then 'นม'
    end, ', ')
  into v_missing
  from seed_category_map
  where category_id is null;

  if v_missing is not null then
    raise exception 'ไม่พบประเภทสินค้าที่เปิดใช้งาน: % (ตรวจชื่อประเภทในระบบ หรือแก้เงื่อนไขในไฟล์นี้)', v_missing;
  end if;
end;
$$;

insert into public.products (name, category, category_id, price_box, price_pack, packs_per_box, stock, stock_pack, created_by)
select
  seed.name,
  category.name,
  category.id,
  seed.price_box,
  seed.price_pack,
  seed.packs_per_box,
  seed.stock,
  0,
  (select profile.id from public.profiles as profile where profile.role = 'OWNER' order by profile.created_at limit 1)
from seed_products as seed
join seed_category_map as map on map.category_group = seed.category_group
join public.product_categories as category on category.id = map.category_id
where not exists (
  select 1 from public.products as existing
  where lower(regexp_replace(trim(existing.name), '\s+', ' ', 'g')) = lower(regexp_replace(trim(seed.name), '\s+', ' ', 'g'))
);

-- แสดงผลว่าแต่ละกลุ่มลงประเภทไหน
select map.category_group, category.name as category_name,
  (select count(*) from seed_products as seed where seed.category_group = map.category_group) as seed_rows
from seed_category_map as map
join public.product_categories as category on category.id = map.category_id
order by map.category_group;

commit;
