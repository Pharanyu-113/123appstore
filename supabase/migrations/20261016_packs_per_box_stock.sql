-- 123พาณิชย์ปลีกส่ง · ระบบแพ็คต่อกล่อง
-- รันหลังจาก schema.sql, custom-auth-migration.sql, 20260906_daily_receipts_categories_checkout_stock.sql
-- และ 20260930_product_category_management.sql
--
-- products.stock      = จำนวนกล่องเต็ม
-- products.stock_pack = แพ็คที่แกะกล่องแล้ว (น้อยกว่า packs_per_box เสมอ)
-- เมื่อขายแพ็ค ระบบจะแกะกล่องให้อัตโนมัติ เช่น 1 กล่องมี 5 แพ็ค เหลือ 3 กล่อง สั่ง 3 แพ็ค -> เหลือ 2 กล่อง 2 แพ็ค
-- สินค้าเดิมที่ยังไม่ได้ตั้ง packs_per_box จะตัดแบบเดิม (1 กล่อง หรือ 1 แพ็ค = 1 หน่วยของ stock)

begin;

alter table public.products
  add column if not exists packs_per_box integer,
  add column if not exists stock_pack integer not null default 0;

alter table public.products drop constraint if exists products_packs_per_box_check;
alter table public.products
  add constraint products_packs_per_box_check
  check (packs_per_box is null or (packs_per_box > 0 and price_pack is not null));

alter table public.products drop constraint if exists products_stock_pack_check;
alter table public.products
  add constraint products_stock_pack_check
  check (stock_pack >= 0 and (packs_per_box is null and stock_pack = 0 or stock_pack < packs_per_box));

-- ตัดสต๊อกสินค้า 1 รายการ: ต้องเรียกภายใน transaction ของ RPC ที่ล็อกแถวสินค้าไว้แล้ว
create or replace function public.custom_deduct_product_stock(
  p_product_id uuid,
  p_boxes integer,
  p_packs integer
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v_product public.products;
  v_available integer;
  v_needed integer;
  v_remaining integer;
begin
  select *
  into v_product
  from public.products as product
  where product.id = p_product_id
  for update;

  if not found then
    raise exception 'พบสินค้าที่ไม่มีอยู่ในระบบ';
  end if;

  if v_product.packs_per_box is null then
    v_needed := p_boxes + p_packs;
    if v_product.stock < v_needed then
      raise exception 'สินค้า "%" คงเหลือไม่เพียงพอ (เหลือ %, ต้องใช้ %)', v_product.name, v_product.stock, v_needed;
    end if;

    update public.products as product
    set stock = product.stock - v_needed
    where product.id = p_product_id;
    return;
  end if;

  v_available := v_product.stock * v_product.packs_per_box + v_product.stock_pack;
  v_needed := p_boxes * v_product.packs_per_box + p_packs;
  if v_available < v_needed then
    raise exception 'สินค้า "%" คงเหลือไม่เพียงพอ (เหลือ % กล่อง % แพ็ค)', v_product.name, v_product.stock, v_product.stock_pack;
  end if;

  v_remaining := v_available - v_needed;
  update public.products as product
  set stock = v_remaining / v_product.packs_per_box,
      stock_pack = v_remaining % v_product.packs_per_box
  where product.id = p_product_id;
end;
$$;

-- เหมือนเวอร์ชันใน 20260906 ทุกอย่าง ยกเว้นการตรวจและตัดสต๊อกที่ย้ายไปอยู่ใน custom_deduct_product_stock
create or replace function public.custom_create_retailer_order(
  p_retailer_id uuid,
  p_items jsonb,
  p_request_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_receipt_id uuid;
  v_order_id uuid;
  v_existing_receipt_id uuid;
  v_receipt_date date := (now() at time zone 'Asia/Bangkok')::date;
  v_item record;
  v_product_name text;
  v_category text;
  v_price_box numeric(12, 2);
  v_price_pack numeric(12, 2);
  v_line_total numeric(12, 2);
  v_total numeric(12, 2) := 0;
  v_existing_receipt_item_id uuid;
begin
  if not exists (
    select 1 from public.profiles where id = p_retailer_id and role = 'RETAILER'
  ) then
    raise exception 'เฉพาะผู้ค้าปลีกเท่านั้นที่สั่งซื้อได้';
  end if;

  if jsonb_typeof(p_items) <> 'array' or jsonb_array_length(p_items) = 0 then
    raise exception 'กรุณาเลือกสินค้าอย่างน้อย 1 รายการ';
  end if;
  if p_request_id is null then
    raise exception 'รหัสยืนยันคำสั่งซื้อไม่ถูกต้อง';
  end if;

  -- ป้องกัน retry รหัสเดิมที่คร่อมเที่ยงคืน: ต้องรอคำขอเดิมจบก่อนจึงตรวจผลลัพธ์
  perform pg_advisory_xact_lock(hashtext(p_retailer_id::text), hashtext(p_request_id::text));

  -- Retry ด้วย request เดิมจะได้ผลลัพธ์เดิมโดยไม่ตัดสต๊อกหรือเพิ่มยอดซ้ำ
  select order_row.id, order_row.receipt_id
  into v_order_id, v_existing_receipt_id
  from public.orders as order_row
  where order_row.retailer_id = p_retailer_id
    and order_row.checkout_request_id = p_request_id;
  if found then
    return jsonb_build_object('order_id', v_order_id, 'receipt_id', v_existing_receipt_id);
  end if;

  -- UNIQUE(retailer_id, receipt_date) ทำให้ร้านหนึ่งมีใบเสร็จเดียวต่อวัน และ DO UPDATE ล็อกแถวไว้ทั้ง transaction
  insert into public.daily_receipts (retailer_id, receipt_date, total_amount)
  values (p_retailer_id, v_receipt_date, 0)
  on conflict (retailer_id, receipt_date) do update
  set updated_at = now()
  returning id into v_receipt_id;

  -- อีกคำขอเดียวกันอาจเริ่มพร้อมกันและรอ receipt lock อยู่ จึงตรวจซ้ำหลังได้ lock
  select order_row.id, order_row.receipt_id
  into v_order_id, v_existing_receipt_id
  from public.orders as order_row
  where order_row.retailer_id = p_retailer_id
    and order_row.checkout_request_id = p_request_id;
  if found then
    return jsonb_build_object('order_id', v_order_id, 'receipt_id', v_existing_receipt_id);
  end if;

  -- รวมสินค้าที่ซ้ำใน payload ก่อน จากนั้นล็อกตาม product_id เพื่อป้องกันสต๊อกติดลบระหว่างหลายร้านสั่งพร้อมกัน
  for v_item in
    select
      (entry.value ->> 'product_id')::uuid as product_id,
      sum(greatest(0, coalesce(nullif(entry.value ->> 'quantity_box', '')::integer, 0)))::integer as quantity_box,
      sum(greatest(0, coalesce(nullif(entry.value ->> 'quantity_pack', '')::integer, 0)))::integer as quantity_pack
    from jsonb_array_elements(p_items) as entry(value)
    group by (entry.value ->> 'product_id')::uuid
    order by (entry.value ->> 'product_id')::uuid
  loop
    if v_item.quantity_box = 0 and v_item.quantity_pack = 0 then
      continue;
    end if;

    select product.name, product.price_box, product.price_pack
    into v_product_name, v_price_box, v_price_pack
    from public.products as product
    where product.id = v_item.product_id
    for update;

    if not found then
      raise exception 'พบสินค้าที่ไม่มีอยู่ในระบบ';
    end if;
    if v_item.quantity_pack > 0 and v_price_pack is null then
      raise exception 'สินค้าบางรายการไม่มีราคาต่อแพ็ค';
    end if;

    v_total := v_total + (v_item.quantity_box * v_price_box) + (v_item.quantity_pack * coalesce(v_price_pack, 0));
  end loop;

  if v_total <= 0 then
    raise exception 'กรุณาระบุจำนวนสินค้าที่ถูกต้อง';
  end if;

  -- สร้างงานจัดส่งของการยืนยันครั้งนี้ โดยเชื่อมกับใบเสร็จที่รวมยอดทั้งวัน
  insert into public.orders (retailer_id, receipt_id, total_amount, stock_deducted_at, checkout_request_id)
  values (p_retailer_id, v_receipt_id, v_total, now(), p_request_id)
  returning id into v_order_id;

  for v_item in
    select
      (entry.value ->> 'product_id')::uuid as product_id,
      sum(greatest(0, coalesce(nullif(entry.value ->> 'quantity_box', '')::integer, 0)))::integer as quantity_box,
      sum(greatest(0, coalesce(nullif(entry.value ->> 'quantity_pack', '')::integer, 0)))::integer as quantity_pack
    from jsonb_array_elements(p_items) as entry(value)
    group by (entry.value ->> 'product_id')::uuid
    order by (entry.value ->> 'product_id')::uuid
  loop
    if v_item.quantity_box = 0 and v_item.quantity_pack = 0 then
      continue;
    end if;

    select product.name, product.category, product.price_box, product.price_pack
    into v_product_name, v_category, v_price_box, v_price_pack
    from public.products as product
    where product.id = v_item.product_id;

    v_line_total := (v_item.quantity_box * v_price_box) + (v_item.quantity_pack * coalesce(v_price_pack, 0));

    insert into public.order_items (
      order_id, product_id, product_name, quantity_box, quantity_pack,
      unit_price_box, unit_price_pack, line_total
    ) values (
      v_order_id, v_item.product_id, v_product_name, v_item.quantity_box, v_item.quantity_pack,
      v_price_box, v_price_pack, v_line_total
    );

    -- แยกราคาที่ไม่เท่ากันเป็นคนละบรรทัด เพื่อให้ใบเสร็จย้อนหลังไม่บิดเบือนเมื่อเจ้าของร้านเปลี่ยนราคา
    select receipt_item.id
    into v_existing_receipt_item_id
    from public.receipt_items as receipt_item
    where receipt_item.receipt_id = v_receipt_id
      and receipt_item.product_id = v_item.product_id
      and receipt_item.unit_price_box = v_price_box
      and receipt_item.unit_price_pack is not distinct from v_price_pack
    order by receipt_item.id
    limit 1
    for update;

    if found then
      update public.receipt_items
      set quantity_box = quantity_box + v_item.quantity_box,
          quantity_pack = quantity_pack + v_item.quantity_pack,
          line_total = line_total + v_line_total
      where id = v_existing_receipt_item_id;
    else
      insert into public.receipt_items (
        receipt_id, product_id, product_name, category, quantity_box, quantity_pack,
        unit_price_box, unit_price_pack, line_total
      ) values (
        v_receipt_id, v_item.product_id, v_product_name, v_category, v_item.quantity_box, v_item.quantity_pack,
        v_price_box, v_price_pack, v_line_total
      );
    end if;

    -- ตัดสต๊อกทันทีเมื่อผู้ค้าปลีกยืนยัน ถ้าไม่พอจะ raise และย้อนกลับทั้งคำสั่งซื้อ
    perform public.custom_deduct_product_stock(v_item.product_id, v_item.quantity_box, v_item.quantity_pack);
  end loop;

  update public.daily_receipts
  set total_amount = total_amount + v_total
  where id = v_receipt_id;

  return jsonb_build_object('order_id', v_order_id, 'receipt_id', v_receipt_id);
end;
$$;

-- พนักงานเปลี่ยนสถานะจัดส่งอย่างเดียว; block ตัดสต๊อกนี้มีไว้เฉพาะออเดอร์เก่าที่ยังไม่ถูกตัดก่อน migration 20260906
create or replace function public.custom_update_delivery_status(
  p_employee_id uuid,
  p_order_id uuid,
  p_status public.order_status
)
returns public.orders
language plpgsql
security definer
set search_path = public
as $$
declare
  v_order public.orders;
  v_item record;
begin
  if not exists (
    select 1 from public.profiles where id = p_employee_id and role = 'EMPLOYEE'
  ) then
    raise exception 'เฉพาะพนักงานเท่านั้นที่อัปเดตสถานะจัดส่งได้';
  end if;
  if p_status not in ('PACKED', 'DELIVERING', 'COMPLETED') then
    raise exception 'ไม่สามารถใช้สถานะนี้ได้';
  end if;

  select *
  into v_order
  from public.orders
  where id = p_order_id
    and (assigned_employee_id is null or assigned_employee_id = p_employee_id)
  for update;

  if not found then
    raise exception 'ไม่พบคำสั่งซื้อ หรือคำสั่งซื้อนี้ถูกพนักงานคนอื่นรับงานแล้ว';
  end if;
  if (v_order.status = 'PENDING' and p_status <> 'PACKED')
    or (v_order.status = 'PACKED' and p_status <> 'DELIVERING')
    or (v_order.status = 'DELIVERING' and p_status <> 'COMPLETED')
  then
    raise exception 'ไม่สามารถเปลี่ยนสถานะข้ามขั้น หรือทำรายการเดิมซ้ำได้';
  end if;

  -- Compatibility only: orders created before 20260906 had no checkout-time stock deduction.
  if p_status = 'DELIVERING' and v_order.stock_deducted_at is null then
    for v_item in
      select oi.product_id,
        sum(oi.quantity_box)::integer as quantity_box,
        sum(oi.quantity_pack)::integer as quantity_pack
      from public.order_items as oi
      where oi.order_id = p_order_id
      group by oi.product_id
      order by oi.product_id
    loop
      perform public.custom_deduct_product_stock(v_item.product_id, v_item.quantity_box, v_item.quantity_pack);
    end loop;
  end if;

  update public.orders
  set status = p_status,
      assigned_employee_id = p_employee_id,
      stock_deducted_at = case
        when p_status = 'DELIVERING' and v_order.stock_deducted_at is null then now()
        else stock_deducted_at
      end
  where id = p_order_id
  returning * into v_order;

  return v_order;
end;
$$;

revoke all on function public.custom_deduct_product_stock(uuid, integer, integer) from public, anon, authenticated;
revoke all on function public.custom_create_retailer_order(uuid, jsonb, uuid) from public, anon, authenticated;
revoke all on function public.custom_update_delivery_status(uuid, uuid, public.order_status) from public, anon, authenticated;
grant execute on function public.custom_deduct_product_stock(uuid, integer, integer) to service_role;
grant execute on function public.custom_create_retailer_order(uuid, jsonb, uuid) to service_role;
grant execute on function public.custom_update_delivery_status(uuid, uuid, public.order_status) to service_role;

commit;
