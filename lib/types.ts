export type Role = "OWNER" | "ADMIN" | "EMPLOYEE" | "RETAILER";

export type OrderStatus =
  | "PENDING"
  | "PACKED"
  | "DELIVERING"
  | "COMPLETED";

export type DashboardView = "overview" | "orders" | "add-product" | "categories" | "tracking" | "cart";

export interface Product {
  id: string;
  name: string;
  category?: string;
  category_id?: string | null;
  price_box: number;
  price_pack: number | null;
  // Required when price_pack is set. Null on older products, which still treat one pack as one stock unit.
  packs_per_box?: number | null;
  // `stock` counts full boxes; `stock_pack` counts loose packs from an opened box.
  stock: number;
  stock_pack?: number;
  created_at?: string;
}

export interface ProductCategory {
  id: string;
  name: string;
  normalized_name: string;
  is_active: boolean;
  created_at?: string;
  updated_at?: string;
}

export interface Retailer {
  id: string;
  shop_name: string;
  phone?: string | null;
  role?: Role;
}

export interface OrderItem {
  id?: string;
  // Products can be deleted after an order is placed. The database deliberately
  // keeps the historical order line and clears this foreign key in that case.
  product_id: string | null;
  product_name?: string;
  quantity_box: number;
  quantity_pack: number;
  unit_price_box: number;
  unit_price_pack: number | null;
  line_total: number;
  product?: Product | null;
}

export interface ReceiptItem {
  id?: string;
  product_id: string | null;
  product_name: string;
  category?: string | null;
  quantity_box: number;
  quantity_pack: number;
  unit_price_box: number;
  unit_price_pack: number | null;
  line_total: number;
  product?: Product | null;
}

export interface DailyReceipt {
  id: string;
  retailer_id: string;
  receipt_date: string;
  total_amount: number;
  created_at: string;
  updated_at?: string;
  retailer?: Retailer | null;
  items: ReceiptItem[];
}

export interface DeliveryTracking {
  order_id: string;
  employee_id?: string | null;
  latitude: number;
  longitude: number;
  updated_at: string;
  employee?: Pick<Retailer, "id" | "shop_name"> | null;
}

export interface Order {
  id: string;
  retailer_id: string;
  receipt_id?: string | null;
  assigned_employee_id?: string | null;
  status: OrderStatus;
  total_amount: number;
  created_at: string;
  retailer?: Retailer | null;
  items: OrderItem[];
  tracking?: DeliveryTracking | null;
}

export interface CartLine {
  product: Product;
  boxes: number;
  packs: number;
}

export interface AppSession {
  expires_at?: number;
  user: {
    id: string;
    phone?: string | null;
  };
}

export interface DashboardSnapshot {
  products: Product[];
  categories: ProductCategory[];
  orders: Order[];
  receipts: DailyReceipt[];
  retailers: Retailer[];
  tracking: DeliveryTracking[];
}
