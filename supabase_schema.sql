-- ═══════════════════════════════════════════
-- МАГАЗИН ЕСАБЫ — SUPABASE SQL SCHEMA
-- ═══════════════════════════════════════════

-- 1. ПРОФИЛЛЕР (роллер)
CREATE TABLE profiles (
  id UUID REFERENCES auth.users(id) PRIMARY KEY,
  full_name TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('director', 'seller', 'supply')),
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. КАТЕГОРИЯЛАР
CREATE TABLE categories (
  id SERIAL PRIMARY KEY,
  name TEXT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. ТОВАРЛАР
CREATE TABLE products (
  id SERIAL PRIMARY KEY,
  name TEXT NOT NULL,
  barcode TEXT UNIQUE,
  category_id INT REFERENCES categories(id),
  buy_price NUMERIC NOT NULL DEFAULT 0,
  sell_price NUMERIC NOT NULL DEFAULT 0,
  stock NUMERIC NOT NULL DEFAULT 0,
  min_stock NUMERIC NOT NULL DEFAULT 5,
  unit TEXT NOT NULL DEFAULT 'дана',
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 4. КЛИЕНТЛЕР
CREATE TABLE clients (
  id SERIAL PRIMARY KEY,
  name TEXT NOT NULL,
  phone TEXT,
  debt NUMERIC DEFAULT 0,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. САТЫЎЛАР
CREATE TABLE sales (
  id SERIAL PRIMARY KEY,
  seller_id UUID REFERENCES profiles(id),
  client_id INT REFERENCES clients(id),
  payment_type TEXT NOT NULL CHECK (payment_type IN ('cash', 'card', 'qr', 'debt')),
  total NUMERIC NOT NULL DEFAULT 0,
  comment TEXT,
  date DATE NOT NULL DEFAULT CURRENT_DATE,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 6. САТЫЎ ЭЛЕМЕНТЛЕРИ
CREATE TABLE sale_items (
  id SERIAL PRIMARY KEY,
  sale_id INT REFERENCES sales(id) ON DELETE CASCADE,
  product_id INT REFERENCES products(id),
  product_name TEXT NOT NULL,
  qty NUMERIC NOT NULL,
  sell_price NUMERIC NOT NULL,
  buy_price NUMERIC NOT NULL
);

-- 7. ПРИХОД (кириш)
CREATE TABLE purchases (
  id SERIAL PRIMARY KEY,
  supply_id UUID REFERENCES profiles(id),
  comment TEXT,
  date DATE NOT NULL DEFAULT CURRENT_DATE,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 8. ПРИХОД ЭЛЕМЕНТЛЕРИ
CREATE TABLE purchase_items (
  id SERIAL PRIMARY KEY,
  purchase_id INT REFERENCES purchases(id) ON DELETE CASCADE,
  product_id INT REFERENCES products(id),
  product_name TEXT NOT NULL,
  qty NUMERIC NOT NULL,
  buy_price NUMERIC NOT NULL,
  sell_price NUMERIC NOT NULL
);

-- 9. КАССА ТАПСЫРЫЎ
CREATE TABLE cash_handovers (
  id SERIAL PRIMARY KEY,
  seller_id UUID REFERENCES profiles(id),
  amount NUMERIC NOT NULL,
  comment TEXT,
  date DATE NOT NULL DEFAULT CURRENT_DATE,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 10. КҮН ЖАБЫЎ
CREATE TABLE day_closings (
  id SERIAL PRIMARY KEY,
  closed_by UUID REFERENCES profiles(id),
  total_sales NUMERIC DEFAULT 0,
  total_cash NUMERIC DEFAULT 0,
  total_card NUMERIC DEFAULT 0,
  total_qr NUMERIC DEFAULT 0,
  total_debt NUMERIC DEFAULT 0,
  total_profit NUMERIC DEFAULT 0,
  comment TEXT,
  date DATE NOT NULL DEFAULT CURRENT_DATE,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 11. СОРАНЫСЛАР (жоқ товар)
CREATE TABLE requests (
  id SERIAL PRIMARY KEY,
  seller_id UUID REFERENCES profiles(id),
  product_name TEXT NOT NULL,
  qty NUMERIC,
  comment TEXT,
  status TEXT DEFAULT 'new' CHECK (status IN ('new', 'ordered', 'done')),
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- ═══════════════════════════════════════════
-- ROW LEVEL SECURITY (RLS)
-- ═══════════════════════════════════════════

ALTER TABLE profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE products ENABLE ROW LEVEL SECURITY;
ALTER TABLE categories ENABLE ROW LEVEL SECURITY;
ALTER TABLE clients ENABLE ROW LEVEL SECURITY;
ALTER TABLE sales ENABLE ROW LEVEL SECURITY;
ALTER TABLE sale_items ENABLE ROW LEVEL SECURITY;
ALTER TABLE purchases ENABLE ROW LEVEL SECURITY;
ALTER TABLE purchase_items ENABLE ROW LEVEL SECURITY;
ALTER TABLE cash_handovers ENABLE ROW LEVEL SECURITY;
ALTER TABLE day_closings ENABLE ROW LEVEL SECURITY;
ALTER TABLE requests ENABLE ROW LEVEL SECURITY;

-- Ҳәмме аутентификацияланған пайдаланыўшылар оқый алады
CREATE POLICY "authenticated_read_profiles" ON profiles FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_products" ON products FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_categories" ON categories FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_clients" ON clients FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_sales" ON sales FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_sale_items" ON sale_items FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_purchases" ON purchases FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_purchase_items" ON purchase_items FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_handovers" ON cash_handovers FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_closings" ON day_closings FOR SELECT TO authenticated USING (true);
CREATE POLICY "authenticated_read_requests" ON requests FOR SELECT TO authenticated USING (true);

-- Жазыў политикалары
CREATE POLICY "authenticated_insert_products" ON products FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "authenticated_update_products" ON products FOR UPDATE TO authenticated USING (true);
CREATE POLICY "authenticated_delete_products" ON products FOR DELETE TO authenticated USING (true);

CREATE POLICY "authenticated_insert_categories" ON categories FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "authenticated_update_categories" ON categories FOR UPDATE TO authenticated USING (true);

CREATE POLICY "authenticated_insert_clients" ON clients FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "authenticated_update_clients" ON clients FOR UPDATE TO authenticated USING (true);

CREATE POLICY "authenticated_insert_sales" ON sales FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "authenticated_insert_sale_items" ON sale_items FOR INSERT TO authenticated WITH CHECK (true);

CREATE POLICY "authenticated_insert_purchases" ON purchases FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "authenticated_insert_purchase_items" ON purchase_items FOR INSERT TO authenticated WITH CHECK (true);

CREATE POLICY "authenticated_insert_handovers" ON cash_handovers FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "authenticated_insert_closings" ON day_closings FOR INSERT TO authenticated WITH CHECK (true);

CREATE POLICY "authenticated_insert_requests" ON requests FOR INSERT TO authenticated WITH CHECK (true);
CREATE POLICY "authenticated_update_requests" ON requests FOR UPDATE TO authenticated USING (true);

CREATE POLICY "profiles_insert" ON profiles FOR INSERT TO authenticated WITH CHECK (auth.uid() = id);
CREATE POLICY "profiles_update" ON profiles FOR UPDATE TO authenticated USING (auth.uid() = id);

-- ═══════════════════════════════════════════
-- TRIGGERS — жаңа пайдаланыўшы профиль
-- ═══════════════════════════════════════════
CREATE OR REPLACE FUNCTION handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO profiles (id, full_name, role)
  VALUES (NEW.id, COALESCE(NEW.raw_user_meta_data->>'full_name', 'Пайдаланыўшы'), COALESCE(NEW.raw_user_meta_data->>'role', 'seller'));
  RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION handle_new_user();

-- ═══════════════════════════════════════════
-- БАСЛАПҚЫ МАҒЛЫЎМАТЛАР
-- ═══════════════════════════════════════════
INSERT INTO categories (name) VALUES
  ('Азық-аўқат'),
  ('Ишимликлер'),
  ('Хожалық'),
  ('Гигиена'),
  ('Басқа');
