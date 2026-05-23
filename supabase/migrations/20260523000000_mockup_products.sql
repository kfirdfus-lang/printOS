-- ============================================================
-- טבלת מוצרי הדמיה
-- ============================================================
CREATE TABLE IF NOT EXISTS public.mockup_products (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  
  -- שמות
  name text NOT NULL,                  -- "חולצת T"
  name_en text NOT NULL,               -- "T-shirt" (חשוב ל-AI)
  emoji text NOT NULL,                 -- "👕"
  category text NOT NULL,              -- "clothing" / "accessories" / "drink" / וכו'
  
  -- תיאור ל-AI
  ai_description text,                 -- "casual cotton t-shirt"
  
  -- אופציות
  colors jsonb DEFAULT '[]'::jsonb,    -- ["שחור", "לבן", "כחול"]
  views jsonb DEFAULT '[]'::jsonb,     -- ["חזית", "גב", "צד"]
  print_locations jsonb DEFAULT '[]'::jsonb,  
  -- [{name: "חזית מרכז", max_width_cm: 25, max_height_cm: 30, view: "חזית"}]
  
  -- תבניות (יתווסף בעתיד)
  template_files jsonb DEFAULT '[]'::jsonb,
  -- [{path, view, color}]
  
  -- מטא
  is_active boolean DEFAULT true,
  is_default boolean DEFAULT false,    -- מוצר ברירת מחדל (לא ניתן למחיקה)
  sort_order int DEFAULT 100,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_mockup_products_active 
  ON public.mockup_products(is_active) WHERE is_active = true;
CREATE INDEX IF NOT EXISTS idx_mockup_products_category 
  ON public.mockup_products(category);

-- RLS
ALTER TABLE public.mockup_products ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "mockup_products_all" ON public.mockup_products;
CREATE POLICY "mockup_products_all"
  ON public.mockup_products FOR ALL USING (true) WITH CHECK (true);

GRANT ALL ON public.mockup_products TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
INSERT INTO public.mockup_products (name, name_en, emoji, category, ai_description, colors, views, print_locations, is_default, sort_order)
VALUES
-- 👕 ביגוד
('חולצת T', 'T-shirt', '👕', 'clothing', 'casual cotton t-shirt', 
  '["שחור", "לבן", "כחול", "אדום", "אפור", "ירוק"]'::jsonb,
  '["חזית", "גב", "צד"]'::jsonb,
  '[
    {"name":"חזית מרכז","view":"חזית","max_width_cm":25,"max_height_cm":30},
    {"name":"גב גדול","view":"גב","max_width_cm":30,"max_height_cm":40},
    {"name":"שרוול שמאל","view":"צד","max_width_cm":8,"max_height_cm":8},
    {"name":"שרוול ימין","view":"צד","max_width_cm":8,"max_height_cm":8}
  ]'::jsonb,
  true, 10),

('חולצת שרוול ארוך', 'long sleeve t-shirt', '🥼', 'clothing', 'long sleeve cotton t-shirt',
  '["שחור", "לבן", "כחול", "אפור"]'::jsonb,
  '["חזית", "גב", "צד"]'::jsonb,
  '[
    {"name":"חזית מרכז","view":"חזית","max_width_cm":25,"max_height_cm":30},
    {"name":"גב גדול","view":"גב","max_width_cm":30,"max_height_cm":40},
    {"name":"שרוול שמאל","view":"צד","max_width_cm":10,"max_height_cm":10},
    {"name":"שרוול ימין","view":"צד","max_width_cm":10,"max_height_cm":10}
  ]'::jsonb,
  true, 20),

('חולצת פולו', 'polo shirt', '👔', 'clothing', 'classic polo shirt with collar',
  '["שחור", "לבן", "כחול נייבי", "אפור"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית שמאל (לוגו)","view":"חזית","max_width_cm":8,"max_height_cm":8},
    {"name":"גב גדול","view":"גב","max_width_cm":30,"max_height_cm":35}
  ]'::jsonb,
  true, 30),

('קפוצ''ון', 'hoodie', '🧥', 'clothing', 'cotton hoodie sweatshirt',
  '["שחור", "אפור", "לבן", "כחול נייבי"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית מרכז","view":"חזית","max_width_cm":25,"max_height_cm":25},
    {"name":"גב גדול","view":"גב","max_width_cm":30,"max_height_cm":40},
    {"name":"קפוצ''ון מאחור","view":"גב","max_width_cm":15,"max_height_cm":5}
  ]'::jsonb,
  true, 40),

('וסט', 'vest', '🦺', 'clothing', 'work vest',
  '["צהוב", "כתום", "ירוק", "שחור"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית שמאל","view":"חזית","max_width_cm":10,"max_height_cm":10},
    {"name":"גב גדול","view":"גב","max_width_cm":25,"max_height_cm":15}
  ]'::jsonb,
  true, 50),

('בגד תינוק', 'baby onesie', '👶', 'clothing', 'baby cotton onesie',
  '["לבן", "ורוד", "תכלת", "צהוב"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית מרכז","view":"חזית","max_width_cm":15,"max_height_cm":15},
    {"name":"גב","view":"גב","max_width_cm":15,"max_height_cm":15}
  ]'::jsonb,
  true, 60),

('כובע', 'cap', '🧢', 'clothing', 'baseball cap',
  '["שחור", "לבן", "אפור", "כחול נייבי", "אדום"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"חזית מרכז","view":"חזית","max_width_cm":10,"max_height_cm":6}
  ]'::jsonb,
  true, 70),

-- 🎒 אביזרים
('תיק טוט (בד)', 'cotton tote bag', '👜', 'accessories', 'natural canvas tote bag with handles',
  '["טבעי בז''", "לבן", "שחור", "ירוק"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית מרכז","view":"חזית","max_width_cm":25,"max_height_cm":25}
  ]'::jsonb,
  true, 80),

('תיק לא-ארוג', 'non-woven bag', '🛍️', 'accessories', 'non-woven shopping bag',
  '["לבן", "שחור", "כחול", "אדום", "ירוק"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית מרכז","view":"חזית","max_width_cm":30,"max_height_cm":25}
  ]'::jsonb,
  true, 90),

('תיק שרוכים', 'drawstring bag', '🎒', 'accessories', 'drawstring sports bag',
  '["שחור", "לבן", "כחול", "אדום"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית מרכז","view":"חזית","max_width_cm":25,"max_height_cm":30}
  ]'::jsonb,
  true, 100),

-- ☕ כלי שתייה
('ספל קפה', 'coffee mug', '☕', 'drinkware', 'ceramic coffee mug',
  '["לבן", "שחור"]'::jsonb,
  '["חזית", "צד", "סביב"]'::jsonb,
  '[
    {"name":"צד אחד","view":"צד","max_width_cm":8,"max_height_cm":8},
    {"name":"שני צדדים","view":"סביב","max_width_cm":20,"max_height_cm":8}
  ]'::jsonb,
  true, 110),

-- 📋 פרסום ושיווק
('פלייר A4', 'flyer A4', '📋', 'marketing', 'A4 flyer print',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"מלא חזית","view":"חזית","max_width_cm":21,"max_height_cm":29.7},
    {"name":"מלא גב","view":"גב","max_width_cm":21,"max_height_cm":29.7}
  ]'::jsonb,
  true, 120),

('פלייר A5', 'flyer A5', '📋', 'marketing', 'A5 flyer print',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"מלא חזית","view":"חזית","max_width_cm":14.8,"max_height_cm":21},
    {"name":"מלא גב","view":"גב","max_width_cm":14.8,"max_height_cm":21}
  ]'::jsonb,
  true, 130),

('פלייר A6', 'flyer A6', '📋', 'marketing', 'A6 small flyer',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"מלא חזית","view":"חזית","max_width_cm":10.5,"max_height_cm":14.8},
    {"name":"מלא גב","view":"גב","max_width_cm":10.5,"max_height_cm":14.8}
  ]'::jsonb,
  true, 140),

('פולדר', 'folder', '📂', 'marketing', 'presentation folder with pockets',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב", "פנים"]'::jsonb,
  '[
    {"name":"חזית","view":"חזית","max_width_cm":23,"max_height_cm":31},
    {"name":"גב","view":"גב","max_width_cm":23,"max_height_cm":31}
  ]'::jsonb,
  true, 150),

('ברושור מתקפל', 'tri-fold brochure', '📰', 'marketing', 'tri-fold marketing brochure',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"מלא חזית","view":"חזית","max_width_cm":29.7,"max_height_cm":21},
    {"name":"מלא גב","view":"גב","max_width_cm":29.7,"max_height_cm":21}
  ]'::jsonb,
  true, 160),

('כרטיס ביקור', 'business card', '🎟️', 'marketing', 'business card 9x5cm',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית","view":"חזית","max_width_cm":9,"max_height_cm":5},
    {"name":"גב","view":"גב","max_width_cm":9,"max_height_cm":5}
  ]'::jsonb,
  true, 170),

('כרטיס מתנה', 'gift card', '📇', 'marketing', 'gift card / voucher',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית","view":"חזית","max_width_cm":8.5,"max_height_cm":5.5},
    {"name":"גב","view":"גב","max_width_cm":8.5,"max_height_cm":5.5}
  ]'::jsonb,
  true, 180),

('גלויה', 'postcard', '📨', 'marketing', 'postcard 10x15cm',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית","view":"חזית","max_width_cm":10,"max_height_cm":15},
    {"name":"גב","view":"גב","max_width_cm":10,"max_height_cm":15}
  ]'::jsonb,
  true, 190),

-- 🏷️ מדבקות
('מדבקה עגולה', 'round sticker', '🔘', 'stickers', 'circular product sticker',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":10,"max_height_cm":10}
  ]'::jsonb,
  true, 200),

('מדבקה מלבנית', 'rectangular sticker', '📦', 'stickers', 'rectangular packaging sticker',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":10,"max_height_cm":5}
  ]'::jsonb,
  true, 210),

('מדבקה מותאמת', 'die-cut sticker', '🎨', 'stickers', 'custom shape die-cut sticker',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":15,"max_height_cm":15}
  ]'::jsonb,
  true, 220),

('מדבקת רכב', 'car sticker', '🚗', 'stickers', 'vehicle sticker for car',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":30,"max_height_cm":15}
  ]'::jsonb,
  true, 230),

('מדבקה עמידת מים', 'waterproof sticker', '💧', 'stickers', 'waterproof outdoor sticker',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":15,"max_height_cm":15}
  ]'::jsonb,
  true, 240),

('מדבקה הולוגרמה', 'hologram sticker', '✨', 'stickers', 'holographic shiny sticker',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":10,"max_height_cm":10}
  ]'::jsonb,
  true, 250),

-- 📝 משרד
('מחברת / יומן', 'notebook', '📓', 'office', 'spiral notebook journal',
  '["שחור", "לבן", "כחול", "אדום"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"חזית","view":"חזית","max_width_cm":15,"max_height_cm":20},
    {"name":"גב","view":"גב","max_width_cm":15,"max_height_cm":20}
  ]'::jsonb,
  true, 260),

('עט', 'pen', '🖊️', 'office', 'branded pen',
  '["שחור", "לבן", "כחול", "אדום"]'::jsonb,
  '["צד"]'::jsonb,
  '[
    {"name":"גוף העט","view":"צד","max_width_cm":5,"max_height_cm":1}
  ]'::jsonb,
  true, 270),

('פנקס דביק', 'sticky notes', '📌', 'office', 'branded sticky notes pad',
  '["לבן", "צהוב"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"כריכה","view":"חזית","max_width_cm":7.5,"max_height_cm":7.5}
  ]'::jsonb,
  true, 280),

('לוח שנה', 'calendar', '🗓️', 'office', 'wall calendar',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"כריכה","view":"חזית","max_width_cm":30,"max_height_cm":40}
  ]'::jsonb,
  true, 290),

('פוסטר', 'poster', '🖼️', 'office', 'framed poster',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"A3","view":"חזית","max_width_cm":29.7,"max_height_cm":42},
    {"name":"A2","view":"חזית","max_width_cm":42,"max_height_cm":59.4}
  ]'::jsonb,
  true, 300),

('פד עכבר', 'mousepad', '🖱️', 'office', 'computer mousepad',
  '["שחור", "לבן", "צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":22,"max_height_cm":18}
  ]'::jsonb,
  true, 310),

-- 🎁 מזכרות
('מזכרת', 'souvenir', '🎁', 'gifts', 'custom souvenir item',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"חזית","view":"חזית","max_width_cm":15,"max_height_cm":15}
  ]'::jsonb,
  true, 320),

-- 🏭 דפוס מיוחד
('קנבס', 'canvas print', '🎨', 'special', 'stretched canvas print on wooden frame',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":50,"max_height_cm":70}
  ]'::jsonb,
  true, 330),

('מגנט', 'fridge magnet', '🧱', 'special', 'rectangular fridge magnet',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":10,"max_height_cm":7}
  ]'::jsonb,
  true, 340),

-- 🏗️ מתקני תצוגה
('רולאפ', 'roll-up banner', '🚩', 'displays', 'roll-up retractable banner stand',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"85x200","view":"חזית","max_width_cm":85,"max_height_cm":200},
    {"name":"100x200","view":"חזית","max_width_cm":100,"max_height_cm":200}
  ]'::jsonb,
  true, 350),

('מולטי קיוב', 'multi-cube display', '🎯', 'displays', 'multi-sided cube display stand',
  '["צבע מלא"]'::jsonb,
  '["חזית", "צד"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":80,"max_height_cm":120}
  ]'::jsonb,
  true, 360),

('מתקן חמור 70/100 אלומיניום', 'aluminum donkey stand 70x100', '🪧', 'displays', 'aluminum sandwich board sidewalk sign 70x100',
  '["צבע מלא"]'::jsonb,
  '["חזית", "גב"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":70,"max_height_cm":100}
  ]'::jsonb,
  true, 370),

('מתקן מגנטי 160/60', 'magnetic stand 160x60', '🎌', 'displays', 'magnetic banner stand 160x60 with magnetic poster',
  '["צבע מלא"]'::jsonb,
  '["חזית"]'::jsonb,
  '[
    {"name":"מלא","view":"חזית","max_width_cm":160,"max_height_cm":60}
  ]'::jsonb,
  true, 380)
;

NOTIFY pgrst, 'reload schema';
