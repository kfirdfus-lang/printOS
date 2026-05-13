-- מאפשר NULL בעמודות נוספות (בינה מחזירים רשומות חלקיות לפעמים)
ALTER TABLE debt_snapshots ALTER COLUMN doc_balance DROP NOT NULL;
