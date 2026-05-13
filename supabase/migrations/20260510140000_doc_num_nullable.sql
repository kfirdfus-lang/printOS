-- מאפשר doc_num להיות NULL (לרשומות יתרת פתיחה ללא חשבונית)
ALTER TABLE debt_snapshots ALTER COLUMN doc_num DROP NOT NULL;

-- מסיר את האינדקס היחודי הישן ויוצר אחד שמתעלם מ-NULL
DROP INDEX IF EXISTS idx_debt_snapshots_unique;
CREATE UNIQUE INDEX idx_debt_snapshots_unique
  ON debt_snapshots(snapshot_date, bina_customer_id, doc_num)
  WHERE doc_num IS NOT NULL;

-- אינדקס נוסף לרשומות בלי doc_num (יתרת פתיחה)
CREATE UNIQUE INDEX idx_debt_snapshots_no_docnum
  ON debt_snapshots(snapshot_date, bina_customer_id)
  WHERE doc_num IS NULL;
