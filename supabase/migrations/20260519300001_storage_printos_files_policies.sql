-- ============================================================
-- Storage policies for bucket: printos-files
-- יש ליצור את ה-bucket ידנית ב-Dashboard לפני הרצת מדיניות אלה
-- ============================================================

DROP POLICY IF EXISTS "Allow public read access to printos-files" ON storage.objects;
CREATE POLICY "Allow public read access to printos-files"
  ON storage.objects FOR SELECT
  USING (bucket_id = 'printos-files');

DROP POLICY IF EXISTS "Allow public insert to printos-files" ON storage.objects;
CREATE POLICY "Allow public insert to printos-files"
  ON storage.objects FOR INSERT
  WITH CHECK (bucket_id = 'printos-files');

DROP POLICY IF EXISTS "Allow public update to printos-files" ON storage.objects;
CREATE POLICY "Allow public update to printos-files"
  ON storage.objects FOR UPDATE
  USING (bucket_id = 'printos-files');

DROP POLICY IF EXISTS "Allow public delete to printos-files" ON storage.objects;
CREATE POLICY "Allow public delete to printos-files"
  ON storage.objects FOR DELETE
  USING (bucket_id = 'printos-files');
