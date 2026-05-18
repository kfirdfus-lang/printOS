import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const indexPath = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "index.html");
let s = fs.readFileSync(indexPath, "utf8");

const start = s.indexOf("function EmailModal({ client, mode, emailType");
const end = s.indexOf("\nfunction DebtActionModal(", start);
if (start < 0 || end < 0) {
  console.error("markers not found", start, end);
  process.exit(1);
}

const newFn = `function EmailModal({ client, mode, emailType, onClose, onSendEmail, sb, createdBy }) {
  const [savedEmails, setSavedEmails] = useState([]);
  const [emailMode, setEmailMode] = useState('loading');
  const [selectedEmailId, setSelectedEmailId] = useState(null);
  const [recipient, setRecipient] = useState('');
  const [showAddNew, setShowAddNew] = useState(false);
  const [newEmail, setNewEmail] = useState('');
  const [newLabel, setNewLabel] = useState('');
  const [saving, setSaving] = useState(false);
  const prevEmailIdRef = useRef(null);
  const clientId = client?.client_id || null;
  const isEditMode = mode === 'editEmail';
  const isConfirmMode = mode === 'confirmSend';
  const overlayStyle = { position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 9999, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 };
  const modalStyle = { background: '#fff', borderRadius: 16, maxWidth: 520, width: '100%', direction: 'rtl', overflow: 'hidden', maxHeight: '90vh', display: 'flex', flexDirection: 'column' };
  const inputStyle = { width: '100%', padding: '10px 12px', fontSize: 14, border: '1px solid #e5e7eb', borderRadius: 8, fontFamily: 'inherit', boxSizing: 'border-box', direction: 'ltr', textAlign: 'left' };
  const handleSelectEmail = (em) => { setSelectedEmailId(em.id); setRecipient(em.email); setShowAddNew(false); };
  const handleStartAddNew = () => {
    if (selectedEmailId) prevEmailIdRef.current = selectedEmailId;
    setSelectedEmailId(null); setRecipient(''); setNewEmail(''); setNewLabel(''); setShowAddNew(true);
  };
  const handleSaveNewEmail = async () => {
    if (!newEmail.trim() || !newEmail.includes('@')) { window.alert('כתובת מייל לא תקינה'); return; }
    if (!clientId) { window.alert('שגיאה: לא ניתן לזהות את הלקוח'); return; }
    try {
      const result = await sb('client_billing_emails', { method: 'POST', body: JSON.stringify({
        client_id: clientId, email: newEmail.trim().toLowerCase(), label: newLabel.trim() || null,
        is_default: savedEmails.length === 0, usage_count: 0, created_by: createdBy || 'unknown',
      }) });
      const newE = Array.isArray(result) ? result[0] : result;
      if (newE) {
        setSavedEmails((prev) => [...prev, newE]);
        setSelectedEmailId(newE.id); setRecipient(newE.email);
        setShowAddNew(false); setEmailMode('pick'); setNewEmail(''); setNewLabel('');
      }
    } catch (e) {
      const msg = e.message || String(e);
      if (msg.includes('duplicate') || msg.includes('23505')) window.alert('כתובת מייל זו כבר קיימת במאגר');
      else window.alert('שגיאה: ' + msg);
    }
  };
  const incrementUsage = async () => {
    if (!selectedEmailId) return;
    const em = savedEmails.find((e) => e.id === selectedEmailId);
    if (!em) return;
    try {
      await sb(\`client_billing_emails?id=eq.\${selectedEmailId}\`, { method: 'PATCH', body: JSON.stringify({
        usage_count: (em.usage_count || 0) + 1, last_used_at: new Date().toISOString(),
      }) });
    } catch (_) {}
  };
  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setEmailMode('loading');
      if (!clientId) {
        if (!cancelled) {
          setSavedEmails([]);
          setEmailMode('addFirst');
          if (client.collection_email_primary) {
            setRecipient(client.collection_email_primary);
            setNewEmail(client.collection_email_primary);
          }
        }
        return;
      }
      try {
        const data = await loadClientBillingEmailPool(clientId);
        if (cancelled) return;
        setSavedEmails(data);
        if (data.length > 0) {
          setEmailMode('pick');
          const def = data.find((e) => e.is_default) || data[0];
          setSelectedEmailId(def.id);
          setRecipient(def.email);
        } else {
          setEmailMode('addFirst');
          if (client.collection_email_primary) {
            setRecipient(client.collection_email_primary);
            setNewEmail(client.collection_email_primary);
          }
        }
      } catch (e) {
        console.error(e);
        if (!cancelled) setEmailMode('addFirst');
      }
    };
    load();
    return () => { cancelled = true; };
  }, [clientId, client.collection_email_primary]);
  const billingPickerUi = <>
    <motion.div style={{ marginBottom: 15 }}>
      <label style={{ display: 'block', marginBottom: 6, fontWeight: 600, fontSize: 14 }}>💼 מייל הנה"ח / גבייה:</label>
      {emailMode === 'loading' && <div style={{ padding: 12, background: '#f9fafb', borderRadius: 6, color: '#6b7280', fontSize: 13, textAlign: 'center' }}>טוען מיילים שמורים...</div>}
      {emailMode === 'pick' && !showAddNew && <>
        <div style={{ background: '#fef3c7', padding: 12, borderRadius: 8, border: '1px solid #fde68a', marginBottom: 8 }}>
          <div style={{ fontSize: 12, color: '#78350f', marginBottom: 8, fontWeight: 600 }}>💼 מיילים שמורים ללקוח ({savedEmails.length})</div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxHeight: 220, overflowY: 'auto' }}>
            {savedEmails.map((em) => <button type="button" key={em.id} onClick={() => handleSelectEmail(em)} style={{ padding: 10, textAlign: 'right', background: selectedEmailId === em.id ? '#92400e' : 'white', color: selectedEmailId === em.id ? 'white' : '#374151', border: \`1px solid \${selectedEmailId === em.id ? '#92400e' : '#d1d5db'}\`, borderRadius: 6, cursor: 'pointer', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontFamily: 'inherit' }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontSize: 13, direction: 'ltr', textAlign: 'right', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontWeight: 600 }}>{em.is_default && '⭐ '}{em.email}</motion.div>
                {em.label && <div style={{ fontSize: 11, color: selectedEmailId === em.id ? 'rgba(255,255,255,0.85)' : '#6b7280', marginTop: 2 }}>🏷️ {em.label}</div>}
              </div>
              {em.usage_count > 0 && <div style={{ fontSize: 11, marginRight: 8, color: selectedEmailId === em.id ? 'rgba(255,255,255,0.7)' : '#9ca3af', whiteSpace: 'nowrap' }}>{em.usage_count}×</div>}
            </button>)}
          </div>
        </div>
        <button type="button" onClick={handleStartAddNew} style={{ width: '100%', padding: 10, background: 'transparent', border: '2px dashed #d1d5db', borderRadius: 8, cursor: 'pointer', color: '#6b7280', fontSize: 13, fontFamily: 'inherit' }}>➕ הוסף מייל גבייה נוסף</button>
      </>}
      {(emailMode === 'addFirst' || showAddNew) && <div style={{ background: '#fefce8', padding: 12, borderRadius: 8, border: '2px solid #fde68a' }}>
        {emailMode === 'addFirst' && <div style={{ fontSize: 12, color: '#78350f', marginBottom: 10 }}>ℹ️ אין מיילי גבייה שמורים ללקוח - הוסף את הראשון</div>}
        <div style={{ marginBottom: 8 }}><input type="email" value={showAddNew ? newEmail : (newEmail || recipient)} onChange={(e) => { const v = e.target.value; setNewEmail(v); if (emailMode === 'addFirst') setRecipient(v); }} placeholder="email@company.co.il *" style={inputStyle} autoFocus /></div>
        <div style={{ marginBottom: emailMode === 'pick' && showAddNew ? 10 : 0 }}><input type="text" value={newLabel} onChange={(e) => setNewLabel(e.target.value)} placeholder='תווית (אופציונלי) - "הנה"ח", "גבייה"...' style={{ ...inputStyle, direction: 'rtl', textAlign: 'right' }} /></div>
        {emailMode === 'pick' && showAddNew && <div style={{ display: 'flex', gap: 6, marginTop: 8 }}>
          <button type="button" onClick={() => { setShowAddNew(false); const prevId = prevEmailIdRef.current; if (prevId) { const prev = savedEmails.find((e) => e.id === prevId); if (prev) { setSelectedEmailId(prev.id); setRecipient(prev.email); } } }} style={{ flex: 1, padding: 8, fontSize: 12, background: '#f3f4f6', border: '1px solid #d1d5db', borderRadius: 6, cursor: 'pointer', fontFamily: 'inherit' }}>ביטול</button>
          <button type="button" onClick={handleSaveNewEmail} style={{ flex: 2, padding: 8, fontSize: 12, background: '#92400e', color: 'white', border: 'none', borderRadius: 6, fontWeight: 600, cursor: 'pointer', fontFamily: 'inherit' }}>💾 שמור והשתמש</button>
        </div>}
        {emailMode === 'addFirst' && <div style={{ fontSize: 11, color: '#6b7280', marginTop: 8, textAlign: 'center' }}>💡 המייל יישמר אוטומטית כשתשלח</div>}
      </div>}
    </div>
  </>;
  return (
    <div style={overlayStyle} onClick={onClose}>
      <div style={modalStyle} onClick={(e) => e.stopPropagation()}>
        <div style={{ padding: '16px 20px', borderBottom: '1px solid #e5e7eb', background: '#fffbeb', flexShrink: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <div>
              <motion.div style={{ fontSize: 17, fontWeight: 700, color: '#1E3A52' }}>{isEditMode ? '✏️ מיילי גבייה' : '📧 אישור שליחה'}</div>
              <div style={{ fontSize: 13, color: '#6b7280', marginTop: 3 }}>{client.name}</div>
            </div>
            <button type="button" onClick={onClose} style={{ background: 'none', border: 'none', fontSize: 24, cursor: 'pointer', color: '#9ca3af' }}>✕</button>
          </div>
        </div>
        <div style={{ padding: 20, overflowY: 'auto', flex: 1 }}>
          {isConfirmMode && <p style={{ fontSize: 14, color: '#1E3A52', lineHeight: 1.7, marginTop: 0, marginBottom: 16 }}>
            האם לשלוח <strong>{emailType === 'inquiry' ? 'בקשת בירור' : 'תזכורת תשלום'}</strong>?
            {client.invoices?.length ? <> עם <strong>{client.invoices.length} חשבוניות</strong></> : null}
          </p>}
          {isEditMode && <p style={{ fontSize: 13, color: '#6b7280', marginTop: 0, marginBottom: 12 }}>נהל את מאגר מיילי הגבייה ללקוח זה.</p>}
          {billingPickerUi}
        </div>
        <div style={{ padding: 16, borderTop: '1px solid #e5e7eb', display: 'flex', gap: 8, justifyContent: 'flex-end', flexShrink: 0 }}>
          <button type="button" onClick={onClose} disabled={saving} style={{ padding: '10px 20px', borderRadius: 8, fontSize: 14, fontWeight: 600, background: '#fff', color: '#6b7280', border: '1px solid #e5e7eb', cursor: 'pointer', fontFamily: 'inherit' }}>ביטול</button>
          {!isEditMode && <button type="button" onClick={async () => {
            const toSend = (emailMode === 'addFirst' ? (newEmail || recipient) : recipient).trim().toLowerCase();
            if (!toSend || !toSend.includes('@')) { window.alert('יש לבחור או להזין מייל תקין'); return; }
            setSaving(true);
            try {
              if (!selectedEmailId && clientId && toSend) {
                try {
                  await sb('client_billing_emails', { method: 'POST', body: JSON.stringify({
                    client_id: clientId, email: toSend, label: newLabel.trim() || null,
                    is_default: savedEmails.length === 0, usage_count: 1, last_used_at: new Date().toISOString(),
                    created_by: createdBy || 'unknown',
                  }) });
                } catch (e) { console.warn('Failed to auto-save billing email:', e); }
              } else if (selectedEmailId) { await incrementUsage(); }
              const ok = await onSendEmail(client, emailType, { recipient: toSend, selectedEmailId });
              if (ok !== false) onClose();
            } finally { setSaving(false); }
          }} disabled={saving || emailMode === 'loading'} style={{ padding: '10px 24px', borderRadius: 8, fontSize: 14, fontWeight: 700, background: isConfirmMode && emailType === 'inquiry' ? '#dc2626' : '#92400e', color: '#fff', border: 'none', cursor: saving ? 'wait' : 'pointer', fontFamily: 'inherit', opacity: saving || emailMode === 'loading' ? 0.6 : 1 }}>
            {saving ? '⏳ שולח...' : '📤 שלח'}
          </button>}
          {isEditMode && <button type="button" onClick={onClose} style={{ padding: '10px 24px', borderRadius: 8, fontSize: 14, fontWeight: 700, background: '#92400e', color: '#fff', border: 'none', cursor: 'pointer', fontFamily: 'inherit' }}>סגור</button>}
        </div>
      </div>
    </div>
  );
}
`;

const clean = newFn.replace(/<\/?motion\.div/g, (m) => m.replace("motion.", ""));
s = s.slice(0, start) + clean + s.slice(end);
fs.writeFileSync(indexPath, s, "utf8");
console.log("OK: EmailModal replaced");
