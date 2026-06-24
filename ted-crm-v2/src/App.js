import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import { supabase } from "./supabase";

// ─── Constants ────────────────────────────────────────────────────────────────
const GENRES = ["Homme", "Femme", "Entreprise", "Non renseigné"];
const MONTHS_FR = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"];
const PAGE_SIZES = [25, 50, 100];
const G = "#E8C547";

// ─── Utilities ────────────────────────────────────────────────────────────────
function capitalize(s) { if (!s) return ""; return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase(); }
function formatDate(iso) { if (!iso) return ""; const d = new Date(iso); if (isNaN(d)) return ""; return d.toLocaleDateString("fr-FR", { day: "2-digit", month: "2-digit", year: "numeric" }); }
function normalizeStr(s) { return (s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase(); }
function getMonthName(iso) { if (!iso) return ""; const d = new Date(iso); return isNaN(d) ? "" : MONTHS_FR[d.getMonth()]; }
function getCurrentMonthName() { return MONTHS_FR[new Date().getMonth()]; }
function isCurrentMonth(iso) { if (!iso) return false; const d = new Date(iso), n = new Date(); return d.getMonth() === n.getMonth() && d.getFullYear() === n.getFullYear(); }
function genId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

function downloadBlob(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportToCSV(clients) {
  const header = ["Genre","Nom","Prénom","Téléphone","Mail","Date d'ajout","Commentaire"];
  const rows = clients.map(c => [c.genre,c.nom,c.prenom,c.tel,c.mail,formatDate(c.created_at),c.commentaire].map(v => `"${(v||"").replace(/"/g,'""')}"`));
  const csv = "\uFEFF" + [header, ...rows].map(r => r.join(";")).join("\n");
  downloadBlob(csv, "clients_TED.csv", "text/csv;charset=utf-8;");
}

function exportToXLSX(clients) {
  const header = ["Genre","Nom","Prénom","Téléphone","Mail","Date d'ajout","Commentaire"];
  const rows = clients.map(c => [c.genre||"",c.nom||"",c.prenom||"",c.tel?`\t${c.tel}`:"",c.mail||"",formatDate(c.created_at),c.commentaire||""]);
  let xml = `<?xml version="1.0" encoding="UTF-8"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="Clients"><Table>`;
  const encCell = v => `<Cell><Data ss:Type="String">${(v||"").toString().replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}</Data></Cell>`;
  xml += `<Row>${header.map(encCell).join("")}</Row>`;
  rows.forEach(r => { xml += `<Row>${r.map(encCell).join("")}</Row>`; });
  xml += `</Table></Worksheet></Workbook>`;
  downloadBlob(xml, "clients_TED.xls", "application/vnd.ms-excel");
}

function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const headers = lines[0].split(/[;,]/).map(h => h.replace(/^"|"$/g,"").trim().toLowerCase());
  return lines.slice(1).map(line => {
    const cols = line.match(/(".*?"|[^;,\n]+)(?=[;,]|$)/g) || [];
    const row = {};
    headers.forEach((h, i) => { row[h] = (cols[i]||"").replace(/^"|"$/g,"").trim(); });
    return row;
  }).filter(r => r["nom"] || r["prénom"] || r["prenom"]);
}

function mapImportRow(row) {
  const nom = capitalize(row["nom"] || "");
  const prenom = capitalize(row["prénom"] || row["prenom"] || "");
  const genre = GENRES.find(g => g.toLowerCase() === (row["genre"]||"").toLowerCase()) || "Non renseigné";
  const tel = (row["téléphone"]||row["telephone"]||"").replace(/\D/g,"").slice(0,10);
  const mail = row["mail"] || row["email"] || "";
  const commentaire = row["commentaire"] || "";
  let created_at = new Date().toISOString();
  const rawDate = row["date d'ajout"] || row["date"] || "";
  if (rawDate) {
    const parts = rawDate.match(/(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
    if (parts) { const d = new Date(parseInt(parts[3]), parseInt(parts[2])-1, parseInt(parts[1])); if (!isNaN(d)) created_at = d.toISOString(); }
  }
  return { genre, nom, prenom, tel, mail, commentaire, created_at };
}

// ─── Styles ───────────────────────────────────────────────────────────────────
const badge = (g) => ({ display:"inline-block", fontSize:11, fontWeight:600, borderRadius:99, padding:"2px 8px", background: g==="Homme"?"#dbeafe":g==="Femme"?"#fce7f3":g==="Entreprise"?"#d1fae5":"#f3f4f6", color: g==="Homme"?"#1e40af":g==="Femme"?"#be185d":g==="Entreprise"?"#065f46":"#6b7280" });
const btnPrimary = { background:G, color:"#111", border:"none", borderRadius:8, padding:"0 18px", height:40, fontWeight:700, fontSize:14, cursor:"pointer", whiteSpace:"nowrap" };
const btnSecondary = { background:"#fff", color:"#333", border:"1.5px solid #ddd", borderRadius:7, padding:"0 12px", height:36, fontWeight:500, fontSize:13, cursor:"pointer", whiteSpace:"nowrap" };
const btnDanger = { background:"#dc2626", color:"#fff", border:"none", borderRadius:8, padding:"0 18px", height:38, fontWeight:700, fontSize:14, cursor:"pointer" };
const inp = (err) => ({ width:"100%", height:44, border:`1.5px solid ${err?"#dc2626":"#ddd"}`, borderRadius:7, padding:"0 12px", fontSize:'16px', outline:"none", boxSizing:"border-box" });
const lbl = { display:"block", fontSize:12, fontWeight:600, color:"#444", marginBottom:5 };
const fg = { marginBottom:14 };

// ─── Brevo Email ─────────────────────────────────────────────────────────────
async function sendBrevoEmail(toEmail, toName, subject, htmlContent) {
  if (!toEmail) return;
  try {
    const res = await fetch('/send-email', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ to: toEmail, toName, subject, html: htmlContent })
    });
    const data = await res.json();
    return { success: data.success };
  } catch(e) {
    return { success: false };
  }
}

// ─── Toast ────────────────────────────────────────────────────────────────────
function Toast({ msg, type, onClose }) {
  useEffect(() => { const t = setTimeout(onClose, 1500); return () => clearTimeout(t); }, [onClose]);
  return (
    <>
      <style>{`
        @keyframes popIn { 0% { opacity:0; transform:scale(0.6); } 70% { transform:scale(1.05); } 100% { opacity:1; transform:scale(1); } }
        @keyframes checkIn { 0% { transform:scale(0) rotate(-90deg); opacity:0; } 60% { transform:scale(1.2) rotate(10deg); } 100% { transform:scale(1) rotate(0deg); opacity:1; } }
        @keyframes fadeOverlay { 0% { opacity:0; } 100% { opacity:1; } }
      `}</style>
      <div style={{ position:"fixed", inset:0, zIndex:99999, display:"flex", alignItems:"center", justifyContent:"center", background:"rgba(0,0,0,0.35)", pointerEvents:"none", animation:"fadeOverlay 0.15s ease" }}>
        <div style={{ background:"#fff", borderRadius:24, padding:"32px 44px", display:"flex", flexDirection:"column", alignItems:"center", gap:16, boxShadow:"0 24px 80px rgba(0,0,0,0.25)", animation:"popIn 0.25s cubic-bezier(0.34, 1.56, 0.64, 1)", minWidth:220, maxWidth:300, textAlign:"center" }}>
          <div style={{ width:72, height:72, borderRadius:"50%", background: type==="error" ? "#fef2f2" : "#f0fdf4", border:`4px solid ${type==="error" ? "#dc2626" : "#22c55e"}`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:38, color: type==="error" ? "#dc2626" : "#22c55e", fontWeight:700, animation:"checkIn 0.3s cubic-bezier(0.34, 1.56, 0.64, 1) 0.1s both" }}>
            {type === "error" ? "✕" : "✓"}
          </div>
          <p style={{ color: type==="error" ? "#dc2626" : "#16a34a", fontWeight:700, fontSize:17, margin:0, lineHeight:1.4 }}>{msg}</p>
        </div>
      </div>
    </>
  );
}

// ─── Modal ────────────────────────────────────────────────────────────────────
function Modal({ title, onClose, children, footer, maxW=520, zIndex=1000 }) {
  const isMobile = window.innerWidth < 768;
  return (
    <div
      style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex, display:"flex", alignItems: isMobile ? "flex-end" : "center", justifyContent:"center", padding: isMobile ? 0 : "1rem" }}
      onPointerDown={e => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div
        style={{ background:"#fff", borderRadius: isMobile ? '16px 16px 0 0' : 12, width:"100%", maxWidth: isMobile ? '100%' : maxW, overflow:"hidden", maxHeight:"90vh", display:"flex", flexDirection:"column" }}
        onPointerDown={e => e.stopPropagation()}
      >
        <div style={{ background:"#111", color:"#fff", padding:"14px 18px", display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0 }}>
          <span style={{ fontWeight:700, fontSize:15 }}>{title}</span>
          <button type="button" onClick={onClose} style={{ background:"none", border:"none", color:"#fff", fontSize:20, cursor:"pointer" }}>✕</button>
        </div>
        <div style={{ padding:"18px", overflowY:"auto", flex:1, WebkitOverflowScrolling:"touch" }}>{children}</div>
        {footer && <div style={{ padding: isMobile ? "12px 16px 24px" : "0 18px 18px", display:"flex", gap:8, justifyContent:"flex-end", flexShrink:0, background:'#fff' }}>{footer}</div>}
      </div>
    </div>
  );
}

function ConfirmModal({ title, msg, onOk, onCancel, okLabel="Confirmer", danger=false }) {
  return (
    <Modal title={title} onClose={onCancel} maxW={400} footer={[
      <button key="c" type="button" onPointerDown={onCancel} style={{...btnSecondary, touchAction:"manipulation"}}>Annuler</button>,
      <button key="o" type="button" onPointerDown={onOk} style={{...(danger?btnDanger:btnPrimary), touchAction:"manipulation"}}>{okLabel}</button>
    ]}>
      <p style={{ fontSize:14, lineHeight:1.65, margin:0 }}>{msg}</p>
    </Modal>
  );
}

// ─── Login Page ───────────────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function handleLogin(e) {
    e.preventDefault();
    setLoading(true);
    setError("");
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) { setError("Email ou mot de passe incorrect."); setLoading(false); return; }
    onLogin();
  }

  return (
    <div style={{ minHeight:"100vh", background:"#f8f8f8", display:"flex", alignItems:"center", justifyContent:"center", padding:"1rem" }}>
      <div style={{ background:"#fff", borderRadius:14, border:"1.5px solid #e5e5e5", padding:"2.5rem 2rem", width:"100%", maxWidth:380, boxShadow:"0 4px 30px rgba(0,0,0,0.08)" }}>
        <div style={{ textAlign:"center", marginBottom:"2rem" }}>
          <div style={{ background:"#111", borderRadius:10, display:"inline-block", padding:"10px 20px", marginBottom:16 }}>
            <span style={{ color:G, fontWeight:700, fontSize:20, letterSpacing:2 }}>TED</span>
          </div>
          <h1 style={{ fontSize:18, fontWeight:700, color:"#111", margin:0 }}>Fichier Clients</h1>
          <p style={{ fontSize:13, color:"#999", marginTop:6 }}>Connectez-vous pour accéder au CRM</p>
        </div>
        <form onSubmit={handleLogin}>
          <div style={fg}>
            <label style={lbl}>Adresse email</label>
            <input style={inp(false)} type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="votre@email.fr" required />
          </div>
          <div style={fg}>
            <label style={lbl}>Mot de passe</label>
            <input style={inp(false)} type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••" required />
          </div>
          {error && <p style={{ fontSize:12, color:"#dc2626", marginBottom:12, textAlign:"center" }}>{error}</p>}
          <button type="submit" style={{ ...btnPrimary, width:"100%", height:44, fontSize:15 }} disabled={loading}>
            {loading ? "Connexion..." : "Se connecter"}
          </button>
        </form>
      </div>
    </div>
  );
}

// ─── Client Form ──────────────────────────────────────────────────────────────
function ClientForm({ initial, onSave, onCancel, existingClients }) {
  const isEdit = !!initial?.id;
  const isMobile = window.innerWidth < 768;
  const [form, setForm] = useState({
    genre: initial?.genre || "Non renseigné",
    nom: initial?.nom || "",
    prenom: initial?.prenom || "",
    tel: initial?.tel || "",
    mail: initial?.mail || "",
    commentaire: initial?.commentaire || "",
    entreprise: initial?.entreprise || ""
  });
  const [errors, setErrors] = useState({});
  const [dupWarn, setDupWarn] = useState(null);
  const [success, setSuccess] = useState(false);
  const [dupClient, setDupClient] = useState(null);

  function set(k, v) { setForm(f => ({ ...f, [k]: v })); setErrors(e => ({ ...e, [k]: "" })); }
  function handleTel(v) {
    const clean = v.replace(/\D/g, "").slice(0, 10);
    set("tel", clean);
    if (clean.length === 10) {
      const others = existingClients.filter(c => !isEdit || c.id !== initial?.id);
      const found = others.find(c => c.tel === clean);
      if (found) { setDupClient(found); } else { setDupClient(null); }
    } else {
      setDupClient(null);
    }
  }

  function validate() {
    const e = {};
    if (!form.genre || form.genre === "Non renseigné") {
      e.genre = "Veuillez sélectionner un genre.";
    }
    if (form.genre === "Entreprise") {
      if (!form.entreprise || !form.entreprise.trim()) e.entreprise = "Le nom de l'entreprise est obligatoire.";
    } else {
      if (!form.nom.trim()) e.nom = "Le nom est obligatoire.";
      if (!form.prenom.trim()) e.prenom = "Le prénom est obligatoire.";
    }
    if (!form.tel || !form.tel.trim()) e.tel = "Le téléphone est obligatoire.";
    if (form.tel && !/^\d{10}$/.test(form.tel)) e.tel = "Le numéro doit contenir uniquement 10 chiffres.";
    if (form.mail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(form.mail)) e.mail = "Adresse mail invalide.";
    return e;
  }

  function checkDup() {
    const others = existingClients.filter(c => !isEdit || c.id !== initial.id);
    if (form.tel && others.some(c => c.tel === form.tel)) return `Le téléphone ${form.tel} est déjà utilisé par un autre client.`;
    if (form.mail && form.mail.trim() && others.some(c => c.mail && c.mail.toLowerCase() === form.mail.toLowerCase())) return `L'adresse ${form.mail} est déjà utilisée.`;
    return null;
  }

  function doSave() {
    const saved = {
      ...(initial || {}),
      id: initial?.id,
      genre: form.genre,
      nom: capitalize(form.nom.trim()),
      prenom: capitalize(form.prenom.trim()),
      tel: form.tel,
      mail: form.mail.trim().toLowerCase(),
      commentaire: form.commentaire.trim(),
      entreprise: form.entreprise.trim(),
      created_at: initial?.created_at || new Date().toISOString()
    };
    onSave(saved);
    setDupWarn(null);
  }

  function handleSubmit() {
    const e = validate();
    if (Object.keys(e).length > 0) { setErrors(e); return; }
    const dup = checkDup();
    if (dup) { setDupWarn(dup); return; }
    setSuccess(true);
    setTimeout(() => { doSave(); }, 800);
  }

  const inputStyle = (err) => ({
    width: "100%", height: 44, border: `1.5px solid ${err ? "#dc2626" : "#ddd"}`,
    borderRadius: 7, padding: "0 12px", fontSize: 16, outline: "none", boxSizing: "border-box"
  });
  const labelStyle = { display: "block", fontSize: 12, fontWeight: 600, color: "#444", marginBottom: 5 };
  const fieldGroup = { marginBottom: 14 };

  return (
    <>
      {dupWarn && (
        <ConfirmModal
          title="Doublon détecté"
          msg={`Attention : ${dupWarn} Voulez-vous tout de même continuer ?`}
          onOk={() => { setDupWarn(null); doSave(); }}
          onCancel={() => setDupWarn(null)}
          okLabel="Ajouter quand même"
        />
      )}
      <Modal
        title={isEdit ? "Modifier le client" : "Ajouter un client"}
        onClose={onCancel}
        footer={[
          <button key="c" type="button" onPointerDown={onCancel} style={{
            background: "#fff", border: "1.5px solid #ddd", borderRadius: 8,
            padding: "0 14px", height: 48, fontWeight: 500, fontSize: 15,
            cursor: "pointer", flex: 1, touchAction: "manipulation"
          }}>Annuler</button>,
          <button key="s" type="button" onPointerDown={dupClient ? undefined : handleSubmit} disabled={!!dupClient} style={{
            background: dupClient ? "#ddd" : (success ? "#22c55e" : "#E8C547"),
            color: dupClient ? "#999" : (success ? "#fff" : "#111"),
            border: "none", borderRadius: 12,
            height: 52, fontWeight: 700, fontSize: 16,
            cursor: dupClient ? "not-allowed" : "pointer", flex: 2, touchAction: "manipulation",
            transition: "all 0.4s cubic-bezier(0.34, 1.56, 0.64, 1)",
            transform: success ? "scale(1.05)" : "scale(1)",
            display: "flex", alignItems: "center", justifyContent: "center", gap: 8,
            boxShadow: success ? "0 4px 20px rgba(34,197,94,0.4)" : "none"
          }}>
            {success ? (<><span style={{ display:"inline-block", animation:"scaleIn 0.3s cubic-bezier(0.34, 1.56, 0.64, 1)" }}>✓</span>Enregistré !</>) : dupClient ? "⚠ Client existant" : isEdit ? "Modifier" : "Enregistrer"}
          </button>
        ]}
      >
        <div style={fieldGroup}>
          <label style={labelStyle}>Genre <span style={{color:"#dc2626"}}>*</span></label>
          <select style={{ width:"100%", height:44, border:`1.5px solid ${errors.genre ? "#dc2626" : "#ddd"}`, borderRadius:7, padding:"0 12px", fontSize:16, background:"#fff", outline:"none" }}
            value={form.genre} onChange={e => set("genre", e.target.value)}>
            <option value="Non renseigné">-- Sélectionner --</option>
            {GENRES.filter(g => g !== "Non renseigné").map(g => <option key={g}>{g}</option>)}
          </select>
          {errors.genre && <p style={{fontSize:12, color:"#dc2626", marginTop:4}}>{errors.genre}</p>}
        </div>

        <div style={fieldGroup}>
          <label style={labelStyle}>Téléphone <span style={{ color: "#dc2626" }}>*</span></label>
          <input style={inputStyle(errors.tel)} value={form.tel}
            onChange={e => handleTel(e.target.value)} inputMode="numeric" placeholder="0612345678" maxLength={10} />
          {errors.tel && <p style={{ fontSize: 12, color: "#dc2626", marginTop: 4 }}>{errors.tel}</p>}
          {dupClient && (
            <div style={{ background:"#fef2f2", border:"2px solid #dc2626", borderRadius:10, padding:"12px 14px", marginTop:8, display:"flex", alignItems:"center", gap:10 }}>
              <span style={{fontSize:24}}>⚠️</span>
              <div>
                <p style={{fontWeight:700, color:"#dc2626", fontSize:14, margin:0}}>Client déjà existant !</p>
                <p style={{fontSize:13, color:"#333", margin:"4px 0 0"}}><strong>{dupClient.prenom} {dupClient.nom}{dupClient.entreprise ? ` — ${dupClient.entreprise}` : ""}</strong></p>
                <p style={{fontSize:12, color:"#666", margin:"2px 0 0"}}>📞 {dupClient.tel}{dupClient.mail ? ` · ${dupClient.mail}` : ""}</p>
              </div>
            </div>
          )}
        </div>

        {form.genre === "Entreprise" && (
          <div style={fieldGroup}>
            <label style={labelStyle}>Nom de l'entreprise <span style={{ color: "#dc2626" }}>*</span></label>
            <input style={inputStyle(errors.entreprise)} value={form.entreprise}
              onChange={e => set("entreprise", e.target.value)} placeholder="Nom de l'entreprise" />
            {errors.entreprise && <p style={{ fontSize: 12, color: "#dc2626", marginTop: 4 }}>{errors.entreprise}</p>}
          </div>
        )}

        <div style={fieldGroup}>
          <label style={labelStyle}>
            Nom {form.genre !== "Entreprise" ? <span style={{ color: "#dc2626" }}>*</span> : <span style={{ color: "#999", fontSize: 11 }}> (facultatif)</span>}
          </label>
          <input style={inputStyle(errors.nom)} value={form.nom}
            onChange={e => set("nom", e.target.value)} placeholder="Dupont" />
          {errors.nom && <p style={{ fontSize: 12, color: "#dc2626", marginTop: 4 }}>{errors.nom}</p>}
        </div>

        <div style={fieldGroup}>
          <label style={labelStyle}>
            Prénom {form.genre !== "Entreprise" ? <span style={{ color: "#dc2626" }}>*</span> : <span style={{ color: "#999", fontSize: 11 }}> (facultatif)</span>}
          </label>
          <input style={inputStyle(errors.prenom)} value={form.prenom}
            onChange={e => set("prenom", e.target.value)} placeholder="Jean" />
          {errors.prenom && <p style={{ fontSize: 12, color: "#dc2626", marginTop: 4 }}>{errors.prenom}</p>}
        </div>

        <div style={fieldGroup}>
          <label style={labelStyle}>Mail</label>
          <input style={inputStyle(errors.mail)} value={form.mail}
            onChange={e => set("mail", e.target.value)} placeholder="exemple@mail.fr" type="email" />
          {errors.mail && <p style={{ fontSize: 12, color: "#dc2626", marginTop: 4 }}>{errors.mail}</p>}
        </div>

        <div style={fieldGroup}>
          <label style={labelStyle}>Commentaire</label>
          <textarea style={{ width: "100%", border: "1.5px solid #ddd", borderRadius: 7, padding: "10px 12px", fontSize: 16, outline: "none", boxSizing: "border-box", resize: "vertical", minHeight: 80 }}
            value={form.commentaire} onChange={e => set("commentaire", e.target.value)} placeholder="Notes sur ce client…" />
        </div>

        {isEdit && <p style={{ fontSize: 12, color: "#999", margin: 0 }}>Date d'ajout : {formatDate(initial.created_at)} — non modifiable</p>}
      </Modal>
    </>
  );
}

// ─── Import Modal ─────────────────────────────────────────────────────────────
function ImportModal({ onImport, onCancel, existingClients }) {
  const [parsed, setParsed] = useState(null);
  const [dups, setDups] = useState([]);

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const rows = parseCSV(ev.target.result);
      const mapped = rows.map(mapImportRow);
      const dupList = mapped.filter(c => (c.tel && existingClients.some(ex=>ex.tel===c.tel)) || (c.mail && existingClients.some(ex=>ex.mail&&ex.mail.toLowerCase()===c.mail.toLowerCase())));
      setParsed(mapped); setDups(dupList);
    };
    reader.readAsText(file, "UTF-8");
  }

  return (
    <Modal title="Importer des clients (CSV)" onClose={onCancel} maxW={560} footer={parsed ? [
      <button key="c" onClick={onCancel} style={btnSecondary}>Annuler</button>,
      <button key="i" onClick={()=>onImport(parsed)} style={btnPrimary}>Importer {parsed.length} client(s)</button>
    ] : null}>
      {!parsed && (
        <>
          <p style={{ fontSize:13, color:"#555", marginBottom:12 }}>Importez un fichier CSV avec les colonnes : Genre, Nom, Prénom, Téléphone, Mail, Date d'ajout, Commentaire.</p>
          <input type="file" accept=".csv,.txt" onChange={handleFile} style={{ fontSize:13 }} />
        </>
      )}
      {parsed && (
        <>
          {dups.length > 0 && <div style={{ background:"#fffbeb", border:"1.5px solid #fbbf24", borderRadius:8, padding:"10px 14px", fontSize:13, color:"#92400e", marginBottom:12 }}>⚠ {dups.length} doublon(s) potentiel(s) détecté(s).</div>}
          <p style={{ fontWeight:600, marginBottom:8 }}>{parsed.length} client(s) détecté(s)</p>
          <div style={{ maxHeight:180, overflowY:"auto", fontSize:12, border:"1px solid #eee", borderRadius:6, padding:"8px" }}>
            {parsed.map((c,i) => <div key={i} style={{ padding:"3px 0", borderBottom:"1px solid #f0f0f0" }}><span style={{fontWeight:600}}>{c.nom} {c.prenom}</span><span style={{color:"#999",marginLeft:8}}>{c.tel} {c.mail}</span></div>)}
          </div>
          <p style={{ fontSize:11, color:"#999", marginTop:8 }}>Les clients existants ne seront pas écrasés.</p>
        </>
      )}
    </Modal>
  );
}

// ─── Corbeille Modal ──────────────────────────────────────────────────────────
function CorbeilleModal({ onClose, showToast }) {
  const [deleted, setDeleted] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    supabase.from("clients").select("*").not("deleted_at", "is", null).order("deleted_at", { ascending: false })
      .then(({ data }) => { setDeleted(data || []); setLoading(false); });
  }, []);

  async function restore(id) {
    await supabase.from("clients").update({ deleted_at: null, deleted_by: null }).eq("id", id);
    setDeleted(prev => prev.filter(c => c.id !== id));
    showToast("Client restauré ✓");
  }

  async function deletePermanently(id) {
    if (!window.confirm("Supprimer définitivement ce client ? Cette action est irréversible.")) return;
    await supabase.from("clients").delete().eq("id", id);
    setDeleted(prev => prev.filter(c => c.id !== id));
    showToast("Client supprimé définitivement");
  }

  async function emptyTrash() {
    if (!window.confirm(`Vider la corbeille ? ${deleted.length} client(s) seront supprimés définitivement.`)) return;
    await supabase.from("clients").delete().not("deleted_at", "is", null);
    setDeleted([]);
    showToast("Corbeille vidée ✓");
  }

  const footer = deleted.length > 0 ? [
    <button key="empty" onClick={emptyTrash} style={{ background:"#dc2626", color:"#fff", border:"none", borderRadius:8, padding:"0 16px", height:44, fontWeight:700, fontSize:13, cursor:"pointer" }}>🗑 Vider la corbeille ({deleted.length})</button>
  ] : null;

  return (
    <Modal title="🗑 Corbeille" onClose={onClose} maxW={600} footer={footer}>
      {loading && <p style={{ textAlign:"center", color:"#999", padding:"2rem" }}>Chargement...</p>}
      {!loading && deleted.length === 0 && (
        <div style={{ textAlign:"center", padding:"3rem" }}>
          <div style={{ fontSize:48, marginBottom:12 }}>🗑</div>
          <p style={{ color:"#bbb", fontSize:15 }}>La corbeille est vide</p>
        </div>
      )}
      {!loading && deleted.map(c => (
        <div key={c.id} style={{ background:"#fff", border:"1.5px solid #f0f0f0", borderRadius:12, padding:"14px", marginBottom:10, display:"flex", alignItems:"center", justifyContent:"space-between", gap:12 }}>
          <div style={{ flex:1 }}>
            <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:4 }}>
              <span style={badge(c.genre)}>{c.genre}</span>
              <span style={{ fontWeight:700, fontSize:15 }}>{c.genre === "Entreprise" ? (c.entreprise||c.nom) : `${c.nom} ${c.prenom}`}</span>
            </div>
            {c.tel && <p style={{ fontSize:13, color:"#555", margin:"2px 0" }}>📞 {c.tel}</p>}
            <p style={{ fontSize:11, color:"#bbb", margin:"4px 0 0" }}>
              Supprimé le {new Date(c.deleted_at).toLocaleDateString("fr-FR")} à {new Date(c.deleted_at).toLocaleTimeString("fr-FR", { hour:"2-digit", minute:"2-digit" })}
              {c.deleted_by ? ` par ${c.deleted_by}` : ""}
            </p>
          </div>
          <div style={{ display:"flex", gap:8, flexShrink:0 }}>
            <button onClick={()=>restore(c.id)} style={{ background:"#f0fdf4", border:"1.5px solid #22c55e", borderRadius:8, padding:"0 12px", height:36, fontWeight:600, fontSize:12, cursor:"pointer", color:"#16a34a" }}>↩ Restaurer</button>
            <button onClick={()=>deletePermanently(c.id)} style={{ background:"#fef2f2", border:"1.5px solid #dc2626", borderRadius:8, padding:"0 12px", height:36, fontWeight:600, fontSize:12, cursor:"pointer", color:"#dc2626" }}>✕ Supprimer</button>
          </div>
        </div>
      ))}
    </Modal>
  );
}

// ─── Mobile hook ──────────────────────────────────────────────────────────────
function useIsMobile() {
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768);
  useEffect(() => {
    const handler = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener('resize', handler);
    return () => window.removeEventListener('resize', handler);
  }, []);
  return isMobile;
}

// ─── Réservations Page ────────────────────────────────────────────────────────
const FORM_URL = "https://ted-crm.pages.dev/reserver.html";

const OCCASIONS = ["Anniversaire","Saint-Valentin","Repas d'affaires","Mariage","Fiançailles","Autre"];
const HEURES_MIDI = ["12:00","12:15","12:30","12:45","13:00","13:15","13:30","13:45","14:00"];
const HEURES_SOIR = ["19:00","19:15","19:30","19:45","20:00","20:15","20:30","20:45","21:00","21:15","21:30"];

// Génère les 30 prochains jours comme options de select
function buildDateOptions() {
  const opts = [];
  const joursL = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
  const moisC = ['jan','fév','mar','avr','mai','jun','jul','aoû','sep','oct','nov','déc'];
  const today = new Date();
  for (let i = 0; i < 30; i++) {
    const d = new Date(today);
    d.setDate(today.getDate() + i);
    const iso = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
    let label;
    if (i === 0) label = `Aujourd'hui ${d.getDate()} ${moisC[d.getMonth()]}`;
    else if (i === 1) label = `Demain ${d.getDate()} ${moisC[d.getMonth()]}`;
    else label = `${joursL[d.getDay()]} ${d.getDate()} ${moisC[d.getMonth()]}`;
    opts.push({ iso, label });
  }
  return opts;
}

function AddResaModal({ onClose, onSaved, showToast, user }) {
  const DATE_OPTS = useMemo(() => buildDateOptions(), []);

  const [tel, setTel] = useState('');
  const [clientFound, setClientFound] = useState(null); // objet client ou null
  const [lookingUp, setLookingUp] = useState(false);
  const [genre, setGenre] = useState('');
  const [prenom, setPrenom] = useState('');
  const [nom, setNom] = useState('');
  const [entreprise, setEntreprise] = useState('');
  const [email, setEmail] = useState('');
  const [dateIso, setDateIso] = useState(DATE_OPTS[0].iso);
  const [service, setService] = useState('soir');
  const [heure, setHeure] = useState('');
  const [nbPersonnes, setNbPersonnes] = useState(2);
  const [occasion, setOccasion] = useState('');
  const [saving, setSaving] = useState(false);
  const [heureError, setHeureError] = useState(false);

  const heures = service === 'midi' ? HEURES_MIDI : HEURES_SOIR;

  // Recherche automatique dès 10 chiffres
  async function handleTelChange(val) {
    setTel(val);
    setClientFound(null);
    const digits = val.replace(/\D/g, '');
    if (digits.length < 10) return;
    setLookingUp(true);
    const telNorm = val.replace(/[\s.\-()]/g,'').replace(/^0/,'+33');
    const { data } = await supabase
      .from('clients')
      .select('id,prenom,nom,mail,genre,entreprise,tel_normalise')
      .or(`tel_normalise.eq.${telNorm},tel.eq.${val.trim()}`)
      .maybeSingle();
    setLookingUp(false);
    if (data) {
      setClientFound(data);
      setPrenom(data.prenom || '');
      setNom(data.nom || '');
      setEmail(data.mail || '');
      setGenre(data.genre || '');
      setEntreprise(data.entreprise || '');
    }
  }

  async function handleSave() {
    if (!tel.trim()) { showToast('Téléphone requis', 'error'); return; }
    if (!genre) { showToast('Genre requis', 'error'); return; }
    if (!prenom.trim()) { showToast('Prénom requis', 'error'); return; }
    if (!nom.trim()) { showToast('Nom requis', 'error'); return; }
    if (genre === 'Entreprise' && !entreprise.trim()) { showToast('Nom d\'entreprise requis', 'error'); return; }
    if (!heure) { setHeureError(true); return; }
    setSaving(true);
    const telNorm = tel.replace(/[\s.\-()]/g,'').replace(/^0/,'+33');
    let clientId = clientFound?.id || null;
    if (!clientId) {
      const { data: newClient, error: errClient } = await supabase.from('clients').insert({
        prenom: capitalize(prenom.trim()),
        nom: capitalize(nom.trim()),
        tel: tel.trim(),
        mail: email.trim() || null,
        genre,
        entreprise: genre === 'Entreprise' ? entreprise.trim() : null,
        tel_normalise: telNorm,
        source: 'manuel',
      }).select('id').single();
      if (errClient) { setSaving(false); showToast('Erreur création client', 'error'); return; }
      clientId = newClient.id;
    }
    const { error } = await supabase.from('reservations').insert({
      client_id: clientId,
      date: dateIso,
      service,
      heure: heure || null,
      nb_personnes: nbPersonnes,
      occasion: occasion || null,
      statut: 'attente',
      source: 'manuel',
    });
    setSaving(false);
    if (error) { showToast('Erreur lors de la création', 'error'); return; }
    showToast('Réservation créée ✓');
    onSaved();
    onClose();
  }

  const btnSvc = (s) => ({
    flex: 1, height: 42, border: `1.5px solid ${service === s ? '#111' : '#eee'}`,
    borderRadius: 8, background: service === s ? '#111' : '#f8f8f8',
    color: service === s ? '#fff' : '#666', fontWeight: 700, fontSize: 14, cursor: 'pointer'
  });

  const btnGenre = (g) => ({
    flex: 1, height: 42, border: `1.5px solid ${genre === g ? '#111' : '#eee'}`,
    borderRadius: 8, background: genre === g ? '#111' : '#f8f8f8',
    color: genre === g ? G : '#666', fontWeight: 700, fontSize: 13, cursor: 'pointer'
  });

  return (
    <Modal title="Ajouter une réservation" onClose={onClose} footer={[
      <button key="cancel" onClick={onClose} style={btnSecondary}>Annuler</button>,
      <button key="save" onClick={handleSave} disabled={saving || !heure} style={{ ...btnPrimary, opacity: (saving || !heure) ? 0.5 : 1 }}>{saving ? 'Enregistrement…' : 'Enregistrer'}</button>
    ]}>
      <div style={{ display:'flex', flexDirection:'column', gap:14 }}>

        {/* Bandeau info statut */}
        <div style={{ background:'#fffbeb', border:'1.5px solid #fbbf24', borderRadius:10, padding:'10px 14px', fontSize:13, color:'#92400e', display:'flex', alignItems:'center', gap:8 }}>
          <span style={{ fontSize:16 }}>⏳</span>
          <span>Cette réservation sera créée comme <strong>demande en attente</strong> — vous pourrez l'accepter depuis la page Réservations.</span>
        </div>

        {/* 1. Téléphone */}
        <div>
          <label style={lbl}>Téléphone *</label>
          <div style={{ position:'relative' }}>
            <input value={tel} onChange={e=>handleTelChange(e.target.value)} placeholder="06 12 34 56 78" type="tel" style={inp(false)} />
            {lookingUp && <span style={{ position:'absolute', right:12, top:'50%', transform:'translateY(-50%)', fontSize:12, color:'#888' }}>Recherche…</span>}
          </div>
          {clientFound && (
            <div style={{ marginTop:6, background:'#f0fdf4', border:'1.5px solid #22c55e', borderRadius:8, padding:'8px 12px', fontSize:13, color:'#166534', fontWeight:600, display:'flex', alignItems:'center', gap:6 }}>
              ✓ Client trouvé — {clientFound.prenom} {clientFound.nom}
            </div>
          )}
        </div>

        {/* 2. Genre */}
        <div>
          <label style={lbl}>Vous êtes *</label>
          <div style={{ display:'flex', gap:8 }}>
            {['Homme','Femme','Entreprise'].map(g => (
              <button key={g} onClick={()=>setGenre(g)} style={btnGenre(g)}>{g}</button>
            ))}
          </div>
        </div>

        {/* Nom d'entreprise si Entreprise */}
        {genre === 'Entreprise' && (
          <div>
            <label style={lbl}>Nom de l'entreprise *</label>
            <input value={entreprise} onChange={e=>setEntreprise(e.target.value)} placeholder="Nom de l'entreprise" style={inp(false)} />
          </div>
        )}

        {/* 3. Prénom + Nom */}
        <div style={{ display:'flex', gap:8 }}>
          <div style={{ flex:1 }}>
            <label style={lbl}>Prénom *</label>
            <input value={prenom} onChange={e=>setPrenom(e.target.value)} placeholder="Jean" style={inp(false)} />
          </div>
          <div style={{ flex:1 }}>
            <label style={lbl}>Nom *</label>
            <input value={nom} onChange={e=>setNom(e.target.value)} placeholder="Dupont" style={inp(false)} />
          </div>
        </div>

        {/* 4. Email */}
        <div>
          <label style={lbl}>Email</label>
          <input value={email} onChange={e=>setEmail(e.target.value)} placeholder="optionnel" type="email" style={inp(false)} />
        </div>

        {/* 5. Date — select 30 jours */}
        <div>
          <label style={lbl}>Date *</label>
          <select value={dateIso} onChange={e=>setDateIso(e.target.value)} style={{ width:'100%', height:44, border:'1.5px solid #ddd', borderRadius:7, padding:'0 12px', fontSize:15, background:'#fff', outline:'none', cursor:'pointer' }}>
            {DATE_OPTS.map(o => <option key={o.iso} value={o.iso}>{o.label}</option>)}
          </select>
        </div>

        {/* 6. Service */}
        <div>
          <label style={lbl}>Service *</label>
          <div style={{ display:'flex', gap:8 }}>
            <button style={btnSvc('midi')} onClick={()=>{ setService('midi'); setHeure(''); setHeureError(false); }}>🌞 Midi</button>
            <button style={btnSvc('soir')} onClick={()=>{ setService('soir'); setHeure(''); setHeureError(false); }}>🌙 Soir</button>
          </div>
        </div>

        {/* 7. Heure */}
        <div>
          <label style={lbl}>Heure *</label>
          <div style={{ display:'flex', flexWrap:'wrap', gap:6 }}>
            {heures.map(h => (
              <button key={h} onClick={()=>{ setHeure(heure===h?'':h); setHeureError(false); }} style={{ padding:'6px 12px', borderRadius:8, border:`1.5px solid ${heure===h?'#111':heureError?'#dc2626':'#eee'}`, background:heure===h?'#111':'#f8f8f8', color:heure===h?'#fff':'#555', fontWeight:600, fontSize:13, cursor:'pointer' }}>{h}</button>
            ))}
          </div>
          {heureError && <p style={{ fontSize:12, color:'#dc2626', marginTop:6, margin:'6px 0 0' }}>* Sélectionnez un créneau horaire</p>}
        </div>

        {/* 8. Nb personnes */}
        <div>
          <label style={lbl}>Nombre de personnes *</label>
          <div style={{ display:'flex', alignItems:'center', gap:12 }}>
            <button onClick={()=>setNbPersonnes(n=>Math.max(1,n-1))} style={{ width:40, height:40, borderRadius:8, border:'1.5px solid #eee', background:'#f8f8f8', fontSize:20, cursor:'pointer', fontWeight:700 }}>−</button>
            <input type="number" value={nbPersonnes} min={1} max={50} onChange={e=>setNbPersonnes(Math.max(1,parseInt(e.target.value)||1))} style={{ width:70, height:40, border:'1.5px solid #ddd', borderRadius:8, textAlign:'center', fontSize:16, fontWeight:700, outline:'none' }} />
            <button onClick={()=>setNbPersonnes(n=>Math.min(50,n+1))} style={{ width:40, height:40, borderRadius:8, border:'1.5px solid #eee', background:'#f8f8f8', fontSize:20, cursor:'pointer', fontWeight:700 }}>+</button>
          </div>
        </div>

        {/* 9. Occasion */}
        <div>
          <label style={lbl}>Occasion</label>
          <select value={occasion} onChange={e=>setOccasion(e.target.value)} style={{ width:'100%', height:44, border:'1.5px solid #ddd', borderRadius:7, padding:'0 12px', fontSize:15, background:'#fff', outline:'none', cursor:'pointer' }}>
            <option value="">— Aucune —</option>
            {OCCASIONS.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>

      </div>
    </Modal>
  );
}

const MOTIFS_REFUS = ["Complet","Fermé","Horaire indispo","Groupe trop grand","Autre"];

const statutBadge = (s) => {
  const map = {
    attente:   { bg:'#fffbeb', color:'#92400e', label:'En attente' },
    rappeler:  { bg:'#fff7ed', color:'#9a3412', label:'À rappeler' },
    confirmee: { bg:'#f0fdf4', color:'#166534', label:'Confirmée' },
    refusee:   { bg:'#fef2f2', color:'#991b1b', label:'Refusée' },
    annulee:   { bg:'#f3f4f6', color:'#374151', label:'Annulée' },
    venue:     { bg:'#dcfce7', color:'#14532d', label:'Venue' },
    absente:   { bg:'#fef2f2', color:'#7f1d1d', label:'Absente' },
  };
  const s2 = map[s] || { bg:'#f3f4f6', color:'#374151', label: s };
  return <span style={{ display:'inline-block', fontSize:11, fontWeight:700, borderRadius:99, padding:'3px 9px', background:s2.bg, color:s2.color }}>{s2.label}</span>;
};

function fmtResaDate(iso) {
  if (!iso) return '';
  const d = new Date(iso + 'T00:00:00');
  const jours = ['Dim','Lun','Mar','Mer','Jeu','Ven','Sam'];
  const mois = ['jan','fév','mar','avr','mai','jun','jul','aoû','sep','oct','nov','déc'];
  return `${jours[d.getDay()]} ${d.getDate()} ${mois[d.getMonth()]}`;
}

function RefusModal({ onConfirm, onCancel }) {
  const [motif, setMotif] = useState(MOTIFS_REFUS[0]);
  const [autre, setAutre] = useState('');
  function confirm() {
    const raison = motif === 'Autre' ? (autre.trim() || 'Autre') : motif;
    onConfirm(raison);
  }
  return (
    <Modal title="Refuser la réservation" onClose={onCancel} maxW={400}
      footer={[
        <button key="c" type="button" onClick={onCancel} style={{...btnSecondary}}>Annuler</button>,
        <button key="o" type="button" onClick={confirm} style={{...btnDanger}}>Refuser</button>
      ]}>
      <p style={{ fontSize:13, color:'#555', marginBottom:14 }}>Sélectionnez le motif du refus :</p>
      {MOTIFS_REFUS.map(m => (
        <label key={m} style={{ display:'flex', alignItems:'center', gap:10, padding:'10px 12px', borderRadius:8, marginBottom:6, cursor:'pointer', background: motif===m ? '#fef2f2' : '#f8f8f8', border:`1.5px solid ${motif===m?'#dc2626':'#eee'}` }}>
          <input type="radio" name="motif" value={m} checked={motif===m} onChange={()=>setMotif(m)} style={{ accentColor:'#dc2626' }} />
          <span style={{ fontSize:14, fontWeight: motif===m?700:400, color: motif===m?'#dc2626':'#333' }}>{m}</span>
        </label>
      ))}
      {motif === 'Autre' && (
        <input value={autre} onChange={e=>setAutre(e.target.value)} placeholder="Précisez le motif…"
          style={{ width:'100%', height:42, border:'1.5px solid #ddd', borderRadius:8, padding:'0 12px', fontSize:14, outline:'none', marginTop:4 }} />
      )}
    </Modal>
  );
}

function AccepterModal({ resa, onConfirm, onCancel }) {
  const c = resa.clients || {};
  const nom = c.entreprise ? c.entreprise : `${c.prenom || ''} ${c.nom || ''}`.trim();
  return (
    <Modal title="✓ Confirmer la réservation" onClose={onCancel} maxW={420} zIndex={3000}
      footer={[
        <button key="c" type="button" onClick={onCancel} style={{...btnSecondary}}>Annuler</button>,
        <button key="o" type="button" onClick={onConfirm} style={{ background:'#16a34a', color:'#fff', border:'none', borderRadius:8, padding:'0 20px', height:40, fontWeight:700, fontSize:14, cursor:'pointer' }}>✓ Confirmer</button>
      ]}>
      <div style={{ textAlign:'center', padding:'8px 0 16px' }}>
        <div style={{ fontSize:22, fontWeight:800, marginBottom:12 }}>{nom || '—'}</div>
        <div style={{ display:'flex', flexWrap:'wrap', justifyContent:'center', gap:8 }}>
          <span style={{ background:'#f8f8f8', border:'1.5px solid #eee', borderRadius:8, padding:'6px 14px', fontSize:13, fontWeight:600 }}>{fmtResaDate(resa.date)}</span>
          <span style={{ background:'#f8f8f8', border:'1.5px solid #eee', borderRadius:8, padding:'6px 14px', fontSize:13, fontWeight:600 }}>{resa.service === 'midi' ? '🌞 Midi' : '🌙 Soir'}{resa.heure ? ` · ${resa.heure}` : ''}</span>
          <span style={{ background:'#f8f8f8', border:'1.5px solid #eee', borderRadius:8, padding:'6px 14px', fontSize:13, fontWeight:600 }}>👥 {resa.nb_personnes} pers.</span>
        </div>
      </div>
    </Modal>
  );
}

const STATUTS_OPTIONS = [
  { val:'attente',   label:'En attente' },
  { val:'rappeler',  label:'À rappeler' },
  { val:'confirmee', label:'Confirmée' },
  { val:'refusee',   label:'Refusée' },
  { val:'annulee',   label:'Annulée' },
  { val:'venue',     label:'Venue' },
  { val:'absente',   label:'Absente' },
];

function DetailResaModal({ resa, onClose, onSaved }) {
  const c = resa.clients || {};
  const nom = c.entreprise ? c.entreprise : `${c.prenom || ''} ${c.nom || ''}`.trim();
  const [statut, setStatut] = useState(resa.statut);
  const [saving, setSaving] = useState(false);

  async function saveStatut() {
    setSaving(true);
    const { error } = await supabase.from('reservations').update({ statut, updated_at: new Date().toISOString() }).eq('id', resa.id);
    setSaving(false);
    if (error) { alert('Erreur lors de la mise à jour'); return; }
    onSaved();
    onClose();
  }

  return (
    <Modal title="Détail de la réservation" onClose={onClose} maxW={460}
      footer={[
        <button key="f" type="button" onClick={onClose} style={{...btnSecondary}}>Fermer</button>,
        <button key="s" type="button" onClick={saveStatut} disabled={saving || statut === resa.statut} style={{ background: statut !== resa.statut ? '#111' : '#ddd', color: statut !== resa.statut ? '#fff' : '#999', border:'none', borderRadius:8, padding:'0 18px', height:38, fontWeight:700, fontSize:14, cursor: statut !== resa.statut ? 'pointer' : 'not-allowed' }}>
          {saving ? 'Enregistrement…' : 'Enregistrer'}
        </button>
      ]}>
      <div style={{ display:'flex', flexDirection:'column', gap:10 }}>
        <div style={{ fontSize:18, fontWeight:800 }}>{nom || '—'}</div>
        {c.prenom && c.nom && c.entreprise && <div style={{ fontSize:13, color:'#888' }}>{c.prenom} {c.nom}</div>}
        {c.tel && <a href={`tel:${c.tel}`} style={{ display:'inline-flex', alignItems:'center', gap:8, background:G, color:'#111', borderRadius:9, padding:'8px 18px', fontSize:14, fontWeight:700, textDecoration:'none', width:'fit-content' }}>📞 Appeler · {c.tel}</a>}
        {c.mail && <div style={{ fontSize:13, color:'#3b82f6' }}>{c.mail}</div>}
        <div style={{ height:1, background:'#f0f0f0', margin:'4px 0' }} />
        {[
          ['Date', fmtResaDate(resa.date)],
          ['Service', resa.service === 'midi' ? '🌞 Midi' : '🌙 Soir'],
          resa.heure ? ['Heure', resa.heure] : null,
          ['Personnes', `${resa.nb_personnes} personne${resa.nb_personnes > 1 ? 's' : ''}`],
          resa.occasion ? ['Occasion', resa.occasion] : null,
        ].filter(Boolean).map(([l,v]) => (
          <div key={l} style={{ display:'flex', justifyContent:'space-between', fontSize:14, padding:'4px 0', borderBottom:'1px solid #f8f8f8' }}>
            <span style={{ color:'#888' }}>{l}</span><span style={{ fontWeight:600 }}>{v}</span>
          </div>
        ))}
        {resa.commentaire_client && <p style={{ fontSize:13, color:'#aaa', fontStyle:'italic', borderLeft:'3px solid #eee', paddingLeft:10, margin:'4px 0' }}>"{resa.commentaire_client}"</p>}
        {resa.raison_refus && <div style={{ background:'#fef2f2', borderRadius:8, padding:'8px 12px', fontSize:13, color:'#dc2626' }}>Motif refus : {resa.raison_refus}</div>}
        <div style={{ height:1, background:'#f0f0f0', margin:'4px 0' }} />
        <div style={{ display:'flex', alignItems:'center', gap:12 }}>
          <span style={{ fontSize:13, color:'#555', fontWeight:600, flexShrink:0 }}>Statut</span>
          <select value={statut} onChange={e=>setStatut(e.target.value)} style={{ flex:1, height:40, border:'1.5px solid #ddd', borderRadius:8, padding:'0 10px', fontSize:14, background:'#fff', outline:'none', cursor:'pointer' }}>
            {STATUTS_OPTIONS.map(o => <option key={o.val} value={o.val}>{o.label}</option>)}
          </select>
        </div>
      </div>
    </Modal>
  );
}

function ReservationsPage({ onBack, showToast, user, inline = false, onResaCountChange }) {
  const [resaList, setResaList] = useState([]);
  const [loading, setLoading] = useState(true);
  const [refusResa, setRefusResa] = useState(null);
  const [acceptResa, setAcceptResa] = useState(null);
  const [detailResa, setDetailResa] = useState(null);
  const [histOpen, setHistOpen] = useState(false);
  const isMobile = useIsMobile();
  const qr = `https://api.qrserver.com/v1/create-qr-code/?size=120x120&data=${encodeURIComponent(FORM_URL)}`;

  useEffect(() => { loadResa(); }, []);

  async function loadResa() {
    setLoading(true);
    const { data, error } = await supabase
      .from('reservations')
      .select('*, clients(id, nom, prenom, tel, mail, genre, entreprise)')
      .order('created_at', { ascending: false });
    if (error) showToast('Erreur chargement réservations', 'error');
    else {
      setResaList(data || []);
      onResaCountChange?.((data||[]).filter(r=>r.statut==='attente').length);
    }
    setLoading(false);
  }

  async function accepter(r) {
    const { error } = await supabase.from('reservations').update({
      statut: 'confirmee', traited_at: new Date().toISOString(), traited_by: user?.email
    }).eq('id', r.id);
    setAcceptResa(null);
    if (error) { showToast('Erreur', 'error'); return; }
    showToast('Réservation confirmée ✓');
    loadResa();
    if (!r.clients?.mail) { showToast("⚠️ Email non envoyé (pas d'adresse)"); return; }
    const dateFormatee = new Date(r.date).toLocaleDateString('fr-FR', {weekday:'long', day:'numeric', month:'long', year:'numeric'});
    const heureBase = r.heure || '19:00';
    const calStart = r.date.replace(/-/g,'') + 'T' + heureBase.replace(':','') + '00';
    const calEnd = r.date.replace(/-/g,'') + 'T' + (parseInt(heureBase.split(':')[0])+2) + heureBase.split(':')[1] + '00';
    const calUrl = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=R%C3%A9servation+Le+TED&dates=${calStart}/${calEnd}&details=R%C3%A9servation+confirm%C3%A9e+au+TED+pour+${r.nb_personnes}+personne(s)&location=28+Av.+des+Fr%C3%A8res+Montgolfier+69680+Chassieu`;
    const htmlConfirmation = `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;background:#f8f8f8;padding:20px">
  <div style="background:#111111;padding:28px 24px;text-align:center;border-radius:12px 12px 0 0;border-bottom:4px solid #E8C547">
    <img src="https://ted-crm.pages.dev/favicon.png" alt="Le TED" style="height:60px;margin-bottom:12px" />
    <h1 style="color:#E8C547;margin:0;font-size:28px;letter-spacing:2px;font-weight:800">LE TED</h1>
    <p style="color:#888;margin:4px 0 0;font-size:13px;letter-spacing:1px">RESTAURANT &amp; CLUB — CHASSIEU</p>
  </div>
  <div style="background:#fff;padding:28px 24px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08)">
    <h2 style="color:#111;margin:0 0 8px;font-size:22px">Bonjour ${r.clients.prenom} 👋</h2>
    <p style="color:#444;font-size:16px;margin:0 0 24px">Votre réservation est <strong style="color:#16a34a">confirmée</strong> ✅</p>
    <div style="background:#f9f9f9;border-left:4px solid #E8C547;padding:20px;border-radius:0 8px 8px 0;margin-bottom:24px">
      <p style="margin:0 0 10px;font-size:15px">📅 <strong>Date :</strong> ${dateFormatee}</p>
      <p style="margin:0 0 10px;font-size:15px">🕐 <strong>Heure :</strong> ${r.heure || 'À confirmer'}</p>
      <p style="margin:0 0 10px;font-size:15px">👥 <strong>Nombre de personnes :</strong> ${r.nb_personnes}</p>
      <p style="margin:0;font-size:15px">🍽 <strong>Service :</strong> ${r.service === 'midi' ? 'Déjeuner' : 'Dîner'}</p>
      ${r.occasion ? `<p style="margin:8px 0 0;font-size:15px">🎉 <strong>Occasion :</strong> ${r.occasion}</p>` : ''}
    </div>
    <div style="text-align:center;margin-bottom:24px">
      <a href="${calUrl}" target="_blank" style="display:inline-block;background:#E8C547;color:#111;text-decoration:none;padding:14px 28px;border-radius:8px;font-weight:800;font-size:15px;letter-spacing:0.5px">📆 Ajouter à mon agenda</a>
    </div>
    <div style="background:#fff8e1;border:1.5px solid #E8C547;border-radius:8px;padding:16px;margin-bottom:24px">
      <p style="margin:0;font-size:14px;color:#555;line-height:1.6">⚠️ <strong>En cas d'annulation ou de modification</strong>, merci de nous prévenir au plus tôt au <strong>04 78 90 67 80</strong> ou par email afin que nous puissions libérer la table pour d'autres clients. Merci de votre compréhension.</p>
    </div>
    <div style="border-top:1px solid #eee;padding-top:20px;text-align:center">
      <p style="color:#111;font-weight:700;font-size:15px;margin:0 0 6px">Le TED — Restaurant &amp; Club</p>
      <p style="color:#888;font-size:13px;margin:0 0 4px">📍 28 Av. des Frères Montgolfier, 69680 Chassieu</p>
      <p style="color:#888;font-size:13px;margin:0 0 4px">📞 04 78 90 67 80</p>
      <p style="margin:0;display:flex;align-items:center;justify-content:center;gap:6px"><img src="https://ted-crm.pages.dev/favicon.png" alt="" style="height:16px;width:16px;vertical-align:middle" /><a href="https://leted.fr" style="color:#888;font-size:13px;text-decoration:none">leted.fr</a></p>
    </div>
    <p style="text-align:center;color:#bbb;font-size:12px;margin-top:20px">Nous avons hâte de vous accueillir ! 🎉</p>
  </div>
</div>`;
    const resEmail = await sendBrevoEmail(
      r.clients.mail,
      `${r.clients.prenom || ''} ${r.clients.nom || ''}`.trim(),
      `✅ Réservation confirmée au TED — ${dateFormatee}`,
      htmlConfirmation
    );
    showToast(resEmail?.success ? '📧 Email envoyé' : '⚠️ Email non envoyé');
  }

  async function refuser(r, raison) {
    const { error } = await supabase.from('reservations').update({
      statut: 'refusee', raison_refus: raison, traited_at: new Date().toISOString(), traited_by: user?.email
    }).eq('id', r.id);
    setRefusResa(null);
    if (error) { showToast('Erreur', 'error'); return; }
    showToast('Réservation refusée');
    loadResa();
    if (!r.clients?.mail) { showToast("⚠️ Email non envoyé (pas d'adresse)"); return; }
    const dateFormateeRefus = new Date(r.date).toLocaleDateString('fr-FR', {weekday:'long', day:'numeric', month:'long', year:'numeric'});
    const htmlRefus = `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;background:#f8f8f8;padding:20px">
  <div style="background:#111111;padding:28px 24px;text-align:center;border-radius:12px 12px 0 0;border-bottom:4px solid #E8C547">
    <img src="https://ted-crm.pages.dev/favicon.png" alt="Le TED" style="height:60px;margin-bottom:12px" />
    <h1 style="color:#E8C547;margin:0;font-size:28px;letter-spacing:2px;font-weight:800">LE TED</h1>
    <p style="color:#888;margin:4px 0 0;font-size:13px;letter-spacing:1px">RESTAURANT &amp; CLUB — CHASSIEU</p>
  </div>
  <div style="background:#fff;padding:28px 24px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08)">
    <h2 style="color:#111;margin:0 0 8px;font-size:22px">Bonjour ${r.clients.prenom},</h2>
    <p style="color:#444;font-size:16px;margin:0 0 24px">Merci pour votre demande de réservation au TED.</p>
    <div style="background:#f9f9f9;border-left:4px solid #ccc;padding:20px;border-radius:0 8px 8px 0;margin-bottom:24px">
      <p style="margin:0 0 8px;font-size:15px">📅 <strong>Date demandée :</strong> ${dateFormateeRefus}</p>
      <p style="margin:0 0 8px;font-size:15px">👥 <strong>Nombre de personnes :</strong> ${r.nb_personnes}</p>
      <p style="margin:0;font-size:15px">🍽 <strong>Service :</strong> ${r.service === 'midi' ? 'Déjeuner' : 'Dîner'}</p>
    </div>
    <div style="background:#fff2f2;border:1.5px solid #dc2626;border-radius:8px;padding:16px;margin-bottom:24px">
      <p style="margin:0;font-size:14px;color:#dc2626;font-weight:700">Motif : ${raison}</p>
    </div>
    <p style="color:#444;font-size:15px;line-height:1.6">Nous sommes désolés de ne pas pouvoir donner suite à votre demande. N'hésitez pas à nous contacter directement au <strong>04 78 90 67 80</strong> pour trouver une autre disponibilité ou pour toute question.</p>
    <div style="border-top:1px solid #eee;padding-top:20px;text-align:center;margin-top:24px">
      <p style="color:#111;font-weight:700;font-size:15px;margin:0 0 6px">Le TED — Restaurant &amp; Club</p>
      <p style="color:#888;font-size:13px;margin:0 0 4px">📍 28 Av. des Frères Montgolfier, 69680 Chassieu</p>
      <p style="color:#888;font-size:13px;margin:0 0 4px">📞 04 78 90 67 80</p>
      <p style="margin:0;display:flex;align-items:center;justify-content:center;gap:6px"><img src="https://ted-crm.pages.dev/favicon.png" alt="" style="height:16px;width:16px;vertical-align:middle" /><a href="https://leted.fr" style="color:#888;font-size:13px;text-decoration:none">leted.fr</a></p>
    </div>
    <p style="text-align:center;color:#bbb;font-size:12px;margin-top:20px">À bientôt au TED 🙏</p>
  </div>
</div>`;
    const resEmail = await sendBrevoEmail(
      r.clients.mail,
      `${r.clients.prenom || ''} ${r.clients.nom || ''}`.trim(),
      `Votre demande de réservation au TED — ${dateFormateeRefus}`,
      htmlRefus
    );
    showToast(resEmail?.success ? '📧 Email envoyé' : '⚠️ Email non envoyé');
  }

  function copyLink() {
    navigator.clipboard.writeText(FORM_URL).then(() => showToast('Lien copié ! ✓')).catch(() => {
      const el = document.createElement('textarea');
      el.value = FORM_URL; document.body.appendChild(el); el.select(); document.execCommand('copy'); document.body.removeChild(el);
      showToast('Lien copié ! ✓');
    });
  }

  const attente = resaList.filter(r => r.statut === 'attente');
  const historique = resaList.filter(r => r.statut !== 'attente');

  const cardStyle = { background:'#fff', borderRadius:14, border:'1.5px solid #f0f0f0', padding:16, marginBottom:10, boxShadow:'0 2px 8px rgba(0,0,0,0.04)' };

  return (
    <div style={{ fontFamily:"'Inter','Segoe UI',Arial,sans-serif", background:'#f8f8f8', minHeight: inline ? undefined : '100vh' }}>
      {/* Header — desktop full-page mode only */}
      {!inline && (
        <header style={{ background:'#111', color:'#fff', padding:'0 20px', height:56, display:'flex', alignItems:'center', gap:14, borderBottom:`3px solid ${G}`, flexShrink:0 }}>
          <button onClick={onBack} style={{ background:'rgba(255,255,255,0.1)', border:'none', borderRadius:8, height:34, padding:'0 14px', color:'#fff', fontWeight:600, fontSize:13, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}>← Retour</button>
          <span style={{ fontWeight:700, fontSize:15, flex:1 }}>📅 Réservations</span>
          {attente.length > 0 && <span style={{ background:'#dc2626', color:'#fff', borderRadius:99, padding:'2px 8px', fontSize:12, fontWeight:700 }}>{attente.length}</span>}
        </header>
      )}

      <main style={{ maxWidth:800, margin:'0 auto', padding: isMobile ? '16px 12px 100px' : '24px 20px 40px' }}>

        {/* ── Bloc partage lien ── */}
        <div style={{ background:'#fff', borderRadius:14, border:'1.5px solid #eee', padding:18, marginBottom:24, boxShadow:'0 2px 8px rgba(0,0,0,0.04)' }}>
          <p style={{ fontSize:12, fontWeight:700, color:'#888', textTransform:'uppercase', letterSpacing:0.5, marginBottom:12 }}>🔗 Formulaire de réservation en ligne</p>
          <div style={{ display:'flex', gap:12, alignItems:'flex-start', flexWrap: isMobile ? 'wrap' : 'nowrap' }}>
            <div style={{ flex:1, minWidth:0 }}>
              <div style={{ background:'#f8f8f8', border:'1.5px solid #eee', borderRadius:8, padding:'8px 12px', fontSize:12, color:'#555', fontFamily:'monospace', wordBreak:'break-all', marginBottom:10, lineHeight:1.5 }}>{FORM_URL}</div>
              <div style={{ display:'flex', gap:8 }}>
                <button onClick={copyLink} style={{ flex:1, background:'#111', color:'#fff', border:'none', borderRadius:8, height:38, fontWeight:700, fontSize:13, cursor:'pointer' }}>📋 Copier</button>
                <button onClick={()=>window.open(FORM_URL,'_blank')} style={{ flex:1, background:G, color:'#111', border:'none', borderRadius:8, height:38, fontWeight:700, fontSize:13, cursor:'pointer' }}>🔗 Ouvrir</button>
              </div>
            </div>
            <div style={{ flexShrink:0, textAlign:'center' }}>
              <img src={qr} alt="QR" width={90} height={90} style={{ display:'block', borderRadius:8, border:'1.5px solid #eee' }} />
              <p style={{ fontSize:10, color:'#bbb', marginTop:4 }}>Scanner</p>
            </div>
          </div>
        </div>

        {/* ── Demandes en attente ── */}
        <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:14 }}>
          <h2 style={{ fontSize:17, fontWeight:800, margin:0 }}>Demandes en attente</h2>
          {attente.length > 0 && <span style={{ background:'#dc2626', color:'#fff', borderRadius:99, padding:'2px 9px', fontSize:12, fontWeight:700 }}>{attente.length}</span>}
        </div>

        {loading && <p style={{ color:'#bbb', fontSize:14, padding:'20px 0' }}>Chargement…</p>}

        {!loading && attente.length === 0 && (
          <div style={{ textAlign:'center', padding:'32px 0', color:'#bbb' }}>
            <div style={{ fontSize:40, marginBottom:10 }}>📭</div>
            <p style={{ fontSize:15 }}>Aucune nouvelle demande</p>
          </div>
        )}

        {attente.map(r => {
          const c = r.clients || {};
          const nom = c.entreprise ? c.entreprise : `${c.prenom || ''} ${c.nom || ''}`.trim();
          return (
            <div key={r.id} style={cardStyle}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', gap:8, marginBottom:10 }}>
                <div>
                  <span style={{ fontWeight:700, fontSize:15 }}>{nom || '—'}</span>
                  {c.mail && <span style={{ fontSize:12, color:'#888', marginTop:2, display:'block' }}>{c.mail}</span>}
                  {c.tel && (
                    <a href={`tel:${c.tel}`} style={{ display:'inline-flex', alignItems:'center', gap:6, background:G, color:'#111', borderRadius:8, padding:'6px 14px', fontSize:13, fontWeight:700, textDecoration:'none', marginTop:8 }}>
                      📞 Appeler · {c.tel}
                    </a>
                  )}
                </div>
                <div style={{ textAlign:'right', flexShrink:0 }}>
                  <div style={{ fontSize:13, fontWeight:700 }}>{fmtResaDate(r.date)}</div>
                  <div style={{ fontSize:12, color:'#888' }}>{r.service === 'midi' ? '🌞 Midi' : '🌙 Soir'}{r.heure ? ` · ${r.heure}` : ''}</div>
                  <div style={{ fontSize:12, color:'#555', marginTop:2 }}>👥 {r.nb_personnes} pers.</div>
                </div>
              </div>
              {r.occasion && <p style={{ fontSize:12, color:'#6b7280', marginBottom:6 }}>🎉 {r.occasion}</p>}
              {r.commentaire_client && <p style={{ fontSize:12, color:'#aaa', fontStyle:'italic', marginBottom:8, borderLeft:`3px solid #eee`, paddingLeft:8 }}>"{r.commentaire_client}"</p>}
              <div style={{ fontSize:11, color:'#bbb', marginBottom:10 }}>Reçue le {new Date(r.created_at).toLocaleDateString('fr-FR')} à {new Date(r.created_at).toLocaleTimeString('fr-FR',{hour:'2-digit',minute:'2-digit'})}</div>
              <div style={{ display:'flex', gap:8 }}>
                <button onClick={()=>setAcceptResa(r)} style={{ flex:1, background:'#16a34a', color:'#fff', border:'none', borderRadius:8, height:40, fontWeight:700, fontSize:14, cursor:'pointer' }}>✓ Accepter</button>
                <button onClick={()=>setRefusResa(r)} style={{ flex:1, background:'#fef2f2', color:'#dc2626', border:'1.5px solid #dc2626', borderRadius:8, height:40, fontWeight:700, fontSize:14, cursor:'pointer' }}>✕ Refuser</button>
              </div>
            </div>
          );
        })}

        {/* ── Historique accordéon ── */}
        {historique.length > 0 && (
          <div style={{ marginTop:28 }}>
            <button onClick={()=>setHistOpen(o=>!o)} style={{ width:'100%', display:'flex', alignItems:'center', justifyContent:'space-between', background:'#fff', border:'1.5px solid #eee', borderRadius:12, padding:'14px 18px', cursor:'pointer', fontWeight:800, fontSize:16 }}>
              <span>Historique <span style={{ color:'#bbb', fontWeight:400, fontSize:14 }}>({historique.length})</span></span>
              <span style={{ fontSize:20, color:'#888', transform: histOpen ? 'rotate(90deg)' : 'none', transition:'transform 0.2s' }}>›</span>
            </button>
            {histOpen && (
              <div style={{ marginTop:8 }}>
                {historique.map(r => {
                  const c = r.clients || {};
                  const nom = c.entreprise ? c.entreprise : `${c.prenom || ''} ${c.nom || ''}`.trim();
                  return (
                    <div key={r.id} onClick={()=>setDetailResa(r)} style={{ ...cardStyle, display:'flex', alignItems:'center', gap:12, padding:'12px 14px', cursor:'pointer' }}
                      onMouseEnter={e=>e.currentTarget.style.background='#fffbeb'} onMouseLeave={e=>e.currentTarget.style.background='#fff'}>
                      <div style={{ flex:1, minWidth:0 }}>
                        <span style={{ fontWeight:600, fontSize:14 }}>{nom || '—'}</span>
                        <span style={{ fontSize:12, color:'#888', marginLeft:8 }}>{fmtResaDate(r.date)} · {r.nb_personnes} pers.</span>
                      </div>
                      <div style={{ flexShrink:0 }}>{statutBadge(r.statut)}</div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
      </main>

      {acceptResa && <AccepterModal resa={acceptResa} onConfirm={()=>accepter(acceptResa)} onCancel={()=>setAcceptResa(null)} />}
      {refusResa && <RefusModal onConfirm={raison=>refuser(refusResa, raison)} onCancel={()=>setRefusResa(null)} />}
      {detailResa && <DetailResaModal resa={detailResa} onClose={()=>setDetailResa(null)} onSaved={()=>{ loadResa(); setDetailResa(null); }} />}
    </div>
  );
}

// ─── Main CRM App ─────────────────────────────────────────────────────────────

function CRMApp({ user, onLogout }) {
  const [clients, setClients] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [sortKey, setSortKey] = useState("created_at");
  const [sortDir, setSortDir] = useState("desc");
  const [filterGenre, setFilterGenre] = useState("");
  const [filterMonth, setFilterMonth] = useState("");
  const [pageSize, setPageSize] = useState(25);
  const [page, setPage] = useState(1);
  const [modalAdd, setModalAdd] = useState(false);
  const [modalEdit, setModalEdit] = useState(null);
  const [modalDelete, setModalDelete] = useState(null);
  const [modalImport, setModalImport] = useState(false);
  const [modalComment, setModalComment] = useState(null);
  const [toast, setToast] = useState(null);
  const [hoverRow, setHoverRow] = useState(null);
  const [showFilters, setShowFilters] = useState(false);
  const [activeTab, setActiveTab] = useState("tous");
  const [showSearch, setShowSearch] = useState(false);
  const [modalCorbeille, setModalCorbeille] = useState(false);
  const [mobileAction, setMobileAction] = useState(null);
  const [showResaPage, setShowResaPage] = useState(false);
  const [resaAttenteCount, setResaAttenteCount] = useState(0);
  const [showPlusSheet, setShowPlusSheet] = useState(false);
  const [mobileTab, setMobileTab] = useState('clients'); // 'clients' | 'reservations'
  const [showAddResa, setShowAddResa] = useState(false);
  const deleteGuard = useRef(false);
  const isMobile = useIsMobile();

  const showToast = useCallback((msg, type="success") => setToast({msg,type}), []);

  // ─── Load from Supabase ───────────────────────────────────────────────────
  useEffect(() => {
    loadClients();
    loadResaCount();
  }, []);

  async function loadResaCount() {
    const { count } = await supabase.from('reservations').select('id', { count:'exact', head:true }).eq('statut','attente');
    setResaAttenteCount(count || 0);
  }

  async function loadClients() {
    setLoading(true);
    const { data, error } = await supabase.from("clients").select("*").is("deleted_at", null).order("created_at", { ascending: false });
    if (error) { showToast("Erreur de chargement", "error"); }
    else { setClients(data || []); }
    setLoading(false);
  }

  // ─── CRUD ─────────────────────────────────────────────────────────────────
  async function addClient(c) {
    const tempId = genId();
    const newClient = { ...c, id: tempId, genre: c.genre||"Non renseigné", nom: c.nom||"", prenom: c.prenom||"", tel: c.tel||"", mail: c.mail||"", commentaire: c.commentaire||"", entreprise: c.entreprise||"", created_at: c.created_at||new Date().toISOString() };
    setClients(prev => [newClient, ...prev]);
    setModalAdd(false);
    showToast("Client ajouté avec succès ✓");
    setPage(1);
    const { data, error } = await supabase.from("clients").insert([{ genre:newClient.genre, nom:newClient.nom, prenom:newClient.prenom, tel:newClient.tel, mail:newClient.mail, commentaire:newClient.commentaire, entreprise:newClient.entreprise, created_at:newClient.created_at }]).select().single();
    if (error) {
      setClients(prev => prev.filter(x => x.id !== tempId));
      showToast("Erreur lors de l'ajout : " + error.message, "error");
      return;
    }
    setClients(prev => prev.map(x => x.id === tempId ? data : x));
  }

  async function editClient(c) {
    setClients(prev => prev.map(x => x.id === c.id ? {...x, ...c} : x));
    setModalEdit(null);
    showToast("Client modifié avec succès ✓");
    const { error } = await supabase.from("clients").update({ genre:c.genre, nom:c.nom, prenom:c.prenom, tel:c.tel, mail:c.mail, commentaire:c.commentaire, entreprise:c.entreprise||"" }).eq("id", c.id);
    if (error) {
      showToast("Erreur lors de la modification", "error");
      loadClients();
    }
  }

  async function deleteClient(id) {
    if (deleteGuard.current) return;
    deleteGuard.current = true;
    setClients(prev => prev.filter(x => x.id !== id));
    setModalDelete(null);
    showToast("Client déplacé dans la corbeille ✓");
    const { error } = await supabase.from("clients").update({ deleted_at: new Date().toISOString(), deleted_by: user.email }).eq("id", id);
    if (error) {
      showToast("Erreur lors de la suppression", "error");
      loadClients();
    }
    setTimeout(() => { deleteGuard.current = false; }, 500);
  }

  async function importClients(rows) {
    const { data, error } = await supabase.from("clients").insert(rows).select();
    if (error) { showToast("Erreur lors de l'import", "error"); return; }
    setClients(prev => [...(data||[]), ...prev]);
    setModalImport(false);
    showToast(`${rows.length} client(s) importé(s) ✓`);
  }

  // ─── Export / Backup ──────────────────────────────────────────────────────
  function saveBackup() {
    const json = JSON.stringify({ version:2, date:new Date().toISOString(), clients });
    downloadBlob(json, `backup_TED_${new Date().toISOString().slice(0,10)}.json`, "application/json");
    showToast("Sauvegarde téléchargée ✓");
  }

  const restoreRef = useRef();
  function handleRestoreFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async ev => {
      try {
        const d = JSON.parse(ev.target.result);
        if (d.clients && Array.isArray(d.clients)) {
          if (window.confirm(`Restaurer ${d.clients.length} clients ? Les clients actuels seront supprimés.`)) {
            await supabase.from("clients").delete().neq("id", "00000000-0000-0000-0000-000000000000");
            const toInsert = d.clients.map(c => ({ genre:c.genre, nom:c.nom, prenom:c.prenom, tel:c.tel, mail:c.mail, commentaire:c.commentaire, created_at:c.created_at||c.createdAt }));
            const { data } = await supabase.from("clients").insert(toInsert).select();
            setClients(data || []);
            showToast(`${d.clients.length} clients restaurés ✓`);
          }
        }
      } catch { showToast("Fichier invalide", "error"); }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  // ─── Sort & Filter ────────────────────────────────────────────────────────
  function toggleSort(k) {
    if (sortKey === k) setSortDir(d => d==="asc"?"desc":"asc");
    else { setSortKey(k); setSortDir("asc"); }
    setPage(1);
  }

  const filtered = useMemo(() => {
    let list = [...clients];
    if (activeTab === "particuliers") list = list.filter(c => c.genre !== "Entreprise");
    if (activeTab === "entreprises") list = list.filter(c => c.genre === "Entreprise");
    if (filterGenre) list = list.filter(c => c.genre === filterGenre);
    if (filterMonth) list = list.filter(c => { const d = new Date(c.created_at); return !isNaN(d) && (d.getMonth()+1) === parseInt(filterMonth); });
    if (search.trim()) {
      const terms = normalizeStr(search).split(/\s+/).filter(Boolean);
      list = list.filter(c => {
        const blob = [normalizeStr(c.genre),normalizeStr(c.nom),normalizeStr(c.prenom),c.tel||"",normalizeStr(c.mail),normalizeStr(formatDate(c.created_at)),normalizeStr(getMonthName(c.created_at)),c.created_at?new Date(c.created_at).getFullYear().toString():"",normalizeStr(c.commentaire)].join(" ");
        return terms.every(t => blob.includes(t));
      });
    }
    list.sort((a,b) => {
      let va = a[sortKey]||"", vb = b[sortKey]||"";
      if (sortKey === "created_at") { va = new Date(va).getTime()||0; vb = new Date(vb).getTime()||0; return sortDir==="asc"?va-vb:vb-va; }
      va = normalizeStr(va); vb = normalizeStr(vb);
      return sortDir==="asc"?va.localeCompare(vb):vb.localeCompare(va);
    });
    return list;
  }, [clients, search, filterGenre, filterMonth, sortKey, sortDir, activeTab]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const pageClients = filtered.slice((safePage-1)*pageSize, safePage*pageSize);
  const newMonth = clients.filter(c => isCurrentMonth(c.created_at)).length;

  const sel = { height:36, border:"1.5px solid #ddd", borderRadius:7, padding:"0 10px", fontSize:13, background:"#fff", cursor:"pointer", outline:"none" };

  function Th({ col, label }) {
    const active = sortKey === col;
    return <th onClick={()=>toggleSort(col)} style={{ background:"#111", color:"#fff", padding:"10px 12px", textAlign:"left", fontWeight:600, fontSize:12, letterSpacing:0.5, cursor:"pointer", userSelect:"none", whiteSpace:"nowrap" }}>{label} <span style={{color:active?G:"#666",fontSize:10}}>{active?(sortDir==="asc"?"▲":"▼"):"⇅"}</span></th>;
  }

  if (loading) return <div style={{ textAlign:"center", paddingTop:80, fontSize:16, color:"#888" }}>Chargement des clients…</div>;
  if (showResaPage && !isMobile) return <ReservationsPage onBack={()=>{ setShowResaPage(false); loadResaCount(); }} showToast={showToast} user={user} />;

  return (
    <div style={{ fontFamily:"'Inter','Segoe UI',Arial,sans-serif", minHeight:"100vh", background:"#f8f8f8", color:"#111" }}>
      <style>{`
        @keyframes popIn { 0%{opacity:0;transform:scale(0.5)} 70%{transform:scale(1.05)} 100%{opacity:1;transform:scale(1)} }
        @keyframes scaleIn { from{transform:scale(0)} to{transform:scale(1)} }
        @keyframes slideUpFade { from{opacity:0;transform:translateY(24px)} to{opacity:1;transform:translateY(0)} }
        @keyframes slideDownFade { from{opacity:0;transform:translateY(-12px)} to{opacity:1;transform:translateY(0)} }
        @keyframes fadeIn { from{opacity:0} to{opacity:1} }
        @keyframes shimmer { 0%{background-position:-200% 0} 100%{background-position:200% 0} }
        .card-mobile { animation: slideUpFade 0.22s cubic-bezier(0.34,1.2,0.64,1) both; }
        .btn-mobile:active { transform: scale(0.96); opacity: 0.85; }
        .tab-pill { transition: background 0.15s, color 0.15s, transform 0.1s; }
        .tab-pill:active { transform: scale(0.95); }
        .client-card { background:#fff; border-radius:16px; border:1.5px solid #efefef; margin-bottom:12px; overflow:hidden; box-shadow:0 1px 6px rgba(0,0,0,0.05); }
      `}</style>

      {/* ═══ MOBILE HEADER FIXE (header + tabs + recherche) ═══ */}
      {isMobile && (
        <div style={{ position:'fixed', top:0, left:0, right:0, zIndex:200, background:'#fff' }}>
          {/* Barre titre */}
          <div style={{ background:'#111', borderBottom:`3px solid ${G}`, display:'flex', alignItems:'center', justifyContent:'space-between', padding:'0 14px', height:50 }}>
            <div style={{ display:'flex', alignItems:'center', gap:8 }}>
              <img src={require('./logo.png')} alt="TED" style={{ height:26, filter:'brightness(0) invert(1)' }} onError={e=>e.target.style.display='none'} />
              <span style={{ color:'#fff', fontWeight:800, fontSize:15 }}>TED <span style={{color:G}}>CRM</span></span>
            </div>
            <button onClick={()=>{ if(window.confirm('Voulez-vous vraiment vous déconnecter ?')) onLogout(); }} style={{ background:'rgba(255,255,255,0.1)', border:'none', borderRadius:8, width:34, height:34, display:'flex', alignItems:'center', justifyContent:'center', fontSize:14, cursor:'pointer', color:'#fff' }}>⎋</button>
          </div>
          {/* Onglets + Recherche — uniquement sur l'onglet Clients */}
          {mobileTab === 'clients' && (
            <>
              <div style={{ display:'flex', gap:6, padding:'8px 12px 6px', background:'#fff', overflowX:'auto', scrollbarWidth:'none' }}>
                {[
                  { id:'tous', label:'Tous', count:clients.length },
                  { id:'particuliers', label:'👤 Particuliers', count:clients.filter(c=>c.genre!=='Entreprise').length },
                  { id:'entreprises', label:'🏢 Entreprises', count:clients.filter(c=>c.genre==='Entreprise').length }
                ].map(tab => (
                  <button key={tab.id} onClick={()=>{setActiveTab(tab.id);setPage(1)}} style={{ background:activeTab===tab.id?'#111':'#f0f0f0', color:activeTab===tab.id?'#fff':'#666', border:'none', borderRadius:99, padding:'6px 14px', fontSize:12, fontWeight:activeTab===tab.id?700:500, cursor:'pointer', whiteSpace:'nowrap', flexShrink:0 }}>
                    {tab.label} <span style={{ background:activeTab===tab.id?G:'#ddd', color:activeTab===tab.id?'#111':'#999', borderRadius:99, padding:'1px 6px', fontSize:10, fontWeight:700, marginLeft:2 }}>{tab.count}</span>
                  </button>
                ))}
              </div>
              <div style={{ padding:'0 12px 8px', background:'#fff', borderBottom:'1px solid #eee' }}>
                <div style={{ position:'relative' }}>
                  <span style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', color:'#bbb', fontSize:14 }}>🔍</span>
                  <input value={search} onChange={e=>{setSearch(e.target.value);setPage(1)}} placeholder="Rechercher..." style={{ width:'100%', height:38, border:'1.5px solid #eee', borderRadius:10, padding:'0 36px 0 36px', fontSize:16, outline:'none', boxSizing:'border-box', background:'#f8f8f8' }} />
                  {search && <button onClick={()=>{setSearch('');setPage(1)}} style={{ position:'absolute', right:10, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', fontSize:16, cursor:'pointer', color:'#aaa' }}>✕</button>}
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* ═══ DESKTOP HEADER ═══ */}
      {!isMobile && (
        <header style={{background:"#111", color:"#fff", padding:"0 20px", display:"flex", alignItems:"center", justifyContent:"space-between", height:56, borderBottom:`3px solid ${G}`, flexShrink:0}}>
          <div style={{display:"flex", alignItems:"center", gap:10}}>
            <img src={require('./logo.png')} alt="TED" style={{height:30, filter:'brightness(0) invert(1)'}} onError={e=>e.target.style.display='none'} />
            <h1 style={{fontSize:15, fontWeight:700, letterSpacing:2, color:"#fff", margin:0}}>
              <span style={{color:G}}>TED</span> — FICHIER CLIENTS
            </h1>
          </div>
          <div style={{display:"flex", gap:6, alignItems:"center", flexShrink:0}}>
            <span style={{fontSize:11, color:"#666", marginRight:4, maxWidth:120, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap"}}>{user.email}</span>
            <button onClick={saveBackup} title="Sauvegarder" style={{background:"transparent", color:"#ccc", border:"1px solid #444", borderRadius:7, width:32, height:32, cursor:"pointer", fontSize:16, display:"flex", alignItems:"center", justifyContent:"center"}}>💾</button>
            <button onClick={()=>restoreRef.current?.click()} title="Restaurer" style={{background:"transparent", color:"#ccc", border:"1px solid #444", borderRadius:7, width:32, height:32, cursor:"pointer", fontSize:16, display:"flex", alignItems:"center", justifyContent:"center"}}>🔄</button>
            <input ref={restoreRef} type="file" accept=".json" style={{display:"none"}} onChange={handleRestoreFile} />
            <button onClick={()=>setModalCorbeille(true)} style={{background:"transparent", color:G, border:`1px solid ${G}`, borderRadius:7, padding:"0 10px", height:32, fontSize:12, fontWeight:700, cursor:"pointer", whiteSpace:"nowrap"}}>🗑 Corbeille</button>
            <button onClick={()=>setShowResaPage(true)} style={{ position:'relative', background:'transparent', color:'#ccc', border:'1px solid #444', borderRadius:7, padding:'0 10px', height:32, fontSize:12, cursor:'pointer', whiteSpace:'nowrap' }}>
              📅 Réservations
              {resaAttenteCount > 0 && <span style={{ marginLeft:6, background:'#dc2626', color:'#fff', borderRadius:99, fontSize:10, fontWeight:700, padding:'1px 6px' }}>{resaAttenteCount}</span>}
            </button>
            <button onClick={onLogout} style={{background:"transparent", color:"#ccc", border:"1px solid #444", borderRadius:7, padding:"0 10px", height:32, fontSize:12, cursor:"pointer"}}>⎋ Quitter</button>
          </div>
        </header>
      )}
      {isMobile && <input ref={restoreRef} type="file" accept=".json" style={{display:"none"}} onChange={handleRestoreFile} />}

      {/* ═══ MOBILE — RÉSERVATIONS INLINE ═══ */}
      {isMobile && mobileTab === 'reservations' && (
        <div style={{ paddingTop:56 }}>
          <ReservationsPage
            inline
            showToast={showToast}
            user={user}
            onResaCountChange={setResaAttenteCount}
          />
        </div>
      )}

      {/* ═══ MOBILE CARDS ═══ */}
      {isMobile && mobileTab === 'clients' && (
        <div style={{ paddingTop:146, paddingBottom:'calc(65px + env(safe-area-inset-bottom, 16px))', paddingLeft:12, paddingRight:12 }}>
          {pageClients.length === 0 && (
            <div style={{ textAlign:'center', padding:'4rem 2rem' }}>
              <div style={{ fontSize:48, marginBottom:12 }}>🔍</div>
              <p style={{ color:'#bbb', fontSize:15 }}>Aucun client trouvé</p>
            </div>
          )}
          {pageClients.map((c,i) => (
            <div key={c.id} style={{ background:'#fff', borderRadius:14, border:'1.5px solid #f0f0f0', padding:'12px', marginBottom:8, boxShadow:'0 2px 8px rgba(0,0,0,0.05)', animation:'slideUpFade 0.25s ease both', animationDelay:`${i*0.04}s` }}>
              {/* Ligne principale : infos + boutons ✏️🗑 */}
              <div style={{ display:'flex', alignItems:'flex-start', gap:8 }}>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:6, marginBottom:4 }}>
                    <span style={badge(c.genre)}>{c.genre}</span>
                    <span style={{ fontWeight:700, fontSize:15, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                      {c.genre==='Entreprise' ? (c.entreprise||c.nom||'—') : `${c.nom||''} ${c.prenom||''}`}
                    </span>
                  </div>
                  {c.mail && <p style={{ fontSize:11, color:'#3b82f6', margin:'2px 0', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{c.mail}</p>}
                  {c.tel && <p style={{ fontSize:12, color:'#555', fontWeight:600, margin:'2px 0' }}>{c.tel}</p>}
                  {c.commentaire && <p style={{ fontSize:11, color:'#aaa', margin:'3px 0 0', fontStyle:'italic', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>"{c.commentaire}"</p>}
                </div>
                <div style={{ flexShrink:0 }}>
                  <button onClick={e=>{ const r=e.currentTarget.getBoundingClientRect(); setMobileAction(mobileAction?.id===c.id ? null : {...c, _rect:r}); }} style={{ background: mobileAction?.id===c.id ? '#111' : '#f5f5f5', border:'none', borderRadius:8, width:34, height:34, display:'flex', alignItems:'center', justifyContent:'center', fontSize:13, fontWeight:700, color: mobileAction?.id===c.id ? '#fff' : '#555', cursor:'pointer', letterSpacing:1 }}>•••</button>
                </div>
              </div>
              {/* Bouton Appeler — séparé en dessous à droite */}
              {c.tel && (
                <div style={{ marginTop:10, display:'flex', justifyContent:'flex-end' }}>
                  <a href={`tel:${c.tel}`} style={{ display:'inline-flex', alignItems:'center', gap:6, background:G, color:'#111', borderRadius:10, padding:'7px 18px', fontSize:14, fontWeight:700, textDecoration:'none', boxShadow:'0 2px 6px rgba(232,197,71,0.35)' }}>
                    📞 Appeler
                  </a>
                </div>
              )}
            </div>
          ))}
          {totalPages > 1 && (
            <div style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:16, padding:'8px 0 16px' }}>
              <button disabled={safePage<=1} onClick={()=>setPage(p=>p-1)} style={{ width:44, height:44, borderRadius:12, border:'1.5px solid #eee', background:'#fff', fontSize:20, cursor:safePage<=1?'not-allowed':'pointer', color:safePage<=1?'#ddd':'#111', display:'flex', alignItems:'center', justifyContent:'center' }}>‹</button>
              <span style={{ fontSize:14, fontWeight:700, color:'#555' }}>{safePage} <span style={{ color:'#bbb', fontWeight:400 }}>/ {totalPages}</span></span>
              <button disabled={safePage>=totalPages} onClick={()=>setPage(p=>p+1)} style={{ width:44, height:44, borderRadius:12, border:'1.5px solid #eee', background:'#fff', fontSize:20, cursor:safePage>=totalPages?'not-allowed':'pointer', color:safePage>=totalPages?'#ddd':'#111', display:'flex', alignItems:'center', justifyContent:'center' }}>›</button>
            </div>
          )}
        </div>
      )}

      {/* ═══ DESKTOP MAIN ═══ */}
      {!isMobile && (
        <main style={{ maxWidth:1400, margin:"0 auto", padding:"20px 16px" }}>
          {/* Dashboard */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:14, marginBottom:20 }}>
            <div style={{ background:"#fff", borderRadius:10, border:"1.5px solid #e5e5e5", padding:"14px 18px", textAlign:"center" }}>
              <div style={{ fontSize:11, fontWeight:700, letterSpacing:1, color:"#888", textTransform:"uppercase", marginBottom:4 }}>Total</div>
              <div style={{ fontSize:36, fontWeight:700, color:"#111" }}>{clients.length}</div>
              <div style={{ fontSize:12, color:"#bbb", marginTop:3 }}>dans la base</div>
            </div>
            <div style={{ background:"#fff", borderRadius:10, border:"1.5px solid #e5e5e5", padding:"14px 18px", textAlign:"center" }}>
              <div style={{ fontSize:11, fontWeight:700, letterSpacing:1, color:"#888", textTransform:"uppercase", marginBottom:4 }}>Aujourd'hui</div>
              <div style={{ fontSize:22, fontWeight:700, color:"#111", paddingTop:7 }}>{new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit",year:"numeric"})}</div>
            </div>
            <div style={{ background:"#fff", borderRadius:10, border:"1.5px solid #e5e5e5", padding:"14px 18px", textAlign:"center" }}>
              <div style={{ fontSize:11, fontWeight:700, letterSpacing:1, color:"#888", textTransform:"uppercase", marginBottom:4 }}>{`Nouveaux — ${getCurrentMonthName().toUpperCase()}`}</div>
              <div style={{ fontSize:36, fontWeight:700, color:G }}>{newMonth}</div>
              <div style={{ fontSize:12, color:"#bbb", marginTop:3 }}>ce mois-ci</div>
            </div>
          </div>

          {/* Tabs */}
          <div style={{ display:"flex", gap:0, marginBottom:16, background:"#f0f0f0", borderRadius:10, padding:3, width:"fit-content" }}>
            {[
              { id:"tous", label:"👥 Tous", count: clients.length },
              { id:"particuliers", label:"🙍 Particuliers", count: clients.filter(c=>c.genre!=="Entreprise").length },
              { id:"entreprises", label:"🏢 Entreprises", count: clients.filter(c=>c.genre==="Entreprise").length }
            ].map(tab => (
              <button key={tab.id} onClick={()=>{setActiveTab(tab.id);setPage(1)}} style={{
                background: activeTab===tab.id ? "#111" : "transparent",
                color: activeTab===tab.id ? "#fff" : "#666",
                border:"none", borderRadius:8, padding:"8px 16px", fontSize:13,
                fontWeight: activeTab===tab.id ? 700 : 400,
                cursor:"pointer", whiteSpace:"nowrap", transition:"all 0.15s"
              }}>
                {tab.label} <span style={{ background: activeTab===tab.id ? G : "#ddd", color: activeTab===tab.id ? "#111" : "#888", borderRadius:99, padding:"1px 7px", fontSize:11, fontWeight:700, marginLeft:4 }}>{tab.count}</span>
              </button>
            ))}
          </div>

          {/* Search + Add */}
          <div style={{ display:"flex", flexWrap:"wrap", gap:10, alignItems:"center", marginBottom:12 }}>
            <div style={{ position:"relative", flex:1, minWidth:220 }}>
              <span style={{ position:"absolute", left:12, top:"50%", transform:"translateY(-50%)", color:"#bbb", fontSize:15, pointerEvents:"none" }}>🔍</span>
              <input style={{ width:"100%", height:40, border:"1.5px solid #ddd", borderRadius:8, padding:"0 36px 0 40px", fontSize:13, background:"#fff", outline:"none", boxSizing:"border-box" }} value={search} onChange={e=>{setSearch(e.target.value);setPage(1)}} placeholder="Rechercher par nom, prénom, téléphone, mail, date, mois, genre ou commentaire…" />
              {search && <button onClick={()=>{setSearch("");setPage(1)}} style={{ position:"absolute", right:10, top:"50%", transform:"translateY(-50%)", background:"none", border:"none", cursor:"pointer", color:"#aaa", fontSize:16, padding:2 }}>✕</button>}
            </div>
            <button onClick={()=>setModalAdd(true)} style={btnPrimary}>+ Ajouter un client</button>
          </div>

          {/* Filters desktop */}
          <div style={{ display:"flex", flexWrap:"wrap", gap:8, alignItems:"center", marginBottom:14 }}>
            <select style={sel} value={filterGenre} onChange={e=>{setFilterGenre(e.target.value);setPage(1)}}><option value="">Tous les genres</option>{GENRES.map(g=><option key={g}>{g}</option>)}</select>
            <select style={sel} value={filterMonth} onChange={e=>{setFilterMonth(e.target.value);setPage(1)}}><option value="">Tous les mois</option>{MONTHS_FR.map((m,i)=><option key={i+1} value={i+1}>{m}</option>)}</select>
            <select style={sel} value={pageSize} onChange={e=>{setPageSize(Number(e.target.value));setPage(1)}}>{PAGE_SIZES.map(n=><option key={n} value={n}>{n} par page</option>)}</select>
            {(filterGenre||filterMonth||search) && <button onClick={()=>{setFilterGenre("");setFilterMonth("");setSearch("");setPage(1)}} style={{ ...btnSecondary, fontSize:12 }}>✕ Réinitialiser</button>}
            <div style={{ marginLeft:"auto", display:"flex", gap:8 }}>
              <button onClick={()=>exportToCSV(filtered)} style={btnSecondary}>⬇ CSV</button>
              <button onClick={()=>exportToXLSX(filtered)} style={btnSecondary}>⬇ Excel</button>
              <button onClick={()=>setModalImport(true)} style={btnSecondary}>⬆ Importer</button>
            </div>
          </div>

          {/* Table desktop */}
          <div style={{ background:"#fff", borderRadius:10, border:"1.5px solid #e5e5e5", overflow:"hidden" }}>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                <thead>
                  <tr>
                    <Th col="genre" label="Genre"/>
                    {activeTab === "entreprises" ? (
                      <>
                        <Th col="entreprise" label="Entreprise"/>
                        <Th col="nom" label="Contact"/>
                      </>
                    ) : (
                      <>
                        <Th col="nom" label="Nom"/>
                        <Th col="prenom" label="Prénom"/>
                      </>
                    )}
                    <Th col="tel" label="Téléphone"/>
                    <Th col="mail" label="Mail"/>
                    <Th col="created_at" label="Date d'ajout"/>
                    <th style={{ background:"#111", color:"#fff", padding:"10px 12px", textAlign:"left", fontWeight:600, fontSize:12 }}>Commentaire</th>
                    <th style={{ background:"#111", color:"#fff", padding:"10px 12px", textAlign:"left", fontWeight:600, fontSize:12 }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {pageClients.length === 0 && (
                    <tr><td colSpan={8} style={{ textAlign:"center", padding:"3rem", color:"#bbb", fontSize:14 }}>{(search||filterGenre||filterMonth)?"Aucun client trouvé":"Aucun client dans la base"}</td></tr>
                  )}
                  {pageClients.map((c, i) => {
                    const isHov = hoverRow === c.id;
                    const bg = isHov?"#fffbea":i%2===0?"#fff":"#f9f9f9";
                    const td = { padding:"9px 12px", borderBottom:"1px solid #f0f0f0", verticalAlign:"middle", background:bg };
                    return (
                      <tr key={c.id} onMouseEnter={()=>setHoverRow(c.id)} onMouseLeave={()=>setHoverRow(null)}>
                        <td style={td}><span style={badge(c.genre)}>{c.genre||"—"}</span></td>
                        {activeTab === "entreprises" ? (
                          <>
                            <td style={{...td,fontWeight:700,color:"#065f46"}}>{c.entreprise||"—"}</td>
                            <td style={td}>{c.nom||""} {c.prenom||""}</td>
                          </>
                        ) : (
                          <>
                            <td style={{...td,fontWeight:600}}>{c.genre==="Entreprise" ? <span style={{color:'#065f46',fontWeight:700}}>{c.entreprise||"—"}</span> : c.nom||"—"}</td>
                            <td style={td}>{c.genre==="Entreprise" ? <span style={{fontSize:11,color:'#999'}}>{c.nom} {c.prenom}</span> : c.prenom||"—"}</td>
                          </>
                        )}
                        <td style={{...td,fontFamily:"'Courier New',monospace"}}>{c.tel||"—"}</td>
                        <td style={{...td,fontSize:12,color:"#3b82f6",maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.mail||"—"}</td>
                        <td style={{...td,whiteSpace:"nowrap"}}>{formatDate(c.created_at)}</td>
                        <td style={{...td,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",cursor:c.commentaire?"pointer":"default",color:c.commentaire?"#555":"#ccc"}} onClick={()=>c.commentaire&&setModalComment(c)} title={c.commentaire||""}>{c.commentaire||"—"}</td>
                        <td style={{...td,whiteSpace:"nowrap"}}>
                          <button onClick={()=>setModalEdit(c)} style={{ background:"none", border:"none", cursor:"pointer", borderRadius:5, padding:"3px 6px", fontSize:16, color:"#3b82f6" }} title="Modifier">✏️</button>
                          <button onClick={()=>setModalDelete(c)} style={{ background:"none", border:"none", cursor:"pointer", borderRadius:5, padding:"3px 6px", fontSize:16, color:"#dc2626" }} title="Supprimer">🗑</button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {/* Pagination desktop */}
            <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"12px 16px", borderTop:"1px solid #eee", flexWrap:"wrap", gap:8 }}>
              <span style={{ fontSize:12, color:"#999" }}>{filtered.length===0?"0 résultat":`${(safePage-1)*pageSize+1}–${Math.min(safePage*pageSize,filtered.length)} sur ${filtered.length} client(s)`}</span>
              <div style={{ display:"flex", gap:4 }}>
                {[["«",1],["‹",safePage-1]].map(([label,p])=><button key={label} disabled={safePage<=1} onClick={()=>setPage(p)} style={{ height:30, minWidth:30, border:"1.5px solid #ddd", borderRadius:6, background:"#fff", cursor:safePage<=1?"not-allowed":"pointer", fontSize:12, padding:"0 8px", color:safePage<=1?"#ccc":"#333" }}>{label}</button>)}
                {Array.from({length:Math.min(5,totalPages)},(_,i)=>{let p=i+1;if(totalPages>5){if(safePage<=3)p=i+1;else if(safePage>=totalPages-2)p=totalPages-4+i;else p=safePage-2+i}return <button key={p} onClick={()=>setPage(p)} style={{ height:30, minWidth:30, border:`1.5px solid ${p===safePage?G:"#ddd"}`, borderRadius:6, background:p===safePage?G:"#fff", cursor:"pointer", fontSize:12, fontWeight:p===safePage?700:400, padding:"0 8px" }}>{p}</button>})}
                {[["›",safePage+1],["»",totalPages]].map(([label,p])=><button key={label} disabled={safePage>=totalPages} onClick={()=>setPage(p)} style={{ height:30, minWidth:30, border:"1.5px solid #ddd", borderRadius:6, background:"#fff", cursor:safePage>=totalPages?"not-allowed":"pointer", fontSize:12, padding:"0 8px", color:safePage>=totalPages?"#ccc":"#333" }}>{label}</button>)}
              </div>
            </div>
          </div>
        </main>
      )}

      {/* Barre nav fixe mobile */}
      {isMobile && (
        <>
          <div style={{ position:'fixed', bottom:0, left:0, right:0, background:'#fff', borderRadius:'20px 20px 0 0', boxShadow:'0 -4px 20px rgba(0,0,0,0.10)', display:'flex', alignItems:'center', justifyContent:'space-around', height:65, marginBottom:0, paddingBottom:0, zIndex:1000, paddingLeft:16, paddingRight:16 }}>

            {/* Gauche — Clients */}
            <button onClick={()=>setMobileTab('clients')} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:4, border:'none', background:'none', cursor:'pointer', color: mobileTab==='clients' ? '#111' : '#bbb', fontWeight: mobileTab==='clients' ? 700 : 500, fontSize:11 }}>
              <svg width="24" height="24" fill="none" viewBox="0 0 24 24">
                <path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z" fill="currentColor"/>
              </svg>
              Clients
            </button>

            {/* Centre — + surélevé */}
            <div style={{ position:'relative', marginTop:-28, flex:'0 0 auto' }}>
              <button onClick={()=>setShowPlusSheet(true)} style={{ width:60, height:60, borderRadius:'50%', background:G, border:'4px solid #fff', boxShadow:'0 4px 16px rgba(232,197,71,0.5)', display:'flex', alignItems:'center', justifyContent:'center', cursor:'pointer', fontSize:32, fontWeight:300, color:'#111', lineHeight:1 }}>+</button>
            </div>

            {/* Droite — Réservations */}
            <button onClick={()=>setMobileTab('reservations')} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:4, border:'none', background:'none', cursor:'pointer', color: mobileTab==='reservations' ? '#111' : '#bbb', fontWeight: mobileTab==='reservations' ? 700 : 500, fontSize:11, position:'relative' }}>
              <div style={{ position:'relative' }}>
                <svg width="24" height="24" fill="none" viewBox="0 0 24 24">
                  <path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z" fill="currentColor"/>
                </svg>
                {resaAttenteCount > 0 && (
                  <div style={{ position:'absolute', top:-8, right:-10, background:'#dc2626', color:'#fff', borderRadius:99, minWidth:22, height:22, display:'flex', alignItems:'center', justifyContent:'center', fontSize:12, fontWeight:800, border:'2px solid #fff', padding:'0 5px', boxSizing:'border-box' }}>{resaAttenteCount}</div>
                )}
              </div>
              Réservations
            </button>
          </div>
          <div style={{ position:'fixed', bottom:0, left:0, right:0, height:'env(safe-area-inset-bottom, 0px)', background:'#fff', zIndex:999 }} />

          {/* Bottom sheet + */}
          {showPlusSheet && (
            <>
              <div onPointerDown={()=>setShowPlusSheet(false)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:2999 }} />
              <div style={{ position:'fixed', bottom:0, left:0, right:0, background:'#fff', borderRadius:'20px 20px 0 0', zIndex:3000, paddingTop:20, paddingBottom:'calc(16px + env(safe-area-inset-bottom))' }}>
                <div style={{ width:40, height:4, background:'#e5e5e5', borderRadius:99, margin:'0 auto 16px' }} />
                <button onClick={()=>{ setShowPlusSheet(false); setModalAdd(true); }} style={{ display:'flex', alignItems:'center', gap:12, padding:'14px 16px', border:'none', background:'none', width:'100%', cursor:'pointer', borderBottom:'1px solid #f0f0f0' }}>
                  <div style={{ width:44, height:44, background:'#f0fdf4', borderRadius:12, display:'flex', alignItems:'center', justifyContent:'center', fontSize:22, flexShrink:0 }}>👤</div>
                  <div style={{ textAlign:'left' }}>
                    <div style={{ fontWeight:700, fontSize:14, color:'#111' }}>Ajouter un client</div>
                    <div style={{ fontSize:12, color:'#888' }}>Créer une nouvelle fiche client</div>
                  </div>
                </button>
                <button onClick={()=>{ setShowPlusSheet(false); setShowAddResa(true); }} style={{ display:'flex', alignItems:'center', gap:12, padding:'14px 16px', border:'none', background:'none', width:'100%', cursor:'pointer' }}>
                  <div style={{ width:44, height:44, background:'#fffbea', borderRadius:12, display:'flex', alignItems:'center', justifyContent:'center', fontSize:22, flexShrink:0 }}>📅</div>
                  <div style={{ textAlign:'left' }}>
                    <div style={{ fontWeight:700, fontSize:14, color:'#111' }}>Ajouter une réservation</div>
                    <div style={{ fontSize:12, color:'#888' }}>Saisie manuelle par téléphone</div>
                  </div>
                </button>
              </div>
            </>
          )}
        </>
      )}

      {/* Menu ••• fixe positionné au bouton */}
      {mobileAction && isMobile && (() => {
        const r = mobileAction._rect;
        const menuW = 170;
        const left = Math.min(r.right - menuW, window.innerWidth - menuW - 8);
        const top = r.bottom + 6;
        return (
          <>
            <div onPointerDown={()=>setMobileAction(null)} style={{ position:'fixed', inset:0, zIndex:300 }} />
            <div style={{ position:'fixed', top, left, width:menuW, background:'#fff', borderRadius:12, boxShadow:'0 6px 24px rgba(0,0,0,0.18)', zIndex:301, overflow:'hidden', border:'1px solid #f0f0f0' }}>
              <button onPointerDown={()=>{ setModalEdit(mobileAction); setMobileAction(null); }} style={{ width:'100%', padding:'13px 16px', background:'none', border:'none', borderBottom:'1px solid #f5f5f5', fontSize:14, fontWeight:600, color:'#1d4ed8', cursor:'pointer', textAlign:'left', display:'flex', alignItems:'center', gap:10 }}>✏️ Modifier</button>
              <button onPointerDown={()=>{ setModalDelete(mobileAction); setMobileAction(null); }} style={{ width:'100%', padding:'13px 16px', background:'none', border:'none', fontSize:14, fontWeight:600, color:'#dc2626', cursor:'pointer', textAlign:'left', display:'flex', alignItems:'center', gap:10 }}>🗑 Supprimer</button>
            </div>
          </>
        );
      })()}


      {/* Modals */}
      {showAddResa && <AddResaModal onClose={()=>setShowAddResa(false)} onSaved={()=>{ loadResaCount(); }} showToast={showToast} user={user} />}
      {modalAdd && <ClientForm existingClients={clients} onSave={addClient} onCancel={()=>setModalAdd(false)} />}
      {modalEdit && <ClientForm initial={modalEdit} existingClients={clients} onSave={editClient} onCancel={()=>setModalEdit(null)} />}
      {modalDelete && <ConfirmModal title="Supprimer ce client ?" msg={`Êtes-vous sûr de vouloir supprimer définitivement ${modalDelete.prenom} ${modalDelete.nom} ? Cette action est irréversible.`} onOk={()=>deleteClient(modalDelete.id)} onCancel={()=>setModalDelete(null)} okLabel="Supprimer définitivement" danger />}
      {modalImport && <ImportModal existingClients={clients} onImport={importClients} onCancel={()=>setModalImport(false)} />}
      {modalComment && <Modal title={`Commentaire — ${modalComment.prenom} ${modalComment.nom}`} onClose={()=>setModalComment(null)}><p style={{fontSize:14,lineHeight:1.7,margin:0}}>{modalComment.commentaire}</p></Modal>}
      {modalCorbeille && !isMobile && <CorbeilleModal onClose={()=>{ setModalCorbeille(false); loadClients(); }} showToast={showToast} />}
      {toast && <Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)} />}
    </div>
  );
}

// ─── Root App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState(null);
  const [checking, setChecking] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setUser(session?.user ?? null);
      setChecking(false);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null);
    });
    return () => subscription.unsubscribe();
  }, []);

  async function handleLogout() {
    await supabase.auth.signOut();
    setUser(null);
  }

  if (checking) return <div style={{ textAlign:"center", paddingTop:80, fontSize:16, color:"#888" }}>Chargement…</div>;
  if (!user) return <LoginPage onLogin={()=>{}} />;
  return <CRMApp user={user} onLogout={handleLogout} />;
}
