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
function Modal({ title, onClose, children, footer, maxW=520 }) {
  const isMobile = window.innerWidth < 768;
  return (
    <div
      style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:1000, display:"flex", alignItems: isMobile ? "flex-end" : "center", justifyContent:"center", padding: isMobile ? 0 : "1rem" }}
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
  const deleteGuard = useRef(false);
  const isMobile = useIsMobile();

  const showToast = useCallback((msg, type="success") => setToast({msg,type}), []);

  // ─── Load from Supabase ───────────────────────────────────────────────────
  useEffect(() => {
    loadClients();
  }, []);

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

      {/* ═══ MOBILE HEADER (fixed) ═══ */}
      {isMobile && (
        <header style={{ position:'fixed', top:0, left:0, right:0, height:56, background:'#111', zIndex:200, display:'flex', alignItems:'center', justifyContent:'space-between', padding:'0 14px', borderBottom:`3px solid ${G}` }}>
          <div style={{ display:'flex', alignItems:'center', gap:8 }}>
            <img src={require('./logo.png')} alt="TED" style={{ height:28, filter:'brightness(0) invert(1)', verticalAlign:'middle' }} />
            <span style={{ color:'#fff', fontWeight:700, fontSize:14, letterSpacing:1 }}><span style={{color:G}}>TED</span> CRM</span>
          </div>
          <div style={{ display:'flex', gap:6, alignItems:'center' }}>
            <button type="button" onPointerDown={()=>setShowSearch(s=>!s)} className="btn-mobile" style={{ background:showSearch?G:'rgba(255,255,255,0.1)', border:'none', borderRadius:10, width:38, height:38, fontSize:18, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', touchAction:'manipulation' }}>🔍</button>
            <button type="button" onPointerDown={onLogout} className="btn-mobile" style={{ background:'rgba(255,255,255,0.08)', border:'1px solid #333', borderRadius:10, height:38, padding:'0 12px', fontSize:12, color:'#ccc', cursor:'pointer', fontWeight:500, touchAction:'manipulation' }}>✕</button>
          </div>
        </header>
      )}

      {/* ═══ MOBILE SEARCH BAR (slides down) ═══ */}
      {isMobile && showSearch && (
        <div style={{ position:'fixed', top:59, left:0, right:0, zIndex:190, background:'#fff', padding:'10px 14px', borderBottom:'1px solid #eee', boxShadow:'0 4px 20px rgba(0,0,0,0.1)', animation:'slideDownFade 0.18s ease' }}>
          <div style={{ position:'relative' }}>
            <span style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', color:'#bbb', fontSize:16, pointerEvents:'none' }}>🔍</span>
            <input
              autoFocus
              style={{ width:'100%', height:44, border:'1.5px solid #e5e5e5', borderRadius:12, padding:'0 40px 0 42px', fontSize:16, background:'#f8f8f8', outline:'none', boxSizing:'border-box' }}
              value={search}
              onChange={e=>{setSearch(e.target.value);setPage(1)}}
              placeholder="Nom, téléphone, mail…"
            />
            {search && <button type="button" onPointerDown={()=>{setSearch("");setPage(1)}} style={{ position:'absolute', right:10, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', cursor:'pointer', color:'#aaa', fontSize:18, padding:4, touchAction:'manipulation' }}>✕</button>}
          </div>
        </div>
      )}

      {/* ═══ DESKTOP HEADER ═══ */}
      {!isMobile && (
        <header style={{ background:"#111", color:"#fff", padding:"0 20px", display:"flex", alignItems:"center", justifyContent:"space-between", height:56, borderBottom:`3px solid ${G}` }}>
          <h1 style={{ fontSize:16, fontWeight:700, letterSpacing:2, color:"#fff", margin:0 }}>
            <img src={require('./logo.png')} alt="TED" style={{height:32, marginRight:8, verticalAlign:'middle', filter:'brightness(0) invert(1)'}} />
            <span style={{color:G}}>TED</span> — FICHIER CLIENTS
          </h1>
          <div style={{ display:"flex", gap:8, alignItems:"center" }}>
            <span style={{ fontSize:12, color:"#888", marginRight:4 }}>{user.email}</span>
            <button onClick={saveBackup} style={{ ...btnSecondary, background:"transparent", color:"#ccc", border:"1px solid #444", height:32, fontSize:12 }}>💾 Sauvegarder</button>
            <button onClick={()=>restoreRef.current?.click()} style={{ ...btnSecondary, background:"transparent", color:"#ccc", border:"1px solid #444", height:32, fontSize:12 }}>🔄 Restaurer</button>
            <input ref={restoreRef} type="file" accept=".json" style={{display:"none"}} onChange={handleRestoreFile} />
            <button onClick={()=>setModalCorbeille(true)} style={{ ...btnSecondary, background:"transparent", color:"#ccc", border:"1px solid #444", height:32, fontSize:12 }}>🗑 Corbeille</button>
            <button onClick={onLogout} style={{ ...btnSecondary, background:"transparent", color:"#ccc", border:"1px solid #444", height:32, fontSize:12 }}>Déconnexion</button>
          </div>
        </header>
      )}
      {isMobile && <input ref={restoreRef} type="file" accept=".json" style={{display:"none"}} onChange={handleRestoreFile} />}

      {/* ═══ MOBILE CONTENT ═══ */}
      {isMobile && (
        <div style={{ paddingTop: showSearch ? 116 : 68, paddingBottom: 80, paddingLeft:12, paddingRight:12 }}>

          {/* Dashboard cards */}
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:10, marginBottom:16 }}>
            <div style={{ background:'#fff', borderRadius:14, border:'1.5px solid #efefef', padding:'12px 8px', textAlign:'center', boxShadow:'0 1px 4px rgba(0,0,0,0.06)' }}>
              <div style={{ fontSize:20, marginBottom:4 }}>👥</div>
              <div style={{ fontSize:28, fontWeight:800, color:'#111', lineHeight:1 }}>{clients.length}</div>
              <div style={{ fontSize:9, fontWeight:700, color:'#aaa', letterSpacing:0.5, textTransform:'uppercase', marginTop:4 }}>Total</div>
            </div>
            <div style={{ background:'#fff', borderRadius:14, border:'1.5px solid #efefef', padding:'12px 8px', textAlign:'center', boxShadow:'0 1px 4px rgba(0,0,0,0.06)' }}>
              <div style={{ fontSize:20, marginBottom:4 }}>✨</div>
              <div style={{ fontSize:28, fontWeight:800, color:G, lineHeight:1 }}>{newMonth}</div>
              <div style={{ fontSize:9, fontWeight:700, color:'#aaa', letterSpacing:0.5, textTransform:'uppercase', marginTop:4 }}>Ce mois</div>
            </div>
            <div style={{ background:'#fff', borderRadius:14, border:'1.5px solid #efefef', padding:'12px 8px', textAlign:'center', boxShadow:'0 1px 4px rgba(0,0,0,0.06)' }}>
              <div style={{ fontSize:20, marginBottom:4 }}>📅</div>
              <div style={{ fontSize:11, fontWeight:700, color:'#111', lineHeight:1.3, marginTop:2 }}>{new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit",year:"2-digit"})}</div>
              <div style={{ fontSize:9, fontWeight:700, color:'#aaa', letterSpacing:0.5, textTransform:'uppercase', marginTop:4 }}>Aujourd'hui</div>
            </div>
          </div>

          {/* Tabs pills */}
          <div style={{ display:'flex', gap:8, marginBottom:14, overflowX:'auto', paddingBottom:2 }}>
            {[
              { id:"tous", label:"Tous", icon:"👥", count: clients.length },
              { id:"particuliers", label:"Particuliers", icon:"🙍", count: clients.filter(c=>c.genre!=="Entreprise").length },
              { id:"entreprises", label:"Entreprises", icon:"🏢", count: clients.filter(c=>c.genre==="Entreprise").length }
            ].map(tab => (
              <button key={tab.id} type="button" onPointerDown={()=>{setActiveTab(tab.id);setPage(1)}} className="tab-pill btn-mobile" style={{
                background: activeTab===tab.id ? '#111' : '#fff',
                color: activeTab===tab.id ? '#fff' : '#666',
                border: activeTab===tab.id ? '2px solid #111' : '1.5px solid #e0e0e0',
                borderRadius:99,
                padding:'8px 14px',
                fontSize:13,
                fontWeight: activeTab===tab.id ? 700 : 500,
                cursor:'pointer',
                whiteSpace:'nowrap',
                touchAction:'manipulation',
                display:'flex',
                alignItems:'center',
                gap:5
              }}>
                <span>{tab.icon}</span>
                <span>{tab.label}</span>
                <span style={{ background: activeTab===tab.id ? G : '#eee', color: activeTab===tab.id ? '#111' : '#888', borderRadius:99, padding:'1px 7px', fontSize:10, fontWeight:700, marginLeft:2 }}>{tab.count}</span>
              </button>
            ))}
          </div>

          {/* Filters toggle */}
          <div style={{ marginBottom:12 }}>
            <div style={{ display:'flex', gap:8, marginBottom: showFilters ? 8 : 0 }}>
              <button type="button" onPointerDown={()=>setShowFilters(f=>!f)} className="btn-mobile" style={{ flex:1, height:38, border:'1.5px solid #e0e0e0', borderRadius:10, background:'#fff', fontSize:13, color:'#555', fontWeight:500, cursor:'pointer', touchAction:'manipulation' }}>
                {showFilters ? '▲ Masquer' : '▼ Filtres'}{(filterGenre||filterMonth) ? ' •' : ''}
              </button>
              {(filterGenre||filterMonth||search) && <button type="button" onPointerDown={()=>{setFilterGenre("");setFilterMonth("");setSearch("");setPage(1)}} className="btn-mobile" style={{ height:38, width:38, border:'1.5px solid #fca5a5', borderRadius:10, background:'#fff5f5', fontSize:16, cursor:'pointer', touchAction:'manipulation' }}>✕</button>}
            </div>
            {showFilters && (
              <div style={{ display:'flex', flexDirection:'column', gap:8, animation:'slideDownFade 0.15s ease' }}>
                <select style={{...sel, width:'100%', height:44, fontSize:15, borderRadius:10}} value={filterGenre} onChange={e=>{setFilterGenre(e.target.value);setPage(1)}}><option value="">Tous les genres</option>{GENRES.map(g=><option key={g}>{g}</option>)}</select>
                <select style={{...sel, width:'100%', height:44, fontSize:15, borderRadius:10}} value={filterMonth} onChange={e=>{setFilterMonth(e.target.value);setPage(1)}}><option value="">Tous les mois</option>{MONTHS_FR.map((m,i)=><option key={i+1} value={i+1}>{m}</option>)}</select>
              </div>
            )}
          </div>

          {/* Client cards */}
          {pageClients.length === 0 && (
            <div style={{ textAlign:'center', padding:'3rem', color:'#ccc', fontSize:14 }}>{(search||filterGenre||filterMonth)?"Aucun résultat":"Aucun client"}</div>
          )}
          {pageClients.map((c, idx) => (
            <div key={c.id} className="client-card card-mobile" style={{ animationDelay:`${idx*0.04}s` }}>
              <div style={{ padding:'14px 16px' }}>
                {/* Header row */}
                <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between', marginBottom:8 }}>
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ display:'flex', alignItems:'center', gap:7, flexWrap:'wrap', marginBottom:3 }}>
                      <span style={badge(c.genre)}>{c.genre}</span>
                      <span style={{ fontWeight:700, fontSize:16, color:'#111' }}>
                        {c.genre==="Entreprise" ? (c.entreprise||c.nom||"—") : `${c.nom||""} ${c.prenom||""}`.trim()||"—"}
                      </span>
                    </div>
                    {c.genre==="Entreprise" && (c.nom||c.prenom) && (
                      <div style={{ fontSize:12, color:'#999' }}>{c.nom} {c.prenom}</div>
                    )}
                  </div>
                  <div style={{ display:'flex', gap:2, flexShrink:0, marginLeft:8 }}>
                    <button type="button" onPointerDown={()=>setModalEdit(c)} className="btn-mobile" style={{ background:'#f0f4ff', border:'none', borderRadius:9, width:36, height:36, fontSize:17, cursor:'pointer', touchAction:'manipulation', display:'flex', alignItems:'center', justifyContent:'center' }}>✏️</button>
                    <button type="button" onPointerDown={()=>setModalDelete(c)} className="btn-mobile" style={{ background:'#fff0f0', border:'none', borderRadius:9, width:36, height:36, fontSize:17, cursor:'pointer', touchAction:'manipulation', display:'flex', alignItems:'center', justifyContent:'center' }}>🗑</button>
                  </div>
                </div>
                {/* Info rows */}
                {c.tel && (
                  <a href={`tel:${c.tel}`} style={{ display:'flex', alignItems:'center', gap:6, textDecoration:'none', color:'#111', fontSize:14, fontWeight:600, marginBottom:4 }}>
                    <span style={{ fontSize:15 }}>📞</span>
                    <span>{c.tel}</span>
                  </a>
                )}
                {c.mail && <div style={{ fontSize:12, color:'#3b82f6', marginBottom:4, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{c.mail}</div>}
                <div style={{ fontSize:11, color:'#ccc', marginTop:2 }}>{formatDate(c.created_at)}</div>
                {c.commentaire && <div style={{ fontSize:12, color:'#888', marginTop:6, fontStyle:'italic', borderTop:'1px solid #f5f5f5', paddingTop:6 }}>"{c.commentaire}"</div>}
              </div>
              {/* Appeler button */}
              {c.tel && (
                <a href={`tel:${c.tel}`} style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:6, height:44, background:`${G}22`, color:'#111', fontWeight:700, fontSize:14, textDecoration:'none', borderTop:'1px solid #f0f0f0' }}>
                  <span>📞</span> Appeler
                </a>
              )}
            </div>
          ))}

          {/* Pagination mobile */}
          {totalPages > 1 && (
            <div style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:16, padding:'12px 0 4px' }}>
              <button type="button" disabled={safePage<=1} onPointerDown={()=>setPage(safePage-1)} className="btn-mobile" style={{ height:44, width:44, border:'1.5px solid #e0e0e0', borderRadius:12, background:'#fff', cursor:safePage<=1?'not-allowed':'pointer', fontSize:20, color:safePage<=1?'#ccc':'#333', touchAction:'manipulation' }}>‹</button>
              <span style={{ fontSize:13, color:'#888', fontWeight:600 }}>{safePage} / {totalPages}</span>
              <button type="button" disabled={safePage>=totalPages} onPointerDown={()=>setPage(safePage+1)} className="btn-mobile" style={{ height:44, width:44, border:'1.5px solid #e0e0e0', borderRadius:12, background:'#fff', cursor:safePage>=totalPages?'not-allowed':'pointer', fontSize:20, color:safePage>=totalPages?'#ccc':'#333', touchAction:'manipulation' }}>›</button>
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

      {/* Bouton fixe mobile */}
      {isMobile && !modalAdd && !modalEdit && !modalDelete && (
        <button
          type="button"
          onPointerDown={()=>setModalAdd(true)}
          className="btn-mobile"
          style={{ position:'fixed', bottom:0, left:0, right:0, height:64, paddingBottom:'env(safe-area-inset-bottom)', background:G, color:'#111', border:'none', fontSize:16, fontWeight:800, cursor:'pointer', zIndex:500, boxShadow:'0 -4px 24px rgba(0,0,0,0.18)', letterSpacing:0.5, touchAction:'manipulation' }}
        >
          + Ajouter un client
        </button>
      )}

      {/* Modals */}
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
