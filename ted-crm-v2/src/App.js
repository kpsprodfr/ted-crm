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
const inp = (err) => ({ width:"100%", height:38, border:`1.5px solid ${err?"#dc2626":"#ddd"}`, borderRadius:7, padding:"0 10px", fontSize:13, outline:"none", boxSizing:"border-box" });
const lbl = { display:"block", fontSize:12, fontWeight:600, color:"#444", marginBottom:5 };
const fg = { marginBottom:14 };

// ─── Toast ────────────────────────────────────────────────────────────────────
function Toast({ msg, type, onClose }) {
  useEffect(() => { const t = setTimeout(onClose, 3500); return () => clearTimeout(t); }, [onClose]);
  const bc = type==="success"?"#22c55e":type==="error"?"#dc2626":"#f59e0b";
  return <div style={{ position:"fixed", bottom:24, right:24, zIndex:9999, background:"#111", color:"#fff", borderRadius:8, padding:"12px 18px", fontSize:13, fontWeight:500, borderLeft:`4px solid ${bc}`, boxShadow:"0 4px 20px rgba(0,0,0,0.25)", maxWidth:340 }}>{msg}</div>;
}

// ─── Modal ────────────────────────────────────────────────────────────────────
function Modal({ title, onClose, children, footer, maxW=520 }) {
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:1000, display:"flex", alignItems:"center", justifyContent:"center", padding:"1rem" }} onClick={onClose}>
      <div style={{ background:"#fff", borderRadius:12, width:"100%", maxWidth:maxW, overflow:"hidden", maxHeight:"90vh", display:"flex", flexDirection:"column" }} onClick={e=>e.stopPropagation()}>
        <div style={{ background:"#111", color:"#fff", padding:"14px 18px", display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0 }}>
          <span style={{ fontWeight:700, fontSize:15 }}>{title}</span>
          <button onClick={onClose} style={{ background:"none", border:"none", color:"#fff", fontSize:20, cursor:"pointer" }}>✕</button>
        </div>
        <div style={{ padding:"18px", overflowY:"auto", flex:1 }}>{children}</div>
        {footer && <div style={{ padding:"0 18px 18px", display:"flex", gap:8, justifyContent:"flex-end", flexShrink:0 }}>{footer}</div>}
      </div>
    </div>
  );
}

function ConfirmModal({ title, msg, onOk, onCancel, okLabel="Confirmer", danger=false }) {
  return (
    <Modal title={title} onClose={onCancel} maxW={400} footer={[
      <button key="c" onClick={onCancel} style={btnSecondary}>Annuler</button>,
      <button key="o" onClick={onOk} style={danger?btnDanger:btnPrimary}>{okLabel}</button>
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
  const [form, setForm] = useState({ genre:initial?.genre||"Non renseigné", nom:initial?.nom||"", prenom:initial?.prenom||"", tel:initial?.tel||"", mail:initial?.mail||"", commentaire:initial?.commentaire||"" });
  const [errors, setErrors] = useState({});
  const [dupWarn, setDupWarn] = useState(null);

  function set(k,v) { setForm(f=>({...f,[k]:v})); setErrors(e=>({...e,[k]:""})); }
  function handleTel(v) { set("tel", v.replace(/\D/g,"").slice(0,10)); }

  function validate() {
    const e = {};
    if (!form.nom.trim()) e.nom = "Le nom est obligatoire.";
    if (!form.prenom.trim()) e.prenom = "Le prénom est obligatoire.";
    if (form.tel && !/^\d{10}$/.test(form.tel)) e.tel = "Le numéro doit contenir uniquement 10 chiffres.";
    if (form.mail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(form.mail)) e.mail = "Adresse mail invalide.";
    return e;
  }

  function checkDup() {
    const others = existingClients.filter(c => !isEdit || c.id !== initial.id);
    if (form.tel && others.some(c => c.tel === form.tel)) return `Le téléphone ${form.tel} est déjà utilisé.`;
    if (form.mail && others.some(c => c.mail && c.mail.toLowerCase() === form.mail.toLowerCase())) return `L'adresse ${form.mail} est déjà utilisée.`;
    return null;
  }

  function doSave() {
    const saved = { ...(initial||{}), id:initial?.id, genre:form.genre, nom:capitalize(form.nom.trim()), prenom:capitalize(form.prenom.trim()), tel:form.tel, mail:form.mail.trim().toLowerCase(), commentaire:form.commentaire.trim(), created_at:initial?.created_at||new Date().toISOString() };
    onSave(saved);
    setDupWarn(null);
  }

  function handleSubmit() {
    const e = validate();
    if (Object.keys(e).length) { setErrors(e); return; }
    const dup = checkDup();
    if (dup) { setDupWarn(dup); return; }
    doSave();
  }

  return (
    <>
      {dupWarn && <ConfirmModal title="Doublon potentiel" msg={`Un client similaire existe déjà. ${dupWarn} Voulez-vous tout de même continuer ?`} onOk={()=>{setDupWarn(null);doSave()}} onCancel={()=>setDupWarn(null)} okLabel="Ajouter quand même" />}
      <Modal title={isEdit?"Modifier le client":"Ajouter un client"} onClose={onCancel} footer={[
        <button key="c" onClick={onCancel} style={btnSecondary}>Annuler</button>,
        <button key="s" onClick={handleSubmit} style={btnPrimary}>{isEdit?"Enregistrer les modifications":"Enregistrer"}</button>
      ]}>
        <div style={fg}>
          <label style={lbl}>Genre</label>
          <select style={{ ...inp(false), height:38 }} value={form.genre} onChange={e=>set("genre",e.target.value)}>
            {GENRES.map(g=><option key={g}>{g}</option>)}
          </select>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
          <div style={fg}>
            <label style={lbl}>Nom <span style={{color:"#dc2626"}}>*</span></label>
            <input style={inp(errors.nom)} value={form.nom} onChange={e=>set("nom",e.target.value)} placeholder="Dupont" />
            {errors.nom && <p style={{ fontSize:11, color:"#dc2626", marginTop:4 }}>{errors.nom}</p>}
          </div>
          <div style={fg}>
            <label style={lbl}>Prénom <span style={{color:"#dc2626"}}>*</span></label>
            <input style={inp(errors.prenom)} value={form.prenom} onChange={e=>set("prenom",e.target.value)} placeholder="Jean" />
            {errors.prenom && <p style={{ fontSize:11, color:"#dc2626", marginTop:4 }}>{errors.prenom}</p>}
          </div>
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
          <div style={fg}>
            <label style={lbl}>Téléphone</label>
            <input style={inp(errors.tel)} value={form.tel} onChange={e=>handleTel(e.target.value)} inputMode="numeric" placeholder="0612345678" maxLength={10} />
            {errors.tel && <p style={{ fontSize:11, color:"#dc2626", marginTop:4 }}>{errors.tel}</p>}
          </div>
          <div style={fg}>
            <label style={lbl}>Mail</label>
            <input style={inp(errors.mail)} value={form.mail} onChange={e=>set("mail",e.target.value)} placeholder="exemple@mail.fr" type="email" />
            {errors.mail && <p style={{ fontSize:11, color:"#dc2626", marginTop:4 }}>{errors.mail}</p>}
          </div>
        </div>
        <div style={fg}>
          <label style={lbl}>Commentaire</label>
          <textarea style={{ width:"100%", border:"1.5px solid #ddd", borderRadius:7, padding:"8px 10px", fontSize:13, outline:"none", boxSizing:"border-box", resize:"vertical", minHeight:65 }} value={form.commentaire} onChange={e=>set("commentaire",e.target.value)} placeholder="Notes sur ce client…" />
        </div>
        {isEdit && <p style={{ fontSize:12, color:"#999", margin:0 }}>Date d'ajout : {formatDate(initial.created_at)} — non modifiable</p>}
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

// ─── Main CRM App ─────────────────────────────────────────────────────────────
const mobileCSS = `
  @media (max-width: 768px) {
    header { padding: 0 12px !important; height: 50px !important; }
    header h1 { font-size: 13px !important; letter-spacing: 1px !important; }
    header img { height: 26px !important; }
    main { padding: 12px 8px !important; }
    .dash-grid { grid-template-columns: 1fr 1fr !important; gap: 8px !important; margin-bottom: 12px !important; }
    table { font-size: 11px !important; }
    th, td { padding: 6px 8px !important; }
    .search-bar { font-size: 12px !important; }
    .filters-row { gap: 6px !important; }
    .filters-row select { font-size: 11px !important; height: 32px !important; padding: 0 6px !important; }
    .export-btns { display: none !important; }
    button { font-size: 12px !important; }
  }
`;

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
  const deleteGuard = useRef(false);

  const showToast = useCallback((msg, type="success") => setToast({msg,type}), []);

  // ─── Load from Supabase ───────────────────────────────────────────────────
  useEffect(() => {
    loadClients();
  }, []);

  async function loadClients() {
    setLoading(true);
    const { data, error } = await supabase.from("clients").select("*").order("created_at", { ascending: false });
    if (error) { showToast("Erreur de chargement", "error"); }
    else { setClients(data || []); }
    setLoading(false);
  }

  // ─── CRUD ─────────────────────────────────────────────────────────────────
  async function addClient(c) {
    const { data, error } = await supabase.from("clients").insert([{ genre:c.genre, nom:c.nom, prenom:c.prenom, tel:c.tel, mail:c.mail, commentaire:c.commentaire, created_at:c.created_at }]).select().single();
    if (error) { showToast("Erreur lors de l'ajout", "error"); return; }
    setClients(prev => [data, ...prev]);
    setModalAdd(false);
    showToast("Client ajouté avec succès ✓");
    setPage(1);
  }

  async function editClient(c) {
    const { error } = await supabase.from("clients").update({ genre:c.genre, nom:c.nom, prenom:c.prenom, tel:c.tel, mail:c.mail, commentaire:c.commentaire }).eq("id", c.id);
    if (error) { showToast("Erreur lors de la modification", "error"); return; }
    setClients(prev => prev.map(x => x.id === c.id ? {...x, ...c} : x));
    setModalEdit(null);
    showToast("Client modifié avec succès ✓");
  }

  async function deleteClient(id) {
    if (deleteGuard.current) return;
    deleteGuard.current = true;
    const { error } = await supabase.from("clients").delete().eq("id", id);
    if (error) { showToast("Erreur lors de la suppression", "error"); deleteGuard.current = false; return; }
    setClients(prev => prev.filter(x => x.id !== id));
    setModalDelete(null);
    showToast("Client supprimé ✓");
    setTimeout(() => { deleteGuard.current = false; }, 1000);
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
  }, [clients, search, filterGenre, filterMonth, sortKey, sortDir]);

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
      <style>{mobileCSS}</style>
      {/* Header */}
      <header style={{ background:"#111", color:"#fff", padding:"0 20px", display:"flex", alignItems:"center", justifyContent:"space-between", height:56, borderBottom:`3px solid ${G}` }}>
        <h1 style={{ fontSize:16, fontWeight:700, letterSpacing:2, color:"#fff", margin:0 }}><img src={require('./logo.png')} alt="TED" style={{height:32, marginRight:10, verticalAlign:'middle', filter:'brightness(0) invert(1)'}} /><span style={{color:G}}>TED</span> — FICHIER CLIENTS</h1>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <span style={{ fontSize:12, color:"#888", marginRight:4 }}>{user.email}</span>
          <button onClick={saveBackup} style={{ ...btnSecondary, background:"transparent", color:"#ccc", border:"1px solid #444", height:32, fontSize:12 }}>💾 Sauvegarder</button>
          <button onClick={()=>restoreRef.current?.click()} style={{ ...btnSecondary, background:"transparent", color:"#ccc", border:"1px solid #444", height:32, fontSize:12 }}>🔄 Restaurer</button>
          <input ref={restoreRef} type="file" accept=".json" style={{display:"none"}} onChange={handleRestoreFile} />
          <button onClick={onLogout} style={{ ...btnSecondary, background:"transparent", color:"#ccc", border:"1px solid #444", height:32, fontSize:12 }}>Déconnexion</button>
        </div>
      </header>

      <main style={{ maxWidth:1400, margin:"0 auto", padding:"20px 16px" }}>
        {/* Dashboard */}
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))", gap:14, marginBottom:20 }}>
          <div style={{ background:"#fff", borderRadius:10, border:"1.5px solid #e5e5e5", padding:"14px 18px", textAlign:"center" }}>
            <div style={{ fontSize:11, fontWeight:700, letterSpacing:1.5, color:"#888", textTransform:"uppercase", marginBottom:6 }}>Total clients</div>
            <div style={{ fontSize:36, fontWeight:700, color:"#111" }}>{clients.length}</div>
            <div style={{ fontSize:12, color:"#bbb", marginTop:3 }}>dans la base</div>
          </div>
          <div style={{ background:"#fff", borderRadius:10, border:"1.5px solid #e5e5e5", padding:"14px 18px", textAlign:"center" }}>
            <div style={{ fontSize:11, fontWeight:700, letterSpacing:1.5, color:"#888", textTransform:"uppercase", marginBottom:6 }}>Date du jour</div>
            <div style={{ fontSize:22, fontWeight:700, color:"#111", paddingTop:7 }}>{new Date().toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit",year:"numeric"})}</div>
          </div>
          <div style={{ background:"#fff", borderRadius:10, border:"1.5px solid #e5e5e5", padding:"14px 18px", textAlign:"center" }}>
            <div style={{ fontSize:11, fontWeight:700, letterSpacing:1.5, color:"#888", textTransform:"uppercase", marginBottom:6 }}>Nouveaux — {getCurrentMonthName().toUpperCase()}</div>
            <div style={{ fontSize:36, fontWeight:700, color:G }}>{newMonth}</div>
            <div style={{ fontSize:12, color:"#bbb", marginTop:3 }}>ce mois-ci</div>
          </div>
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

        {/* Filters */}
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

        {/* Table */}
        <div style={{ background:"#fff", borderRadius:10, border:"1.5px solid #e5e5e5", overflow:"hidden" }}>
          <div style={{ overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
              <thead>
                <tr>
                  <Th col="genre" label="Genre"/>
                  <Th col="nom" label="Nom"/>
                  <Th col="prenom" label="Prénom"/>
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
                      <td style={{...td,fontWeight:600}}>{c.nom||"—"}</td>
                      <td style={td}>{c.prenom||"—"}</td>
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

          {/* Pagination */}
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

      {/* Modals */}
      {modalAdd && <ClientForm existingClients={clients} onSave={addClient} onCancel={()=>setModalAdd(false)} />}
      {modalEdit && <ClientForm initial={modalEdit} existingClients={clients} onSave={editClient} onCancel={()=>setModalEdit(null)} />}
      {modalDelete && <ConfirmModal title="Supprimer ce client ?" msg={`Êtes-vous sûr de vouloir supprimer définitivement ${modalDelete.prenom} ${modalDelete.nom} ? Cette action est irréversible.`} onOk={()=>deleteClient(modalDelete.id)} onCancel={()=>setModalDelete(null)} okLabel="Supprimer définitivement" danger />}
      {modalImport && <ImportModal existingClients={clients} onImport={importClients} onCancel={()=>setModalImport(false)} />}
      {modalComment && <Modal title={`Commentaire — ${modalComment.prenom} ${modalComment.nom}`} onClose={()=>setModalComment(null)}><p style={{fontSize:14,lineHeight:1.7,margin:0}}>{modalComment.commentaire}</p></Modal>}
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
