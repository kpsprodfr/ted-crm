import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import { Mail, LockKeyhole, Eye, EyeOff, RefreshCw, ShieldCheck, MonitorSmartphone, Headphones, ArrowRight, AlertCircle, Users, UtensilsCrossed, Phone, Download, CalendarDays, Megaphone, Link, LogOut, Copy, ExternalLink, Share2, ClipboardList, CircleCheck, User, ChevronRight, ChevronDown, Pencil, Sun, Moon, ArrowLeft, MessageSquare, UserX, Clock, Star, Trash2, Send, History, Building2, CheckCircle, Check, Search, RotateCcw, Save, Plus, UserPlus, Trophy, ArrowUpDown } from 'lucide-react';
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

function exportToCSV(clients, opts = {}) {
  const { filtreLabel = 'Tous les clients', recherche = '' } = opts;
  const header = ["Genre","Entreprise","Nom","Prénom","Téléphone","Mail","Date d'ajout","Commentaire"];
  const rows = clients.map(c => [c.genre, c.genre==='Entreprise'?(c.entreprise||''):'', c.nom,c.prenom,c.tel,c.mail,formatDate(c.created_at),c.commentaire].map(v => `"${(v||"").replace(/"/g,'""')}"`));
  const now = new Date();
  const dateLabel = now.toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long',year:'numeric'}) + ' à ' + now.toLocaleTimeString('fr-FR',{hour:'2-digit',minute:'2-digit'});
  const recapLines = [
    '',
    '---,---,---,---,---,---,---,---',
    'RÉCAPITULATIF,,,,,,,',
    `"Date d'export";"${dateLabel}";;;;;;`,
    `"Total exporté";"${clients.length} client${clients.length>1?'s':''}";;;;;;`,
    `"Filtre appliqué";"${filtreLabel}";;;;;;`,
    recherche ? `"Recherche appliquée";"${recherche}";;;;;;` : '',
    `"Hommes";"${clients.filter(c=>c.genre==='Homme').length}";;;;;;`,
    `"Femmes";"${clients.filter(c=>c.genre==='Femme').length}";;;;;;`,
    `"Entreprises";"${clients.filter(c=>c.genre==='Entreprise').length}";;;;;;`,
  ].filter(Boolean).join('\n');
  const csvContent = "\uFEFF" + [header, ...rows].map(r => r.join(";")).join("\n") + '\n' + recapLines;
  const nomFichier = `clients_TED_${filtreLabel.replace(/ /g,'_')}_${now.toISOString().split('T')[0]}.csv`;
  downloadBlob(csvContent, nomFichier, "text/csv;charset=utf-8;");
}

function exportToXLSX(clients) {
  const header = ["Genre","Entreprise","Nom","Prénom","Téléphone","Mail","Date d'ajout","Commentaire"];
  const rows = clients.map(c => [c.genre||"", c.genre==='Entreprise'?(c.entreprise||''):'', c.nom||"",c.prenom||"",c.tel?`\t${c.tel}`:"",c.mail||"",formatDate(c.created_at),c.commentaire||""]);
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
  const isMob = window.innerWidth < 768;
  useEffect(() => { const t = setTimeout(onClose, 2000); return () => clearTimeout(t); }, [onClose]);
  return (
    <div style={{
      position:'fixed', top:16,
      right: isMob ? 'auto' : 16,
      left: isMob ? '50%' : 'auto',
      transform: isMob ? 'translateX(-50%)' : 'none',
      zIndex:99999, pointerEvents:'none',
      background:'#fff',
      border: `1.5px solid ${type==='error' ? '#dc2626' : '#22c55e'}`,
      borderRadius:12, padding:'10px 16px',
      boxShadow:'0 4px 20px rgba(0,0,0,0.12)',
      display:'flex', alignItems:'center', gap:8,
      fontSize:14, fontWeight:600, color:'#111',
      maxWidth:280, whiteSpace:'nowrap',
      animation:'slideDownFade 0.25s cubic-bezier(0.34,1.56,0.64,1)'
    }}>
      <span style={{ color: type==='error' ? '#dc2626' : '#22c55e', fontWeight:800, fontSize:16 }}>{type==='error' ? '✕' : '✓'}</span>
      {msg}
    </div>
  );
}

// ─── Modal ────────────────────────────────────────────────────────────────────
function Modal({ title, onClose, children, footer, maxW=520, zIndex=3000 }) {
  const isMobile = window.innerWidth < 768;
  return (
    <div
      style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex, display:"flex", alignItems: isMobile ? "flex-end" : "center", justifyContent:"center", padding: isMobile ? 0 : "1rem", pointerEvents:'all', cursor:'default' }}
      onPointerDown={e => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div
        style={{ background:"#fff", borderRadius: isMobile ? '20px 20px 0 0' : 12, width:"100%", maxWidth: isMobile ? '100%' : maxW, overflow:"hidden", maxHeight: isMobile ? 'none' : '90vh', height: isMobile ? '90vh' : 'auto', display:"flex", flexDirection:"column" }}
        onPointerDown={e => e.stopPropagation()}
      >
        <div style={{ background:"#111", color:"#fff", padding:"14px 18px", display:"flex", alignItems:"center", justifyContent:"space-between", flexShrink:0 }}>
          <span style={{ fontWeight:700, fontSize:15 }}>{title}</span>
          <button type="button" onPointerDown={onClose} style={{ background:"none", border:"none", color:"#fff", fontSize:20, cursor:"pointer", touchAction:'manipulation' }}>✕</button>
        </div>
        <div style={{ padding:"18px", overflowY:"auto", flex:1, WebkitOverflowScrolling:"touch" }}>{children}</div>
        {footer && <div style={{ padding: isMobile ? "12px 16px" : "0 18px 18px", paddingBottom: isMobile ? 'calc(16px + env(safe-area-inset-bottom))' : 18, display:"flex", gap:8, justifyContent:"flex-end", flexShrink:0, background:'#fff' }}>{footer}</div>}
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
  const [loginEmail, setLoginEmail] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [tentativesEchouees, setTentativesEchouees] = useState(0);
  const [delaiRestant, setDelaiRestant] = useState(0);
  const [enAttente, setEnAttente] = useState(false);
  const isMob = window.innerWidth < 768;

  function lancerDelai(nbTentatives) {
    const delai = nbTentatives * 5;
    setDelaiRestant(delai);
    setEnAttente(true);
    const interval = setInterval(() => {
      setDelaiRestant(prev => {
        if (prev <= 1) { clearInterval(interval); setEnAttente(false); return 0; }
        return prev - 1;
      });
    }, 1000);
  }

  async function handleLogin() {
    if (enAttente) return;
    setLoginLoading(true);
    setLoginError("");
    const { error } = await supabase.auth.signInWithPassword({ email: loginEmail, password: loginPassword });
    if (error) {
      const nouvTentatives = tentativesEchouees + 1;
      setTentativesEchouees(nouvTentatives);
      setLoginError("Email ou mot de passe incorrect.");
      setLoginLoading(false);
      lancerDelai(nouvTentatives);
      return;
    }
    setTentativesEchouees(0);
    onLogin();
  }

  const carteBlanche = (
    <div style={{ width:500, maxWidth:'100%', background:'#fff', borderRadius:24, padding:'42px 44px', boxShadow:'0 24px 55px rgba(0,0,0,0.28)' }}>
      <h2 style={{ fontSize:34, fontWeight:700, color:'#111', margin:'0 0 8px' }}>Connexion</h2>
      <p style={{ fontSize:16, color:'#777', margin:'0 0 32px' }}>Accédez à votre espace TED CRM</p>

      {/* Email */}
      <div style={{ marginBottom:22 }}>
        <label style={{ fontSize:15, fontWeight:600, color:'#111', display:'block', marginBottom:8 }}>Email</label>
        <div style={{ position:'relative' }}>
          <Mail size={20} color="#aaa" strokeWidth={1.8} style={{ position:'absolute', left:16, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }} />
          <input type="email" autoComplete="email" placeholder="votre@email.com" value={loginEmail} onChange={e=>setLoginEmail(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleLogin()}
            style={{ width:'100%', height:60, border:'1px solid #ddd', borderRadius:12, padding:'0 18px 0 48px', fontSize:16, outline:'none', boxSizing:'border-box', background:'#fff', color:'#111', transition:'border-color 0.2s, box-shadow 0.2s' }}
            onFocus={e=>{ e.target.style.borderColor='#efc434'; e.target.style.boxShadow='0 0 0 4px rgba(239,196,52,0.14)'; }}
            onBlur={e=>{ e.target.style.borderColor='#ddd'; e.target.style.boxShadow='none'; }} />
        </div>
      </div>

      {/* Mot de passe */}
      <div style={{ marginBottom:28 }}>
        <label style={{ fontSize:15, fontWeight:600, color:'#111', display:'block', marginBottom:8 }}>Mot de passe</label>
        <div style={{ position:'relative' }}>
          <LockKeyhole size={20} color="#aaa" strokeWidth={1.8} style={{ position:'absolute', left:16, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }} />
          <input type={showPassword ? 'text' : 'password'} autoComplete="current-password" placeholder="Votre mot de passe" value={loginPassword} onChange={e=>setLoginPassword(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleLogin()}
            style={{ width:'100%', height:60, border:'1px solid #ddd', borderRadius:12, padding:'0 48px 0 48px', fontSize:16, outline:'none', boxSizing:'border-box', background:'#fff', color:'#111', transition:'border-color 0.2s, box-shadow 0.2s' }}
            onFocus={e=>{ e.target.style.borderColor='#efc434'; e.target.style.boxShadow='0 0 0 4px rgba(239,196,52,0.14)'; }}
            onBlur={e=>{ e.target.style.borderColor='#ddd'; e.target.style.boxShadow='none'; }} />
          <button onClick={()=>setShowPassword(v=>!v)} aria-label={showPassword ? 'Masquer' : 'Afficher le mot de passe'} style={{ position:'absolute', right:16, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', cursor:'pointer', padding:4, display:'flex', alignItems:'center' }}>
            {showPassword ? <EyeOff size={20} color="#aaa" strokeWidth={1.8} /> : <Eye size={20} color="#aaa" strokeWidth={1.8} />}
          </button>
        </div>
      </div>

      {loginError && (
        <div style={{ background:'#fef2f2', border:'1px solid #fca5a5', borderRadius:10, padding:'10px 14px', marginBottom:16, fontSize:14, color:'#d92d20', display:'flex', alignItems:'center', gap:8 }}>
          <AlertCircle size={16} color="#d92d20" strokeWidth={1.8} style={{ flexShrink:0 }} />
          <span>{loginError}{enAttente && <strong> Nouvelle tentative dans {delaiRestant}s.</strong>}</span>
        </div>
      )}

      <button onClick={handleLogin} disabled={loginLoading || enAttente}
        style={{ width:'100%', height:60, background: enAttente ? '#f0f0f0' : '#efc434', border:'none', borderRadius:12, fontSize:17, fontWeight:700, cursor: (loginLoading || enAttente) ? 'not-allowed' : 'pointer', color: enAttente ? '#999' : '#111', display:'flex', alignItems:'center', justifyContent:'center', gap:10, boxShadow: enAttente ? 'none' : '0 4px 14px rgba(239,196,52,0.28)', transition:'all 0.2s' }}
        onMouseEnter={e=>{ if(!loginLoading && !enAttente) e.currentTarget.style.background='#ddb226'; }}
        onMouseLeave={e=>{ if(!loginLoading && !enAttente) e.currentTarget.style.background='#efc434'; }}>
        {enAttente ? (
          <><span>Veuillez patienter</span><span style={{background:'#ddd', color:'#666', borderRadius:20, padding:'2px 10px', fontSize:15, fontWeight:800}}>{delaiRestant}s</span></>
        ) : loginLoading ? 'Connexion...' : (
          <><span>Se connecter</span><ArrowRight size={20} strokeWidth={2}/></>
        )}
      </button>

      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:20, paddingTop:20, marginTop:22, borderTop:'1px solid #eee' }}>
        <div style={{ display:'flex', alignItems:'flex-start', gap:10 }}>
          <LockKeyhole size={20} color="#999" strokeWidth={1.8} style={{ marginTop:1, flexShrink:0 }} />
          <div><div style={{ fontSize:13, fontWeight:700, color:'#111' }}>Vos données sont protégées.</div><div style={{ fontSize:12, color:'#999' }}>Confidentialité garantie.</div></div>
        </div>
        <div style={{ display:'flex', alignItems:'flex-start', gap:10 }}>
          <Headphones size={20} color="#999" strokeWidth={1.8} style={{ marginTop:1, flexShrink:0 }} />
          <div><div style={{ fontSize:13, fontWeight:700, color:'#111' }}>Besoin d'aide ?</div><div style={{ fontSize:12, color:'#999' }}>Contactez votre responsable.</div></div>
        </div>
      </div>
    </div>
  );

  if (isMob) return (
    <div style={{ minHeight:'100vh', background:'#0a0a0a', display:'flex', alignItems:'center', justifyContent:'center', padding:24, boxSizing:'border-box' }}>
      {carteBlanche}
    </div>
  );

  return (
    <div style={{ minHeight:'100vh', background:'#0a0a0a', display:'flex', alignItems:'center', justifyContent:'center', padding:24, boxSizing:'border-box' }}>
      <div style={{ width:'min(1220px, calc(100vw - 48px))', height:'min(760px, calc(100vh - 48px))', background:'linear-gradient(135deg, #0d0d0d 0%, #151515 100%)', borderRadius:28, overflow:'hidden', boxShadow:'0 24px 60px rgba(0,0,0,0.5)', display:'grid', gridTemplateColumns:'46% 54%' }}>
        {/* Colonne gauche */}
        <div style={{ padding:'44px 52px', display:'flex', flexDirection:'column', justifyContent:'space-between', position:'relative', overflow:'hidden' }}>
          <img src="/favicon.png" style={{ position:'absolute', left:-20, top:'50%', transform:'translateY(-50%)', width:360, height:360, opacity:0.035, filter:'sepia(1) saturate(4) hue-rotate(355deg)', pointerEvents:'none' }} alt="" />
          <div style={{ display:'flex', alignItems:'center', gap:12, position:'relative' }}>
            <img src="/favicon.png" style={{ width:52, height:52 }} alt="TED" />
            <span style={{ fontSize:36, fontWeight:900, color:'#fff', letterSpacing:0.5 }}>TED <span style={{ color:'#efc434' }}>CRM</span></span>
          </div>
          <div style={{ position:'relative' }}>
            <h1 style={{ fontSize:48, fontWeight:700, color:'#fff', margin:'0 0 14px', lineHeight:1.05 }}>Connexion</h1>
            <p style={{ fontSize:18, color:'rgba(255,255,255,0.62)', margin:0 }}>Accédez à votre espace TED CRM</p>
          </div>
          <div style={{ display:'flex', position:'relative' }}>
            {[
              { icon:<RefreshCw size={26} color="#efc434" strokeWidth={1.8} />, title:'Synchronisé', sub:'en temps réel' },
              { icon:<ShieldCheck size={26} color="#efc434" strokeWidth={1.8} />, title:'Sécurisé', sub:'et fiable' },
              { icon:<MonitorSmartphone size={26} color="#efc434" strokeWidth={1.8} />, title:'iPad / PC / Mobile', sub:'Partout avec vous' },
            ].map((f, i) => (
              <div key={i} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:6, padding:'0 12px', borderLeft: i > 0 ? '1px solid rgba(255,255,255,0.08)' : 'none' }}>
                {f.icon}
                <span style={{ fontSize:14, fontWeight:700, color:'#fff', textAlign:'center' }}>{f.title}</span>
                <span style={{ fontSize:12, color:'rgba(255,255,255,0.4)', textAlign:'center' }}>{f.sub}</span>
              </div>
            ))}
          </div>
        </div>
        {/* Colonne droite */}
        <div style={{ display:'flex', alignItems:'center', justifyContent:'center', padding:40, background:'rgba(0,0,0,0.2)' }}>
          {carteBlanche}
        </div>
      </div>
    </div>
  );
}

// ─── Client Form ──────────────────────────────────────────────────────────────
function ClientForm({ initial, onSave, onCancel, existingClients, reservations = [] }) {
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

  function checkDupTel() {
    const others = existingClients.filter(c => !isEdit || c.id !== initial?.id);
    return others.find(c => c.tel && form.tel && c.tel === form.tel) || null;
  }

  function checkDupMail() {
    const others = existingClients.filter(c => !isEdit || c.id !== initial?.id);
    if (!form.mail.trim()) return null;
    return others.find(c => c.mail && c.mail.toLowerCase() === form.mail.trim().toLowerCase()) || null;
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
    // Tel : doublon bloquant
    const dupTel = checkDupTel();
    if (dupTel) {
      const nom = dupTel.genre === 'Entreprise' ? dupTel.entreprise : `${dupTel.prenom} ${dupTel.nom}`;
      setErrors(ex => ({...ex, tel: `Ce numéro est déjà utilisé par ${nom}`}));
      return;
    }
    // Mail : doublon avertissement (contournable)
    const dupMail = checkDupMail();
    if (dupMail) { setDupWarn(`L'adresse ${form.mail} est déjà utilisée par ${dupMail.prenom} ${dupMail.nom}.`); return; }
    setSuccess(true);
    setTimeout(() => { doSave(); }, 800);
  }

  const inputStyle = (err) => ({
    width: "100%", height: 44, border: `1.5px solid ${err ? "#dc2626" : "#ddd"}`,
    borderRadius: 7, padding: "0 12px", fontSize: 16, outline: "none", boxSizing: "border-box"
  });
  const labelStyle = { display: "block", fontSize: 12, fontWeight: 600, color: "#444", marginBottom: 5 };
  const fieldGroup = { marginBottom: 14 };

  const clientValide = !!(form.tel && form.nom && form.prenom && form.genre && form.genre !== 'Non renseigné');

  // ─── Design 2 colonnes desktop en mode édition ───────────────────────────
  if (isEdit && !isMobile) {
    const aujourd = new Date().toISOString().split('T')[0];
    const [localConfirmQuitter, setLocalConfirmQuitter] = useState(false);
    const isDirty = Object.keys(form).some(k => (form[k]||'') !== (initial[k]||''));
    const handleOverlayClick = (e) => {
      if (e.target !== e.currentTarget) return;
      if (isDirty) { setLocalConfirmQuitter(true); } else { onCancel(); }
    };
    return (
      <>
        {dupWarn && <ConfirmModal title="Doublon détecté" msg={`Attention : ${dupWarn} Voulez-vous tout de même continuer ?`} onOk={()=>{ setDupWarn(null); doSave(); }} onCancel={()=>setDupWarn(null)} okLabel="Ajouter quand même" />}
        {localConfirmQuitter && (
          <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.7)',zIndex:9000,display:'flex',alignItems:'center',justifyContent:'center', pointerEvents:'all'}}>
            <div style={{background:'#fff',borderRadius:16,padding:'28px 24px',maxWidth:320,width:'90%',textAlign:'center'}} onClick={e=>e.stopPropagation()}>
              <h3 style={{margin:'0 0 8px',fontSize:17,fontWeight:800,color:'#111'}}>Quitter sans enregistrer ?</h3>
              <p style={{margin:'0 0 20px',fontSize:14,color:'#666'}}>Les informations saisies seront perdues.</p>
              <div style={{display:'flex',gap:10}}>
                <button onClick={()=>setLocalConfirmQuitter(false)} style={{flex:1,height:44,border:'1.5px solid #ddd',borderRadius:10,background:'#fff',fontSize:14,fontWeight:600,cursor:'pointer',color:'#666'}}>Continuer la saisie</button>
                <button onClick={()=>{setLocalConfirmQuitter(false); onCancel();}} style={{flex:1,height:44,border:'none',borderRadius:10,background:'#dc2626',fontSize:14,fontWeight:800,cursor:'pointer',color:'#fff'}}>Quitter</button>
              </div>
            </div>
          </div>
        )}
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.3)', zIndex:3000, display:'flex', pointerEvents:'all', cursor:'default', touchAction:'none' }} onMouseDown={e=>{e.preventDefault();e.stopPropagation();handleOverlayClick(e);}} onClick={handleOverlayClick}>

          {/* Colonne gauche — liste clients grisée */}
          <div style={{ flex:1, background:'#f5f5f5', padding:'32px', overflowY:'auto', opacity:0.6, pointerEvents:'none' }}>
            <h2 style={{ fontSize:28, fontWeight:900, color:'#111', margin:'0 0 20px' }}>Clients</h2>
            <div style={{ background:'#fff', borderRadius:10, padding:'10px 14px', marginBottom:20, display:'flex', alignItems:'center', gap:8 }}>
              <Search size={16} color="#999" strokeWidth={2}/>
              <span style={{ fontSize:14, color:'#bbb' }}>Rechercher un client...</span>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', padding:'8px 0', borderBottom:'1px solid #eee', marginBottom:8 }}>
              {['NOM','TÉLÉPHONE','DERNIÈRE RÉSERVATION'].map(h => (
                <span key={h} style={{ fontSize:11, fontWeight:700, color:'#999', letterSpacing:0.5 }}>{h}</span>
              ))}
            </div>
            {existingClients.slice(0,10).map(c => {
              const derniereResa = reservations.filter(r=>r.client_id===c.id&&r.date<=aujourd).sort((a,b)=>b.date.localeCompare(a.date))[0];
              return (
                <div key={c.id} style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', padding:'12px 0', borderBottom:'1px solid #f0f0f0' }}>
                  <span style={{ fontWeight:700, fontSize:14, color:'#111' }}>{c.genre==='Entreprise'?c.entreprise:`${c.prenom} ${c.nom}`}</span>
                  <span style={{ fontSize:14, color:'#444' }}>{c.tel}</span>
                  <span style={{ fontSize:14, color:'#444' }}>{derniereResa ? new Date(derniereResa.date+'T12:00:00').toLocaleDateString('fr-FR') : '—'}</span>
                </div>
              );
            })}
            <div style={{ marginTop:16, fontSize:13, color:'#999' }}>{existingClients.length} clients</div>
          </div>

          {/* Colonne droite — formulaire */}
          <div style={{ width:480, background:'#fff', padding:'40px', display:'flex', flexDirection:'column', boxShadow:'-8px 0 40px rgba(0,0,0,0.1)' }} onClick={e=>e.stopPropagation()}>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:32 }}>
              <h2 style={{ margin:0, fontSize:24, fontWeight:900, color:'#111' }}>Modifier le client</h2>
              <button onClick={onCancel} style={{ width:36, height:36, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:18, color:'#666', display:'flex', alignItems:'center', justifyContent:'center' }}>✕</button>
            </div>

            <div style={{ display:'flex', flexDirection:'column', gap:20, flex:1, overflowY:'auto' }}>

              {/* Genre */}
              <div>
                <label style={{ fontSize:14, fontWeight:700, color:'#111', display:'block', marginBottom:8 }}>Genre</label>
                <select value={form.genre} onChange={e=>set('genre', e.target.value)} style={{ width:'100%', height:52, border:`1.5px solid ${errors.genre?'#dc2626':'#eee'}`, borderRadius:12, padding:'0 16px', fontSize:15, outline:'none', background:'#fff', cursor:'pointer' }}>
                  <option value="Non renseigné">-- Sélectionner --</option>
                  {GENRES.filter(g=>g!=='Non renseigné').map(g=><option key={g}>{g}</option>)}
                </select>
                {errors.genre && <p style={{ fontSize:12, color:'#dc2626', marginTop:4 }}>{errors.genre}</p>}
              </div>

              {/* Entreprise si Entreprise */}
              {form.genre==='Entreprise' && (
                <div>
                  <label style={{ fontSize:14, fontWeight:700, color:'#111', display:'block', marginBottom:8 }}>Nom de l'entreprise</label>
                  <input value={form.entreprise} onChange={e=>set('entreprise', e.target.value)}
                    style={{ width:'100%', height:52, border:`1.5px solid ${errors.entreprise?'#dc2626':'#eee'}`, borderRadius:12, padding:'0 16px', fontSize:15, outline:'none', boxSizing:'border-box' }}
                    onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor=errors.entreprise?'#dc2626':'#eee'}/>
                  {errors.entreprise && <p style={{ fontSize:12, color:'#dc2626', marginTop:4 }}>{errors.entreprise}</p>}
                </div>
              )}

              {/* Prénom */}
              <div>
                <label style={{ fontSize:14, fontWeight:700, color:'#111', display:'block', marginBottom:8 }}>Prénom</label>
                <input value={form.prenom} onChange={e=>set('prenom', e.target.value)}
                  style={{ width:'100%', height:52, border:`1.5px solid ${errors.prenom?'#dc2626':'#eee'}`, borderRadius:12, padding:'0 16px', fontSize:15, outline:'none', boxSizing:'border-box' }}
                  onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor=errors.prenom?'#dc2626':'#eee'}/>
                {errors.prenom && <p style={{ fontSize:12, color:'#dc2626', marginTop:4 }}>{errors.prenom}</p>}
              </div>

              {/* Nom */}
              <div>
                <label style={{ fontSize:14, fontWeight:700, color:'#111', display:'block', marginBottom:8 }}>Nom</label>
                <input value={form.nom} onChange={e=>set('nom', e.target.value)}
                  style={{ width:'100%', height:52, border:`1.5px solid ${errors.nom?'#dc2626':'#eee'}`, borderRadius:12, padding:'0 16px', fontSize:15, outline:'none', boxSizing:'border-box' }}
                  onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor=errors.nom?'#dc2626':'#eee'}/>
                {errors.nom && <p style={{ fontSize:12, color:'#dc2626', marginTop:4 }}>{errors.nom}</p>}
              </div>

              {/* Téléphone */}
              <div>
                <label style={{ fontSize:14, fontWeight:700, color:'#111', display:'block', marginBottom:8 }}>Téléphone</label>
                <div style={{ position:'relative' }}>
                  <Phone size={16} strokeWidth={2} color="#999" style={{ position:'absolute', left:16, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }}/>
                  <input value={form.tel} onChange={e=>handleTel(e.target.value)} type="tel" inputMode="numeric" maxLength={10}
                    style={{ width:'100%', height:52, border:`1.5px solid ${errors.tel?'#dc2626':'#eee'}`, borderRadius:12, padding:'0 16px 0 44px', fontSize:15, outline:'none', boxSizing:'border-box' }}
                    onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor=errors.tel?'#dc2626':'#eee'}/>
                </div>
                {errors.tel && <p style={{ fontSize:12, color:'#dc2626', marginTop:4 }}>{errors.tel}</p>}
                {dupClient && <div style={{ background:'#fef2f2', border:'2px solid #dc2626', borderRadius:10, padding:'10px 14px', marginTop:8, fontSize:13, color:'#dc2626' }}>⚠️ Ce numéro est déjà utilisé par <strong>{dupClient.prenom} {dupClient.nom}</strong></div>}
              </div>

              {/* Email */}
              <div>
                <label style={{ fontSize:14, fontWeight:700, color:'#111', display:'block', marginBottom:8 }}>Email</label>
                <div style={{ position:'relative' }}>
                  <Mail size={16} strokeWidth={2} color="#999" style={{ position:'absolute', left:16, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }}/>
                  <input value={form.mail} onChange={e=>set('mail', e.target.value)} type="email"
                    style={{ width:'100%', height:52, border:`1.5px solid ${errors.mail?'#dc2626':'#eee'}`, borderRadius:12, padding:'0 16px 0 44px', fontSize:15, outline:'none', boxSizing:'border-box' }}
                    onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor=errors.mail?'#dc2626':'#eee'}/>
                </div>
                {errors.mail && <p style={{ fontSize:12, color:'#dc2626', marginTop:4 }}>{errors.mail}</p>}
              </div>

              {/* Commentaire */}
              <div>
                <label style={{ fontSize:14, fontWeight:700, color:'#111', display:'block', marginBottom:8 }}>Commentaire</label>
                <textarea value={form.commentaire} onChange={e=>set('commentaire', e.target.value)} placeholder="Notes sur ce client…"
                  style={{ width:'100%', border:'1.5px solid #eee', borderRadius:12, padding:'12px 16px', fontSize:15, outline:'none', boxSizing:'border-box', resize:'vertical', minHeight:80, fontFamily:'inherit' }}
                  onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
              </div>
            </div>

            <div style={{ display:'flex', gap:12, marginTop:32 }}>
              <button onClick={onCancel} style={{ flex:1, height:52, border:'1.5px solid #eee', borderRadius:12, background:'#fff', fontSize:15, fontWeight:600, cursor:'pointer', color:'#666' }}>Annuler</button>
              <button onClick={dupClient ? undefined : handleSubmit} disabled={!!dupClient || !clientValide} style={{ flex:2, height:52, border:'none', borderRadius:12, background: dupClient ? '#ddd' : (success ? '#22c55e' : (clientValide ? '#E8C547' : '#f0f0f0')), color: dupClient ? '#999' : (success ? '#fff' : (clientValide ? '#111' : '#bbb')), fontSize:15, fontWeight:800, cursor: dupClient || !clientValide ? 'not-allowed' : 'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                {success ? '✓ Enregistré !' : <><Save size={18} strokeWidth={2}/> Enregistrer</>}
              </button>
            </div>
          </div>
        </div>
      </>
    );
  }

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
          <button key="s" type="button" onPointerDown={dupClient || !clientValide ? undefined : handleSubmit} disabled={!!dupClient || !clientValide} style={{
            background: dupClient ? "#ddd" : (success ? "#22c55e" : (clientValide ? "#E8C547" : "#f0f0f0")),
            color: dupClient ? "#999" : (success ? "#fff" : (clientValide ? "#111" : "#bbb")),
            border: "none", borderRadius: 12,
            height: 52, fontWeight: 700, fontSize: 16,
            cursor: dupClient || !clientValide ? "not-allowed" : "pointer", flex: 2, touchAction: "manipulation",
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
  const [confirmVider, setConfirmVider] = useState(false);
  const [confirmSuppr, setConfirmSuppr] = useState(null);

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
    await supabase.from("clients").delete().eq("id", id);
    setDeleted(prev => prev.filter(c => c.id !== id));
    setConfirmSuppr(null);
    showToast("Client supprimé définitivement");
  }

  async function emptyTrash() {
    await supabase.from("clients").delete().not("deleted_at", "is", null);
    setDeleted([]);
    setConfirmVider(false);
    showToast("Corbeille vidée ✓");
  }

  function nomClient(c) {
    return c.genre === "Entreprise" ? (c.entreprise || c.nom) : `${c.nom || ""} ${c.prenom || ""}`.trim();
  }

  return (
    <>
      <div
        onClick={onClose}
        onMouseDown={e => { e.preventDefault(); e.stopPropagation(); }}
        style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:2999, cursor:"pointer" }}
      />

      <div style={{
        position:"fixed", top:"50%", left:"50%", transform:"translate(-50%,-50%)",
        background:"#fff", borderRadius:20, width:"min(600px,calc(100vw - 48px))",
        maxHeight:"85vh", display:"flex", flexDirection:"column",
        boxShadow:"0 24px 80px rgba(0,0,0,0.18)", zIndex:3000, overflow:"hidden"
      }}>
        {/* Header */}
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"24px 28px 20px", borderBottom:"1.5px solid #f0f0f0", flexShrink:0 }}>
          <div style={{ display:"flex", alignItems:"center", gap:12 }}>
            <div style={{ width:40, height:40, borderRadius:12, background:"#fef2f2", display:"flex", alignItems:"center", justifyContent:"center" }}>
              <span style={{ fontSize:20 }}>🗑</span>
            </div>
            <div>
              <div style={{ fontWeight:800, fontSize:18, color:"#111" }}>Corbeille</div>
              {!loading && <div style={{ fontSize:12, color:"#999", marginTop:1 }}>{deleted.length} client{deleted.length !== 1 ? "s" : ""}</div>}
            </div>
          </div>
          <button onClick={onClose} style={{ width:36, height:36, borderRadius:10, border:"1.5px solid #eee", background:"#f5f5f5", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", fontSize:18, color:"#555", fontWeight:700 }}>✕</button>
        </div>

        {/* Contenu scrollable */}
        <div style={{ flex:1, overflowY:"auto", padding:"20px 28px" }}>
          {loading && (
            <div style={{ textAlign:"center", padding:"40px 0", color:"#bbb", fontSize:15 }}>Chargement…</div>
          )}
          {!loading && deleted.length === 0 && (
            <div style={{ textAlign:"center", padding:"48px 0" }}>
              <div style={{ fontSize:52, marginBottom:12 }}>🗑</div>
              <div style={{ fontWeight:700, fontSize:16, color:"#333", marginBottom:6 }}>La corbeille est vide</div>
              <div style={{ fontSize:13, color:"#bbb" }}>Les clients supprimés apparaîtront ici</div>
            </div>
          )}
          {!loading && deleted.map(c => (
            <div key={c.id} style={{ background:"#fafafa", border:"1.5px solid #f0f0f0", borderRadius:14, padding:"16px 18px", marginBottom:10, display:"flex", alignItems:"center", justifyContent:"space-between", gap:12 }}>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:4 }}>
                  <span style={badge(c.genre)}>{c.genre}</span>
                  <span style={{ fontWeight:700, fontSize:15, color:"#111" }}>{nomClient(c)}</span>
                </div>
                {c.tel && <div style={{ fontSize:13, color:"#666", marginBottom:2 }}>📞 {c.tel}</div>}
                <div style={{ fontSize:11, color:"#bbb" }}>
                  Supprimé le {new Date(c.deleted_at).toLocaleDateString("fr-FR")} à {new Date(c.deleted_at).toLocaleTimeString("fr-FR", { hour:"2-digit", minute:"2-digit" })}
                  {c.deleted_by ? ` par ${c.deleted_by}` : ""}
                </div>
              </div>
              <div style={{ display:"flex", gap:8, flexShrink:0 }}>
                <button onClick={() => restore(c.id)} style={{ height:38, padding:"0 14px", borderRadius:10, border:"1.5px solid #22c55e", background:"#f0fdf4", color:"#16a34a", fontWeight:700, fontSize:13, cursor:"pointer" }}>↩ Restaurer</button>
                <button onClick={() => setConfirmSuppr(c)} style={{ height:38, padding:"0 14px", borderRadius:10, border:"1.5px solid #f0f0f0", background:"#fff", color:"#dc2626", fontWeight:700, fontSize:13, cursor:"pointer" }}>✕</button>
              </div>
            </div>
          ))}
        </div>

        {/* Footer */}
        <div style={{ padding:"16px 28px 24px", borderTop:"1.5px solid #f0f0f0", flexShrink:0, display:"flex", gap:12 }}>
          <button onClick={onClose} style={{ flex:1, height:50, borderRadius:14, border:"1.5px solid #eee", background:"#f5f5f5", fontWeight:700, fontSize:15, cursor:"pointer", color:"#333" }}>Fermer</button>
          {deleted.length > 0 && (
            <button onClick={() => setConfirmVider(true)} style={{ flex:1, height:50, borderRadius:14, border:"none", background:"#dc2626", color:"#fff", fontWeight:700, fontSize:15, cursor:"pointer" }}>
              Vider la corbeille ({deleted.length})
            </button>
          )}
        </div>
      </div>

      {/* Confirm suppression définitive */}
      {confirmSuppr && (
        <>
          <div onClick={() => setConfirmSuppr(null)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.3)", zIndex:3100 }} />
          <div style={{ position:"fixed", top:"50%", left:"50%", transform:"translate(-50%,-50%)", background:"#fff", borderRadius:20, width:"min(420px,calc(100vw - 48px))", padding:"28px", zIndex:3101, boxShadow:"0 24px 80px rgba(0,0,0,0.18)" }}>
            <div style={{ fontWeight:800, fontSize:17, color:"#111", marginBottom:8 }}>Supprimer définitivement ?</div>
            <div style={{ fontSize:14, color:"#666", marginBottom:24 }}>
              <strong>{nomClient(confirmSuppr)}</strong> sera supprimé définitivement. Cette action est irréversible.
            </div>
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={() => setConfirmSuppr(null)} style={{ flex:1, height:48, borderRadius:12, border:"1.5px solid #eee", background:"#f5f5f5", fontWeight:700, fontSize:14, cursor:"pointer" }}>Annuler</button>
              <button onClick={() => deletePermanently(confirmSuppr.id)} style={{ flex:1, height:48, borderRadius:12, border:"none", background:"#dc2626", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer" }}>Supprimer</button>
            </div>
          </div>
        </>
      )}

      {/* Confirm vider corbeille */}
      {confirmVider && (
        <>
          <div onClick={() => setConfirmVider(false)} style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.3)", zIndex:3100 }} />
          <div style={{ position:"fixed", top:"50%", left:"50%", transform:"translate(-50%,-50%)", background:"#fff", borderRadius:20, width:"min(420px,calc(100vw - 48px))", padding:"28px", zIndex:3101, boxShadow:"0 24px 80px rgba(0,0,0,0.18)" }}>
            <div style={{ fontWeight:800, fontSize:17, color:"#111", marginBottom:8 }}>Vider la corbeille ?</div>
            <div style={{ fontSize:14, color:"#666", marginBottom:24 }}>
              {deleted.length} client{deleted.length !== 1 ? "s" : ""} seront supprimés définitivement. Cette action est irréversible.
            </div>
            <div style={{ display:"flex", gap:10 }}>
              <button onClick={() => setConfirmVider(false)} style={{ flex:1, height:48, borderRadius:12, border:"1.5px solid #eee", background:"#f5f5f5", fontWeight:700, fontSize:14, cursor:"pointer" }}>Annuler</button>
              <button onClick={emptyTrash} style={{ flex:1, height:48, borderRadius:12, border:"none", background:"#dc2626", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer" }}>Vider ({deleted.length})</button>
            </div>
          </div>
        </>
      )}
    </>
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
const FORM_URL = "https://ted-crm.pages.dev/reserver";

const OCCASIONS = ["Anniversaire","EVG — Enterrement de vie de garçon","EVJF — Enterrement de vie de jeune fille","Privatisation","Autre"];
const HEURES_MIDI = ["12:00","12:15","12:30","12:45","13:00","13:15","13:30"];
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

function AddResaModal({ onClose, onSaved, showToast, user, initialResa, onViewClient, reservations=[] }) {
  const DATE_OPTS = useMemo(() => buildDateOptions(), []);
  const isEdit = !!initialResa?.id;
  const initClient = initialResa?.clients || {};

  const [tel, setTel] = useState(initClient.tel || '');
  const [clientFound, setClientFound] = useState(isEdit ? initClient : null);
  const [statsClient, setStatsClient] = useState(null);
  const [lookingUp, setLookingUp] = useState(false);
  const [genre, setGenre] = useState(initClient.genre || '');
  const [prenom, setPrenom] = useState(initClient.prenom || '');
  const [nom, setNom] = useState(initClient.nom || '');
  const [entreprise, setEntreprise] = useState(initClient.entreprise || '');
  const [email, setEmail] = useState(initClient.mail || '');
  const [dateIso, setDateIso] = useState(initialResa?.date || DATE_OPTS[0].iso);
  const [service, setService] = useState(initialResa?.service || 'soir');
  const [heure, setHeure] = useState(initialResa?.heure || '');
  const [nbPersonnes, setNbPersonnes] = useState(initialResa?.nb_personnes || 2);
  const [occasion, setOccasion] = useState(initialResa?.occasion || '');
  const [commentaire, setCommentaire] = useState(initialResa?.commentaire_client || '');
  const [saving, setSaving] = useState(false);
  const [heureError, setHeureError] = useState(false);
  const [showCalPicker, setShowCalPicker] = useState(false);
  const [dateFlash, setDateFlash] = useState(null);
  const [calFermeture, setCalFermeture] = useState(false);
  const [showEditClientInline, setShowEditClientInline] = useState(false);
  const [editClientForm, setEditClientForm] = useState({});
  const [resaCree, setResaCree] = useState(null);
  const [showConfirmQuitter, setShowConfirmQuitter] = useState(false);
  const [calPickerDate, setCalPickerDate] = useState(() => {
    if (initialResa?.date) {
      const d = new Date(initialResa.date + 'T12:00:00');
      return isNaN(d.getTime()) ? new Date() : d;
    }
    return new Date();
  });
  const isMobile = useIsMobile();
  const calPickerRef = useRef(null);
  const submitLockRef = useRef(false);
  const refTel = useRef(null);
  const refDate = useRef(null);
  const refService = useRef(null);
  const refHeure = useRef(null);
  const refGenre = useRef(null);
  const refEmail = useRef(null);

  useEffect(() => {
    if (showCalPicker) {
      if (dateIso) {
        const d = new Date(dateIso + 'T12:00:00');
        if (!isNaN(d.getTime())) setCalPickerDate(d);
      }
      if (calPickerRef.current) {
        setTimeout(() => {
          calPickerRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 100);
      }
    }
  }, [showCalPicker]);

  const heures = service === 'midi' ? HEURES_MIDI : HEURES_SOIR;

  // Recherche automatique dès 10 chiffres
  async function handleTelChange(val) {
    setTel(val);
    setClientFound(null);
    setStatsClient(null);
    setGenre(''); setPrenom(''); setNom(''); setEmail(''); setEntreprise('');
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
      const { data: resas } = await supabase.from('reservations').select('statut,date').eq('client_id', data.id);
      if (resas) {
        const total = resas.filter(r => r.statut !== 'annulee' && r.statut !== 'absente' && r.statut !== 'refusee').length;
        const noshow = resas.filter(r => r.statut === 'absente').length;
        const derniereVisite = resas.filter(r => r.statut === 'venue' || r.statut === 'confirmee').sort((a,b) => b.date.localeCompare(a.date))[0];
        setStatsClient({
          total,
          noshow,
          derniereVisite: derniereVisite
            ? new Date(derniereVisite.date+'T12:00:00').toLocaleDateString('fr-FR', {day:'numeric', month:'short', year:'numeric'})
            : 'Jamais'
        });
      }
    }
  }

  async function handleSave() {
    if (submitLockRef.current) return;
    if (!tel.trim()) { showToast('Téléphone requis', 'error'); return; }
    if (!clientFound) {
      if (!genre) { showToast('Genre requis', 'error'); return; }
      if (genre !== 'Entreprise' && !prenom.trim()) { showToast('Prénom requis', 'error'); return; }
      if (genre !== 'Entreprise' && !nom.trim()) { showToast('Nom requis', 'error'); return; }
      if (genre === 'Entreprise' && !entreprise.trim()) { showToast("Nom d'entreprise requis", 'error'); return; }
    }
    if (!heure) { setHeureError(true); return; }
    submitLockRef.current = true;
    setSaving(true);

    const telSaisi = tel.replace(/\s/g, '');
    const nomSaisi = nom.toLowerCase().trim();
    const prenomSaisi = prenom.toLowerCase().trim();
    const mailSaisi = email.toLowerCase().trim();
    const telNorm = tel.replace(/[\s.\-()]/g,'').replace(/^0/,'+33');

    const { data: clientParTel } = await supabase
      .from('clients').select('*')
      .or(`tel.eq.${telSaisi},tel_normalise.eq.${telNorm}`)
      .maybeSingle();

    const { data: clientParMail } = (!clientParTel && mailSaisi)
      ? await supabase.from('clients').select('*').eq('mail', mailSaisi).maybeSingle()
      : { data: null };

    let clientId = null;

    if (clientParTel) {
      const nomMatch = clientParTel.nom?.toLowerCase().trim() === nomSaisi;
      const prenomMatch = clientParTel.prenom?.toLowerCase().trim() === prenomSaisi;
      if (nomMatch && prenomMatch) {
        // Tel + Nom + Prénom correspondent → met à jour si champs modifiés
        const updates = {};
        if (mailSaisi && mailSaisi !== clientParTel.mail) updates.mail = mailSaisi;
        if (genre && genre !== clientParTel.genre) updates.genre = genre;
        if (capitalize(nom.trim()) !== clientParTel.nom) updates.nom = capitalize(nom.trim());
        if (capitalize(prenom.trim()) !== clientParTel.prenom) updates.prenom = capitalize(prenom.trim());
        if (Object.keys(updates).length > 0) {
          await supabase.from('clients').update(updates).eq('id', clientParTel.id);
        }
        clientId = clientParTel.id;
      } else {
        // Tel identique mais nom/prénom différents → nouveau client
        const { data: newClient, error: errClient } = await supabase.from('clients').insert({
          prenom: capitalize(prenom.trim()), nom: capitalize(nom.trim()),
          tel: tel.trim(), tel_normalise: telNorm,
          mail: mailSaisi || null, genre,
          entreprise: genre === 'Entreprise' ? entreprise.trim() : null,
          source: 'manuel',
        }).select('id').single();
        if (errClient) { setSaving(false); showToast('Erreur création client', 'error'); return; }
        clientId = newClient.id;
      }
    } else if (clientParMail) {
      const nomMatch = clientParMail.nom?.toLowerCase().trim() === nomSaisi;
      const prenomMatch = clientParMail.prenom?.toLowerCase().trim() === prenomSaisi;
      if (nomMatch && prenomMatch) {
        // Mail + Nom + Prénom correspondent → met à jour le téléphone si différent
        const updates = {};
        if (telSaisi && telSaisi !== clientParMail.tel) { updates.tel = telSaisi; updates.tel_normalise = telNorm; }
        if (genre && genre !== clientParMail.genre) updates.genre = genre;
        if (Object.keys(updates).length > 0) {
          await supabase.from('clients').update(updates).eq('id', clientParMail.id);
        }
        clientId = clientParMail.id;
      } else {
        // Mail identique mais nom/prénom différents → nouveau client
        const { data: newClient, error: errClient } = await supabase.from('clients').insert({
          prenom: capitalize(prenom.trim()), nom: capitalize(nom.trim()),
          tel: tel.trim(), tel_normalise: telNorm,
          mail: mailSaisi || null, genre,
          entreprise: genre === 'Entreprise' ? entreprise.trim() : null,
          source: 'manuel',
        }).select('id').single();
        if (errClient) { setSaving(false); showToast('Erreur création client', 'error'); return; }
        clientId = newClient.id;
      }
    } else {
      // Aucun client trouvé → nouveau client
      const { data: newClient, error: errClient } = await supabase.from('clients').insert({
        prenom: capitalize(prenom.trim()), nom: capitalize(nom.trim()),
        tel: tel.trim(), tel_normalise: telNorm,
        mail: mailSaisi || null, genre,
        entreprise: genre === 'Entreprise' ? entreprise.trim() : null,
        source: 'manuel',
      }).select('id').single();
      if (errClient) { setSaving(false); showToast('Erreur création client', 'error'); return; }
      clientId = newClient.id;
    }

    let error;
    if (isEdit) {
      ({ error } = await supabase.from('reservations').update({
        date: dateIso,
        service,
        heure: heure || null,
        nb_personnes: nbPersonnes,
        occasion: occasion || null,
        commentaire_client: commentaire.trim() || null,
        updated_at: new Date().toISOString(),
      }).eq('id', initialResa.id));
    } else {
      ({ error } = await supabase.from('reservations').insert({
        client_id: clientId,
        date: dateIso,
        service,
        heure: heure || null,
        nb_personnes: nbPersonnes,
        occasion: occasion || null,
        commentaire_client: commentaire.trim() || null,
        statut: 'attente',
        source: 'manuel',
      }));
    }
    setSaving(false);
    submitLockRef.current = false;
    if (error) { showToast(isEdit ? 'Erreur lors de la modification' : 'Erreur lors de la création', 'error'); return; }
    onSaved();
    if (isEdit) {
      showToast('Réservation modifiée ✓');
      onClose();
    } else {
      setResaCree({
        client: clientFound || { prenom, nom },
        date: dateIso, service, heure, nb_personnes: nbPersonnes, occasion
      });
    }
  }

  const btnSvc = (s) => ({
    flex: 1, height: 42, border: `1.5px solid ${service === s ? '#111' : '#eee'}`,
    borderRadius: 8, background: service === s ? '#111' : '#f8f8f8',
    color: service === s ? '#fff' : '#666', fontWeight: 700, fontSize: 14, cursor: 'pointer'
  });

  const GENRE_STYLES = {
    'Homme':      { bg:'#dbeafe', border:'#3b82f6', color:'#1d4ed8' },
    'Femme':      { bg:'#fce7f3', border:'#ec4899', color:'#be185d' },
    'Entreprise': { bg:'#dcfce7', border:'#22c55e', color:'#15803d' },
  };
  const btnGenre = (g) => {
    const sel = genre === g;
    const s = GENRE_STYLES[g] || {};
    return {
      flex:1, height:44, borderRadius:10, cursor:'pointer', fontSize:14, fontWeight:700,
      border: sel ? `2px solid ${s.border}` : '1.5px solid #ddd',
      background: sel ? s.bg : '#fff',
      color: sel ? s.color : '#666',
      transition:'all 0.15s'
    };
  };

  const emailValide = (e) => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(e);
  const showNouveauClient = !clientFound && tel?.replace(/\D/g,'').length >= 10 && !lookingUp;
  const nouveauClientValide = genre === 'Entreprise'
    ? !!entreprise?.trim() && emailValide(email||'')
    : !!genre && !!prenom?.trim() && !!nom?.trim() && emailValide(email||'');
  const clientOk = clientFound || (showNouveauClient ? nouveauClientValide : true);
  const resaValide = clientOk && tel?.replace(/\D/g,'').length >= 10 && dateIso && service && heure && nbPersonnes >= 1;

  const telValide = tel?.replace(/\D/g,'').length >= 10;
  const getConsigne = () => {
    if (!telValide)
      return { msg: 'Entrez un numéro de téléphone valide', ref: refTel, invalide: false };
    if (!clientFound && !genre)
      return { msg: 'Choisissez un genre', ref: refGenre, invalide: false };
    if (!clientFound && genre !== 'Entreprise' && !prenom?.trim())
      return { msg: 'Entrez un prénom', ref: refGenre, invalide: false };
    if (!clientFound && genre !== 'Entreprise' && !nom?.trim())
      return { msg: 'Entrez un nom', ref: refGenre, invalide: false };
    if (!clientFound && genre === 'Entreprise' && !entreprise?.trim())
      return { msg: "Entrez le nom de l'entreprise", ref: refGenre, invalide: false };
    if (!clientFound && email && !emailValide(email))
      return { msg: 'Email invalide — ex: prenom@gmail.com', ref: refEmail, invalide: true };
    if (!clientFound && !emailValide(email||''))
      return { msg: 'Entrez un email valide', ref: refEmail, invalide: false };
    if (!dateIso)
      return { msg: 'Choisissez une date', ref: refDate, invalide: false };
    if (!service)
      return { msg: 'Choisissez Midi ou Soir', ref: refService, invalide: false };
    if (!heure)
      return { msg: 'Choisissez une heure', ref: refHeure, invalide: false };
    return null;
  };
  const consigne = getConsigne();
  const formValide = !consigne;

  const calendarJSX = showCalPicker && (() => {
    const anneeP = calPickerDate.getFullYear();
    const moisP = calPickerDate.getMonth();
    const premierJourSemaine = new Date(anneeP, moisP, 1).getDay() || 7;
    const nbJours = new Date(anneeP, moisP + 1, 0).getDate();
    const casesP = Array(premierJourSemaine - 1).fill(null).concat(Array.from({length: nbJours}, (_, i) => i + 1));
    const todayIso = new Date().toISOString().split('T')[0];
    return (
      <div ref={calPickerRef} className={calFermeture ? 'cal-fermeture' : ''} style={{ marginTop:8, background:'#fff', borderRadius:12, border:'1.5px solid #eee', boxShadow:'0 4px 16px rgba(0,0,0,0.08)', overflow:'hidden' }}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'12px 12px', borderBottom:'1px solid #eee' }}>
          <button onPointerDown={()=>setCalPickerDate(new Date(anneeP, moisP-1))} style={{ width:40, height:40, borderRadius:10, border:'1.5px solid #ddd', background:'#fff', fontSize:20, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', touchAction:'manipulation', WebkitTapHighlightColor:'transparent' }}>‹</button>
          <span style={{ fontSize:15, fontWeight:800, color:'#111', textTransform:'capitalize' }}>{calPickerDate.toLocaleDateString('fr-FR', {month:'long', year:'numeric'})}</span>
          <button onPointerDown={()=>setCalPickerDate(new Date(anneeP, moisP+1))} style={{ width:40, height:40, borderRadius:10, border:'1.5px solid #ddd', background:'#fff', fontSize:20, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', touchAction:'manipulation', WebkitTapHighlightColor:'transparent' }}>›</button>
        </div>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', padding:'6px 6px 2px' }}>
          {['L','M','M','J','V','S','D'].map((j,i) => <div key={i} style={{ textAlign:'center', fontSize:11, fontWeight:700, color:'#aaa', padding:'3px 0' }}>{j}</div>)}
        </div>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', padding:'2px 6px 8px', gap:2 }}>
          {casesP.map((jour, i) => {
            if (!jour) return <div key={i}/>;
            const iso = `${anneeP}-${String(moisP+1).padStart(2,'0')}-${String(jour).padStart(2,'0')}`;
            const estAujourdhui = iso === todayIso;
            const estSelectionne = dateIso === iso;
            const aujourd2 = new Date(); aujourd2.setHours(0,0,0,0);
            const estPasse = new Date(anneeP, moisP, jour) < aujourd2;
            return (
              <button key={i} disabled={estPasse} className={dateFlash === iso ? 'date-flash' : ''} onPointerDown={()=>{ if (estPasse) return; setDateFlash(iso); setDateIso(iso); setTimeout(()=>{ setCalFermeture(true); setTimeout(()=>{ setShowCalPicker(false); setCalFermeture(false); setDateFlash(null); }, 300); }, 200); }} style={{
                height:44, borderRadius:10,
                border: estAujourdhui && !estSelectionne ? '2px solid #E8C547' : '1.5px solid transparent',
                background: estSelectionne ? '#E8C547' : 'transparent',
                fontWeight: estAujourdhui || estSelectionne ? 800 : 400,
                fontSize:15, cursor: estPasse ? 'not-allowed' : 'pointer',
                color: estPasse ? '#ccc' : '#111', opacity: estPasse ? 0.4 : 1,
                pointerEvents: estPasse ? 'none' : 'auto',
                touchAction:'manipulation', WebkitTapHighlightColor:'transparent'
              }}>{jour}</button>
            );
          })}
        </div>
      </div>
    );
  })();

  function handleClickBoutonDisabled() {
    if (!consigne) return;
    consigne.ref.current?.scrollIntoView({ behavior:'smooth', block:'center' });
    const el = consigne.ref.current;
    if (el) {
      el.style.borderColor = '#E8C547';
      el.style.boxShadow = '0 0 0 3px rgba(232,197,71,0.3)';
      el.style.transition = 'all 0.3s';
      setTimeout(()=>{ el.style.borderColor = '#eee'; el.style.boxShadow = 'none'; }, 2000);
    }
  }

  const fermerFormulaireResa = () => {
    const aDesDonnees = tel || prenom || nom || (heure && heure !== '') || (dateIso && dateIso !== (DATE_OPTS[0]?.iso));
    if (aDesDonnees && !resaCree) { setShowConfirmQuitter(true); } else { onClose(); }
  };

  const ctaFooter = !resaCree ? (
    <div style={{ width:'100%' }}>
      <button onClick={formValide ? handleSave : handleClickBoutonDisabled} disabled={saving} style={{ width:'100%', height:56, background: formValide ? '#E8C547' : '#f0f0f0', color: formValide ? '#111' : '#bbb', border:'none', borderRadius:14, fontSize:17, fontWeight:800, cursor: formValide ? 'pointer' : 'not-allowed', display:'flex', alignItems:'center', justifyContent:'center', gap:8, opacity: saving ? 0.6 : 1, transition:'all 0.3s', boxShadow: formValide ? '0 2px 8px rgba(232,197,71,0.3)' : 'none' }}>
        {saving ? 'Enregistrement…' : (isEdit ? '✏️ Modifier la réservation' : (formValide ? '✓ Créer la réservation' : 'Créer la réservation'))}
      </button>
      <div style={{ textAlign:'center', fontSize:12, marginTop:8, minHeight:20, transition:'opacity 0.2s' }}>
        {consigne && (
          <span style={{ color: consigne.invalide ? '#dc2626' : '#999' }}>
            {consigne.invalide ? '⚠️ ' : '→ '}{consigne.msg}
          </span>
        )}
      </div>
      <button onClick={fermerFormulaireResa} style={{ width:'100%', background:'none', border:'none', color:'#999', fontSize:14, cursor:'pointer', padding:'8px', marginTop:4 }}>Annuler</button>
    </div>
  ) : null;

  const formContent = (
    <>
      {resaCree && (
        <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', padding:32, textAlign:'center', minHeight:400 }}>
          <div style={{ width:72, height:72, borderRadius:'50%', background:'#f0fdf4', border:'3px solid #22c55e', display:'flex', alignItems:'center', justifyContent:'center', fontSize:32, marginBottom:20 }}>✓</div>
          <h2 style={{ fontSize:22, fontWeight:800, color:'#111', margin:'0 0 8px' }}>Réservation créée !</h2>
          <p style={{ color:'#666', fontSize:15, margin:'0 0 24px' }}>{resaCree.client.prenom} {resaCree.client.nom}</p>
          <div style={{ background:'#f9f9f9', borderRadius:12, padding:16, width:'100%', marginBottom:24, textAlign:'left' }}>
            <div style={{ display:'flex', justifyContent:'space-between', padding:'8px 0', borderBottom:'1px solid #eee' }}>
              <span style={{ color:'#999', fontSize:14 }}>Date</span>
              <span style={{ fontWeight:700, fontSize:14 }}>{new Date(resaCree.date+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long'})}</span>
            </div>
            <div style={{ display:'flex', justifyContent:'space-between', padding:'8px 0', borderBottom:'1px solid #eee' }}>
              <span style={{ color:'#999', fontSize:14 }}>Service</span>
              <span style={{ fontWeight:700, fontSize:14 }}>{resaCree.service==='midi'?'☀️ Midi':'🌙 Soir'}</span>
            </div>
            <div style={{ display:'flex', justifyContent:'space-between', padding:'8px 0', borderBottom:'1px solid #eee' }}>
              <span style={{ color:'#999', fontSize:14 }}>Heure</span>
              <span style={{ fontWeight:700, fontSize:14 }}>{resaCree.heure}</span>
            </div>
            <div style={{ display:'flex', justifyContent:'space-between', padding:'8px 0' }}>
              <span style={{ color:'#999', fontSize:14 }}>Personnes</span>
              <span style={{ fontWeight:700, fontSize:14 }}>{resaCree.nb_personnes} pers.</span>
            </div>
          </div>
          <button onClick={onClose} style={{ width:'100%', height:52, background:'#E8C547', border:'none', borderRadius:14, fontSize:16, fontWeight:800, cursor:'pointer', color:'#111', marginBottom:8 }}>✓ Parfait !</button>
          <button onClick={()=>{ setResaCree(null); setTel(''); setClientFound(null); setStatsClient(null); setPrenom(''); setNom(''); setEmail(''); setGenre(''); setDateIso(DATE_OPTS[0].iso); setService('soir'); setHeure(''); setNbPersonnes(2); setOccasion(''); setCommentaire(''); }} style={{ width:'100%', background:'none', border:'none', color:'#999', fontSize:14, cursor:'pointer', padding:'8px' }}>+ Ajouter une autre réservation</button>
        </div>
      )}
      {!resaCree && <>
      {!isEdit && <div style={{ background:'#fffbea', border:'1.5px solid #E8C547', borderRadius:10, padding:'10px 14px', marginBottom:16, display:'flex', alignItems:'center', gap:10 }}>
        <span style={{ fontSize:20 }}>⏳</span>
        <p style={{ margin:0, fontSize:13, color:'#92400e' }}>Cette réservation sera créée comme <strong>demande en attente</strong>.</p>
      </div>}

      <div style={{ display:'flex', flexDirection:'column', gap:20, paddingBottom:8 }}>

        {/* ── Section 1 : Téléphone ── */}
        <div ref={refTel}>
          <div style={{ fontSize:13, fontWeight:800, color:'#555', marginBottom:8, textTransform:'uppercase', letterSpacing:0.5 }}>1. Téléphone du client</div>
          <div style={{ position:'relative' }}>
            <span style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', fontSize:16, pointerEvents:'none' }}>📞</span>
            <input value={tel} onChange={e=>handleTelChange(e.target.value)} placeholder="06 12 34 56 78" type="tel" style={{ ...inp(false), paddingLeft:40 }} />
            {lookingUp && <span style={{ position:'absolute', right:12, top:'50%', transform:'translateY(-50%)', fontSize:12, color:'#888' }}>Recherche…</span>}
          </div>
          {clientFound && (
            <div style={{ marginTop:8 }}>
              <div style={{ background:'#f0fdf4', border:'1.5px solid #22c55e', borderRadius:10, padding:'10px 14px', display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
                <span style={{ fontSize:14 }}>✅</span>
                <span onClick={()=>{ if(onViewClient) onViewClient(clientFound); }} style={{ fontSize:14, fontWeight:800, color:'#111', flex:1, minWidth:0, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', cursor: onViewClient ? 'pointer' : 'default', textDecoration: onViewClient ? 'underline' : 'none', textDecorationColor:'#E8C547' }}>{clientFound.prenom} {clientFound.nom}</span>
                {statsClient && <>
                  <span style={{ fontSize:12, color:'#555', whiteSpace:'nowrap' }}>·&nbsp;{statsClient.total} résa</span>
                  {statsClient.noshow > 0 && <span style={{ fontSize:12, color:'#dc2626', whiteSpace:'nowrap' }}>·&nbsp;{statsClient.noshow} no-show</span>}
                </>}
              </div>
              <button onClick={()=>{ setEditClientForm({ prenom: clientFound.prenom||'', nom: clientFound.nom||'', mail: clientFound.mail||'', genre: clientFound.genre||'', entreprise: clientFound.entreprise||'' }); setShowEditClientInline(v=>!v); }} style={{ background:'none', border:'none', color:'#888', fontSize:12, cursor:'pointer', padding:'6px 2px', textDecoration:'underline' }}>Modifier les infos client ›</button>
              {showEditClientInline && (
                <div style={{ background:'#f9f9f9', borderRadius:10, padding:14, marginTop:8, border:'1.5px solid #eee' }}>
                  <p style={{ fontSize:12, fontWeight:700, color:'#999', marginBottom:10, textTransform:'uppercase', letterSpacing:0.5 }}>Modifier les infos client</p>
                  <div style={{ display:'flex', gap:8, marginBottom:10 }}>
                    {['Homme','Femme','Entreprise'].map(g => {
                      const sel = editClientForm.genre === g;
                      const s = GENRE_STYLES[g] || {};
                      return <button key={g} onClick={()=>setEditClientForm(f=>({...f,genre:g}))} style={{ flex:1, height:38, borderRadius:8, cursor:'pointer', fontSize:13, fontWeight:700, border: sel?`2px solid ${s.border}`:'1.5px solid #ddd', background: sel?s.bg:'#fff', color: sel?s.color:'#666', transition:'all 0.15s' }}>{g}</button>;
                    })}
                  </div>
                  {editClientForm.genre === 'Entreprise' && (
                    <input value={editClientForm.entreprise||''} onChange={e=>setEditClientForm(f=>({...f,entreprise:e.target.value}))} placeholder="Nom de l'entreprise" style={{ ...inp(false), marginBottom:8 }} />
                  )}
                  <div style={{ display:'flex', gap:8, marginBottom:8 }}>
                    <input value={editClientForm.prenom||''} onChange={e=>setEditClientForm(f=>({...f,prenom:e.target.value}))} placeholder="Prénom" style={{ ...inp(false), flex:1 }} />
                    <input value={editClientForm.nom||''} onChange={e=>setEditClientForm(f=>({...f,nom:e.target.value}))} placeholder="Nom" style={{ ...inp(false), flex:1 }} />
                  </div>
                  <input value={editClientForm.mail||''} onChange={e=>setEditClientForm(f=>({...f,mail:e.target.value}))} placeholder="Email" type="email" style={{ ...inp(false), marginBottom:12 }} />
                  <div style={{ display:'flex', gap:8 }}>
                    <button onClick={()=>setShowEditClientInline(false)} style={{ flex:1, height:40, border:'1.5px solid #ddd', borderRadius:8, background:'#fff', fontSize:13, cursor:'pointer', color:'#666' }}>Annuler</button>
                    {(()=>{ const ok = editClientForm.genre==='Entreprise' ? !!editClientForm.entreprise?.trim()&&emailValide(editClientForm.mail||'') : !!editClientForm.genre&&!!editClientForm.prenom?.trim()&&!!editClientForm.nom?.trim()&&emailValide(editClientForm.mail||''); return (
                    <button onClick={ok?async()=>{ await supabase.from('clients').update(editClientForm).eq('id', clientFound.id); setClientFound(prev=>({...prev,...editClientForm})); setShowEditClientInline(false); showToast('✅ Infos client mises à jour'); }:undefined} disabled={!ok} style={{ flex:2, height:40, background:ok?'#E8C547':'#f0f0f0', border:'none', borderRadius:8, fontSize:13, fontWeight:800, cursor:ok?'pointer':'not-allowed', color:ok?'#111':'#bbb', transition:'all 0.2s' }}>Enregistrer les modifications</button>
                    ); })()}
                  </div>
                </div>
              )}
            </div>
          )}
          {showNouveauClient && (
            <div ref={refGenre} style={{display:'flex', flexDirection:'column', gap:10, marginTop:10}}>
              <div>
                <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px'}}>Genre <span style={{color:'#dc2626'}}>*</span></p>
                <div style={{display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:8}}>
                  {[
                    {id:'Homme', label:'M. Monsieur', activeColor:'#1d4ed8', activeBg:'#dbeafe'},
                    {id:'Femme', label:'Mme Madame', activeColor:'#be185d', activeBg:'#fce7f3'},
                    {id:'Entreprise', label:'Entreprise', activeColor:'#15803d', activeBg:'#dcfce7'},
                  ].map(g=>(
                    <button key={g.id} onMouseDown={e=>e.stopPropagation()} onClick={e=>{e.stopPropagation(); setGenre(g.id); setPrenom(''); setNom(''); setEntreprise('');}} style={{height:44, borderRadius:10, cursor:'pointer', fontSize:13, fontWeight:700, border:'1.5px solid', borderColor: genre===g.id?g.activeBg:'#eee', background: genre===g.id?g.activeBg:'#fff', color: genre===g.id?g.activeColor:'#666', transition:'all 0.15s'}}>{g.label}</button>
                  ))}
                </div>
              </div>
              {genre === 'Entreprise' && (
                <div>
                  <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px'}}>Nom de l'entreprise <span style={{color:'#dc2626'}}>*</span></p>
                  <input value={entreprise} onChange={e=>setEntreprise(e.target.value)} placeholder="Nom de l'entreprise"
                    style={{width:'100%', height:48, border:'1.5px solid', borderColor: entreprise?.trim()?'#22c55e':'#eee', borderRadius:10, padding:'0 14px', fontSize:14, outline:'none', boxSizing:'border-box'}}
                    onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor=entreprise?.trim()?'#22c55e':'#eee'}/>
                </div>
              )}
              {genre && (
                <div>
                  <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px'}}>
                    {genre==='Entreprise' ? <>Nom du contact <span style={{fontSize:12, fontWeight:400, color:'#999'}}>(optionnel)</span></> : <>Prénom et Nom <span style={{color:'#dc2626'}}>*</span></>}
                  </p>
                  <div style={{display:'flex', gap:8}}>
                    <input value={prenom} onChange={e=>setPrenom(e.target.value)} placeholder="Prénom"
                      style={{flex:1, height:48, border:'1.5px solid #eee', borderRadius:10, padding:'0 14px', fontSize:14, outline:'none', boxSizing:'border-box'}}
                      onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                    <input value={nom} onChange={e=>setNom(e.target.value)} placeholder="Nom"
                      style={{flex:1, height:48, border:'1.5px solid #eee', borderRadius:10, padding:'0 14px', fontSize:14, outline:'none', boxSizing:'border-box'}}
                      onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                  </div>
                </div>
              )}
              {genre && (
                <div>
                  <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px'}}>Email <span style={{color:'#dc2626'}}>*</span></p>
                  <input ref={refEmail} type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="prenom.nom@gmail.com"
                    style={{width:'100%', height:48, border:'1.5px solid', borderColor:'#eee', borderRadius:10, padding:'0 14px', fontSize:14, outline:'none', boxSizing:'border-box'}}
                    onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                </div>
              )}
              <p style={{fontSize:11, color:'#999', margin:'2px 0 0', textAlign:'right'}}><span style={{color:'#dc2626'}}>*</span> Champs obligatoires</p>
            </div>
          )}
        </div>

        {/* ── Section 2 : Quand ? ── */}
        <div ref={refDate}>
          <div style={{ fontSize:13, fontWeight:800, color:'#555', marginBottom:8, textTransform:'uppercase', letterSpacing:0.5 }}>2. Quand ?</div>
          <button onPointerDown={()=>setShowCalPicker(!showCalPicker)} style={{ width:'100%', height:48, border:`1.5px solid ${showCalPicker ? '#E8C547' : '#ddd'}`, borderRadius:10, background:'#fff', fontSize:14, fontWeight:600, cursor:'pointer', textAlign:'left', padding:'0 14px', color: dateIso ? '#111' : '#aaa', display:'flex', alignItems:'center', justifyContent:'space-between', touchAction:'manipulation', WebkitTapHighlightColor:'transparent' }}>
            <span>📅 {dateIso ? new Date(dateIso+'T12:00:00').toLocaleDateString('fr-FR', {weekday:'long', day:'numeric', month:'long', year:'numeric'}) : 'Choisir une date'}</span>
            <span style={{ color:'#ccc', fontSize:20 }}>›</span>
          </button>
          {calendarJSX}
          <div ref={refService} style={{ display:'flex', gap:8, marginTop:10 }}>
            <button style={btnSvc('midi')} onClick={()=>{ setService('midi'); setHeure(''); setHeureError(false); }}>☀️ Midi</button>
            <button style={btnSvc('soir')} onClick={()=>{ setService('soir'); setHeure(''); setHeureError(false); }}>🌙 Soir</button>
          </div>
          {service && (
            <div ref={refHeure} style={{ display:'flex', flexWrap:'wrap', gap:6, marginTop:10 }}>
              {heures.map(h => (
                <button key={h} onMouseDown={e=>e.stopPropagation()} onClick={e=>{e.stopPropagation(); setHeure(heure===h?'':h); setHeureError(false);}} style={{ padding:'8px 14px', borderRadius:20, border:`1.5px solid ${heure===h?'#111':heureError?'#dc2626':'#eee'}`, background:heure===h?'#111':'#f8f8f8', color:heure===h?'#fff':'#555', fontWeight:700, fontSize:13, cursor:'pointer' }}>{h}</button>
              ))}
            </div>
          )}
          {heureError && <p style={{ fontSize:12, color:'#dc2626', marginTop:6 }}>* Sélectionnez un créneau horaire</p>}
        </div>

        {/* ── Section 3 : Combien ? ── */}
        <div>
          <div style={{ fontSize:13, fontWeight:800, color:'#555', marginBottom:8, textTransform:'uppercase', letterSpacing:0.5 }}>3. Combien de personnes ?</div>
          <div style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:0, border:'1.5px solid #eee', borderRadius:12, overflow:'hidden', width:'100%' }}>
            <button style={{ width:64, height:64, background:'#f8f8f8', border:'none', borderRight:'1.5px solid #eee', fontSize:28, fontWeight:700, cursor:'pointer', color:'#111', flexShrink:0 }} onClick={()=>setNbPersonnes(n=>{const v=typeof n==='number'&&n>0?n:1;return Math.max(1,v-1);})}>−</button>
            <input type="number" inputMode="numeric" pattern="[0-9]*" min={1} max={500} value={nbPersonnes === undefined || nbPersonnes === '' ? '' : nbPersonnes} onChange={e=>{ const v=e.target.value; if(v===''||v==='0'){ setNbPersonnes(''); } else { const val=parseInt(v); if(!isNaN(val)&&val>=1&&val<=500) setNbPersonnes(val); } }} onBlur={()=>{ if(!nbPersonnes||nbPersonnes<1) setNbPersonnes(1); if(nbPersonnes>500) setNbPersonnes(500); }} style={{ flex:1, height:64, border:'none', textAlign:'center', fontSize:28, fontWeight:800, outline:'none', color:'#111' }} />
            <button style={{ width:64, height:64, background:'#f8f8f8', border:'none', borderLeft:'1.5px solid #eee', fontSize:28, fontWeight:700, cursor:'pointer', color:'#111', flexShrink:0 }} onClick={()=>setNbPersonnes(n=>{const v=typeof n==='number'&&n>0?n:1;return Math.min(500,v+1);})}>+</button>
          </div>
        </div>

        {/* ── Section 4 : Occasion & Commentaire ── */}
        <div>
          <div style={{ fontSize:13, fontWeight:800, color:'#555', marginBottom:8, textTransform:'uppercase', letterSpacing:0.5 }}>4. Occasion & Commentaire <span style={{ fontWeight:400, color:'#bbb' }}>(optionnels)</span></div>
          <select value={occasion} onChange={e=>setOccasion(e.target.value)} style={{ width:'100%', height:44, border:'1.5px solid #ddd', borderRadius:8, padding:'0 12px', fontSize:14, background:'#fff', outline:'none', cursor:'pointer', marginBottom:10 }}>
            <option value="">— Aucune occasion —</option>
            {OCCASIONS.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
          <textarea value={commentaire} onChange={e=>setCommentaire(e.target.value)} placeholder="Allergies, demandes particulières…" rows={3} style={{ width:'100%', minHeight:70, border:'1.5px solid #ddd', borderRadius:8, padding:10, fontSize:14, resize:'none', outline:'none', boxSizing:'border-box', fontFamily:'inherit' }} />
        </div>

      </div>

      </>}
    </>
  );

  return (
    <>
    {isMobile ? (
      <div style={{ position:'fixed', inset:0, background:'#f8f8f8', zIndex:2000, display:'flex', flexDirection:'column' }}>
        <div style={{ background:'#111', padding:'16px 20px', paddingTop:'calc(16px + env(safe-area-inset-top))', borderBottom:'3px solid #E8C547', flexShrink:0, display:'flex', alignItems:'center', justifyContent:'space-between' }}>
          <h2 style={{ color:'#fff', margin:0, fontSize:18, fontWeight:800 }}>{isEdit ? 'Modifier la réservation' : 'Nouvelle réservation'}</h2>
          <button type="button" onClick={fermerFormulaireResa} style={{ background:'none', border:'none', color:'#fff', fontSize:20, cursor:'pointer', touchAction:'manipulation' }}>✕</button>
        </div>
        <div style={{ flex:1, overflowY:'auto', padding:'16px', WebkitOverflowScrolling:'touch' }}>
          {formContent}
        </div>
        {ctaFooter && <div style={{ background:'#fff', padding:'12px 16px', paddingBottom:'calc(12px + env(safe-area-inset-bottom))', borderTop:'1px solid #eee', flexShrink:0 }}>{ctaFooter}</div>}
      </div>
    ) : (
      <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:2000, display:'flex', alignItems:'center', justifyContent:'center', pointerEvents:'all', cursor:'default', touchAction:'none' }}
        onMouseDown={e=>{e.preventDefault();e.stopPropagation();fermerFormulaireResa();}}>
        <div style={{ background:'#fff', borderRadius:20, width:'min(560px, calc(100vw - 48px))', maxHeight:'90vh', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.25)', overflow:'hidden' }}
          onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()}>

          {/* Header */}
          <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'24px 28px 16px', flexShrink:0 }}>
            <h2 style={{ margin:0, fontSize:22, fontWeight:800, color:'#111' }}>{isEdit ? 'Modifier la réservation' : 'Nouvelle réservation'}</h2>
            <button onClick={fermerFormulaireResa} style={{ width:36, height:36, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', fontSize:18, color:'#666' }}>✕</button>
          </div>

          {/* Contenu scrollable */}
          <div style={{ flex:1, overflowY:'auto', padding:'0 28px 20px' }}>
          {!resaCree && !isEdit && (
            <div style={{ background:'#fffbea', border:'1.5px solid #E8C547', borderRadius:10, padding:'10px 14px', marginBottom:16 }}>
              <p style={{ margin:0, fontSize:13, color:'#92400e' }}>Cette réservation sera créée comme <strong>demande en attente</strong>.</p>
            </div>
          )}
            {resaCree && (
              <div style={{ display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', padding:32, textAlign:'center', minHeight:340 }}>
                <div style={{ width:72, height:72, borderRadius:'50%', background:'#f0fdf4', border:'3px solid #22c55e', display:'flex', alignItems:'center', justifyContent:'center', fontSize:32, marginBottom:20 }}>✓</div>
                <h2 style={{ fontSize:22, fontWeight:800, color:'#111', margin:'0 0 8px' }}>Réservation créée !</h2>
                <p style={{ color:'#666', fontSize:15, margin:'0 0 24px' }}>{resaCree.client.prenom} {resaCree.client.nom}</p>
                <div style={{ background:'#f9f9f9', borderRadius:12, padding:16, width:'100%', marginBottom:24, textAlign:'left' }}>
                  {[['Date', new Date(resaCree.date+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long'})],['Service',resaCree.service==='midi'?'☀️ Midi':'🌙 Soir'],['Heure',resaCree.heure],['Personnes',`${resaCree.nb_personnes} pers.`]].map(([k,v],i,arr)=>(
                    <div key={k} style={{ display:'flex', justifyContent:'space-between', padding:'8px 0', borderBottom:i<arr.length-1?'1px solid #eee':'none' }}>
                      <span style={{ color:'#999', fontSize:14 }}>{k}</span>
                      <span style={{ fontWeight:700, fontSize:14 }}>{v}</span>
                    </div>
                  ))}
                </div>
                <button onClick={onClose} style={{ width:'100%', height:52, background:'#E8C547', border:'none', borderRadius:14, fontSize:16, fontWeight:800, cursor:'pointer', color:'#111', marginBottom:8 }}>✓ Parfait !</button>
                <button onClick={()=>{ setResaCree(null); setTel(''); setClientFound(null); setStatsClient(null); setPrenom(''); setNom(''); setEmail(''); setGenre(''); setDateIso(DATE_OPTS[0].iso); setService('soir'); setHeure(''); setNbPersonnes(2); setOccasion(''); setCommentaire(''); }} style={{ width:'100%', background:'none', border:'none', color:'#999', fontSize:14, cursor:'pointer', padding:'8px' }}>+ Ajouter une autre réservation</button>
              </div>
            )}

            {!resaCree && (
              <div style={{ display:'flex', flexDirection:'column', gap:0 }}>

                {/* 1. Téléphone */}
                <div ref={refTel} style={{ marginBottom:24, marginTop:8 }}>
                  <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>1. Téléphone du client</p>
                  <div style={{ position:'relative' }}>
                    <Phone size={18} strokeWidth={2} color="#999" style={{ position:'absolute', left:14, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }} />
                    <input type="tel" inputMode="numeric" value={tel} onChange={e=>handleTelChange(e.target.value)} placeholder="06 43 00 49 87"
                      style={{ width:'100%', height:52, border:'1.5px solid #eee', borderRadius:12, padding:'0 46px', fontSize:16, outline:'none', boxSizing:'border-box' }} />
                    {clientFound && <CircleCheck size={20} color="#22c55e" style={{ position:'absolute', right:14, top:'50%', transform:'translateY(-50%)' }} />}
                    {lookingUp && <span style={{ position:'absolute', right:14, top:'50%', transform:'translateY(-50%)', fontSize:12, color:'#888' }}>Recherche…</span>}
                  </div>

                  {clientFound && (
                    <div onClick={()=>{ if(onViewClient) onViewClient(clientFound); }} style={{ marginTop:8, background:'#f0fdf4', border:'1.5px solid #22c55e', borderRadius:10, padding:'10px 14px', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between' }}>
                      <div style={{ display:'flex', alignItems:'center', gap:10 }}>
                        <User size={18} strokeWidth={2} color="#16a34a" />
                        <div>
                          <span style={{ fontWeight:800, fontSize:14, color:'#111' }}>{clientFound.prenom} {clientFound.nom}</span>
                          <div style={{ fontSize:12, color:'#666', marginTop:2 }}>
                            {statsClient?.total} réservations · {statsClient?.noshow} no-show{statsClient ? '' : ''} · <span style={{ color:'#16a34a', fontWeight:600 }}>Voir la fiche</span>
                          </div>
                        </div>
                      </div>
                      <ChevronRight size={16} color="#16a34a" />
                    </div>
                  )}

                  {clientFound && (
                    <button onClick={()=>{ setEditClientForm({ prenom: clientFound.prenom||'', nom: clientFound.nom||'', mail: clientFound.mail||'', genre: clientFound.genre||'', entreprise: clientFound.entreprise||'' }); setShowEditClientInline(v=>!v); }} style={{ marginTop:8, width:'100%', padding:'8px 14px', background:'none', border:'1.5px solid #eee', borderRadius:8, fontSize:13, color:'#666', cursor:'pointer', textAlign:'left', display:'flex', alignItems:'center', gap:8 }}>
                      <Pencil size={14} strokeWidth={2} color="#999" /> Modifier les informations du client
                    </button>
                  )}
                  {showEditClientInline && clientFound && (
                    <div style={{ background:'#f9f9f9', borderRadius:10, padding:14, marginTop:8, border:'1.5px solid #eee' }}>
                      <p style={{ fontSize:12, fontWeight:700, color:'#999', marginBottom:10, textTransform:'uppercase', letterSpacing:0.5 }}>Modifier les infos client</p>
                      <div style={{ display:'flex', gap:8, marginBottom:10 }}>
                        {['Homme','Femme','Entreprise'].map(g => {
                          const sel = editClientForm.genre === g;
                          const s2 = GENRE_STYLES[g] || {};
                          return <button key={g} onClick={()=>setEditClientForm(f=>({...f,genre:g}))} style={{ flex:1, height:38, borderRadius:8, cursor:'pointer', fontSize:13, fontWeight:700, border: sel?`2px solid ${s2.border}`:'1.5px solid #ddd', background: sel?s2.bg:'#fff', color: sel?s2.color:'#666' }}>{g}</button>;
                        })}
                      </div>
                      <div style={{ display:'flex', gap:8, marginBottom:8 }}>
                        <input value={editClientForm.prenom||''} onChange={e=>setEditClientForm(f=>({...f,prenom:e.target.value}))} placeholder="Prénom" style={{ flex:1, height:44, border:'1.5px solid #eee', borderRadius:8, padding:'0 12px', fontSize:14, outline:'none' }} />
                        <input value={editClientForm.nom||''} onChange={e=>setEditClientForm(f=>({...f,nom:e.target.value}))} placeholder="Nom" style={{ flex:1, height:44, border:'1.5px solid #eee', borderRadius:8, padding:'0 12px', fontSize:14, outline:'none' }} />
                      </div>
                      <input value={editClientForm.mail||''} onChange={e=>setEditClientForm(f=>({...f,mail:e.target.value}))} placeholder="Email" type="email" style={{ width:'100%', height:44, border:'1.5px solid #eee', borderRadius:8, padding:'0 12px', fontSize:14, outline:'none', boxSizing:'border-box', marginBottom:12 }} />
                      <div style={{ display:'flex', gap:8 }}>
                        <button onClick={()=>setShowEditClientInline(false)} style={{ flex:1, height:40, border:'1.5px solid #ddd', borderRadius:8, background:'#fff', fontSize:13, cursor:'pointer', color:'#666' }}>Annuler</button>
                        {(()=>{ const ok = editClientForm.genre==='Entreprise' ? !!editClientForm.entreprise?.trim()&&emailValide(editClientForm.mail||'') : !!editClientForm.genre&&!!editClientForm.prenom?.trim()&&!!editClientForm.nom?.trim()&&emailValide(editClientForm.mail||''); return (
                        <button onClick={ok?async()=>{ await supabase.from('clients').update(editClientForm).eq('id', clientFound.id); setClientFound(prev=>({...prev,...editClientForm})); setShowEditClientInline(false); showToast('✅ Infos client mises à jour'); }:undefined} disabled={!ok} style={{ flex:2, height:40, background:ok?'#E8C547':'#f0f0f0', border:'none', borderRadius:8, fontSize:13, fontWeight:800, cursor:ok?'pointer':'not-allowed', color:ok?'#111':'#bbb', transition:'all 0.2s' }}>Enregistrer les modifications</button>
                        ); })()}
                      </div>
                    </div>
                  )}

                  {showNouveauClient && (
                    <div ref={refGenre} style={{display:'flex', flexDirection:'column', gap:10, marginTop:10}}>
                      <div>
                        <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px'}}>Genre <span style={{color:'#dc2626'}}>*</span></p>
                        <div style={{display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:8}}>
                          {[
                            {id:'Homme', label:'M. Monsieur', activeColor:'#1d4ed8', activeBg:'#dbeafe'},
                            {id:'Femme', label:'Mme Madame', activeColor:'#be185d', activeBg:'#fce7f3'},
                            {id:'Entreprise', label:'Entreprise', activeColor:'#15803d', activeBg:'#dcfce7'},
                          ].map(g=>(
                            <button key={g.id} onMouseDown={e=>e.stopPropagation()} onClick={e=>{e.stopPropagation(); setGenre(g.id); setPrenom(''); setNom(''); setEntreprise('');}} style={{height:44, borderRadius:10, cursor:'pointer', fontSize:13, fontWeight:700, border:'1.5px solid', borderColor: genre===g.id?g.activeBg:'#eee', background: genre===g.id?g.activeBg:'#fff', color: genre===g.id?g.activeColor:'#666', transition:'all 0.15s'}}>{g.label}</button>
                          ))}
                        </div>
                      </div>
                      {genre === 'Entreprise' && (
                        <div>
                          <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px'}}>Nom de l'entreprise <span style={{color:'#dc2626'}}>*</span></p>
                          <input value={entreprise} onChange={e=>setEntreprise(e.target.value)} placeholder="Nom de l'entreprise"
                            style={{width:'100%', height:48, border:'1.5px solid', borderColor: entreprise?.trim()?'#22c55e':'#eee', borderRadius:10, padding:'0 14px', fontSize:14, outline:'none', boxSizing:'border-box'}}
                            onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor=entreprise?.trim()?'#22c55e':'#eee'}/>
                        </div>
                      )}
                      {genre && (
                        <div>
                          <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px'}}>
                            {genre==='Entreprise' ? <>Nom du contact <span style={{fontSize:12, fontWeight:400, color:'#999'}}>(optionnel)</span></> : <>Prénom et Nom <span style={{color:'#dc2626'}}>*</span></>}
                          </p>
                          <div style={{display:'flex', gap:8}}>
                            <input value={prenom} onChange={e=>setPrenom(e.target.value)} placeholder="Prénom"
                              style={{flex:1, height:48, border:'1.5px solid #eee', borderRadius:10, padding:'0 14px', fontSize:14, outline:'none', boxSizing:'border-box'}}
                              onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                            <input value={nom} onChange={e=>setNom(e.target.value)} placeholder="Nom"
                              style={{flex:1, height:48, border:'1.5px solid #eee', borderRadius:10, padding:'0 14px', fontSize:14, outline:'none', boxSizing:'border-box'}}
                              onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                          </div>
                        </div>
                      )}
                      {genre && (
                        <div>
                          <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px'}}>Email <span style={{color:'#dc2626'}}>*</span></p>
                          <input ref={refEmail} type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="prenom.nom@gmail.com"
                            style={{width:'100%', height:48, border:'1.5px solid', borderColor: emailValide(email||'')?'#22c55e':'#eee', borderRadius:10, padding:'0 14px', fontSize:14, outline:'none', boxSizing:'border-box'}}
                            onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor=emailValide(email||'')?'#22c55e':'#eee'}/>
                        </div>
                      )}
                      <p style={{fontSize:11, color:'#999', margin:'2px 0 0', textAlign:'right'}}><span style={{color:'#dc2626'}}>*</span> Champs obligatoires</p>
                    </div>
                  )}
                </div>

                {/* 2. Date */}
                <div ref={refDate} style={{ marginBottom:24 }}>
                  <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>2. Date</p>
                  <button onPointerDown={()=>setShowCalPicker(!showCalPicker)} style={{ width:'100%', height:52, border:`1.5px solid ${showCalPicker?'#E8C547':'#eee'}`, borderRadius:12, background:'#fff', display:'flex', alignItems:'center', gap:12, padding:'0 16px', cursor:'pointer', boxSizing:'border-box', touchAction:'manipulation' }}>
                    <CalendarDays size={18} strokeWidth={2} color="#999" />
                    <span style={{ flex:1, textAlign:'left', fontSize:15, color:dateIso?'#111':'#bbb' }}>
                      {dateIso ? new Date(dateIso+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long',year:'numeric'}) : 'Choisir une date'}
                    </span>
                    <ChevronDown size={16} color="#999" />
                  </button>
                  {calendarJSX}
                </div>

                {/* 3. Service */}
                <div ref={refService} style={{ marginBottom:24 }}>
                  <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>3. Service</p>
                  <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10 }}>
                    <button onClick={()=>{ setService('midi'); setHeure(''); setHeureError(false); }} style={{ height:52, borderRadius:12, cursor:'pointer', fontSize:15, fontWeight:700, border:`1.5px solid ${service==='midi'?'#E8C547':'#eee'}`, background:service==='midi'?'#fffbea':'#fff', color:'#111', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                      <Sun size={18} strokeWidth={2} color={service==='midi'?'#E8C547':'#999'} /> Midi
                    </button>
                    <button onClick={()=>{ setService('soir'); setHeure(''); setHeureError(false); }} style={{ height:52, borderRadius:12, cursor:'pointer', fontSize:15, fontWeight:700, border:service==='soir'?'none':'1.5px solid #eee', background:service==='soir'?'#111':'#fff', color:service==='soir'?'#E8C547':'#111', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                      <Moon size={18} strokeWidth={2} color={service==='soir'?'#E8C547':'#999'} /> Soir
                    </button>
                  </div>
                </div>

                {/* 4. Heure */}
                {service && (
                  <div ref={refHeure} style={{ marginBottom:24 }}>
                    <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>4. Heure</p>
                    <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:8 }}>
                      {heures.map(h=>(
                        <button key={h} onMouseDown={e=>e.stopPropagation()} onClick={e=>{e.stopPropagation(); setHeure(heure===h?'':h); setHeureError(false);}} style={{ height:44, borderRadius:10, cursor:'pointer', fontSize:14, fontWeight:600, border:`1.5px solid ${heure===h?'#111':heureError?'#dc2626':'#eee'}`, background:heure===h?'#111':'#fff', color:heure===h?'#E8C547':'#111' }}>{h}</button>
                      ))}
                    </div>
                    {heureError && <p style={{ fontSize:12, color:'#dc2626', marginTop:6 }}>* Sélectionnez un créneau horaire</p>}
                  </div>
                )}

                {/* 5. Nombre de personnes */}
                <div style={{ marginBottom:24 }}>
                  <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>5. Nombre de personnes</p>
                  <div style={{ display:'flex', alignItems:'center', gap:16, justifyContent:'center' }}>
                    <button onClick={()=>setNbPersonnes(n=>{const v=typeof n==='number'&&n>0?n:1;return Math.max(1,v-1);})} style={{ width:52, height:52, borderRadius:12, border:'1.5px solid #eee', background:'#fff', cursor:'pointer', fontSize:24, color:'#111', display:'flex', alignItems:'center', justifyContent:'center', fontWeight:300 }}>−</button>
                    <div style={{ textAlign:'center', minWidth:80 }}>
                      <input
                        type="number"
                        inputMode="numeric"
                        pattern="[0-9]*"
                        min={1}
                        max={500}
                        value={nbPersonnes===undefined||nbPersonnes===null?'':nbPersonnes}
                        onChange={e=>{ const raw=e.target.value; if(raw===''||raw==='0'){setNbPersonnes('');}else{const val=parseInt(raw);if(!isNaN(val)&&val>=1&&val<=500)setNbPersonnes(val);} }}
                        onFocus={e=>e.target.select()}
                        onBlur={()=>{ if(!nbPersonnes||nbPersonnes<1)setNbPersonnes(1); if(nbPersonnes>500)setNbPersonnes(500); }}
                        style={{ width:80, height:52, fontSize:32, fontWeight:800, color:'#111', textAlign:'center', border:'1.5px solid #eee', borderRadius:12, outline:'none', background:'#fff', cursor:'text', MozAppearance:'textfield' }}
                      />
                      <div style={{ fontSize:12, color:'#999', marginTop:4 }}>pers.</div>
                    </div>
                    <button onClick={()=>setNbPersonnes(n=>{const v=typeof n==='number'&&n>0?n:1;return Math.min(500,v+1);})} style={{ width:52, height:52, borderRadius:12, border:'1.5px solid #eee', background:'#fff', cursor:'pointer', fontSize:24, color:'#111', display:'flex', alignItems:'center', justifyContent:'center', fontWeight:300 }}>+</button>
                  </div>
                </div>

                {/* Occasion & Commentaire */}
                <div style={{ marginBottom:8 }}>
                  <select value={occasion} onChange={e=>setOccasion(e.target.value)} style={{ width:'100%', height:44, border:'1.5px solid #eee', borderRadius:10, padding:'0 12px', fontSize:14, outline:'none', background:'#fff', marginBottom:10, boxSizing:'border-box' }}>
                    <option value="">— Aucune occasion particulière</option>
                    {OCCASIONS.map(o=><option key={o} value={o}>{o}</option>)}
                  </select>
                  <textarea value={commentaire} onChange={e=>setCommentaire(e.target.value)} placeholder="Commentaire (allergies, demandes particulières...)"
                    style={{ width:'100%', height:80, border:'1.5px solid #eee', borderRadius:10, padding:'10px 12px', fontSize:14, outline:'none', resize:'none', fontFamily:'inherit', boxSizing:'border-box' }} />
                </div>
              </div>
            )}
          </div>

          {/* Footer fixe */}
          {!resaCree && (
            <div style={{ flexShrink:0, padding:'16px 28px', borderTop:'1px solid #eee', background:'#fff' }}>
              <button onClick={formValide ? handleSave : handleClickBoutonDisabled} disabled={saving} style={{ width:'100%', height:54, background:formValide?'#E8C547':'#f0f0f0', color:formValide?'#111':'#bbb', border:'none', borderRadius:14, fontSize:16, fontWeight:800, cursor:formValide?'pointer':'not-allowed', transition:'all 0.3s', boxShadow:formValide?'0 2px 8px rgba(232,197,71,0.3)':'none' }}>
                {saving ? 'Enregistrement...' : (isEdit ? '✏️ Modifier la réservation' : (formValide ? '✓ Créer la réservation' : 'Créer la réservation'))}
              </button>
              <div style={{ textAlign:'center', fontSize:12, marginTop:8, minHeight:20, transition:'opacity 0.2s' }}>
                {consigne && (
                  <span style={{ color: consigne.invalide ? '#dc2626' : '#999' }}>
                    {consigne.invalide ? '⚠️ ' : '→ '}{consigne.msg}
                  </span>
                )}
              </div>
            </div>
          )}
        </div>
      </div>
    )}

    {showConfirmQuitter && (
      <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:6000, display:'flex', alignItems:'center', justifyContent:'center', pointerEvents:'all', cursor:'default', touchAction:'none' }} onMouseDown={e=>{e.preventDefault();e.stopPropagation();}}>
        <div style={{ background:'#fff', borderRadius:16, padding:'28px 24px', maxWidth:320, width:'90%', textAlign:'center', boxShadow:'0 20px 60px rgba(0,0,0,0.2)' }}>
          <div style={{ fontSize:40, marginBottom:12 }}>⚠️</div>
          <h3 style={{ margin:'0 0 8px', fontSize:17, fontWeight:800, color:'#111' }}>Quitter sans enregistrer ?</h3>
          <p style={{ margin:'0 0 20px', fontSize:14, color:'#666' }}>Les informations saisies seront perdues.</p>
          <div style={{ display:'flex', gap:10 }}>
            <button onClick={()=>setShowConfirmQuitter(false)} style={{ flex:1, height:44, border:'1.5px solid #ddd', borderRadius:10, background:'#fff', fontSize:14, fontWeight:600, cursor:'pointer', color:'#666' }}>Continuer la saisie</button>
            <button onClick={()=>{ setShowConfirmQuitter(false); onClose(); }} style={{ flex:1, height:44, border:'none', borderRadius:10, background:'#dc2626', fontSize:14, fontWeight:800, cursor:'pointer', color:'#fff' }}>Quitter</button>
          </div>
        </div>
      </div>
    )}
    </>
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
    <Modal title="Refuser la réservation" onClose={onCancel} maxW={400} zIndex={4000}
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
    <>
      <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();}} onClick={onCancel} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:4999, pointerEvents:'all' }}/>
      <div onClick={e=>e.stopPropagation()} onMouseDown={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:20, width:'min(440px, calc(100vw - 48px))', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.25)', zIndex:5000, overflow:'hidden' }}>
        {/* Header */}
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'24px 28px 20px', flexShrink:0, borderBottom:'1px solid #f0f0f0' }}>
          <h2 style={{margin:0, fontSize:20, fontWeight:800, color:'#111'}}>Confirmer la réservation</h2>
          <button onClick={onCancel} style={{ width:36, height:36, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:18, color:'#666', display:'flex', alignItems:'center', justifyContent:'center' }}>✕</button>
        </div>
        {/* Contenu */}
        <div style={{padding:'20px 28px 24px', display:'flex', flexDirection:'column', gap:14}}>
          {/* Nom */}
          <div style={{textAlign:'center', marginBottom:4}}>
            <h3 style={{fontSize:22, fontWeight:900, color:'#111', margin:0}}>{nom || '—'}</h3>
          </div>
          {/* Infos */}
          <div style={{display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:10}}>
            {[
              {label:'Date', value: fmtResaDate(resa.date)},
              {label:'Service', value: resa.service==='midi'?`☀️ Midi · ${resa.heure}`:`🌙 Soir · ${resa.heure}`},
              {label:'Personnes', value: `👥 ${resa.nb_personnes} pers.`},
            ].map((item,i)=>(
              <div key={i} style={{background:'#f9f9f9', borderRadius:10, padding:'10px 12px', textAlign:'center'}}>
                <div style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, marginBottom:4}}>{item.label}</div>
                <div style={{fontSize:13, fontWeight:700, color:'#111'}}>{item.value}</div>
              </div>
            ))}
          </div>
          {/* Boutons */}
          <div style={{display:'flex', gap:10, marginTop:4}}>
            <button onClick={onCancel} style={{ flex:1, height:52, border:'1.5px solid #eee', borderRadius:12, background:'#fff', fontSize:15, fontWeight:600, cursor:'pointer', color:'#666' }}>Annuler</button>
            <button onClick={onConfirm} style={{ flex:2, height:52, border:'none', borderRadius:12, background:'#E8C547', fontSize:15, fontWeight:800, cursor:'pointer', color:'#111', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
              <Check size={18} strokeWidth={2}/> Confirmer
            </button>
          </div>
        </div>
      </div>
    </>
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

function DetailResaModal({ resa, onClose, onSaved, onEdit, resaList = [], showToast }) {
  const c = resa.clients || {};
  const nom = c.entreprise ? c.entreprise : `${c.prenom || ''} ${c.nom || ''}`.trim();
  const [statut, setStatut] = useState(resa.statut);
  const [statutEnCours, setStatutEnCours] = useState(resa.statut);
  const [statutModifie, setStatutModifie] = useState(false);
  const [saving, setSaving] = useState(false);
  const [showSmsPanel, setShowSmsPanel] = useState(false);
  const [smsTexte, setSmsTexte] = useState('');
  const [showStatutPanel, setShowStatutPanel] = useState(false);

  const STATUTS_COLORS = [
    { value:'confirmee', label:'Confirmée',  desc:'La réservation est confirmée',    color:'#16a34a' },
    { value:'attente',   label:'En attente', desc:'Demande en attente',              color:'#f59e0b' },
    { value:'absente',   label:'Absente',    desc:"Le client ne s'est pas présenté", color:'#dc2626' },
    { value:'annulee',   label:'Annulée',    desc:'Réservation annulée',             color:'#9ca3af' },
  ];

  const aujourd = new Date().toISOString().split('T')[0];
  const demain = new Date(Date.now() + 86400000).toISOString().split('T')[0];

  const smsSuggestions = [
    resa.date === demain
      ? `Bonjour ${c.prenom || ''} 👋 Rappel : votre résa au TED est demain à ${resa.heure} pour ${resa.nb_personnes} pers. À demain !`
      : resa.date < aujourd
      ? `Bonjour ${c.prenom || ''}, merci pour votre visite au TED. À bientôt ! 🙏`
      : `Bonjour ${c.prenom || ''}, votre résa au TED le ${new Date(resa.date + 'T12:00:00').toLocaleDateString('fr-FR', {weekday:'long', day:'numeric', month:'long'})} à ${resa.heure} est confirmée ✅`,
    `Bonjour ${c.prenom || ''}, pouvez-vous confirmer votre présence au TED ? Merci 🙏`,
  ];

  const resasClient = resaList.filter(r => r.client_id === resa.client_id);
  const nbVenues = resasClient.filter(r => r.statut === 'venue').length;
  const nbAbsentes = resasClient.filter(r => r.statut === 'absente').length;
  const totalResas = resasClient.filter(r => r.statut !== 'annulee' && r.statut !== 'absente' && r.statut !== 'refusee').length;
  const noshow = nbAbsentes;
  const derniereVisite = resasClient
    .filter(r => (r.statut === 'venue' || r.statut === 'confirmee') && r.date <= aujourd)
    .sort((a,b) => b.date.localeCompare(a.date))[0];
  const derniereVisiteFormatee = derniereVisite
    ? new Date(derniereVisite.date+'T12:00:00').toLocaleDateString('fr-FR', {day:'numeric', month:'long', year:'numeric'})
    : 'Jamais';

  function fermerModal() {
    if (statutModifie) {
      const confirme = window.confirm('Vous avez modifié le statut sans valider. Quitter sans sauvegarder ?');
      if (!confirme) return;
    }
    onClose();
  }

  async function sauvegarderStatut() {
    setSaving(true);
    const updates = { statut: statutEnCours, updated_at: new Date().toISOString() };
    if (statutEnCours === 'annulee') updates.raison_annulation = '';
    if (statutEnCours === 'absente') {
      await supabase.from('clients').update({ nb_absences: (c.nb_absences || 0) + 1 }).eq('id', resa.client_id);
    }
    const { error } = await supabase.from('reservations').update(updates).eq('id', resa.id);
    setSaving(false);
    if (error) { showToast('Erreur lors de la mise à jour', 'error'); return; }
    showToast('✅ Statut mis à jour');
    onSaved(statutEnCours);
    onClose();
  }

  function envoyerSms() {
    if (!c.tel || !smsTexte.trim()) return;
    window.location.href = `sms:${c.tel}?body=${encodeURIComponent(smsTexte)}`;
  }

  function containsEmoji(str) {
    return /[\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]|[\u{2700}-\u{27BF}]/u.test(str||'');
  }
  const smsLimit = containsEmoji(smsTexte) ? 70 : 160;
  const avatarBg = c.genre==='Homme'?'#dbeafe':c.genre==='Femme'?'#fce7f3':'#dcfce7';
  const avatarColor = c.genre==='Homme'?'#1d4ed8':c.genre==='Femme'?'#be185d':'#15803d';
  const initiales = c.genre==='Entreprise'
    ? (c.entreprise||'?').slice(0,2).toUpperCase()
    : `${(c.prenom||'?')[0]}${(c.nom||'')[0]||''}`.toUpperCase();

  return (
    <>
      {/* Overlay bloquant */}
      <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();}} onClick={fermerModal}
        style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:2999,pointerEvents:'all'}}/>

      {/* Modal */}
      <div onClick={e=>e.stopPropagation()} onMouseDown={e=>e.stopPropagation()}
        style={{position:'fixed',top:'50%',left:'50%',transform:'translate(-50%,-50%)',background:'#fff',borderRadius:20,width:'min(560px,calc(100vw - 48px))',maxHeight:'90vh',display:'flex',flexDirection:'column',boxShadow:'0 32px 80px rgba(0,0,0,0.25)',zIndex:3000,overflow:'hidden'}}>

        {/* Header */}
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'24px 28px 20px',flexShrink:0,borderBottom:'1px solid #f0f0f0'}}>
          <div>
            <h2 style={{margin:0,fontSize:20,fontWeight:800,color:'#111'}}>Détail de la réservation</h2>
            <p style={{margin:'4px 0 0',fontSize:13,color:'#999'}}>
              {new Date(resa.date+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long',year:'numeric'})}
            </p>
          </div>
          <button onClick={fermerModal} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666',display:'flex',alignItems:'center',justifyContent:'center'}}>✕</button>
        </div>

        {/* Contenu scrollable */}
        <div style={{flex:1,overflowY:'auto',padding:'20px 28px',display:'flex',flexDirection:'column',gap:16}}>

          {/* Bloc client */}
          <div style={{background:'#f9f9f9',borderRadius:14,padding:'16px 18px',display:'flex',alignItems:'center',gap:14}}>
            <div style={{width:48,height:48,borderRadius:'50%',flexShrink:0,background:avatarBg,display:'flex',alignItems:'center',justifyContent:'center',fontSize:16,fontWeight:800,color:avatarColor}}>
              {initiales}
            </div>
            <div style={{flex:1}}>
              <div style={{fontWeight:800,fontSize:16,color:'#111'}}>{nom||'—'}</div>
              {c.prenom && c.nom && c.entreprise && <div style={{fontSize:13,color:'#888',marginTop:1}}>{c.prenom} {c.nom}</div>}
              {c.tel && <a href={`tel:${c.tel}`} style={{fontSize:13,color:'#666',textDecoration:'none',display:'flex',alignItems:'center',gap:4,marginTop:3}}><Phone size={12} strokeWidth={2} color="#999"/> {c.tel}</a>}
              {c.mail && <div style={{fontSize:12,color:'#3b82f6',marginTop:2}}>{c.mail}</div>}
            </div>
            {c.tel && (
              <div style={{display:'flex',gap:8}}>
                <a href={`tel:${c.tel}`} style={{width:38,height:38,borderRadius:10,background:'#E8C547',border:'none',display:'flex',alignItems:'center',justifyContent:'center',textDecoration:'none',flexShrink:0}}>
                  <Phone size={16} strokeWidth={2} color="#111"/>
                </a>
                <button onClick={()=>{ setShowSmsPanel(!showSmsPanel); if(!showSmsPanel) setSmsTexte(smsSuggestions[0]); }} style={{width:38,height:38,borderRadius:10,background:'#fff',border:'1.5px solid #eee',display:'flex',alignItems:'center',justifyContent:'center',cursor:'pointer',flexShrink:0}}>
                  <MessageSquare size={16} strokeWidth={2} color="#666"/>
                </button>
              </div>
            )}
          </div>

          {/* Infos réservation */}
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
            {[
              {label:'Service', value: resa.service==='midi'?'☀️ Midi':'🌙 Soir'},
              {label:'Heure', value: resa.heure||'—'},
              {label:'Personnes', value: `${resa.nb_personnes} pers.`},
              {label:'Occasion', value: resa.occasion||'—'},
            ].map((item,i)=>(
              <div key={i} style={{background:'#f9f9f9',borderRadius:10,padding:'12px 14px'}}>
                <div style={{fontSize:10,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:0.5,marginBottom:4}}>{item.label}</div>
                <div style={{fontSize:14,fontWeight:600,color:'#111'}}>{item.value}</div>
              </div>
            ))}
            {resa.commentaire_client && (
              <div style={{gridColumn:'1/-1',background:'#f9f9f9',borderRadius:10,padding:'12px 14px'}}>
                <div style={{fontSize:10,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:0.5,marginBottom:4}}>Commentaire</div>
                <div style={{fontSize:14,color:'#555',fontStyle:'italic'}}>"{resa.commentaire_client}"</div>
              </div>
            )}
            {resa.raison_refus && (
              <div style={{gridColumn:'1/-1',background:'#fef2f2',borderRadius:10,padding:'12px 14px'}}>
                <div style={{fontSize:10,fontWeight:700,color:'#dc2626',textTransform:'uppercase',letterSpacing:0.5,marginBottom:4}}>Motif refus</div>
                <div style={{fontSize:14,color:'#dc2626'}}>{resa.raison_refus}</div>
              </div>
            )}
          </div>

          {/* Statut */}
          <div>
            <p style={{fontSize:13,fontWeight:700,color:'#111',margin:'0 0 8px'}}>Statut</p>
            <div style={{position:'relative'}}>
              {(()=>{
                const s = STATUTS_COLORS.find(x=>x.value===statutEnCours)||STATUTS_COLORS[0];
                return (
                  <button onClick={()=>setShowStatutPanel(!showStatutPanel)} style={{width:'100%',height:48,borderRadius:12,border:`2px solid ${s.color}`,background:`${s.color}18`,cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'space-between',padding:'0 16px',fontSize:15,fontWeight:700,color:s.color}}>
                    <span style={{display:'flex',alignItems:'center',gap:8}}><span style={{width:10,height:10,borderRadius:'50%',background:s.color,display:'inline-block'}}/>{s.label}{statutModifie&&<span style={{fontSize:10,opacity:0.8}}>●</span>}</span>
                    <ChevronDown size={16} strokeWidth={2} color={s.color}/>
                  </button>
                );
              })()}
              {showStatutPanel && (
                <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:5000,display:'flex',alignItems:'center',justifyContent:'center',pointerEvents:'all',cursor:'default',touchAction:'none'}} onMouseDown={e=>{e.preventDefault();e.stopPropagation();setShowStatutPanel(false);}} onClick={()=>setShowStatutPanel(false)}>
                  <div style={{background:'#fff',borderRadius:16,padding:24,width:320,boxShadow:'0 20px 60px rgba(0,0,0,0.2)'}} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()}>
                    <h3 style={{margin:'0 0 16px',fontSize:16,fontWeight:800}}>Changer le statut</h3>
                    {STATUTS_COLORS.map(s=>(
                      <div key={s.value} onClick={()=>{setStatutEnCours(s.value);setStatutModifie(s.value!==resa.statut);setShowStatutPanel(false);}}
                        style={{display:'flex',alignItems:'center',gap:12,padding:12,borderRadius:10,cursor:'pointer',marginBottom:6,background:statutEnCours===s.value?`${s.color}10`:'#fff'}}
                        onMouseEnter={e=>e.currentTarget.style.background='#f5f5f5'}
                        onMouseLeave={e=>e.currentTarget.style.background=statutEnCours===s.value?`${s.color}10`:'#fff'}>
                        <div style={{width:12,height:12,borderRadius:'50%',background:s.color,flexShrink:0}}/>
                        <div style={{flex:1}}>
                          <div style={{fontWeight:700,fontSize:14,color:statutEnCours===s.value?s.color:'#111'}}>{s.label}</div>
                          <div style={{fontSize:12,color:'#999'}}>{s.desc}</div>
                        </div>
                        {statutEnCours===s.value && <span style={{color:s.color,fontSize:18}}>✓</span>}
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Panneau SMS */}
          {showSmsPanel && c.tel && (
            <div style={{background:'#f9f9f9',borderRadius:12,padding:16,display:'flex',flexDirection:'column',gap:8}}>
              <p style={{fontSize:12,fontWeight:700,color:'#999',margin:'0 0 4px',textTransform:'uppercase',letterSpacing:0.5}}>Suggestions</p>
              <div style={{display:'flex',flexDirection:'column',gap:6}}>
                {smsSuggestions.map((s,i)=>(
                  <button key={i} onClick={()=>setSmsTexte(s.slice(0,smsLimit))} style={{width:'100%',textAlign:'left',background:smsTexte===s?'#E8C547':'#fff',border:'1.5px solid #eee',borderRadius:8,padding:'8px 12px',fontSize:12,cursor:'pointer',color:'#111',fontWeight:smsTexte===s?700:400}}>{s}</button>
                ))}
              </div>
              <textarea value={smsTexte} onChange={e=>setSmsTexte(e.target.value.slice(0,smsLimit))} placeholder="Votre message…"
                style={{width:'100%',height:70,border:'1.5px solid #eee',borderRadius:8,padding:'8px 12px',fontSize:13,resize:'none',outline:'none',fontFamily:'inherit',boxSizing:'border-box'}}/>
              <div style={{display:'flex',flexWrap:'wrap',gap:5}}>
                <span style={{fontSize:11,color:'#999',alignSelf:'center'}}>Insérer :</span>
                {[{label:'{prénom}',val:c.prenom||'{prénom}'},{label:'{nom}',val:c.nom||'{nom}'},{label:'🔗 Lien',val:'https://ted-crm.pages.dev/reserver.html'}].map(v=>(
                  <button key={v.label} onClick={()=>setSmsTexte((smsTexte+v.val).slice(0,smsLimit))} style={{background:'#fffbea',border:'1.5px solid #E8C547',borderRadius:6,padding:'3px 10px',fontSize:12,fontWeight:600,color:'#111',cursor:'pointer'}}>{v.label}</button>
                ))}
              </div>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <span style={{fontSize:11,color:smsTexte.length>smsLimit*0.9?'#dc2626':'#999',fontWeight:smsTexte.length>smsLimit*0.9?700:400}}>{smsTexte.length}/{smsLimit}{containsEmoji(smsTexte)&&' ⚠️ Emoji'}</span>
                <button onClick={envoyerSms} disabled={!smsTexte.trim()} style={{background:smsTexte.trim()?'#111':'#ddd',color:smsTexte.trim()?'#fff':'#999',border:'none',borderRadius:8,padding:'6px 16px',fontSize:13,fontWeight:800,cursor:smsTexte.trim()?'pointer':'not-allowed'}}>Envoyer</button>
              </div>
            </div>
          )}

          {/* Historique client */}
          <div style={{background:'#f9f9f9',borderRadius:12,padding:'12px 16px',display:'flex',gap:0}}>
            <div style={{textAlign:'center',flex:1}}>
              <div style={{fontSize:20,fontWeight:800,color:'#111'}}>{totalResas}</div>
              <div style={{fontSize:10,color:'#999',textTransform:'uppercase',letterSpacing:0.5}}>Résa total</div>
            </div>
            <div style={{textAlign:'center',flex:1}}>
              <div style={{fontSize:20,fontWeight:800,color:noshow>0?'#dc2626':'#111'}}>{noshow}</div>
              <div style={{fontSize:10,color:'#999',textTransform:'uppercase',letterSpacing:0.5}}>No-show</div>
            </div>
            <div style={{textAlign:'center',flex:2}}>
              <div style={{fontSize:13,fontWeight:700,color:'#111'}}>{derniereVisiteFormatee}</div>
              <div style={{fontSize:10,color:'#999',textTransform:'uppercase',letterSpacing:0.5}}>Dernière visite</div>
            </div>
          </div>
        </div>

        {/* Boutons fixes en bas */}
        <div style={{flexShrink:0,padding:'16px 28px',borderTop:'1px solid #eee',background:'#fff',display:'flex',gap:10}}>
          <button onClick={fermerModal} style={{flex:1,height:52,border:'1.5px solid #eee',borderRadius:12,background:'#fff',fontSize:15,fontWeight:600,cursor:'pointer',color:'#666'}}>Fermer</button>
          {onEdit && (
            <button onClick={()=>{onClose();onEdit(resa);}} style={{flex:1,height:52,border:'none',borderRadius:12,background:'#f0f0f0',fontSize:15,fontWeight:700,cursor:'pointer',color:'#111',display:'flex',alignItems:'center',justifyContent:'center',gap:6}}>✏️ Modifier</button>
          )}
          {statutModifie && (
            <button onClick={sauvegarderStatut} disabled={saving} style={{flex:2,height:52,border:'none',borderRadius:12,background:saving?'#ddd':'#E8C547',fontSize:15,fontWeight:800,cursor:saving?'not-allowed':'pointer',color:saving?'#999':'#111',display:'flex',alignItems:'center',justifyContent:'center',gap:8}}>
              <Check size={18} strokeWidth={2}/> {saving?'Enregistrement…':'Valider le statut'}
            </button>
          )}
        </div>
      </div>
    </>
  );
}

function ReservationsPage({ onBack, showToast, user, onLogout, inline = false, onResaCountChange }) {
  const [resaList, setResaList] = useState([]);
  const [loading, setLoading] = useState(true);
  const [refusResa, setRefusResa] = useState(null);
  const [acceptResa, setAcceptResa] = useState(null);
  const [detailResa, setDetailResa] = useState(null);
  const [editResa, setEditResa] = useState(null);
  const [showAddResa, setShowAddResa] = useState(false);
  const [ficheClientRP, setFicheClientRP] = useState(null);
  const [showConfirmDecoRP, setShowConfirmDecoRP] = useState(false);
  const [calDate, setCalDate] = useState(new Date());
  const [calDragX, setCalDragX] = useState(0);
  const [calIsDragging, setCalIsDragging] = useState(false);
  const [calNoTransition, setCalNoTransition] = useState(false);
  const [calDragDir, setCalDragDir] = useState(null);
  const calContainerRef = useRef(null);
  const calSwipeTouchStartX = useRef(null);
  const calTouchStartY = useRef(null);
  const [calSlideDir, setCalSlideDir] = useState(null);
  const [calAnimating, setCalAnimating] = useState(false);
  const [calMensuelOuvert, setCalMensuelOuvert] = useState(false);
  const now0 = new Date();
  const todayLocal = `${now0.getFullYear()}-${String(now0.getMonth()+1).padStart(2,'0')}-${String(now0.getDate()).padStart(2,'0')}`;
  const [calJourSelectionne, setCalJourSelectionne] = useState(todayLocal);
  const joursScrollRef = useRef(null);
  const [calServiceSelectionne, setCalServiceSelectionne] = useState(new Date().getHours() < 15 ? 'midi' : 'soir');
  const [resaSearchPanel, setResaSearchPanel] = useState('');
const [showDemandesAttente, setShowDemandesAttente] = useState(false);
  const [showFormDropdown, setShowFormDropdown] = useState(false);
  const isMobile = useIsMobile();
  const qr = `https://api.qrserver.com/v1/create-qr-code/?size=120x120&data=${encodeURIComponent(FORM_URL)}`;

  useEffect(() => {
    function handleClickOutside(e) {
      if (!e.target.closest('#formulaire-dropdown')) setShowFormDropdown(false);
    }
    if (showFormDropdown) document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [showFormDropdown]);

  useEffect(() => { loadResa(); }, []);

  // Sync calendrier quand le jour sélectionné change de mois
  useEffect(() => {
    if (!calJourSelectionne) return;
    const d = new Date(calJourSelectionne + 'T12:00:00');
    if (d.getFullYear() !== calDate.getFullYear() || d.getMonth() !== calDate.getMonth()) {
      const dir = d > calDate ? 1 : -1;
      setCalDate(new Date(d.getFullYear(), d.getMonth(), 1));
      setCalSlideDir(dir > 0 ? 'right' : 'left');
      setTimeout(() => setCalSlideDir(null), 300);
    }
  }, [calJourSelectionne]);

  useEffect(() => {
    const ch = supabase
      .channel('resa-page-realtime')
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'reservations' }, async (payload) => {
        const nouvelleResa = payload.new;
        const { data: clientData } = await supabase.from('clients').select('*').eq('id', nouvelleResa.client_id).single();
        const resaComplete = { ...nouvelleResa, clients: clientData };
        setResaList(prev => {
          if (prev.some(r => r.id === resaComplete.id)) return prev; // déjà chargée via loadResa, on ignore
          const updated = [resaComplete, ...prev];
          onResaCountChange?.(updated.filter(r => r.statut === 'attente').length);
          return updated;
        });
      })
      .on('postgres_changes', { event: 'UPDATE', schema: 'public', table: 'reservations' }, (payload) => {
        setResaList(prev => {
          const updated = prev.map(r => r.id === payload.new.id ? { ...r, ...payload.new } : r);
          onResaCountChange?.(updated.filter(r => r.statut === 'attente').length);
          return updated;
        });
      })
      .subscribe();
    return () => { supabase.removeChannel(ch); };
  }, []);

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
    // Recharge les infos fraîches du client (mail/nom peuvent avoir changé)
    const { data: clientFrais } = await supabase.from('clients').select('*').eq('id', r.client_id).single();
    const clientPourEmail = clientFrais || r.clients;
    if (!clientPourEmail?.mail) { showToast("⚠️ Email non envoyé (pas d'adresse)"); return; }
    const dateFormatee = new Date(r.date).toLocaleDateString('fr-FR', {weekday:'long', day:'numeric', month:'long', year:'numeric'});
    const dateStr = r.date.replace(/-/g,'');
    const heureArr = (r.heure||'19:00').split(':');
    const hStart = heureArr[0];
    const mStart = heureArr[1];
    const hEnd = String(parseInt(hStart) + 2).padStart(2,'0');
    const calStart = `${dateStr}T${hStart}${mStart}00`;
    const calEnd = `${dateStr}T${hEnd}${mStart}00`;
    const titre = encodeURIComponent('Réservation Le TED');
    const lieu = encodeURIComponent('28 Av. des Frères Montgolfier, 69680 Chassieu');
    const details = encodeURIComponent(`Réservation confirmée au TED pour ${r.nb_personnes} personne(s) — ${r.service === 'midi' ? 'Déjeuner' : 'Dîner'}`);
    const agendaUrl = `https://ted-crm.pages.dev/agenda.html?date=${r.date}&heure=${encodeURIComponent(r.heure||'19:00')}&nb=${r.nb_personnes}&service=${r.service}&prenom=${encodeURIComponent(clientPourEmail?.prenom||'')}`;
    const htmlConfirmation = `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto;background:#f8f8f8;padding:20px">
  <div style="background:#111111;padding:28px 24px;text-align:center;border-radius:12px 12px 0 0;border-bottom:4px solid #E8C547">
    <img src="https://ted-crm.pages.dev/favicon.png" alt="Le TED" style="height:60px;margin-bottom:12px" />
    <h1 style="color:#E8C547;margin:0;font-size:28px;letter-spacing:2px;font-weight:800">LE TED</h1>
    <p style="color:#888;margin:4px 0 0;font-size:13px;letter-spacing:1px">RESTAURANT &amp; CLUB — CHASSIEU</p>
  </div>
  <div style="background:#fff;padding:28px 24px;border-radius:0 0 12px 12px;box-shadow:0 4px 20px rgba(0,0,0,0.08)">
    <h2 style="color:#111;margin:0 0 8px;font-size:22px">Bonjour ${clientPourEmail.prenom} 👋</h2>
    <p style="color:#444;font-size:16px;margin:0 0 24px">Votre réservation est <strong style="color:#16a34a">confirmée</strong> ✅</p>
    <div style="background:#f9f9f9;border-left:4px solid #E8C547;padding:20px;border-radius:0 8px 8px 0;margin-bottom:24px">
      <p style="margin:0 0 10px;font-size:15px">📅 <strong>Date :</strong> ${dateFormatee}</p>
      <p style="margin:0 0 10px;font-size:15px">🕐 <strong>Heure :</strong> ${r.heure || 'À confirmer'}</p>
      <p style="margin:0 0 10px;font-size:15px">👥 <strong>Nombre de personnes :</strong> ${r.nb_personnes}</p>
      <p style="margin:0;font-size:15px">🍽 <strong>Service :</strong> ${r.service === 'midi' ? 'Déjeuner' : 'Dîner'}</p>
      ${r.occasion ? `<p style="margin:8px 0 0;font-size:15px">🎉 <strong>Occasion :</strong> ${r.occasion}</p>` : ''}
    </div>
    <div style="text-align:center;margin-bottom:24px">
      <a href="${agendaUrl}" target="_blank" style="display:inline-block;background:#E8C547;color:#111;text-decoration:none;padding:14px 32px;border-radius:10px;font-weight:800;font-size:15px">📅 Ajouter à mon agenda</a>
    </div>
    <div style="background:#f9f9f9;border:1.5px solid #ddd;border-radius:8px;padding:16px;margin-bottom:20px">
      <p style="margin:0 0 10px;font-size:14px;font-weight:800;color:#111">👔 Dress code</p>
      <p style="margin:0 0 8px;font-size:13px;color:#555;line-height:1.6">Afin de garantir une ambiance soignée à tous nos clients, nous vous remercions de respecter notre dress code :</p>
      <ul style="margin:0;padding-left:18px;font-size:13px;color:#555;line-height:2">
        <li>Pas de pulls à capuches ni de joggings ou pantalons style cargo</li>
        <li>Pas de couvre-chef, quel qu'il soit</li>
        <li>Pas de baskets type Air Max, TN ou similaires</li>
      </ul>
      <p style="margin:8px 0 0;font-size:12px;color:#999;font-style:italic">Merci de votre compréhension — nous nous réservons le droit de refuser l'accès en cas de non-respect.</p>
    </div>
    <div style="background:#fff8e1;border:1.5px solid #E8C547;border-radius:8px;padding:16px;margin-bottom:24px">
      <p style="margin:0;font-size:14px;color:#555;line-height:1.6">⚠️ <strong>En cas d'annulation ou de modification</strong>, merci de nous prévenir au plus tôt au <strong>04 78 90 67 80</strong> ou par email afin que nous puissions libérer la table pour d'autres clients. Merci de votre compréhension.</p>
    </div>
    <div style="border-top:1px solid #eee;padding-top:20px;text-align:center">
      <p style="color:#111;font-weight:700;font-size:15px;margin:0 0 6px">Le TED — Restaurant &amp; Club</p>
      <p style="color:#888;font-size:13px;margin:0 0 4px">📍 28 Av. des Frères Montgolfier, 69680 Chassieu</p>
      <p style="color:#888;font-size:13px;margin:0 0 4px">📞 04 78 90 67 80</p>
      <p style="margin:8px 0 0;text-align:center"><a href="https://leted.fr" style="display:inline-flex;align-items:center;justify-content:center;gap:8px;text-decoration:none;color:#111;font-size:15px;font-weight:700"><img src="https://ted-crm.pages.dev/favicon.png" alt="TED" style="height:24px;width:24px;vertical-align:middle" />leted.fr</a></p>
    </div>
    <p style="text-align:center;color:#bbb;font-size:12px;margin-top:20px">Nous avons hâte de vous accueillir ! 🎉</p>
  </div>
</div>`;
    const resEmail = await sendBrevoEmail(
      clientPourEmail.mail,
      `${clientPourEmail.prenom || ''} ${clientPourEmail.nom || ''}`.trim(),
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
      <p style="margin:8px 0 0;text-align:center"><a href="https://leted.fr" style="display:inline-flex;align-items:center;justify-content:center;gap:8px;text-decoration:none;color:#111;font-size:15px;font-weight:700"><img src="https://ted-crm.pages.dev/favicon.png" alt="TED" style="height:24px;width:24px;vertical-align:middle" />leted.fr</a></p>
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

  const cardStyle = { background:'#fff', borderRadius:14, border:'1.5px solid #f0f0f0', padding:16, marginBottom:10, boxShadow:'0 2px 8px rgba(0,0,0,0.04)' };

  function telechargerTableau(date, service, reservations) {
    const dateFormatee = new Date(date + 'T12:00:00').toLocaleDateString('fr-FR', {weekday:'long', day:'numeric', month:'long', year:'numeric'});
    const serviceLabel2 = service === 'midi' ? '☀️ Déjeuner' : '🌙 Dîner';
    const resasConfirmees = reservations.filter(r => r.statut !== 'annulee');
    const lignes = resasConfirmees.map((r) => {
      const nom = r.clients?.genre === 'Entreprise' ? (r.clients?.entreprise || '') : `${r.clients?.prenom || ''} ${r.clients?.nom || ''}`;
      return `<tr><td>${nom}</td><td style="text-align:center">${r.heure || ''}</td><td style="text-align:center">${r.nb_personnes || ''}</td><td></td><td>${r.commentaire_client || ''}</td><td></td></tr>`;
    }).join('');
    const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><title>Réservations TED - ${dateFormatee}</title><style>* { margin:0; padding:0; box-sizing:border-box; } body { font-family: Arial, sans-serif; background: #fff; padding: 40px; } .header { text-align: center; margin-bottom: 32px; border-bottom: 3px solid #E8C547; padding-bottom: 20px; } .logo { font-size: 32px; font-weight: 900; letter-spacing: 4px; color: #111; } .subtitle { font-size: 13px; color: #888; letter-spacing: 2px; margin-top: 4px; text-transform: uppercase; } .date-title { font-size: 20px; font-weight: 700; color: #111; margin-top: 16px; } .service-badge { display: inline-block; background: #E8C547; color: #111; padding: 4px 16px; border-radius: 20px; font-size: 13px; font-weight: 700; margin-top: 8px; } table { width: 100%; border-collapse: collapse; margin-top: 24px; } th { background: #111; color: #E8C547; padding: 12px 16px; text-align: left; font-size: 12px; letter-spacing: 1px; text-transform: uppercase; } td { padding: 12px 16px; border-bottom: 1px solid #eee; font-size: 14px; color: #333; } tr:last-child td { border-bottom: 2px solid #111; } .footer { margin-top: 32px; text-align: center; font-size: 11px; color: #bbb; } @media print { body { padding: 20px; } }</style></head><body><div class="header"><div style="display:flex;align-items:center;justify-content:center;gap:12px;margin-bottom:8px"><img src="https://leted.fr/wp-content/uploads/2023/01/logo-Le-TED.png" style="height:60px;width:auto" onerror="this.src='https://ted-crm.pages.dev/favicon.png'" /><div class="logo">LE TED</div></div><div class="subtitle">Restaurant &amp; Club — Chassieu</div><div class="date-title">${dateFormatee}</div><div class="service-badge">${serviceLabel2}</div></div><table><thead><tr><th>Nom Prénom</th><th style="text-align:center">Heure</th><th style="text-align:center">Couverts</th><th style="text-align:center">N° Table</th><th>Commentaire</th><th style="text-align:center">Validé</th></tr></thead><tbody>${lignes}${(() => { const n = resasConfirmees.length; const nbTotal = Math.max(20, Math.ceil(n / 4) * 4); const nb = nbTotal - n; return Array(nb).fill('<tr><td style="padding:14px 16px;border-bottom:1px solid #eee">&nbsp;</td><td></td><td></td><td></td><td></td><td></td></tr>').join(''); })()}</tbody></table><div class="footer">Imprimé le ${new Date().toLocaleDateString('fr-FR')} · ${resasConfirmees.length} réservation(s) — Le TED · 28 Av. des Frères Montgolfier, 69680 Chassieu · 04 78 90 67 80</div></body></html>`;
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `reservations-ted-${date}-${service}.html`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  return (
    <div style={{ fontFamily:"'Inter','Segoe UI',Arial,sans-serif", background:'#f8f8f8', minHeight: inline ? undefined : '100vh', overflow: !isMobile ? 'hidden' : undefined, height: (!isMobile && inline) ? '100vh' : undefined }}>
      {/* Header — desktop full-page mode only */}
      {!inline && (
        <header style={{ background:'#111', color:'#fff', padding:'0 20px', height:56, display:'flex', alignItems:'center', justifyContent:'space-between', borderBottom:`3px solid ${G}`, flexShrink:0 }}>
          <div style={{ display:'flex', alignItems:'center', gap:10 }}>
            <span style={{ fontWeight:700, fontSize:15, color:'#fff' }}>📅 <span style={{ color:G }}>TED</span> — Réservations</span>
            <div id="formulaire-dropdown" style={{ position:'relative' }}>
              <button onClick={() => setShowFormDropdown(v => !v)} style={{ background:'rgba(255,255,255,0.08)', border:'1px solid #444', borderRadius:8, height:34, padding:'0 14px', color:'#ccc', fontWeight:600, fontSize:13, cursor:'pointer' }}>🔗 Formulaire</button>
              {showFormDropdown && (
                <div style={{ position:'absolute', top:40, left:0, background:'#fff', borderRadius:10, border:'1.5px solid #eee', boxShadow:'0 8px 24px rgba(0,0,0,0.12)', padding:8, zIndex:200, minWidth:180 }}>
                  <button type="button" onMouseDown={async(e)=>{ e.preventDefault(); e.stopPropagation(); try{ await navigator.clipboard.writeText('https://ted-crm.pages.dev/reserver.html'); }catch{ const t=document.createElement('textarea'); t.value='https://ted-crm.pages.dev/reserver.html'; document.body.appendChild(t); t.select(); document.execCommand('copy'); document.body.removeChild(t); } showToast('✅ Lien copié !'); setShowFormDropdown(false); }} style={{ display:'block', width:'100%', textAlign:'left', padding:'9px 14px', border:'none', background:'none', cursor:'pointer', fontSize:13, fontWeight:600, borderRadius:7 }}>📋 Copier</button>
                  <button type="button" onMouseDown={(e)=>{ e.preventDefault(); e.stopPropagation(); window.open('https://ted-crm.pages.dev/reserver.html','_blank'); setShowFormDropdown(false); }} style={{ display:'block', width:'100%', textAlign:'left', padding:'9px 14px', border:'none', background:'none', cursor:'pointer', fontSize:13, fontWeight:600, borderRadius:7 }}>🔗 Ouvrir</button>
                  <button type="button" onMouseDown={async(e)=>{ e.preventDefault(); e.stopPropagation(); const url='https://ted-crm.pages.dev/reserver.html'; if(navigator.share){ try{ await navigator.share({title:'Réservation Le TED',url}); }catch{} }else{ try{ await navigator.clipboard.writeText(url); }catch{} showToast('✅ Lien copié !'); } setShowFormDropdown(false); }} style={{ display:'block', width:'100%', textAlign:'left', padding:'9px 14px', border:'none', background:'none', cursor:'pointer', fontSize:13, fontWeight:600, borderRadius:7 }}>📤 Partager</button>
                </div>
              )}
            </div>
          </div>
          <div style={{ display:'flex', alignItems:'center', gap:8 }}>
            <button onClick={onBack} style={{ background:'rgba(255,255,255,0.08)', border:'1px solid #444', borderRadius:8, height:34, padding:'0 14px', color:'#ccc', fontWeight:600, fontSize:13, cursor:'pointer' }}>👥 Mes Clients</button>
            <button onClick={()=>onBack('communications')} style={{ background:'rgba(255,255,255,0.08)', border:'1px solid #444', borderRadius:8, height:34, padding:'0 14px', color:'#ccc', fontWeight:600, fontSize:13, cursor:'pointer' }}>📣 Communications</button>
            <button onClick={()=>setShowConfirmDecoRP(true)} style={{ background:'transparent', color:'#ccc', border:'1px solid #444', borderRadius:7, padding:'0 12px', height:32, fontSize:12, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}>🔓 Déconnexion</button>
          </div>
        </header>
      )}

      <div style={{ display: !isMobile ? 'grid' : 'block', gridTemplateColumns: !isMobile ? '1fr 380px' : undefined, gap: !isMobile ? 16 : undefined, padding: !isMobile ? '24px 32px' : undefined, maxWidth: !isMobile ? 1440 : undefined, margin: !isMobile ? '0 auto' : undefined, alignItems: !isMobile ? 'stretch' : 'start', height: !isMobile ? 'calc(100vh - 48px)' : undefined, boxSizing: !isMobile ? 'border-box' : undefined, background: !isMobile ? '#f5f5f5' : undefined }}>
      <main style={{ maxWidth: isMobile ? 800 : 'none', margin: isMobile ? '0 auto' : 0, padding: isMobile ? '12px 16px 100px' : '0', display: !isMobile ? 'flex' : 'block', flexDirection: !isMobile ? 'column' : undefined, gap: !isMobile ? 12 : undefined, height: !isMobile ? '100%' : undefined, overflow: !isMobile ? 'hidden' : undefined }}>

        {!isMobile && (
          <div style={{ display:'flex', alignItems:'center', gap:12, flexShrink:0, position:'relative' }}>
            <h1 style={{ fontSize:28, fontWeight:900, color:'#111', margin:0 }}>Réservations</h1>
            <div style={{ position:'relative' }}>
              <button onClick={()=>setShowFormDropdown(v=>!v)} style={{ display:'flex', alignItems:'center', gap:6, height:38, padding:'0 14px', background:'#fff', border:'1.5px solid #eee', borderRadius:10, fontSize:13, fontWeight:600, cursor:'pointer', color:'#666' }}>
                <Link size={14} strokeWidth={2} /> Formulaire
              </button>
              {showFormDropdown && (
                <>
                  <div onClick={()=>setShowFormDropdown(false)} style={{ position:'fixed', inset:0, zIndex:299 }} />
                  <div style={{ position:'absolute', top:'calc(100% + 8px)', left:0, background:'#fff', borderRadius:10, boxShadow:'0 8px 32px rgba(0,0,0,0.15)', padding:6, minWidth:200, zIndex:300 }}>
                    <button type="button" onMouseDown={async(e)=>{ e.preventDefault(); e.stopPropagation(); try{ await navigator.clipboard.writeText('https://ted-crm.pages.dev/reserver.html'); }catch{ const t=document.createElement('textarea'); t.value='https://ted-crm.pages.dev/reserver.html'; document.body.appendChild(t); t.select(); document.execCommand('copy'); document.body.removeChild(t); } showToast('✅ Lien copié !'); setShowFormDropdown(false); }} style={{ width:'100%', padding:'10px 14px', border:'none', background:'none', textAlign:'left', cursor:'pointer', fontSize:13, borderRadius:6, display:'flex', alignItems:'center', gap:10, color:'#111' }} onMouseEnter={e=>e.currentTarget.style.background='#f5f5f5'} onMouseLeave={e=>e.currentTarget.style.background='none'}><Copy size={15} strokeWidth={2} color="#666" /> Copier le lien</button>
                    <button type="button" onMouseDown={(e)=>{ e.preventDefault(); e.stopPropagation(); window.open('https://ted-crm.pages.dev/reserver.html','_blank'); setShowFormDropdown(false); }} style={{ width:'100%', padding:'10px 14px', border:'none', background:'none', textAlign:'left', cursor:'pointer', fontSize:13, borderRadius:6, display:'flex', alignItems:'center', gap:10, color:'#111' }} onMouseEnter={e=>e.currentTarget.style.background='#f5f5f5'} onMouseLeave={e=>e.currentTarget.style.background='none'}><ExternalLink size={15} strokeWidth={2} color="#666" /> Ouvrir</button>
                    <button type="button" onMouseDown={async(e)=>{ e.preventDefault(); e.stopPropagation(); const url='https://ted-crm.pages.dev/reserver.html'; if(navigator.share){ try{ await navigator.share({title:'Réservation Le TED',url}); }catch{} }else{ try{ await navigator.clipboard.writeText(url); }catch{} showToast('✅ Lien copié !'); } setShowFormDropdown(false); }} style={{ width:'100%', padding:'10px 14px', border:'none', background:'none', textAlign:'left', cursor:'pointer', fontSize:13, borderRadius:6, display:'flex', alignItems:'center', gap:10, color:'#111' }} onMouseEnter={e=>e.currentTarget.style.background='#f5f5f5'} onMouseLeave={e=>e.currentTarget.style.background='none'}><Share2 size={15} strokeWidth={2} color="#666" /> Partager</button>
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {/* ── Bouton Demandes en attente ── */}
        {(() => {
          const nbAttente = resaList.filter(r => r.statut === 'attente').length;
          return (
            <div onClick={()=>setShowDemandesAttente(true)} className={nbAttente > 0 ? 'alarm-blink' : ''} style={{ background: nbAttente > 0 ? '#dc2626' : '#fff', border: nbAttente > 0 ? 'none' : '1.5px solid #f0f0f0', borderRadius:16, padding:'14px 20px', display:'flex', alignItems:'center', justifyContent:'space-between', cursor:'pointer', flexShrink:0, transition:'background 0.1s', boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
              <span style={{ fontSize:15, fontWeight:800, color: nbAttente > 0 ? '#fff' : '#111', display:'flex', alignItems:'center', gap:8 }}><ClipboardList size={16} strokeWidth={2} color={nbAttente > 0 ? '#fff' : '#666'} /> Demandes de réservation en attente</span>
              <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                {nbAttente > 0 ? (
                  <span style={{ background:'#fff', color:'#dc2626', borderRadius:'50%', width:26, height:26, display:'flex', alignItems:'center', justifyContent:'center', fontSize:13, fontWeight:800 }}>{nbAttente}</span>
                ) : (
                  <span style={{ fontSize:13, color:'#999', fontWeight:600 }}>Aucune</span>
                )}
                <span style={{ color: nbAttente > 0 ? '#fff' : '#ccc', fontSize:18 }}>›</span>
              </div>
            </div>
          );
        })()}

        {/* ── Bloc unique : 7 jours + calendrier + Midi/Soir ── */}
        {(() => {
          const nowLocal = new Date();
          const todayStr = `${nowLocal.getFullYear()}-${String(nowLocal.getMonth()+1).padStart(2,'0')}-${String(nowLocal.getDate()).padStart(2,'0')}`;
          const quinzeJours = Array.from({length:15}, (_,i) => {
            const d = new Date(); d.setDate(d.getDate() + i);
            const str = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
            return { date:str, jour:d.toLocaleDateString('fr-FR',{weekday:'short'}).toUpperCase().replace('.',''), num:d.getDate(), mois:d.toLocaleDateString('fr-FR',{month:'short'}), isAujourd: str===todayStr };
          });
          const JOURS = ['Lun','Mar','Mer','Jeu','Ven','Sam','Dim'];
          const MOIS = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
          const annee = calDate.getFullYear();
          const mois = calDate.getMonth();
          const premierJour = new Date(annee, mois, 1);
          const dernierJour = new Date(annee, mois + 1, 0);
          const debutSemaine = (premierJour.getDay() + 6) % 7;
          const confirmeesParJour = {};
          resaList.filter(r => r.statut === 'confirmee').forEach(r => {
            if (!confirmeesParJour[r.date]) confirmeesParJour[r.date] = [];
            confirmeesParJour[r.date].push(r);
          });
          const cases = [];
          for (let i = 0; i < debutSemaine; i++) cases.push(null);
          for (let d = 1; d <= dernierJour.getDate(); d++) cases.push(d);
          while (cases.length % 7 !== 0) cases.push(null);
          const today = new Date();
          const dateLabel = calJourSelectionne ? new Date(calJourSelectionne + 'T12:00:00').toLocaleDateString('fr-FR', { weekday:'long', day:'numeric', month:'long', year:'numeric' }) : null;
          const couvertsMidi = calJourSelectionne ? resaList.filter(r => r.date === calJourSelectionne && r.service === 'midi' && r.statut === 'confirmee').reduce((sum, r) => sum + (r.nb_personnes || 0), 0) : 0;
          const couvertsSoir = calJourSelectionne ? resaList.filter(r => r.date === calJourSelectionne && r.service === 'soir' && r.statut === 'confirmee').reduce((sum, r) => sum + (r.nb_personnes || 0), 0) : 0;
          return (
            <div style={{ background:'#fff', borderRadius:16, border:'1.5px solid #f0f0f0', padding:14, flex:1, display:'flex', flexDirection:'column', overflow:'hidden', boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
              {/* 1. 15 jours scroll natif */}
              <div ref={joursScrollRef} className="jours-strip"
                style={{ display:'flex', gap:8, overflowX:'scroll', marginBottom:16, flexShrink:0, WebkitOverflowScrolling:'touch', scrollSnapType:'x mandatory', userSelect:'none', WebkitUserSelect:'none' }}>
                {quinzeJours.map(j => {
                  const totalCouverts = resaList.filter(r => r.date===j.date && r.statut==='confirmee').reduce((sum,r)=>sum+(r.nb_personnes||0),0);
                  const isSelected = calJourSelectionne === j.date;
                  return (
                    <div key={j.date} onClick={()=>setCalJourSelectionne(j.date)}
                      style={{ borderRadius:12, padding:'10px 6px', textAlign:'center', cursor:'pointer', border:'2px solid', borderColor: isSelected?'#E8C547':'#eee', background: isSelected?'#fffbea':'#fff', transition:'border-color 0.15s, background 0.15s', flexShrink:0, width:'calc((100% - 40px) / 6)', scrollSnapAlign:'start' }}>
                      <div style={{ fontSize:10, fontWeight:700, marginBottom:4, color: isSelected?'#E8C547': j.isAujourd?'#E8C547':'#999' }}>{j.isAujourd?'AUJ.':j.jour}</div>
                      <div style={{ fontSize:20, fontWeight:900, marginBottom:2, color:'#111' }}>{j.num}</div>
                      <div style={{ fontSize:10, color:'#999', marginBottom:4 }}>{j.mois}</div>
                      <div style={{ fontSize:11, fontWeight:700, color: totalCouverts>0?'#111':'#ccc' }}>{totalCouverts>0?`${totalCouverts} p.`:'—'}</div>
                      {j.isAujourd && <div style={{ width:5, height:5, borderRadius:'50%', background:'#E8C547', margin:'4px auto 0' }}/>}
                    </div>
                  );
                })}
              </div>
              <div style={{ display: !isMobile ? 'grid' : 'block', gridTemplateColumns: !isMobile ? '3fr 2fr' : undefined, gap: !isMobile ? 16 : 0, marginTop: !isMobile ? 16 : 0, flex: !isMobile ? 1 : undefined, minHeight: !isMobile ? 0 : undefined, overflow: !isMobile ? 'hidden' : undefined }}>
                {/* Colonne calendrier */}
                <div style={!isMobile ? { background:'#f8f8f8', borderRadius:12, padding:12, overflow:'auto' } : {}}>
                  {/* 2. Bouton toggle (mobile only) */}
                  {isMobile && (
                    <button onClick={()=>setCalMensuelOuvert(v=>!v)} style={{ width:'100%', padding:'10px 12px', background:'#f8f8f8', border:'1.5px solid #eee', borderRadius:8, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom: calMensuelOuvert ? 12 : 0 }}>
                      <span style={{ fontSize:13, fontWeight:700, color:'#555' }}>📅 Calendrier</span>
                      <span style={{ fontSize:12, color:'#999' }}>{calMensuelOuvert ? '▲' : '▼'}</span>
                    </button>
                  )}
                  {/* 3. Grand calendrier mensuel — toujours visible sur desktop */}
                  {(!isMobile || calMensuelOuvert) && (() => {
                    const changerMois = (direction) => {
                      if (calAnimating) return;
                      setCalAnimating(true);
                      setCalDragX(0);
                      // Change le mois immédiatement → le nouveau mois entre en animation
                      setCalDate(new Date(annee, mois + direction, 1));
                      // direction > 0 = mois suivant → entre depuis la droite
                      setCalSlideDir(direction > 0 ? 'right' : 'left');
                      setTimeout(() => { setCalSlideDir(null); setCalAnimating(false); }, 300);
                    };
                    // Mois adjacent pour le swipe (visible pendant le drag)
                    const adjDate = calDragDir === 'left' ? new Date(annee, mois+1, 1)
                                  : calDragDir === 'right' ? new Date(annee, mois-1, 1) : null;
                    let adjCases = [];
                    if (adjDate) {
                      const aA = adjDate.getFullYear(), aM = adjDate.getMonth();
                      const aPremier = new Date(aA, aM, 1), aDernier = new Date(aA, aM+1, 0);
                      const aDebut = (aPremier.getDay() + 6) % 7;
                      for (let i=0; i<aDebut; i++) adjCases.push(null);
                      for (let d=1; d<=aDernier.getDate(); d++) adjCases.push(d);
                      while (adjCases.length % 7 !== 0) adjCases.push(null);
                    }
                    const containerW = calContainerRef.current?.offsetWidth || 320;
                    const calTransition = (calIsDragging || calNoTransition) ? 'none' : 'transform 0.28s cubic-bezier(0.4,0,0.2,1)';
                    const renderCalGrid = (gridCases, gridAnnee, gridMois, dx) => (
                      <div style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', gap:3, transform:`translateX(${dx}px)`, transition:calTransition, willChange:'transform', touchAction:'pan-y', position:'absolute', top:0, left:0, width:'100%' }}>
                        {gridCases.map((d, i) => {
                          if (!d) return <div key={i} />;
                          const iso = `${gridAnnee}-${String(gridMois+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
                          const hasResa = !!confirmeesParJour[iso];
                          const isToday2 = today.getFullYear()===gridAnnee && today.getMonth()===gridMois && today.getDate()===d;
                          const isSelected2 = calJourSelectionne === iso;
                          const estPasse2 = new Date(iso) < new Date(new Date().setHours(0,0,0,0));
                          return (
                            <button key={i} onClick={() => setCalJourSelectionne(iso)}
                              style={{ textAlign:'center', height:48, borderRadius:6, cursor:'pointer', position:'relative',
                                border: isToday2 && !isSelected2 ? '2px solid #E8C547' : '2px solid transparent',
                                background: isSelected2 ? '#111' : isToday2 ? '#fffbea' : 'transparent',
                                color: isSelected2 ? '#fff' : '#111',
                                fontWeight: isSelected2 ? 800 : isToday2 ? 900 : 400, fontSize:16,
                                boxSizing:'border-box', opacity: estPasse2 ? 0.4 : 1, transition:'background 0.15s' }}>
                              {d}
                              {hasResa && <span style={{ display:'block', width:4, height:4, borderRadius:'50%', background:'#E8C547', margin:'2px auto 0' }} />}
                            </button>
                          );
                        })}
                      </div>
                    );
                    const handleCalTouchStart = (e) => {
                      if (calAnimating) return;
                      calSwipeTouchStartX.current = e.touches[0].clientX;
                      calTouchStartY.current = e.touches[0].clientY;
                      setCalDragDir(null);
                      setCalIsDragging(false);
                      setCalDragX(0);
                    };
                    const handleCalTouchMove = (e) => {
                      if (calSwipeTouchStartX.current === null) return;
                      const dx = e.touches[0].clientX - calSwipeTouchStartX.current;
                      const dy = e.touches[0].clientY - calTouchStartY.current;
                      if (!calIsDragging && Math.abs(dy) > Math.abs(dx)) return;
                      e.preventDefault();
                      if (!calDragDir && Math.abs(dx) > 8) setCalDragDir(dx < 0 ? 'left' : 'right');
                      setCalIsDragging(true);
                      setCalDragX(dx);
                    };
                    const handleCalTouchEnd = (e) => {
                      if (calSwipeTouchStartX.current === null) return;
                      const dx = e.changedTouches[0].clientX - calSwipeTouchStartX.current;
                      const dy = e.changedTouches[0].clientY - calTouchStartY.current;
                      setCalIsDragging(false);
                      if (Math.abs(dy) <= Math.abs(dx) && Math.abs(dx) > containerW * 0.28) {
                        const dir = dx < 0 ? 1 : -1;
                        const target = dx < 0 ? -containerW : containerW;
                        setCalDragX(target); // anime la sortie
                        setTimeout(() => {
                          setCalNoTransition(true);
                          setCalDate(new Date(annee, mois+dir, 1));
                          setCalDragDir(null);
                          setCalDragX(0);
                          setTimeout(() => setCalNoTransition(false), 30);
                        }, 280);
                      } else {
                        setCalDragX(0);
                        setCalDragDir(null);
                      }
                      calSwipeTouchStartX.current = null;
                    };
                    return (
                    <div style={{ marginBottom:4, userSelect:'none', WebkitUserSelect:'none' }}>
                      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:10 }}>
                        <button onClick={()=>changerMois(-1)} style={{ background:'#f0f0f0', border:'none', borderRadius:8, width:34, height:34, fontSize:16, cursor:'pointer', fontWeight:700 }}>‹</button>
                        <span style={{ fontWeight:800, fontSize:18 }}>{MOIS[mois]} {annee}</span>
                        <button onClick={()=>changerMois(1)} style={{ background:'#f0f0f0', border:'none', borderRadius:8, width:34, height:34, fontSize:16, cursor:'pointer', fontWeight:700 }}>›</button>
                      </div>
                      <div style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', gap:2, marginBottom:4 }}>
                        {JOURS.map(j => <div key={j} style={{ textAlign:'center', fontSize:13, fontWeight:700, color:'#999', padding:'8px 0' }}>{j}</div>)}
                      </div>
                      <div ref={calContainerRef}
                        onTouchStart={handleCalTouchStart}
                        onTouchMove={handleCalTouchMove}
                        onTouchEnd={handleCalTouchEnd}
                        style={{ overflow:'hidden', position:'relative', height: `${Math.ceil(cases.length/7)*51}px` }}
                      >
                        {renderCalGrid(cases, annee, mois, calDragX)}
                        {adjDate && renderCalGrid(adjCases, adjDate.getFullYear(), adjDate.getMonth(), calDragX + (calDragDir==='left' ? containerW : -containerW))}
                      </div>
                    </div>
                    );
                  })()}
                </div>
                {/* Colonne Midi/Soir */}
                <div style={!isMobile ? { background:'#f8f8f8', borderRadius:12 } : {}}>
                  {!isMobile ? (
                    <div style={{ display:'flex', flexDirection:'column', gap:16, padding:16 }}>
                      <p style={{ fontSize:12, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:0 }}>SERVICE</p>
                      <p style={{ fontSize:14, color:'#666', margin:0 }}>{dateLabel || 'Sélectionner un jour'}</p>
                      <button onClick={() => setCalServiceSelectionne('midi')} style={{ width:'100%', padding:'20px 16px', borderRadius:12, border: calServiceSelectionne==='midi' ? '2px solid #111' : '1.5px solid #eee', background: calServiceSelectionne==='midi' ? '#111' : '#fff', cursor:'pointer', textAlign:'center' }}>
                        <div style={{ fontSize:22 }}>☀️</div>
                        <div style={{ fontSize:16, fontWeight:800, color: calServiceSelectionne==='midi' ? '#E8C547' : '#111', marginTop:6 }}>Midi</div>
                        <div style={{ fontSize:13, color:'#999', marginTop:4 }}>{couvertsMidi} couvert{couvertsMidi > 1 ? 's' : ''}</div>
                      </button>
                      <button onClick={() => setCalServiceSelectionne('soir')} style={{ width:'100%', padding:'20px 16px', borderRadius:12, border: calServiceSelectionne==='soir' ? '2px solid #111' : '1.5px solid #eee', background: calServiceSelectionne==='soir' ? '#111' : '#fff', cursor:'pointer', textAlign:'center' }}>
                        <div style={{ fontSize:22 }}>🌙</div>
                        <div style={{ fontSize:16, fontWeight:800, color: calServiceSelectionne==='soir' ? '#E8C547' : '#111', marginTop:6 }}>Soir</div>
                        <div style={{ fontSize:13, color:'#999', marginTop:4 }}>{couvertsSoir} couvert{couvertsSoir > 1 ? 's' : ''}</div>
                      </button>
                    </div>
                  ) : (
                    /* Mobile Midi/Soir */
                    calJourSelectionne && (
                      <div style={{ borderTop:'1px solid #f0f0f0', paddingTop:12, marginTop:4 }}>
                        <div style={{ fontSize:12, fontWeight:700, color:'#555', marginBottom:10 }}>{dateLabel}</div>
                        <div style={{ display:'flex', gap:8 }}>
                          <div style={{ flex:1, display:'flex', flexDirection:'column', gap:4 }}>
                            <button onClick={() => setCalServiceSelectionne(calServiceSelectionne === 'midi' ? null : 'midi')}
                              style={{ height:40, borderRadius:9, border:'1.5px solid', fontSize:13, fontWeight:700, cursor:'pointer',
                                background: calServiceSelectionne === 'midi' ? '#111' : '#fff',
                                color: calServiceSelectionne === 'midi' ? '#E8C547' : '#111',
                                borderColor: calServiceSelectionne === 'midi' ? '#111' : '#ddd' }}>☀️ Midi</button>
                            <div style={{ textAlign:'center', fontSize:11, color:'#888' }}>{couvertsMidi} couvert{couvertsMidi > 1 ? 's' : ''}</div>
                          </div>
                          <div style={{ flex:1, display:'flex', flexDirection:'column', gap:4 }}>
                            <button onClick={() => setCalServiceSelectionne(calServiceSelectionne === 'soir' ? null : 'soir')}
                              style={{ height:40, borderRadius:9, border:'1.5px solid', fontSize:13, fontWeight:700, cursor:'pointer',
                                background: calServiceSelectionne === 'soir' ? '#111' : '#fff',
                                color: calServiceSelectionne === 'soir' ? '#E8C547' : '#111',
                                borderColor: calServiceSelectionne === 'soir' ? '#111' : '#ddd' }}>🌙 Soir</button>
                            <div style={{ textAlign:'center', fontSize:11, color:'#888' }}>{couvertsSoir} couvert{couvertsSoir > 1 ? 's' : ''}</div>
                          </div>
                        </div>
                      </div>
                    )
                  )}
                </div>
              </div>
            </div>
          );
        })()}

        {/* ── Tableau réservations du jour (mobile only — desktop uses right panel) ── */}
        {isMobile && calJourSelectionne && calServiceSelectionne && (() => {
          const resasDuJour = resaList
            .filter(r => (r.statut === 'confirmee' || r.statut === 'annulee' || r.statut === 'absente') && r.date === calJourSelectionne && r.service === calServiceSelectionne)
            .sort((a,b) => (a.heure||'').localeCompare(b.heure||''));
          const dateLabel = new Date(calJourSelectionne + 'T12:00:00').toLocaleDateString('fr-FR', { weekday:'long', day:'numeric', month:'long', year:'numeric' });
          const serviceLabel = calServiceSelectionne === 'midi' ? '☀️ Midi' : '🌙 Soir';
          return (
            <div style={{ background:'#fff', borderRadius:14, padding:'14px 16px', marginBottom:12, boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16 }}>
                <div>
                  <div style={{ fontWeight:800, fontSize:17 }}>Réservations TED</div>
                  <div style={{ fontSize:13, color:'#888', marginTop:2 }}>{dateLabel} — {serviceLabel}</div>
                </div>
                <button onClick={() => telechargerTableau(calJourSelectionne, calServiceSelectionne, resasDuJour)} style={{ background:'#111', color:'#fff', border:'none', borderRadius:9, padding:'0 18px', height:38, fontSize:13, fontWeight:700, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}>
                  📥 Télécharger
                </button>
              </div>
              <div id="print-tableau">
                <div style={{ textAlign:'center', marginBottom:16, display:'none' }} className="print-only">
                  <div style={{ fontWeight:800, fontSize:22 }}>Réservations TED</div>
                  <div style={{ fontSize:15, color:'#555', marginTop:4 }}>{dateLabel} — {serviceLabel}</div>
                </div>
                {resasDuJour.length === 0 ? (
                  <div style={{ textAlign:'center', padding:'24px 0', color:'#bbb', fontSize:14 }}>Aucune réservation confirmée pour ce service</div>
                ) : isMobile ? (
                  <div style={{padding:'0 16px 16px'}}>
                    <div style={{display:'grid', gridTemplateColumns:'2fr 1fr 1fr auto', gap:8, padding:'8px 0', borderBottom:'2px solid #E8C547', marginBottom:4}}>
                      <span style={{fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1}}>Nom Prénom</span>
                      <span style={{fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1}}>Heure</span>
                      <span style={{fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1}}>Couverts</span>
                      <span/>
                    </div>
                    {resasDuJour.map((r,ri) => {
                      const sMobile = ({confirmee:{bg:'#dcfce7',color:'#16a34a',label:'Confirmée'},attente:{bg:'#fef9c3',color:'#ca8a04',label:'En attente'},venue:{bg:'#d1fae5',color:'#059669',label:'Venue'},absente:{bg:'#fee2e2',color:'#dc2626',label:'No-show'},annulee:{bg:'#f3f4f6',color:'#6b7280',label:'Annulée'}})[r.statut]||{bg:'#f3f4f6',color:'#666',label:r.statut};
                      return (
                      <div key={r.id} onClick={() => setDetailResa(r)} style={{display:'flex', alignItems:'center', gap:12, padding:'10px 0', borderBottom: ri<resasDuJour.length-1?'1px solid #f5f5f5':'none', cursor:'pointer'}}>
                        <span style={{fontSize:13, fontWeight:800, color:'#111', minWidth:40, flexShrink:0}}>{r.heure||'—'}</span>
                        <div style={{flex:1, minWidth:0}}>
                          <div style={{fontSize:14, fontWeight:700, color: r.statut==='absente'?'#dc2626':'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}}>
                            {r.clients?.genre==='Entreprise' ? r.clients?.entreprise : `${r.clients?.prenom||''} ${r.clients?.nom||''}`}
                          </div>
                          <div style={{fontSize:12, color:'#999'}}>{r.nb_personnes} pers.</div>
                        </div>
                        <span style={{background:sMobile.bg, color:sMobile.color, borderRadius:20, padding:'3px 8px', fontSize:11, fontWeight:700, flexShrink:0}}>{sMobile.label}</span>
                      </div>
                      );
                    })}
                  </div>
                ) : (
                  <table style={{ width:'100%', borderCollapse:'collapse', fontSize:13 }}>
                    <tbody>
                      {resasDuJour.map((r) => (
                        <tr key={r.id} onClick={() => setDetailResa(r)} style={{ borderBottom:'1px solid #f0f0f0', cursor:'pointer', background: r.statut==='absente' ? '#fff0f0' : r.statut==='annulee' ? '#fff5f5' : 'white', opacity: r.statut==='annulee' ? 0.8 : 1 }}
                          onMouseEnter={e => e.currentTarget.style.background= r.statut==='absente' ? '#ffe0e0' : r.statut==='annulee' ? '#ffe8e8' : '#fffbea'}
                          onMouseLeave={e => e.currentTarget.style.background= r.statut==='absente' ? '#fff0f0' : r.statut==='annulee' ? '#fff5f5' : ''}
                        >
                          <td style={{ padding:'12px 16px' }}>
                            <div style={{ display:'flex', alignItems:'center', gap:16, flexWrap:'nowrap' }}>
                              <span style={{ fontWeight:700, minWidth:140, color: r.statut==='absente' ? '#dc2626' : r.clients?.genre==='Entreprise' ? '#E8C547' : '#111', display:'flex', alignItems:'center', gap:6 }}>
                                {r.clients?.genre==='Entreprise' ? r.clients?.entreprise : `${r.clients?.prenom||''} ${r.clients?.nom||''}`}
                                {r.statut==='annulee' && <span style={{background:'#f97316', color:'#fff', fontSize:10, fontWeight:700, borderRadius:4, padding:'2px 6px', textTransform:'uppercase'}}>Annulée</span>}
                                {r.statut==='absente' && <span style={{background:'#dc2626', color:'#fff', fontSize:10, fontWeight:700, borderRadius:4, padding:'2px 6px', textTransform:'uppercase'}}>Absente</span>}
                              </span>
                              <span style={{ color:'#666', minWidth:50 }}>{r.heure || '—'}</span>
                              <span style={{ color:'#666', minWidth:60 }}>{r.nb_personnes} pers.</span>
                              {r.clients?.tel && (
                                <a href={`tel:${r.clients.tel}`} onClick={e => e.stopPropagation()} style={{ display:'inline-flex', alignItems:'center', gap:4, background:'#f0f0f0', borderRadius:6, padding:'3px 10px', fontSize:12, color:'#111', textDecoration:'none', fontWeight:600, whiteSpace:'nowrap' }}>📞 {r.clients.tel}</a>
                              )}
                              {r.commentaire_client && (
                                <span style={{ fontSize:12, color:'#999', fontStyle:'italic' }}>{r.commentaire_client}</span>
                              )}
                              <span style={{ color:'#ccc', fontSize:16, marginLeft:'auto' }}>›</span>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          );
        })()}

        {/* ── Modal Demandes en attente ── */}
        {showDemandesAttente && (
          <>
            <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();}} onClick={()=>setShowDemandesAttente(false)} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:2999,pointerEvents:'all'}}/>
            <div onClick={e=>e.stopPropagation()} onMouseDown={e=>e.stopPropagation()} style={{position:'fixed',top:'50%',left:'50%',transform:'translate(-50%,-50%)',background:'#fff',borderRadius:20,width:'min(620px, calc(100vw - 48px))',maxHeight:'85vh',display:'flex',flexDirection:'column',boxShadow:'0 32px 80px rgba(0,0,0,0.25)',zIndex:3000,overflow:'hidden'}}>
              {/* Header */}
              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'24px 28px 20px',flexShrink:0,borderBottom:'1px solid #f0f0f0'}}>
                <div>
                  <h2 style={{margin:0,fontSize:20,fontWeight:800,color:'#111'}}>Demandes en attente</h2>
                  <p style={{margin:'4px 0 0',fontSize:13,color:'#999'}}>{attente.length} demande{attente.length>1?'s':''} à traiter</p>
                </div>
                <button onClick={()=>setShowDemandesAttente(false)} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666',display:'flex',alignItems:'center',justifyContent:'center'}}>✕</button>
              </div>
              {/* Liste scrollable */}
              <div style={{flex:1,overflowY:'auto',padding:'16px 28px'}}>
                {attente.length === 0 && (
                  <div style={{textAlign:'center',padding:'48px 0',color:'#bbb'}}>
                    <div style={{fontSize:48,marginBottom:12}}>✓</div>
                    <p style={{fontSize:15,fontWeight:600,margin:0}}>Aucune demande en attente</p>
                  </div>
                )}
                {attente.sort((a,b)=>a.date.localeCompare(b.date)).map(r=>{
                  const cl = r.clients || {};
                  const avatarBg = cl.genre==='Homme'?'#dbeafe':cl.genre==='Femme'?'#fce7f3':'#dcfce7';
                  const avatarColor = cl.genre==='Homme'?'#1d4ed8':cl.genre==='Femme'?'#be185d':'#15803d';
                  const initiales = cl.genre==='Entreprise'?(cl.entreprise||'?').slice(0,2).toUpperCase():`${(cl.prenom||'?')[0]}${(cl.nom||'')[0]||''}`.toUpperCase();
                  return (
                    <div key={r.id} style={{background:'#f9f9f9',borderRadius:14,padding:16,marginBottom:12,border:'1.5px solid #f0f0f0'}}>
                      <div style={{display:'flex',alignItems:'center',gap:12,marginBottom:14}}>
                        <div style={{width:44,height:44,borderRadius:'50%',flexShrink:0,background:avatarBg,display:'flex',alignItems:'center',justifyContent:'center',fontSize:14,fontWeight:800,color:avatarColor}}>{initiales}</div>
                        <div style={{flex:1}}>
                          <div style={{fontWeight:800,fontSize:15,color:'#111'}}>{cl.genre==='Entreprise'?cl.entreprise:`${cl.prenom||''} ${cl.nom||''}`}</div>
                          <div style={{fontSize:13,color:'#999',display:'flex',gap:8,marginTop:2,flexWrap:'wrap'}}>
                            <span>{cl.tel}</span>
                            {cl.mail && <><span>·</span><span>{cl.mail}</span></>}
                          </div>
                        </div>
                        {cl.tel && <a href={`tel:${cl.tel}`} onClick={e=>e.stopPropagation()} style={{width:36,height:36,borderRadius:'50%',background:'#fff',border:'1.5px solid #eee',display:'flex',alignItems:'center',justifyContent:'center',textDecoration:'none',flexShrink:0}}><Phone size={16} strokeWidth={2} color="#666"/></a>}
                      </div>
                      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr 1fr',gap:8,marginBottom:14}}>
                        {[
                          {label:'Date', value:new Date(r.date+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long'})},
                          {label:'Service', value:r.service==='midi'?'☀️ Midi':'🌙 Soir'},
                          {label:'Heure', value:r.heure||'—'},
                          {label:'Personnes', value:`${r.nb_personnes} pers.`},
                          {label:'Occasion', value:r.occasion||'—'},
                        ].map((item,i)=>(
                          <div key={i} style={{background:'#fff',borderRadius:8,padding:'8px 12px'}}>
                            <div style={{fontSize:10,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:0.5,marginBottom:3}}>{item.label}</div>
                            <div style={{fontSize:13,fontWeight:700,color:'#111'}}>{item.value}</div>
                          </div>
                        ))}
                      </div>
                      {r.commentaire_client && <div style={{background:'#fffbea',borderRadius:8,padding:'8px 12px',marginBottom:14,fontSize:13,color:'#666',fontStyle:'italic'}}>💬 {r.commentaire_client}</div>}
                      <div style={{display:'flex',gap:10}}>
                        <button onClick={()=>setRefusResa(r)} style={{flex:1,height:44,border:'none',borderRadius:10,background:'#dc2626',fontSize:14,fontWeight:700,cursor:'pointer',color:'#fff'}}>✕ Refuser</button>
                        <button onClick={()=>setAcceptResa(r)} style={{flex:2,height:44,border:'none',borderRadius:10,background:'#16a34a',fontSize:14,fontWeight:800,cursor:'pointer',color:'#fff'}}>✓ Accepter la réservation</button>
                      </div>
                    </div>
                  );
                })}
              </div>
              {/* Fermer fixe en bas */}
              <div style={{flexShrink:0,padding:'16px 28px',borderTop:'1px solid #eee'}}>
                <button onClick={()=>setShowDemandesAttente(false)} style={{width:'100%',height:48,border:'1.5px solid #eee',borderRadius:12,background:'#fff',fontSize:14,fontWeight:600,cursor:'pointer',color:'#666'}}>Fermer</button>
              </div>
            </div>
          </>
        )}

      </main>

      {/* Right column — desktop only */}
      {!isMobile && (() => {
        const resasDuJour = (calJourSelectionne && calServiceSelectionne)
          ? resaList.filter(r => (r.statut==='confirmee'||r.statut==='annulee'||r.statut==='absente') && r.date===calJourSelectionne && r.service===calServiceSelectionne).sort((a,b)=>(a.heure||'').localeCompare(b.heure||''))
          : [];
        const resasDuJourFiltrees = resaSearchPanel
          ? resasDuJour.filter(r => { const n = `${r.clients?.prenom||''} ${r.clients?.nom||''} ${r.clients?.entreprise||''} ${r.clients?.tel||''}`.toLowerCase(); return n.includes(resaSearchPanel.toLowerCase()); })
          : resasDuJour;
        return (
          <div style={{ background:'#fff', borderRadius:16, border:'1.5px solid #f0f0f0', height:'100%', display:'flex', flexDirection:'column', overflow:'hidden', boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
            {/* Header fixe */}
            <div style={{padding:'16px 20px 12px', flexShrink:0, borderBottom:'1px solid #f5f5f5'}}>
              <div style={{display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:8}}>
                <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:0}}>
                  Réservations du
                </p>
                {calJourSelectionne && calServiceSelectionne && (
                  <button onClick={()=>telechargerTableau(calJourSelectionne, calServiceSelectionne, resasDuJour.filter(r=>r.statut==='confirmee'))} style={{
                    height:28, padding:'0 12px', borderRadius:8,
                    border:'1.5px solid #eee', background:'#fff',
                    fontSize:11, fontWeight:600, cursor:'pointer', color:'#666',
                    display:'flex', alignItems:'center', gap:5
                  }}
                  onMouseEnter={e=>{e.currentTarget.style.background='#f5f5f5';}}
                  onMouseLeave={e=>{e.currentTarget.style.background='#fff';}}
                  >
                    <Download size={12} strokeWidth={2} color="#666"/> Télécharger
                  </button>
                )}
              </div>
              <h3 style={{margin:'0 0 8px', fontSize:16, fontWeight:800, color:'#111'}}>
                {calJourSelectionne ? new Date(calJourSelectionne+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long'}) : 'Sélectionner un jour'}
                {calServiceSelectionne ? ` — ${calServiceSelectionne==='midi'?'☀️ Midi':'🌙 Soir'}` : ''}
              </h3>
              {calJourSelectionne && calServiceSelectionne && (
                <div style={{display:'flex', gap:16, fontSize:12, color:'#666', marginBottom:10}}>
                  <span style={{display:'flex', alignItems:'center', gap:4}}>
                    <Users size={12} strokeWidth={2} color="#999"/>
                    {resasDuJour.filter(r=>r.statut==='confirmee').length} réservation{resasDuJour.filter(r=>r.statut==='confirmee').length>1?'s':''}
                  </span>
                  <span style={{display:'flex', alignItems:'center', gap:4}}>
                    <UtensilsCrossed size={12} strokeWidth={2} color="#999"/>
                    {resasDuJour.filter(r=>r.statut==='confirmee').reduce((s,r)=>s+(r.nb_personnes||0),0)} couverts
                  </span>
                </div>
              )}
              <div style={{position:'relative'}}>
                <Search size={13} strokeWidth={2} color="#999" style={{position:'absolute', left:10, top:'50%', transform:'translateY(-50%)', pointerEvents:'none'}}/>
                <input value={resaSearchPanel} onChange={e=>setResaSearchPanel(e.target.value)} placeholder="Rechercher une réservation..." style={{
                  width:'100%', height:34, border:'1.5px solid #eee',
                  borderRadius:9, padding:'0 10px 0 30px',
                  fontSize:12, outline:'none', boxSizing:'border-box'
                }}/>
              </div>
            </div>
            {/* Liste scrollable */}
            <div style={{ flex:1, overflowY:'auto' }}>
              {resasDuJourFiltrees.map(r => {
                const statutColors = {
                  'confirmee': {bg:'#dcfce7', color:'#16a34a', label:'Confirmée'},
                  'attente':   {bg:'#fef9c3', color:'#ca8a04', label:'En attente'},
                  'venue':     {bg:'#d1fae5', color:'#059669', label:'Venue'},
                  'absente':   {bg:'#fee2e2', color:'#dc2626', label:'No-show'},
                  'annulee':   {bg:'#f3f4f6', color:'#6b7280', label:'Annulée'},
                };
                const s = statutColors[r.statut] || statutColors['confirmee'];
                return (
                  <div key={r.id} onClick={()=>setDetailResa(r)} style={{ display:'flex', alignItems:'center', gap:12, padding:'12px 20px', borderBottom:'1px solid #f5f5f5', cursor:'pointer' }}
                    onMouseEnter={e=>e.currentTarget.style.background='#fafafa'}
                    onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                    <span style={{ fontSize:14, fontWeight:800, color:'#111', minWidth:44, flexShrink:0 }}>{r.heure||'—'}</span>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ fontWeight:700, fontSize:14, color: r.statut==='absente'?'#dc2626':'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                        {r.clients?.genre==='Entreprise' ? r.clients?.entreprise : `${r.clients?.prenom||''} ${r.clients?.nom||''}`.trim()}
                      </div>
                      <div style={{ fontSize:12, color:'#999' }}>{r.nb_personnes} pers.</div>
                    </div>
                    <span style={{ background:s.bg, color:s.color, borderRadius:20, padding:'3px 10px', fontSize:11, fontWeight:700, flexShrink:0, whiteSpace:'nowrap' }}>{s.label}</span>
                  </div>
                );
              })}
              {resasDuJour.length === 0 && (
                <div style={{ padding:'48px', textAlign:'center', color:'#bbb' }}>
                  <CalendarDays size={32} strokeWidth={1.5} color="#ddd" style={{ marginBottom:12 }} />
                  <p style={{ fontSize:14, margin:0 }}>{calJourSelectionne && calServiceSelectionne ? 'Aucune réservation confirmée' : 'Sélectionner un jour et un service'}</p>
                </div>
              )}
            </div>
            {/* Bouton fixe en bas */}
            <div style={{ flexShrink:0, padding:'14px 20px', borderTop:'1px solid #f5f5f5' }}>
              <button onClick={()=>setShowAddResa(true)} style={{ width:'100%', height:52, background:'#E8C547', border:'none', borderRadius:14, fontSize:15, fontWeight:800, cursor:'pointer', color:'#111', display:'flex', alignItems:'center', justifyContent:'center', gap:8, boxShadow:'0 2px 8px rgba(232,197,71,0.3)' }}>
                <Plus size={18} strokeWidth={2} /> Nouvelle réservation
              </button>
            </div>
          </div>
        );
      })()}

      </div>{/* end 2-col grid */}

      {acceptResa && <AccepterModal resa={acceptResa} onConfirm={()=>accepter(acceptResa)} onCancel={()=>setAcceptResa(null)} />}
      {refusResa && <RefusModal onConfirm={raison=>refuser(refusResa, raison)} onCancel={()=>setRefusResa(null)} />}
      {detailResa && <DetailResaModal resa={detailResa} resaList={resaList} showToast={showToast} onClose={()=>setDetailResa(null)} onEdit={(r)=>setEditResa(r)} onSaved={(newStatut)=>{ setResaList(prev => prev.map(r => r.id === detailResa.id ? {...r, statut: newStatut} : r)); setDetailResa(null); loadResa(); }} />}
      {showAddResa && <AddResaModal onClose={()=>setShowAddResa(false)} onSaved={()=>{ loadResa(); setShowAddResa(false); }} showToast={showToast} user={user} onViewClient={(c)=>setFicheClientRP(c)} reservations={resaList} />}
      {editResa && <AddResaModal initialResa={editResa} onClose={()=>setEditResa(null)} onSaved={()=>{ loadResa(); setEditResa(null); }} showToast={showToast} user={user} onViewClient={(c)=>setFicheClientRP(c)} reservations={resaList} />}
      {ficheClientRP && (() => {
        const c = ficheClientRP;
        const resasC = resaList.filter(r => r.client_id === c.id);
        const aujourd = new Date().toISOString().split('T')[0];
        const total = resasC.filter(r=>r.statut!=='annulee'&&r.statut!=='absente').length;
        const noshow = resasC.filter(r => r.statut === 'absente').length;
        const derniereVisite = resasC.filter(r => (r.statut==='venue'||r.statut==='confirmee') && r.date <= aujourd).sort((a,b)=>b.date.localeCompare(a.date))[0];
        const prochaineResa = resasC.filter(r => r.date >= aujourd && (r.statut==='confirmee'||r.statut==='attente')).sort((a,b)=>a.date.localeCompare(b.date))[0];
        const nomAffiche = c.genre==='Entreprise' ? (c.entreprise||c.nom||'—') : `${c.prenom||''} ${c.nom||''}`.trim()||'—';
        return (
          <>
            <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();}} onClick={()=>setFicheClientRP(null)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:4999, pointerEvents:'all' }}/>
            <div onClick={e=>e.stopPropagation()} onMouseDown={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:20, width:'min(420px, calc(100vw - 48px))', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.25)', zIndex:5000, overflow:'hidden' }}>
              {/* Header */}
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'22px 26px 18px', flexShrink:0, borderBottom:'1px solid #f0f0f0' }}>
                <h2 style={{margin:0, fontSize:18, fontWeight:800, color:'#111'}}>{nomAffiche}</h2>
                <button onClick={()=>setFicheClientRP(null)} style={{ width:34, height:34, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:16, color:'#666', display:'flex', alignItems:'center', justifyContent:'center' }}>✕</button>
              </div>
              {/* Contenu */}
              <div style={{padding:'18px 26px 22px', display:'flex', flexDirection:'column', gap:14}}>
                {/* Téléphone */}
                {c.tel && <a href={`tel:${c.tel}`} style={{ display:'flex', alignItems:'center', gap:10, background:'#E8C547', borderRadius:10, padding:'12px 16px', textDecoration:'none', color:'#111', fontWeight:700, fontSize:15 }}><Phone size={16} strokeWidth={2}/> {c.tel}</a>}
                {/* Email */}
                {c.mail && <div style={{display:'flex', alignItems:'center', gap:10}}><Mail size={15} strokeWidth={2} color="#3b82f6"/><span style={{fontSize:13, color:'#3b82f6'}}>{c.mail}</span></div>}
                {/* Stats */}
                <div style={{display:'grid', gridTemplateColumns:'1fr 1fr', gap:10}}>
                  <div style={{background:'#f9f9f9', borderRadius:10, padding:'12px 14px', textAlign:'center'}}>
                    <p style={{fontSize:22, fontWeight:900, color:'#111', margin:'0 0 3px'}}>{total}</p>
                    <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:0}}>Résa total</p>
                  </div>
                  <div style={{background:'#fef2f2', borderRadius:10, padding:'12px 14px', textAlign:'center'}}>
                    <p style={{fontSize:22, fontWeight:900, color:'#dc2626', margin:'0 0 3px'}}>{noshow}</p>
                    <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:0}}>No-show</p>
                  </div>
                </div>
                {/* Dernière visite */}
                <div style={{display:'flex', alignItems:'center', gap:10, padding:'10px 12px', background:'#f9f9f9', borderRadius:10}}>
                  <Clock size={15} strokeWidth={2} color="#666" style={{flexShrink:0}}/>
                  <div>
                    <span style={{fontSize:11, color:'#999'}}>Dernière visite : </span>
                    <span style={{fontSize:13, fontWeight:600, color:'#111'}}>{derniereVisite ? new Date(derniereVisite.date+'T12:00:00').toLocaleDateString('fr-FR',{day:'numeric',month:'long',year:'numeric'}) : 'Jamais'}</span>
                  </div>
                </div>
                {/* Bouton fermer */}
                <button onClick={()=>setFicheClientRP(null)} style={{ width:'100%', height:48, border:'1.5px solid #eee', borderRadius:12, background:'#fff', fontSize:14, fontWeight:600, cursor:'pointer', color:'#666', marginTop:4 }}>Fermer</button>
              </div>
            </div>
          </>
        );
      })()}
      {showConfirmDecoRP && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:9000, display:'flex', alignItems:'center', justifyContent:'center', pointerEvents:'all', cursor:'default', touchAction:'none' }} onMouseDown={e=>{e.preventDefault();e.stopPropagation();setShowConfirmDecoRP(false);}} onClick={(e)=>{if(e.target===e.currentTarget)setShowConfirmDecoRP(false);}}>
          <div style={{ background:'#fff', borderRadius:16, padding:'28px 24px', maxWidth:320, width:'90%', textAlign:'center' }} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()}>
            <h3 style={{ margin:'0 0 8px', fontSize:17, fontWeight:800 }}>Se déconnecter ?</h3>
            <p style={{ margin:'0 0 20px', fontSize:14, color:'#666' }}>Vous devrez vous reconnecter pour accéder au CRM.</p>
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={()=>setShowConfirmDecoRP(false)} style={{ flex:1, height:44, border:'1.5px solid #ddd', borderRadius:10, background:'#fff', fontSize:14, cursor:'pointer', color:'#666' }}>Annuler</button>
              <button onClick={()=>{ supabase.auth.signOut(); setShowConfirmDecoRP(false); }} style={{ flex:1, height:44, border:'none', borderRadius:10, background:'#111', fontSize:14, fontWeight:800, cursor:'pointer', color:'#fff' }}>Se déconnecter</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Sending Progress Modal ───────────────────────────────────────────────────

function SendingProgressModal({ type, total, done, successCount, onClose }) {
  const [progress, setProgress] = useState(0);
  const [showSuccess, setShowSuccess] = useState(false);

  useEffect(() => {
    const steps = 200;
    const target = 88;
    let step = 0;
    const timer = setInterval(() => {
      step++;
      const eased = target * (1 - Math.pow(1 - step / steps, 2.5));
      setProgress(Math.min(eased, target));
      if (step >= steps) clearInterval(timer);
    }, 50);
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    if (!done) return;
    setProgress(100);
    const t = setTimeout(() => {
      setShowSuccess(true);
      setTimeout(onClose, 2200);
    }, 500);
    return () => clearTimeout(t);
  }, [done]);

  const isEmail = type === 'email';
  const emoji = isEmail ? '✉️' : '📱';
  const label = isEmail ? 'email' : 'SMS';
  const labelP = isEmail ? 'emails' : 'SMS';

  return (
    <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)', zIndex:9999, display:'flex', alignItems:'center', justifyContent:'center', backdropFilter:'blur(6px)', WebkitBackdropFilter:'blur(6px)' }}>
      <div style={{ background:'#fff', borderRadius:28, padding:'44px 52px', width:'min(460px,calc(100vw - 48px))', textAlign:'center', boxShadow:'0 40px 100px rgba(0,0,0,0.25)', position:'relative', overflow:'hidden' }}>
        {/* Fond décoratif */}
        <div style={{ position:'absolute', top:-60, right:-60, width:200, height:200, borderRadius:'50%', background:'rgba(232,197,71,0.08)', pointerEvents:'none' }}/>
        <div style={{ position:'absolute', bottom:-40, left:-40, width:140, height:140, borderRadius:'50%', background:'rgba(232,197,71,0.05)', pointerEvents:'none' }}/>

        {!showSuccess ? (
          <>
            <div style={{ fontSize:44, marginBottom:18, filter:'drop-shadow(0 4px 8px rgba(0,0,0,0.12))' }}>{emoji}</div>
            <h2 style={{ margin:'0 0 6px', fontSize:22, fontWeight:900, color:'#111', letterSpacing:-0.5 }}>Envoi en cours…</h2>
            <p style={{ margin:'0 0 32px', fontSize:14, color:'#999', fontWeight:500 }}>
              {total} destinataire{total > 1 ? 's' : ''}
            </p>

            {/* Barre principale */}
            <div style={{ background:'#f0f0f0', borderRadius:99, height:12, overflow:'hidden', marginBottom:10, position:'relative' }}>
              <div style={{ height:'100%', borderRadius:99, background:'linear-gradient(90deg, #E8C547 0%, #f5d76e 60%, #ffe680 100%)', width:`${progress}%`, transition:'width 0.08s linear', boxShadow:'0 0 16px rgba(232,197,71,0.6)', position:'relative' }}>
                {/* Shimmer */}
                <div style={{ position:'absolute', top:0, right:0, bottom:0, width:40, background:'linear-gradient(90deg, transparent, rgba(255,255,255,0.4), transparent)', animation:'shimmer 1.2s infinite' }}/>
              </div>
            </div>

            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
              <p style={{ fontSize:12, color:'#bbb', margin:0, fontWeight:500 }}>
                {isEmail ? 'Connexion au serveur d\'envoi…' : 'Transmission vers les opérateurs…'}
              </p>
              <p style={{ fontSize:13, fontWeight:800, color:'#E8C547', margin:0 }}>{Math.round(progress)}%</p>
            </div>
          </>
        ) : (
          <>
            <div style={{ width:72, height:72, borderRadius:'50%', background:'linear-gradient(135deg, #22c55e, #16a34a)', display:'flex', alignItems:'center', justifyContent:'center', margin:'0 auto 20px', boxShadow:'0 8px 24px rgba(34,197,94,0.35)', fontSize:32 }}>✓</div>
            <h2 style={{ margin:'0 0 8px', fontSize:24, fontWeight:900, color:'#111', letterSpacing:-0.5 }}>
              {successCount} {successCount > 1 ? labelP : label} envoyé{successCount > 1 ? 's' : ''} !
            </h2>
            <p style={{ margin:0, fontSize:14, color:'#999' }}>Livraison en cours chez les destinataires</p>
          </>
        )}
      </div>
    </div>
  );
}

// ─── Menu Components ──────────────────────────────────────────────────────────

const MENU_BADGES = ['Fait maison', 'Fumé maison', 'Nouveau', 'Signature du chef', 'Best-seller'];
const MENU_ALLERGENES = ['Gluten','Crustacés','Œufs','Poisson','Arachides','Soja','Lait','Fruits à coque','Céleri','Moutarde','Graines de sésame','Anhydride sulfureux','Lupin','Mollusques'];

function formatPrix(p) {
  if (!p) return '';
  const s = String(p).trim();
  if (!s || s.includes('€')) return s;
  return /^[\d]+([,.][\d]{1,2})?$/.test(s) ? s + ' €' : s;
}

function slugify(str) {
  return str.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
}

function MenuToggle({ value, onChange }) {
  return (
    <div onClick={onChange} style={{ width:44, height:24, borderRadius:12, background: value ? '#E8C547' : '#ddd', cursor:'pointer', position:'relative', transition:'background 0.2s', flexShrink:0 }}>
      <div style={{ position:'absolute', top:2, left: value ? 22 : 2, width:20, height:20, borderRadius:10, background:'#fff', transition:'left 0.2s', boxShadow:'0 1px 4px rgba(0,0,0,0.2)' }} />
    </div>
  );
}

function MenuBottomSheet({ title, onClose, children, footer }) {
  const isMobile = window.innerWidth < 768;
  const [vis, setVis] = useState(false);
  useEffect(() => { const id = requestAnimationFrame(() => setVis(true)); return () => cancelAnimationFrame(id); }, []);
  function close() { setVis(false); setTimeout(onClose, 260); }

  const overlayBase = { position:'fixed', inset:0, background: vis ? 'rgba(0,0,0,0.45)' : 'rgba(0,0,0,0)', transition:'background 0.25s', zIndex:3000 };
  const headerBar = (
    <div style={{ padding:'10px 20px 14px', borderBottom:'1px solid #f0f0f0', display:'flex', alignItems:'center', justifyContent:'space-between', flexShrink:0 }}>
      <span style={{ fontWeight:700, fontSize:15, color:'#111' }}>{title}</span>
      <button onClick={close} style={{ background:'none', border:'none', fontSize:22, cursor:'pointer', color:'#aaa', lineHeight:1, padding:0 }}>✕</button>
    </div>
  );

  if (!isMobile) {
    return (
      <div style={{ ...overlayBase, display:'flex', alignItems:'center', justifyContent:'center', padding:20 }} onClick={close}>
        <div style={{ background:'#fff', borderRadius:16, width:'100%', maxWidth:540, maxHeight:'90vh', display:'flex', flexDirection:'column', overflow:'hidden', transform: vis ? 'scale(1)' : 'scale(0.96)', transition:'transform 0.2s' }} onClick={e => e.stopPropagation()}>
          {headerBar}
          <div style={{ flex:1, overflowY:'auto', padding:'16px 20px' }}>{children}</div>
          {footer && <div style={{ padding:'12px 20px', borderTop:'1px solid #f0f0f0', display:'flex', gap:8, justifyContent:'flex-end', flexShrink:0 }}>{footer}</div>}
        </div>
      </div>
    );
  }

  return (
    <div style={overlayBase} onClick={close}>
      <div style={{ position:'absolute', bottom:0, left:0, right:0, background:'#fff', borderRadius:'20px 20px 0 0', maxHeight:'92vh', display:'flex', flexDirection:'column', transform: vis ? 'translateY(0)' : 'translateY(100%)', transition:'transform 0.26s cubic-bezier(0.32,0.72,0,1)', paddingBottom:'env(safe-area-inset-bottom,16px)' }} onClick={e => e.stopPropagation()}>
        <div style={{ width:36, height:4, background:'#ddd', borderRadius:2, margin:'12px auto 4px', flexShrink:0 }} />
        {headerBar}
        <div style={{ flex:1, overflowY:'auto', padding:'16px 20px', WebkitOverflowScrolling:'touch' }}>{children}</div>
        {footer && <div style={{ padding:'12px 20px', borderTop:'1px solid #f0f0f0', display:'flex', gap:8, background:'#fff', flexShrink:0 }}>{footer}</div>}
      </div>
    </div>
  );
}

function PlatJourSheet({ item, onClose, onSaved }) {
  const [form, setForm] = useState({ ...item });
  const [saving, setSaving] = useState(false);

  async function save() {
    setSaving(true);
    await supabase.from('menu_plat_jour').update({ nom: form.nom, description: form.description, prix: form.prix, actif: form.actif, updated_at: new Date().toISOString() }).eq('id', form.id);
    onSaved(form);
    setSaving(false);
    onClose();
  }

  const sheetTitle = form.type === 'plat' ? '🍽 Plat du jour' : form.type === 'dessert' ? '🍮 Dessert du jour' : '👨‍🍳 Suggestion du chef';
  return (
    <MenuBottomSheet
      title={sheetTitle}
      onClose={onClose}
      footer={<><button onClick={onClose} style={{ ...btnSecondary, flex:1 }}>Annuler</button><button onClick={save} disabled={saving} style={{ ...btnPrimary, flex:2 }}>{saving ? '...' : 'Enregistrer'}</button></>}
    >
      <div style={{ display:'flex', flexDirection:'column', gap:14 }}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'4px 0' }}>
          <span style={{ fontSize:14, fontWeight:600, color:'#333' }}>Affiché sur la carte</span>
          <MenuToggle value={!!form.actif} onChange={() => setForm(p => ({ ...p, actif: !p.actif }))} />
        </div>
        <div><label style={lbl}>Nom</label><input value={form.nom||''} onChange={e=>setForm(p=>({...p,nom:e.target.value}))} style={inp(false)} placeholder="Nom du plat" autoFocus /></div>
        <div><label style={lbl}>Description courte</label><input value={form.description||''} onChange={e=>setForm(p=>({...p,description:e.target.value}))} style={inp(false)} placeholder="Description" /></div>
        <div><label style={lbl}>Prix</label><input value={form.prix||''} onChange={e=>setForm(p=>({...p,prix:e.target.value}))} style={inp(false)} placeholder="ex: 13,50 €" /></div>
      </div>
    </MenuBottomSheet>
  );
}

function ProduitSheet({ produit, categories, carte: defaultCarte, onSave, onClose, saving }) {
  const [form, setForm] = useState({ carte: defaultCarte, disponible: true, mise_en_avant: false, badges: [], allergenes: [], ordre: 0, ...produit });
  const [showMore, setShowMore] = useState(!!produit._focusCat);
  const [showBadges, setShowBadges] = useState(false);
  const [showAllergenes, setShowAllergenes] = useState(false);

  function toggleArr(field, val) {
    const arr = form[field] || [];
    setForm(p => ({ ...p, [field]: arr.includes(val) ? arr.filter(x => x !== val) : [...arr, val] }));
  }

  return (
    <MenuBottomSheet
      title={produit.id ? 'Modifier le produit' : 'Nouveau produit'}
      onClose={onClose}
      footer={<><button onClick={onClose} style={{ ...btnSecondary, flex:1 }}>Annuler</button><button onClick={() => onSave(form)} disabled={saving || !form.nom?.trim()} style={{ ...btnPrimary, flex:2 }}>{saving ? 'Enregistrement...' : 'Enregistrer'}</button></>}
    >
      <div style={{ display:'flex', flexDirection:'column', gap:14 }}>
        <div><label style={lbl}>Nom *</label><input value={form.nom||''} onChange={e=>setForm(p=>({...p,nom:e.target.value}))} style={inp(false)} placeholder="Nom du produit" autoFocus /></div>
        <div>
          <label style={lbl}>Prix <span style={{ fontWeight:400, color:'#bbb' }}>(ex: 18 ou 13,50 — € ajouté automatiquement)</span></label>
          <input value={form.prix||''} onChange={e=>setForm(p=>({...p,prix:e.target.value}))} style={inp(false)} placeholder="ex: 18" />
        </div>
        <div>
          <label style={lbl}>Prix détaillé <span style={{ fontWeight:400, color:'#bbb' }}>(optionnel — format libre : 25cl 4€ / 50cl 7€)</span></label>
          <input value={form.prix_detail||''} onChange={e=>setForm(p=>({...p,prix_detail:e.target.value}))} style={inp(false)} placeholder="Laissez vide pour utiliser le prix simple" />
        </div>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'2px 0' }}>
          <span style={{ fontSize:14, fontWeight:500, color:'#333' }}>Disponible</span>
          <MenuToggle value={!!form.disponible} onChange={() => setForm(p=>({...p,disponible:!p.disponible}))} />
        </div>

        <button onClick={() => setShowMore(s=>!s)} style={{ background:'none', border:'none', color:'#888', fontSize:13, fontWeight:600, cursor:'pointer', textAlign:'left', padding:'2px 0', display:'flex', alignItems:'center', gap:4 }}>
          <ChevronRight size={14} strokeWidth={2.5} style={{ transform: showMore ? 'rotate(90deg)' : 'none', transition:'transform 0.2s' }} />
          {showMore ? "Moins d'options" : "Plus d'options ›"}
        </button>

        {showMore && <>
          <div><label style={lbl}>Description</label><textarea value={form.description||''} onChange={e=>setForm(p=>({...p,description:e.target.value}))} style={{...inp(false),height:70,resize:'vertical',padding:'10px 12px'}} placeholder="Description" /></div>
          <div><label style={lbl}>Accord vin</label><input value={form.accord_vin||''} onChange={e=>setForm(p=>({...p,accord_vin:e.target.value}))} style={inp(false)} placeholder="ex: Vacqueyras 7,50 €" /></div>
          <div>
            <label style={lbl}>Catégorie</label>
            <select value={form.categorie_id||''} onChange={e=>setForm(p=>({...p,categorie_id:e.target.value}))} style={{...inp(false),cursor:'pointer'}}>
              <option value="">— Choisir —</option>
              {categories.map(c=><option key={c.id} value={c.id}>{c.nom}</option>)}
            </select>
          </div>
          <div>
            <label style={lbl}>Carte</label>
            <div style={{ display:'flex', gap:8 }}>
              {[['restaurant','Restaurant'],['brasero','Brasero'],['les-deux','Les deux']].map(([v,l])=>(
                <button key={v} type="button" onClick={()=>setForm(p=>({...p,carte:v}))} style={{ flex:1, height:38, borderRadius:10, border:`1.5px solid ${form.carte===v?'#111':'#ddd'}`, background:form.carte===v?'#111':'#fff', color:form.carte===v?'#E8C547':'#666', fontSize:12, fontWeight:600, cursor:'pointer' }}>{l}</button>
              ))}
            </div>
          </div>
          <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'2px 0' }}>
            <span style={{ fontSize:14, fontWeight:500, color:'#333' }}>Mise en avant</span>
            <MenuToggle value={!!form.mise_en_avant} onChange={() => setForm(p=>({...p,mise_en_avant:!p.mise_en_avant}))} />
          </div>
          <div>
            <button onClick={()=>setShowBadges(s=>!s)} style={{ width:'100%', background:'#f5f5f5', border:'none', borderRadius:10, padding:'10px 14px', textAlign:'left', cursor:'pointer', fontSize:13, fontWeight:600, color:'#333', display:'flex', alignItems:'center', justifyContent:'space-between' }}>
              <span>Badges — {(form.badges||[]).length} sélectionné(s)</span>
              <ChevronRight size={14} style={{ color:'#aaa', transform:showBadges?'rotate(90deg)':'none', transition:'transform 0.2s' }} />
            </button>
            {showBadges && <div style={{ display:'flex', flexWrap:'wrap', gap:8, marginTop:8, paddingLeft:2 }}>
              {MENU_BADGES.map(b=>{const on=(form.badges||[]).includes(b);return <button key={b} type="button" onClick={()=>toggleArr('badges',b)} style={{padding:'6px 14px',borderRadius:20,border:`1.5px solid ${on?'#b8860b':'#ddd'}`,background:on?'#fff8e1':'#fff',color:on?'#b8860b':'#666',fontSize:12,fontWeight:600,cursor:'pointer'}}>{b}</button>;})}
            </div>}
          </div>
          <div>
            <button onClick={()=>setShowAllergenes(s=>!s)} style={{ width:'100%', background:'#f5f5f5', border:'none', borderRadius:10, padding:'10px 14px', textAlign:'left', cursor:'pointer', fontSize:13, fontWeight:600, color:'#333', display:'flex', alignItems:'center', justifyContent:'space-between' }}>
              <span>Allergènes — {(form.allergenes||[]).length} sélectionné(s)</span>
              <ChevronRight size={14} style={{ color:'#aaa', transform:showAllergenes?'rotate(90deg)':'none', transition:'transform 0.2s' }} />
            </button>
            {showAllergenes && <div style={{ display:'flex', flexWrap:'wrap', gap:8, marginTop:8, paddingLeft:2 }}>
              {MENU_ALLERGENES.map(a=>{const on=(form.allergenes||[]).includes(a);return <button key={a} type="button" onClick={()=>toggleArr('allergenes',a)} style={{padding:'5px 11px',borderRadius:20,border:`1.5px solid ${on?'#dc2626':'#ddd'}`,background:on?'#fef2f2':'#fff',color:on?'#dc2626':'#666',fontSize:12,cursor:'pointer'}}>{a}</button>;})}
            </div>}
          </div>
        </>}
      </div>
    </MenuBottomSheet>
  );
}

function CartesSheet({ onClose, showToast, produits }) {
  const [cartes, setCartes] = useState([]);
  const [newNom, setNewNom] = useState('');
  const [adding, setAdding] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [editNom, setEditNom] = useState('');
  const dragIdx = useRef(null);
  const dragOverIdx = useRef(null);

  useEffect(() => {
    supabase.from('menu_cartes').select('*').order('ordre').then(({ data }) => setCartes(data || []));
  }, []);

  async function addCarte() {
    if (!newNom.trim()) return;
    setAdding(true);
    const slug = slugify(newNom.trim());
    const ordre = cartes.length > 0 ? Math.max(...cartes.map(c=>c.ordre||0))+1 : 1;
    const { data, error } = await supabase.from('menu_cartes').insert({ nom: newNom.trim(), slug, ordre }).select().single();
    if (error) { showToast('Erreur : slug déjà utilisé ?'); }
    else { setCartes(prev => [...prev, data]); setNewNom(''); showToast('Carte ajoutée ✓'); }
    setAdding(false);
  }

  async function saveNom(c) {
    if (!editNom.trim()) { setEditingId(null); return; }
    await supabase.from('menu_cartes').update({ nom: editNom.trim() }).eq('id', c.id);
    setCartes(prev => prev.map(x => x.id === c.id ? { ...x, nom: editNom.trim() } : x));
    setEditingId(null);
    showToast('Renommée ✓');
  }

  async function toggleVisible(c) {
    const val = !c.visible;
    await supabase.from('menu_cartes').update({ visible: val }).eq('id', c.id);
    setCartes(prev => prev.map(x => x.id === c.id ? { ...x, visible: val } : x));
  }

  async function deleteCarte(c) {
    const count = produits.filter(p => p.carte === c.slug || p.carte === c.slug).length;
    if (count > 0) { showToast(`Impossible : ${count} produit(s) utilisent cette carte`); return; }
    await supabase.from('menu_cartes').delete().eq('id', c.id);
    setCartes(prev => prev.filter(x => x.id !== c.id));
    showToast('Carte supprimée');
  }

  async function drop() {
    if (dragIdx.current === null || dragOverIdx.current === null || dragIdx.current === dragOverIdx.current) { dragIdx.current=null; dragOverIdx.current=null; return; }
    const next = [...cartes];
    const [moved] = next.splice(dragIdx.current, 1);
    next.splice(dragOverIdx.current, 0, moved);
    const updated = next.map((c, i) => ({ ...c, ordre: i+1 }));
    setCartes(updated);
    await Promise.all(updated.map(c => supabase.from('menu_cartes').update({ ordre: c.ordre }).eq('id', c.id)));
    dragIdx.current=null; dragOverIdx.current=null;
  }

  return (
    <MenuBottomSheet title="🗂 Gérer les cartes" onClose={onClose} footer={<button onClick={onClose} style={{ ...btnPrimary, width:'100%' }}>Fermer</button>}>
      <div style={{ display:'flex', gap:8, marginBottom:16 }}>
        <input value={newNom} onChange={e=>setNewNom(e.target.value)} placeholder="Nouvelle carte..." style={{ ...inp(false), flex:1, height:42 }} onKeyDown={e=>e.key==='Enter'&&addCarte()} />
        <button onClick={addCarte} disabled={adding} style={{ ...btnPrimary, height:42, whiteSpace:'nowrap' }}>+ Ajouter</button>
      </div>
      <p style={{ fontSize:11, color:'#bbb', marginBottom:12, lineHeight:1.5 }}>Le slug est généré automatiquement. Utilisé dans "carte" des produits. Impossible de supprimer une carte contenant des produits.</p>
      <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
        {cartes.map((c, i) => (
          <div key={c.id}
            draggable onDragStart={() => { dragIdx.current=i; }} onDragEnter={() => { dragOverIdx.current=i; }} onDragEnd={drop} onDragOver={e => e.preventDefault()}
            style={{ display:'flex', alignItems:'center', gap:10, padding:'11px 12px', background:'#f9f9f9', borderRadius:10, cursor:'grab', userSelect:'none', opacity: c.visible ? 1 : 0.5 }}>
            <span style={{ color:'#ccc', fontSize:16, flexShrink:0 }}>⠿</span>
            {editingId === c.id ? (
              <input value={editNom} onChange={e=>setEditNom(e.target.value)} onBlur={()=>saveNom(c)} onKeyDown={e=>{if(e.key==='Enter')saveNom(c);if(e.key==='Escape')setEditingId(null);}} style={{ flex:1, height:34, border:'1.5px solid #E8C547', borderRadius:7, padding:'0 10px', fontSize:13, outline:'none' }} autoFocus />
            ) : (
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontWeight:600, fontSize:13, color:'#111' }}>{c.nom}</div>
                <div style={{ fontSize:11, color:'#bbb' }}>{c.slug}</div>
              </div>
            )}
            <button onClick={()=>{setEditingId(c.id);setEditNom(c.nom);}} style={{ background:'none', border:'none', cursor:'pointer', color:'#888', display:'flex', padding:4 }}><Pencil size={13}/></button>
            <MenuToggle value={!!c.visible} onChange={()=>toggleVisible(c)} />
            <button onClick={()=>deleteCarte(c)} style={{ background:'none', border:'none', cursor:'pointer', color:'#ddd', display:'flex', padding:4 }}><Trash2 size={14}/></button>
          </div>
        ))}
      </div>
    </MenuBottomSheet>
  );
}

function CatsSheet({ categories: initCats, onClose, showToast, carte }) {
  const [cats, setCats] = useState([...initCats]);
  const [newNom, setNewNom] = useState('');
  const [newCarte, setNewCarte] = useState(carte);
  const [editingId, setEditingId] = useState(null);
  const [editNom, setEditNom] = useState('');
  const [adding, setAdding] = useState(false);
  const dragIdx = useRef(null);
  const dragOverIdx = useRef(null);

  async function addCat() {
    if (!newNom.trim()) return;
    setAdding(true);
    const ordre = cats.length > 0 ? Math.max(...cats.map(c=>c.ordre||0))+1 : 1;
    const { data } = await supabase.from('menu_categories').insert({ nom: newNom.trim(), carte: newCarte, ordre }).select().single();
    if (data) { setCats(prev=>[...prev,data]); setNewNom(''); showToast('Catégorie ajoutée ✓'); }
    setAdding(false);
  }

  async function saveNom(cat) {
    if (!editNom.trim()) { setEditingId(null); return; }
    await supabase.from('menu_categories').update({ nom: editNom.trim() }).eq('id', cat.id);
    setCats(prev=>prev.map(c=>c.id===cat.id?{...c,nom:editNom.trim()}:c));
    setEditingId(null);
    showToast('Renommée ✓');
  }

  async function toggleVisible(cat) {
    const val = cat.visible === false ? true : false;
    await supabase.from('menu_categories').update({ visible: val }).eq('id', cat.id);
    setCats(prev=>prev.map(c=>c.id===cat.id?{...c,visible:val}:c));
  }

  async function dropCat() {
    if (dragIdx.current === null || dragOverIdx.current === null || dragIdx.current === dragOverIdx.current) { dragIdx.current=null; dragOverIdx.current=null; return; }
    const next = [...cats];
    const [moved] = next.splice(dragIdx.current, 1);
    next.splice(dragOverIdx.current, 0, moved);
    const updated = next.map((c,i) => ({ ...c, ordre: i+1 }));
    setCats(updated);
    await Promise.all(updated.map(c => supabase.from('menu_categories').update({ ordre: c.ordre }).eq('id', c.id)));
    dragIdx.current=null; dragOverIdx.current=null;
  }

  return (
    <MenuBottomSheet title="Gérer les catégories" onClose={onClose} footer={<button onClick={onClose} style={{ ...btnPrimary, width:'100%' }}>Fermer</button>}>
      <div style={{ display:'flex', gap:8, marginBottom:16 }}>
        <input value={newNom} onChange={e=>setNewNom(e.target.value)} placeholder="Nouvelle catégorie..." style={{ ...inp(false), flex:1, height:42 }} onKeyDown={e=>e.key==='Enter'&&addCat()} />
        <select value={newCarte} onChange={e=>setNewCarte(e.target.value)} style={{ height:42, border:'1.5px solid #ddd', borderRadius:7, padding:'0 8px', fontSize:12, cursor:'pointer', outline:'none' }}>
          <option value="restaurant">Restaurant</option>
          <option value="brasero">Brasero</option>
          <option value="les-deux">Les deux</option>
        </select>
        <button onClick={addCat} disabled={adding} style={{ ...btnPrimary, height:42, whiteSpace:'nowrap' }}>+ Ajouter</button>
      </div>
      <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
        {cats.map((cat, i) => (
          <div key={cat.id}
            draggable
            onDragStart={() => { dragIdx.current=i; }}
            onDragEnter={() => { dragOverIdx.current=i; }}
            onDragEnd={dropCat}
            onDragOver={e => e.preventDefault()}
            style={{ display:'flex', alignItems:'center', gap:10, padding:'11px 12px', background:'#f9f9f9', borderRadius:10, cursor:'grab', userSelect:'none' }}
          >
            <span style={{ color:'#ccc', fontSize:16, flexShrink:0 }}>⠿</span>
            {editingId === cat.id ? (
              <input value={editNom} onChange={e=>setEditNom(e.target.value)} onBlur={()=>saveNom(cat)} onKeyDown={e=>{if(e.key==='Enter')saveNom(cat);if(e.key==='Escape')setEditingId(null);}} style={{ flex:1, height:34, border:'1.5px solid #E8C547', borderRadius:7, padding:'0 10px', fontSize:13, outline:'none' }} autoFocus />
            ) : (
              <span style={{ flex:1, fontSize:13, fontWeight:600, color: cat.visible===false ? '#bbb' : '#111' }}>{cat.nom}</span>
            )}
            <span style={{ fontSize:10, color:'#aaa', background:'#e8e8e8', borderRadius:5, padding:'2px 6px', flexShrink:0 }}>{cat.carte}</span>
            <button onClick={()=>{setEditingId(cat.id);setEditNom(cat.nom);}} style={{ background:'none', border:'none', cursor:'pointer', color:'#888', display:'flex', padding:4, flexShrink:0 }}><Pencil size={13}/></button>
            <MenuToggle value={cat.visible!==false} onChange={()=>toggleVisible(cat)} />
          </div>
        ))}
      </div>
    </MenuBottomSheet>
  );
}

function MenuPage({ showToast }) {
  const [carte, setCarte] = useState('restaurant');
  const [cartes, setCartes] = useState([{id:'restaurant',l:'Restaurant'},{id:'brasero',l:'Brasero'}]);
  const [categories, setCategories] = useState([]);
  const [produits, setProduits] = useState([]);
  const [platJour, setPlatJour] = useState(null);
  const [dessertJour, setDessertJour] = useState(null);
  const [suggestionJour, setSuggestionJour] = useState(null);
  const [loading, setLoading] = useState(true);
  const [openCats, setOpenCats] = useState(new Set());
  const [menuSearch, setMenuSearch] = useState('');
  const [editProduit, setEditProduit] = useState(null);
  const [catPickerOpen, setCatPickerOpen] = useState(false);
  const [platSheet, setPlatSheet] = useState(null);
  const [showGererCats, setShowGererCats] = useState(false);
  const [showCartesSheet, setShowCartesSheet] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [saving, setSaving] = useState(false);
  const [ctxMenu, setCtxMenu] = useState(null);
  const [editingPrice, setEditingPrice] = useState(null);
  const [editingCatId, setEditingCatId] = useState(null);
  const [editingCatName, setEditingCatName] = useState('');
  const dragProd = useRef(null);
  const dragOverProd = useRef(null);
  const dragCat = useRef(null);
  const dragOverCat = useRef(null);

  // Soirées
  const [soirees, setSoirees] = useState([]);
  const [soireeSheet, setSoireeSheet] = useState(null); // {} pour new, {...s} pour edit
  const [confirmDeleteSoiree, setConfirmDeleteSoiree] = useState(null);
  const dragSoiree = useRef(null);
  const dragOverSoiree = useRef(null);

  useEffect(() => { loadMenu(); setMenuSearch(''); setOpenCats(new Set()); }, [carte]);
  useEffect(() => { loadSoirees(); loadCartes(); }, []);

  useEffect(() => {
    const ch = supabase.channel('menu-rt')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_produits' }, () => loadMenu())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_categories' }, () => loadMenu())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_plat_jour' }, () => loadMenu())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_cartes' }, () => loadCartes())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_soirees' }, () => loadSoirees())
      .subscribe();
    return () => supabase.removeChannel(ch);
  }, []);

  useEffect(() => {
    if (!ctxMenu) return;
    function handle() { setCtxMenu(null); }
    document.addEventListener('pointerdown', handle);
    return () => document.removeEventListener('pointerdown', handle);
  }, [ctxMenu]);

  async function loadCartes() {
    const { data, error } = await supabase.from('menu_cartes').select('*').eq('visible', true).order('ordre');
    if (!error && data?.length) setCartes(data.map(c => ({ id: c.slug, l: c.nom, dbId: c.id })));
  }

  async function loadSoirees() {
    const { data } = await supabase.from('menu_soirees').select('*').order('ordre');
    setSoirees(data || []);
  }

  async function saveSoiree(data) {
    const { _new, ...clean } = data;
    if (clean.id) {
      await supabase.from('menu_soirees').update(clean).eq('id', clean.id);
      setSoirees(prev => prev.map(s => s.id === clean.id ? { ...s, ...clean } : s));
    } else {
      const ordre = soirees.length > 0 ? Math.max(...soirees.map(s => s.ordre || 0)) + 1 : 1;
      const { data: newS } = await supabase.from('menu_soirees').insert({ ...clean, ordre }).select().single();
      if (newS) setSoirees(prev => [...prev, newS]);
    }
    setSoireeSheet(null);
    showToast('Soirée enregistrée ✓');
  }

  async function deleteSoiree(id) {
    await supabase.from('menu_soirees').delete().eq('id', id);
    setSoirees(prev => prev.filter(s => s.id !== id));
    setConfirmDeleteSoiree(null);
    showToast('Soirée supprimée');
  }

  async function toggleSoireeVisible(s) {
    const val = !s.visible;
    setSoirees(prev => prev.map(x => x.id === s.id ? { ...x, visible: val } : x));
    await supabase.from('menu_soirees').update({ visible: val }).eq('id', s.id);
  }

  async function dropSoiree() {
    if (dragSoiree.current === null || dragOverSoiree.current === null || dragSoiree.current === dragOverSoiree.current) { dragSoiree.current=null; dragOverSoiree.current=null; return; }
    const next = [...soirees];
    const [moved] = next.splice(dragSoiree.current, 1);
    next.splice(dragOverSoiree.current, 0, moved);
    const updated = next.map((s, i) => ({ ...s, ordre: i + 1 }));
    setSoirees(updated);
    await Promise.all(updated.map(s => supabase.from('menu_soirees').update({ ordre: s.ordre }).eq('id', s.id)));
    dragSoiree.current=null; dragOverSoiree.current=null;
  }

  async function loadMenu() {
    const [cR, pR, jR] = await Promise.all([
      supabase.from('menu_categories').select('*').order('ordre'),
      supabase.from('menu_produits').select('*').order('ordre'),
      supabase.from('menu_plat_jour').select('*')
    ]);
    setCategories(cR.data || []);
    setProduits(pR.data || []);
    const pj = jR.data || [];
    setPlatJour(pj.find(p => p.type === 'plat' && p.carte === carte) || null);
    setDessertJour(pj.find(p => p.type === 'dessert' && p.carte === carte) || null);
    setSuggestionJour(pj.find(p => p.type === 'suggestion' && p.carte === carte) || null);
    setLoading(false);
  }

  async function toggleDisponible(produit) {
    const val = !produit.disponible;
    setProduits(prev => prev.map(p => p.id === produit.id ? { ...p, disponible: val } : p));
    await supabase.from('menu_produits').update({ disponible: val }).eq('id', produit.id);
  }

  async function savePrixInline(id, val) {
    setProduits(prev => prev.map(p => p.id === id ? { ...p, prix: val } : p));
    await supabase.from('menu_produits').update({ prix: val }).eq('id', id);
    setEditingPrice(null);
    showToast('Prix mis à jour ✓');
  }

  async function saveProduit(data) {
    setSaving(true);
    const { _focusCat, ...clean } = data;
    if (clean.id) {
      const { error } = await supabase.from('menu_produits').update(clean).eq('id', clean.id);
      if (!error) { setProduits(prev => prev.map(p => p.id === clean.id ? { ...p, ...clean } : p)); showToast('Produit modifié ✓'); }
    } else {
      const { data: newP, error } = await supabase.from('menu_produits').insert(clean).select().single();
      if (!error && newP) { setProduits(prev => [...prev, newP]); showToast('Produit ajouté ✓'); }
    }
    setSaving(false);
    setEditProduit(null);
  }

  async function deleteProduit(id) {
    await supabase.from('menu_produits').delete().eq('id', id);
    setProduits(prev => prev.filter(p => p.id !== id));
    setConfirmDelete(null);
    showToast('Produit supprimé');
  }

  async function dropProd() {
    if (!dragProd.current || !dragOverProd.current) return;
    const { catId: fromCat, idx: fromIdx } = dragProd.current;
    const { catId: toCat, idx: toIdx } = dragOverProd.current;
    dragProd.current = null; dragOverProd.current = null;
    if (fromCat === toCat && fromIdx === toIdx) return;

    if (fromCat === toCat) {
      const catProds = [...produits.filter(p => p.categorie_id === fromCat)].sort((a,b) => (a.ordre||0)-(b.ordre||0));
      const [moved] = catProds.splice(fromIdx, 1);
      catProds.splice(toIdx, 0, moved);
      const updated = catProds.map((p,i) => ({ ...p, ordre: i+1 }));
      setProduits(prev => { const ids = new Set(updated.map(u=>u.id)); return [...prev.filter(p=>!ids.has(p.id)), ...updated]; });
      await Promise.all(updated.map(p => supabase.from('menu_produits').update({ ordre: p.ordre }).eq('id', p.id)));
    } else {
      const fromProds = [...produits.filter(p => p.categorie_id === fromCat)].sort((a,b) => (a.ordre||0)-(b.ordre||0));
      const [moved] = fromProds.splice(fromIdx, 1);
      const toProds = [...produits.filter(p => p.categorie_id === toCat)].sort((a,b) => (a.ordre||0)-(b.ordre||0));
      const movedNew = { ...moved, categorie_id: toCat };
      toProds.splice(toIdx, 0, movedNew);
      const updatedFrom = fromProds.map((p,i) => ({ ...p, ordre: i+1 }));
      const updatedTo = toProds.map((p,i) => ({ ...p, ordre: i+1 }));
      const all = [...updatedFrom, ...updatedTo];
      setProduits(prev => { const ids = new Set(all.map(u=>u.id)); return [...prev.filter(p=>!ids.has(p.id)), ...all]; });
      await Promise.all([
        ...updatedFrom.map(p => supabase.from('menu_produits').update({ ordre: p.ordre }).eq('id', p.id)),
        ...updatedTo.map(p => supabase.from('menu_produits').update({ ordre: p.ordre, categorie_id: p.categorie_id }).eq('id', p.id)),
      ]);
      const destCat = categories.find(c => c.id === toCat);
      showToast(`Déplacé vers ${destCat?.nom || 'autre catégorie'} ✓`);
    }
  }

  async function dropCatAccordion() {
    if (dragCat.current === null || dragOverCat.current === null || dragCat.current === dragOverCat.current) { dragCat.current=null; dragOverCat.current=null; return; }
    const next = [...catsVisible];
    const [moved] = next.splice(dragCat.current, 1);
    next.splice(dragOverCat.current, 0, moved);
    const updated = next.map((c, i) => ({ ...c, ordre: i + 1 }));
    setCategories(prev => { const ids = new Set(updated.map(u=>u.id)); return [...prev.filter(c=>!ids.has(c.id)), ...updated].sort((a,b)=>(a.ordre||0)-(b.ordre||0)); });
    await Promise.all(updated.map(c => supabase.from('menu_categories').update({ ordre: c.ordre }).eq('id', c.id)));
    dragCat.current=null; dragOverCat.current=null;
  }

  async function saveCatName(cat) {
    if (!editingCatName.trim()) { setEditingCatId(null); return; }
    await supabase.from('menu_categories').update({ nom: editingCatName.trim() }).eq('id', cat.id);
    setCategories(prev => prev.map(c => c.id === cat.id ? { ...c, nom: editingCatName.trim() } : c));
    setEditingCatId(null);
    showToast('Catégorie renommée ✓');
  }

  function toggleCat(id) {
    setOpenCats(prev => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  }

  function highlight(text, q) {
    if (!q || !text) return text;
    const t = String(text);
    const idx = normalizeStr(t).indexOf(normalizeStr(q));
    if (idx < 0) return t;
    return <>{t.slice(0,idx)}<mark style={{ background:'#fffbea', color:'#b8860b', borderRadius:3, padding:'0 1px' }}>{t.slice(idx,idx+q.length)}</mark>{t.slice(idx+q.length)}</>;
  }

  const catsFiltered = categories.filter(c => c.visible !== false && (c.carte === carte || c.carte === 'les-deux') && c.nom !== 'Plat du jour');
  const searchQ = menuSearch.trim();

  function produitsForCat(catId) {
    const base = produits.filter(p => p.categorie_id === catId && (p.carte === carte || p.carte === 'les-deux')).sort((a,b) => (a.ordre||0)-(b.ordre||0));
    if (!searchQ) return base;
    return base.filter(p => normalizeStr(p.nom||'').includes(normalizeStr(searchQ)) || normalizeStr(p.description||'').includes(normalizeStr(searchQ)));
  }

  const catsVisible = searchQ ? catsFiltered.filter(c => produitsForCat(c.id).length > 0) : catsFiltered;
  const platItems = [platJour, dessertJour, suggestionJour].filter(Boolean);

  if (loading) return <div style={{ display:'flex', alignItems:'center', justifyContent:'center', height:'60vh', color:'#888', fontSize:15 }}>Chargement du menu...</div>;

  return (
    <div style={{ padding:'24px 28px', maxWidth:900, margin:'0 auto' }}>

      {/* Header */}
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16 }}>
        <div>
          <h1 style={{ margin:0, fontSize:24, fontWeight:900, color:'#111' }}>Menu</h1>
          <p style={{ margin:'3px 0 0', fontSize:13, color:'#aaa' }}>Gérez la carte en temps réel</p>
        </div>
        <div style={{ display:'flex', alignItems:'center', gap:8 }}>
          <button onClick={() => setShowCartesSheet(true)} style={{ background:'none', border:'none', cursor:'pointer', color:'#888', fontSize:18, padding:'6px', borderRadius:8, display:'flex', alignItems:'center' }} title="Gérer les cartes">🗂</button>
          <button onClick={() => setShowGererCats(true)} style={{ background:'none', border:'none', cursor:'pointer', color:'#888', fontSize:20, padding:'6px', borderRadius:8, display:'flex', alignItems:'center' }} title="Gérer les catégories">⚙️</button>
          <button onClick={() => catsFiltered.length > 0 ? setCatPickerOpen(true) : setEditProduit({ carte, disponible: true, mise_en_avant: false, badges: [], allergenes: [] })} style={{ ...btnPrimary, height:38, fontSize:13 }}>+ Ajouter</button>
        </div>
      </div>

      {/* Onglets + lien carte client */}
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:20 }}>
        <div style={{ display:'flex', gap:6 }}>
          {cartes.map(c => (
            <button key={c.id} onClick={() => setCarte(c.id)} style={{ padding:'7px 18px', borderRadius:20, fontWeight:700, fontSize:13, cursor:'pointer', border:'none', background: carte===c.id ? '#E8C547' : '#efefef', color: carte===c.id ? '#111' : '#888', transition:'all 0.15s' }}>{c.l}</button>
          ))}
        </div>
        <a href="/accueil.html" target="_blank" rel="noopener noreferrer" style={{ fontSize:12, color:'#888', display:'flex', alignItems:'center', gap:4, textDecoration:'none', border:'1px solid #ddd', borderRadius:8, padding:'5px 10px', background:'#fff' }}>
          <ExternalLink size={12} strokeWidth={2} /> Carte client
        </a>
      </div>

      {/* Plat du jour / Dessert du jour */}
      {platItems.length > 0 && (
        <div style={{ display:'grid', gridTemplateColumns: `repeat(${platItems.length},1fr)`, gap:12, marginBottom:24 }}>
          {platItems.map(item => (
            <div key={item.id} style={{ background:'#fff', borderRadius:14, padding:'14px 16px', border:`1.5px solid ${item.actif ? '#E8C547' : '#eee'}`, transition:'border-color 0.2s' }}>
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:8 }}>
                <span style={{ fontSize:11, fontWeight:700, color:'#888', textTransform:'uppercase', letterSpacing:0.5 }}>{item.type==='plat' ? '🍽 Plat du jour' : item.type==='dessert' ? '🍮 Dessert du jour' : '👨‍🍳 Suggestion du chef'}</span>
                <MenuToggle value={!!item.actif} onChange={async () => { const v=!item.actif; item.type==='plat'?setPlatJour(p=>({...p,actif:v})):item.type==='dessert'?setDessertJour(p=>({...p,actif:v})):setSuggestionJour(p=>({...p,actif:v})); await supabase.from('menu_plat_jour').update({actif:v,updated_at:new Date().toISOString()}).eq('id',item.id); }} />
              </div>
              <div style={{ fontSize:14, fontWeight:600, color: item.actif ? '#111' : '#bbb', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', marginBottom:2 }}>
                {item.nom || <span style={{ color:'#ddd', fontStyle:'italic', fontWeight:400 }}>Non défini</span>}
              </div>
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginTop:4 }}>
                <span style={{ fontSize:12, color:'#aaa' }}>{formatPrix(item.prix)}</span>
                <button onClick={() => setPlatSheet(item)} style={{ fontSize:12, color:'#666', background:'none', border:'none', cursor:'pointer', fontWeight:600, textDecoration:'underline', padding:0 }}>Modifier</button>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Recherche */}
      <div style={{ position:'relative', marginBottom:20 }}>
        <span style={{ position:'absolute', left:13, top:'50%', transform:'translateY(-50%)', color:'#bbb', fontSize:14, pointerEvents:'none' }}>🔍</span>
        <input value={menuSearch} onChange={e => { setMenuSearch(e.target.value); if (e.target.value.trim()) setOpenCats(new Set(catsFiltered.map(c => c.id))); }} placeholder="Rechercher un produit..." style={{ width:'100%', height:44, border:'1.5px solid #eee', borderRadius:12, padding:'0 36px 0 38px', fontSize:14, outline:'none', boxSizing:'border-box', background:'#f9f9f9' }} />
        {menuSearch && <button onClick={() => { setMenuSearch(''); setOpenCats(new Set()); }} style={{ position:'absolute', right:12, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', fontSize:16, cursor:'pointer', color:'#aaa' }}>✕</button>}
      </div>

      {searchQ && catsVisible.length === 0 && <div style={{ textAlign:'center', padding:'40px 0', color:'#bbb', fontSize:14 }}>Aucun produit pour "{menuSearch}"</div>}

      {/* Accordéon catégories */}
      {catsVisible.map((cat, catIdx) => {
        const ps = produitsForCat(cat.id);
        const allPs = produits.filter(p => p.categorie_id === cat.id).length;
        const isOpen = searchQ ? true : openCats.has(cat.id);
        return (
          <div key={cat.id}
            draggable={!searchQ}
            onDragStart={() => { dragCat.current = catIdx; }}
            onDragEnter={() => { dragOverCat.current = catIdx; }}
            onDragEnd={dropCatAccordion}
            onDragOver={e => e.preventDefault()}
            style={{ marginBottom:8, borderRadius:14, background:'#fff', border:'1px solid #eee', overflow:'hidden' }}>
            <div style={{ display:'flex', alignItems:'center', gap:10, padding:'13px 16px', borderBottom: isOpen ? '1px solid #f5f5f5' : 'none' }}>
              {!searchQ && <span style={{ color:'#d0d0d0', fontSize:15, cursor:'grab', flexShrink:0, userSelect:'none' }}>⠿</span>}
              {!searchQ && <ChevronRight size={15} strokeWidth={2.5} onClick={() => toggleCat(cat.id)} style={{ color:'#ccc', flexShrink:0, transform: isOpen ? 'rotate(90deg)' : 'none', transition:'transform 0.2s', cursor:'pointer' }} />}
              {editingCatId === cat.id ? (
                <input
                  value={editingCatName}
                  onChange={e => setEditingCatName(e.target.value)}
                  onBlur={() => saveCatName(cat)}
                  onKeyDown={e => { if(e.key==='Enter') saveCatName(cat); if(e.key==='Escape') setEditingCatId(null); }}
                  style={{ flex:1, height:28, border:'1.5px solid #E8C547', borderRadius:7, padding:'0 8px', fontSize:11, fontWeight:700, textTransform:'uppercase', letterSpacing:1.1, outline:'none' }}
                  autoFocus
                  onClick={e => e.stopPropagation()}
                />
              ) : (
                <span onClick={() => !searchQ && toggleCat(cat.id)} style={{ flex:1, fontSize:11, fontWeight:700, color:'#888', textTransform:'uppercase', letterSpacing:1.1, cursor: searchQ ? 'default' : 'pointer', userSelect:'none' }}>
                  {cat.nom}
                  <span style={{ fontWeight:400, color:'#ccc', marginLeft:6, textTransform:'none', letterSpacing:0 }}>({allPs})</span>
                  {searchQ && ps.length < allPs && <span style={{ color:'#E8C547', marginLeft:6, fontWeight:700 }}> · {ps.length} résultat{ps.length>1?'s':''}</span>}
                </span>
              )}
              {!searchQ && (
                <button onClick={e => { e.stopPropagation(); setEditingCatId(cat.id); setEditingCatName(cat.nom); }} style={{ background:'none', border:'none', cursor:'pointer', color:'#d0d0d0', display:'flex', padding:'2px 4px', flexShrink:0 }} title="Renommer">
                  <Pencil size={12} />
                </button>
              )}
              <button onClick={e => { e.stopPropagation(); setEditProduit({ categorie_id: cat.id, carte, disponible: true, mise_en_avant: false, badges: [], allergenes: [], ordre: allPs+1 }); }} style={{ ...btnSecondary, height:28, fontSize:11, display:'inline-flex', alignItems:'center', gap:3, padding:'0 9px', flexShrink:0 }}>
                <Plus size={11} strokeWidth={2.5} /> Ajouter
              </button>
            </div>

            {isOpen && (
              <div>
                {ps.length === 0 ? (
                  <div style={{ padding:'14px 16px', color:'#ccc', fontSize:13, textAlign:'center' }}>Aucun produit</div>
                ) : ps.map((p, i) => (
                  <div key={p.id}
                    draggable
                    onDragStart={() => { dragProd.current = { catId: cat.id, idx: i }; }}
                    onDragEnter={() => { dragOverProd.current = { catId: cat.id, idx: i }; }}
                    onDragEnd={dropProd}
                    onDragOver={e => e.preventDefault()}
                    style={{ display:'flex', alignItems:'center', gap:10, padding:'10px 16px', borderBottom: i < ps.length-1 ? '1px solid #f7f7f7' : 'none', opacity: p.disponible ? 1 : 0.45, transition:'opacity 0.15s', background:'#fff' }}
                  >
                    <span style={{ color:'#ddd', fontSize:15, cursor:'grab', flexShrink:0, userSelect:'none' }}>⠿</span>

                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ fontWeight:500, fontSize:13, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                        {highlight(p.nom, searchQ)}
                        {p.mise_en_avant && <span style={{ marginLeft:5, fontSize:10, background:'#fffbea', color:'#b8860b', borderRadius:5, padding:'1px 5px', fontWeight:700 }}>★</span>}
                      </div>
                    </div>

                    {editingPrice?.id === p.id ? (
                      <input value={editingPrice.val} onChange={e => setEditingPrice(prev => ({ ...prev, val: e.target.value }))} onBlur={() => savePrixInline(p.id, editingPrice.val)} onKeyDown={e => { if(e.key==='Enter') savePrixInline(p.id, editingPrice.val); if(e.key==='Escape') setEditingPrice(null); }} style={{ width:80, height:30, border:'1.5px solid #E8C547', borderRadius:7, padding:'0 8px', fontSize:12, outline:'none', textAlign:'right' }} autoFocus />
                    ) : (
                      <span onClick={() => setEditingPrice({ id: p.id, val: p.prix||'' })} style={{ fontSize:12, color:'#999', cursor:'pointer', whiteSpace:'nowrap', minWidth:50, textAlign:'right', padding:'4px 6px', borderRadius:6, border:'1px solid transparent' }} onMouseEnter={e => e.currentTarget.style.borderColor='#eee'} onMouseLeave={e => e.currentTarget.style.borderColor='transparent'}>
                        {p.prix_detail || formatPrix(p.prix) || <span style={{ color:'#ddd' }}>—</span>}
                      </span>
                    )}

                    <MenuToggle value={!!p.disponible} onChange={() => toggleDisponible(p)} />

                    <button
                      onClick={e => { e.stopPropagation(); const r=e.currentTarget.getBoundingClientRect(); setCtxMenu({ produit: p, x: r.right, y: r.bottom+4 }); }}
                      style={{ background:'none', border:'none', cursor:'pointer', color:'#bbb', fontSize:17, padding:'4px 5px', borderRadius:6, lineHeight:1, flexShrink:0 }}
                    >···</button>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}

      {/* Context menu ••• */}
      {ctxMenu && (
        <div onPointerDown={e => e.stopPropagation()} style={{ position:'fixed', top: ctxMenu.y, right: `calc(100vw - ${ctxMenu.x}px)`, background:'#fff', borderRadius:10, boxShadow:'0 4px 20px rgba(0,0,0,0.14)', zIndex:4000, minWidth:170, overflow:'hidden', border:'1px solid #eee' }}>
          {[
            { label:'Modifier', action: () => { setEditProduit({ ...ctxMenu.produit }); setCtxMenu(null); } },
            { label:'Changer de catégorie', action: () => { setEditProduit({ ...ctxMenu.produit, _focusCat: true }); setCtxMenu(null); } },
            { label:'Supprimer', danger: true, action: () => { setConfirmDelete(ctxMenu.produit.id); setCtxMenu(null); } },
          ].map(item => (
            <button key={item.label} onClick={item.action} style={{ display:'block', width:'100%', textAlign:'left', padding:'12px 16px', border:'none', background:'none', cursor:'pointer', fontSize:14, color: item.danger ? '#dc2626' : '#111', fontWeight: item.danger ? 600 : 500 }}
              onMouseEnter={e => e.currentTarget.style.background='#f9f9f9'} onMouseLeave={e => e.currentTarget.style.background='none'}>
              {item.label}
            </button>
          ))}
        </div>
      )}

      {/* Sélecteur de catégorie avant création produit */}
      {catPickerOpen && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.45)', zIndex:3000, display:'flex', alignItems:'center', justifyContent:'center', padding:20 }} onClick={() => setCatPickerOpen(false)}>
          <div style={{ background:'#fff', borderRadius:16, width:'100%', maxWidth:400, overflow:'hidden', boxShadow:'0 8px 32px rgba(0,0,0,0.18)' }} onClick={e => e.stopPropagation()}>
            <div style={{ padding:'16px 20px', borderBottom:'1px solid #f0f0f0', fontWeight:700, fontSize:15, color:'#111' }}>Dans quelle catégorie ?</div>
            <div style={{ maxHeight:360, overflowY:'auto' }}>
              {catsFiltered.map(cat => {
                const n = produits.filter(p => p.categorie_id === cat.id).length;
                return (
                  <button key={cat.id} onClick={() => { setCatPickerOpen(false); setEditProduit({ categorie_id: cat.id, carte, disponible: true, mise_en_avant: false, badges: [], allergenes: [], ordre: n+1 }); }}
                    style={{ display:'block', width:'100%', textAlign:'left', padding:'13px 20px', border:'none', background:'none', cursor:'pointer', fontSize:14, color:'#111', borderBottom:'1px solid #f7f7f7' }}
                    onMouseEnter={e => e.currentTarget.style.background='#f9f9f9'} onMouseLeave={e => e.currentTarget.style.background='none'}>
                    {cat.nom} <span style={{ color:'#ccc', fontSize:12 }}>({n})</span>
                  </button>
                );
              })}
              <button onClick={() => { setCatPickerOpen(false); setEditProduit({ carte, disponible: true, mise_en_avant: false, badges: [], allergenes: [] }); }}
                style={{ display:'block', width:'100%', textAlign:'left', padding:'13px 20px', border:'none', background:'none', cursor:'pointer', fontSize:13, color:'#aaa', borderTop:'1px solid #f0f0f0' }}>
                Choisir plus tard →
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Bottom sheet produit */}
      {editProduit && (
        <ProduitSheet
          produit={editProduit}
          categories={catsFiltered}
          carte={carte}
          onSave={saveProduit}
          onClose={() => setEditProduit(null)}
          saving={saving}
        />
      )}

      {/* Bottom sheet plat du jour */}
      {platSheet && (
        <PlatJourSheet
          item={platSheet}
          onClose={() => setPlatSheet(null)}
          onSaved={updated => { updated.type==='plat' ? setPlatJour(updated) : updated.type==='dessert' ? setDessertJour(updated) : setSuggestionJour(updated); showToast('Enregistré ✓'); }}
        />
      )}

      {/* Confirmation suppression */}
      {confirmDelete && (
        <ConfirmModal
          title="Supprimer ce produit ?"
          msg="Cette action est irréversible."
          danger
          okLabel="Supprimer"
          onOk={() => deleteProduit(confirmDelete)}
          onCancel={() => setConfirmDelete(null)}
        />
      )}

      {/* Gestion catégories */}
      {showGererCats && (
        <CatsSheet
          categories={catsFiltered}
          onClose={() => { setShowGererCats(false); loadMenu(); }}
          showToast={showToast}
          carte={carte}
        />
      )}

      {/* Gestion cartes dynamiques */}
      {showCartesSheet && (
        <CartesSheet
          onClose={() => { setShowCartesSheet(false); loadCartes(); }}
          showToast={showToast}
          produits={produits}
        />
      )}

      {/* ── Section Soirées ── */}
      <div style={{ marginTop:40, paddingTop:32, borderTop:'2px solid #f0f0f0' }}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16 }}>
          <div>
            <h2 style={{ margin:0, fontSize:20, fontWeight:800, color:'#111' }}>Nos soirées</h2>
            <p style={{ margin:'2px 0 0', fontSize:12, color:'#aaa' }}>Affiché sur la page d'accueil</p>
          </div>
          <button onClick={() => setSoireeSheet({ nom:'', jour:'', horaire:'', description:'', visible:true })} style={{ ...btnPrimary, height:36, fontSize:12 }}>+ Ajouter</button>
        </div>

        {soirees.length === 0 ? (
          <div style={{ textAlign:'center', padding:'32px 0', color:'#ccc', fontSize:14 }}>Aucune soirée configurée</div>
        ) : (
          <div style={{ display:'flex', flexDirection:'column', gap:6 }}>
            {soirees.map((s, i) => (
              <div key={s.id}
                draggable
                onDragStart={() => { dragSoiree.current = i; }}
                onDragEnter={() => { dragOverSoiree.current = i; }}
                onDragEnd={dropSoiree}
                onDragOver={e => e.preventDefault()}
                style={{ display:'flex', alignItems:'center', gap:10, padding:'12px 14px', background:'#fff', borderRadius:12, border:'1px solid #eee', opacity: s.visible ? 1 : 0.45, transition:'opacity 0.15s', cursor:'grab', userSelect:'none' }}
              >
                <span style={{ color:'#ddd', fontSize:15, flexShrink:0 }}>⠿</span>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontWeight:600, fontSize:14, color:'#111' }}>{s.nom}</div>
                  {(s.jour || s.horaire) && (
                    <div style={{ fontSize:12, color:'#aaa', marginTop:1 }}>
                      {[s.jour, s.horaire].filter(Boolean).join(' · ')}
                    </div>
                  )}
                </div>
                <MenuToggle value={!!s.visible} onChange={() => toggleSoireeVisible(s)} />
                <button onClick={() => setSoireeSheet({ ...s })} style={{ background:'none', border:'none', cursor:'pointer', color:'#aaa', display:'flex', padding:4 }}><Pencil size={14}/></button>
                <button onClick={() => setConfirmDeleteSoiree(s.id)} style={{ background:'none', border:'none', cursor:'pointer', color:'#ddd', display:'flex', padding:4 }}><Trash2 size={14}/></button>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Sheet soirée */}
      {soireeSheet && (
        <MenuBottomSheet
          title={soireeSheet.id ? 'Modifier la soirée' : 'Nouvelle soirée'}
          onClose={() => setSoireeSheet(null)}
          footer={<>
            <button onClick={() => setSoireeSheet(null)} style={{ ...btnSecondary, flex:1 }}>Annuler</button>
            <button onClick={() => saveSoiree(soireeSheet)} disabled={!soireeSheet.nom?.trim()} style={{ ...btnPrimary, flex:2 }}>Enregistrer</button>
          </>}
        >
          <div style={{ display:'flex', flexDirection:'column', gap:14 }}>
            <div><label style={lbl}>Nom *</label><input value={soireeSheet.nom||''} onChange={e=>setSoireeSheet(p=>({...p,nom:e.target.value}))} style={inp(false)} placeholder="Ex : La Bringue" autoFocus /></div>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12 }}>
              <div><label style={lbl}>Jour</label><input value={soireeSheet.jour||''} onChange={e=>setSoireeSheet(p=>({...p,jour:e.target.value}))} style={inp(false)} placeholder="Ex : Vendredi" /></div>
              <div><label style={lbl}>Horaire</label><input value={soireeSheet.horaire||''} onChange={e=>setSoireeSheet(p=>({...p,horaire:e.target.value}))} style={inp(false)} placeholder="Ex : 22h00" /></div>
            </div>
            <div><label style={lbl}>Description</label><textarea value={soireeSheet.description||''} onChange={e=>setSoireeSheet(p=>({...p,description:e.target.value}))} style={{...inp(false),height:70,resize:'vertical',padding:'10px 12px'}} placeholder="Description optionnelle" /></div>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'2px 0' }}>
              <span style={{ fontSize:14, fontWeight:500, color:'#333' }}>Visible sur la page d'accueil</span>
              <MenuToggle value={!!soireeSheet.visible} onChange={() => setSoireeSheet(p=>({...p,visible:!p.visible}))} />
            </div>
            {soireeSheet.id && (
              <button onClick={() => { setConfirmDeleteSoiree(soireeSheet.id); setSoireeSheet(null); }} style={{ ...btnDanger, marginTop:4 }}>Supprimer cette soirée</button>
            )}
          </div>
        </MenuBottomSheet>
      )}

      {/* Confirmation suppression soirée */}
      {confirmDeleteSoiree && (
        <ConfirmModal
          title="Supprimer cette soirée ?"
          msg="Elle disparaîtra de la page d'accueil."
          danger okLabel="Supprimer"
          onOk={() => deleteSoiree(confirmDeleteSoiree)}
          onCancel={() => setConfirmDeleteSoiree(null)}
        />
      )}
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
  const [addClientForm, setAddClientForm] = useState({});
  const [filtreGenreClients, setFiltreGenreClients] = useState('Tous');
  const [rechercheClients, setRechercheClients] = useState('');
  const [showConfirmQuitterClient, setShowConfirmQuitterClient] = useState(false);
  const [modalDetailClient, setModalDetailClient] = useState(null);
  const [ficheClientReadOnly, setFicheClientReadOnly] = useState(false);
  const [showToutesResas, setShowToutesResas] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [showConfirmEnvoi, setShowConfirmEnvoi] = useState(false);
  const [sendingModal, setSendingModal] = useState(null); // { type, total, done, successCount }
  const [showTop300, setShowTop300] = useState(false);
  const [showTopClients, setShowTopClients] = useState(false);
  const [triColonne, setTriColonne] = useState('nom');
  const [triSens, setTriSens] = useState('asc');
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [statsClients, setStatsClients] = useState({});
  const [topJours, setTopJours] = useState([]);
  const [resasData, setResasData] = useState([]);
  const [modalEdit, setModalEdit] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [showConfirmDeconnexion, setShowConfirmDeconnexion] = useState(false);
  const [showFormulaireDropdown, setShowFormulaireDropdown] = useState(false);
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
  const [mobileTab, setMobileTab] = useState(window.innerWidth < 768 ? 'reservations' : 'clients'); // 'clients' | 'reservations'
  const [showAddResa, setShowAddResa] = useState(false);
  const [activeView, setActiveView] = useState('reservations'); // 'reservations' | 'clients' | 'communications'
  const [screenWidth, setScreenWidth] = useState(window.innerWidth);
  const [showConfirmQuitter, setShowConfirmQuitter] = useState(false);
  const [pendingFermer, setPendingFermer] = useState(null);
  const [commFilter, setCommFilter] = useState('tous');
  const [filtreGenresComm, setFiltreGenresComm] = useState(new Set());
  function toggleGenreComm(genre) {
    setFiltreGenresComm(prev => { const next = new Set(prev); next.has(genre)?next.delete(genre):next.add(genre); return next; });
  }
  const [commType, setCommType] = useState('email');
  const [nomCampagne, setNomCampagne] = useState('');
  const [showHistorique, setShowHistorique] = useState(false);
  const [filtreJours, setFiltreJours] = useState(new Set());
  const [filtreServices, setFiltreServices] = useState(new Set());
  const [showJoursDropdown, setShowJoursDropdown] = useState(false);
  const [showSegmentDropdown, setShowSegmentDropdown] = useState(false);
  const [showServiceDropdown, setShowServiceDropdown] = useState(false);
  function toggleFiltreJour(jour) { setFiltreJours(prev => { const n=new Set(prev); n.has(jour)?n.delete(jour):n.add(jour); return n; }); }
  function toggleFiltreService(service) { setFiltreServices(prev => { const n=new Set(prev); n.has(service)?n.delete(service):n.add(service); return n; }); }
  const [filtreAbsentsMois, setFiltreAbsentsMois] = useState(0);
  const [filtreAbsentsActif, setFiltreAbsentsActif] = useState(false);
  const [commSearch, setCommSearch] = useState('');
  const [commSelected, setCommSelected] = useState([]);
  const [commObjet, setCommObjet] = useState('');
  const [commMessage, setCommMessage] = useState('');
  const [commSending, setCommSending] = useState(false);
  const [showConfirmComm, setShowConfirmComm] = useState(false);
  const [showEmojiPicker, setShowEmojiPicker] = useState(false);
  const [doublons, setDoublons] = useState([]);
  const [emailsHistorique, setEmailsHistorique] = useState([]);
  const [emailsExpanded, setEmailsExpanded] = useState({});
  const commTextareaRef = useRef(null);
  const [commMode, setCommMode] = useState('email');
  const [smsMessage, setSmsMessage] = useState('');
  const [smsSelected, setSmsSelected] = useState([]);
  const [smsFilter, setSmsFilter] = useState('tous');
  const [smsSearch, setSmsSearch] = useState('');
  const [showConfirmSms, setShowConfirmSms] = useState(false);
  const [showSmsEmojiPicker, setShowSmsEmojiPicker] = useState(false);
  const [smsHistorique, setSmsHistorique] = useState([]);
  const [smsExpanded, setSmsExpanded] = useState({});
  const smsTextareaRef = useRef(null);
  const [notifResa, setNotifResa] = useState(null);
  const [showNotifPrePrompt, setShowNotifPrePrompt] = useState(false);

  useEffect(()=>{
    const notifAsked = localStorage.getItem('ted_notif_asked');
    if (!notifAsked && typeof Notification !== 'undefined' && Notification.permission === 'default') {
      setTimeout(()=>setShowNotifPrePrompt(true), 3000);
    }
  }, []);

  async function initOneSignal() {
    if (typeof window.OneSignalDeferred === 'undefined') window.OneSignalDeferred = [];
    window.OneSignalDeferred.push(async function(OneSignal) {
      await OneSignal.init({
        appId: '87b29550-ffb0-412a-9682-05fdace514fc',
        safari_web_id: 'web.onesignal.auto.87b29550-ffb0-412a-9682-05fdace514fc',
        notifyButton: { enable: false },
        allowLocalhostAsSecureOrigin: true
      });
      console.log('OneSignal initialisé');
      const userId = user?.id || 'ted-admin';
      await OneSignal.login(userId);
      console.log('OneSignal user logged in:', userId);
    });
  }

  async function demanderPermissionNotif() {
    window.OneSignalDeferred = window.OneSignalDeferred || [];
    window.OneSignalDeferred.push(async function(OneSignal) {
      const isSubscribed = await OneSignal.User.PushSubscription.optedIn;
      if (isSubscribed) {
        showToast('🔔 Notifications déjà activées !');
        return;
      }
      const permission = await OneSignal.Notifications.requestPermission();
      if (permission) {
        showToast('🔔 Les nouvelles demandes de réservation arriveront ici !');
      } else {
        showToast('⚠️ Notifications refusées', 'error');
      }
    });
  }
  const deleteGuard = useRef(false);
  const notifEnCoursRef = useRef(false);
  const isMobile = useIsMobile();

  const showToast = useCallback((msg, type="success") => setToast({msg,type}), []);

  function toggleEmailExpanded(id) {
    setEmailsExpanded(prev => ({...prev, [id]: !prev[id]}));
  }

  function toggleSmsExpanded(id) {
    setSmsExpanded(prev => ({...prev, [id]: !prev[id]}));
  }

  async function loadEmailsHistorique() {
    const { data } = await supabase.from('emails_envoyes').select('*').order('created_at', {ascending:false}).limit(50);
    setEmailsHistorique(data || []);
  }

  async function loadSmsHistorique() {
    const { data } = await supabase.from('sms_envoyes').select('*').order('created_at', {ascending:false}).limit(50);
    setSmsHistorique(data || []);
  }

  // ─── Load from Supabase ───────────────────────────────────────────────────
  useEffect(() => {
    loadClients();
    loadResaCount();
  }, []);

  useEffect(() => {
    if (modalEdit) setEditForm({ genre: modalEdit.genre||'', prenom: modalEdit.prenom||'', nom: modalEdit.nom||'', tel: modalEdit.tel||'', mail: modalEdit.mail||'', entreprise: modalEdit.entreprise||'', commentaire: modalEdit.commentaire||'' });
  }, [modalEdit]);

  useEffect(() => {
    const handler = () => setScreenWidth(window.innerWidth);
    window.addEventListener('resize', handler);
    return () => window.removeEventListener('resize', handler);
  }, []);

  useEffect(() => {
    if (activeView === 'communications') { loadEmailsHistorique(); loadSmsHistorique(); }
    setModalDetailClient(null);
    setFicheClientReadOnly(false);
  }, [activeView]);

  useEffect(() => { if (user) initOneSignal(); }, [user]);

  useEffect(() => {
    const channel = supabase
      .channel('nouvelles-reservations')
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'reservations', filter: 'statut=eq.attente' }, async (payload) => {
        if (notifEnCoursRef.current) return;
        notifEnCoursRef.current = true;
        console.log('PAYLOAD REÇU:', payload);
        console.log('Permission notifications:', Notification.permission);
        console.log('Service Worker disponible:', 'serviceWorker' in navigator);
        const { data: client } = await supabase.from('clients').select('nom, prenom, tel').eq('id', payload.new.client_id).single();
        loadResaCount();
        const nom = client ? `${client.prenom} ${client.nom}` : 'Nouveau client';
        const date = new Date(payload.new.date + 'T12:00:00').toLocaleDateString('fr-FR', { weekday: 'long', day: 'numeric', month: 'long' });
        const isMob = window.innerWidth < 768;
        if (!isMob) {
          setNotifResa({ nom, message: `${date} · ${payload.new.heure || ''} · ${payload.new.nb_personnes} pers.`, id: payload.new.id });
          setTimeout(() => setNotifResa(null), 6000);
        }
        setResaAttenteCount(prev => { const n = prev + 1; updateBadge(n); return n; });
        await fetch('/send-push-onesignal', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            title: '📅 Nouvelle réservation !',
            body: `${nom} · ${date} · ${payload.new.heure || ''} · ${payload.new.nb_personnes} pers.`
          })
        });
        notifEnCoursRef.current = false;
      })
      .subscribe((status) => {
        console.log('Statut Realtime:', status);
      });
    return () => { supabase.removeChannel(channel); };
  }, []);

  async function updateBadge(count) {
    if ('setAppBadge' in navigator) {
      try {
        if (count > 0) await navigator.setAppBadge(count);
        else await navigator.clearAppBadge();
      } catch(e) {}
    }
  }

  async function loadResaCount() {
    const { count } = await supabase.from('reservations').select('id', { count:'exact', head:true }).eq('statut','attente');
    const n = count || 0;
    setResaAttenteCount(n);
    updateBadge(n);
  }

  async function loadClients(silent = false) {
    if (!silent) setLoading(true);
    const { data, error } = await supabase.from("clients").select("*").is("deleted_at", null).order("created_at", { ascending: false });
    if (error) { showToast("Erreur de chargement", "error"); }
    else { setClients(data || []); }
    if (!silent) setLoading(false);
    chargerToutesStatsClients();
  }

  async function chargerToutesStatsClients() {
    const { data } = await supabase.from('reservations').select('client_id, statut, date, service');
    setResasData(data || []);
    const stats = {};
    const joursSemaine = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
    const compteParJourService = {};
    (data||[]).forEach(r => {
      if (!stats[r.client_id]) stats[r.client_id] = { total:0, noshow:0, derniereVisite:null };
      stats[r.client_id].total++;
      if (r.statut === 'absente') stats[r.client_id].noshow++;
      if (r.statut === 'venue' || r.statut === 'confirmee') {
        if (!stats[r.client_id].derniereVisite || r.date > stats[r.client_id].derniereVisite)
          stats[r.client_id].derniereVisite = r.date;
      }
      if ((r.statut === 'confirmee' || r.statut === 'venue') && r.date) {
        const jour = joursSemaine[new Date(r.date+'T12:00:00').getDay()];
        const service = r.service === 'midi' ? 'Midi' : 'Soir';
        const key = `${jour} ${service}`;
        compteParJourService[key] = (compteParJourService[key] || 0) + 1;
      }
    });
    setStatsClients(stats);
    setTopJours(Object.entries(compteParJourService).sort((a,b) => b[1]-a[1]).slice(0,3));
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
    setModalDetailClient(prev => prev && prev.id === c.id ? {...prev, ...c} : prev);
    setModalEdit(null);
    showToast("Client modifié avec succès ✓");
    const { error } = await supabase.from("clients").update({ genre:c.genre, nom:c.nom, prenom:c.prenom, tel:c.tel, mail:c.mail, commentaire:c.commentaire, entreprise:c.entreprise||"" }).eq("id", c.id);
    if (error) {
      showToast("Erreur lors de la modification", "error");
    }
    loadClients(); // toujours recharger pour garantir la sync (BUG 3 : nouveau mail pour emails en attente)
  }

  async function sauvegarderEditClient() {
    if (!modalEdit) return;
    await editClient({ ...modalEdit, ...editForm });
  }

  async function deleteClient(id) {
    if (deleteGuard.current) return;
    deleteGuard.current = true;
    setClients(prev => prev.filter(x => x.id !== id));
    setModalDelete(null);
    setModalDetailClient(null);
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

  const sidebarDesktop = !isMobile ? (
    <div style={{ position:'fixed', top:0, left:0, bottom:0, width:120, background:'#111', display:'flex', flexDirection:'column', alignItems:'center', padding:'20px 0', zIndex:100, borderRight:'1px solid #222' }}>
      <img src="/favicon.png" style={{ width:44, height:44 }} alt="TED" />
      <span style={{ fontSize:10, fontWeight:800, color:'#E8C547', letterSpacing:2, marginTop:4, marginBottom:28 }}>LE TED</span>
      {[
        { id:'reservations', label:'Réservations', icon:<CalendarDays size={24} strokeWidth={1.8} /> },
        { id:'clients', label:'Clients', icon:<Users size={24} strokeWidth={1.8} /> },
        { id:'communications', label:'Communications', icon:<Megaphone size={24} strokeWidth={1.8} /> },
        { id:'menu', label:'Menu', icon:<UtensilsCrossed size={24} strokeWidth={1.8} /> },
      ].map(item => {
        const nbAttenteSidebar = item.id === 'reservations' ? resaAttenteCount : 0;
        return (
          <button key={item.id} onClick={()=>setActiveView(item.id)} style={{ width:'100%', padding:'12px 8px', border:'none', display:'flex', flexDirection:'column', alignItems:'center', gap:6, cursor:'pointer', marginBottom:4, borderLeft: activeView===item.id ? '3px solid #E8C547' : '3px solid transparent', background: activeView===item.id ? 'rgba(232,197,71,0.1)' : 'transparent', color: activeView===item.id ? '#E8C547' : '#555', position:'relative' }}>
            {item.icon}
            <span style={{ fontSize:10, fontWeight:600, textAlign:'center', lineHeight:1.2 }}>{item.label}</span>
            {nbAttenteSidebar > 0 && (
              <div className="notif-badge-alarm" style={{
                position:'absolute', top:2, right:6,
                minWidth:22, height:22, borderRadius:11,
                background:'#dc2626', border:'2.5px solid #111',
                boxShadow:'0 0 12px rgba(220,38,38,1)',
                display:'flex', alignItems:'center', justifyContent:'center',
                fontSize:11, fontWeight:900, color:'#fff',
                padding:'0 5px'
              }}>
                {nbAttenteSidebar}
              </div>
            )}
          </button>
        );
      })}
      <div style={{ flex:1 }} />
      <button onClick={()=>setShowConfirmDeconnexion(true)} style={{ width:'100%', padding:'12px 8px', border:'none', background:'none', display:'flex', flexDirection:'column', alignItems:'center', gap:6, cursor:'pointer', color:'#555' }}>
        <LogOut size={22} strokeWidth={1.8} />
        <span style={{ fontSize:10, fontWeight:600 }}>Déconnexion</span>
      </button>
      {showConfirmDeconnexion && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:9000, display:'flex', alignItems:'center', justifyContent:'center', pointerEvents:'all', cursor:'default', touchAction:'none' }} onMouseDown={e=>{e.preventDefault();e.stopPropagation();setShowConfirmDeconnexion(false);}}>
          <div style={{ background:'#fff', borderRadius:16, padding:'28px 24px', maxWidth:320, width:'90%', textAlign:'center', boxShadow:'0 20px 60px rgba(0,0,0,0.2)' }} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()}>
            <h3 style={{ margin:'0 0 8px', fontSize:17, fontWeight:800 }}>Se déconnecter ?</h3>
            <p style={{ color:'#888', fontSize:14, margin:'0 0 20px' }}>Vous devrez vous reconnecter pour accéder au CRM.</p>
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={()=>setShowConfirmDeconnexion(false)} style={{ flex:1, height:44, border:'1.5px solid #ddd', borderRadius:10, background:'#fff', fontSize:14, cursor:'pointer', color:'#666' }}>Annuler</button>
              <button onClick={()=>{ supabase.auth.signOut(); setShowConfirmDeconnexion(false); }} style={{ flex:1, height:44, border:'none', borderRadius:10, background:'#111', fontSize:14, fontWeight:800, cursor:'pointer', color:'#fff' }}>Se déconnecter</button>
            </div>
          </div>
        </div>
      )}
    </div>
  ) : null;

  const notifPrePromptModal = showNotifPrePrompt ? (
    <>
      <div onClick={()=>setShowNotifPrePrompt(false)} style={{
        position:'fixed', inset:0, background:'rgba(0,0,0,0.4)',
        zIndex:8000, pointerEvents:'all'
      }}/>
      <div onClick={e=>e.stopPropagation()} style={{
        position:'fixed', top:'50%', left:'50%',
        transform:'translate(-50%,-50%)',
        background:'#fff', borderRadius:20,
        width:'min(380px, calc(100vw - 48px))',
        padding:'32px 28px', textAlign:'center',
        boxShadow:'0 32px 80px rgba(0,0,0,0.25)',
        zIndex:8001
      }}>
        <div style={{
          width:64, height:64, borderRadius:'50%',
          background:'#fffbea', border:'3px solid #E8C547',
          display:'flex', alignItems:'center', justifyContent:'center',
          margin:'0 auto 20px', fontSize:28
        }}>
          🔔
        </div>
        <h2 style={{fontSize:20, fontWeight:900, color:'#111', margin:'0 0 10px'}}>
          Activer les notifications
        </h2>
        <p style={{fontSize:14, color:'#666', lineHeight:1.6, margin:'0 0 24px'}}>
          Soyez alerté instantanément quand une nouvelle réservation arrive, même si l'app est en arrière-plan.
        </p>
        <button onClick={async()=>{
          localStorage.setItem('ted_notif_asked', 'true');
          setShowNotifPrePrompt(false);
          await Notification.requestPermission();
        }} style={{
          width:'100%', height:50, border:'none', borderRadius:14,
          background:'#E8C547', color:'#111',
          fontSize:15, fontWeight:800, cursor:'pointer', marginBottom:10
        }}>
          🔔 Activer les notifications
        </button>
        <button onClick={()=>{
          localStorage.setItem('ted_notif_asked', 'true');
          setShowNotifPrePrompt(false);
        }} style={{
          width:'100%', background:'none', border:'none',
          color:'#999', fontSize:13, cursor:'pointer', padding:'8px'
        }}>
          Plus tard
        </button>
      </div>
    </>
  ) : null;

  if (!isMobile && activeView === 'reservations') return (
    <>
      {sidebarDesktop}
      <div style={{ marginLeft:120, minHeight:'100vh' }}>
        <ReservationsPage inline showToast={showToast} user={user} onResaCountChange={(n)=>{ setResaAttenteCount(n); updateBadge(n); }} />
      </div>
      {toast && <Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)} />}
      {notifPrePromptModal}
    </>
  );


  if (!isMobile && activeView === 'menu') return (
    <>
      {sidebarDesktop}
      <div style={{ marginLeft:120, minHeight:'100vh', background:'#f5f5f5' }}>
        <MenuPage showToast={showToast} />
      </div>
      {toast && <Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)} />}
      {notifPrePromptModal}
    </>
  );

  if (activeView === 'communications' && !isMobile) {
    const limiteCommDate = (() => { const d = new Date(); d.setMonth(d.getMonth() - filtreAbsentsMois); return d.toISOString().split('T')[0]; })();
    const il6MoisComm = new Date(Date.now() - 180*24*60*60*1000).toISOString().split('T')[0];
    const joursSem = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
    const isNumeroMobile = (tel) => /^(\+336|\+337|06|07)/.test((tel||'').replace(/[\s.\-()]/g,''));

    // Filtre unifié pour les deux modes
    const clientsFiltresComm = clients.filter(c => {
      if (filtreGenresComm.size > 0 && !filtreGenresComm.has(c.genre)) return false;
      const q = commSearch.toLowerCase();
      if (q && !normalizeStr(c.nom||'').includes(normalizeStr(q)) && !normalizeStr(c.prenom||'').includes(normalizeStr(q)) && !(c.mail||'').toLowerCase().includes(q)) return false;
      if (filtreAbsentsMois > 0) {
        const aujourd = new Date().toISOString().split('T')[0];
        const resasC = resasData.filter(r => r.client_id === c.id);
        const aResaFuture = resasC.some(r => r.date > aujourd && (r.statut === 'confirmee' || r.statut === 'attente'));
        if (aResaFuture) return false;
        const derniereResa = resasC.filter(r => r.date <= aujourd && (r.statut === 'venue' || r.statut === 'confirmee')).sort((a,b) => b.date.localeCompare(a.date))[0];
        if (derniereResa && derniereResa.date >= limiteCommDate) return false;
      }
      if (filtreJours.size > 0 || filtreServices.size > 0) {
        const resasC = resasData.filter(r => r.client_id === c.id && (r.statut === 'confirmee' || r.statut === 'venue') && r.date >= il6MoisComm);
        const compteJ = {};
        resasC.forEach(r => { const key = `${joursSem[new Date(r.date+'T12:00:00').getDay()]}_${r.service}`; compteJ[key] = (compteJ[key]||0)+1; });
        const top3 = Object.entries(compteJ).sort((a,b)=>b[1]-a[1]).slice(0,3).map(e=>e[0]);
        let match = false;
        if (filtreJours.size > 0 && filtreServices.size > 0) {
          for (const jour of filtreJours) { for (const srv of filtreServices) { if (top3.includes(`${jour}_${srv}`)) { match = true; break; } } if (match) break; }
        } else if (filtreJours.size > 0) {
          for (const jour of filtreJours) { if (top3.some(k => k.startsWith(jour+'_'))) { match = true; break; } }
        } else {
          for (const srv of filtreServices) { if (top3.some(k => k.endsWith('_'+srv))) { match = true; break; } }
        }
        if (!match) return false;
      }
      return true;
    });

    // Sélection unifiée selon le mode
    const selectedComm = commType === 'email' ? commSelected : smsSelected;
    const setSelectedComm = commType === 'email' ? setCommSelected : setSmsSelected;
    const toggleSelectionClient = (id) => {
      if (commType === 'sms' && !isNumeroMobile(clients.find(c=>c.id===id)?.tel||'')) return;
      setSelectedComm(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
    };
    const tousSelectionnes = clientsFiltresComm.length > 0 && clientsFiltresComm.filter(c => commType==='sms'?isNumeroMobile(c.tel||''):true).every(c => selectedComm.includes(c.id));
    const toggleToutSelection = () => {
      if (tousSelectionnes) setSelectedComm([]);
      else setSelectedComm(clientsFiltresComm.filter(c => commType==='sms'?isNumeroMobile(c.tel||''):true).map(c => c.id));
    };

    // Logique email
    const selectedClients = clients.filter(c => commSelected.includes(c.id) && c.mail);
    const buildHtml = (client) => {
      const msg = commMessage.replace(/\n/g,'<br>').replace(/{prenom}/g, client.prenom||'').replace(/{nom}/g, client.nom||'').replace(/{tel}/g, client.tel||'').replace(/{entreprise}/g, client.entreprise||'');
      return `<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto;padding:32px 24px;background:#fff"><p style="font-size:15px;line-height:1.7;color:#222">${msg}</p><div style="margin-top:32px;padding-top:16px;border-top:1px solid #eee"><table cellpadding="0" cellspacing="0"><tr><td style="padding-right:12px;vertical-align:top"><img src="https://ted-crm.pages.dev/favicon.png" style="height:36px;width:36px"/></td><td><p style="margin:0;font-weight:800;font-size:14px;color:#111">Le TED — Restaurant &amp; Club</p><p style="margin:4px 0 0;font-size:12px;color:#888">📍 28 Av. des Frères Montgolfier, 69680 Chassieu</p><p style="margin:2px 0 0;font-size:12px;color:#888">📞 04 78 90 67 80</p><p style="margin:2px 0 0;font-size:12px"><a href="https://leted.fr" style="color:#E8C547;text-decoration:none;font-weight:700">leted.fr</a></p></td></tr></table></div></div>`;
    };
    const doSendComm = async () => {
      setCommSending(true);
      setSendingModal({ type:'email', total: selectedClients.length, done: false, successCount: 0 });
      let sent = 0;
      for (const client of selectedClients) {
        try {
          const res = await fetch('/send-email', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({ to:client.mail, toName:`${client.prenom||''} ${client.nom||''}`.trim(), subject:commObjet, html:buildHtml(client) }) });
          const text = await res.text(); let data = {}; try { data = JSON.parse(text); } catch(_) {}
          if (data.success) sent++;
        } catch(e) { console.error('[Comm] Erreur réseau pour', client.mail, e); }
      }
      await supabase.from('emails_envoyes').insert([{ objet:commObjet, message:commMessage, nb_destinataires:commSelected.length, destinataires:commSelected.map(id => { const c = clients.find(x=>x.id===id); return {id, nom:c?.nom, prenom:c?.prenom, mail:c?.mail}; }), envoye_par:user.email, statut:'envoye' }]);
      setCommSending(false);
      setSendingModal(prev => prev ? { ...prev, done: true, successCount: sent } : null);
      setCommObjet(''); setCommMessage(''); setCommSelected([]); setNomCampagne('');
      loadEmailsHistorique();
    };
    const handleSendAll = async () => {
      const { data: dejaSent } = await supabase.from('emails_envoyes').select('destinataires').eq('objet', commObjet);
      const dejaSentIds = new Set((dejaSent||[]).flatMap(e => (e.destinataires||[]).map(d => d.id)));
      setDoublons(commSelected.filter(id => dejaSentIds.has(id)));
      setShowConfirmComm(true);
    };

    // Logique SMS
    const containsEmoji = (str) => /[\u{1F300}-\u{1F9FF}]|[\u{2600}-\u{26FF}]/u.test(str);
    const smsLimit = containsEmoji(smsMessage) ? 70 : 160;
    const doSendSms = async () => {
      const BREVO_KEY = process.env.REACT_APP_BREVO_API_KEY;

      let destinatairesFinaux = [...smsSelected];
      setSendingModal(null); // reset au cas où
      try {
        const { data: dejaSent, error: errDoublons } = await supabase.from('sms_envoyes').select('destinataires, message').eq('message', smsMessage);
        if (!errDoublons) {
          const dejaSentIds = new Set((dejaSent||[]).flatMap(s => (s.destinataires||[]).map(d => d.id)));
          const doublons = smsSelected.filter(id => dejaSentIds.has(id));
          const nouveaux = smsSelected.filter(id => !dejaSentIds.has(id));
          if (doublons.length > 0 && nouveaux.length === 0) { showToast('Ces clients ont déjà reçu ce message', 'error'); return; }
          if (doublons.length > 0) { destinatairesFinaux = nouveaux; }
        }
      } catch(e) {}

      const destinatairesMobiles = destinatairesFinaux.filter(id => {
        const client = clients.find(c => c.id === id);
        return /^(06|07|\+336|\+337)/.test((client?.tel||'').replace(/[\s.\-()]/g, ''));
      });

      if (destinatairesMobiles.length === 0) {
        showToast('Aucun numéro mobile valide (06/07)', 'error');
        return;
      }

      setSendingModal({ type:'sms', total: destinatairesMobiles.length, done: false, successCount: 0 });
      let success = 0, errors = 0;
      for (const id of destinatairesMobiles) {
        const client = clients.find(c => c.id === id);
        if (!client?.tel) { errors++; continue; }
        const tel = client.tel.replace(/[\s.\-()]/g, '').replace(/^0/, '+33');
        const msg = smsMessage
          .replace(/{prenom}/g, client.prenom || client.entreprise || '')
          .replace(/{nom}/g, client.nom || '')
          .replace(/{tel}/g, client.tel || '')
          .replace(/{entreprise}/g, client.entreprise || '')
          .replace(/{lien_resa}/g, 'https://ted-crm.pages.dev/reserver.html');
        try {
          const res = await fetch('https://api.brevo.com/v3/transactionalSMS/sms', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'api-key': BREVO_KEY },
            body: JSON.stringify({ sender: 'LE TED', recipient: tel, content: msg, type: 'marketing' })
          });
          const data = await res.json();
          console.log('Status Brevo:', res.status);
          console.log('Réponse Brevo complète:', JSON.stringify(data));
          if (!res.ok) { console.error('ERREUR Brevo:', data.message || data.error || JSON.stringify(data)); errors++; }
          else { success++; }
        } catch(err) { errors++; console.error('Fetch erreur:', err); }
        await new Promise(r => setTimeout(r, 100));
      }

      await supabase.from('sms_envoyes').insert([{
        message: smsMessage, nb_destinataires: success,
        destinataires: destinatairesMobiles.map(id => { const c = clients.find(x=>x.id===id); return {id, nom:c?.nom, prenom:c?.prenom, tel:c?.tel}; }),
        envoye_par: user.email
      }]);

      setSendingModal(prev => prev ? { ...prev, done: true, successCount: success } : null);

      setSmsMessage(''); setSmsSelected([]); setNomCampagne('');
      setShowConfirmEnvoi(false);
      loadSmsHistorique();
    };

    // Historique combiné
    const historiqueEnvois = [
      ...emailsHistorique.map(e => ({...e, type:'email'})),
      ...smsHistorique.map(s => ({...s, type:'sms'}))
    ].sort((a,b) => (b.created_at||'').localeCompare(a.created_at||''));

    return (
      <>
        {sidebarDesktop}
        <div style={{marginLeft:120, minHeight:'100vh', background:'#f5f5f5', padding:'24px 32px', boxSizing:'border-box'}}>

          {/* Header */}
          <div style={{display:'flex', alignItems:'center', gap:16, marginBottom:20, flexShrink:0}}>
            <h1 style={{fontSize:28, fontWeight:900, color:'#111', margin:0}}>Communications</h1>
            <button onClick={()=>{ loadEmailsHistorique(); loadSmsHistorique(); setShowHistorique(true); }} style={{ display:'flex', alignItems:'center', gap:6, background:'#fff', border:'1.5px solid #eee', borderRadius:8, height:40, padding:'0 16px', fontSize:14, fontWeight:600, cursor:'pointer', color:'#444' }}>
              <History size={14} strokeWidth={2} /> Historique
            </button>
          </div>

          <div style={{display:'grid', gridTemplateColumns:'240px 1fr 380px', gap:16, height:'calc(100vh - 130px)', overflow:'hidden'}}>

            {/* ─── Colonne 1 — Ciblage ─── */}
            <div style={{background:'#fff', borderRadius:16, boxShadow:'0 1px 4px rgba(0,0,0,0.04)', padding:14, display:'flex', flexDirection:'column', height:'100%', overflow:'hidden'}}>
              <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:'0 0 8px', flexShrink:0}}>Cibler vos destinataires</p>

              {/* Contenu scrollable */}
              <div style={{flex:1, overflowY:'auto', display:'flex', flexDirection:'column', gap:8}}>
                {/* Segment */}
                <div style={{flexShrink:0}}>
                  <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 5px'}}>Segment</p>
                  <div style={{position:'relative'}}>
                    <button onClick={()=>setShowSegmentDropdown(v=>!v)} style={{width:'100%', height:36, border:'1.5px solid #eee', borderRadius:9, background:'#fff', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', padding:'0 12px', fontSize:13, color:'#111', fontWeight:500}}>
                      <span style={{overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', flex:1, textAlign:'left'}}>
                        {filtreGenresComm?.size>0 ? [...filtreGenresComm].join(', ') : 'Tous les clients'}
                      </span>
                      <ChevronDown size={14} strokeWidth={2} color="#999" style={{flexShrink:0, marginLeft:6}}/>
                    </button>
                    {showSegmentDropdown && (
                      <>
                        <div onClick={()=>setShowSegmentDropdown(false)} style={{position:'fixed', inset:0, zIndex:299}}/>
                        <div style={{position:'absolute', top:'calc(100% + 4px)', left:0, right:0, background:'#fff', borderRadius:12, boxShadow:'0 8px 24px rgba(0,0,0,0.12)', zIndex:300, overflow:'hidden', border:'1.5px solid #eee'}}>
                          <div onClick={()=>{setFiltreGenresComm(new Set()); setShowSegmentDropdown(false);}}
                            style={{display:'flex', alignItems:'center', gap:10, padding:'10px 14px', cursor:'pointer', background:filtreGenresComm?.size===0?'#fffbea':'#fff', borderBottom:'1px solid #f5f5f5'}}
                            onMouseEnter={e=>e.currentTarget.style.background='#f9f9f9'}
                            onMouseLeave={e=>e.currentTarget.style.background=filtreGenresComm?.size===0?'#fffbea':'#fff'}
                          >
                            <div style={{width:16,height:16,borderRadius:4,flexShrink:0,border:'1.5px solid',borderColor:filtreGenresComm?.size===0?'#E8C547':'#ddd',background:filtreGenresComm?.size===0?'#E8C547':'#fff',display:'flex',alignItems:'center',justifyContent:'center'}}>
                              {filtreGenresComm?.size===0 && <Check size={10} strokeWidth={3} color="#111"/>}
                            </div>
                            <span style={{fontSize:13, fontWeight:500, color:'#111'}}>Tous les clients</span>
                          </div>
                          {[{id:'Homme',label:'Hommes'},{id:'Femme',label:'Femmes'},{id:'Entreprise',label:'Entreprises'}].map(s=>{
                            const actif = filtreGenresComm?.has(s.id);
                            return (
                              <div key={s.id} onClick={()=>toggleGenreComm(s.id)}
                                style={{display:'flex', alignItems:'center', gap:10, padding:'10px 14px', cursor:'pointer', background:actif?'#fffbea':'#fff', borderBottom:'1px solid #f5f5f5'}}
                                onMouseEnter={e=>e.currentTarget.style.background=actif?'#fffbea':'#f9f9f9'}
                                onMouseLeave={e=>e.currentTarget.style.background=actif?'#fffbea':'#fff'}
                              >
                                <div style={{width:16,height:16,borderRadius:4,flexShrink:0,border:'1.5px solid',borderColor:actif?'#E8C547':'#ddd',background:actif?'#E8C547':'#fff',display:'flex',alignItems:'center',justifyContent:'center'}}>
                                  {actif && <Check size={10} strokeWidth={3} color="#111"/>}
                                </div>
                                <span style={{fontSize:13, fontWeight:500, color:'#111'}}>{s.label}</span>
                              </div>
                            );
                          })}
                        </div>
                      </>
                    )}
                  </div>
                </div>

                {/* Jour favori */}
                <div style={{flexShrink:0}}>
                  <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 5px'}}>Jour favori</p>
                  <div style={{position:'relative'}}>
                    <button
                      onClick={()=>setShowJoursDropdown(v=>!v)}
                      style={{width:'100%', height:36, border:'1.5px solid #eee', borderRadius:9, background:'#fff', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', padding:'0 12px', fontSize:13, color:'#111', fontWeight:500}}
                    >
                      <span style={{overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', flex:1, textAlign:'left'}}>
                        {filtreJours.size>0 ? [...filtreJours].map(j=>j.slice(0,3)).join(', ') : 'Tous les jours'}
                      </span>
                      <ChevronDown size={14} strokeWidth={2} color="#999" style={{flexShrink:0, marginLeft:6}}/>
                    </button>
                    {showJoursDropdown && (
                      <>
                        <div onClick={()=>setShowJoursDropdown(false)} style={{position:'fixed', inset:0, zIndex:299}}/>
                        <div style={{position:'absolute', top:'calc(100% + 4px)', left:0, right:0, background:'#fff', borderRadius:12, boxShadow:'0 8px 24px rgba(0,0,0,0.12)', zIndex:300, overflow:'hidden', border:'1.5px solid #eee'}}>
                          <div
                            onClick={()=>{setFiltreJours(new Set()); setShowJoursDropdown(false);}}
                            style={{display:'flex', alignItems:'center', gap:10, padding:'10px 14px', cursor:'pointer', borderBottom:'1px solid #f5f5f5', background: filtreJours.size===0?'#fffbea':'#fff'}}
                            onMouseEnter={e=>e.currentTarget.style.background='#f9f9f9'}
                            onMouseLeave={e=>e.currentTarget.style.background=filtreJours.size===0?'#fffbea':'#fff'}
                          >
                            <div style={{width:16,height:16,borderRadius:4,flexShrink:0,border:'1.5px solid',borderColor:filtreJours.size===0?'#E8C547':'#ddd',background:filtreJours.size===0?'#E8C547':'#fff',display:'flex',alignItems:'center',justifyContent:'center'}}>
                              {filtreJours.size===0 && <Check size={10} strokeWidth={3} color="#111"/>}
                            </div>
                            <span style={{fontSize:13, fontWeight:500, color:'#111'}}>Tous les jours</span>
                          </div>
                          {['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'].map(jour=>{
                            const actif = filtreJours.has(jour);
                            return (
                              <div key={jour} onClick={()=>toggleFiltreJour(jour)}
                                style={{display:'flex', alignItems:'center', gap:10, padding:'10px 14px', cursor:'pointer', background:actif?'#fffbea':'#fff', borderBottom:'1px solid #f5f5f5'}}
                                onMouseEnter={e=>e.currentTarget.style.background=actif?'#fffbea':'#f9f9f9'}
                                onMouseLeave={e=>e.currentTarget.style.background=actif?'#fffbea':'#fff'}
                              >
                                <div style={{width:16,height:16,borderRadius:4,flexShrink:0,border:'1.5px solid',borderColor:actif?'#E8C547':'#ddd',background:actif?'#E8C547':'#fff',display:'flex',alignItems:'center',justifyContent:'center'}}>
                                  {actif && <Check size={10} strokeWidth={3} color="#111"/>}
                                </div>
                                <span style={{fontSize:13, fontWeight:500, color:'#111'}}>{jour}</span>
                              </div>
                            );
                          })}
                        </div>
                      </>
                    )}
                  </div>
                </div>

                {/* Service préféré */}
                <div style={{flexShrink:0}}>
                  <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 5px'}}>Service préféré</p>
                  <div style={{position:'relative'}}>
                    <button onClick={()=>setShowServiceDropdown(v=>!v)} style={{width:'100%', height:36, border:'1.5px solid #eee', borderRadius:9, background:'#fff', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', padding:'0 12px', fontSize:13, color:'#111', fontWeight:500}}>
                      <span style={{overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', flex:1, textAlign:'left'}}>
                        {filtreServices?.size>0 ? [...filtreServices].map(s=>s==='midi'?'☀️ Midi':'🌙 Soir').join(', ') : 'Tous les services'}
                      </span>
                      <ChevronDown size={14} strokeWidth={2} color="#999" style={{flexShrink:0, marginLeft:6}}/>
                    </button>
                    {showServiceDropdown && (
                      <>
                        <div onClick={()=>setShowServiceDropdown(false)} style={{position:'fixed', inset:0, zIndex:299}}/>
                        <div style={{position:'absolute', top:'calc(100% + 4px)', left:0, right:0, background:'#fff', borderRadius:12, boxShadow:'0 8px 24px rgba(0,0,0,0.12)', zIndex:300, overflow:'hidden', border:'1.5px solid #eee'}}>
                          <div onClick={()=>{setFiltreServices(new Set()); setShowServiceDropdown(false);}}
                            style={{display:'flex', alignItems:'center', gap:10, padding:'10px 14px', cursor:'pointer', background:filtreServices?.size===0?'#fffbea':'#fff', borderBottom:'1px solid #f5f5f5'}}
                            onMouseEnter={e=>e.currentTarget.style.background='#f9f9f9'}
                            onMouseLeave={e=>e.currentTarget.style.background=filtreServices?.size===0?'#fffbea':'#fff'}
                          >
                            <div style={{width:16,height:16,borderRadius:4,flexShrink:0,border:'1.5px solid',borderColor:filtreServices?.size===0?'#E8C547':'#ddd',background:filtreServices?.size===0?'#E8C547':'#fff',display:'flex',alignItems:'center',justifyContent:'center'}}>
                              {filtreServices?.size===0 && <Check size={10} strokeWidth={3} color="#111"/>}
                            </div>
                            <span style={{fontSize:13, fontWeight:500, color:'#111'}}>Tous les services</span>
                          </div>
                          {[{id:'midi',label:'☀️ Midi'},{id:'soir',label:'🌙 Soir'}].map(s=>{
                            const actif = filtreServices?.has(s.id);
                            return (
                              <div key={s.id} onClick={()=>toggleFiltreService(s.id)}
                                style={{display:'flex', alignItems:'center', gap:10, padding:'10px 14px', cursor:'pointer', background:actif?'#fffbea':'#fff', borderBottom:'1px solid #f5f5f5'}}
                                onMouseEnter={e=>e.currentTarget.style.background=actif?'#fffbea':'#f9f9f9'}
                                onMouseLeave={e=>e.currentTarget.style.background=actif?'#fffbea':'#fff'}
                              >
                                <div style={{width:16,height:16,borderRadius:4,flexShrink:0,border:'1.5px solid',borderColor:actif?'#E8C547':'#ddd',background:actif?'#E8C547':'#fff',display:'flex',alignItems:'center',justifyContent:'center'}}>
                                  {actif && <Check size={10} strokeWidth={3} color="#111"/>}
                                </div>
                                <span style={{fontSize:13, fontWeight:500, color:'#111'}}>{s.label}</span>
                              </div>
                            );
                          })}
                        </div>
                      </>
                    )}
                  </div>
                </div>

                {/* Clients absents depuis */}
                <div style={{flexShrink:0}}>
                  <p style={{fontSize:13, fontWeight:700, color:'#111', margin:'0 0 5px'}}>Clients absents depuis</p>
                  <div style={{position:'relative'}}>
                    <select value={filtreAbsentsMois} onChange={e=>setFiltreAbsentsMois(Number(e.target.value))} style={{width:'100%', height:36, border:'1.5px solid #eee', borderRadius:9, padding:'0 32px 0 12px', fontSize:13, color:'#111', fontWeight:500, outline:'none', background:'#fff', cursor:'pointer', appearance:'none', WebkitAppearance:'none'}}>
                      <option value={0}>Indifférent</option>
                      <option value={1}>1 mois</option>
                      <option value={2}>2 mois</option>
                      <option value={3}>3 mois</option>
                      <option value={6}>6 mois</option>
                      <option value={12}>12 mois</option>
                    </select>
                    <ChevronDown size={14} strokeWidth={2} color="#999" style={{position:'absolute', right:12, top:'50%', transform:'translateY(-50%)', pointerEvents:'none'}}/>
                  </div>
                </div>
              </div>

              {/* Résumé de la cible — toujours visible en bas */}
              <div style={{flexShrink:0, marginTop:12, background:'#f9f9f9', borderRadius:12, padding:'12px 14px'}}>
                <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:'0 0 6px'}}>Résumé de la cible</p>
                {(()=>{
                  const total = clientsFiltresComm.length;
                  const h = clientsFiltresComm.filter(c=>c.genre==='Homme').length;
                  const f = clientsFiltresComm.filter(c=>c.genre==='Femme').length;
                  const e = clientsFiltresComm.filter(c=>c.genre==='Entreprise').length;
                  return [
                    {label:'Total ciblé', value:`${total} clients`, bold:true},
                    {label:'Hommes', value:`${h} (${total?Math.round(h/total*100):0}%)`},
                    {label:'Femmes', value:`${f} (${total?Math.round(f/total*100):0}%)`},
                    {label:'Entreprises', value:`${e} (${total?Math.round(e/total*100):0}%)`},
                  ].map((r,i)=>(
                    <div key={i} style={{display:'flex', justifyContent:'space-between', marginBottom:3}}>
                      <span style={{fontSize:12, fontWeight:500, color:'#666'}}>{r.label}</span>
                      <span style={{fontSize:12, fontWeight:r.bold?700:600, color:'#111'}}>{r.value}</span>
                    </div>
                  ));
                })()}
                <button onClick={()=>{ setFiltreGenresComm(new Set()); setFiltreAbsentsMois(0); setFiltreJours(new Set()); setFiltreServices(new Set()); }} style={{width:'100%', marginTop:6, padding:'4px', border:'none', background:'none', fontSize:11, color:'#999', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:4}}>
                  <RotateCcw size={10} strokeWidth={2}/> Réinitialiser
                </button>
              </div>
            </div>

            {/* ─── Colonne 2 — Destinataires ─── */}
            <div style={{background:'#fff', borderRadius:16, boxShadow:'0 1px 4px rgba(0,0,0,0.04)', padding:16, display:'flex', flexDirection:'column', height:'100%', overflow:'hidden'}}>
              <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:'0 0 12px', flexShrink:0}}>
                Destinataires ({clientsFiltresComm.length})
                {selectedComm.length > 0 && <span style={{marginLeft:6, background:'#E8C547', color:'#111', borderRadius:20, padding:'1px 8px', fontSize:11, fontWeight:800}}>{selectedComm.length} sél.</span>}
              </p>

              {/* Recherche */}
              <div style={{position:'relative', marginBottom:10, flexShrink:0}}>
                <Search size={14} strokeWidth={2} color="#999" style={{position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', pointerEvents:'none'}}/>
                <input placeholder="Rechercher un client..." value={commSearch} onChange={e=>setCommSearch(e.target.value)}
                  style={{width:'100%', height:36, border:'1.5px solid #eee', borderRadius:9, padding:'0 12px 0 34px', fontSize:13, outline:'none', boxSizing:'border-box'}}/>
              </div>

              {/* Tout sélectionner */}
              <div onClick={toggleToutSelection} style={{display:'flex', alignItems:'center', gap:10, padding:'8px 12px', borderRadius:9, cursor:'pointer', marginBottom:8, background:'#f9f9f9', flexShrink:0}}>
                <div style={{width:16, height:16, borderRadius:4, border:'1.5px solid', borderColor: tousSelectionnes?'#E8C547':'#ddd', background: tousSelectionnes?'#E8C547':'#fff', display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0}}>
                  {tousSelectionnes && <Check size={10} strokeWidth={3} color="#111"/>}
                </div>
                <span style={{fontWeight:500, fontSize:13, color:'#111', flex:1}}>Tout sélectionner</span>
                <span style={{fontSize:13, color:'#999'}}>{clientsFiltresComm.length}</span>
              </div>

              {/* Liste scrollable */}
              <div style={{flex:1, overflowY:'auto'}}>
                {clientsFiltresComm.map(c => {
                  const estSel = selectedComm.includes(c.id);
                  const isMobileNum = isNumeroMobile(c.tel||'');
                  const disabled = commType==='sms' && !isMobileNum;
                  return (
                    <div key={c.id} onClick={()=>!disabled&&toggleSelectionClient(c.id)} style={{display:'flex', alignItems:'center', gap:10, padding:'8px 10px', borderRadius:8, cursor: disabled?'not-allowed':'pointer', opacity: disabled?0.4:1, marginBottom:2, background: estSel?'#fffbea':'transparent'}}
                      onMouseEnter={e=>{ if(!disabled) e.currentTarget.style.background = estSel?'#fffbea':'#f9f9f9'; }}
                      onMouseLeave={e=>{ e.currentTarget.style.background = estSel?'#fffbea':'transparent'; }}>
                      <div style={{width:16, height:16, borderRadius:4, border:'1.5px solid', flexShrink:0, borderColor: estSel?'#E8C547':'#ddd', background: estSel?'#E8C547':'#fff', display:'flex', alignItems:'center', justifyContent:'center'}}>
                        {estSel && <Check size={10} strokeWidth={3} color="#111"/>}
                      </div>
                      <div style={{width:32, height:32, borderRadius:'50%', flexShrink:0, background: c.genre==='Homme'?'#dbeafe':c.genre==='Femme'?'#fce7f3':'#dcfce7', display:'flex', alignItems:'center', justifyContent:'center', fontSize:11, fontWeight:600, color: c.genre==='Homme'?'#1d4ed8':c.genre==='Femme'?'#be185d':'#15803d'}}>
                        {(c.prenom||c.entreprise||'?')[0]?.toUpperCase()}
                      </div>
                      <div style={{flex:1, minWidth:0}}>
                        <div style={{fontWeight:500, fontSize:13, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}}>{c.genre==='Entreprise'?c.entreprise:`${c.prenom} ${c.nom}`}</div>
                        <div style={{fontSize:11, color:'#999'}}>{c.tel}</div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* ─── Colonne 3 — Message ─── */}
            <div style={{background:'#fff', borderRadius:16, boxShadow:'0 1px 4px rgba(0,0,0,0.04)', padding:16, display:'flex', flexDirection:'column', height:'100%', overflow:'hidden'}}>
              <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:'0 0 12px', flexShrink:0}}>Créer une campagne</p>

              {/* Onglets SMS / Email */}
              <div style={{display:'flex', gap:0, marginBottom:14, flexShrink:0, borderBottom:'2px solid #f0f0f0'}}>
                {[{id:'email',label:'Email',icon:<Mail size={14} strokeWidth={2}/>},{id:'sms',label:'SMS',icon:<MessageSquare size={14} strokeWidth={2}/>}].map(t => (
                  <button key={t.id} onClick={()=>setCommType(t.id)} style={{flex:1, height:36, border:'none', background:'none', fontSize:13, fontWeight:600, cursor:'pointer', color: commType===t.id?'#111':'#999', borderBottom: commType===t.id?'2px solid #E8C547':'2px solid transparent', marginBottom:-2, display:'flex', alignItems:'center', justifyContent:'center', gap:6}}>
                    {t.icon} {t.label}
                  </button>
                ))}
              </div>

              {/* Contenu scrollable */}
              <div style={{flex:1, overflowY:'auto', display:'flex', flexDirection:'column', gap:12}}>
                <div>
                  <label style={{fontSize:13, fontWeight:700, color:'#111', display:'block', marginBottom:5}}>Nom de la campagne</label>
                  <input value={nomCampagne} onChange={e=>setNomCampagne(e.target.value.slice(0,100))} placeholder="Ex: Offre spéciale été – Juin 2026" style={{width:'100%', height:36, border:'1.5px solid #eee', borderRadius:9, padding:'0 12px', fontSize:13, outline:'none', boxSizing:'border-box'}}/>
                  <div style={{textAlign:'right', fontSize:10, color:'#999', marginTop:2}}>{nomCampagne.length}/100</div>
                </div>

                {commType==='email' && (
                  <div>
                    <label style={{fontSize:13, fontWeight:700, color:'#111', display:'block', marginBottom:5}}>Objet</label>
                    <input value={commObjet} onChange={e=>setCommObjet(e.target.value)} placeholder="Objet de l'email..." style={{width:'100%', height:36, border:'1.5px solid #eee', borderRadius:9, padding:'0 12px', fontSize:13, outline:'none', boxSizing:'border-box'}}/>
                  </div>
                )}

                <div>
                  <label style={{fontSize:13, fontWeight:700, color:'#111', display:'block', marginBottom:5}}>Message</label>
                  <textarea
                    value={commType==='sms'?smsMessage:commMessage}
                    onChange={e => { const limit = commType==='sms'?smsLimit:2000; commType==='sms'?setSmsMessage(e.target.value.slice(0,limit)):setCommMessage(e.target.value.slice(0,limit)); }}
                    placeholder="Écrivez votre message..."
                    style={{width:'100%', height:80, border:'1.5px solid #eee', borderRadius:9, padding:'8px 12px', fontSize:13, outline:'none', resize:'none', boxSizing:'border-box', fontFamily:'inherit'}}
                  />
                  {commType==='sms' && (
                    <div style={{display:'flex', justifyContent:'space-between', fontSize:10, color:'#999', marginTop:2}}>
                      <span>{smsMessage.length}/{smsLimit} caractères</span>
                      <span>~0.04€/dest.</span>
                    </div>
                  )}
                </div>

                <div>
                  <p style={{fontSize:12, fontWeight:600, color:'#666', margin:'0 0 5px'}}>Variables disponibles</p>
                  <div style={{display:'flex', flexWrap:'wrap', gap:5}}>
                    {['{prenom}','{nom}','{tel}','{entreprise}','{lien_resa}'].map(v => (
                      <button key={v} onClick={()=>{ const setter=commType==='sms'?setSmsMessage:setCommMessage; const val=commType==='sms'?smsMessage:commMessage; setter(val+v); }} style={{padding:'3px 10px', borderRadius:16, fontSize:11, fontWeight:600, background:'#fffbea', border:'1.5px solid #E8C547', color:'#111', cursor:'pointer'}}>{v}</button>
                    ))}
                  </div>
                </div>
              </div>

              {/* Bouton fixe en bas */}
              <div style={{flexShrink:0, borderTop:'1px solid #eee', paddingTop:10, marginTop:8}}>
                <button onClick={()=>setShowConfirmEnvoi(true)} disabled={selectedComm.length===0||(commType==='sms'?!smsMessage.trim():(!commObjet.trim()||!commMessage.trim()))} style={{width:'100%', height:44, border:'none', borderRadius:10, background: (selectedComm.length>0&&(commType==='sms'?smsMessage.trim():commObjet.trim()&&commMessage.trim()))?'#E8C547':'#f0f0f0', color: (selectedComm.length>0&&(commType==='sms'?smsMessage.trim():commObjet.trim()&&commMessage.trim()))?'#111':'#bbb', fontSize:14, fontWeight:700, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:6}}>
                  <Send size={16} strokeWidth={2}/> Envoyer la campagne
                </button>
                {selectedComm.length>0 && <p style={{textAlign:'center', fontSize:11, color:'#999', margin:'4px 0 0'}}>Envoi immédiat à {selectedComm.length} destinataire{selectedComm.length>1?'s':''}</p>}
              </div>
            </div>
          </div>

          {/* ─── Modal Historique ─── */}
          {showHistorique && (
            <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:4000,display:'flex',alignItems:'center',justifyContent:'center',padding:24,pointerEvents:'all',cursor:'default',touchAction:'none'}} onMouseDown={e=>{e.preventDefault();e.stopPropagation();setShowHistorique(false);}} onClick={(e)=>{if(e.target===e.currentTarget)setShowHistorique(false);}}>
              <div style={{background:'#fff',borderRadius:20,width:'min(600px,calc(100vw-48px))',maxHeight:'80vh',display:'flex',flexDirection:'column',overflow:'hidden'}} onClick={e=>e.stopPropagation()}>
                <div style={{padding:'24px 28px 20px',borderBottom:'1px solid #f0f0f0',display:'flex',alignItems:'center',justifyContent:'space-between',flexShrink:0}}>
                  <h2 style={{margin:0,fontSize:18,fontWeight:800,color:'#111'}}>Historique des envois</h2>
                  <button onClick={()=>setShowHistorique(false)} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666'}}>✕</button>
                </div>
                <div style={{flex:1,overflowY:'auto',padding:'20px 28px'}}>
                  {historiqueEnvois.length===0 ? (
                    <p style={{color:'#bbb',textAlign:'center',padding:'32px 0'}}>Aucun envoi pour l'instant</p>
                  ) : historiqueEnvois.map((h,i) => (
                    <div key={i} style={{display:'flex',alignItems:'center',gap:12,padding:'12px 0',borderBottom:'1px solid #f5f5f5'}}>
                      <div style={{width:36,height:36,borderRadius:8,background:h.type==='sms'?'#f0fdf4':'#eff6ff',display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>
                        {h.type==='sms'?<MessageSquare size={16} color="#16a34a" strokeWidth={2}/>:<Mail size={16} color="#3b82f6" strokeWidth={2}/>}
                      </div>
                      <div style={{flex:1}}>
                        <div style={{fontWeight:700,fontSize:13,color:'#111'}}>{h.type==='sms'?'SMS':'Email'} — {h.nb_destinataires} destinataire{h.nb_destinataires>1?'s':''}</div>
                        <div style={{fontSize:12,color:'#999'}}>{h.objet||h.message?.slice(0,40)||''}</div>
                        <div style={{fontSize:11,color:'#bbb'}}>{new Date(h.created_at).toLocaleDateString('fr-FR',{day:'numeric',month:'long',year:'numeric',hour:'2-digit',minute:'2-digit'})}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* ─── Modal Récap + Confirmation envoi ─── */}
          {showConfirmEnvoi && (
            <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.6)',zIndex:4000,display:'flex',alignItems:'center',justifyContent:'center',padding:24,pointerEvents:'all',cursor:'default',touchAction:'none'}} onMouseDown={e=>{e.preventDefault();e.stopPropagation();setShowConfirmEnvoi(false);}} onClick={(e)=>{if(e.target===e.currentTarget)setShowConfirmEnvoi(false);}}>
              <div style={{background:'#fff',borderRadius:20,width:'min(720px,calc(100vw-48px))',maxHeight:'88vh',display:'flex',flexDirection:'column',boxShadow:'0 32px 80px rgba(0,0,0,0.3)',overflow:'hidden'}} onClick={e=>e.stopPropagation()} onMouseDown={e=>e.stopPropagation()}>

                {/* Header */}
                <div style={{padding:'24px 32px 20px',borderBottom:'1px solid #f0f0f0',display:'flex',alignItems:'center',justifyContent:'space-between',flexShrink:0}}>
                  <div>
                    <h2 style={{margin:0,fontSize:20,fontWeight:800,color:'#111'}}>Récapitulatif de l'envoi</h2>
                    <p style={{margin:'4px 0 0',fontSize:13,color:'#999'}}>Vérifiez avant d'envoyer</p>
                  </div>
                  <button onClick={()=>setShowConfirmEnvoi(false)} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666',display:'flex',alignItems:'center',justifyContent:'center'}}>✕</button>
                </div>

                <div style={{flex:1,overflowY:'auto',padding:'24px 32px'}}>
                  {/* Ligne infos envoi */}
                  <div style={{display:'flex',gap:12,marginBottom:20}}>
                    <div style={{flex:1,background:'#f9f9f9',borderRadius:12,padding:'14px 16px'}}>
                      <p style={{fontSize:11,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:1,margin:'0 0 6px'}}>Destinataires</p>
                      <p style={{fontSize:22,fontWeight:900,color:'#111',margin:'0 0 2px'}}>{selectedComm.length}</p>
                      <p style={{fontSize:12,color:'#666',margin:0}}>
                        {clients.filter(c=>selectedComm.includes(c.id)).slice(0,3).map(c=>c.prenom||c.entreprise).join(', ')}
                        {selectedComm.length>3?` et ${selectedComm.length-3} autres`:''}
                      </p>
                    </div>
                    <div style={{flex:1,background:'#f9f9f9',borderRadius:12,padding:'14px 16px'}}>
                      <p style={{fontSize:11,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:1,margin:'0 0 6px'}}>Type</p>
                      <p style={{fontSize:16,fontWeight:800,color:'#111',margin:'0 0 2px'}}>{commType==='email'?'✉️ Email':'💬 SMS'}</p>
                      <p style={{fontSize:12,color:'#666',margin:0}}>{commType==='sms'?`~${(selectedComm.length*0.04).toFixed(2)}€ estimés`:'Envoi gratuit'}</p>
                    </div>
                    <div style={{flex:2,background:'#f9f9f9',borderRadius:12,padding:'14px 16px'}}>
                      <p style={{fontSize:11,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:1,margin:'0 0 8px'}}>Filtres actifs</p>
                      <div style={{display:'flex',flexWrap:'wrap',gap:6}}>
                        {filtreGenresComm.size>0 && [...filtreGenresComm].map(g=><span key={g} style={{background:'#111',color:'#E8C547',borderRadius:20,padding:'3px 10px',fontSize:12,fontWeight:700}}>{g}s</span>)}
                        {filtreAbsentsMois>0 && <span style={{background:'#111',color:'#E8C547',borderRadius:20,padding:'3px 10px',fontSize:12,fontWeight:700}}>Absents {filtreAbsentsMois}m</span>}
                        {filtreJours?.size>0 && [...filtreJours].map(j=><span key={j} style={{background:'#111',color:'#E8C547',borderRadius:20,padding:'3px 10px',fontSize:12,fontWeight:700}}>{j}</span>)}
                        {filtreServices?.size>0 && [...filtreServices].map(s=><span key={s} style={{background:'#111',color:'#E8C547',borderRadius:20,padding:'3px 10px',fontSize:12,fontWeight:700}}>{s==='midi'?'☀️ Midi':'🌙 Soir'}</span>)}
                        {filtreGenresComm.size===0 && !filtreAbsentsMois && !filtreJours?.size && !filtreServices?.size && <span style={{fontSize:13,color:'#999'}}>Aucun — tous les clients</span>}
                      </div>
                    </div>
                  </div>

                  {/* Aperçu réaliste */}
                  {(()=>{
                    const premier = clients.find(c=>selectedComm.includes(c.id));
                    const replaceVars = (txt) => (txt||'')
                      .replace(/{prenom}/g, premier?.prenom||'Prénom')
                      .replace(/{nom}/g, premier?.nom||'Nom')
                      .replace(/{tel}/g, premier?.tel||'Téléphone')
                      .replace(/{entreprise}/g, premier?.entreprise||'Entreprise')
                      .replace(/{lien_resa}/g, 'https://ted-crm.pages.dev/reserver.html');
                    return (
                      <div style={{border:'1.5px solid #eee',borderRadius:14,overflow:'hidden'}}>
                        {commType==='email' ? (
                          <>
                            <div style={{background:'#f8f8f8',padding:'14px 20px',borderBottom:'1px solid #eee'}}>
                              <div style={{display:'flex',gap:8,marginBottom:6,fontSize:13}}>
                                <span style={{color:'#999',minWidth:60}}>De :</span>
                                <span style={{fontWeight:600,color:'#111'}}>Le TED &lt;com.astegal@gmail.com&gt;</span>
                              </div>
                              <div style={{display:'flex',gap:8,marginBottom:6,fontSize:13}}>
                                <span style={{color:'#999',minWidth:60}}>À :</span>
                                <span style={{fontWeight:600,color:'#111'}}>
                                  {premier?.mail||`${premier?.prenom||''} ${premier?.nom||''}`.trim()||'destinataire'}
                                  {selectedComm.length>1?` + ${selectedComm.length-1} autres`:''}
                                </span>
                              </div>
                              <div style={{display:'flex',gap:8,fontSize:13}}>
                                <span style={{color:'#999',minWidth:60}}>Objet :</span>
                                <span style={{fontWeight:800,color:'#111'}}>{commObjet||'(sans objet)'}</span>
                              </div>
                            </div>
                            <div style={{padding:'24px 28px',minHeight:120,fontSize:15,color:'#333',lineHeight:1.8,whiteSpace:'pre-wrap'}}>
                              {replaceVars(commMessage)||'(message vide)'}
                            </div>
                            <div style={{background:'#f8f8f8',padding:'12px 20px',borderTop:'1px solid #eee',fontSize:12,color:'#999',textAlign:'center'}}>
                              Le TED · Restaurant & Club · 28 Av. des Frères Montgolfier, 69680 Chassieu
                            </div>
                          </>
                        ) : (
                          <div style={{padding:24,background:'#f0f0f0',display:'flex',justifyContent:'flex-end'}}>
                            <div style={{maxWidth:'80%'}}>
                              <div style={{background:'#111',borderRadius:'18px 18px 4px 18px',padding:'14px 18px'}}>
                                <p style={{color:'#fff',fontSize:15,lineHeight:1.6,margin:0,whiteSpace:'pre-wrap'}}>
                                  {replaceVars(smsMessage)||'(message vide)'}
                                </p>
                              </div>
                              <p style={{fontSize:11,color:'#999',textAlign:'right',marginTop:6}}>
                                {(smsMessage||'').length}/160 · LE TED
                              </p>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </div>

                {/* Boutons */}
                <div style={{flexShrink:0,padding:'16px 32px',borderTop:'1px solid #eee',display:'flex',gap:12}}>
                  <button onClick={()=>setShowConfirmEnvoi(false)} style={{flex:1,height:50,border:'1.5px solid #ddd',borderRadius:12,background:'#fff',fontSize:15,fontWeight:600,cursor:'pointer',color:'#666'}}>Modifier</button>
                  <button onClick={async()=>{ setShowConfirmEnvoi(false); if(commType==='email'){ await doSendComm(); }else{ await doSendSms(); } }} style={{flex:2,height:50,border:'none',borderRadius:12,background:'#E8C547',fontSize:15,fontWeight:800,cursor:'pointer',color:'#111',display:'flex',alignItems:'center',justifyContent:'center',gap:8}}>
                    <Send size={18} strokeWidth={2}/> Confirmer l'envoi ({selectedComm.length})
                  </button>
                </div>
              </div>
            </div>
          )}


          {/* ─── Modal Confirmation Email doublons ─── */}

        </div>
        {toast && <Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)} />}
        {notifPrePromptModal}
        {sendingModal && <SendingProgressModal type={sendingModal.type} total={sendingModal.total} done={sendingModal.done} successCount={sendingModal.successCount} onClose={()=>setSendingModal(null)} />}
      </>
    );
  }

  return (
    <div style={{ fontFamily:"'Inter','Segoe UI',Arial,sans-serif", minHeight:"100vh", background:"#f8f8f8", color:"#111" }}>
      {notifResa && (() => { const isMob = window.innerWidth < 768; return (
        <div style={{ position:'fixed', top:16, right:isMob?'auto':20, left:isMob?'50%':'auto', transform:isMob?'translateX(-50%)':'none', background:'rgba(17,17,17,0.92)', backdropFilter:'blur(12px)', WebkitBackdropFilter:'blur(12px)', color:'#fff', borderRadius:16, padding:'14px 16px', zIndex:9999, boxShadow:'0 8px 32px rgba(0,0,0,0.25)', display:'flex', alignItems:'center', gap:12, maxWidth:isMob?'90vw':340, minWidth:280, animation:'slideDownFade 0.3s cubic-bezier(0.34,1.56,0.64,1)', cursor:'pointer', border:'1px solid rgba(255,255,255,0.08)' }}
          onClick={() => { setActiveView('reservations'); setNotifResa(null); }}>
          <div style={{ width:42, height:42, borderRadius:10, background:'#E8C547', display:'flex', alignItems:'center', justifyContent:'center', fontSize:20, flexShrink:0 }}>📅</div>
          <div style={{ flex:1, minWidth:0 }}>
            <p style={{ margin:'0 0 2px', fontWeight:800, fontSize:13, color:'#fff' }}>Nouvelle réservation !</p>
            <p style={{ margin:'0 0 1px', fontSize:13, color:'#E8C547', fontWeight:600, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{notifResa.nom}</p>
            <p style={{ margin:0, fontSize:11, color:'rgba(255,255,255,0.5)' }}>{notifResa.message}</p>
          </div>
          <button onClick={(e) => { e.stopPropagation(); setNotifResa(null); }}
            style={{ background:'rgba(255,255,255,0.1)', border:'none', borderRadius:8, color:'rgba(255,255,255,0.6)', fontSize:14, cursor:'pointer', padding:'6px 8px', flexShrink:0, lineHeight:1, transition:'background 0.15s' }}
            onMouseEnter={e => e.target.style.background='rgba(255,255,255,0.2)'}
            onMouseLeave={e => e.target.style.background='rgba(255,255,255,0.1)'}>✕</button>
        </div>
      ); })()}
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
            <div style={{ display:'flex', alignItems:'center', gap:6 }}>
              {'Notification' in window && Notification.permission !== 'granted' && (
                <button onClick={demanderPermissionNotif} style={{ background:'#E8C547', color:'#111', border:'none', borderRadius:8, width:36, height:36, display:'flex', alignItems:'center', justifyContent:'center', fontSize:18, cursor:'pointer' }}>🔔</button>
              )}
              <button onClick={()=>setShowConfirmDeconnexion(true)} style={{ background:'rgba(255,255,255,0.1)', border:'none', borderRadius:8, width:36, height:36, display:'flex', alignItems:'center', justifyContent:'center', cursor:'pointer', color:'#fff', fontSize:16 }}>🔓</button>
            </div>
          </div>
          {/* Onglets + Recherche — uniquement sur l'onglet Clients */}
          {mobileTab === 'clients' && (
            <>
              <div style={{ padding:'8px 12px 6px', background:'#f5f5f5' }}>
                <div style={{ position:'relative', marginBottom:8 }}>
                  <span style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', color:'#bbb', fontSize:14 }}>🔍</span>
                  <input value={search} onChange={e=>{setSearch(e.target.value);setPage(1)}} placeholder="Rechercher..." style={{ width:'100%', height:44, border:'1.5px solid #eee', borderRadius:12, padding:'0 36px 0 38px', fontSize:14, outline:'none', boxSizing:'border-box', background:'#fff' }} />
                  {search && <button onClick={()=>{setSearch('');setPage(1)}} style={{ position:'absolute', right:10, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', fontSize:16, cursor:'pointer', color:'#aaa' }}>✕</button>}
                </div>
                <div style={{ display:'flex', gap:8, overflowX:'auto', scrollbarWidth:'none', paddingBottom:4 }}>
                  {[
                    { id:'tous', label:'Tous' },
                    { id:'particuliers', label:'Particuliers' },
                    { id:'entreprises', label:'Entreprises' }
                  ].map(tab => (
                    <button key={tab.id} onClick={()=>{setActiveTab(tab.id);setPage(1)}} style={{ height:36, padding:'0 14px', borderRadius:10, fontSize:13, fontWeight:700, border:'none', flexShrink:0, cursor:'pointer', background:activeTab===tab.id?'#111':'#fff', color:activeTab===tab.id?'#fff':'#666' }}>
                      {tab.label}
                    </button>
                  ))}
                </div>
              </div>
            </>
          )}
        </div>
      )}

      {/* ═══ SIDEBAR DESKTOP ═══ */}
      {sidebarDesktop}

      <div style={{ marginLeft: isMobile ? 0 : 120 }}>

      {/* ═══ MOBILE — RÉSERVATIONS INLINE ═══ */}
      {isMobile && mobileTab === 'reservations' && (
        <div style={{ paddingTop:56, overflowX:'hidden', maxWidth:'100vw', width:'100%', background:'#f5f5f5', minHeight:'100vh' }}>
          <ReservationsPage
            inline
            showToast={showToast}
            user={user}
            onResaCountChange={(n) => { setResaAttenteCount(n); updateBadge(n); }}
          />
        </div>
      )}

      {/* ═══ MOBILE CARDS ═══ */}
      {isMobile && mobileTab === 'clients' && (
        <div style={{ paddingTop:146, paddingBottom:'calc(90px + env(safe-area-inset-bottom, 16px))', background:'#f5f5f5', minHeight:'100vh' }}>
          {pageClients.length === 0 && (
            <div style={{ textAlign:'center', padding:'4rem 2rem' }}>
              <div style={{ fontSize:48, marginBottom:12 }}>🔍</div>
              <p style={{ color:'#bbb', fontSize:15 }}>Aucun client trouvé</p>
            </div>
          )}
          <div style={{ background:'#fff', borderRadius:14, margin:'12px 16px', overflow:'hidden', boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
          {pageClients.map((c,i) => {
            const avatarBgM = c.genre==='Homme'?'#dbeafe':c.genre==='Femme'?'#fce7f3':'#dcfce7';
            const avatarColorM = c.genre==='Homme'?'#1d4ed8':c.genre==='Femme'?'#be185d':'#15803d';
            const initialesM = c.genre==='Entreprise'?(c.entreprise||'?').slice(0,2).toUpperCase():`${(c.prenom||'?')[0]}${(c.nom||'')[0]||''}`.toUpperCase();
            return (
            <div key={c.id} onClick={()=>setModalDetailClient(c)}
              style={{ display:'flex', alignItems:'center', gap:12, padding:'12px 16px', borderBottom: i<pageClients.length-1?'1px solid #f5f5f5':'none', cursor:'pointer', background:'#fff' }}
              onTouchStart={e=>e.currentTarget.style.background='#fafafa'}
              onTouchEnd={e=>e.currentTarget.style.background='#fff'}>
              <div style={{ width:40, height:40, borderRadius:'50%', flexShrink:0, background:avatarBgM, display:'flex', alignItems:'center', justifyContent:'center', fontSize:13, fontWeight:800, color:avatarColorM }}>{initialesM}</div>
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontWeight:500, fontSize:14, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                  {c.genre==='Entreprise' ? (c.entreprise||c.nom||'—') : `${c.prenom||''} ${c.nom||''}`}
                </div>
                <div style={{ fontSize:12, color:'#999', marginTop:2 }}>{c.tel||'—'}</div>
              </div>
              <ChevronRight size={16} strokeWidth={2} color="#ddd"/>
            </div>
            );
          })}
          </div>
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
      {!isMobile && (() => {
        const aujourd = new Date().toISOString().split('T')[0];
        const debutMois = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
        const debutMoisDernier = new Date(new Date().getFullYear(), new Date().getMonth()-1, 1);
        const finMoisDernier = new Date(new Date().getFullYear(), new Date().getMonth(), 0);

        const clientsFiltres = clients
          .filter(c => filtreGenreClients==='Tous' || c.genre===filtreGenreClients)
          .filter(c => !rechercheClients ||
            `${c.prenom||''} ${c.nom||''} ${c.tel||''} ${c.mail||''} ${c.entreprise||''} ${c.commentaire||''}`
              .toLowerCase().includes(rechercheClients.toLowerCase()))
          .sort((a,b)=>`${a.prenom||''}${a.nom||''}`.localeCompare(`${b.prenom||''}${b.nom||''}`));

        const topClients = clients.map(c=>({
          ...c,
          nb: resasData.filter(r=>r.client_id===c.id&&r.statut!=='absente'&&r.statut!=='annulee'&&r.statut!=='refusee').length
        })).filter(c=>c.nb>0).sort((a,b)=>b.nb-a.nb).slice(0,3);

        const nbCeMois = clients.filter(c=>c.created_at && new Date(c.created_at)>=debutMois).length;
        const nbMoisDernier = clients.filter(c=>{ if(!c.created_at) return false; const d=new Date(c.created_at); return d>=debutMoisDernier && d<=finMoisDernier; }).length;
        const pctEvol = nbMoisDernier>0 ? Math.round((nbCeMois-nbMoisDernier)/nbMoisDernier*100) : 0;

        return (
        <div style={{minHeight:'100vh', background:'#f5f5f5'}}>

          {/* 1. HEADER — scrolle et disparaît */}
          <div style={{padding:'24px 32px 16px'}}>
            <div style={{display:'flex', alignItems:'center', justifyContent:'space-between'}}>
              <h1 style={{fontSize:28, fontWeight:900, color:'#111', margin:0}}>Clients</h1>
              <div style={{display:'flex', gap:8}}>
                <div style={{position:'relative'}}>
                  <button onClick={()=>setShowExportMenu(v=>!v)} style={{height:34, padding:'0 12px', borderRadius:8, border:'1.5px solid #eee', background:'#fff', fontSize:12, fontWeight:600, cursor:'pointer', color:'#666', display:'flex', alignItems:'center', gap:6}}>
                    <ArrowUpDown size={14} strokeWidth={2} color="#666"/> Import / Export
                  </button>
                  {showExportMenu && (
                    <>
                      <div onClick={()=>setShowExportMenu(false)} style={{position:'fixed',inset:0,zIndex:199,background:'transparent'}}/>
                      <div style={{position:'absolute', right:0, top:'calc(100% + 4px)', background:'#fff', border:'1.5px solid #eee', borderRadius:10, boxShadow:'0 8px 24px rgba(0,0,0,0.1)', zIndex:200, minWidth:180, overflow:'hidden'}}>
                        <button onMouseDown={e=>e.stopPropagation()} onClick={()=>{ const fl = filtreGenreClients==='Tous'?'Tous les clients':filtreGenreClients==='Homme'?'Hommes uniquement':filtreGenreClients==='Femme'?'Femmes uniquement':'Entreprises uniquement'; exportToCSV(clientsFiltres,{filtreLabel:fl,recherche:rechercheClients}); setShowExportMenu(false); }} style={{display:'block',width:'100%',textAlign:'left',padding:'10px 16px',border:'none',background:'none',cursor:'pointer',fontSize:13,borderBottom:'1px solid #f5f5f5'}} onMouseEnter={e=>e.currentTarget.style.background='#f9f9f9'} onMouseLeave={e=>e.currentTarget.style.background='none'}>⬇ Exporter CSV</button>
                        <button onMouseDown={e=>e.stopPropagation()} onClick={()=>{exportToXLSX(clients);setShowExportMenu(false);}} style={{display:'block',width:'100%',textAlign:'left',padding:'10px 16px',border:'none',background:'none',cursor:'pointer',fontSize:13,borderBottom:'1px solid #f5f5f5'}} onMouseEnter={e=>e.currentTarget.style.background='#f9f9f9'} onMouseLeave={e=>e.currentTarget.style.background='none'}>⬇ Exporter Excel</button>
                        <button onMouseDown={e=>e.stopPropagation()} onClick={()=>{setModalImport(true);setShowExportMenu(false);}} style={{display:'block',width:'100%',textAlign:'left',padding:'10px 16px',border:'none',background:'none',cursor:'pointer',fontSize:13}} onMouseEnter={e=>e.currentTarget.style.background='#f9f9f9'} onMouseLeave={e=>e.currentTarget.style.background='none'}>⬆ Importer clients</button>
                      </div>
                    </>
                  )}
                </div>
                <button onClick={()=>setModalCorbeille(true)} style={{height:34, padding:'0 12px', borderRadius:8, border:'1.5px solid #eee', background:'#fff', fontSize:12, fontWeight:600, cursor:'pointer', color:'#666', display:'flex', alignItems:'center', gap:6}}>
                  <Trash2 size={14} strokeWidth={2} color="#666"/> Corbeille
                </button>
              </div>
            </div>
          </div>

          {/* 2. BARRE STICKY */}
          <div style={{position:'sticky', top:0, zIndex:100, background:'#f5f5f5', padding:'10px 32px 14px', borderBottom:'1px solid #eee', boxShadow:'0 2px 12px rgba(0,0,0,0.06)'}}>
            <div style={{display:'flex', alignItems:'center', gap:12}}>
              <div style={{position:'relative', flex:1}}>
                <Search size={16} strokeWidth={2} color="#999" style={{position:'absolute', left:16, top:'50%', transform:'translateY(-50%)', pointerEvents:'none'}}/>
                <input placeholder="Rechercher un client..." value={rechercheClients} onChange={e=>setRechercheClients(e.target.value)}
                  style={{width:'100%', height:36, border:'1.5px solid #eee', borderRadius:10, padding:'0 16px 0 44px', fontSize:13, outline:'none', background:'#fff', boxSizing:'border-box'}}
                  onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                {rechercheClients && <button onClick={()=>setRechercheClients('')} style={{position:'absolute', right:14, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', cursor:'pointer', color:'#aaa', fontSize:16, padding:0}}>✕</button>}
              </div>
              <div style={{display:'flex', gap:6, flexShrink:0}}>
                {[{id:'Tous',label:'Tous'},{id:'Homme',label:'Hommes'},{id:'Femme',label:'Femmes'},{id:'Entreprise',label:'Entreprises'}].map(f=>(
                  <button key={f.id} onClick={()=>setFiltreGenreClients(f.id)} style={{height:36, padding:'0 14px', borderRadius:10, cursor:'pointer', fontSize:13, fontWeight:700, border:'none', background: filtreGenreClients===f.id?'#111':'#fff', color: filtreGenreClients===f.id?'#fff':'#666', boxShadow: filtreGenreClients===f.id?'none':'0 1px 4px rgba(0,0,0,0.06)'}}>{f.label}</button>
                ))}
              </div>
              <button onClick={()=>setModalAdd(true)} style={{height:36, padding:'0 16px', borderRadius:10, border:'none', background:'#E8C547', color:'#111', fontSize:13, fontWeight:800, cursor:'pointer', display:'flex', alignItems:'center', gap:8, flexShrink:0, boxShadow:'0 2px 8px rgba(232,197,71,0.3)'}}>
                <Plus size={16} strokeWidth={2}/> Nouveau client
              </button>
            </div>
          </div>

          {/* 3. STATS + LISTE */}
          <div style={{padding:'20px 32px 32px'}}>

            {/* Blocs stats */}
            <div style={{display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:16, marginBottom:20}}>
              {/* Total clients */}
              <div style={{background:'#fff', borderRadius:16, padding:'14px 18px', display:'flex', alignItems:'center', justifyContent:'space-between'}}>
                <div>
                  <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:'0 0 6px'}}>Total clients</p>
                  <p style={{fontSize:24, fontWeight:900, color:'#111', margin:'0 0 3px'}}>{clients.length}</p>
                  <p style={{fontSize:11, color:'#22c55e', fontWeight:600, margin:0}}>+{nbCeMois} ce mois-ci</p>
                </div>
                <div style={{width:44, height:44, borderRadius:12, background:'#f5f5f5', display:'flex', alignItems:'center', justifyContent:'center'}}>
                  <Users size={20} strokeWidth={2} color="#666"/>
                </div>
              </div>

              {/* Nouveaux ce mois */}
              <div style={{background:'#fff', borderRadius:16, padding:'14px 18px', display:'flex', alignItems:'center', justifyContent:'space-between'}}>
                <div>
                  <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:'0 0 6px'}}>Nouveaux ce mois-ci</p>
                  <p style={{fontSize:24, fontWeight:900, color:'#111', margin:'0 0 3px'}}>{nbCeMois}</p>
                  <p style={{fontSize:11, color:pctEvol>=0?'#22c55e':'#dc2626', fontWeight:600, margin:0}}>{pctEvol>=0?'+':''}{pctEvol}% vs mois dernier</p>
                </div>
                <div style={{width:44, height:44, borderRadius:12, background:'#fffbea', display:'flex', alignItems:'center', justifyContent:'center'}}>
                  <UserPlus size={20} strokeWidth={2} color="#E8C547"/>
                </div>
              </div>

              {/* Top client */}
              <div style={{background:'#fff', borderRadius:16, padding:'14px 18px', display:'flex', alignItems:'center', justifyContent:'space-between'}}>
                {(()=>{
                  const top = clients.map(c=>({...c, nb:resasData.filter(r=>r.client_id===c.id&&r.statut!=='annulee'&&r.statut!=='absente'&&r.statut!=='refusee').length})).filter(c=>c.nb>0).sort((a,b)=>b.nb-a.nb)[0];
                  return (
                    <>
                      <div style={{flex:1, minWidth:0}}>
                        <p style={{fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:'0 0 6px'}}>Top client</p>
                        {top ? (
                          <>
                            <p style={{fontSize:18, fontWeight:900, color:'#111', margin:'0 0 3px', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}}>
                              {top.genre==='Entreprise'?top.entreprise:`${top.prenom} ${top.nom}`}
                            </p>
                            <p style={{fontSize:11, color:'#999', fontWeight:600, margin:0, cursor:'pointer', display:'flex', alignItems:'center', gap:4}} onClick={()=>setShowTopClients(true)}>
                              {top.nb} réservations · <span style={{color:'#E8C547'}}>Voir classement</span>
                            </p>
                          </>
                        ) : <p style={{fontSize:11, color:'#bbb', margin:0}}>Pas encore de données</p>}
                      </div>
                      <div style={{width:44, height:44, borderRadius:12, background:'#fffbea', display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0, cursor:'pointer'}} onClick={()=>setShowTopClients(true)}>
                        <Trophy size={20} strokeWidth={2} color="#E8C547"/>
                      </div>
                    </>
                  );
                })()}
              </div>
            </div>

            {/* Liste clients */}
            <div style={{background:'#fff', borderRadius:16, overflow:'hidden'}}>
              {clientsFiltres.length === 0 ? (
                <div style={{padding:'48px', textAlign:'center', color:'#bbb'}}>
                  <Users size={32} strokeWidth={1.5} color="#ddd" style={{marginBottom:12}}/>
                  <p style={{fontSize:14, margin:0}}>Aucun client trouvé</p>
                </div>
              ) : clientsFiltres.map((c, idx) => {
                const resasClient = resasData.filter(r=>r.client_id===c.id);
                const total = resasClient.filter(r=>r.statut!=='absente'&&r.statut!=='annulee'&&r.statut!=='refusee').length;
                const derniereVisite = resasClient.filter(r=>r.date<=aujourd&&(r.statut==='venue'||r.statut==='confirmee')).sort((a,b)=>b.date.localeCompare(a.date))[0];
                const prochaineResa = resasClient.filter(r=>r.date>aujourd&&(r.statut==='confirmee'||r.statut==='attente')).sort((a,b)=>a.date.localeCompare(b.date))[0];
                const avatarBg = c.genre==='Homme'?'#dbeafe':c.genre==='Femme'?'#fce7f3':'#dcfce7';
                const avatarColor = c.genre==='Homme'?'#1d4ed8':c.genre==='Femme'?'#be185d':'#15803d';
                const initiales = c.genre==='Entreprise'?(c.entreprise||'?').slice(0,2).toUpperCase():`${(c.prenom||'?')[0]}${(c.nom||'')[0]||''}`.toUpperCase();
                return (
                  <div key={c.id} onClick={()=>setModalDetailClient(c)}
                    style={{display:'flex', alignItems:'center', gap:12, padding:'12px 20px', borderBottom: idx<clientsFiltres.length-1?'1px solid #f5f5f5':'none', cursor:'pointer'}}
                    onMouseEnter={e=>e.currentTarget.style.background='#fafafa'}
                    onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                    <div style={{width:36, height:36, borderRadius:'50%', flexShrink:0, background:avatarBg, display:'flex', alignItems:'center', justifyContent:'center', fontSize:12, fontWeight:800, color:avatarColor}}>{initiales}</div>
                    <div style={{minWidth:180, flex:'0 0 180px'}}>
                      <div style={{fontWeight:700, fontSize:14, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}}>{c.genre==='Entreprise'?c.entreprise:`${c.prenom||''} ${c.nom||''}`}</div>
                      <div style={{fontSize:12, color:'#999', marginTop:1}}>{c.tel||'—'}</div>
                    </div>
                    <div style={{display:'flex', alignItems:'center', gap:8, flex:'0 0 110px'}}>
                      <CalendarDays size={15} strokeWidth={2} color="#ccc"/>
                      <div>
                        <span style={{fontSize:14, fontWeight:800, color:'#111'}}>{total}</span>
                        <div style={{fontSize:11, color:'#999'}}>réservations</div>
                      </div>
                    </div>
                    <div style={{flex:1, minWidth:0}}>
                      <div style={{fontSize:11, color:'#999', marginBottom:2}}>Dernière visite</div>
                      <div style={{fontSize:13, fontWeight:600, color: derniereVisite?'#111':'#ccc', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}}>
                        {derniereVisite ? new Date(derniereVisite.date+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long',year:'numeric'}) : 'Jamais'}
                      </div>
                    </div>
                    <div style={{flex:1, minWidth:0}}>
                      <div style={{fontSize:11, color:'#999', marginBottom:2}}>Prochaine réservation</div>
                      <div style={{fontSize:13, fontWeight:600, color: prochaineResa?'#16a34a':'#ccc', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}}>
                        {prochaineResa ? `${new Date(prochaineResa.date+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long'})}${prochaineResa.heure?` à ${prochaineResa.heure}`:''}` : '—'}
                      </div>
                    </div>
                    <ChevronRight size={14} strokeWidth={2} color="#ddd" style={{flexShrink:0}}/>
                  </div>
                );
              })}
            </div>
          </div>

        </div>
        );
      })()}

      {/* Barre nav fixe mobile */}
      {isMobile && (
        <>
          <div style={{ position:'fixed', bottom:0, left:0, right:0, background:'#fff', borderTop:'1px solid #eee', display:'flex', alignItems:'center', zIndex:1000, paddingTop:10, paddingBottom:'env(safe-area-inset-bottom, 16px)', minHeight:70 }}>
            {/* Clients */}
            <button onClick={()=>setMobileTab('clients')} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:5, border:'none', background:'none', cursor:'pointer', color: mobileTab==='clients' ? '#111' : '#aaa', paddingBottom:4 }}>
              <svg width="28" height="28" fill="none" viewBox="0 0 24 24">
                <path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z" fill="currentColor"/>
              </svg>
              <span style={{ fontSize:12, fontWeight: mobileTab==='clients' ? 700 : 500 }}>Clients</span>
            </button>
            {/* Réservations */}
            <button onClick={()=>setMobileTab('reservations')} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:5, border:'none', background:'none', cursor:'pointer', color: mobileTab==='reservations' ? '#111' : '#aaa', paddingBottom:4, position:'relative' }}>
              <div style={{ position:'relative' }}>
                <svg width="28" height="28" fill="none" viewBox="0 0 24 24">
                  <path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z" fill="currentColor"/>
                </svg>
                {resaAttenteCount > 0 && (
                  <div style={{ position:'absolute', top:-8, right:-10, background:'#dc2626', color:'#fff', borderRadius:'99px', minWidth:22, height:22, display:'flex', alignItems:'center', justifyContent:'center', fontSize:12, fontWeight:800, border:'2px solid #fff', padding:'0 5px' }}>{resaAttenteCount}</div>
                )}
              </div>
              <span style={{ fontSize:12, fontWeight: mobileTab==='reservations' ? 700 : 500 }}>Réservations</span>
            </button>
          </div>
          {/* Bouton flottant + */}
          {mobileTab === 'reservations' ? (
            <div style={{ position:'fixed', bottom:'calc(85px + env(safe-area-inset-bottom))', right:16, zIndex:1000 }}>
              <button
                className="btn-pulse"
                onClick={()=>setShowAddResa(true)}
                onMouseEnter={e => e.currentTarget.style.transform='scale(1.05)'}
                onMouseLeave={e => e.currentTarget.style.transform='scale(1)'}
                onTouchStart={e => e.currentTarget.style.transform='scale(0.95)'}
                onTouchEnd={e => e.currentTarget.style.transform='scale(1)'}
                style={{ background:'#E8C547', border:'3px solid #fff', borderRadius:50, padding:'14px 20px', fontSize:13, fontWeight:800, cursor:'pointer', color:'#111', whiteSpace:'nowrap', transition:'transform 0.15s ease' }}
              >+ Nouvelle réservation</button>
            </div>
          ) : (
            <div style={{ position:'fixed', bottom:'calc(85px + env(safe-area-inset-bottom))', right:16, zIndex:1000 }}>
              <button
                className="btn-pulse"
                onClick={()=>setModalAdd(true)}
                onMouseEnter={e => e.currentTarget.style.transform='scale(1.05)'}
                onMouseLeave={e => e.currentTarget.style.transform='scale(1)'}
                onTouchStart={e => e.currentTarget.style.transform='scale(0.95)'}
                onTouchEnd={e => e.currentTarget.style.transform='scale(1)'}
                style={{ background:'#E8C547', border:'3px solid #fff', borderRadius:50, padding:'14px 20px', fontSize:13, fontWeight:800, cursor:'pointer', color:'#111', whiteSpace:'nowrap', transition:'transform 0.15s ease' }}
              >+ Nouveau client</button>
            </div>
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
      {showAddResa && <AddResaModal onClose={()=>setShowAddResa(false)} onSaved={()=>{ loadResaCount(); loadClients(); }} showToast={showToast} user={user} onViewClient={(c)=>{ setFicheClientReadOnly(true); setModalDetailClient(c); }} reservations={resasData} />}
      {modalDetailClient && (() => {
        const c = modalDetailClient;
        const s = statsClients[c.id] || { total:0, noshow:0, derniereVisite:null };
        const nomAffiche = c.genre === 'Entreprise' ? (c.entreprise || c.nom || '—') : `${c.prenom||''} ${c.nom||''}`.trim() || '—';
        const fermerFiche = () => { setModalDetailClient(null); setFicheClientReadOnly(false); };
        const ficheBody = (
            <div style={{ display:'flex', flexDirection:'column', gap:12 }}>
              <div style={{ display:'flex', gap:8, alignItems:'center' }}>
                <span style={badge(c.genre)}>{c.genre}</span>
                {c.genre === 'Entreprise' && c.nom && <span style={{ fontSize:13, color:'#888' }}>{c.nom} {c.prenom}</span>}
              </div>
              {c.tel && <div style={{ fontSize:14, color:'#333' }}>📞 <a href={`tel:${c.tel}`} style={{ color:'#111', textDecoration:'none', fontWeight:600 }}>{c.tel}</a></div>}
              {c.mail && <div style={{ fontSize:13, color:'#3b82f6' }}>✉️ <a href={`mailto:${c.mail}`} style={{ color:'#3b82f6', textDecoration:'none' }}>{c.mail}</a></div>}
              {c.tel && (
                <div style={{ display:'flex', gap:8, marginTop:4, marginBottom:4 }}>
                  <a href={`sms:${c.tel}`} style={{ flex:1, height:44, background:'#fff', border:'1.5px solid #ddd', borderRadius:10, fontSize:13, fontWeight:700, cursor:'pointer', color:'#111', display:'flex', alignItems:'center', justifyContent:'center', textDecoration:'none' }}>💬 SMS</a>
                  <a href={`tel:${c.tel}`} style={{ flex:1, height:44, background:'#111', border:'none', borderRadius:10, fontSize:13, fontWeight:700, cursor:'pointer', color:'#fff', display:'flex', alignItems:'center', justifyContent:'center', textDecoration:'none' }}>📞 Appeler</a>
                </div>
              )}
              {c.created_at && <div style={{ fontSize:12, color:'#999' }}>📋 Client depuis le {formatDate(c.created_at)}</div>}
              {c.commentaire && <div style={{ fontSize:13, color:'#555', background:'#f9f9f9', borderRadius:8, padding:'10px 12px', fontStyle:'italic' }}>"{c.commentaire}"</div>}
              <div style={{ background:'#f9f9f9', borderRadius:10, padding:'12px 16px', display:'flex', gap:16 }}>
                <div style={{ textAlign:'center', flex:1 }}><div style={{ fontSize:22, fontWeight:800, color:'#111' }}>{s.total}</div><div style={{ fontSize:11, color:'#999', textTransform:'uppercase', letterSpacing:0.5 }}>Résa total</div></div>
                <div style={{ textAlign:'center', flex:1 }}><div style={{ fontSize:22, fontWeight:800, color: s.noshow > 0 ? '#dc2626' : '#111' }}>{s.noshow}</div><div style={{ fontSize:11, color:'#999', textTransform:'uppercase', letterSpacing:0.5 }}>No-show</div></div>
                <div style={{ textAlign:'center', flex:2 }}><div style={{ fontSize:13, fontWeight:700, color:'#111' }}>{s.derniereVisite ? new Date(s.derniereVisite+'T12:00:00').toLocaleDateString('fr-FR',{day:'numeric',month:'long',year:'numeric'}) : 'Jamais'}</div><div style={{ fontSize:11, color:'#999', textTransform:'uppercase', letterSpacing:0.5 }}>Dernière visite</div></div>
              </div>
              {(() => {
                const aujourd2 = new Date();
                const il6Mois = new Date(); il6Mois.setMonth(il6Mois.getMonth() - 6);
                const il6MoisStr = il6Mois.toISOString().split('T')[0];
                const periodeLabel = `${il6Mois.toLocaleDateString('fr-FR',{month:'short',year:'numeric'})} — ${aujourd2.toLocaleDateString('fr-FR',{month:'short',year:'numeric'})}`;
                const jours = ['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
                const compteJours = {};
                resasData.filter(r => r.client_id === c.id && (r.statut === 'confirmee' || r.statut === 'venue') && r.date >= il6MoisStr).forEach(r => {
                  const jour = jours[new Date(r.date+'T12:00:00').getDay()];
                  const service = r.service === 'midi' ? 'Midi' : 'Soir';
                  const key = `${jour} ${service}`;
                  compteJours[key] = (compteJours[key] || 0) + 1;
                });
                const topJoursClient = Object.entries(compteJours).sort((a,b) => b[1]-a[1]).slice(0,3);
                return (
                  <div style={{ background:'#f9f9f9', borderRadius:10, padding:'12px 14px' }}>
                    <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:8 }}>
                      <p style={{ fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:0 }}>🏆 Jours favoris</p>
                      <span style={{ fontSize:10, color:'#bbb' }}>↻ {periodeLabel}</span>
                    </div>
                    {topJoursClient.length === 0
                      ? <p style={{ fontSize:12, color:'#bbb', margin:0 }}>Pas de données sur cette période</p>
                      : topJoursClient.map(([label, count], i) => (
                        <div key={i} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:4 }}>
                          <span style={{ fontSize:13, color:'#444' }}>{i===0?'🥇':i===1?'🥈':'🥉'} {label}</span>
                          <span style={{ fontSize:13, fontWeight:700, color:'#111' }}>{count} résa</span>
                        </div>
                      ))
                    }
                  </div>
                );
              })()}
            </div>
        );
        const ficheFooter = (
          <div style={{ display:'flex', flexDirection:'column', gap:8, width:'100%' }}>
            {!ficheClientReadOnly && (
              <div style={{ display:'flex', gap:8 }}>
                <button onClick={()=>{ fermerFiche(); setModalDelete(c); }} style={{ flex:1, height:44, border:'1.5px solid #dc2626', borderRadius:10, background:'#fef2f2', color:'#dc2626', fontSize:14, fontWeight:700, cursor:'pointer' }}>🗑️ Supprimer le client</button>
                <button onClick={()=>{ setModalEdit(c); }} style={{ flex:2, height:44, border:'none', borderRadius:10, background:'#E8C547', color:'#111', fontSize:14, fontWeight:700, cursor:'pointer' }}>✏️ Modifier le client</button>
              </div>
            )}
            <button onClick={fermerFiche} style={{ width:'100%', height:44, background:'#fff', border:'1.5px solid #ddd', borderRadius:10, fontSize:14, fontWeight:600, cursor:'pointer', color:'#666' }}>Fermer</button>
          </div>
        );
        if (isMobile) return (
          <div style={{ position:'fixed', inset:0, background:'#f5f5f5', zIndex:6000, display:'flex', flexDirection:'column' }}>
            <div style={{ background:'#f5f5f5', padding:'16px 16px 12px', paddingTop:'calc(16px + env(safe-area-inset-top))', borderBottom:'1px solid #eee', flexShrink:0, display:'flex', alignItems:'center', justifyContent:'space-between', gap:12 }}>
              <h2 style={{ color:'#111', margin:0, fontSize:18, fontWeight:900, flex:1, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{nomAffiche}</h2>
              <button onClick={fermerFiche} style={{ height:36, padding:'0 14px', borderRadius:10, background:'#111', border:'none', fontSize:13, fontWeight:700, color:'#fff', display:'flex', alignItems:'center', gap:6, flexShrink:0, cursor:'pointer', touchAction:'manipulation' }}>
                <ArrowLeft size={14} strokeWidth={2} color="#fff"/> Retour
              </button>
            </div>
            <div style={{ flex:1, overflowY:'auto', padding:'12px 16px', WebkitOverflowScrolling:'touch' }}>{ficheBody}</div>
            <div style={{ background:'#fff', padding:'12px 16px', paddingBottom:'calc(12px + env(safe-area-inset-bottom))', borderTop:'1px solid #eee', flexShrink:0 }}>{ficheFooter}</div>
          </div>
        );
        // Desktop — full page overlay
        const aujourd = new Date().toISOString().split('T')[0];
        const resasClient = resasData.filter(r => r.client_id === c.id);
        const totalResas = resasClient.filter(r => r.statut !== 'absente' && r.statut !== 'annulee' && r.statut !== 'refusee').length;
        const noshowResas = resasClient.filter(r => r.statut === 'absente').length;
        const pct = totalResas > 0 ? Math.round(noshowResas / totalResas * 100) : 0;
        const derniereVisite = resasClient.filter(r => r.date <= aujourd && (r.statut === 'venue' || r.statut === 'confirmee')).sort((a,b) => b.date.localeCompare(a.date))[0];
        const prochaineResa = resasClient.filter(r => r.date > aujourd && (r.statut === 'confirmee' || r.statut === 'attente')).sort((a,b) => a.date.localeCompare(b.date))[0];
        const derniereVisiteIlYA = derniereVisite ? Math.floor((new Date() - new Date(derniereVisite.date+'T12:00:00')) / (1000*60*60*24)) : null;
        const createdAtLabel = c.created_at ? new Date(c.created_at).toLocaleDateString('fr-FR',{day:'numeric',month:'long',year:'numeric'}) : '';
        const avatarBg = c.genre==='Homme'?'#dbeafe':c.genre==='Femme'?'#fce7f3':'#dcfce7';
        const avatarColor = c.genre==='Homme'?'#1d4ed8':c.genre==='Femme'?'#be185d':'#15803d';
        const statutColors2 = {confirmee:{bg:'#dcfce7',color:'#16a34a',label:'Confirmée'},attente:{bg:'#fef9c3',color:'#ca8a04',label:'En attente'},venue:{bg:'#d1fae5',color:'#059669',label:'Venue'},absente:{bg:'#fee2e2',color:'#dc2626',label:'No-show'},annulee:{bg:'#f3f4f6',color:'#6b7280',label:'Annulée'}};
        const joursSemaine2=['Dimanche','Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi'];
        const joursAbr=['DIM','LUN','MAR','MER','JEU','VEN','SAM'];
        const il6MoisStr2 = new Date(Date.now()-180*24*60*60*1000).toISOString().split('T')[0];
        const resasFav = resasData.filter(r => r.client_id===c.id && (r.statut==='confirmee'||r.statut==='venue') && r.date>=il6MoisStr2);
        const compteJoursFav = {};
        resasFav.forEach(r => { const j = joursSemaine2[new Date(r.date+'T12:00:00').getDay()]; const service = r.service==='midi'?'Midi':'Soir'; const key=`${j}|${service}`; compteJoursFav[key]=(compteJoursFav[key]||0)+1; });
        const top3Jours = Object.entries(compteJoursFav).sort((a,b)=>b[1]-a[1]).slice(0,3);

        return (
          <div style={{ position:'fixed', inset:0, background:'#f5f5f5', zIndex:500, overflowY:'auto', marginLeft:120 }}>
            <div style={{ maxWidth:1100, margin:'0 auto', padding:'24px 32px' }}>

              {/* Header */}
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:20 }}>
                <div style={{ display:'flex', alignItems:'center', gap:12 }}>
                  <h1 style={{ margin:0, fontSize:32, fontWeight:900, color:'#111' }}>{nomAffiche}</h1>
                  {!ficheClientReadOnly && (
                    <button onClick={()=>setModalDelete(c)}
                      style={{ height:38, padding:'0 14px', borderRadius:10, border:'1.5px solid #ddd', background:'#f5f5f5', fontSize:13, fontWeight:600, cursor:'pointer', color:'#999', display:'flex', alignItems:'center', gap:6, flexShrink:0 }}
                      onMouseEnter={e=>{ e.currentTarget.style.background='#fee2e2'; e.currentTarget.style.borderColor='#fca5a5'; e.currentTarget.style.color='#dc2626'; }}
                      onMouseLeave={e=>{ e.currentTarget.style.background='#f5f5f5'; e.currentTarget.style.borderColor='#ddd'; e.currentTarget.style.color='#999'; }}>
                      <Trash2 size={14} strokeWidth={2} color="currentColor"/> Supprimer
                    </button>
                  )}
                </div>
                <button onClick={fermerFiche}
                  style={{ height:38, padding:'0 16px', borderRadius:10, background:'#111', border:'none', fontSize:13, fontWeight:700, cursor:'pointer', color:'#fff', display:'flex', alignItems:'center', gap:6, flexShrink:0 }}
                  onMouseEnter={e=>e.currentTarget.style.background='#333'}
                  onMouseLeave={e=>e.currentTarget.style.background='#111'}>
                  <ArrowLeft size={16} strokeWidth={2} color="#fff"/> Retour
                </button>
              </div>

              {/* Infos + actions */}
              <div style={{ background:'#fff', borderRadius:16, padding:'20px 24px', marginBottom:16, display:'flex', alignItems:'center', gap:24, flexWrap:'wrap' }}>
                <div style={{ width:72, height:72, borderRadius:'50%', flexShrink:0, background:avatarBg, display:'flex', alignItems:'center', justifyContent:'center', fontSize:24, fontWeight:900, color:avatarColor }}>
                  {(((c.prenom||c.entreprise||'?')[0])+(c.nom||'')[0]||'').toUpperCase()}
                </div>
                <div style={{ flex:1, display:'flex', flexDirection:'column', gap:6 }}>
                  {c.tel && <div style={{ display:'flex', alignItems:'center', gap:10 }}><Phone size={16} strokeWidth={2} color="#666" /><span style={{ fontSize:16, fontWeight:600, color:'#111' }}>{c.tel}</span></div>}
                  {c.mail && <div style={{ display:'flex', alignItems:'center', gap:10 }}><Mail size={16} strokeWidth={2} color="#666" /><span style={{ fontSize:15, color:'#3b82f6' }}>{c.mail}</span></div>}
                  {createdAtLabel && <div style={{ display:'flex', alignItems:'center', gap:10 }}><User size={16} strokeWidth={2} color="#666" /><span style={{ fontSize:14, color:'#999' }}>Client depuis le {createdAtLabel}</span></div>}
                </div>
                <div style={{ display:'flex', gap:12 }}>
                  {c.tel && <a href={`tel:${c.tel}`} style={{ display:'flex', flexDirection:'column', alignItems:'center', gap:6, padding:'12px 20px', borderRadius:12, border:'1.5px solid #eee', background:'#fff', cursor:'pointer', minWidth:80, color:'#111', textDecoration:'none', fontSize:13, fontWeight:600 }}><Phone size={20} strokeWidth={2}/>Appeler</a>}
                  {c.tel && <a href={`sms:${c.tel}`} style={{ display:'flex', flexDirection:'column', alignItems:'center', gap:6, padding:'12px 20px', borderRadius:12, border:'1.5px solid #eee', background:'#fff', cursor:'pointer', minWidth:80, color:'#111', textDecoration:'none', fontSize:13, fontWeight:600 }}><MessageSquare size={20} strokeWidth={2}/>SMS</a>}
                  {!ficheClientReadOnly && <button onClick={()=>{ setModalEdit(c); }} style={{ display:'flex', flexDirection:'column', alignItems:'center', gap:6, padding:'12px 20px', borderRadius:12, border:'1.5px solid #eee', background:'#fff', cursor:'pointer', minWidth:80, color:'#111', fontSize:13, fontWeight:600 }}><Pencil size={20} strokeWidth={2}/>Modifier</button>}
                </div>
              </div>

              {/* 4 blocs stats */}
              <div style={{ display:'grid', gridTemplateColumns:'repeat(4,1fr)', gap:16, marginBottom:16 }}>
                {[
                  { icon:<CalendarDays size={20} strokeWidth={2} color="#E8C547"/>, bg:'#fffbea', label:'RÉSA TOTALES', value:totalResas, sub:createdAtLabel?`Depuis le ${createdAtLabel}`:'' },
                  { icon:<UserX size={20} strokeWidth={2} color="#ef4444"/>, bg:'#fef2f2', label:'NO-SHOW', value:noshowResas, sub:`${pct}% des résa` },
                  { icon:<Clock size={20} strokeWidth={2} color="#3b82f6"/>, bg:'#eff6ff', label:'DERNIÈRE VISITE', value:derniereVisite?new Date(derniereVisite.date+'T12:00:00').toLocaleDateString('fr-FR',{day:'numeric',month:'long',year:'numeric'}):'Jamais', sub:derniereVisiteIlYA!==null?`Il y a ${derniereVisiteIlYA} jours`:'' },
                  { icon:<CalendarDays size={20} strokeWidth={2} color="#22c55e"/>, bg:'#f0fdf4', label:'PROCHAINE RÉSA', value:prochaineResa?new Date(prochaineResa.date+'T12:00:00').toLocaleDateString('fr-FR',{day:'numeric',month:'long',year:'numeric'}):'Aucune', sub:prochaineResa?`Dans ${Math.ceil((new Date(prochaineResa.date+'T12:00:00')-new Date())/(1000*60*60*24))}j à ${prochaineResa.heure}`:'' }
                ].map((stat,i)=>(
                  <div key={i} style={{ background:'#fff', borderRadius:16, padding:'16px 20px' }}>
                    <div style={{ width:36, height:36, borderRadius:10, background:stat.bg, display:'flex', alignItems:'center', justifyContent:'center', marginBottom:10 }}>{stat.icon}</div>
                    <p style={{ fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 4px' }}>{stat.label}</p>
                    <p style={{ fontSize:18, fontWeight:900, color:'#111', margin:'0 0 2px' }}>{stat.value}</p>
                    <p style={{ fontSize:11, color:'#999', margin:0 }}>{stat.sub}</p>
                  </div>
                ))}
              </div>

              {/* Commentaire */}
              {c.commentaire && (
                <div style={{ background:'#fff', borderRadius:16, padding:'20px 24px', marginBottom:16, border:'1.5px solid #f0f0f0' }}>
                  <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:10 }}>
                    <MessageSquare size={18} strokeWidth={2} color="#111"/>
                    <h3 style={{ margin:0, fontSize:15, fontWeight:800, color:'#111' }}>Commentaire</h3>
                  </div>
                  <p style={{ margin:0, fontSize:14, color:'#444', lineHeight:1.7, whiteSpace:'pre-wrap' }}>{c.commentaire}</p>
                </div>
              )}
              {!c.commentaire && !ficheClientReadOnly && (
                <button onClick={()=>setModalEdit(c)} style={{ width:'100%', padding:'12px', marginBottom:16, border:'1.5px dashed #ddd', borderRadius:12, background:'transparent', cursor:'pointer', fontSize:13, color:'#999', display:'flex', alignItems:'center', justifyContent:'center', gap:6 }}>
                  <Plus size={14} strokeWidth={2} color="#999"/> Ajouter un commentaire
                </button>
              )}

              {/* Grille historique + jours favoris */}
              <div style={{ display:'grid', gridTemplateColumns:'1fr 320px', gap:16, alignItems:'stretch' }}>
                <div style={{ background:'#fff', borderRadius:16, padding:'20px 24px', display:'flex', flexDirection:'column', maxHeight:340 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:14, flexShrink:0 }}>
                    <CalendarDays size={18} strokeWidth={2} color="#111" />
                    <h3 style={{ margin:0, fontSize:15, fontWeight:800, color:'#111' }}>Historique des réservations</h3>
                  </div>
                  <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr 1fr', padding:'6px 0', borderBottom:'2px solid #f0f0f0', marginBottom:4, flexShrink:0 }}>
                    {['DATE','SERVICE','COUVERTS','STATUT'].map(h=>(
                      <span key={h} style={{ fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5 }}>{h}</span>
                    ))}
                  </div>
                  <div style={{ overflowY:'auto', flex:1 }}>
                    {resasData.filter(r=>r.client_id===c.id).sort((a,b)=>b.date.localeCompare(a.date)).map(r=>{
                      const sc = statutColors2[r.statut] || statutColors2.confirmee;
                      return (
                        <div key={r.id} style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr 1fr', alignItems:'center', padding:'10px 0', borderBottom:'1px solid #f5f5f5' }}>
                          <div>
                            <div style={{ fontWeight:600, fontSize:13, color:'#111' }}>{new Date(r.date+'T12:00:00').toLocaleDateString('fr-FR',{day:'numeric',month:'short',year:'numeric'})}</div>
                            <div style={{ fontSize:11, color:'#999' }}>{r.heure}</div>
                          </div>
                          <div style={{ display:'flex', alignItems:'center', gap:4, fontSize:13, color:'#444' }}>
                            {r.service==='midi'?<><Sun size={13} strokeWidth={2} color="#E8C547"/> Midi</>:<><Moon size={13} strokeWidth={2} color="#666"/> Soir</>}
                          </div>
                          <div style={{ fontSize:13, color:'#444' }}>{r.nb_personnes ? `${r.nb_personnes} pers.` : '—'}</div>
                          <div><span style={{ background:sc.bg, color:sc.color, borderRadius:20, padding:'3px 8px', fontSize:11, fontWeight:700 }}>{sc.label}</span></div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div style={{ background:'#fff', borderRadius:16, padding:'20px 24px', display:'flex', flexDirection:'column', maxHeight:340 }}>
                  <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:6, flexShrink:0 }}>
                    <Star size={16} strokeWidth={2} color="#111" />
                    <h3 style={{ margin:0, fontSize:15, fontWeight:800, color:'#111' }}>Jours favoris</h3>
                  </div>
                  <p style={{ fontSize:11, color:'#999', margin:'0 0 14px', flexShrink:0 }}>Basé sur les 6 derniers mois</p>
                  <div style={{ flex:1 }}>
                    {top3Jours.length > 0 ? top3Jours.map(([key,count],i)=>{
                      const [jour, service] = key.split('|');
                      const abr = joursAbr[joursSemaine2.indexOf(jour)];
                      return (
                        <div key={key} style={{ display:'flex', alignItems:'center', gap:12, padding:'8px 0', borderBottom:i<top3Jours.length-1?'1px solid #f5f5f5':'none' }}>
                          <div style={{ width:44, height:44, borderRadius:8, flexShrink:0, background:'#fffbea', border:'1.5px solid #E8C547', display:'flex', alignItems:'center', justifyContent:'center', fontSize:11, fontWeight:900, color:'#E8C547' }}>
                            {abr}
                          </div>
                          <div style={{ flex:1 }}>
                            <div style={{ fontWeight:700, fontSize:14, color:'#111' }}>{jour}</div>
                            <div style={{ fontSize:12, color:'#999', display:'flex', alignItems:'center', gap:4 }}>
                              {service==='Midi' ? '☀️' : '🌙'} {service} · {count} résa
                            </div>
                          </div>
                        </div>
                      );
                    }) : <p style={{ fontSize:13, color:'#bbb', margin:0 }}>Pas encore de données</p>}
                  </div>
                </div>
              </div>
            </div>
          </div>
        );
      })()}
      {showConfirmDeconnexion && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:9000, display:'flex', alignItems:'center', justifyContent:'center', pointerEvents:'all', cursor:'default', touchAction:'none' }} onMouseDown={e=>{e.preventDefault();e.stopPropagation();setShowConfirmDeconnexion(false);}} onClick={(e)=>{if(e.target===e.currentTarget)setShowConfirmDeconnexion(false);}}>
          <div style={{ background:'#fff', borderRadius:16, padding:'28px 24px', maxWidth:320, width:'90%', textAlign:'center', boxShadow:'0 20px 60px rgba(0,0,0,0.2)' }} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()}>
            <h3 style={{ margin:'0 0 8px', fontSize:17, fontWeight:800 }}>Se déconnecter ?</h3>
            <p style={{ margin:'0 0 20px', fontSize:14, color:'#666' }}>Vous devrez vous reconnecter pour accéder au CRM.</p>
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={()=>setShowConfirmDeconnexion(false)} style={{ flex:1, height:44, border:'1.5px solid #ddd', borderRadius:10, background:'#fff', fontSize:14, cursor:'pointer', color:'#666' }}>Annuler</button>
              <button onClick={()=>{ supabase.auth.signOut(); setShowConfirmDeconnexion(false); }} style={{ flex:1, height:44, border:'none', borderRadius:10, background:'#111', fontSize:14, fontWeight:800, cursor:'pointer', color:'#fff' }}>Se déconnecter</button>
            </div>
          </div>
        </div>
      )}
      </div>{/* end marginLeft wrapper */}

      {modalAdd && (() => {
        const fermerAdd = () => {
          const aDesDonnees = addClientForm.prenom || addClientForm.nom || addClientForm.tel;
          if (aDesDonnees) { setShowConfirmQuitterClient(true); }
          else { setModalAdd(false); setAddClientForm({}); }
        };
        const valide = addClientForm.genre === 'Entreprise'
          ? (addClientForm.tel||'').replace(/\s/g,'').length >= 10 && addClientForm.entreprise?.trim() && (addClientForm.mail||'').includes('@')
          : (addClientForm.tel||'').replace(/\s/g,'').length >= 10 && addClientForm.prenom?.trim() && addClientForm.nom?.trim() && addClientForm.genre && (addClientForm.mail||'').includes('@');
        const sauvegarderNouveauClient = () => { if (!valide) return; addClient(addClientForm); setAddClientForm({}); };
        return (
          <>
            <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();fermerAdd();}} onClick={fermerAdd} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:2999,pointerEvents:'all',cursor:'default',touchAction:'none'}}/>
            <div style={{position:'fixed', ...(isMobile?{inset:0,transform:'none',width:'100%',height:'100%',borderRadius:0}:{top:'50%',left:'50%',transform:'translate(-50%,-50%)',width:'min(520px,calc(100vw - 48px))',maxHeight:'90vh',borderRadius:20}), background:'#fff', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.25)', zIndex:3000, overflow:'hidden'}} onClick={e=>e.stopPropagation()}>
              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'24px 28px 20px',flexShrink:0,borderBottom:'1px solid #f0f0f0'}}>
                <h2 style={{margin:0,fontSize:22,fontWeight:800,color:'#111'}}>Nouveau client</h2>
                <button onClick={fermerAdd} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666',display:'flex',alignItems:'center',justifyContent:'center'}}>✕</button>
              </div>
              <div style={{flex:1,overflowY:'auto',padding:'20px 28px',display:'flex',flexDirection:'column',gap:20}}>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>1. Téléphone <span style={{color:'#dc2626'}}>*</span></p>
                  <div style={{position:'relative'}}>
                    <Phone size={18} strokeWidth={2} color="#999" style={{position:'absolute',left:16,top:'50%',transform:'translateY(-50%)',pointerEvents:'none'}}/>
                    <input type="tel" inputMode="numeric" value={addClientForm.tel||''} onChange={e=>setAddClientForm({...addClientForm,tel:e.target.value})} placeholder="06 43 00 49 87" style={{width:'100%',height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px 0 48px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                  </div>
                </div>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>2. Genre <span style={{color:'#dc2626'}}>*</span></p>
                  <div style={{display:'flex',gap:8}}>
                    {['Homme','Femme','Entreprise'].map(g=>(
                      <button key={g} onClick={()=>setAddClientForm({...addClientForm,genre:g})} style={{flex:1,height:46,borderRadius:10,cursor:'pointer',fontSize:13,fontWeight:700,border:'1.5px solid',borderColor:addClientForm.genre===g?(g==='Homme'?'#3b82f6':g==='Femme'?'#ec4899':'#22c55e'):'#eee',background:addClientForm.genre===g?(g==='Homme'?'#dbeafe':g==='Femme'?'#fce7f3':'#dcfce7'):'#fff',color:addClientForm.genre===g?(g==='Homme'?'#1d4ed8':g==='Femme'?'#be185d':'#15803d'):'#666'}}>{g}</button>
                    ))}
                  </div>
                </div>
                {addClientForm.genre === 'Entreprise' && (
                  <div>
                    <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>3. Nom de l'entreprise <span style={{color:'#dc2626'}}>*</span></p>
                    <input value={addClientForm.entreprise||''} onChange={e=>setAddClientForm({...addClientForm,entreprise:e.target.value})} placeholder="Nom de l'entreprise" style={{width:'100%',height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                  </div>
                )}
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>
                    {addClientForm.genre === 'Entreprise'
                      ? <>{addClientForm.genre==='Entreprise'?'4.':''} Nom du contact <span style={{fontSize:12,fontWeight:400,color:'#999'}}>(optionnel)</span></>
                      : <>3. Prénom et Nom <span style={{color:'#dc2626'}}>*</span></>}
                  </p>
                  <div style={{display:'flex',gap:10}}>
                    <input value={addClientForm.prenom||''} onChange={e=>setAddClientForm({...addClientForm,prenom:e.target.value})} placeholder="Prénom" style={{flex:1,height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                    <input value={addClientForm.nom||''} onChange={e=>setAddClientForm({...addClientForm,nom:e.target.value})} placeholder="Nom" style={{flex:1,height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                  </div>
                </div>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>4. Email <span style={{color:'#dc2626'}}>*</span></p>
                  <div style={{position:'relative'}}>
                    <Mail size={18} strokeWidth={2} color="#999" style={{position:'absolute',left:16,top:'50%',transform:'translateY(-50%)',pointerEvents:'none'}}/>
                    <input type="email" value={addClientForm.mail||''} onChange={e=>setAddClientForm({...addClientForm,mail:e.target.value})} placeholder="email@exemple.com" style={{width:'100%',height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px 0 48px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                  </div>
                </div>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>5. Commentaire <span style={{fontSize:12,fontWeight:400,color:'#999'}}>(optionnel)</span></p>
                  <textarea value={addClientForm.commentaire||''} onChange={e=>setAddClientForm({...addClientForm,commentaire:e.target.value})} placeholder="Notes internes (allergies, préférences...)" style={{width:'100%',height:80,border:'1.5px solid #eee',borderRadius:12,padding:'12px 16px',fontSize:14,outline:'none',resize:'none',fontFamily:'inherit',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                </div>
              </div>
              <div style={{flexShrink:0,padding:'16px 28px',paddingBottom:'calc(16px + env(safe-area-inset-bottom))',borderTop:'1px solid #eee',background:'#fff'}}>
                <button disabled={!valide} onClick={sauvegarderNouveauClient} style={{width:'100%',height:54,background:valide?'#E8C547':'#f0f0f0',color:valide?'#111':'#bbb',border:'none',borderRadius:14,fontSize:16,fontWeight:800,cursor:valide?'pointer':'not-allowed',display:'flex',alignItems:'center',justifyContent:'center',gap:8}}>
                  <UserPlus size={18} strokeWidth={2}/> Créer le client
                </button>
                {!valide && addClientForm.genre && (
                  <p style={{textAlign:'center',fontSize:12,color:'#999',margin:'6px 0 0'}}>
                    {!(addClientForm.tel||'').replace(/\s/g,'').length>=10
                      ? 'Renseignez un numéro de téléphone'
                      : addClientForm.genre==='Entreprise' && !addClientForm.entreprise?.trim()
                      ? "Renseignez le nom de l'entreprise"
                      : !(addClientForm.mail||'').includes('@')
                      ? 'Renseignez un email valide'
                      : addClientForm.genre!=='Entreprise' && (!addClientForm.prenom?.trim()||!addClientForm.nom?.trim())
                      ? 'Renseignez le prénom et le nom'
                      : ''}
                  </p>
                )}
                <p style={{fontSize:11,color:'#bbb',textAlign:'center',margin:'6px 0 0'}}><span style={{color:'#dc2626'}}>*</span> Champs obligatoires</p>
                <button onClick={fermerAdd} style={{width:'100%',background:'none',border:'none',color:'#999',fontSize:14,cursor:'pointer',padding:'8px',marginTop:4}}>Annuler</button>
              </div>
            </div>
            {showConfirmQuitterClient && (
              <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.7)',zIndex:6000,display:'flex',alignItems:'center',justifyContent:'center',pointerEvents:'all',touchAction:'none'}} onMouseDown={e=>{e.preventDefault();e.stopPropagation();setShowConfirmQuitterClient(false);}} onClick={()=>setShowConfirmQuitterClient(false)}>
                <div style={{background:'#fff',borderRadius:16,padding:'28px 24px',maxWidth:320,width:'90%',textAlign:'center'}} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()}>
                  <h3 style={{margin:'0 0 8px',fontSize:17,fontWeight:800,color:'#111'}}>Quitter sans enregistrer ?</h3>
                  <p style={{margin:'0 0 20px',fontSize:14,color:'#666'}}>Les informations saisies seront perdues.</p>
                  <div style={{display:'flex',gap:10}}>
                    <button onClick={()=>setShowConfirmQuitterClient(false)} style={{flex:1,height:44,border:'1.5px solid #ddd',borderRadius:10,background:'#fff',fontSize:14,fontWeight:600,cursor:'pointer',color:'#666'}}>Continuer</button>
                    <button onClick={()=>{setShowConfirmQuitterClient(false);setModalAdd(false);setAddClientForm({});}} style={{flex:1,height:44,border:'none',borderRadius:10,background:'#dc2626',fontSize:14,fontWeight:800,cursor:'pointer',color:'#fff'}}>Quitter</button>
                  </div>
                </div>
              </div>
            )}
          </>
        );
      })()}
      {modalEdit && (() => {
        const aDesDonnees = editForm.prenom !== (modalEdit.prenom||'') || editForm.nom !== (modalEdit.nom||'') || editForm.tel !== (modalEdit.tel||'') || editForm.mail !== (modalEdit.mail||'') || editForm.genre !== (modalEdit.genre||'') || editForm.entreprise !== (modalEdit.entreprise||'') || editForm.commentaire !== (modalEdit.commentaire||'');
        const fermerEdit = () => {
          if (aDesDonnees) { setPendingFermer(()=>()=>setModalEdit(null)); setShowConfirmQuitter(true); }
          else { setModalEdit(null); }
        };
        return (
          <>
            <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();fermerEdit();}} onClick={fermerEdit} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:2999,pointerEvents:'all',cursor:'default',touchAction:'none'}} />
            <div style={{position:'fixed',top:'50%',left:'50%',transform:'translate(-50%,-50%)',background:'#fff',borderRadius:20,width:'min(520px,calc(100vw - 48px))',maxHeight:'90vh',display:'flex',flexDirection:'column',boxShadow:'0 32px 80px rgba(0,0,0,0.25)',zIndex:3000,overflow:'hidden'}} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()}>
              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'24px 28px 20px',flexShrink:0,borderBottom:'1px solid #f0f0f0'}}>
                <h2 style={{margin:0,fontSize:22,fontWeight:800,color:'#111'}}>Modifier le client</h2>
                <button onClick={fermerEdit} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666',display:'flex',alignItems:'center',justifyContent:'center'}}>✕</button>
              </div>
              <div style={{flex:1,overflowY:'auto',padding:'20px 28px',display:'flex',flexDirection:'column',gap:18}}>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>Genre</p>
                  <div style={{display:'flex',gap:8}}>
                    {['Homme','Femme','Entreprise'].map(g=>(
                      <button key={g} onClick={()=>setEditForm({...editForm,genre:g})} style={{flex:1,height:44,borderRadius:10,cursor:'pointer',fontSize:13,fontWeight:700,border:'1.5px solid',borderColor:editForm.genre===g?(g==='Homme'?'#3b82f6':g==='Femme'?'#ec4899':'#22c55e'):'#eee',background:editForm.genre===g?(g==='Homme'?'#dbeafe':g==='Femme'?'#fce7f3':'#dcfce7'):'#fff',color:editForm.genre===g?(g==='Homme'?'#1d4ed8':g==='Femme'?'#be185d':'#15803d'):'#666'}}>{g}</button>
                    ))}
                  </div>
                </div>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>Prénom</p>
                  <input value={editForm.prenom||''} onChange={e=>setEditForm({...editForm,prenom:e.target.value})} style={{width:'100%',height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                </div>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>Nom</p>
                  <input value={editForm.nom||''} onChange={e=>setEditForm({...editForm,nom:e.target.value})} style={{width:'100%',height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                </div>
                {editForm.genre==='Entreprise' && (
                  <div>
                    <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>Nom de l'entreprise</p>
                    <input value={editForm.entreprise||''} onChange={e=>setEditForm({...editForm,entreprise:e.target.value})} style={{width:'100%',height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                  </div>
                )}
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>Téléphone</p>
                  <div style={{position:'relative'}}>
                    <Phone size={18} strokeWidth={2} color="#999" style={{position:'absolute',left:16,top:'50%',transform:'translateY(-50%)',pointerEvents:'none'}}/>
                    <input value={editForm.tel||''} onChange={e=>setEditForm({...editForm,tel:e.target.value})} type="tel" style={{width:'100%',height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px 0 48px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                  </div>
                </div>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>Email</p>
                  <div style={{position:'relative'}}>
                    <Mail size={18} strokeWidth={2} color="#999" style={{position:'absolute',left:16,top:'50%',transform:'translateY(-50%)',pointerEvents:'none'}}/>
                    <input value={editForm.mail||''} onChange={e=>setEditForm({...editForm,mail:e.target.value})} type="email" style={{width:'100%',height:52,border:'1.5px solid #eee',borderRadius:12,padding:'0 16px 0 48px',fontSize:15,outline:'none',boxSizing:'border-box'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                  </div>
                </div>
                <div>
                  <p style={{fontSize:14,fontWeight:800,color:'#111',margin:'0 0 10px'}}>Commentaire</p>
                  <textarea value={editForm.commentaire||''} onChange={e=>setEditForm({...editForm,commentaire:e.target.value})} rows={3} style={{width:'100%',border:'1.5px solid #eee',borderRadius:12,padding:'12px 16px',fontSize:14,outline:'none',resize:'vertical',boxSizing:'border-box',fontFamily:'inherit'}} onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                </div>
              </div>
              <div style={{flexShrink:0,padding:'16px 28px',paddingBottom:'calc(16px + env(safe-area-inset-bottom))',borderTop:'1px solid #eee',background:'#fff',display:'flex',gap:10}}>
                <button onClick={fermerEdit} style={{flex:1,height:52,border:'1.5px solid #eee',borderRadius:12,background:'#fff',fontSize:15,fontWeight:600,cursor:'pointer',color:'#666'}}>Annuler</button>
                <button onClick={sauvegarderEditClient} style={{flex:2,height:52,border:'none',borderRadius:12,background:'#E8C547',fontSize:15,fontWeight:800,cursor:'pointer',color:'#111',display:'flex',alignItems:'center',justifyContent:'center',gap:8}}>
                  <Save size={18} strokeWidth={2}/> Enregistrer
                </button>
              </div>
            </div>
          </>
        );
      })()}
      {modalDelete && (
        <>
          <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();}} onClick={()=>setModalDelete(null)} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.6)',zIndex:4000,pointerEvents:'all'}}/>
          <div onClick={e=>e.stopPropagation()} onMouseDown={e=>e.stopPropagation()} style={{position:'fixed',top:'50%',left:'50%',transform:'translate(-50%,-50%)',background:'#fff',borderRadius:20,width:'min(440px,calc(100vw - 48px))',display:'flex',flexDirection:'column',boxShadow:'0 32px 80px rgba(0,0,0,0.25)',zIndex:4001,overflow:'hidden'}}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'24px 28px 20px',flexShrink:0}}>
              <h2 style={{margin:0,fontSize:20,fontWeight:800,color:'#111'}}>
                Supprimer {modalDelete.genre==='Entreprise'?(modalDelete.entreprise||modalDelete.nom):`${modalDelete.prenom} ${modalDelete.nom}`} ?
              </h2>
              <button onClick={()=>setModalDelete(null)} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666',display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}>✕</button>
            </div>
            <div style={{padding:'0 28px 24px'}}>
              <div style={{background:'#fff5f5',border:'1.5px solid #fca5a5',borderRadius:12,padding:'14px 16px',marginBottom:20,display:'flex',alignItems:'center',gap:12}}>
                <Trash2 size={20} strokeWidth={2} color="#dc2626" style={{flexShrink:0}}/>
                <p style={{margin:0,fontSize:14,color:'#dc2626',lineHeight:1.5}}>Cette action est définitive. Le client sera déplacé dans la corbeille et pourra être restauré.</p>
              </div>
              <div style={{display:'flex',gap:10}}>
                <button onClick={()=>setModalDelete(null)} style={{flex:1,height:52,border:'1.5px solid #eee',borderRadius:12,background:'#fff',fontSize:15,fontWeight:600,cursor:'pointer',color:'#666'}}>Annuler</button>
                <button onClick={()=>deleteClient(modalDelete.id)} style={{flex:1,height:52,border:'none',borderRadius:12,background:'#dc2626',fontSize:15,fontWeight:800,cursor:'pointer',color:'#fff',display:'flex',alignItems:'center',justifyContent:'center',gap:8}}>
                  <Trash2 size={16} strokeWidth={2} color="#fff"/> Supprimer
                </button>
              </div>
            </div>
          </div>
        </>
      )}
      {modalImport && <ImportModal existingClients={clients} onImport={importClients} onCancel={()=>setModalImport(false)} />}
      {modalComment && <Modal title={`Commentaire — ${modalComment.prenom} ${modalComment.nom}`} onClose={()=>setModalComment(null)}><p style={{fontSize:14,lineHeight:1.7,margin:0}}>{modalComment.commentaire}</p></Modal>}
      {modalCorbeille && !isMobile && <CorbeilleModal onClose={()=>{ setModalCorbeille(false); loadClients(true); }} showToast={showToast} />}

      {/* Modal Top 50 clients */}
      {showTopClients && (
        <>
          <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();}} onClick={()=>setShowTopClients(false)} style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.5)',zIndex:2999,pointerEvents:'all'}}/>
          <div onClick={e=>e.stopPropagation()} style={{position:'fixed',top:'50%',left:'50%',transform:'translate(-50%,-50%)',background:'#fff',borderRadius:20,width:'min(520px, calc(100vw - 48px))',maxHeight:'80vh',display:'flex',flexDirection:'column',boxShadow:'0 32px 80px rgba(0,0,0,0.25)',zIndex:3000,overflow:'hidden'}}>
            <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'24px 28px 20px',flexShrink:0,borderBottom:'1px solid #f0f0f0'}}>
              <h2 style={{margin:0,fontSize:20,fontWeight:800,color:'#111'}}>🏆 Classement clients</h2>
              <button onClick={()=>setShowTopClients(false)} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666',display:'flex',alignItems:'center',justifyContent:'center'}}>✕</button>
            </div>
            <div style={{flex:1,overflowY:'auto',padding:'16px 28px'}}>
              {clients
                .map(c=>({...c, nb:resasData.filter(r=>r.client_id===c.id&&r.statut!=='absente'&&r.statut!=='annulee'&&r.statut!=='refusee').length}))
                .filter(c=>c.nb>0)
                .sort((a,b)=>b.nb-a.nb)
                .slice(0,50)
                .map((c,i)=>{
                  const medals=['🥇','🥈','🥉'];
                  const avatarBg=c.genre==='Homme'?'#dbeafe':c.genre==='Femme'?'#fce7f3':'#dcfce7';
                  const avatarColor=c.genre==='Homme'?'#1d4ed8':c.genre==='Femme'?'#be185d':'#15803d';
                  const initiales=c.genre==='Entreprise'?(c.entreprise||'?').slice(0,2).toUpperCase():`${(c.prenom||'?')[0]}${(c.nom||'')[0]||''}`.toUpperCase();
                  return (
                    <div key={c.id} onClick={()=>{setModalDetailClient(c);setShowTopClients(false);}} style={{display:'flex',alignItems:'center',gap:12,padding:'10px 0',borderBottom:'1px solid #f5f5f5',cursor:'pointer'}}
                      onMouseEnter={e=>e.currentTarget.style.background='#fafafa'}
                      onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                      <span style={{fontSize:i<3?18:13,minWidth:28,textAlign:'center'}}>{i<3?medals[i]:`#${i+1}`}</span>
                      <div style={{width:34,height:34,borderRadius:'50%',background:avatarBg,display:'flex',alignItems:'center',justifyContent:'center',fontSize:12,fontWeight:800,color:avatarColor,flexShrink:0}}>{initiales}</div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontWeight:700,fontSize:14,color:'#111',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{c.genre==='Entreprise'?c.entreprise:`${c.prenom||''} ${c.nom||''}`}</div>
                        <div style={{fontSize:12,color:'#999'}}>{c.tel}</div>
                      </div>
                      <span style={{background:'#fffbea',color:'#111',borderRadius:20,padding:'3px 12px',fontSize:13,fontWeight:800,flexShrink:0}}>{c.nb} résa</span>
                    </div>
                  );
                })}
            </div>
            <div style={{flexShrink:0,padding:'16px 28px',borderTop:'1px solid #eee'}}>
              <button onClick={()=>setShowTopClients(false)} style={{width:'100%',height:48,border:'1.5px solid #eee',borderRadius:12,background:'#fff',fontSize:14,fontWeight:600,cursor:'pointer',color:'#666'}}>Fermer</button>
            </div>
          </div>
        </>
      )}

      {toast && <Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)} />}
      {notifPrePromptModal}
      {sendingModal && <SendingProgressModal type={sendingModal.type} total={sendingModal.total} done={sendingModal.done} successCount={sendingModal.successCount} onClose={()=>setSendingModal(null)} />}
      {showConfirmQuitter && (
        <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.7)',zIndex:9000,display:'flex',alignItems:'center',justifyContent:'center',pointerEvents:'all',touchAction:'none'}} onMouseDown={e=>{e.preventDefault();e.stopPropagation();setShowConfirmQuitter(false);}} onClick={()=>setShowConfirmQuitter(false)}>
          <div style={{background:'#fff',borderRadius:16,padding:'28px 24px',maxWidth:320,width:'90%',textAlign:'center'}} onMouseDown={e=>e.stopPropagation()} onClick={e=>e.stopPropagation()}>
            <h3 style={{margin:'0 0 8px',fontSize:17,fontWeight:800,color:'#111'}}>Quitter sans enregistrer ?</h3>
            <p style={{margin:'0 0 20px',fontSize:14,color:'#666'}}>Les informations saisies seront perdues.</p>
            <div style={{display:'flex',gap:10}}>
              <button onClick={()=>setShowConfirmQuitter(false)} style={{flex:1,height:44,border:'1.5px solid #ddd',borderRadius:10,background:'#fff',fontSize:14,fontWeight:600,cursor:'pointer',color:'#666'}}>Continuer la saisie</button>
              <button onClick={()=>{setShowConfirmQuitter(false); if(pendingFermer) { pendingFermer(); setPendingFermer(null); }}} style={{flex:1,height:44,border:'none',borderRadius:10,background:'#dc2626',fontSize:14,fontWeight:800,cursor:'pointer',color:'#fff'}}>Quitter</button>
            </div>
          </div>
        </div>
      )}
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
