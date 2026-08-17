import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import { Mail, LockKeyhole, Eye, EyeOff, RefreshCw, ShieldCheck, MonitorSmartphone, Headphones, ArrowRight, AlertCircle, Users, UtensilsCrossed, Phone, Download, CalendarDays, Megaphone, Link, LogOut, Copy, ExternalLink, Share2, ClipboardList, CircleCheck, User, ChevronRight, ChevronDown, Pencil, Sun, Moon, ArrowLeft, MessageSquare, UserX, Clock, Star, Trash2, Send, History, Building2, CheckCircle, Check, Search, RotateCcw, Save, Plus, UserPlus, Trophy, ArrowUpDown, LayoutGrid, Settings, MapPin, Dices, Bell, X, Award, Gift, Image as ImageIcon, BadgeCheck, ShoppingBag, BarChart3, Info } from 'lucide-react';
import { supabase } from "./supabase";
import { safeQuery, resilientChannel, logError } from "./lib/db";

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

// ─── Auth CRM pour les endpoints protégés ────────────────────────────────────
// Les Cloudflare Functions sensibles exigent le JWT Supabase de l'utilisateur.
async function crmAuthHeaders() {
  const { data: { session } } = await supabase.auth.getSession();
  return { 'Content-Type': 'application/json', Authorization: `Bearer ${session?.access_token || ''}` };
}

// ─── Brevo Email ─────────────────────────────────────────────────────────────
async function sendBrevoEmail(toEmail, toName, subject, htmlContent) {
  if (!toEmail) return { success: false };
  const body = { to: toEmail, toName, subject, html: htmlContent };
  try {
    const res = await fetch('/send-email', {
      method: 'POST',
      headers: await crmAuthHeaders(),
      body: JSON.stringify(body)
    });
    const data = await res.json();
    return { success: data.success };
  } catch(e) {
    logError(e.message, 'sendBrevoEmail');
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
                {dupClient && <div style={{ background:'#fef2f2', border:'2px solid #dc2626', borderRadius:10, padding:'10px 14px', marginTop:8, fontSize:13, color:'#dc2626' }}><AlertCircle size={14} color="currentColor" style={{display:'inline',verticalAlign:'middle'}} /> Ce numéro est déjà utilisé par <strong>{dupClient.prenom} {dupClient.nom}</strong></div>}
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
            {success ? (<><span style={{ display:"inline-block", animation:"scaleIn 0.3s cubic-bezier(0.34, 1.56, 0.64, 1)" }}>✓</span>Enregistré !</>) : dupClient ? <><AlertCircle size={14} color="currentColor" style={{display:'inline',verticalAlign:'middle'}} /> Client existant</> : isEdit ? "Modifier" : "Enregistrer"}
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
              <AlertCircle size={24} color="currentColor" style={{display:'inline',verticalAlign:'middle'}} />
              <div>
                <p style={{fontWeight:700, color:"#dc2626", fontSize:14, margin:0}}>Client déjà existant !</p>
                <p style={{fontSize:13, color:"#333", margin:"4px 0 0"}}><strong>{dupClient.prenom} {dupClient.nom}{dupClient.entreprise ? ` — ${dupClient.entreprise}` : ""}</strong></p>
                <p style={{fontSize:12, color:"#666", margin:"2px 0 0"}}><Phone size={14} style={{display:'inline',verticalAlign:'middle'}} /> {dupClient.tel}{dupClient.mail ? ` · ${dupClient.mail}` : ""}</p>
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
          {dups.length > 0 && <div style={{ background:"#fffbeb", border:"1.5px solid #fbbf24", borderRadius:8, padding:"10px 14px", fontSize:13, color:"#92400e", marginBottom:12 }}><AlertCircle size={14} color="currentColor" style={{display:'inline',verticalAlign:'middle'}} /> {dups.length} doublon(s) potentiel(s) détecté(s).</div>}
          <p style={{ fontWeight:600, marginBottom:8 }}>{parsed.length} client(s) détecté(s)</p>
          <div style={{ maxHeight:180, overflowY:"auto", fontSize:12, border:"1px solid #eee", borderRadius:6, padding:"8px" }}>
            {parsed.map((c,i) => <div key={i} style={{ padding:"3px 0", borderBottom:"1px solid #f0f0f0" }}><span style={{fontWeight:600}}>{c.prenom} {c.nom}</span><span style={{color:"#999",marginLeft:8}}>{c.tel} {c.mail}</span></div>)}
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
    return c.genre === "Entreprise" ? (c.entreprise || c.nom) : `${c.prenom || ""} ${c.nom || ""}`.trim();
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
              <Trash2 size={20} color="#dc2626" />
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
              <div style={{ marginBottom:12 }}><Trash2 size={52} color="#dc2626" /></div>
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
                {c.tel && <div style={{ fontSize:13, color:"#666", marginBottom:2 }}><Phone size={14} style={{display:'inline',verticalAlign:'middle'}} /> {c.tel}</div>}
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
// 900 px : au-dessous, la mise en page « mobile » (une colonne, navigation
// basse) sert aussi les tablettes en portrait, où le gabarit bureau déborde.
function useIsMobile() {
  const [isMobile, setIsMobile] = useState(window.innerWidth < 900);
  useEffect(() => {
    const handler = () => setIsMobile(window.innerWidth < 900);
    window.addEventListener('resize', handler);
    window.addEventListener('orientationchange', handler);
    return () => { window.removeEventListener('resize', handler); window.removeEventListener('orientationchange', handler); };
  }, []);
  return isMobile;
}

// Tablettes comprises : en dessous de 1180 px, les mises en page à deux colonnes
// deviennent illisibles (calendrier écrasé, barres de filtres qui débordent).
function useEcranEtroit(seuil = 1180) {
  const [etroit, setEtroit] = useState(window.innerWidth < seuil);
  useEffect(() => {
    const handler = () => setEtroit(window.innerWidth < seuil);
    window.addEventListener('resize', handler);
    window.addEventListener('orientationchange', handler);
    return () => { window.removeEventListener('resize', handler); window.removeEventListener('orientationchange', handler); };
  }, [seuil]);
  return etroit;
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
    // Même validation que les formulaires publics : formats tel FR / email
    if (!/^(\+33|0)[1-9]\d{8}$/.test(tel.replace(/[\s.\-()]/g, ''))) { showToast('Numéro de téléphone invalide', 'error'); return; }
    if (email.trim() && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email.trim())) { showToast('Email invalide', 'error'); return; }
    if (prenom.trim().length > 50 || nom.trim().length > 50) { showToast('Nom/prénom trop long (max 50 caractères)', 'error'); return; }
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
        {saving ? 'Enregistrement…' : (isEdit ? <><Pencil size={15} style={{display:'inline',verticalAlign:'middle'}} /> Modifier la réservation</> : (formValide ? '✓ Créer la réservation' : 'Créer la réservation'))}
      </button>
      <div style={{ textAlign:'center', fontSize:12, marginTop:8, minHeight:20, transition:'opacity 0.2s' }}>
        {consigne && (
          <span style={{ color: consigne.invalide ? '#dc2626' : '#999' }}>
            {consigne.invalide ? <AlertCircle size={13} color="currentColor" style={{display:'inline',verticalAlign:'middle'}} /> : '→ '}{consigne.msg}
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
              <span style={{ fontWeight:700, fontSize:14 }}>{resaCree.service==='midi'?<><Sun size={14} style={{display:'inline',verticalAlign:'middle'}} /> Midi</>:<><Moon size={14} style={{display:'inline',verticalAlign:'middle'}} /> Soir</>}</span>
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
        <Clock size={20} color="#92400e" />
        <p style={{ margin:0, fontSize:13, color:'#92400e' }}>Cette réservation sera créée comme <strong>demande en attente</strong>.</p>
      </div>}

      <div style={{ display:'flex', flexDirection:'column', gap:20, paddingBottom:8 }}>

        {/* ── Section 1 : Téléphone ── */}
        <div ref={refTel}>
          <div style={{ fontSize:13, fontWeight:800, color:'#555', marginBottom:8, textTransform:'uppercase', letterSpacing:0.5 }}>1. Téléphone du client</div>
          <div style={{ position:'relative' }}>
            <span style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }}><Phone size={16} style={{display:'block'}} /></span>
            <input value={tel} onChange={e=>handleTelChange(e.target.value)} placeholder="06 12 34 56 78" type="tel" style={{ ...inp(false), paddingLeft:40 }} />
            {lookingUp && <span style={{ position:'absolute', right:12, top:'50%', transform:'translateY(-50%)', fontSize:12, color:'#888' }}>Recherche…</span>}
          </div>
          {clientFound && (
            <div style={{ marginTop:8 }}>
              <div style={{ background:'#f0fdf4', border:'1.5px solid #22c55e', borderRadius:10, padding:'10px 14px', display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
                <CheckCircle size={14} color="#16a34a" style={{display:'inline',verticalAlign:'middle'}} />
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
            <span><CalendarDays size={14} style={{display:'inline',verticalAlign:'middle'}} /> {dateIso ? new Date(dateIso+'T12:00:00').toLocaleDateString('fr-FR', {weekday:'long', day:'numeric', month:'long', year:'numeric'}) : 'Choisir une date'}</span>
            <span style={{ color:'#ccc', fontSize:20 }}>›</span>
          </button>
          {calendarJSX}
          <div ref={refService} style={{ display:'flex', gap:8, marginTop:10 }}>
            <button style={btnSvc('midi')} onClick={()=>{ setService('midi'); setHeure(''); setHeureError(false); }}><Sun size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Midi</button>
            <button style={btnSvc('soir')} onClick={()=>{ setService('soir'); setHeure(''); setHeureError(false); }}><Moon size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Soir</button>
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
                  {[['Date', new Date(resaCree.date+'T12:00:00').toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long'})],['Service',resaCree.service==='midi'?<><Sun size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Midi</>:<><Moon size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Soir</>],['Heure',resaCree.heure],['Personnes',`${resaCree.nb_personnes} pers.`]].map(([k,v],i,arr)=>(
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
                {saving ? 'Enregistrement...' : (isEdit ? <><Pencil size={15} style={{display:'inline',verticalAlign:'middle'}} /> Modifier la réservation</> : (formValide ? '✓ Créer la réservation' : 'Créer la réservation'))}
              </button>
              <div style={{ textAlign:'center', fontSize:12, marginTop:8, minHeight:20, transition:'opacity 0.2s' }}>
                {consigne && (
                  <span style={{ color: consigne.invalide ? '#dc2626' : '#999' }}>
                    {consigne.invalide ? <AlertCircle size={13} color="currentColor" style={{display:'inline',verticalAlign:'middle'}} /> : '→ '}{consigne.msg}
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
          <div style={{ marginBottom:12 }}><AlertCircle size={40} color="#dc2626" style={{display:'block',margin:'0 auto'}} /></div>
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
              {label:'Service', value: resa.service==='midi'?<><Sun size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/>{` Midi · ${resa.heure}`}</>:<><Moon size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/>{` Soir · ${resa.heure}`}</>},
              {label:'Personnes', value: <><Users size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/>{`${resa.nb_personnes} pers.`}</>},
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
              {label:'Service', value: resa.service==='midi'?<><Sun size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Midi</>:<><Moon size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Soir</>},
              {label:'Heure', value: resa.heure||'—'},
              {label:'Personnes', value: `${resa.nb_personnes} pers.`},
              {label:'Occasion', value: resa.occasion||'—'},
            ].map((item,i)=>(
              <div key={i} style={{background:'#f9f9f9',borderRadius:10,padding:'12px 14px'}}>
                <div style={{fontSize:10,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:0.5,marginBottom:4}}>{item.label}</div>
                <div style={{fontSize:14,fontWeight:600,color:'#111'}}>{item.value}</div>
              </div>
            ))}
            {resa.source === 'Grand Jeux du TED' && (
              <div style={{gridColumn:'1/-1',background:'#fffbea',borderRadius:10,padding:'10px 14px',display:'flex',alignItems:'center',gap:8}}>
                <Dices size={18} style={{display:'inline',verticalAlign:'middle'}} />
                <span style={{fontSize:13,fontWeight:700,color:'#92400e'}}>Grand Jeu du TED</span>
              </div>
            )}
            {resa.note_interne && (
              <div style={{gridColumn:'1/-1',background:'#f9f9f9',borderRadius:10,padding:'12px 14px'}}>
                <div style={{fontSize:10,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:0.5,marginBottom:4}}>Note</div>
                <div style={{fontSize:14,color:'#555'}}>{resa.note_interne}</div>
              </div>
            )}
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
                {[{label:'{prénom}',val:c.prenom||'{prénom}'},{label:'{nom}',val:c.nom||'{nom}'},{label:'Lien',val:'https://ted-crm.pages.dev/reserver.html'}].map(v=>(
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
            <button onClick={()=>{onClose();onEdit(resa);}} style={{flex:1,height:52,border:'none',borderRadius:12,background:'#f0f0f0',fontSize:15,fontWeight:700,cursor:'pointer',color:'#111',display:'flex',alignItems:'center',justifyContent:'center',gap:6}}><Pencil size={15} style={{display:'inline',verticalAlign:'middle'}} /> Modifier</button>
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
  const etroit = useEcranEtroit();
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
    return resilientChannel(supabase, 'resa-page-realtime', (chan) => chan
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
    );
  }, []);

  async function loadResa() {
    setLoading(true);
    const { data, error } = await safeQuery(() => supabase
      .from('reservations')
      .select('*, clients(id, nom, prenom, tel, mail, genre, entreprise)')
      .order('created_at', { ascending: false }), { fallback: [], context: 'loadResa' });
    if (error) showToast('Erreur chargement réservations', 'error');
    else {
      setResaList(data || []);
      onResaCountChange?.((data||[]).filter(r=>r.statut==='attente').length);
    }
    setLoading(false);
  }

  const accepterLockRef = useRef(false);
  async function accepter(r) {
    // Verrou anti double-clic : évite double update + double email de confirmation
    if (accepterLockRef.current) return;
    accepterLockRef.current = true;
    setTimeout(() => { accepterLockRef.current = false; }, 3000);
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
    <div style={{ fontFamily:"'Inter','Segoe UI',Arial,sans-serif", background:'#f8f8f8', minHeight: inline ? undefined : '100vh', overflow: (!isMobile && !etroit) ? 'hidden' : undefined, height: (!isMobile && !etroit && inline) ? '100vh' : undefined }}>
      {/* Header — desktop full-page mode only */}
      {!inline && (
        <header style={{ background:'#111', color:'#fff', padding:'0 20px', height:56, display:'flex', alignItems:'center', justifyContent:'space-between', borderBottom:`3px solid ${G}`, flexShrink:0 }}>
          <div style={{ display:'flex', alignItems:'center', gap:10 }}>
            <span style={{ fontWeight:700, fontSize:15, color:'#fff' }}><CalendarDays size={14} style={{display:'inline',verticalAlign:'middle',marginRight:4}} /><span style={{ color:G }}>TED</span> — Réservations</span>
            <div id="formulaire-dropdown" style={{ position:'relative' }}>
              <button onClick={() => setShowFormDropdown(v => !v)} style={{ background:'rgba(255,255,255,0.08)', border:'1px solid #444', borderRadius:8, height:34, padding:'0 14px', color:'#ccc', fontWeight:600, fontSize:13, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}><Link size={13} /> Formulaire</button>
              {showFormDropdown && (
                <div style={{ position:'absolute', top:40, left:0, background:'#fff', borderRadius:10, border:'1.5px solid #eee', boxShadow:'0 8px 24px rgba(0,0,0,0.12)', padding:8, zIndex:200, minWidth:180 }}>
                  <button type="button" onMouseDown={async(e)=>{ e.preventDefault(); e.stopPropagation(); try{ await navigator.clipboard.writeText('https://ted-crm.pages.dev/reserver.html'); }catch{ const t=document.createElement('textarea'); t.value='https://ted-crm.pages.dev/reserver.html'; document.body.appendChild(t); t.select(); document.execCommand('copy'); document.body.removeChild(t); } showToast('✅ Lien copié !'); setShowFormDropdown(false); }} style={{ display:'block', width:'100%', textAlign:'left', padding:'9px 14px', border:'none', background:'none', cursor:'pointer', fontSize:13, fontWeight:600, borderRadius:7 }}><ClipboardList size={14} style={{display:'inline',verticalAlign:'middle',marginRight:6}} />Copier</button>
                  <button type="button" onMouseDown={(e)=>{ e.preventDefault(); e.stopPropagation(); window.open('https://ted-crm.pages.dev/reserver.html','_blank'); setShowFormDropdown(false); }} style={{ display:'block', width:'100%', textAlign:'left', padding:'9px 14px', border:'none', background:'none', cursor:'pointer', fontSize:13, fontWeight:600, borderRadius:7 }}><Link size={14} style={{display:'inline',verticalAlign:'middle',marginRight:6}} />Ouvrir</button>
                  <button type="button" onMouseDown={async(e)=>{ e.preventDefault(); e.stopPropagation(); const url='https://ted-crm.pages.dev/reserver.html'; if(navigator.share){ try{ await navigator.share({title:'Réservation Le TED',url}); }catch{} }else{ try{ await navigator.clipboard.writeText(url); }catch{} showToast('✅ Lien copié !'); } setShowFormDropdown(false); }} style={{ display:'block', width:'100%', textAlign:'left', padding:'9px 14px', border:'none', background:'none', cursor:'pointer', fontSize:13, fontWeight:600, borderRadius:7 }}><Share2 size={14} style={{display:'inline',verticalAlign:'middle',marginRight:6}} />Partager</button>
                </div>
              )}
            </div>
          </div>
          <div style={{ display:'flex', alignItems:'center', gap:8 }}>
            <button onClick={onBack} style={{ background:'rgba(255,255,255,0.08)', border:'1px solid #444', borderRadius:8, height:34, padding:'0 14px', color:'#ccc', fontWeight:600, fontSize:13, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}><Users size={14} style={{display:'inline',verticalAlign:'middle'}} /> Mes Clients</button>
            <button onClick={()=>onBack('communications')} style={{ background:'rgba(255,255,255,0.08)', border:'1px solid #444', borderRadius:8, height:34, padding:'0 14px', color:'#ccc', fontWeight:600, fontSize:13, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}><Megaphone size={14} style={{display:'inline',verticalAlign:'middle'}} /> Communications</button>
            <button onClick={()=>setShowConfirmDecoRP(true)} style={{ background:'transparent', color:'#ccc', border:'1px solid #444', borderRadius:7, padding:'0 12px', height:32, fontSize:12, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}><LockKeyhole size={14} style={{display:'inline',verticalAlign:'middle'}} /> Déconnexion</button>
          </div>
        </header>
      )}

      <div style={{ display: !isMobile ? 'grid' : 'block', gridTemplateColumns: !isMobile ? (etroit ? '1fr 268px' : '1fr 380px') : undefined, gap: !isMobile ? (etroit ? 12 : 16) : undefined, padding: !isMobile ? (etroit ? '14px 16px' : '24px 32px') : undefined, maxWidth: !isMobile ? 1440 : undefined, margin: !isMobile ? '0 auto' : undefined, alignItems: !isMobile ? 'stretch' : 'start', height: !isMobile ? (etroit ? 'calc(100vh - 28px)' : 'calc(100vh - 48px)') : undefined, boxSizing: !isMobile ? 'border-box' : undefined, background: !isMobile ? '#f5f5f5' : undefined }}>
      <main style={{ maxWidth: isMobile ? 800 : 'none', margin: isMobile ? '0 auto' : 0, padding: isMobile ? '12px 16px 100px' : '0', display: !isMobile ? 'flex' : 'block', flexDirection: !isMobile ? 'column' : undefined, gap: !isMobile ? (etroit ? 10 : 12) : undefined, height: !isMobile ? '100%' : undefined, overflow: !isMobile ? 'hidden' : undefined }}>

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
            <div onClick={()=>setShowDemandesAttente(true)} className={nbAttente > 0 ? 'alarm-blink' : ''} style={{ background: nbAttente > 0 ? '#dc2626' : '#fff', border: nbAttente > 0 ? 'none' : '1.5px solid #f0f0f0', borderRadius:16, padding: etroit ? '9px 16px' : '14px 20px', display:'flex', alignItems:'center', justifyContent:'space-between', cursor:'pointer', flexShrink:0, transition:'background 0.1s', boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
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
            <div style={{ background:'#fff', borderRadius:16, border:'1.5px solid #f0f0f0', padding: etroit ? 10 : 14, flex:1, display:'flex', flexDirection:'column', overflow:'hidden', boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
              {/* 1. 15 jours scroll natif */}
              <div ref={joursScrollRef} className="jours-strip"
                style={{ display:'flex', gap:8, overflowX:'scroll', marginBottom: etroit ? 8 : 16, flexShrink:0, WebkitOverflowScrolling:'touch', scrollSnapType:'x mandatory', userSelect:'none', WebkitUserSelect:'none' }}>
                {quinzeJours.map(j => {
                  const totalCouverts = resaList.filter(r => r.date===j.date && r.statut==='confirmee').reduce((sum,r)=>sum+(r.nb_personnes||0),0);
                  const isSelected = calJourSelectionne === j.date;
                  return (
                    <div key={j.date} onClick={()=>setCalJourSelectionne(j.date)}
                      style={{ borderRadius:12, padding: etroit ? '6px 6px' : '10px 6px', textAlign:'center', cursor:'pointer', border:'2px solid', borderColor: isSelected?'#E8C547':'#eee', background: isSelected?'#fffbea':'#fff', transition:'border-color 0.15s, background 0.15s', flexShrink:0, width:'calc((100% - 40px) / 6)', scrollSnapAlign:'start' }}>
                      <div style={{ fontSize:10, fontWeight:700, marginBottom: etroit ? 1 : 4, color: isSelected?'#E8C547': j.isAujourd?'#E8C547':'#999' }}>{j.isAujourd?'AUJ.':j.jour}</div>
                      <div style={{ fontSize: etroit ? 17 : 20, fontWeight:900, marginBottom:1, color:'#111' }}>{j.num}</div>
                      {!etroit && <div style={{ fontSize:10, color:'#999', marginBottom:4 }}>{j.mois}</div>}
                      <div style={{ fontSize: etroit ? 10.5 : 11, fontWeight:700, color: totalCouverts>0?'#111':'#ccc' }}>{totalCouverts>0?`${totalCouverts} p.`:'—'}</div>
                      {j.isAujourd && !etroit && <div style={{ width:5, height:5, borderRadius:'50%', background:'#E8C547', margin:'4px auto 0' }}/>}
                    </div>
                  );
                })}
              </div>
              <div style={{ display: !isMobile ? 'grid' : 'block', gridTemplateColumns: !isMobile ? (etroit ? '2fr 1fr' : '3fr 2fr') : undefined, gap: !isMobile ? (etroit ? 12 : 16) : 0, marginTop: !isMobile ? (etroit ? 10 : 16) : 0, flex: !isMobile ? 1 : undefined, minHeight: !isMobile ? 0 : undefined, overflow: !isMobile ? 'hidden' : undefined }}>
                {/* Colonne calendrier */}
                <div style={!isMobile ? { background:'#f8f8f8', borderRadius:12, padding:12, overflow:'auto' } : {}}>
                  {/* 2. Bouton toggle (mobile only) */}
                  {isMobile && (
                    <button onClick={()=>setCalMensuelOuvert(v=>!v)} style={{ width:'100%', padding:'10px 12px', background:'#f8f8f8', border:'1.5px solid #eee', borderRadius:8, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom: calMensuelOuvert ? 12 : 0 }}>
                      <span style={{ fontSize:13, fontWeight:700, color:'#555' }}><CalendarDays size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}} />Calendrier</span>
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
                              style={{ textAlign:'center', height: etroit ? 33 : 48, borderRadius:6, cursor:'pointer', position:'relative',
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
                      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom: etroit ? 6 : 10 }}>
                        <button onClick={()=>changerMois(-1)} style={{ background:'#f0f0f0', border:'none', borderRadius:8, width:34, height:34, fontSize:16, cursor:'pointer', fontWeight:700 }}>‹</button>
                        <span style={{ fontWeight:800, fontSize: etroit ? 16 : 18 }}>{MOIS[mois]} {annee}</span>
                        <button onClick={()=>changerMois(1)} style={{ background:'#f0f0f0', border:'none', borderRadius:8, width:34, height:34, fontSize:16, cursor:'pointer', fontWeight:700 }}>›</button>
                      </div>
                      <div style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', gap:2, marginBottom:4 }}>
                        {JOURS.map(j => <div key={j} style={{ textAlign:'center', fontSize: etroit ? 11.5 : 13, fontWeight:700, color:'#999', padding: etroit ? '3px 0' : '8px 0' }}>{j}</div>)}
                      </div>
                      <div ref={calContainerRef}
                        onTouchStart={handleCalTouchStart}
                        onTouchMove={handleCalTouchMove}
                        onTouchEnd={handleCalTouchEnd}
                        style={{ overflow:'hidden', position:'relative', height: `${Math.ceil(cases.length/7)*(etroit ? 36 : 51)}px` }}
                      >
                        {renderCalGrid(cases, annee, mois, calDragX)}
                        {adjDate && renderCalGrid(adjCases, adjDate.getFullYear(), adjDate.getMonth(), calDragX + (calDragDir==='left' ? containerW : -containerW))}
                      </div>
                    </div>
                    );
                  })()}
                </div>
                {/* Colonne Midi/Soir */}
                <div style={!isMobile ? { background:'#f8f8f8', borderRadius:12, height:'100%', boxSizing:'border-box' } : {}}>
                  {!isMobile ? (
                    /* Service : le titre, puis deux boutons qui se partagent la hauteur */
                    <div style={{ display:'flex', flexDirection:'column', gap:10, padding:12, height:'100%', boxSizing:'border-box' }}>
                      <p style={{ fontSize:12, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:0, flexShrink:0 }}>Service</p>
                      {[{ id:'midi', label:'Midi', Icone:Sun, couverts:couvertsMidi },
                        { id:'soir', label:'Soir', Icone:Moon, couverts:couvertsSoir }].map(sv => {
                        const actif = calServiceSelectionne === sv.id;
                        const Icone = sv.Icone;
                        return (
                          <button key={sv.id} onClick={() => setCalServiceSelectionne(sv.id)}
                            style={{ width:'100%', flex:1, minHeight:76, padding:'10px 12px', borderRadius:12,
                              border: actif ? '2px solid #111' : '1.5px solid #eee',
                              background: actif ? '#111' : '#fff', cursor:'pointer', textAlign:'center',
                              display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap:4 }}>
                            <Icone size={20} color={actif ? '#E8C547' : '#999'} />
                            <div style={{ fontSize:16, fontWeight:800, color: actif ? '#E8C547' : '#111' }}>{sv.label}</div>
                            <div style={{ fontSize:13, color: actif ? '#bbb' : '#999' }}>{sv.couverts} couvert{sv.couverts > 1 ? 's' : ''}</div>
                          </button>
                        );
                      })}
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
                                borderColor: calServiceSelectionne === 'midi' ? '#111' : '#ddd' }}><Sun size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Midi</button>
                            <div style={{ textAlign:'center', fontSize:11, color:'#888' }}>{couvertsMidi} couvert{couvertsMidi > 1 ? 's' : ''}</div>
                          </div>
                          <div style={{ flex:1, display:'flex', flexDirection:'column', gap:4 }}>
                            <button onClick={() => setCalServiceSelectionne(calServiceSelectionne === 'soir' ? null : 'soir')}
                              style={{ height:40, borderRadius:9, border:'1.5px solid', fontSize:13, fontWeight:700, cursor:'pointer',
                                background: calServiceSelectionne === 'soir' ? '#111' : '#fff',
                                color: calServiceSelectionne === 'soir' ? '#E8C547' : '#111',
                                borderColor: calServiceSelectionne === 'soir' ? '#111' : '#ddd' }}><Moon size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Soir</button>
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
          const serviceLabel = calServiceSelectionne === 'midi' ? 'Midi' : 'Soir';
          return (
            <div style={{ background:'#fff', borderRadius:14, padding:'14px 16px', marginBottom:12, boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:16 }}>
                <div>
                  <div style={{ fontWeight:800, fontSize:17 }}>Réservations TED</div>
                  <div style={{ fontSize:13, color:'#888', marginTop:2 }}>{dateLabel} — {serviceLabel}</div>
                </div>
                <button onClick={() => telechargerTableau(calJourSelectionne, calServiceSelectionne, resasDuJour)} style={{ background:'#111', color:'#fff', border:'none', borderRadius:9, padding:'0 18px', height:38, fontSize:13, fontWeight:700, cursor:'pointer', display:'flex', alignItems:'center', gap:6 }}>
                  <Download size={14} style={{display:'inline',verticalAlign:'middle',marginRight:4}} /> Télécharger
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
                                <a href={`tel:${r.clients.tel}`} onClick={e => e.stopPropagation()} style={{ display:'inline-flex', alignItems:'center', gap:4, background:'#f0f0f0', borderRadius:6, padding:'3px 10px', fontSize:12, color:'#111', textDecoration:'none', fontWeight:600, whiteSpace:'nowrap' }}><Phone size={12} style={{display:'inline',verticalAlign:'middle'}} /> {r.clients.tel}</a>
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
                          {label:'Service', value:r.service==='midi'?<><Sun size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Midi</>:<><Moon size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Soir</>},
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
                      {r.commentaire_client && <div style={{background:'#fffbea',borderRadius:8,padding:'8px 12px',marginBottom:14,fontSize:13,color:'#666',fontStyle:'italic'}}><MessageSquare size={14} style={{display:'inline',verticalAlign:'middle',marginRight:4}} />{r.commentaire_client}</div>}
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
          <div style={{ height:'100%', display:'flex', flexDirection:'column', gap:10, minHeight:0 }}>
          {/* Prise de réservation : au-dessus du bloc, à la hauteur de « Nouvelle commande » */}
          <button onClick={()=>setShowAddResa(true)} style={{ ...btnPrimary, width:'100%', height:38, flexShrink:0, display:'flex', alignItems:'center', justifyContent:'center', gap:6 }}>
            <Plus size={16} strokeWidth={2.4} /> Nouvelle réservation
          </button>
          <div style={{ background:'#fff', borderRadius:16, border:'1.5px solid #f0f0f0', flex:1, minHeight:0, display:'flex', flexDirection:'column', overflow:'hidden', boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
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
                {calServiceSelectionne ? <> — {calServiceSelectionne==='midi'?<><Sun size={13} style={{display:'inline',verticalAlign:'middle'}} /> Midi</>:<><Moon size={13} style={{display:'inline',verticalAlign:'middle'}} /> Soir</>}</> : ''}
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
                    {r.source === 'Grand Jeux du TED'
                      ? <span style={{ background:'#E8C547', color:'#111', borderRadius:20, padding:'3px 10px', fontSize:11, fontWeight:700, flexShrink:0, whiteSpace:'nowrap', display:'inline-flex', alignItems:'center', gap:4 }}><Dices size={11} /> Jeux</span>
                      : <span style={{ background:s.bg, color:s.color, borderRadius:20, padding:'3px 10px', fontSize:11, fontWeight:700, flexShrink:0, whiteSpace:'nowrap' }}>{s.label}</span>
                    }
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

// ─── Commandes à emporter ─────────────────────────────────────────────────────

const CMD_STATUTS = [
  { id:'nouvelle',       label:'Nouvelle',       court:'Nouvelle',  bg:'#dc2626', fg:'#fff' },
  { id:'en_preparation', label:'En préparation', court:'En prépa',  bg:'#E8C547', fg:'#111' },
  { id:'prete',          label:'Prête',          court:'Prête',     bg:'#16a34a', fg:'#fff' },
  { id:'recuperee',      label:'Récupérée',      court:'Récupérée', bg:'#111111', fg:'#fff' },
  { id:'annulee',        label:'Annulée',        court:'Annulée',   bg:'#f5f5f5', fg:'#999' },
];
const cmdStatut = (id) => CMD_STATUTS.find(s => s.id === id) || CMD_STATUTS[0];
const cmdTotal  = (items) => (items || []).reduce((s, it) => s + (Number(it.prix) || 0) * (Number(it.quantite) || 1), 0);
const fmtEuro   = (n) => (Number(n) || 0).toFixed(2).replace('.', ',') + ' €';

// Rapprochement de deux numéros de téléphone saisis librement : on ne garde que
// les chiffres, et seulement les 9 derniers pour ignorer le 0 ou le +33 initial.
const cleTel = (t) => {
  const chiffres = String(t || '').replace(/\D/g, '');
  return chiffres.length >= 9 ? chiffres.slice(-9) : '';
};

function CommandesPage({ showToast, user }) {
  const [commandes, setCommandes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filtre, setFiltre] = useState('actives'); // actives | arecuperer | terminees
  const [jourSelectionne, setJourSelectionne] = useState(null); // AAAA-MM-JJ ou null
  const [service, setService] = useState(null);                 // 'midi' | 'soir' | null (toute la journée)
  const [detail, setDetail] = useState(null);
  const [editCmd, setEditCmd] = useState(null);
  const [showNouvelle, setShowNouvelle] = useState(false);
  const [showATraiter, setShowATraiter] = useState(false);
  const [showLienDropdown, setShowLienDropdown] = useState(false);
  const [autoAccept, setAutoAccept] = useState(false);
  const [delaiDefaut, setDelaiDefaut] = useState(30);
  const [commandesActives, setCommandesActives] = useState(true);
  const [motifFermeture, setMotifFermeture] = useState('');
  const [horizonJours, setHorizonJours] = useState(15);
  const [showParams, setShowParams] = useState(false);
  const [showCalendrier, setShowCalendrier] = useState(false);
  const [detailCalendrier, setDetailCalendrier] = useState(null);
  const [showStats, setShowStats] = useState(false);
  const [tick, setTick] = useState(0); // rafraîchit les comptes à rebours
  const basculeesRef = useRef(new Set()); // commandes déjà passées en « Prête » automatiquement
  const isMobile = useIsMobile();
  const etroit = useEcranEtroit();
  const LIEN_COMMANDE = 'https://ted-crm.pages.dev/commander.html';

  async function loadCommandes(silent = false) {
    if (!silent) setLoading(true);
    const { data } = await safeQuery(
      () => supabase.from('commandes').select('*, client:clients(prenom,nom,entreprise)').order('created_at', { ascending: false }).limit(500),
      { fallback: [], context: 'loadCommandes' }
    );
    // Les commandes en ligne ne sont pas rattachées à une fiche : on retrouve
    // le client par son téléphone pour afficher « Prénom Nom » plutôt que le
    // seul nom que le client a bien voulu taper.
    const orphelines = (data || []).filter(c => !c.client && c.client_tel);
    let parTel = {};
    if (orphelines.length) {
      const { data: fiches } = await safeQuery(
        () => supabase.from('clients').select('tel,prenom,nom,entreprise').limit(2000),
        { fallback: [], context: 'loadClientsPourCommandes' }
      );
      (fiches || []).forEach(f => { const k = cleTel(f.tel); if (k) parTel[k] = f; });
    }
    setCommandes((data || []).map(c =>
      (!c.client && c.client_tel && parTel[cleTel(c.client_tel)])
        ? { ...c, client: parTel[cleTel(c.client_tel)] }
        : c
    ));
    if (!silent) setLoading(false);
  }

  async function loadConfig() {
    const { data } = await safeQuery(
      () => supabase.from('commandes_config').select('cle,valeur'),
      { fallback: [], context: 'loadCommandesConfig' }
    );
    const conf = {};
    (data || []).forEach(r => { conf[r.cle] = r.valeur; });

    // Un réglage posé un autre jour est périmé : on revient au comportement normal
    const duJour = (cle) => conf[`${cle}_jour`] === dateLocale();

    setDelaiDefaut(parseInt(conf.delai_minutes) || 30);
    setMotifFermeture(conf.motif_fermeture || '');
    setHorizonJours(parseInt(conf.horizon_jours) ?? 15);
    setAutoAccept(duJour('acceptation_auto') ? conf.acceptation_auto === 'true' : conf.acceptation_auto_defaut !== 'false');
    setCommandesActives(duJour('commandes_actives') ? conf.commandes_actives !== 'false' : true);
  }

  useEffect(() => { loadCommandes(); loadConfig(); }, []);

  // Temps réel : nouvelle commande en ligne → apparaît immédiatement
  useEffect(() => {
    return resilientChannel(supabase, 'commandes-rt', (chan) => chan
      .on('postgres_changes', { event: '*', schema: 'public', table: 'commandes' }, () => loadCommandes(true))
      .on('postgres_changes', { event: '*', schema: 'public', table: 'commandes_config' }, () => loadConfig())
    );
  }, []);

  // Horloge : comptes à rebours + passage automatique en « Prête » à échéance
  useEffect(() => {
    const iv = setInterval(() => setTick(t => t + 1), 20000);
    return () => clearInterval(iv);
  }, []);

  useEffect(() => {
    const aBasculer = commandes.filter(c =>
      c.statut === 'en_preparation' && c.pret_estime_a && new Date(c.pret_estime_a) <= new Date()
      && !basculeesRef.current.has(c.id)
    );
    if (!aBasculer.length) return;
    // Verrou : si la mise à jour échoue, on ne réessaie pas en boucle
    aBasculer.forEach(c => {
      basculeesRef.current.add(c.id);
      changerStatut(c, 'prete', true);
    });
  }, [tick, commandes]);

  useEffect(() => {
    function handleClickOutside(e) {
      if (!e.target.closest('#lien-commande-dropdown')) setShowLienDropdown(false);
    }
    document.addEventListener('click', handleClickOutside);
    return () => document.removeEventListener('click', handleClickOutside);
  }, []);

  // Les deux réglages ne valent que pour la journée : on note le jour où ils
  // ont été posés, et ils reprennent leur valeur par défaut le lendemain.
  async function majConfig(cle, valeur) {
    if (cle === 'commandes_actives' || cle === 'acceptation_auto') {
      await safeQuery(
        () => supabase.from('commandes_config').upsert({ cle: `${cle}_jour`, valeur: dateLocale(), updated_at: new Date().toISOString() }, { onConflict: 'cle' }),
        { context: 'majConfigJour' }
      );
    }
    return majConfigBrut(cle, valeur);
  }

  async function majConfigBrut(cle, valeur) {
    if (cle === 'acceptation_auto') setAutoAccept(valeur === 'true');
    if (cle === 'delai_minutes') setDelaiDefaut(parseInt(valeur) || 30);
    if (cle === 'commandes_actives') setCommandesActives(valeur !== 'false');
    if (cle === 'motif_fermeture') setMotifFermeture(valeur || '');
    if (cle === 'horizon_jours') setHorizonJours(parseInt(valeur) ?? 15);
    const { error } = await safeQuery(
      () => supabase.from('commandes_config').upsert({ cle, valeur: String(valeur), updated_at: new Date().toISOString() }, { onConflict: 'cle' }),
      { context: 'majCommandesConfig' }
    );
    if (error) { showToast('Erreur d\'enregistrement du réglage', 'error'); loadConfig(); return; }
    if (cle === 'acceptation_auto') showToast(valeur === 'true' ? '✅ Acceptation automatique activée' : 'Acceptation automatique désactivée');
    if (cle === 'commandes_actives') showToast(valeur !== 'false' ? '✅ Commande en ligne rouverte' : '⏸️ Commande en ligne désactivée');
  }

  async function changerStatut(cmd, statut, silencieux = false, motifRefus = null) {
    const patch = { statut, updated_at: new Date().toISOString() };
    if (statut === 'recuperee' || statut === 'annulee') {
      patch.traited_at = new Date().toISOString();
      patch.traited_by = user?.email || null;
    }
    if (motifRefus) patch.motif_refus = motifRefus;
    setCommandes(prev => prev.map(c => c.id === cmd.id ? { ...c, ...patch } : c));
    setDetail(prev => prev && prev.id === cmd.id ? { ...prev, ...patch } : prev);
    const { error } = await safeQuery(
      () => supabase.from('commandes').update(patch).eq('id', cmd.id),
      { context: 'changerStatutCommande' }
    );
    if (error) { showToast('Erreur de mise à jour', 'error'); loadCommandes(true); return; }
    if (!silencieux) showToast(`Commande ${cmd.numero || ''} → ${cmdStatut(statut).label}`);
  }

  // Suppression définitive. La commande disparaît de la base : à réserver aux
  // saisies erronées. Pour une commande réellement annulée, préférer le statut.
  async function supprimerCommande(cmd) {
    const { error } = await safeQuery(
      () => supabase.from('commandes').delete().eq('id', cmd.id),
      { context: 'supprimerCommande' }
    );
    if (error) { showToast('Erreur lors de la suppression', 'error'); return; }
    setCommandes(prev => prev.filter(c => c.id !== cmd.id));
    setDetail(null);
    showToast(`Commande ${cmd.numero || ''} supprimée`);
  }

  // Acceptation manuelle : démarre le chrono, la suite s'enchaîne toute seule
  async function accepter(cmd, minutes) {
    const delai = parseInt(minutes) || delaiDefaut;
    const jourRetrait = cmd.date_retrait || (cmd.created_at || '').split('T')[0];
    const pourPlusTard = jourRetrait > dateLocale();
    const patch = {
      statut: 'en_preparation',
      acceptee_at: new Date().toISOString(),
      acceptee_auto: false,
      updated_at: new Date().toISOString(),
      // Le compte à rebours ne vaut que pour une commande du jour
      delai_minutes: pourPlusTard ? null : delai,
      pret_estime_a: pourPlusTard ? null : new Date(Date.now() + delai * 60000).toISOString(),
    };
    setCommandes(prev => prev.map(c => c.id === cmd.id ? { ...c, ...patch } : c));
    const { error } = await safeQuery(
      () => supabase.from('commandes').update(patch).eq('id', cmd.id),
      { context: 'accepterCommande' }
    );
    if (error) { showToast('Erreur lors de l\'acceptation', 'error'); loadCommandes(true); return; }
    showToast(pourPlusTard
      ? `✅ Commande ${cmd.numero || ''} acceptée — ${labelJour(jourRetrait)}${cmd.heure_retrait ? ' à ' + cmd.heure_retrait : ''}`
      : `✅ Commande ${cmd.numero || ''} acceptée — prête dans ~${delai} min`);
  }

  const aTraiter = commandes.filter(c => c.statut === 'nouvelle');
  const jourDe = (c) => c.date_retrait || (c.created_at || '').split('T')[0];

  // La liste principale ne montre PAS les commandes à traiter : elles passent
  // d'abord par le panneau dédié (bandeau rouge en haut).
  // La page est l'écran du service en cours : la consultation des autres jours
  // se fait dans le calendrier, sans influencer cet écran.
  //
  // Le repère n'est pas la date du calendrier mais la JOURNÉE DE SERVICE : le
  // service du soir déborde après minuit, si bien qu'à 2 h du matin on travaille
  // encore le service de la veille. La journée bascule à 6 h, pas à minuit.
  const jourSvc = jourServiceActuel();
  const svcEnCours = serviceActuel();

  // Ordre des services dans une journée, pour écarter ce qui n'a pas commencé.
  const rangSvc = (s) => (s === 'midi' ? 0 : 1);

  // Commande relevant du service en cours, ou d'un service déjà passé : une
  // commande du midi jamais terminée doit rester visible le soir venu, sinon
  // elle disparaîtrait sans que personne ne la clôture.
  const jusquAuServiceEnCours = (c) => {
    const j = jourServiceDe(c);
    if (j < jourSvc) return true;
    return j === jourSvc && rangSvc(serviceDe(c)) <= rangSvc(svcEnCours);
  };

  // « Terminées » couvre la journée de service entière, midi et soir réunis.
  const deLaJourneeDeService = (c) => jourServiceDe(c) === jourSvc;

  const listeFiltree = commandes.filter(c => {
    if (c.statut === 'nouvelle') return false;
    if (filtre === 'terminees')   return c.statut === 'recuperee' && deLaJourneeDeService(c);
    if (filtre === 'arecuperer')  return c.statut === 'prete' && jusquAuServiceEnCours(c);
    return c.statut === 'en_preparation' && jusquAuServiceEnCours(c);
  }).sort((a, b) => {
    // Les commandes terminées passent après celles encore en cours…
    const ta = estTerminee(a), tb = estTerminee(b);
    if (ta !== tb) return ta ? 1 : -1;

    // …et se classent entre elles par heure de récupération, la dernière en haut.
    if (ta && tb) {
      const ra = recupereeA(a), rb = recupereeA(b);
      if (ra && rb && ra !== rb) return ra < rb ? 1 : -1;
      if (ra && !rb) return -1;
      if (!ra && rb) return 1;
    }

    // Les commandes en cours restent groupées par jour de retrait,
    // et à l'intérieur d'un jour dans leur ordre de réception.
    const ja = jourDe(a), jb = jourDe(b);
    if (ja !== jb) return ja < jb ? -1 : 1;
    return (b.created_at || '').localeCompare(a.created_at || '');
  });

  // Compteurs des trois filtres.
  const nbEnCours    = commandes.filter(c => c.statut === 'en_preparation' && jusquAuServiceEnCours(c)).length;
  const nbARecuperer = commandes.filter(c => c.statut === 'prete' && jusquAuServiceEnCours(c)).length;
  const nbTerminees  = commandes.filter(c => c.statut === 'recuperee' && deLaJourneeDeService(c)).length;

  // Les deux tuiles couvrent la journée de service, midi et soir réunis.
  const cmdDuJour = commandes.filter(c => deLaJourneeDeService(c) && c.statut !== 'annulee');
  const caJour = cmdDuJour.reduce((s, c) => s + (Number(c.total) || 0), 0);
  const nbJour = cmdDuJour.length;

  if (loading) return <div style={{ textAlign:'center', paddingTop:80, fontSize:16, color:'#888' }}>Chargement des commandes…</div>;

  return (
    <div style={{ padding: isMobile ? '16px 14px 90px' : (etroit ? '14px 16px 28px' : '24px 28px'), minHeight:'100vh', boxSizing:'border-box', background:'#f5f5f5', display:'flex', flexDirection:'column', gap: etroit ? 10 : 12 }}>

      {/* ── En-tête ── */}
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:12, flexWrap:'wrap' }}>
        <div>
          <h1 style={{ margin:0, fontSize: isMobile ? 22 : 26, fontWeight:900, color:'#111', display:'flex', alignItems:'center', gap:10 }}>
            Click and Collect
          </h1>
          <p style={{ margin:'4px 0 0', fontSize:13, color:'#888' }}>Commandes à emporter — téléphone et en ligne</p>
        </div>
        <div style={{ display:'flex', alignItems:'center', gap:8 }}>
          <div style={{ position:'relative' }} id="lien-commande-dropdown">
            <button onClick={()=>setShowLienDropdown(v=>!v)} style={{ ...btnSecondary, height:38, display:'flex', alignItems:'center', gap:6 }}>
              <Link size={14} strokeWidth={2} /> Lien client
            </button>
            {showLienDropdown && (
              <div style={{ position:'absolute', top:'calc(100% + 8px)', right:0, background:'#fff', borderRadius:10, boxShadow:'0 8px 32px rgba(0,0,0,0.15)', padding:6, minWidth:210, zIndex:300 }}>
                <button type="button" onMouseDown={async(e)=>{ e.preventDefault(); e.stopPropagation(); try{ await navigator.clipboard.writeText(LIEN_COMMANDE); }catch{} showToast('✅ Lien copié !'); setShowLienDropdown(false); }} style={{ width:'100%', padding:'10px 14px', border:'none', background:'none', textAlign:'left', cursor:'pointer', fontSize:13, borderRadius:6, display:'flex', alignItems:'center', gap:10, color:'#111' }} onMouseEnter={e=>e.currentTarget.style.background='#f5f5f5'} onMouseLeave={e=>e.currentTarget.style.background='none'}><Copy size={15} strokeWidth={2} color="#666" /> Copier le lien</button>
                <button type="button" onMouseDown={(e)=>{ e.preventDefault(); e.stopPropagation(); window.open(LIEN_COMMANDE,'_blank'); setShowLienDropdown(false); }} style={{ width:'100%', padding:'10px 14px', border:'none', background:'none', textAlign:'left', cursor:'pointer', fontSize:13, borderRadius:6, display:'flex', alignItems:'center', gap:10, color:'#111' }} onMouseEnter={e=>e.currentTarget.style.background='#f5f5f5'} onMouseLeave={e=>e.currentTarget.style.background='none'}><ExternalLink size={15} strokeWidth={2} color="#666" /> Ouvrir la page</button>
              </div>
            )}
          </div>
          <button onClick={()=>setShowParams(true)} style={{ ...btnSecondary, height:38, display:'flex', alignItems:'center', gap:8,
            border: `1.5px solid ${commandesActives ? '#16a34a' : '#dc2626'}`,
            color: commandesActives ? '#15803d' : '#b91c1c', fontWeight:700 }}>
            <span className={commandesActives ? 'pastille-active' : 'pastille-fermee'}
              style={{ width:9, height:9, borderRadius:'50%', background: commandesActives ? '#16a34a' : '#dc2626', flexShrink:0 }} />
            {commandesActives ? 'Actif' : 'Inactif'}
          </button>
          <button onClick={()=>setShowStats(true)} style={{ ...btnSecondary, height:38, display:'flex', alignItems:'center', gap:6 }}>
            <BarChart3 size={15} strokeWidth={2} /> Statistiques
          </button>
          <button onClick={()=>setShowNouvelle(true)} style={{ ...btnPrimary, height:38, display:'flex', alignItems:'center', gap:6 }}>
            <Plus size={16} strokeWidth={2.4} /> Nouvelle commande
          </button>
        </div>
      </div>

      {/* ── Bandeau : porte d'entrée vers les commandes à traiter ── */}
      {/* ── Bandeau à traiter + compteurs à gauche, calendrier à droite ── */}
      <div style={{ display:'flex', alignItems:'stretch', gap: etroit ? 8 : 10, position:'sticky', top: isMobile ? 8 : 10, zIndex:40 }}>
        <div style={{ flex:1, minWidth:0, display:'flex', flexDirection:'column', gap: etroit ? 8 : 10 }}>
        <div onClick={()=>setShowATraiter(true)} className={aTraiter.length > 0 ? 'alarm-blink' : ''} style={{ background: aTraiter.length > 0 ? '#dc2626' : '#fff', border: aTraiter.length > 0 ? 'none' : '1.5px solid #f0f0f0', borderRadius:16, padding:'14px 20px', display:'flex', alignItems:'center', justifyContent:'space-between', cursor:'pointer', boxShadow: aTraiter.length > 0 ? '0 4px 16px rgba(220,38,38,0.3)' : '0 2px 10px rgba(0,0,0,0.08)' }}>
          <span style={{ fontSize:15, fontWeight:800, color: aTraiter.length > 0 ? '#fff' : '#111', display:'flex', alignItems:'center', gap:8 }}>
            <ClipboardList size={16} strokeWidth={2} color={aTraiter.length > 0 ? '#fff' : '#666'} /> Nouvelles commandes à traiter
          </span>
          <div style={{ display:'flex', alignItems:'center', gap:8 }}>
            {aTraiter.length > 0
              ? <span style={{ background:'#fff', color:'#dc2626', borderRadius:'50%', width:26, height:26, display:'flex', alignItems:'center', justifyContent:'center', fontSize:13, fontWeight:800 }}>{aTraiter.length}</span>
              : <span style={{ fontSize:13, color:'#999', fontWeight:600 }}>Aucune</span>}
            <span style={{ color: aTraiter.length > 0 ? '#fff' : '#ccc', fontSize:18 }}>›</span>
          </div>
        </div>
          <div style={{ display:'flex', alignItems:'stretch', gap: etroit ? 8 : 10 }}>
          <div style={{ background:'#fff', borderRadius:12, padding: etroit ? '6px 14px' : '6px 18px', boxShadow:'0 1px 4px rgba(0,0,0,0.04)', display:'flex', alignItems:'center', justifyContent:'center', gap:8, minHeight: etroit ? 42 : 52, flex:1, minWidth:0 }}>
            <span style={{ fontSize: etroit ? 16 : 20, fontWeight:900, color:'#111', lineHeight:1 }}>{nbJour}</span>
            <span style={{ fontSize: etroit ? 9.5 : 11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.4 }}>Commandes</span>
          </div>
          <div style={{ background:'#fff', borderRadius:12, padding: etroit ? '6px 14px' : '6px 18px', boxShadow:'0 1px 4px rgba(0,0,0,0.04)', display:'flex', alignItems:'center', justifyContent:'center', gap:8, minHeight: etroit ? 42 : 52, flex:1, minWidth:0 }}>
            <span style={{ fontSize: etroit ? 16 : 20, fontWeight:900, color:'#111', lineHeight:1 }}>{fmtEuro(caJour)}</span>
            <span style={{ fontSize: etroit ? 9.5 : 11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.4 }}>Total</span>
          </div>
          </div>
        </div>

        {/* Calendrier : colonne de droite, sur toute la hauteur du bloc */}
        <button onClick={()=>setShowCalendrier(true)} style={{ background:'#fff', border:'1.5px solid #f0f0f0', borderRadius:16, width: etroit ? 108 : 132, flexShrink:0, boxShadow:'0 1px 4px rgba(0,0,0,0.04)', cursor:'pointer', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap:8 }}>
          <CalendarDays size={etroit ? 24 : 28} strokeWidth={1.7} color="#666" />
          <span style={{ fontSize: etroit ? 12 : 13.5, fontWeight:700, color:'#444' }}>Calendrier</span>
        </button>
      </div>

      {/* ── Filtres ── */}
      <div style={{ display:'flex', gap:8, overflowX:'auto', paddingBottom:2, alignItems:'center' }}>
        {[
          // Le compteur reste affiché même à zéro : sur une autre date, le
          // libellé ne doit pas changer de forme sous les yeux du commerçant.
          {id:'actives',    label:`En préparation (${nbEnCours})`},
          {id:'arecuperer', label:`À récupérer (${nbARecuperer})`},
          {id:'terminees',  label:`Terminées (${nbTerminees})`},
        ].map(f => (
          <button key={f.id} onClick={()=>{
            setFiltre(f.id);
            // « En cours » ramène toujours à la journée de travail en cours,
            // toute la journée, même si le calendrier affichait autre chose.

          }} style={{ height: etroit ? 31 : 36, padding: etroit ? '0 12px' : '0 16px', borderRadius:9, fontSize: etroit ? 11.5 : 13, fontWeight:700, border:'none', flexShrink:0, cursor:'pointer', background: filtre===f.id ? '#111' : '#fff', color: filtre===f.id ? '#fff' : '#666' }}>
            {f.label}
          </button>
        ))}
      </div>

      {/* ── Liste des commandes acceptées ── */}
      {listeFiltree.length === 0 ? (
        <div style={{ background:'#fff', borderRadius:14, padding:'40px 20px', textAlign:'center', color:'#bbb', fontSize:14, boxShadow:'0 1px 4px rgba(0,0,0,0.04)' }}>
          {`Aucune commande ${filtre === 'terminees' ? 'terminée' : filtre === 'arecuperer' ? 'à récupérer' : 'en préparation'} aujourd'hui`}
        </div>
      ) : (
        <div style={{ display:'flex', flexDirection:'column', gap:10 }}>
          {listeFiltree.map(c => <CommandeCarte key={c.id} cmd={c} onOpen={()=>setDetail(c)} onStatut={changerStatut} />)}
        </div>
      )}

      {showATraiter && (
        <ATraiterPanel
          commandes={aTraiter}
          delaiDefaut={delaiDefaut}
          autoAccept={autoAccept}
          onAccepter={accepter}
          onRefuser={(c, motif)=>changerStatut(c, 'annulee', false, motif)}
          onOuvrir={(c)=>{ setShowATraiter(false); setDetail(c); }}
          onClose={()=>setShowATraiter(false)}
        />
      )}
      {detailCalendrier && !editCmd && (
        <CommandeDetail
          cmd={detailCalendrier}
          auDessus
          statutsSelonDate
          onClose={()=>setDetailCalendrier(null)}
          onStatut={(c, st)=>{ changerStatut(c, st); setDetailCalendrier(prev => prev ? { ...prev, statut: st } : prev); }}
          onEdit={(c)=>setEditCmd(c)}
          onSupprimer={(c)=>{ supprimerCommande(c); setDetailCalendrier(null); }}
        />
      )}
      {detail && !editCmd && <CommandeDetail cmd={detail} onClose={()=>setDetail(null)} onStatut={changerStatut} onEdit={(c)=>setEditCmd(c)} onSupprimer={supprimerCommande} />}
      {showNouvelle && <NouvelleCommandeModal onClose={()=>setShowNouvelle(false)} onSaved={()=>{ setShowNouvelle(false); loadCommandes(true); }} showToast={showToast} delaiDefaut={delaiDefaut} />}
      {editCmd && (
        <NouvelleCommandeModal
          cmd={editCmd}
          auDessus={!!detailCalendrier}
          onClose={()=>setEditCmd(null)}
          onSaved={()=>{ setEditCmd(null); setDetail(null); setDetailCalendrier(null); loadCommandes(true); }}
          showToast={showToast}
          delaiDefaut={delaiDefaut}
        />
      )}
      {showParams && (
        <ParametresCommandesModal
          delaiDefaut={delaiDefaut}
          commandesActives={commandesActives}
          motifFermeture={motifFermeture}
          horizonJours={horizonJours}
          onMaj={majConfig}
          onClose={()=>setShowParams(false)}
        />
      )}
      {showStats && <StatistiquesCommandesModal commandes={commandes} onClose={()=>setShowStats(false)} showToast={showToast} />}
      {showCalendrier && (
        <CalendrierCommandesModal
          commandes={commandes}
          onOuvrirCommande={(c)=>setDetailCalendrier(c)}
          onClose={()=>setShowCalendrier(false)}
        />
      )}
    </div>
  );
}

// Libellé lisible d'un jour de retrait
function labelJour(dateStr) {
  if (!dateStr) return '';
  const auj = dateLocale();
  const d = new Date(); d.setDate(d.getDate() + 1);
  const demain = dateLocale(d);
  if (dateStr === auj) return "Aujourd'hui";
  if (dateStr === demain) return 'Demain';
  return new Date(dateStr + 'T12:00:00').toLocaleDateString('fr-FR', { weekday:'long', day:'numeric', month:'long' });
}

// La nuit appartient au service du soir : une commande encore en cours à 2h
// du matin relève du service de la veille, pas de la journée qui commence.
const HEURE_FIN_DE_NUIT = 6;

// Service correspondant à une heure donnée
const serviceDeLHeure = (h) => (h >= 15 || h < HEURE_FIN_DE_NUIT) ? 'soir' : 'midi';

// Service d'une commande, d'après son heure de retrait (ou de réception).
function serviceDe(c) {
  const h = c.heure_retrait
    ? parseInt(String(c.heure_retrait).slice(0, 2), 10)
    : (c.created_at ? new Date(c.created_at).getHours() : 12);
  return serviceDeLHeure(h);
}

// Jour de service d'une commande : un retrait avant 6h relève de la veille
function jourServiceDe(c) {
  const j = c.date_retrait || (c.created_at || '').split('T')[0];
  if (!j) return '';
  const h = c.heure_retrait ? parseInt(String(c.heure_retrait).slice(0, 2), 10) : null;
  if (h !== null && h < HEURE_FIN_DE_NUIT) {
    const d = new Date(j + 'T12:00:00');
    d.setDate(d.getDate() - 1);
    return dateLocale(d);
  }
  return j;
}

// Ce qu'un client a commandé à emporter. Le chiffre d'affaires ne retient que
// les commandes réellement récupérées ; les goûts, eux, se lisent sur tout ce
// qui n'a pas été annulé.
function statsClickCollect(commandes, client) {
  const cle = cleTel(client.tel);
  const siennes = (commandes || []).filter(c =>
    (c.client_id && c.client_id === client.id) || (cle && cleTel(c.client_tel) === cle));

  const retenues = siennes.filter(c => c.statut !== 'annulee');
  const payees   = siennes.filter(c => c.statut === 'recuperee');
  const ca       = payees.reduce((s2, c) => s2 + (Number(c.total) || 0), 0);
  const enCours  = retenues.length - payees.length;

  const jour = (c) => c.date_retrait || (c.created_at || '').split('T')[0] || '';
  const parDateDesc = retenues.slice().sort((a, b) => jour(b).localeCompare(jour(a)));
  const derniere = parDateDesc[0] || null;
  const ilYA = derniere && jour(derniere)
    ? Math.round((new Date(dateLocale() + 'T12:00:00') - new Date(jour(derniere) + 'T12:00:00')) / 86400000)
    : null;

  const compte = {};
  retenues.forEach(c => (c.items || []).forEach(it => {
    const nom = (it.nom || '').trim();
    if (nom) compte[nom] = (compte[nom] || 0) + (Number(it.quantite) || 1);
  }));
  const topArticles = Object.entries(compte).sort((a, b) => b[1] - a[1]).slice(0, 5);

  const svc = { midi:0, soir:0 };
  retenues.forEach(c => { svc[serviceDe(c)] += 1; });
  const servicePrefere = (svc.midi || svc.soir)
    ? (svc.midi >= svc.soir ? 'midi' : 'soir') : null;

  return {
    nb: payees.length, enCours, ca,
    panier: payees.length ? ca / payees.length : 0,
    derniere, ilYA, jourDeLaCommande: jour,
    dernieres: parDateDesc.slice(0, 5),
    topArticles, svc, servicePrefere,
    aucune: retenues.length === 0,
  };
}

// Volet Click and Collect de la fiche client. `compact` sert la fiche mobile :
// mêmes chiffres, sans les colonnes côte à côte.
function BlocClickCollect({ stats, compact = false }) {
  const st = stats;
  const jour = st.jourDeLaCommande;
  const dateCourte = (c) => {
    const j = jour(c);
    return j ? new Date(j + 'T12:00:00').toLocaleDateString('fr-FR', { day:'numeric', month:'short', year:'numeric' }) : '—';
  };
  const STATUT_FICHE = {
    recuperee:      { bg:'#111',     color:'#fff',    label:'Récupérée' },
    prete:          { bg:'#dcfce7',  color:'#16a34a', label:'Prête' },
    en_preparation: { bg:'#fffbea',  color:'#92400e', label:'En prépa' },
    nouvelle:       { bg:'#fee2e2',  color:'#dc2626', label:'Nouvelle' },
  };

  const tuiles = [
    { label:'COMMANDES',      valeur: st.nb,               sub: st.enCours > 0 ? `+ ${st.enCours} en cours` : 'récupérées' },
    { label:'CHIFFRE D\'AFFAIRES', valeur: fmtEuro(st.ca), sub: 'commandes récupérées' },
    { label:'PANIER MOYEN',   valeur: fmtEuro(st.panier),  sub: st.nb ? `sur ${st.nb} commande${st.nb > 1 ? 's' : ''}` : '—' },
    { label:'DERNIÈRE COMMANDE', valeur: st.derniere ? dateCourte(st.derniere) : 'Jamais',
      sub: st.ilYA === null ? ''
        : st.ilYA === 0 ? "aujourd'hui"
        : st.ilYA < 0   ? `à retirer dans ${-st.ilYA} jour${-st.ilYA > 1 ? 's' : ''}`
        : `il y a ${st.ilYA} jour${st.ilYA > 1 ? 's' : ''}` },
  ];

  const enTete = (
    <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:14 }}>
      <ShoppingBag size={18} strokeWidth={2} color="#111" />
      <h3 style={{ margin:0, fontSize:15, fontWeight:800, color:'#111' }}>Click and Collect</h3>
      {st.servicePrefere && (
        <span style={{ marginLeft:'auto', display:'inline-flex', alignItems:'center', gap:5, fontSize:12, fontWeight:700,
          padding:'4px 10px', borderRadius:20,
          background: st.servicePrefere === 'midi' ? '#fffbea' : '#1e1b4b',
          color:      st.servicePrefere === 'midi' ? '#92400e' : '#c7d2fe',
          border:     st.servicePrefere === 'midi' ? '1.5px solid #fde68a' : '1.5px solid #4338ca' }}>
          {st.servicePrefere === 'midi' ? <Sun size={12} strokeWidth={2.2} /> : <Moon size={12} strokeWidth={2.2} />}
          Plutôt {st.servicePrefere} · {st.svc[st.servicePrefere]}
        </span>
      )}
    </div>
  );

  if (st.aucune) return (
    <div style={{ background:'#fff', borderRadius:16, padding:'20px 24px', marginBottom:16 }}>
      {enTete}
      <p style={{ margin:0, fontSize:13, color:'#bbb' }}>Ce client n'a jamais commandé à emporter.</p>
    </div>
  );

  const listeCommandes = (
    <div>
      <p style={{ fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 8px' }}>Dernières commandes</p>
      {st.dernieres.map((c, i) => {
        const sc = STATUT_FICHE[c.statut] || STATUT_FICHE.recuperee;
        return (
          <div key={c.id || i} style={{ display:'flex', alignItems:'center', gap:10, padding:'9px 0', borderBottom: i < st.dernieres.length - 1 ? '1px solid #f5f5f5' : 'none' }}>
            <div style={{ flex:1, minWidth:0 }}>
              <div style={{ fontSize:13, fontWeight:700, color:'#111' }}>{dateCourte(c)}</div>
              <div style={{ fontSize:11.5, color:'#999', display:'flex', alignItems:'center', gap:4 }}>
                {serviceDe(c) === 'midi' ? <Sun size={11} strokeWidth={2} /> : <Moon size={11} strokeWidth={2} />}
                {c.heure_retrait || '—'} · {(c.items || []).length} article{(c.items || []).length > 1 ? 's' : ''}
              </div>
            </div>
            <span style={{ fontSize:13.5, fontWeight:800, color:'#111', flexShrink:0 }}>{fmtEuro(c.total)}</span>
            <span style={{ background:sc.bg, color:sc.color, borderRadius:20, padding:'3px 9px', fontSize:10.5, fontWeight:700, flexShrink:0 }}>{sc.label}</span>
          </div>
        );
      })}
    </div>
  );

  const maxArticle = st.topArticles.length ? st.topArticles[0][1] : 1;
  const listeArticles = (
    <div>
      <p style={{ fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 8px' }}>Articles préférés</p>
      {st.topArticles.length === 0
        ? <p style={{ fontSize:13, color:'#bbb', margin:0 }}>Pas encore d'article commandé</p>
        : st.topArticles.map(([nom, n], i) => (
          <div key={nom} style={{ padding:'7px 0', borderBottom: i < st.topArticles.length - 1 ? '1px solid #f5f5f5' : 'none' }}>
            <div style={{ display:'flex', alignItems:'baseline', justifyContent:'space-between', gap:10, marginBottom:5 }}>
              <span style={{ fontSize:13, fontWeight:700, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{nom}</span>
              <span style={{ fontSize:12, fontWeight:800, color:'#888', flexShrink:0 }}>×{n}</span>
            </div>
            {/* La barre situe l'article par rapport au plus commandé */}
            <div style={{ height:5, borderRadius:3, background:'#f0f0f0', overflow:'hidden' }}>
              <div style={{ width:`${Math.round(n / maxArticle * 100)}%`, height:'100%', borderRadius:3, background:'#E8C547' }} />
            </div>
          </div>
        ))}
    </div>
  );

  return (
    <div style={{ background:'#fff', borderRadius:16, padding:'20px 24px', marginBottom:16 }}>
      {enTete}
      <div style={{ display:'grid', gridTemplateColumns: compact ? '1fr 1fr' : 'repeat(4, 1fr)', gap:12, marginBottom:18 }}>
        {tuiles.map(t => (
          <div key={t.label} style={{ background:'#f9f9f9', borderRadius:12, padding:'12px 14px' }}>
            <p style={{ fontSize:9.5, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 4px' }}>{t.label}</p>
            <p style={{ fontSize:16, fontWeight:900, color:'#111', margin:'0 0 2px' }}>{t.valeur}</p>
            <p style={{ fontSize:11, color:'#999', margin:0 }}>{t.sub}</p>
          </div>
        ))}
      </div>
      <div style={{ display:'grid', gridTemplateColumns: compact ? '1fr' : '1fr 1fr', gap: compact ? 18 : 28 }}>
        {listeCommandes}
        {listeArticles}
      </div>
    </div>
  );
}

// Service et jour de service en cours à l'instant présent
const serviceActuel = () => serviceDeLHeure(new Date().getHours());
function jourServiceActuel() {
  const d = new Date();
  if (d.getHours() < HEURE_FIN_DE_NUIT) d.setDate(d.getDate() - 1);
  return dateLocale(d);
}

// Une commande récupérée ou annulée ne bouge plus : elle sort du flux en cours.
const estTerminee = (c) => c.statut === 'recuperee' || c.statut === 'annulee';

// Horodatage du passage en « Récupérée » (traited_at est posé à ce moment-là).
const recupereeA = (c) => c.traited_at || c.updated_at || null;

// « Récupérée aujourd'hui à 13:42 » / « Récupérée le mardi 12 août à 19:05 »
function fmtRecuperee(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return '';
  const heure = d.toLocaleTimeString('fr-FR', { hour:'2-digit', minute:'2-digit' });
  const jour = dateLocale(d);
  if (jour === dateLocale()) return `Récupérée aujourd'hui à ${heure}`;
  const veille = new Date(); veille.setDate(veille.getDate() - 1);
  if (jour === dateLocale(veille)) return `Récupérée hier à ${heure}`;
  return `Récupérée le ${d.toLocaleDateString('fr-FR', { weekday:'long', day:'numeric', month:'long' })} à ${heure}`;
}

// Date locale au format AAAA-MM-JJ (sans décalage UTC)
function dateLocale(d = new Date()) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

// Jusqu'à quand un client peut commander à l'avance
const HORIZONS = [
  { j: 0,  label: "Aujourd'hui" },
  { j: 7,  label: '7 jours' },
  { j: 15, label: '15 jours' },
  { j: 30, label: '1 mois' },
  { j: 60, label: '2 mois' },
  { j: 90, label: '3 mois' },
];

// Motifs de fermeture de la prise de commandes
const MOTIFS_FERMETURE = [
  'Établissement fermé',
  'Service complet',
  'Trop de commandes en cours',
  'Fermeture exceptionnelle',
  'Rupture de stock',
];

// Motifs de refus proposés au commerçant (commandes)
const MOTIFS_REFUS_CMD = [
  'Trop de commandes en cours',
  'Produit indisponible',
  'Horaire de retrait impossible',
  'Établissement fermé',
  'Client injoignable',
];

// Délais proposés à l'acceptation (minutes)
const DELAIS_RAPIDES = [10, 15, 20, 25, 30, 45, 60, 90];
function fmtDelai(min) {
  const m = parseInt(min) || 0;
  if (m < 60) return `${m} min`;
  const h = Math.floor(m / 60), r = m % 60;
  return r === 0 ? `${h}h` : `${h}h${String(r).padStart(2, '0')}`;
}

// ── Compte à rebours « prête dans X min » ────────────────────────────────────
function minutesRestantes(cmd) {
  if (!cmd.pret_estime_a) return null;
  return Math.round((new Date(cmd.pret_estime_a) - new Date()) / 60000);
}

// ── Carte d'une commande déjà acceptée (même mise en page que « à traiter ») ──
function CommandeCarte({ cmd, onOpen, onStatut }) {
  const isMobile = useIsMobile();
  const etroit = useEcranEtroit();   // tablette : carte resserrée pour en voir plusieurs
  const st = cmdStatut(cmd.statut);
  const nbArticles = (cmd.items || []).reduce((s, it) => s + (Number(it.quantite) || 1), 0);
  // Prénom + nom depuis la fiche client si la commande y est rattachée
  const nomComplet = cmd.client
    ? (cmd.client.entreprise || `${cmd.client.prenom || ''} ${cmd.client.nom || ''}`.trim())
    : '';
  const nomAffiche = nomComplet || cmd.client_nom || 'Client';
  const heure = cmd.created_at ? new Date(cmd.created_at).toLocaleTimeString('fr-FR', { hour:'2-digit', minute:'2-digit' }) : '';
  const reste = minutesRestantes(cmd);
  const jourRetrait = cmd.date_retrait || (cmd.created_at || '').split('T')[0];
  const futur = jourRetrait > dateLocale();
  const actionnable = ['en_preparation', 'prete'].includes(cmd.statut);

  return (
    <div style={{ background:'#fff', borderRadius:16, padding: etroit ? '11px 13px' : '16px 18px', boxShadow:'0 1px 4px rgba(0,0,0,0.05)', borderLeft:`4px solid ${st.bg}`, display:'flex', gap: etroit ? 12 : 16, alignItems:'stretch', flexDirection: isMobile ? 'column' : 'row' }}>

      {/* Informations + détail de la commande */}
      <div style={{ flex:1, minWidth:0, cursor:'pointer' }} onClick={onOpen}>
        <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
          <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap', flex:1, minWidth:0 }}>
          <span style={{ fontSize: etroit ? 13.5 : 17, fontWeight:800, color:'#111' }}>{nomAffiche}</span>
          <span style={{ background:st.bg, color:st.fg, borderRadius:20, padding:'3px 10px', fontSize:11, fontWeight:700, whiteSpace:'nowrap' }}>{st.court}</span>
          {cmd.acceptee_auto && <span style={{ background:'#f0fdf4', color:'#16a34a', borderRadius:20, padding:'3px 9px', fontSize:11, fontWeight:700 }}>Auto</span>}
          {cmd.source === 'en_ligne'
            ? <span style={{ background:'#eff6ff', color:'#2563eb', borderRadius:20, padding:'3px 9px', fontSize:11, fontWeight:700, display:'inline-flex', alignItems:'center', gap:4 }}><MonitorSmartphone size={11} /> En ligne</span>
            : <span style={{ background:'#f5f5f5', color:'#666', borderRadius:20, padding:'3px 9px', fontSize:11, fontWeight:700, display:'inline-flex', alignItems:'center', gap:4 }}><Phone size={11} /> Téléphone</span>}
          </div>
          {/* Compte à rebours en bout de première ligne */}
          {cmd.statut === 'en_preparation' && !futur && reste !== null && (
            <span style={{ fontSize: etroit ? 12.5 : 14, fontWeight:800, color: reste <= 5 ? '#dc2626' : '#b8860b', display:'inline-flex', alignItems:'center', gap:5, whiteSpace:'nowrap', flexShrink:0 }}>
              <Clock size={etroit ? 13 : 14} strokeWidth={2.2} />
              {reste > 0 ? `Prête dans ${fmtDelai(reste)}` : 'À sortir'}
            </span>
          )}
        </div>

        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:10, marginTop: etroit ? 3 : 5 }}>
          <span style={{ fontSize: etroit ? 10.5 : 13, color:'#888', minWidth:0, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
            N° {cmd.numero || '—'} · reçue à {heure}
            {cmd.client_tel ? ` · ${cmd.client_tel}` : ''}
          </span>
          {/* Jour et heure de retrait, alignés à droite sous le compte à rebours */}
          {cmd.statut === 'recuperee' && recupereeA(cmd) ? (
            <span style={{ display:'inline-flex', alignItems:'center', gap:6, padding: etroit ? '3px 8px' : '5px 11px', borderRadius:8, background:'#f0fdf4', color:'#15803d', fontSize: etroit ? 11 : 13, fontWeight:700, whiteSpace:'nowrap', flexShrink:0 }}>
              <CircleCheck size={etroit ? 12 : 13} strokeWidth={2} /> {fmtRecuperee(recupereeA(cmd))}
            </span>
          ) : (
            <span style={{ display:'inline-flex', alignItems:'center', gap:6, padding: etroit ? '3px 8px' : '5px 11px', borderRadius:8, background: futur ? '#eff6ff' : '#f5f5f5', color: futur ? '#1d4ed8' : '#444', fontSize: etroit ? 11 : 13, fontWeight:700, whiteSpace:'nowrap', flexShrink:0 }}>
              <CalendarDays size={etroit ? 12 : 13} strokeWidth={2} />
              {labelJour(jourRetrait)}{cmd.heure_retrait ? ` · ${cmd.heure_retrait}` : ''}
            </span>
          )}
        </div>

        {/* Détail des articles, directement visible */}
        <div style={{ marginTop: etroit ? 8 : 12, background:'#fafafa', borderRadius:10, padding: etroit ? '8px 11px' : '12px 14px' }}>
          {(cmd.items || []).map((it, i) => (
            <div key={i} style={{ display:'flex', justifyContent:'space-between', gap:14, alignItems:'baseline', padding: etroit ? '5px 0' : '8px 0', borderBottom: i < (cmd.items.length - 1) ? '1px solid #efefef' : 'none' }}>
              <span style={{ color:'#111', fontSize: etroit ? 12.5 : 15.5, lineHeight:1.35 }}>
                <span style={{ fontWeight:800, marginRight:5 }}>{it.quantite || 1}×</span>{it.nom}
                {it.note ? <span style={{ display:'block', color:'#888', fontSize:13, fontStyle:'italic', marginTop:2 }}>{it.note}</span> : null}
              </span>
              <span style={{ fontWeight:400, fontSize: etroit ? 12 : 15, color:'#555', whiteSpace:'nowrap' }}>{fmtEuro((Number(it.prix)||0) * (Number(it.quantite)||1))}</span>
            </div>
          ))}
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'baseline', marginTop: etroit ? 7 : 10, paddingTop: etroit ? 7 : 10, borderTop:'2px solid #111' }}>
            <span style={{ fontSize: etroit ? 12 : 13, fontWeight:700, color:'#888' }}>{nbArticles} article{nbArticles > 1 ? 's' : ''}</span>
            <span style={{ fontSize: etroit ? 15.5 : 19, fontWeight:900, color:'#111' }}>{fmtEuro(cmd.total)}</span>
          </div>
        </div>

        {cmd.note && (
          <div style={{ marginTop:8, background:'#fffbea', border:'1.5px solid #E8C547', borderRadius:9, padding:'8px 12px', fontSize:12.5, color:'#111' }}>
            <strong>Note :</strong> {cmd.note}
          </div>
        )}

        {cmd.motif_refus && (
          <div style={{ marginTop:8, background:'#fef2f2', border:'1.5px solid #fca5a5', borderRadius:9, padding:'8px 12px', fontSize:12.5, color:'#111' }}>
            <strong>Motif du refus :</strong> {cmd.motif_refus}
          </div>
        )}
      </div>

      {/* Gros bouton d'action, à droite */}
      {actionnable && (
        <div style={{ width: isMobile ? '100%' : (etroit ? 118 : 200), flexShrink:0, display:'flex', flexDirection:'column', gap:11, alignItems:'center', justifyContent:'center', borderLeft: isMobile ? 'none' : '1px solid #f0f0f0', paddingLeft: isMobile ? 0 : (etroit ? 12 : 20) }}>
          {/* Le bouton porte la couleur du statut courant : orange tant que la
              commande se prépare, vert une fois qu'elle attend son client. */}
          {cmd.statut === 'en_preparation' && (
            <button onClick={()=>onStatut(cmd, 'prete')} style={{ width: etroit ? 82 : 128, height: etroit ? 82 : 128, flexShrink:0, border:'none', borderRadius:14, background:'#E8C547', color:'#fff', fontSize: etroit ? 11.5 : 15, fontWeight:900, lineHeight:1.15, cursor:'pointer', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap: etroit ? 5 : 8, boxShadow:'0 6px 20px rgba(232,197,71,0.4)' }}>
              <CheckCircle size={etroit ? 19 : 28} strokeWidth={2.4} />
              Marquer prête
            </button>
          )}
          {cmd.statut === 'prete' && (
            <button onClick={()=>onStatut(cmd, 'recuperee')} style={{ width: etroit ? 96 : 150, height: etroit ? 96 : 150, flexShrink:0, border:'none', borderRadius:14, background:'#16a34a', color:'#fff', fontSize: etroit ? 12.5 : 17, fontWeight:900, lineHeight:1.15, cursor:'pointer', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap: etroit ? 5 : 8, boxShadow:'0 6px 20px rgba(22,163,74,0.35)' }}>
              <BadgeCheck size={etroit ? 22 : 34} strokeWidth={2.2} />
              À récupérer
            </button>
          )}
        </div>
      )}
    </div>
  );
}

// ── Panneau « Commandes à traiter » ──────────────────────────────────────────
function ATraiterPanel({ commandes, delaiDefaut, autoAccept, onAccepter, onRefuser, onOuvrir, onClose }) {
  const [delais, setDelais] = useState({});       // délai choisi par commande
  const [choixOuvert, setChoixOuvert] = useState(null); // id de la commande en cours de choix
  const [refusCmd, setRefusCmd] = useState(null);       // commande en cours de refus
  const isMobile = useIsMobile();
  const getDelai = (id) => delais[id] != null ? parseInt(delais[id]) : (parseInt(delaiDefaut) || 30);

  return (
    <>
      <div onClick={onClose} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:4999 }} />
      <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#f5f5f5', borderRadius:20, width:'min(1040px, calc(100vw - 28px))', maxHeight:'92vh', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.25)', zIndex:5000, overflow:'hidden' }}>

        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'20px 24px 16px', borderBottom:'1px solid #e8e8e8', background:'#fff', flexShrink:0 }}>
          <div>
            <h2 style={{ margin:0, fontSize:18, fontWeight:800, color:'#111', display:'flex', alignItems:'center', gap:8 }}>
              <ClipboardList size={18} strokeWidth={2} /> Commandes à traiter
            </h2>
            <p style={{ margin:'3px 0 0', fontSize:12.5, color:'#888' }}>
              {commandes.length === 0 ? 'Tout est traité' : `${commandes.length} commande${commandes.length > 1 ? 's' : ''} en attente d'acceptation`}
            </p>
          </div>
          <button onClick={onClose} style={{ width:34, height:34, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:16, color:'#666' }}>✕</button>
        </div>

        <div style={{ padding:'16px 20px 20px', overflowY:'auto', display:'flex', flexDirection:'column', gap:12 }}>
          {autoAccept && (
            <div style={{ background:'#f0fdf4', border:'1.5px solid #86efac', borderRadius:12, padding:'11px 14px', fontSize:13, color:'#15803d', display:'flex', alignItems:'center', gap:8 }}>
              <CircleCheck size={15} strokeWidth={2} /> Acceptation automatique active — les nouvelles commandes sont acceptées seules avec ~{delaiDefaut} min de délai.
            </div>
          )}

          {commandes.length === 0 ? (
            <div style={{ background:'#fff', borderRadius:14, padding:'50px 20px', textAlign:'center', color:'#bbb', fontSize:14 }}>
              Aucune commande en attente 👌
            </div>
          ) : commandes.map(cmd => {
            const nbArticles = (cmd.items || []).reduce((s, it) => s + (Number(it.quantite) || 1), 0);
            const heure = cmd.created_at ? new Date(cmd.created_at).toLocaleTimeString('fr-FR', { hour:'2-digit', minute:'2-digit' }) : '';
            const d = getDelai(cmd.id);
            return (
              <div key={cmd.id} style={{ background:'#fff', borderRadius:16, padding:'16px 18px', boxShadow:'0 1px 4px rgba(0,0,0,0.05)', borderLeft:'4px solid #dc2626', display:'flex', gap:16, alignItems:'stretch', flexDirection: isMobile ? 'column' : 'row' }}>

                {/* Informations de la commande */}
                <div style={{ flex:1, minWidth:0, cursor:'pointer' }} onClick={()=>onOuvrir(cmd)}>
                  <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
                    <span style={{ fontSize:17, fontWeight:800, color:'#111' }}>{cmd.client_nom || 'Client'}</span>
                    {cmd.source === 'en_ligne'
                      ? <span style={{ background:'#eff6ff', color:'#2563eb', borderRadius:20, padding:'3px 9px', fontSize:11, fontWeight:700, display:'inline-flex', alignItems:'center', gap:4 }}><MonitorSmartphone size={11} /> En ligne</span>
                      : <span style={{ background:'#f5f5f5', color:'#666', borderRadius:20, padding:'3px 9px', fontSize:11, fontWeight:700, display:'inline-flex', alignItems:'center', gap:4 }}><Phone size={11} /> Téléphone</span>}
                  </div>
                  <div style={{ fontSize:13, color:'#888', marginTop:4 }}>
                    N° {cmd.numero || '—'} · reçue à {heure}
                    {cmd.client_tel ? ` · ${cmd.client_tel}` : ''}
                  </div>

                  <div style={{ display:'inline-flex', alignItems:'center', gap:7, marginTop:7, padding:'6px 12px', borderRadius:9, background: (cmd.date_retrait && cmd.date_retrait > dateLocale()) ? '#eff6ff' : '#f5f5f5', color: (cmd.date_retrait && cmd.date_retrait > dateLocale()) ? '#1d4ed8' : '#444', fontSize:13.5, fontWeight:700 }}>
                    <CalendarDays size={14} strokeWidth={2} />
                    {labelJour(cmd.date_retrait)}{cmd.heure_retrait ? ` · ${cmd.heure_retrait}` : ''}
                  </div>

                  {/* Détail des articles */}
                  <div style={{ marginTop:12, background:'#fafafa', borderRadius:10, padding:'12px 14px' }}>
                    {(cmd.items || []).map((it, i) => (
                      <div key={i} style={{ display:'flex', justifyContent:'space-between', gap:14, alignItems:'baseline', padding:'8px 0', borderBottom: i < (cmd.items.length - 1) ? '1px solid #efefef' : 'none' }}>
                        <span style={{ color:'#111', fontSize:15.5, lineHeight:1.45 }}>
                          <span style={{ fontWeight:800, marginRight:5 }}>{it.quantite || 1}×</span>{it.nom}
                          {it.note ? <span style={{ display:'block', color:'#888', fontSize:13, fontStyle:'italic', marginTop:2 }}>{it.note}</span> : null}
                        </span>
                        <span style={{ fontWeight:400, fontSize:15, color:'#555', whiteSpace:'nowrap' }}>{fmtEuro((Number(it.prix)||0) * (Number(it.quantite)||1))}</span>
                      </div>
                    ))}
                    <div style={{ display:'flex', justifyContent:'space-between', alignItems:'baseline', marginTop:10, paddingTop:10, borderTop:'2px solid #111' }}>
                      <span style={{ fontSize:13, fontWeight:700, color:'#888' }}>{nbArticles} article{nbArticles > 1 ? 's' : ''}</span>
                      <span style={{ fontSize:19, fontWeight:900, color:'#111' }}>{fmtEuro(cmd.total)}</span>
                    </div>
                  </div>

                  {cmd.note && (
                    <div style={{ marginTop:8, background:'#fffbea', border:'1.5px solid #E8C547', borderRadius:9, padding:'8px 12px', fontSize:12.5, color:'#111' }}>
                      <strong>Note :</strong> {cmd.note}
                    </div>
                  )}
                </div>

                {/* Bloc d'acceptation, à droite */}
                <div style={{ width: isMobile ? '100%' : 340, flexShrink:0, display:'flex', flexDirection:'column', gap:11, justifyContent:'center', borderLeft: isMobile ? 'none' : '1px solid #f0f0f0', paddingLeft: isMobile ? 0 : 20 }}>
                  {choixOuvert !== cmd.id ? (
                    <>
                      {/* Gros bouton d'acceptation */}
                      <button onClick={()=>setChoixOuvert(cmd.id)} style={{ width:'100%', minHeight:104, border:'none', borderRadius:16, background:'#16a34a', color:'#fff', fontSize:23, fontWeight:900, cursor:'pointer', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap:5, boxShadow:'0 6px 20px rgba(22,163,74,0.35)' }}>
                        <CheckCircle size={30} strokeWidth={2.4} />
                        Accepter
                      </button>
                      <button onClick={()=>setRefusCmd(cmd)} style={{ width:'100%', height:44, border:'1.5px solid #eee', borderRadius:11, background:'#fff', color:'#dc2626', fontSize:14.5, fontWeight:700, cursor:'pointer' }}>Refuser</button>
                    </>
                  ) : (
                    <>
                      {/* Choix du délai puis confirmation */}
                      <label style={{ fontSize:13.5, fontWeight:800, color:'#111', display:'block' }}>Prête dans :</label>
                      <div style={{ display:'grid', gridTemplateColumns:'repeat(4, 1fr)', gap:8 }}>
                        {DELAIS_RAPIDES.map(m => (
                          <button key={m} onClick={()=>setDelais(p=>({...p,[cmd.id]:m}))} style={{ height:56, borderRadius:11, border: d===m ? '2px solid #16a34a' : '1.5px solid #ddd', background: d===m ? '#16a34a' : '#fff', color: d===m ? '#fff' : '#333', fontSize:15.5, fontWeight:800, cursor:'pointer', padding:0 }}>
                            {fmtDelai(m)}
                          </button>
                        ))}
                      </div>
                      <div style={{ display:'flex', alignItems:'center', gap:6 }}>
                        <button onClick={()=>setDelais(p=>({...p,[cmd.id]:Math.max(5, d - 5)}))} style={{ width:46, height:46, borderRadius:11, border:'1.5px solid #ddd', background:'#fff', fontSize:22, fontWeight:700, cursor:'pointer', lineHeight:1 }}>−</button>
                        <div style={{ flex:1, height:46, border:'1.5px solid #ddd', borderRadius:11, display:'flex', alignItems:'center', justifyContent:'center', fontSize:16, fontWeight:800, background:'#fafafa' }}>{fmtDelai(d)}</div>
                        <button onClick={()=>setDelais(p=>({...p,[cmd.id]:Math.min(180, d + 5)}))} style={{ width:46, height:46, borderRadius:11, border:'1.5px solid #ddd', background:'#fff', fontSize:22, fontWeight:700, cursor:'pointer', lineHeight:1 }}>+</button>
                      </div>
                      <button onClick={()=>{ setChoixOuvert(null); onAccepter(cmd, d); }} style={{ width:'100%', minHeight:62, border:'none', borderRadius:14, background:'#16a34a', color:'#fff', fontSize:17, fontWeight:900, cursor:'pointer', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', gap:2, boxShadow:'0 4px 14px rgba(22,163,74,0.3)' }}>
                        <span style={{ display:'flex', alignItems:'center', gap:8 }}><CheckCircle size={19} strokeWidth={2.4} /> Accepter</span>
                        <span style={{ fontSize:11.5, fontWeight:600, opacity:0.9 }}>prête dans ~{fmtDelai(d)}</span>
                      </button>
                      <button onClick={()=>setChoixOuvert(null)} style={{ width:'100%', height:36, border:'none', background:'none', color:'#888', fontSize:13, fontWeight:600, cursor:'pointer' }}>Annuler</button>
                    </>
                  )}
                </div>
              </div>
            );
          })}
        </div>

        <div style={{ padding:'12px 24px calc(16px + env(safe-area-inset-bottom, 0px))', borderTop:'1px solid #e8e8e8', background:'#fff', flexShrink:0 }}>
          <button onClick={onClose} style={{ width:'100%', height:46, border:'1.5px solid #ddd', borderRadius:12, background:'#fff', fontSize:14, fontWeight:600, cursor:'pointer', color:'#666' }}>Fermer</button>
        </div>
      </div>

      {refusCmd && (
        <RefusCommandeModal
          cmd={refusCmd}
          onClose={()=>setRefusCmd(null)}
          onConfirm={(motif)=>{ const c = refusCmd; setRefusCmd(null); onRefuser(c, motif); }}
        />
      )}
    </>
  );
}

// ── Statistiques des commandes ───────────────────────────────────────────────
// Palette de données validée (CVD) : or foncé #b8860b + bleu #2563eb sur fond clair.
const STAT_OR = '#b8860b';
const boutonVoirTout = { height:26, padding:'0 11px', borderRadius:8, border:'1.5px solid #e4e4e4', background:'#fff', color:'#555', fontSize:11.5, fontWeight:700, cursor:'pointer', whiteSpace:'nowrap' };
const STAT_BLEU = '#2563eb';

function StatistiquesCommandesModal({ commandes, onClose, showToast }) {
  const [periode, setPeriode] = useState('7j');       // jour | 7j | mois | annee | perso
  const [triProduits, setTriProduits] = useState('qte'); // qte | ca
  const [voirTout, setVoirTout] = useState(null);        // null | 'vend' | 'dorment'
  const [rechercheTout, setRechercheTout] = useState('');
  const isMobile = useIsMobile();
  const etroit = useEcranEtroit();
  // En tablette l'écran fait ~620 px : un graphe plein format pousserait tout
  // le reste sous le pli.
  const hBarre = etroit ? 116 : 186;

  const auj = new Date();
  const aujStr = dateLocale(auj);
  const jourDe = (c) => c.date_retrait || (c.created_at || '').split('T')[0];
  const decale = (str, n) => { const d = new Date(str + 'T12:00:00'); d.setDate(d.getDate() + n); return dateLocale(d); };
  const nbJoursEntre = (a, b) => Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 86400000) + 1;

  const [debutPerso, setDebutPerso] = useState(decale(aujStr, -29));
  const [finPerso, setFinPerso] = useState(aujStr);

  // Une commande annulée n'a rien vendu : elle ne compte nulle part.
  const valides = commandes.filter(c => c.statut !== 'annulee');

  // Fenêtre observée, et la fenêtre de même durée qui la précède immédiatement :
  // c'est elle qui donne son sens à chaque chiffre.
  const fenetre = (() => {
    if (periode === 'jour')  return { debut: aujStr, fin: aujStr };
    if (periode === '7j')    return { debut: decale(aujStr, -6), fin: aujStr };
    if (periode === 'mois')  return { debut: `${auj.getFullYear()}-${String(auj.getMonth()+1).padStart(2,'0')}-01`, fin: aujStr };
    if (periode === 'annee') return { debut: `${auj.getFullYear()}-01-01`, fin: aujStr };
    const d = debutPerso <= finPerso ? debutPerso : finPerso;
    const f = debutPerso <= finPerso ? finPerso : debutPerso;
    return { debut: d, fin: f };
  })();
  const duree = Math.max(1, nbJoursEntre(fenetre.debut, fenetre.fin));
  const avant = { debut: decale(fenetre.debut, -duree), fin: decale(fenetre.debut, -1) };

  const lotDe = (f) => valides.filter(c => { const j = jourDe(c); return j >= f.debut && j <= f.fin; });
  const dansPeriode = lotDe(fenetre);
  const dansAvant   = lotDe(avant);

  // Un client = un numéro de téléphone, à défaut le nom saisi.
  const cleClient = (c) => cleTel(c.client_tel) || (c.client_nom || '').trim().toLowerCase();
  const mesure = (lot) => {
    const ca = lot.reduce((s, c) => s + (Number(c.total) || 0), 0);
    return {
      nb: lot.length, ca,
      panier: lot.length ? ca / lot.length : 0,
      clients: new Set(lot.map(cleClient).filter(Boolean)).size,
    };
  };
  const m = mesure(dansPeriode);
  const mAvant = mesure(dansAvant);

  // Évolution en pourcentage. Sans passé comparable, on n'invente rien.
  const evolution = (a, b) => (!b ? null : Math.round(((a - b) / b) * 100));

  // Premier achat de chaque client, tous temps confondus : sert à distinguer
  // un nouveau venu d'un habitué.
  const premierAchat = {};
  valides.forEach(c => {
    const k = cleClient(c); if (!k) return;
    const j = jourDe(c) || '';
    if (!premierAchat[k] || j < premierAchat[k]) premierAchat[k] = j;
  });
  const commandesParClient = {};
  valides.forEach(c => { const k = cleClient(c); if (k) commandesParClient[k] = (commandesParClient[k] || 0) + 1; });

  const clientsPeriode = {};
  dansPeriode.forEach(c => {
    const k = cleClient(c); if (!k) return;
    if (!clientsPeriode[k]) clientsPeriode[k] = { cle:k, nom: c.client_nom || c.client_tel || 'Client', nb:0, ca:0 };
    clientsPeriode[k].nb += 1;
    clientsPeriode[k].ca += Number(c.total) || 0;
  });
  const listeClients = Object.values(clientsPeriode);
  const nouveaux = listeClients.filter(cl => premierAchat[cl.cle] >= fenetre.debut).length;
  const habitues = listeClients.length - nouveaux;
  const fideles  = listeClients.filter(cl => (commandesParClient[cl.cle] || 0) >= 2).length;
  const tauxRetour = listeClients.length ? Math.round((fideles / listeClients.length) * 100) : 0;
  const topClients = listeClients.slice().sort((a, b) => b.ca - a.ca).slice(0, 5);

  // Courbe : un point par jour tant que la fenêtre tient en un mois, sinon par mois.
  const parMois = duree > 31;
  const serie = (() => {
    if (parMois) {
      const debutD = new Date(fenetre.debut + 'T12:00:00');
      const finD = new Date(fenetre.fin + 'T12:00:00');
      const points = [];
      const d = new Date(debutD.getFullYear(), debutD.getMonth(), 1);
      while (d <= finD) {
        const prefix = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
        const lot = dansPeriode.filter(c => (jourDe(c) || '').startsWith(prefix));
        points.push({
          cle: d.toLocaleDateString('fr-FR', { month:'short' }).replace('.', ''),
          titre: d.toLocaleDateString('fr-FR', { month:'long', year:'numeric' }),
          nb: lot.length, ca: lot.reduce((s, c) => s + (Number(c.total) || 0), 0),
        });
        d.setMonth(d.getMonth() + 1);
      }
      return points;
    }
    return Array.from({ length: duree }, (_, i) => {
      const str = decale(fenetre.debut, i);
      const d = new Date(str + 'T12:00:00');
      const lot = dansPeriode.filter(c => jourDe(c) === str);
      return {
        cle: duree <= 14 ? d.toLocaleDateString('fr-FR', { weekday:'short' }).slice(0, 1).toUpperCase() : String(d.getDate()),
        titre: d.toLocaleDateString('fr-FR', { weekday:'long', day:'numeric', month:'long' }),
        nb: lot.length, ca: lot.reduce((s, c) => s + (Number(c.total) || 0), 0),
      };
    });
  })();
  const maxCa = Math.max(1, ...serie.map(p => p.ca));
  const meilleur = serie.slice().sort((a, b) => b.ca - a.ca)[0];

  // Ce qui se vend, en quantité comme en argent.
  const produits = (() => {
    const compte = {};
    dansPeriode.forEach(c => (c.items || []).forEach(it => {
      const nom = (it.nom || '').trim();
      if (!nom) return;
      if (!compte[nom]) compte[nom] = { nom, qte:0, ca:0 };
      compte[nom].qte += Number(it.quantite) || 1;
      compte[nom].ca  += (Number(it.prix) || 0) * (Number(it.quantite) || 1);
    }));
    return Object.values(compte);
  })();
  const caArticles = produits.reduce((s, p) => s + p.ca, 0);
  const topProduits = produits.slice()
    .sort((a, b) => triProduits === 'ca' ? b.ca - a.ca : b.qte - a.qte)
    .slice(0, 8);
  const maxProduit = Math.max(1, ...topProduits.map(p => triProduits === 'ca' ? p.ca : p.qte));
  const tousProduits = produits.slice().sort((a, b) => triProduits === 'ca' ? b.ca - a.ca : b.qte - a.qte);

  // La carte, pour savoir ce qui n'est jamais parti sur la période.
  const [carte, setCarte] = useState([]);
  useEffect(() => {
    let vivant = true;
    (async () => {
      const { data } = await safeQuery(
        () => supabase.from('menu_produits').select('nom,prix').eq('disponible', true).limit(500),
        { fallback: [], context: 'statsCarte' }
      );
      if (!vivant) return;
      const vus = new Set();
      const liste = [];
      (data || []).forEach(p => {
        const nom = (p.nom || '').trim();
        const k = normalizeStr(nom);
        if (!nom || vus.has(k)) return;
        vus.add(k); liste.push({ nom, prix: Number(p.prix) || 0 });
      });
      setCarte(liste);
    })();
    return () => { vivant = false; };
  }, []);
  const vendus = new Set(produits.map(p => normalizeStr(p.nom)));
  const dorment = carte.filter(p => !vendus.has(normalizeStr(p.nom)));

  const nbEnLigne = dansPeriode.filter(c => c.source === 'en_ligne').length;
  const refusees = commandes.filter(c => { const j = jourDe(c); return c.statut === 'annulee' && j >= fenetre.debut && j <= fenetre.fin; }).length;

  const libelleFenetre = fenetre.debut === fenetre.fin
    ? new Date(fenetre.debut + 'T12:00:00').toLocaleDateString('fr-FR', { weekday:'long', day:'numeric', month:'long' })
    : `${new Date(fenetre.debut + 'T12:00:00').toLocaleDateString('fr-FR', { day:'numeric', month:'short' })} → ${new Date(fenetre.fin + 'T12:00:00').toLocaleDateString('fr-FR', { day:'numeric', month:'short' })}`;

  function exporterCSV() {
    const lignes = [['Numero','Date retrait','Heure','Client','Telephone','Source','Statut','Articles','Total EUR'].join(';')];
    dansPeriode.slice().sort((a, b) => (jourDe(a) < jourDe(b) ? -1 : 1)).forEach(c => {
      const detail = (c.items || []).map(it => `${it.quantite || 1}x ${it.nom}`).join(' / ');
      lignes.push([
        c.numero || '', jourDe(c) || '', c.heure_retrait || '',
        (c.client_nom || '').replace(/;/g, ','), c.client_tel || '',
        c.source === 'en_ligne' ? 'En ligne' : 'Telephone',
        cmdStatut(c.statut).label,
        detail.replace(/;/g, ','),
        String(Number(c.total) || 0).replace('.', ','),
      ].join(';'));
    });
    const blob = new Blob(['﻿' + lignes.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `commandes-${fenetre.debut}_${fenetre.fin}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showToast(`✅ ${dansPeriode.length} commande${dansPeriode.length > 1 ? 's' : ''} exportée${dansPeriode.length > 1 ? 's' : ''}`);
  }

  // Une tuile ne vaut que par sa comparaison : le chiffre seul ne dit rien.
  const Tuile = ({ valeur, libelle, delta, sub }) => {
    const hausse = delta !== null && delta !== undefined && delta >= 0;
    return (
      <div style={{ background:'#fff', border:'1.5px solid #f0f0f0', borderRadius:14, padding: etroit ? '10px 13px' : '14px 16px' }}>
        <p style={{ fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 5px' }}>{libelle}</p>
        <p style={{ fontSize: etroit ? 19 : 24, fontWeight:900, color:'#111', margin:'0 0 5px', lineHeight:1.1 }}>{valeur}</p>
        {delta === null || delta === undefined ? (
          <p style={{ fontSize:11.5, color:'#bbb', margin:0 }}>{sub || 'pas de comparable'}</p>
        ) : (
          <p style={{ fontSize:11.5, margin:0, color: hausse ? '#15803d' : '#b91c1c', fontWeight:700, display:'flex', alignItems:'center', gap:4 }}>
            {hausse ? '▲' : '▼'} {Math.abs(delta)} %
            <span style={{ color:'#bbb', fontWeight:500 }}>vs période précédente</span>
          </p>
        )}
      </div>
    );
  };

  const Carte = ({ titre, action, children }) => (
    <div style={{ background:'#fff', border:'1.5px solid #f0f0f0', borderRadius:16, padding:'16px 18px' }}>
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:10, marginBottom:14 }}>
        <h3 style={{ margin:0, fontSize:14, fontWeight:800, color:'#111' }}>{titre}</h3>
        {action}
      </div>
      {children}
    </div>
  );

  const vide = (txt) => <p style={{ margin:0, padding:'22px 0', textAlign:'center', color:'#bbb', fontSize:13.5 }}>{txt}</p>;

  return (
    <>
      <div onClick={onClose} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:5200 }} />
      <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#f7f7f7', borderRadius:20, width:'min(1240px, calc(100vw - 24px))', height:'min(940px, calc(100vh - 24px))', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.28)', zIndex:5201, overflow:'hidden' }}>

        {/* En-tête */}
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:12, padding:'20px 28px 16px', borderBottom:'1px solid #e8e8e8', background:'#fff', flexShrink:0 }}>
          <div>
            <h2 style={{ margin:0, fontSize:20, fontWeight:800, color:'#111' }}>Statistiques des commandes</h2>
            <p style={{ margin:'3px 0 0', fontSize:12.5, color:'#999', textTransform:'capitalize' }}>{libelleFenetre}</p>
          </div>
          <div style={{ display:'flex', alignItems:'center', gap:10, flexShrink:0 }}>
            <button onClick={exporterCSV} disabled={m.nb === 0}
              style={{ height:36, padding:'0 15px', border:'none', borderRadius:10, background: m.nb ? '#111' : '#f0f0f0', color: m.nb ? '#fff' : '#bbb', fontSize:13, fontWeight:800, cursor: m.nb ? 'pointer' : 'not-allowed', display:'flex', alignItems:'center', gap:7, whiteSpace:'nowrap' }}>
              <Download size={15} strokeWidth={2} /> Exporter en CSV
            </button>
            <button onClick={onClose} style={{ width:34, height:34, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:16, color:'#666' }}>✕</button>
          </div>
        </div>

        {/* Période */}
        <div style={{ padding:'14px 28px 0', background:'#f7f7f7', flexShrink:0 }}>
          <div style={{ display:'flex', gap:8 }}>
            {[{id:'jour',label:"Aujourd'hui"},{id:'7j',label:'7 jours'},{id:'mois',label:'Ce mois'},{id:'annee',label:'Cette année'},{id:'perso',label:'Dates au choix'}].map(p => (
              <button key={p.id} onClick={()=>setPeriode(p.id)} style={{ flex:1, height:38, borderRadius:10, fontSize:13, fontWeight:700, border:'none', cursor:'pointer', background: periode===p.id ? '#111' : '#fff', color: periode===p.id ? '#fff' : '#666' }}>
                {p.label}
              </button>
            ))}
          </div>
          {periode === 'perso' && (
            <div style={{ display:'flex', alignItems:'center', gap:10, marginTop:10 }}>
              <input type="date" value={debutPerso} max={aujStr} onChange={e=>setDebutPerso(e.target.value)}
                style={{ flex:1, height:40, border:'1.5px solid #e0e0e0', borderRadius:10, padding:'0 12px', fontSize:13.5, background:'#fff', outline:'none', boxSizing:'border-box' }} />
              <span style={{ fontSize:13, color:'#999', fontWeight:700 }}>→</span>
              <input type="date" value={finPerso} max={aujStr} onChange={e=>setFinPerso(e.target.value)}
                style={{ flex:1, height:40, border:'1.5px solid #e0e0e0', borderRadius:10, padding:'0 12px', fontSize:13.5, background:'#fff', outline:'none', boxSizing:'border-box' }} />
            </div>
          )}
        </div>

        <div style={{ padding:'16px 28px 24px', overflowY:'auto', display:'flex', flexDirection:'column', gap:16, flex:1, minHeight:0 }}>

          {/* Les quatre chiffres qui comptent, chacun face à la période précédente */}
          <div style={{ display:'grid', gridTemplateColumns: isMobile ? '1fr 1fr' : 'repeat(4, 1fr)', gap:10 }}>
            <Tuile libelle="Chiffre d'affaires" valeur={fmtEuro(m.ca)} delta={evolution(m.ca, mAvant.ca)} />
            <Tuile libelle="Commandes" valeur={m.nb} delta={evolution(m.nb, mAvant.nb)} />
            <Tuile libelle="Panier moyen" valeur={fmtEuro(m.panier)} delta={evolution(m.panier, mAvant.panier)} />
            <Tuile libelle="Clients" valeur={m.clients} delta={null}
              sub={m.clients ? `${nouveaux} nouveau${nouveaux > 1 ? 'x' : ''} · ${habitues} déjà venu${habitues > 1 ? 's' : ''}` : 'aucun client'} />
          </div>

          {/* Évolution du chiffre d'affaires — une seule série, pas de légende */}
          <Carte
            titre={parMois ? "Chiffre d'affaires par mois" : "Chiffre d'affaires par jour"}
            action={meilleur && meilleur.ca > 0
              ? <span style={{ fontSize:11.5, color:'#999' }}>meilleur : {meilleur.titre} · {fmtEuro(meilleur.ca)}</span>
              : null}>
            {m.nb === 0 ? vide('Aucune commande sur cette période') : (
              <div style={{ display:'flex', alignItems:'flex-end', gap: serie.length > 20 ? 3 : 6, height: hBarre + 30 }}>
                {serie.map((p, i) => (
                  <div key={i} title={`${p.titre} — ${p.nb} commande${p.nb > 1 ? 's' : ''} · ${fmtEuro(p.ca)}`}
                    style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:6, cursor:'default', minWidth:0 }}>
                    <div style={{ width:'100%', height:hBarre, display:'flex', alignItems:'flex-end' }}>
                      <div style={{ width:'100%', height: Math.max(p.ca > 0 ? 4 : 2, Math.round((p.ca / maxCa) * hBarre)), background: p.ca > 0 ? STAT_OR : '#ececec', borderRadius:'4px 4px 0 0' }} />
                    </div>
                    {serie.length <= 31 && <span style={{ fontSize:10, color:'#999', fontWeight:600 }}>{p.cle}</span>}
                  </div>
                ))}
              </div>
            )}
          </Carte>

          <div style={{ display:'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap:14 }}>

            {/* Ce qui se vend : la quantité guide les achats, l'argent guide la carte */}
            <Carte titre="Ce qui se vend" action={
              <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                <div style={{ display:'flex', gap:4, background:'#f2f2f2', borderRadius:9, padding:3 }}>
                  {[{id:'qte',label:'Quantité'},{id:'ca',label:'Chiffre'}].map(o => (
                    <button key={o.id} onClick={()=>setTriProduits(o.id)}
                      style={{ height:26, padding:'0 11px', borderRadius:7, border:'none', cursor:'pointer', fontSize:11.5, fontWeight:700,
                        background: triProduits===o.id ? '#fff' : 'transparent', color: triProduits===o.id ? '#111' : '#888' }}>
                      {o.label}
                    </button>
                  ))}
                </div>
                {produits.length > topProduits.length && (
                  <button onClick={()=>{ setRechercheTout(''); setVoirTout('vend'); }} style={boutonVoirTout}>
                    Voir les {produits.length}
                  </button>
                )}
              </div>
            }>
              {topProduits.length === 0 ? vide('Aucun article vendu') : (
                <div style={{ display:'flex', flexDirection:'column', gap:11 }}>
                  {topProduits.map(p => {
                    const part = caArticles ? Math.round((p.ca / caArticles) * 100) : 0;
                    return (
                      <div key={p.nom}>
                        <div style={{ display:'flex', justifyContent:'space-between', gap:10, marginBottom:5 }}>
                          <span style={{ fontSize:13.5, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{p.nom}</span>
                          <span style={{ fontSize:13, color:'#555', whiteSpace:'nowrap' }}>×{p.qte} · {fmtEuro(p.ca)} <span style={{ color:'#bbb' }}>({part} %)</span></span>
                        </div>
                        <div style={{ height:8, background:'#f2f2f2', borderRadius:4, overflow:'hidden' }}>
                          <div style={{ width:`${Math.round(((triProduits === 'ca' ? p.ca : p.qte) / maxProduit) * 100)}%`, height:'100%', background:STAT_OR, borderRadius:4 }} />
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </Carte>

            {/* Ce qui dort : la carte proposée mais jamais commandée */}
            <Carte titre="Personne n'en a pris" action={
              dorment.length > 14 ? (
                <button onClick={()=>{ setRechercheTout(''); setVoirTout('dorment'); }} style={boutonVoirTout}>
                  Voir les {dorment.length}
                </button>
              ) : null
            }>
              {carte.length === 0 ? vide('Carte non chargée')
                : dorment.length === 0 ? vide('Toute la carte a trouvé preneur')
                : (
                  <>
                    <p style={{ margin:'-6px 0 12px', fontSize:12, color:'#999' }}>
                      {dorment.length} produit{dorment.length > 1 ? 's' : ''} de la carte sur {carte.length} n'{dorment.length > 1 ? 'ont' : 'a'} pas été commandé{dorment.length > 1 ? 's' : ''}.
                      {duree < 28 && <span style={{ color:'#b45309' }}> Sur {duree} jour{duree > 1 ? 's' : ''}, c'est normal : regardez plutôt sur un mois ou une année.</span>}
                    </p>
                    <div style={{ display:'flex', flexWrap:'wrap', gap:7 }}>
                      {dorment.slice(0, 14).map((p, i) => (
                        <span key={p.nom + i} style={{ fontSize:12.5, color:'#666', background:'#f7f7f7', border:'1px solid #eee', borderRadius:8, padding:'5px 10px' }}>{p.nom}</span>
                      ))}
                      {dorment.length > 14 && <span style={{ fontSize:12.5, color:'#bbb', padding:'5px 4px' }}>et {dorment.length - 14} autres</span>}
                    </div>
                  </>
                )}
            </Carte>
          </div>

          <div style={{ display:'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap:14 }}>

            {/* Les clients : qui revient, et qui pèse */}
            <Carte titre="Vos clients">
              {listeClients.length === 0 ? vide('Aucun client sur cette période') : (
                <>
                  <div style={{ display:'flex', gap:20, paddingBottom:14, marginBottom:14, borderBottom:'1px solid #f2f2f2' }}>
                    <div>
                      <div style={{ fontSize:19, fontWeight:900, color:'#111' }}>{nouveaux}</div>
                      <div style={{ fontSize:10.5, color:'#999', fontWeight:700, textTransform:'uppercase', letterSpacing:0.4 }}>Nouveaux</div>
                    </div>
                    <div>
                      <div style={{ fontSize:19, fontWeight:900, color:'#111' }}>{habitues}</div>
                      <div style={{ fontSize:10.5, color:'#999', fontWeight:700, textTransform:'uppercase', letterSpacing:0.4 }}>Déjà venus</div>
                    </div>
                    <div>
                      <div style={{ fontSize:19, fontWeight:900, color:'#111' }}>{tauxRetour} %</div>
                      <div style={{ fontSize:10.5, color:'#999', fontWeight:700, textTransform:'uppercase', letterSpacing:0.4 }}>Reviennent</div>
                    </div>
                  </div>
                  <p style={{ fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 8px' }}>Meilleurs clients</p>
                  {topClients.map((cl, i) => (
                    <div key={cl.cle} style={{ display:'flex', alignItems:'center', gap:10, padding:'7px 0', borderBottom: i < topClients.length - 1 ? '1px solid #f5f5f5' : 'none' }}>
                      <span style={{ fontSize:13.5, color:'#111', flex:1, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{cl.nom}</span>
                      <span style={{ fontSize:12, color:'#999', flexShrink:0 }}>{cl.nb} cmd</span>
                      <span style={{ fontSize:13.5, fontWeight:800, color:'#111', flexShrink:0 }}>{fmtEuro(cl.ca)}</span>
                    </div>
                  ))}
                </>
              )}
            </Carte>

            {/* Le reste, en une ligne : ce n'est pas ce qu'on vient chercher ici */}
            <Carte titre="Autres repères">
              <div style={{ display:'flex', flexWrap:'wrap', gap:26 }}>
                <div>
                  <div style={{ fontSize:19, fontWeight:900, color:'#111' }}>{nbEnLigne}</div>
                  <div style={{ fontSize:10.5, color:'#999', fontWeight:700, textTransform:'uppercase', letterSpacing:0.4 }}>En ligne</div>
                </div>
                <div>
                  <div style={{ fontSize:19, fontWeight:900, color:'#111' }}>{m.nb - nbEnLigne}</div>
                  <div style={{ fontSize:10.5, color:'#999', fontWeight:700, textTransform:'uppercase', letterSpacing:0.4 }}>Téléphone</div>
                </div>
                <div>
                  <div style={{ fontSize:19, fontWeight:900, color:'#111' }}>{refusees}</div>
                  <div style={{ fontSize:10.5, color:'#999', fontWeight:700, textTransform:'uppercase', letterSpacing:0.4 }}>Annulées</div>
                </div>
                <div>
                  <div style={{ fontSize:19, fontWeight:900, color:'#111' }}>{fmtEuro(mAvant.ca)}</div>
                  <div style={{ fontSize:10.5, color:'#999', fontWeight:700, textTransform:'uppercase', letterSpacing:0.4 }}>Période précédente</div>
                </div>
              </div>
            </Carte>
          </div>
        </div>

        {/* La liste entière, quand les huit premières lignes ne suffisent plus */}
        {voirTout && (() => {
          const surLaCarte = voirTout === 'dorment';
          const filtre = normalizeStr(rechercheTout.trim());
          const lignes = (surLaCarte ? dorment : tousProduits)
            .filter(p => !filtre || normalizeStr(p.nom).includes(filtre));
          const maxTout = surLaCarte ? 1
            : Math.max(1, ...tousProduits.map(p => triProduits === 'ca' ? p.ca : p.qte));
          return (
            <>
              <div onClick={()=>setVoirTout(null)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:5400 }} />
              <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:20, width:'min(760px, calc(100vw - 24px))', height:'min(820px, calc(100vh - 24px))', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.3)', zIndex:5401, overflow:'hidden' }}>

                <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:12, padding:'18px 24px 14px', borderBottom:'1px solid #f0f0f0', flexShrink:0 }}>
                  <div style={{ minWidth:0 }}>
                    <h3 style={{ margin:0, fontSize:17, fontWeight:800, color:'#111' }}>
                      {surLaCarte ? "Personne n'en a pris" : 'Ce qui se vend'}
                    </h3>
                    <p style={{ margin:'3px 0 0', fontSize:12.5, color:'#999' }}>
                      {surLaCarte
                        ? `${dorment.length} produit${dorment.length > 1 ? 's' : ''} de la carte sans aucune commande`
                        : `${tousProduits.length} produit${tousProduits.length > 1 ? 's' : ''} vendu${tousProduits.length > 1 ? 's' : ''}`} · {libelleFenetre}
                    </p>
                  </div>
                  <button onClick={()=>setVoirTout(null)} style={{ width:34, height:34, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:16, color:'#666', flexShrink:0 }}>✕</button>
                </div>

                {/* Recherche et tri restent visibles, seule la liste défile */}
                <div style={{ display:'flex', alignItems:'center', gap:10, padding:'12px 24px', borderBottom:'1px solid #f5f5f5', flexShrink:0 }}>
                  <div style={{ position:'relative', flex:1 }}>
                    <Search size={15} strokeWidth={2} color="#999" style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }} />
                    <input value={rechercheTout} onChange={e=>setRechercheTout(e.target.value)} placeholder="Rechercher un produit…"
                      style={{ width:'100%', height:40, border:'1.5px solid #eee', borderRadius:10, padding:'0 12px 0 36px', fontSize:14, outline:'none', boxSizing:'border-box' }} />
                  </div>
                  {!surLaCarte && (
                    <div style={{ display:'flex', gap:4, background:'#f2f2f2', borderRadius:9, padding:3, flexShrink:0 }}>
                      {[{id:'qte',label:'Quantité'},{id:'ca',label:'Chiffre'}].map(o => (
                        <button key={o.id} onClick={()=>setTriProduits(o.id)}
                          style={{ height:28, padding:'0 12px', borderRadius:7, border:'none', cursor:'pointer', fontSize:12, fontWeight:700,
                            background: triProduits===o.id ? '#fff' : 'transparent', color: triProduits===o.id ? '#111' : '#888' }}>
                          {o.label}
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                <div style={{ flex:1, minHeight:0, overflowY:'auto', padding:'14px 24px 20px' }}>
                  {lignes.length === 0 ? vide('Aucun produit ne correspond')
                    : surLaCarte ? (
                      <div style={{ display:'flex', flexDirection:'column' }}>
                        {lignes.map((p, i) => (
                          <div key={p.nom + i} style={{ display:'flex', alignItems:'baseline', justifyContent:'space-between', gap:10, padding:'9px 0', borderBottom: i < lignes.length - 1 ? '1px solid #f5f5f5' : 'none' }}>
                            <span style={{ fontSize:13.5, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                              <span style={{ color:'#bbb', fontWeight:700, marginRight:7 }}>{i + 1}</span>{p.nom}
                            </span>
                            <span style={{ fontSize:13, color:'#999', whiteSpace:'nowrap', flexShrink:0 }}>{p.prix ? fmtEuro(p.prix) : '—'}</span>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div style={{ display:'flex', flexDirection:'column', gap:11 }}>
                        {lignes.map((p, i) => {
                          const part = caArticles ? Math.round((p.ca / caArticles) * 100) : 0;
                          return (
                            <div key={p.nom}>
                              <div style={{ display:'flex', justifyContent:'space-between', gap:10, marginBottom:5 }}>
                                <span style={{ fontSize:13.5, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                                  <span style={{ color:'#bbb', fontWeight:700, marginRight:7 }}>{i + 1}</span>{p.nom}
                                </span>
                                <span style={{ fontSize:13, color:'#555', whiteSpace:'nowrap' }}>×{p.qte} · {fmtEuro(p.ca)} <span style={{ color:'#bbb' }}>({part} %)</span></span>
                              </div>
                              <div style={{ height:8, background:'#f2f2f2', borderRadius:4, overflow:'hidden' }}>
                                <div style={{ width:`${Math.round(((triProduits === 'ca' ? p.ca : p.qte) / maxTout) * 100)}%`, height:'100%', background:STAT_OR, borderRadius:4 }} />
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                </div>

                <div style={{ padding:'12px 24px calc(16px + env(safe-area-inset-bottom, 0px))', borderTop:'1px solid #f0f0f0', flexShrink:0 }}>
                  <button onClick={()=>setVoirTout(null)} style={{ ...btnPrimary, width:'100%', height:46 }}>Fermer</button>
                </div>
              </div>
            </>
          );
        })()}

      </div>
    </>
  );
}

// ── Calendrier des commandes (même grille que la page Réservations) ─────────
function CalendrierCommandesModal({ commandes, onOuvrirCommande, onClose }) {
  const isMobile = useIsMobile();
  const [calDate, setCalDate] = useState(new Date());
  // Glissement horizontal d'un mois à l'autre. Trois mois sont montés en
  // permanence sur une piste, déplacée directement dans le DOM : aucun rendu
  // React pendant le geste, donc aucun scintillement.
  const grilleRef = useRef(null);
  const pisteRef = useRef(null);
  const swipeX0 = useRef(null);
  const swipeY0 = useRef(null);
  const glisseRef = useRef(false);
  const animeRef = useRef(false);
  // Service consulté : celui déjà choisi, sinon celui en cours à cette heure-ci.
  const [svcChoisi, setSvcChoisi] = useState(serviceActuel());
  const [jourLocal, setJourLocal] = useState(dateLocale());
  // Même enchaînement que le calendrier de « Nouvelle réservation » :
  // flash sur la date choisie (200 ms), puis fermeture animée (300 ms).
  const [dateFlash, setDateFlash] = useState(null);
  const [calFermeture, setCalFermeture] = useState(false);
  const timersRef = useRef([]);
  useEffect(() => () => timersRef.current.forEach(clearTimeout), []);

  // Un clic sélectionne le jour (avec le flash) sans refermer : la validation
  // se fait avec le bouton du bas, une fois le service choisi.
  function choisirJour(iso) {
    setJourLocal(iso);
    setDateFlash(iso);
    timersRef.current.push(setTimeout(() => setDateFlash(null), 220));
  }

  const JOURS = ['Lun','Mar','Mer','Jeu','Ven','Sam','Dim'];
  const MOIS = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];

  const annee = calDate.getFullYear();
  const mois = calDate.getMonth();
  const premierJour = new Date(annee, mois, 1);
  const dernierJour = new Date(annee, mois + 1, 0);
  const debutSemaine = (premierJour.getDay() + 6) % 7;

  // Répartition midi / soir par jour. Même bascule que les réservations : 15h.
  // Sans heure de retrait, on se rabat sur l'heure de réception de la commande.
  // Mémorisé : ce calcul parcourt toutes les commandes, il ne doit pas se
  // refaire à chaque image du glissement.
  const parJour = useMemo(() => {
    const acc = {};
    commandes.filter(c => c.statut !== 'annulee').forEach(c => {
      const j = c.date_retrait || (c.created_at || '').split('T')[0];
      if (!j) return;
      const h = c.heure_retrait
        ? parseInt(String(c.heure_retrait).slice(0, 2), 10)
        : (c.created_at ? new Date(c.created_at).getHours() : 12);
      if (!acc[j]) acc[j] = { midi: 0, soir: 0, total: 0 };
      acc[j][h < 15 ? 'midi' : 'soir'] += 1;
      acc[j].total += 1;
    });
    return acc;
  }, [commandes]);

  const grilleDuMois = (a, m) => {
    const debut = (new Date(a, m, 1).getDay() + 6) % 7;
    const out = [];
    for (let i = 0; i < debut; i++) out.push(null);
    for (let d = 1; d <= new Date(a, m + 1, 0).getDate(); d++) out.push(d);
    while (out.length % 7 !== 0) out.push(null);
    return out;
  };
  const cases = useMemo(() => grilleDuMois(annee, mois), [annee, mois]);

  const today = new Date();
  const changerMois = (delta) => setCalDate(new Date(annee, mois + delta, 1));

  // Les deux mois voisins sont montés en permanence : plus rien n'apparaît
  // ni ne disparaît pendant le glissement.
  const casesPrecedent = useMemo(() => grilleDuMois(annee, mois - 1), [annee, mois]);
  const casesSuivant = useMemo(() => grilleDuMois(annee, mois + 1), [annee, mois]);
  const moisPrecedent = new Date(annee, mois - 1, 1);
  const moisSuivant = new Date(annee, mois + 1, 1);

  // Le glissement écrit directement dans le DOM : aucun rendu React tant que
  // le doigt bouge, donc pas de recalcul ni de scintillement.
  const largeurGrille = () => grilleRef.current?.offsetWidth || 320;
  const glisserA = (x, avecTransition) => {
    const p = pisteRef.current;
    if (!p) return;
    p.style.transition = avecTransition ? 'transform 0.26s cubic-bezier(0.33,1,0.68,1)' : 'none';
    p.style.transform = `translate3d(${x - largeurGrille()}px,0,0)`;
  };

  const debutSwipe = (e) => {
    if (animeRef.current) return;
    swipeX0.current = e.touches[0].clientX;
    swipeY0.current = e.touches[0].clientY;
    glisseRef.current = false;
  };
  const pendantSwipe = (e) => {
    if (swipeX0.current === null || animeRef.current) return;
    const dx = e.touches[0].clientX - swipeX0.current;
    const dy = e.touches[0].clientY - swipeY0.current;
    if (!glisseRef.current && Math.abs(dy) > Math.abs(dx)) return;
    e.preventDefault();
    glisseRef.current = true;
    glisserA(dx, false);
  };
  const finSwipe = (e) => {
    if (swipeX0.current === null) return;
    const dx = e.changedTouches[0].clientX - swipeX0.current;
    const dy = e.changedTouches[0].clientY - swipeY0.current;
    swipeX0.current = null;
    if (!glisseRef.current) return;
    glisseRef.current = false;

    const L = largeurGrille();
    if (Math.abs(dy) <= Math.abs(dx) && Math.abs(dx) > L * 0.28) {
      const sens = dx < 0 ? 1 : -1;
      animeRef.current = true;
      glisserA(sens > 0 ? -L : L, true);
      timersRef.current.push(setTimeout(() => {
        // Le nouveau mois est déjà monté : on recentre sans transition
        setCalDate(new Date(annee, mois + sens, 1));
        animeRef.current = false;
      }, 260));
    } else {
      glisserA(0, true);
    }
  };

  // À chaque changement de mois, la piste revient au centre sans animation
  useEffect(() => { glisserA(0, false); }, [annee, mois]);

  return (
    <>
      <div onClick={onClose} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:5200, opacity: calFermeture ? 0 : 1, transition:'opacity 0.3s' }} />
      {/* Conteneur de centrage : la classe cal-fermeture anime `transform`,
          elle ne peut donc pas cohabiter avec un centrage par translate(-50%,-50%). */}
      <div style={{ position:'fixed', inset:0, display:'flex', alignItems:'center', justifyContent:'center', zIndex:5201, pointerEvents:'none' }}>
      <div onClick={e=>e.stopPropagation()} className={calFermeture ? 'cal-fermeture' : ''} style={{ background:'#fff', borderRadius:20, width:'min(1120px, calc(100vw - 20px))', height:'min(760px, calc(100vh - 20px))', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.28)', overflow:'hidden', pointerEvents:'auto' }}>

        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'14px 20px 12px', borderBottom:'1px solid #f0f0f0', flexShrink:0 }}>
          <h2 style={{ margin:0, fontSize:18, fontWeight:800, color:'#111', display:'flex', alignItems:'center', gap:9 }}>
            <CalendarDays size={19} strokeWidth={2} /> Calendrier des commandes
          </h2>
          <button onClick={onClose} style={{ width:34, height:34, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:16, color:'#666' }}>✕</button>
        </div>

        <div style={{ padding:'14px 22px 18px', overflow:'hidden', flex:1, minHeight:0, display:'flex', flexDirection: isMobile ? 'column' : 'row', gap:14 }}>
          <div style={{ background:'#f8f8f8', borderRadius:14, padding:14, flex:1, display:'flex', flexDirection:'column', minWidth:0, minHeight:0 }}>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:10, flexShrink:0 }}>
              <button onClick={()=>changerMois(-1)} style={{ background:'#f0f0f0', border:'none', borderRadius:9, width:38, height:38, fontSize:18, cursor:'pointer', fontWeight:700 }}>‹</button>
              <span style={{ fontWeight:800, fontSize:19 }}>{MOIS[mois]} {annee}</span>
              <button onClick={()=>changerMois(1)} style={{ background:'#f0f0f0', border:'none', borderRadius:9, width:40, height:40, fontSize:18, cursor:'pointer', fontWeight:700 }}>›</button>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', gap:3, marginBottom:6, flexShrink:0 }}>
              {JOURS.map(j => <div key={j} style={{ textAlign:'center', fontSize:13, fontWeight:700, color:'#999', padding:'5px 0' }}>{j}</div>)}
            </div>
            <div ref={grilleRef}
              onTouchStart={debutSwipe} onTouchMove={pendantSwipe} onTouchEnd={finSwipe}
              style={{ position:'relative', overflow:'hidden', flex:1, minHeight:0, touchAction:'pan-y' }}>
              {/* Piste large de trois mois, centrée sur le mois courant */}
              <div ref={pisteRef} style={{ position:'absolute', top:0, left:0, height:'100%', width:'300%', display:'flex', willChange:'transform' }}>
              {[{ a:moisPrecedent.getFullYear(), m:moisPrecedent.getMonth(), cs:casesPrecedent },
                { a:annee, m:mois, cs:cases },
                { a:moisSuivant.getFullYear(), m:moisSuivant.getMonth(), cs:casesSuivant },
              ].map((vue, vi) => (
            <div key={vi} style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', gridTemplateRows:`repeat(${Math.ceil(vue.cs.length/7)}, 1fr)`, gap:5, width:'33.3333%', height:'100%', flexShrink:0 }}>
              {vue.cs.map((d, i) => {
                if (!d) return <div key={i} />;
                const annee = vue.a, mois = vue.m;
                const iso = `${annee}-${String(mois+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
                const svc = parJour[iso];
                const isToday = today.getFullYear()===annee && today.getMonth()===mois && today.getDate()===d;
                const isSelected = jourLocal === iso || dateFlash === iso;
                const estPasse = new Date(iso) < new Date(new Date().setHours(0,0,0,0));
                return (
                  <button key={i} className={dateFlash === iso ? 'date-flash' : ''} onClick={()=>{ if (vi !== 1) { setCalDate(new Date(vue.a, vue.m, 1)); return; } choisirJour(iso); }}
                    style={{ textAlign:'center', minHeight:0, borderRadius:10, cursor:'pointer', position:'relative',
                      display:'flex', alignItems:'center', justifyContent:'center',
                      border: isToday && !isSelected ? '2px solid #E8C547' : '2px solid #f0f0f0',
                      background: isSelected ? '#111' : isToday ? '#fffbea' : '#fff',
                      color: isSelected ? '#fff' : '#111',
                      fontWeight: isSelected ? 800 : isToday ? 900 : 600, fontSize:16,
                      boxSizing:'border-box', opacity: estPasse ? 0.45 : 1,
                      touchAction:'manipulation', WebkitTapHighlightColor:'transparent',
                      transition:'background 0.15s' }}>
                    {d}
                    {/* Compteurs par service, discrets, ancrés en bas de la case */}
                    {svc && (svc.midi || svc.soir) && (
                      <span style={{ position:'absolute', left:0, right:0, bottom:3, display:'flex', alignItems:'center', justifyContent:'center', gap:5, fontSize:8.5, fontWeight:800, lineHeight:1, pointerEvents:'none' }}>
                        {['midi','soir'].filter(sv => svc[sv]).map(sv => {
                          const enAvant = svcChoisi === sv;
                          const teinte = sv === 'midi'
                            ? (isSelected ? '#E8C547' : '#b8860b')
                            : (isSelected ? '#93c5fd' : '#2563eb');
                          return (
                            <span key={sv} style={{ display:'inline-flex', alignItems:'center', gap:1.5, color:teinte, opacity: enAvant ? 1 : 0.4 }}>
                              {sv === 'midi' ? <Sun size={8} strokeWidth={3} /> : <Moon size={8} strokeWidth={3} />}{svc[sv]}
                            </span>
                          );
                        })}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
              ))}
              </div>
            </div>
          </div>

          {/* Panneau de droite : service, puis les commandes du jour choisi */}
          <div style={{ width: isMobile ? '100%' : 380, flexShrink:0, display:'flex', flexDirection:'column', gap:10, minHeight:0 }}>

            {/* Service — chaque camp annonce le nombre de commandes du jour consulté */}
            {(() => {
              const duJourConsulte = commandes.filter(c =>
                (c.date_retrait || (c.created_at || '').split('T')[0]) === jourLocal && c.statut !== 'annulee');
              const nbSvc = { midi:0, soir:0 };
              duJourConsulte.forEach(c => { nbSvc[serviceDe(c)] += 1; });
              return (
            <div style={{ display:'flex', alignItems:'center', gap:10, flexShrink:0 }}>
              <span style={{ fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1 }}>Service</span>
              <div onClick={()=>setSvcChoisi(svcChoisi === 'midi' ? 'soir' : 'midi')}
                style={{ position:'relative', display:'flex', width:200, height:38, borderRadius:10, background:'#f0f0f0', cursor:'pointer', userSelect:'none', flexShrink:0 }}>
                <span style={{ position:'absolute', top:3, left: svcChoisi === 'midi' ? 3 : 100, width:97, height:32, borderRadius:8, background:'#111', transition:'left 0.22s cubic-bezier(0.4,0,0.2,1)' }} />
                {[{id:'midi',label:'Midi',Icone:Sun},{id:'soir',label:'Soir',Icone:Moon}].map(o => {
                  const actif = svcChoisi === o.id;
                  const Icone = o.Icone;
                  return (
                    <span key={o.id} style={{ position:'relative', width:100, height:38, display:'flex', alignItems:'center', justifyContent:'center', gap:4,
                      fontSize:13, fontWeight:800, color: actif ? '#E8C547' : '#888', transition:'color 0.22s' }}>
                      <Icone size={14} strokeWidth={2.2} />{o.label}
                      <span style={{ fontWeight:700, opacity: actif ? 0.85 : 0.7 }}>({nbSvc[o.id]})</span>
                    </span>
                  );
                })}
              </div>
            </div>
              );
            })()}

            {/* Jour consulté */}
            <div style={{ fontSize:14.5, fontWeight:800, color:'#111', textTransform:'capitalize', flexShrink:0 }}>
              {new Date(jourLocal + 'T12:00:00').toLocaleDateString('fr-FR', { weekday:'long', day:'numeric', month:'long' })}
            </div>

            {/* Commandes du jour et du service */}
            {(() => {
              const liste = commandes
                .filter(c => (c.date_retrait || (c.created_at || '').split('T')[0]) === jourLocal)
                .filter(c => c.statut !== 'annulee' && serviceDe(c) === svcChoisi)
                .sort((a, b) => (a.heure_retrait || '99:99').localeCompare(b.heure_retrait || '99:99'));
              const total = liste.reduce((s2, c) => s2 + (Number(c.total) || 0), 0);
              return (
                <>
                  <div style={{ display:'flex', alignItems:'baseline', justifyContent:'space-between', flexShrink:0 }}>
                    <span style={{ fontSize:12.5, color:'#888', fontWeight:600 }}>{liste.length} commande{liste.length > 1 ? 's' : ''}</span>
                    <span style={{ fontSize:15.5, fontWeight:900, color:'#111' }}>{fmtEuro(total)}</span>
                  </div>
                  <div style={{ flex:1, minHeight:0, overflowY:'auto', display:'flex', flexDirection:'column', gap:8 }}>
                    {liste.length === 0 ? (
                      <p style={{ margin:0, padding:'28px 10px', textAlign:'center', fontSize:13.5, color:'#bbb' }}>
                        Aucune commande le {labelJour(jourLocal).toLowerCase()} au service du {svcChoisi}
                      </p>
                    ) : liste.map(c => {
                      const st = cmdStatut(c.statut);
                      const nomC = c.client ? (c.client.entreprise || `${c.client.prenom || ''} ${c.client.nom || ''}`.trim()) : '';
                      return (
                        <button key={c.id} onClick={()=>onOuvrirCommande(c)}
                          style={{ textAlign:'left', border:'1.5px solid #eee', borderLeft:`4px solid ${st.bg}`, borderRadius:11, background:'#fff', padding:'10px 12px', cursor:'pointer', display:'flex', alignItems:'center', gap:10 }}>
                          <span style={{ flex:1, minWidth:0 }}>
                            <span style={{ display:'flex', alignItems:'center', gap:7 }}>
                              <span style={{ fontSize:14, fontWeight:800, color:'#111', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{nomC || c.client_nom || 'Client'}</span>
                              <span style={{ background:st.bg, color:st.fg, borderRadius:20, padding:'2px 8px', fontSize:10, fontWeight:700, whiteSpace:'nowrap' }}>{st.court}</span>
                            </span>
                            <span style={{ display:'block', fontSize:12, color:'#888', marginTop:2 }}>
                              N° {c.numero || '—'}{c.heure_retrait ? ` · ${c.heure_retrait}` : ''}
                            </span>
                          </span>
                          <span style={{ fontSize:14, fontWeight:800, color:'#111', whiteSpace:'nowrap' }}>{fmtEuro(c.total)}</span>
                          <ChevronRight size={15} color="#ccc" />
                        </button>
                      );
                    })}
                  </div>
                </>
              );
            })()}
          </div>


        </div>

        <div style={{ padding:'10px 20px calc(12px + env(safe-area-inset-bottom, 0px))', borderTop:'1px solid #f0f0f0', flexShrink:0 }}>
          <button onClick={onClose} style={{ width:'100%', height:48, border:'1.5px solid #ddd', borderRadius:12, background:'#fff', fontSize:14.5, fontWeight:600, cursor:'pointer', color:'#666' }}>
            Fermer
          </button>
        </div>
      </div>
      </div>
    </>
  );
}

// ── Paramètres des commandes ─────────────────────────────────────────────────
function ParametresCommandesModal({ delaiDefaut, commandesActives, motifFermeture, horizonJours, onMaj, onClose }) {
  const [motif, setMotif] = useState(motifFermeture || '');
  const [infoAuto, setInfoAuto] = useState(false);
  const d = parseInt(delaiDefaut) || 30;

  // Désactiver n'est jamais immédiat : on annonce d'abord ce que ça implique
  const [confirmDesactiver, setConfirmDesactiver] = useState(null);
  const CONSEQUENCES = {
    commande: {
      titre: 'Désactiver la commande en ligne ?',
      points: [
        'Les clients ne pourront plus passer commande depuis le lien public.',
        'La page de commande affichera le motif que vous choisissez ci-dessous.',
        'Les commandes déjà reçues ne sont pas touchées.',
        'La commande en ligne se réactivera d\'elle-même demain.',
      ],
      action: () => fermerPrise(motif),
    },
  };

  function fermerPrise(m) {
    onMaj('motif_fermeture', m || '');
    onMaj('commandes_actives', 'false');
  }
  function rouvrirPrise() {
    onMaj('commandes_actives', 'true');
    onMaj('motif_fermeture', '');
    setMotif('');
  }

  return (
    <>
      <div onClick={onClose} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:5200 }} />
      <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:20, width:'min(560px, calc(100vw - 32px))', maxHeight:'90vh', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.28)', zIndex:5201, overflow:'hidden' }}>

        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'20px 24px 16px', borderBottom:'1px solid #f0f0f0', flexShrink:0 }}>
          <h2 style={{ margin:0, fontSize:18, fontWeight:800, color:'#111', display:'flex', alignItems:'center', gap:9 }}>
            {/* Même pastille que le bouton de la page : verte au ralenti, rouge si fermé */}
            <span className={commandesActives ? 'pastille-active' : 'pastille-fermee'}
              style={{ width:11, height:11, borderRadius:'50%', flexShrink:0,
                background: commandesActives ? '#16a34a' : '#dc2626' }} />
            Statut du jour
          </h2>
          <button onClick={onClose} style={{ width:34, height:34, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:16, color:'#666' }}>✕</button>
        </div>

        <div style={{ padding:'18px 24px 22px', overflowY:'auto', display:'flex', flexDirection:'column', gap:22 }}>

          {/* ── Commande en ligne ── */}
          <div>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:14, marginBottom:12 }}>
              <div style={{ minWidth:0 }}>
                <div style={{ fontSize:15, fontWeight:800, color:'#111', display:'flex', alignItems:'center', gap:8 }}>
                  <ShoppingBag size={17} strokeWidth={2} color={commandesActives ? '#16a34a' : '#dc2626'} /> Commande en ligne <span style={{ fontSize:12.5, fontWeight:600, color:'#999' }}>(du jour)</span>
                </div>
                {!commandesActives && (
                  <div style={{ fontSize:12.5, color:'#b91c1c', marginTop:4, lineHeight:1.5, fontWeight:600 }}>
                    Désactivée temporairement — les clients voient le motif ci-dessous.
                  </div>
                )}
              </div>
              {commandesActives && (
                <button onClick={()=>setConfirmDesactiver('commande')} style={{ flexShrink:0, height:40, padding:'0 16px', borderRadius:11, border:'1.5px solid #fca5a5', background:'#fff', color:'#b91c1c', fontSize:13, fontWeight:800, cursor:'pointer', whiteSpace:'nowrap' }}>
                  Désactiver temporairement
                </button>
              )}
            </div>

            {!commandesActives && motif && (
              <div style={{ background:'#fef2f2', border:'1.5px solid #fecaca', borderRadius:12, padding:'10px 14px', display:'flex', alignItems:'center', gap:9 }}>
                <AlertCircle size={15} strokeWidth={2} color="#b91c1c" />
                <span style={{ fontSize:13, color:'#b91c1c' }}>
                  Motif affiché au client : <strong>{motif}</strong>
                </span>
              </div>
            )}

            {/* Fermé : aucun réglage à faire, une seule action possible */}
            {!commandesActives && (
              <button onClick={rouvrirPrise} style={{ width:'100%', height:60, marginTop:16, border:'none', borderRadius:14, background:'#16a34a', color:'#fff', fontSize:16, fontWeight:800, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:9, boxShadow:'0 3px 12px rgba(22,163,74,0.28)' }}>
                <CircleCheck size={20} strokeWidth={2.2} /> Réactiver les commandes
              </button>
            )}
          </div>

          {commandesActives && (<>
          <div style={{ height:1, background:'#f0f0f0' }} />

          {/* ── Délai d'acceptation automatique ── */}
          {/* Ce n'est plus un interrupteur : couper l'acceptation automatique
              relève des paramètres, pas du statut du jour. Ici on règle le délai. */}
          <div>
            <div style={{ display:'flex', alignItems:'center', gap:7, marginBottom:4 }}>
              <label style={{ fontSize:13.5, fontWeight:800, color:'#111' }}>Délai d'acceptation automatique</label>
              {/* Bulle accrochée au bouton : elle sort au-dessus du reste du formulaire */}
              <span style={{ position:'relative', display:'inline-flex', flexShrink:0 }}>
                <button onClick={()=>setInfoAuto(v=>!v)} aria-label="À propos de l'acceptation automatique"
                  style={{ width:20, height:20, borderRadius:'50%', border:'none', padding:0, cursor:'pointer',
                    background: infoAuto ? '#111' : '#f0f0f0', color: infoAuto ? '#fff' : '#888',
                    display:'flex', alignItems:'center', justifyContent:'center' }}>
                  <Info size={13} strokeWidth={2.4} />
                </button>
                {infoAuto && (
                  <>
                    <div onClick={()=>setInfoAuto(false)} style={{ position:'fixed', inset:0, zIndex:5210 }} />
                    <div onClick={()=>setInfoAuto(false)}
                      style={{ position:'absolute', top:'calc(100% + 9px)', left:-8, width:260, background:'#111', color:'#fff',
                        borderRadius:12, padding:'11px 14px', fontSize:12.5, lineHeight:1.6, cursor:'pointer',
                        boxShadow:'0 12px 32px rgba(0,0,0,0.28)', zIndex:5211 }}>
                      <span style={{ position:'absolute', top:-5, left:14, width:10, height:10, background:'#111', transform:'rotate(45deg)', borderRadius:2 }} />
                      Pour désactiver l'acceptation automatique, rendez-vous dans les paramètres :
                      ce réglage est durable et ne se règle pas depuis le statut du jour.
                    </div>
                  </>
                )}
              </span>
            </div>
            <div style={{ fontSize:12.5, color:'#888', marginBottom:10, lineHeight:1.5 }}>
              Délai annoncé au client quand la commande est acceptée toute seule.
            </div>
            <>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(4, 1fr)', gap:8, marginBottom:10 }}>
              {DELAIS_RAPIDES.map(m => (
                <button key={m} onClick={()=>onMaj('delai_minutes', m)} style={{ height:52, borderRadius:11, border: d===m ? '2px solid #16a34a' : '1.5px solid #ddd', background: d===m ? '#16a34a' : '#fff', color: d===m ? '#fff' : '#333', fontSize:15, fontWeight:800, cursor:'pointer', padding:0 }}>
                  {fmtDelai(m)}
                </button>
              ))}
            </div>
            <div style={{ display:'flex', alignItems:'center', gap:8 }}>
              <button onClick={()=>onMaj('delai_minutes', Math.max(5, d - 5))} style={{ width:52, height:48, borderRadius:11, border:'1.5px solid #ddd', background:'#fff', fontSize:23, fontWeight:700, cursor:'pointer', color:'#111', lineHeight:1 }}>−</button>
              <div style={{ flex:1, height:48, border:'1.5px solid #ddd', borderRadius:11, display:'flex', alignItems:'center', justifyContent:'center', fontSize:17, fontWeight:800, background:'#fafafa' }}>{fmtDelai(d)}</div>
              <button onClick={()=>onMaj('delai_minutes', Math.min(180, d + 5))} style={{ width:52, height:48, borderRadius:11, border:'1.5px solid #ddd', background:'#fff', fontSize:23, fontWeight:700, cursor:'pointer', color:'#111', lineHeight:1 }}>+</button>
            </div>
            </>
          </div>
          </>)}

        </div>

        <div style={{ padding:'14px 24px calc(18px + env(safe-area-inset-bottom, 0px))', borderTop:'1px solid #f0f0f0', flexShrink:0 }}>
          <button onClick={onClose} style={{ ...btnPrimary, width:'100%', height:48 }}>Terminé</button>
        </div>
      </div>

      {/* Confirmation : ce que la désactivation implique concrètement */}
      {confirmDesactiver && (() => {
        const c = CONSEQUENCES[confirmDesactiver];
        return (
          <>
            <div onClick={()=>setConfirmDesactiver(null)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.45)', zIndex:5300 }} />
            <div style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:18, width:'min(460px, calc(100vw - 32px))', padding:'24px 26px', boxShadow:'0 32px 80px rgba(0,0,0,0.3)', zIndex:5301 }}>
              <h3 style={{ margin:'0 0 12px', fontSize:17, fontWeight:800, color:'#111' }}>{c.titre}</h3>
              <ul style={{ margin:'0 0 18px', padding:'0 0 0 18px', fontSize:13.5, color:'#555', lineHeight:1.65 }}>
                {c.points.map((p, i) => <li key={i}>{p}</li>)}
              </ul>

              {/* Commande en ligne : le motif vaut confirmation, un clic suffit */}
              {confirmDesactiver === 'commande' && (
                <>
                  <p style={{ margin:'0 0 8px', fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5 }}>Motif affiché au client</p>
                  <div style={{ display:'flex', flexDirection:'column', gap:7, marginBottom:18 }}>
                    {MOTIFS_FERMETURE.map(m => (
                      <button key={m} onClick={()=>{ setMotif(m); fermerPrise(m); setConfirmDesactiver(null); }}
                        style={{ height:46, borderRadius:11, border:'1.5px solid #eee', background:'#fff', color:'#111', fontSize:14, fontWeight:700, cursor:'pointer', textAlign:'left', padding:'0 14px' }}>
                        {m}
                      </button>
                    ))}
                  </div>
                </>
              )}

              <div style={{ display:'flex', gap:10 }}>
                <button onClick={()=>setConfirmDesactiver(null)} style={{ flex:1, height:48, border:'1.5px solid #ddd', borderRadius:12, background:'#fff', fontSize:14, fontWeight:700, cursor:'pointer', color:'#666' }}>Annuler</button>
                {confirmDesactiver !== 'commande' && (
                  <button onClick={()=>{ c.action(); setConfirmDesactiver(null); }} style={{ flex:1, height:48, border:'none', borderRadius:12, background:'#dc2626', color:'#fff', fontSize:14, fontWeight:800, cursor:'pointer' }}>Désactiver</button>
                )}
              </div>
            </div>
          </>
        );
      })()}
    </>
  );
}

// ── Refus d'une commande : motif obligatoire ─────────────────────────────────
function RefusCommandeModal({ cmd, onClose, onConfirm }) {
  const [motif, setMotif] = useState('');
  const valide = motif.trim().length > 0;

  return (
    <>
      <div onClick={onClose} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)', zIndex:5100 }} />
      <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:20, width:'min(460px, calc(100vw - 32px))', boxShadow:'0 32px 80px rgba(0,0,0,0.3)', zIndex:5101, overflow:'hidden' }}>
        <div style={{ padding:'22px 24px 18px', borderBottom:'1px solid #f0f0f0' }}>
          <h3 style={{ margin:0, fontSize:17, fontWeight:800, color:'#111', display:'flex', alignItems:'center', gap:9 }}>
            <AlertCircle size={18} strokeWidth={2} color="#dc2626" /> Refuser la commande N° {cmd.numero || '—'}
          </h3>
          <p style={{ margin:'5px 0 0', fontSize:13, color:'#888' }}>
            {cmd.client_nom || 'Client'} · {fmtEuro(cmd.total)} — indiquez le motif du refus.
          </p>
        </div>

        <div style={{ padding:'18px 24px 20px', display:'flex', flexDirection:'column', gap:14 }}>
          <div>
            <label style={{ fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, display:'block', marginBottom:8 }}>Motif</label>
            <div style={{ display:'flex', flexDirection:'column', gap:7 }}>
              {MOTIFS_REFUS_CMD.map(m => (
                <button key={m} onClick={()=>setMotif(m)} style={{ width:'100%', minHeight:46, padding:'0 14px', borderRadius:11, border: motif===m ? '2px solid #dc2626' : '1.5px solid #eee', background: motif===m ? '#fef2f2' : '#fff', color: motif===m ? '#b91c1c' : '#333', fontSize:14, fontWeight:700, cursor:'pointer', textAlign:'left' }}>
                  {m}
                </button>
              ))}
            </div>
          </div>

          <div>
            <label style={{ fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, display:'block', marginBottom:6 }}>Précision (facultatif)</label>
            <textarea value={motif} onChange={e=>setMotif(e.target.value.slice(0, 300))} rows={2} placeholder="Motif communiqué au client…" style={{ width:'100%', border:'1.5px solid #ddd', borderRadius:10, padding:'10px 12px', fontSize:14, outline:'none', resize:'vertical', fontFamily:'inherit', boxSizing:'border-box' }} />
          </div>

          <div style={{ display:'flex', gap:10 }}>
            <button onClick={onClose} style={{ ...btnSecondary, flex:1, height:46 }}>Annuler</button>
            <button onClick={()=>valide && onConfirm(motif.trim())} disabled={!valide} style={{ flex:2, height:46, border:'none', borderRadius:10, background: valide ? '#dc2626' : '#f0f0f0', color: valide ? '#fff' : '#bbb', fontSize:14.5, fontWeight:800, cursor: valide ? 'pointer' : 'not-allowed' }}>
              Confirmer le refus
            </button>
          </div>
        </div>
      </div>
    </>
  );
}

// ── Fiche détail d'une commande ──────────────────────────────────────────────
function CommandeDetail({ cmd, onClose, onStatut, onEdit, onSupprimer, auDessus = false, statutsSelonDate = false }) {
  const zVoile = auDessus ? 5400 : 4999;
  const zBoite = auDessus ? 5401 : 5000;
  // Confirmation en deux temps : la suppression est definitive.
  const [confirmeSuppr, setConfirmeSuppr] = useState(false);
  const st = cmdStatut(cmd.statut);

  // Les étapes proposées dépendent du jour : on ne prépare pas une commande
  // de la semaine prochaine, et on ne remet pas en préparation une commande passée.
  const jourCmd = cmd.date_retrait || (cmd.created_at || '').split('T')[0];
  const auj = dateLocale();
  const aVenir = statutsSelonDate && jourCmd > auj;
  const etiquette = { nouvelle:'Nouvelle', en_preparation:'Acceptée', prete:'Prête', recuperee:'Récupérée', annulee:'Annulée' };
  // Une commande à venir n'est pas « en préparation » mais « acceptée », et
  // cette acceptation se lit en vert.
  const passee = statutsSelonDate && jourCmd < auj;
  const habiller = (base) => (aVenir && base.id === 'en_preparation')
    ? { ...base, label:'Acceptée', bg:'#16a34a', fg:'#fff' }
    : base;
  // Une commande passée restée « nouvelle » ou « en préparation » n'a plus
  // d'étape en cours : elle attend d'être clôturée par l'un des trois statuts.
  const aClôturer = passee && ['nouvelle', 'en_preparation'].includes(cmd.statut);
  const statutsProposes = (
    aVenir ? ['nouvelle', 'en_preparation', 'annulee']
    : passee ? ['prete', 'recuperee', 'annulee']
    : CMD_STATUTS.map(x => x.id)
  ).map(id => habiller({ ...cmdStatut(id), label: (aVenir || passee) ? etiquette[id] : cmdStatut(id).label }));
  const stAffiche = aClôturer
    ? { ...st, label:'À clôturer', bg:'#f0a020', fg:'#111' }
    : habiller(st);

  const dateLabel = cmd.created_at ? new Date(cmd.created_at).toLocaleString('fr-FR', { weekday:'long', day:'numeric', month:'long', hour:'2-digit', minute:'2-digit' }) : '';
  return (
    <>
      <div onClick={onClose} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:zVoile }} />
      <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:20, width:'min(680px, calc(100vw - 32px))', maxHeight:'90vh', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.25)', zIndex:zBoite, overflow:'hidden' }}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'20px 24px 16px', borderBottom:'1px solid #f0f0f0', flexShrink:0 }}>
          <div>
            <h2 style={{ margin:0, fontSize:18, fontWeight:800, color:'#111' }}>Commande N° {cmd.numero || '—'}</h2>
            <p style={{ margin:'3px 0 0', fontSize:12.5, color:'#888', textTransform:'capitalize' }}>{dateLabel}</p>
          </div>
          <button onClick={onClose} style={{ width:34, height:34, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:16, color:'#666' }}>✕</button>
        </div>

        <div style={{ padding:'16px 24px 20px', overflowY:'auto', display:'flex', flexDirection:'column', gap:14 }}>
          <span style={{ background:stAffiche.bg, color:stAffiche.fg, borderRadius:20, padding:'5px 14px', fontSize:12, fontWeight:800, alignSelf:'flex-start' }}>{stAffiche.label}</span>

          {/* Client */}
          <div style={{ background:'#f9f9f9', borderRadius:12, padding:'12px 14px' }}>
            <div style={{ fontSize:15, fontWeight:800, color:'#111', marginBottom:6 }}>{cmd.client_nom || 'Client'}</div>
            {cmd.client_tel && <a href={`tel:${cmd.client_tel}`} style={{ display:'flex', alignItems:'center', gap:8, fontSize:14, color:'#111', textDecoration:'none', fontWeight:600, marginBottom:4 }}><Phone size={14} strokeWidth={2} /> {cmd.client_tel}</a>}
            {cmd.client_email && <div style={{ display:'flex', alignItems:'center', gap:8, fontSize:13, color:'#3b82f6' }}><Mail size={13} strokeWidth={2} /> {cmd.client_email}</div>}
            {cmd.heure_retrait && <div style={{ display:'flex', alignItems:'center', gap:8, fontSize:13, color:'#666', marginTop:6 }}><Clock size={13} strokeWidth={2} /> Retrait à {cmd.heure_retrait}</div>}
          </div>

          {cmd.statut === 'recuperee' && recupereeA(cmd) && (
            <div style={{ display:'flex', alignItems:'center', gap:9, background:'#f0fdf4', border:'1.5px solid #bbf7d0', borderRadius:10, padding:'10px 14px', fontSize:13.5, fontWeight:700, color:'#15803d' }}>
              <CircleCheck size={15} strokeWidth={2} /> {fmtRecuperee(recupereeA(cmd))}
            </div>
          )}

          {/* Articles */}
          <div>
            <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', margin:'0 0 8px' }}>
              <p style={{ fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:0 }}>Articles</p>
              {onEdit && (
                <button onClick={()=>onEdit(cmd)} style={{ display:'flex', alignItems:'center', gap:6, height:30, padding:'0 12px', borderRadius:8, border:'1.5px solid #ddd', background:'#fff', fontSize:12.5, fontWeight:700, color:'#111', cursor:'pointer' }}>
                  <Pencil size={13} strokeWidth={2} /> Modifier
                </button>
              )}
            </div>
            {(cmd.items || []).map((it, i) => (
              <div key={i} style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between', gap:12, padding:'10px 0', borderBottom:'1px solid #f5f5f5' }}>
                <div style={{ minWidth:0 }}>
                  <div style={{ fontSize:14, fontWeight:600, color:'#111' }}>
                    <span style={{ color:'#888', fontWeight:800 }}>{it.quantite || 1}×</span> {it.nom}
                  </div>
                  {it.note && <div style={{ fontSize:12, color:'#888', fontStyle:'italic', marginTop:2 }}>{it.note}</div>}
                </div>
                <div style={{ fontSize:14, fontWeight:700, color:'#111', whiteSpace:'nowrap' }}>{fmtEuro((Number(it.prix)||0) * (Number(it.quantite)||1))}</div>
              </div>
            ))}
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', paddingTop:12, marginTop:4, borderTop:'2px solid #111' }}>
              <span style={{ fontSize:14, fontWeight:800, color:'#111' }}>TOTAL</span>
              <span style={{ fontSize:20, fontWeight:900, color:'#111' }}>{fmtEuro(cmd.total)}</span>
            </div>
          </div>

          {cmd.note && (
            <div style={{ background:'#fffbea', border:'1.5px solid #E8C547', borderRadius:10, padding:'10px 14px' }}>
              <p style={{ fontSize:11, fontWeight:700, color:'#92400e', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 4px' }}>Note</p>
              <p style={{ fontSize:13.5, color:'#111', margin:0, lineHeight:1.5 }}>{cmd.note}</p>
            </div>
          )}

          {cmd.motif_refus && (
            <div style={{ background:'#fef2f2', border:'1.5px solid #fca5a5', borderRadius:10, padding:'10px 14px' }}>
              <p style={{ fontSize:11, fontWeight:700, color:'#b91c1c', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 4px' }}>Motif du refus</p>
              <p style={{ fontSize:13.5, color:'#111', margin:0, lineHeight:1.5 }}>{cmd.motif_refus}</p>
            </div>
          )}

          {/* Changement de statut */}
          <div>
            <p style={{ fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:'0 0 8px' }}>Statut</p>
            <div style={{ display:'flex', flexWrap:'wrap', gap:8 }}>
              {statutsProposes.map(s => (
                <button key={s.id} onClick={()=>onStatut(cmd, s.id)} disabled={cmd.statut === s.id} style={{ height:38, padding:'0 14px', borderRadius:9, border: cmd.statut===s.id ? 'none' : '1.5px solid #eee', background: cmd.statut===s.id ? s.bg : '#fff', color: cmd.statut===s.id ? s.fg : '#666', fontSize:13, fontWeight:700, cursor: cmd.statut===s.id ? 'default' : 'pointer' }}>
                  {s.label}
                </button>
              ))}
            </div>
          </div>

          <button onClick={onClose} style={{ width:'100%', height:48, flexShrink:0, border:'1.5px solid #eee', borderRadius:12, background:'#fff', fontSize:14, fontWeight:600, cursor:'pointer', color:'#666' }}>Fermer</button>

          {/* Suppression définitive, en retrait du reste */}
          {onSupprimer && (
            <div style={{ borderTop:'1px solid #f5f5f5', paddingTop:14, marginTop:2, flexShrink:0 }}>
              {!confirmeSuppr ? (
                <button onClick={()=>setConfirmeSuppr(true)} style={{ width:'100%', height:40, flexShrink:0, border:'none', borderRadius:10, background:'none', fontSize:13, fontWeight:700, cursor:'pointer', color:'#dc2626', display:'flex', alignItems:'center', justifyContent:'center', gap:7 }}>
                  <Trash2 size={14} strokeWidth={2} /> Supprimer la commande
                </button>
              ) : (
                <div style={{ background:'#fef2f2', border:'1.5px solid #fca5a5', borderRadius:12, padding:'12px 14px' }}>
                  <p style={{ margin:'0 0 6px', fontSize:13.5, fontWeight:800, color:'#7f1d1d', lineHeight:1.45 }}>
                    Supprimer la commande N° {cmd.numero || '—'} ?
                  </p>
                  <p style={{ margin:'0 0 12px', fontSize:12.5, color:'#7f1d1d', lineHeight:1.55 }}>
                    <strong>Cette action est irréversible.</strong> La commande sera effacée partout :
                    liste des commandes, calendrier, statistiques et exports. Aucune récupération
                    n'est possible ensuite.
                  </p>
                  <div style={{ display:'flex', gap:8 }}>
                    <button onClick={()=>setConfirmeSuppr(false)} style={{ flex:1, height:42, border:'1.5px solid #ddd', borderRadius:10, background:'#fff', fontSize:13.5, fontWeight:700, cursor:'pointer', color:'#666' }}>Annuler</button>
                    <button onClick={()=>onSupprimer(cmd)} style={{ flex:1, height:42, border:'none', borderRadius:10, background:'#dc2626', color:'#fff', fontSize:13.5, fontWeight:800, cursor:'pointer' }}>Supprimer définitivement</button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </>
  );
}

// ── Toute la carte, en grand — pensée pour la prise de commande sur tablette ──
// Cibles tactiles larges : chaque produit est une carte, les boutons − / + font
// 52 px de côté, bien au-delà des 44 px recommandés pour le doigt.
function CatalogueModal({ parCategorie, quantiteDe, onAjouter, onRetirer, nbArticles, total, onEnregistrer, onAbandonner }) {
  const isMobile = useIsMobile();
  const [filtreCat, setFiltreCat] = useState('toutes');
  const [recherche, setRecherche] = useState('');
  // Catégorie en cours de lecture : la pastille correspondante s'allume
  const [catVisible, setCatVisible] = useState(null);
  const zoneRef = useRef(null);
  const pastillesRef = useRef({});
  // Quitter par la croix abandonne les articles ajoutés pendant cette ouverture
  const [confirmeAbandon, setConfirmeAbandon] = useState(false);

  const categoriesAffichees = parCategorie
    .filter(c => filtreCat === 'toutes' || c.id === filtreCat)
    .map(c => ({
      ...c,
      produits: recherche.trim()
        ? c.produits.filter(p => normalizeStr(p.nom || '').includes(normalizeStr(recherche)))
        : c.produits,
    }))
    .filter(c => c.produits.length > 0);

  // Repère la catégorie dont l'en-tête vient de passer en haut de la liste
  // Catégorie en cours de lecture : la dernière dont l'en-tête a passé le haut
  // de la liste. Un IntersectionObserver sert de déclencheur — plus fiable et
  // moins coûteux qu'un gestionnaire de défilement.
  const majCatVisible = () => {
    const zone = zoneRef.current;
    if (!zone) return;
    const hautZone = zone.getBoundingClientRect().top;
    let courante = null;
    zone.querySelectorAll('[data-cat]').forEach(el => {
      if (el.getBoundingClientRect().top - hautZone <= 24) courante = el.dataset.cat;
    });
    setCatVisible(prev => (prev === courante ? prev : courante));
  };

  useEffect(() => {
    const zone = zoneRef.current;
    if (!zone || typeof IntersectionObserver === 'undefined') return;
    const sections = [...zone.querySelectorAll('[data-cat]')];
    if (!sections.length) return;
    majCatVisible();
    const obs = new IntersectionObserver(() => majCatVisible(), { root: zone, threshold: [0, 0.01, 1] });
    sections.forEach(sec => obs.observe(sec));
    return () => obs.disconnect();
  }, [filtreCat, recherche, parCategorie.length]);

  // La pastille allumée reste visible dans la barre de catégories
  useEffect(() => {
    const p = catVisible && pastillesRef.current[catVisible];
    if (p && p.scrollIntoView) p.scrollIntoView({ behavior:'smooth', block:'nearest', inline:'center' });
  }, [catVisible]);

  return (
    <>
      <div onClick={()=>setConfirmeAbandon(true)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)', zIndex:5500 }} />
      <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:22, width:'min(1080px, calc(100vw - 20px))', height:'min(940px, calc(100vh - 20px))', display:'flex', flexDirection:'column', boxShadow:'0 40px 100px rgba(0,0,0,0.35)', zIndex:5501, overflow:'hidden' }}>

        {/* Barre fixe : recherche et catégories restent accessibles au défilement.
            Le titre a disparu au profit de la place à l'écran ; seule la croix reste. */}
        <div style={{ padding:'12px 16px 0', flexShrink:0, borderBottom:'1px solid #f0f0f0' }}>
          <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:10 }}>
            <div style={{ position:'relative', flex:1, minWidth:0 }}>
              <Search size={16} strokeWidth={2} color="#999" style={{ position:'absolute', left:13, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }} />
              <input value={recherche} onChange={e=>setRecherche(e.target.value)} placeholder="Rechercher dans la carte…"
                style={{ width:'100%', height:42, border:'1.5px solid #eee', borderRadius:11, padding:'0 14px 0 38px', fontSize:14, outline:'none', boxSizing:'border-box' }} />
            </div>
            <button onClick={()=>setConfirmeAbandon(true)} style={{ width:40, height:40, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:19, color:'#666', flexShrink:0 }}>✕</button>
          </div>
          <div style={{ display:'flex', gap:7, overflowX:'auto', paddingBottom:9 }}>
            {[{ id:'toutes', nom:'Toutes' }, ...parCategorie].map(c => {
              // Sur « Toutes », la catégorie en cours de lecture s'allume
              const enLecture = filtreCat === 'toutes' && c.id !== 'toutes' && catVisible === c.id;
              const actif = filtreCat === c.id || enLecture;
              return (
                <button key={c.id} ref={el => { pastillesRef.current[c.id] = el; }} onClick={()=>setFiltreCat(c.id)}
                  style={{ height:34, padding:'0 14px', borderRadius:10, fontSize:12, fontWeight:700, border:'none', flexShrink:0, cursor:'pointer', whiteSpace:'nowrap',
                    background: actif ? '#111' : '#f5f5f5',
                    color: actif ? '#fff' : '#666', transition:'background 0.18s, color 0.18s' }}>
                  {c.nom}
                </button>
              );
            })}
          </div>
        </div>

        {/* Grille de produits */}
        <div ref={zoneRef} onScroll={()=>majCatVisible()} style={{ padding:'4px 16px 16px', overflowY:'auto', flex:1, minHeight:0 }}>
          {categoriesAffichees.length === 0 ? (
            <p style={{ margin:0, padding:'60px 20px', textAlign:'center', fontSize:15, color:'#bbb' }}>
              {recherche.trim() ? `Aucun article ne correspond à « ${recherche.trim()} »` : 'Aucun produit disponible dans la carte.'}
            </p>
          ) : categoriesAffichees.map(c => (
            <div key={c.id} data-cat={c.id} style={{ marginBottom:14 }}>
              <p style={{ margin:'8px 0 8px', fontSize:11, fontWeight:800, color:'#888', textTransform:'uppercase', letterSpacing:0.8 }}>{c.nom}</p>
              <div style={{ display:'grid', gridTemplateColumns: isMobile ? '1fr' : 'repeat(auto-fill, minmax(250px, 1fr))', gap:9 }}>
                {c.produits.map(p => {
                  const q = quantiteDe(p);
                  return (
                    <div key={p.id}
                      style={{ display:'flex', alignItems:'center', gap:9, padding:'9px 11px', borderRadius:13,
                        border: q ? '2px solid #E8C547' : '1.5px solid #eee',
                        background: q ? '#fffdf5' : '#fff', minHeight:58, boxSizing:'border-box' }}>
                      {/* Toute la zone texte ajoute l'article d'un simple appui */}
                      <button onClick={()=>onAjouter(p)}
                        style={{ flex:1, minWidth:0, border:'none', background:'none', textAlign:'left', cursor:'pointer', padding:0, alignSelf:'stretch' }}>
                        <div style={{ fontSize:13, fontWeight: q ? 800 : 600, color:'#111', lineHeight:1.25, marginBottom:2 }}>{p.nom}</div>
                        <div style={{ fontSize:12, fontWeight:700, color:'#888' }}>{fmtEuro(p.prix)}</div>
                      </button>
                      <div style={{ display:'flex', alignItems:'center', gap:8, flexShrink:0 }}>
                        <button onClick={()=>onRetirer(p)} disabled={!q}
                          style={{ width:40, height:40, borderRadius:11, border:'1.5px solid #ddd', background:'#fff', fontSize:19, fontWeight:700,
                            cursor: q ? 'pointer' : 'not-allowed', color: q ? '#111' : '#ddd', touchAction:'manipulation' }}>−</button>
                        <span style={{ minWidth:22, textAlign:'center', fontSize:16, fontWeight:900, color: q ? '#111' : '#ccc' }}>{q}</span>
                        <button onClick={()=>onAjouter(p)}
                          style={{ width:40, height:40, borderRadius:11, border:'none', background:'#111', color:'#E8C547', fontSize:19, fontWeight:700,
                            cursor:'pointer', touchAction:'manipulation' }}>+</button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
        </div>

        {/* Pied : récapitulatif et retour à la commande */}
        <div style={{ display:'flex', alignItems:'center', gap:12, padding:'10px 16px calc(12px + env(safe-area-inset-bottom, 0px))', borderTop:'1px solid #f0f0f0', flexShrink:0 }}>
          <div style={{ flex:1, minWidth:0 }}>
            <div style={{ fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5 }}>Au panier</div>
            <div style={{ fontSize:15.5, fontWeight:900, color:'#111' }}>
              {nbArticles} article{nbArticles > 1 ? 's' : ''} · {fmtEuro(total)}
            </div>
          </div>
          <button onClick={onEnregistrer} style={{ height:46, padding:'0 26px', border:'none', borderRadius:12, background:'#E8C547', color:'#111', fontSize:14.5, fontWeight:800, cursor:'pointer', flexShrink:0 }}>
            Enregistrer
          </button>
        </div>
      </div>

      {confirmeAbandon && (
        <>
          <div onClick={()=>setConfirmeAbandon(false)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.45)', zIndex:5600 }} />
          <div style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:18, width:'min(420px, calc(100vw - 40px))', padding:'24px 26px', boxShadow:'0 32px 80px rgba(0,0,0,0.3)', zIndex:5601 }}>
            <h3 style={{ margin:'0 0 8px', fontSize:17, fontWeight:800, color:'#111' }}>Fermer sans enregistrer ?</h3>
            <p style={{ margin:'0 0 20px', fontSize:13.5, color:'#666', lineHeight:1.55 }}>
              Les articles sélectionnés ici ne seront pas ajoutés à la commande.
            </p>
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={()=>setConfirmeAbandon(false)} style={{ flex:1, height:48, border:'1.5px solid #ddd', borderRadius:12, background:'#fff', fontSize:14, fontWeight:700, cursor:'pointer', color:'#666' }}>Continuer</button>
              <button onClick={onAbandonner} style={{ flex:1, height:48, border:'none', borderRadius:12, background:'#dc2626', color:'#fff', fontSize:14, fontWeight:800, cursor:'pointer' }}>Fermer sans enregistrer</button>
            </div>
          </div>
        </>
      )}
    </>
  );
}

// ── Créneaux de retrait proposés en un clic (mêmes services que les résas) ────
const CRENEAUX_MIDI = ['11:30','11:45','12:00','12:15','12:30','12:45','13:00','13:15','13:30','13:45','14:00','14:15'];
const CRENEAUX_SOIR = ['18:00','18:30','19:00','19:15','19:30','19:45','20:00','20:15','20:30','20:45','21:00','21:30'];

// ── Prise de commande au téléphone / modification d'une commande ─────────────
// `cmd` absent → création. `cmd` présent → édition de cette commande.
function NouvelleCommandeModal({ cmd, onClose, onSaved, showToast, delaiDefaut, auDessus = false }) {
  const zVoile = auDessus ? 5400 : 4999;
  const zBoite = auDessus ? 5401 : 5000;
  const edition = !!cmd;
  const [nom, setNom] = useState(cmd?.client_nom || '');
  const [tel, setTel] = useState(cmd?.client_tel || '');
  const [heureRetrait, setHeureRetrait] = useState(cmd?.heure_retrait || '');
  const [dateRetrait, setDateRetrait] = useState(cmd?.date_retrait || dateLocale());
  const [note, setNote] = useState(cmd?.note || '');
  const [items, setItems] = useState(() => (cmd?.items || []).map(it => ({
    nom: it.nom || '', prix: Number(it.prix) || 0, quantite: Number(it.quantite) || 1, note: it.note || '',
  })));
  const [produits, setProduits] = useState([]);
  const [categories, setCategories] = useState([]);
  // Reconnaissance du client à partir du numéro, comme sur une réservation
  const [clientTrouve, setClientTrouve] = useState(null);
  const [rechercheClient, setRechercheClient] = useState(false);
  // Fiche à créer quand le numéro n'est pas connu, comme en réservation
  const [genre, setGenre] = useState('');
  const [prenom, setPrenom] = useState('');
  const [nomFamille, setNomFamille] = useState('');
  const [entreprise, setEntreprise] = useState('');
  const [email, setEmail] = useState('');
  const [recherche, setRecherche] = useState('');
  const [saving, setSaving] = useState(false);
  const lockRef = useRef(false);

  // Calendrier de retrait — même mécanique que « Nouvelle réservation »
  const [showCalPicker, setShowCalPicker] = useState(false);
  const [calPickerDate, setCalPickerDate] = useState(new Date((cmd?.date_retrait || dateLocale()) + 'T12:00:00'));
  const [dateFlash, setDateFlash] = useState(null);
  const [calFermeture, setCalFermeture] = useState(false);

  // Service utilisé pour proposer les créneaux de retrait
  const [svcRetrait, setSvcRetrait] = useState(() => {
    if (cmd?.heure_retrait) return parseInt(String(cmd.heure_retrait).slice(0, 2), 10) < 15 ? 'midi' : 'soir';
    return serviceActuel();
  });

  const [tousLesHoraires, setTousLesHoraires] = useState(false);

  // Catalogue complet, ouvert par le bouton « + ». On mémorise le panier à
  // l'ouverture pour pouvoir revenir en arrière si l'utilisateur abandonne.
  const [showCatalogue, setShowCatalogue] = useState(false);
  const panierAvantCatalogue = useRef([]);

  // Garde-fou de fermeture : on ne perd pas une saisie en cours
  const [toucheAuFormulaire, setToucheAuFormulaire] = useState(false);
  const [confirmeFermeture, setConfirmeFermeture] = useState(false);
  const marquerModifie = () => setToucheAuFormulaire(true);

  useEffect(() => {
    safeQuery(
      () => supabase.from('menu_produits').select('id,nom,prix,disponible,categorie_id,ordre').eq('disponible', true).order('ordre').limit(2000),
      { fallback: [], context: 'commande:produits' }
    ).then(({ data }) => setProduits(data || []));
    safeQuery(
      () => supabase.from('menu_categories').select('id,nom,ordre').order('ordre').limit(200),
      { fallback: [], context: 'commande:categories' }
    ).then(({ data }) => setCategories(data || []));
  }, []);

  async function saisirTel(val) {
    marquerModifie();
    setTel(val);
    setClientTrouve(null);
    setGenre(''); setPrenom(''); setNomFamille(''); setEntreprise(''); setEmail('');
    const digits = val.replace(/\D/g, '');
    if (digits.length < 10) return;
    setRechercheClient(true);
    const telNorm = val.replace(/[\s.\-()]/g, '').replace(/^0/, '+33');
    const { data } = await safeQuery(
      () => supabase.from('clients').select('id,prenom,nom,mail,entreprise,tel_normalise')
        .or(`tel_normalise.eq.${telNorm},tel.eq.${val.trim()}`).maybeSingle(),
      { fallback: null, context: 'commande:rechercheClient' }
    );
    setRechercheClient(false);
    if (!data) return;
    setClientTrouve(data);
    const nomComplet = data.entreprise || [data.prenom, data.nom].filter(Boolean).join(' ');
    if (nomComplet) setNom(nomComplet);
  }

  // Le numéro est complet mais inconnu du fichier : on demande la fiche
  const telComplet = tel.replace(/\D/g, '').length >= 10;
  const nouveauClient = !edition && telComplet && !clientTrouve && !rechercheClient;

  // Nom retenu pour la commande
  const nomPourCommande = nouveauClient
    ? (genre === 'Entreprise' ? entreprise.trim() : `${prenom.trim()} ${nomFamille.trim()}`.trim())
    : nom.trim();


  function ajouter(p) {
    marquerModifie();
    setItems(prev => {
      const i = prev.findIndex(x => x.nom === p.nom && !x.note);
      if (i >= 0) { const n = [...prev]; n[i] = { ...n[i], quantite: (n[i].quantite || 1) + 1 }; return n; }
      return [...prev, { nom: p.nom, prix: Number(p.prix) || 0, quantite: 1, note: '' }];
    });
    setRecherche('');
  }
  function retirerUn(p) {
    marquerModifie();
    setItems(prev => {
      const i = prev.findIndex(x => x.nom === p.nom && !x.note);
      if (i < 0) return prev;
      const n = [...prev];
      if ((n[i].quantite || 1) <= 1) return n.filter((_, idx) => idx !== i);
      n[i] = { ...n[i], quantite: n[i].quantite - 1 };
      return n;
    });
  }
  function ajouterLibre() {
    const nomLibre = recherche.trim();
    if (!nomLibre) return;
    marquerModifie();
    setItems(prev => [...prev, { nom: nomLibre, prix: 0, quantite: 1, note: '' }]);
    setRecherche('');
  }
  const majItem = (i, patch) => { marquerModifie(); setItems(prev => prev.map((it, idx) => idx === i ? { ...it, ...patch } : it)); };
  const retirer  = (i) => { marquerModifie(); setItems(prev => prev.filter((_, idx) => idx !== i)); };

  // Quantité déjà au panier pour un produit du catalogue
  const quantiteDe = (p) => items.filter(x => x.nom === p.nom && !x.note).reduce((s, x) => s + (Number(x.quantite) || 0), 0);

  const total = cmdTotal(items);
  const ficheComplete = !nouveauClient || (genre === 'Entreprise'
    ? !!entreprise.trim()
    : (!!genre && !!prenom.trim() && !!nomFamille.trim()));
  // Le nom vient du client : il faut donc un numéro complet, reconnu ou saisi
  const formValide = nomPourCommande && items.length > 0 && ficheComplete;

  const suggestions = recherche.trim().length >= 1
    ? produits.filter(p => normalizeStr(p.nom || '').includes(normalizeStr(recherche))).slice(0, 6)
    : [];

  // Catalogue groupé par catégorie, dans l'ordre de la carte
  const parCategorie = categories
    .map(c => ({ ...c, produits: produits.filter(p => p.categorie_id === c.id) }))
    .filter(c => c.produits.length > 0);
  const sansCategorie = produits.filter(p => !categories.some(c => c.id === p.categorie_id));
  if (sansCategorie.length) parCategorie.push({ id: '_autres', nom: 'Autres', produits: sansCategorie });

  function demanderFermeture() {
    if (toucheAuFormulaire) { setConfirmeFermeture(true); return; }
    onClose();
  }

  async function enregistrer() {
    if (!formValide || lockRef.current) return;
    if (tel.trim() && !/^(\+33|0)[1-9]\d{8}$/.test(tel.replace(/[\s.\-()]/g, ''))) {
      showToast('Numéro de téléphone invalide', 'error'); return;
    }
    if (email.trim() && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email.trim())) {
      showToast('Email invalide', 'error'); return;
    }
    lockRef.current = true;
    setSaving(true);

    // Numéro inconnu : on crée la fiche client, comme depuis une réservation
    let clientId = clientTrouve?.id || null;
    if (nouveauClient) {
      const telNorm = tel.replace(/[\s.\-()]/g, '').replace(/^0/, '+33');
      const { data: cree } = await safeQuery(
        () => supabase.from('clients').insert({
          genre,
          prenom: genre === 'Entreprise' ? null : capitalize(prenom.trim()),
          nom: genre === 'Entreprise' ? null : capitalize(nomFamille.trim()),
          entreprise: genre === 'Entreprise' ? entreprise.trim() : null,
          mail: email.trim().toLowerCase() || null,
          tel: tel.trim(),
          tel_normalise: telNorm,
        }).select('id').single(),
        { fallback: null, context: 'commande:creerClient' }
      );
      if (cree) clientId = cree.id;
    }

    const itemsPropres = items
      .filter(it => (it.nom || '').trim())
      .map(it => ({
        nom: String(it.nom).trim().slice(0, 120),
        prix: Number(String(it.prix).replace(',', '.')) || 0,
        quantite: Math.max(1, Number(it.quantite) || 1),
        note: (it.note || '').trim().slice(0, 200) || '',
      }));

    if (edition) {
      const { error: errEdit } = await safeQuery(
        () => supabase.from('commandes').update({
          client_nom: nomPourCommande.slice(0, 80),
          client_tel: tel.trim().slice(0, 20) || null,
          items: itemsPropres,
          total: cmdTotal(itemsPropres),
          date_retrait: dateRetrait || cmd.date_retrait,
          heure_retrait: heureRetrait || null,
          note: note.trim().slice(0, 500) || null,
        }).eq('id', cmd.id),
        { context: 'modifierCommande' }
      );
      setSaving(false);
      lockRef.current = false;
      if (errEdit) { showToast('Erreur lors de la modification', 'error'); return; }
      showToast('✅ Commande modifiée');
      onSaved();
      return;
    }

    const { error } = await safeQuery(
      () => supabase.from('commandes').insert({
        client_nom: nomPourCommande.slice(0, 80),
        client_tel: tel.trim().slice(0, 20) || null,
        client_id: clientId,
        client_email: email.trim().toLowerCase() || clientTrouve?.mail || null,
        items: itemsPropres,
        total: cmdTotal(itemsPropres),
        statut: 'en_preparation',
        source: 'telephone',
        date_retrait: dateRetrait || dateLocale(),
        acceptee_at: new Date().toISOString(),
        // Chrono uniquement si la commande est pour aujourd'hui
        delai_minutes: (dateRetrait > dateLocale()) ? null : (delaiDefaut || 30),
        pret_estime_a: (dateRetrait > dateLocale()) ? null : new Date(Date.now() + (delaiDefaut || 30) * 60000).toISOString(),
        heure_retrait: heureRetrait || null,
        note: note.trim().slice(0, 500) || null,
      }),
      { context: 'creerCommandeTelephone' }
    );
    setSaving(false);
    lockRef.current = false;
    if (error) { showToast('Erreur lors de l\'enregistrement', 'error'); return; }
    showToast('✅ Commande enregistrée');
    onSaved();
  }

  // ── Calendrier de retrait, repris à l'identique de « Nouvelle réservation » ──
  const calendarJSX = showCalPicker && (() => {
    const anneeP = calPickerDate.getFullYear();
    const moisP = calPickerDate.getMonth();
    const premierJourSemaine = new Date(anneeP, moisP, 1).getDay() || 7;
    const nbJours = new Date(anneeP, moisP + 1, 0).getDate();
    const casesP = Array(premierJourSemaine - 1).fill(null).concat(Array.from({length: nbJours}, (_, i) => i + 1));
    const todayIso = dateLocale();
    return (
      <div className={calFermeture ? 'cal-fermeture' : ''} style={{ marginTop:8, background:'#fff', borderRadius:12, border:'1.5px solid #eee', boxShadow:'0 4px 16px rgba(0,0,0,0.08)', overflow:'hidden' }}>
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
            const estSelectionne = dateRetrait === iso;
            const aujourd2 = new Date(); aujourd2.setHours(0,0,0,0);
            const estPasse = new Date(anneeP, moisP, jour) < aujourd2;
            return (
              <button key={i} disabled={estPasse} className={dateFlash === iso ? 'date-flash' : ''} onPointerDown={()=>{ if (estPasse) return; marquerModifie(); setDateFlash(iso); setDateRetrait(iso); setTimeout(()=>{ setCalFermeture(true); setTimeout(()=>{ setShowCalPicker(false); setCalFermeture(false); setDateFlash(null); }, 300); }, 200); }} style={{
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

  const creneaux = svcRetrait === 'midi' ? CRENEAUX_MIDI : CRENEAUX_SOIR;

  // Cinq propositions : les prochains créneaux du jour, sinon le début du service.
  // L'heure déjà choisie reste toujours visible, même hors des cinq.
  const creneauxProposes = (() => {
    const tous = [...CRENEAUX_MIDI, ...CRENEAUX_SOIR];
    let base;
    if (dateRetrait === dateLocale()) {
      const maintenant = new Date();
      const hhmm = `${String(maintenant.getHours()).padStart(2,'0')}:${String(maintenant.getMinutes()).padStart(2,'0')}`;
      base = creneaux.filter(h => h > hhmm);
      if (base.length < 5) base = tous.filter(h => h > hhmm);
      if (base.length < 5) base = creneaux.slice(-5);
    } else {
      base = creneaux;
    }
    const cinq = base.slice(0, 5);
    if (heureRetrait && !cinq.includes(heureRetrait)) return [heureRetrait, ...cinq.slice(0, 4)].sort();
    return cinq;
  })();

  return (
    <>
      <div onClick={demanderFermeture} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:zVoile }} />
      <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:20, width:'min(580px, calc(100vw - 32px))', height:'min(880px, calc(100vh - 32px))', display:'flex', flexDirection:'column', boxShadow:'0 32px 80px rgba(0,0,0,0.25)', zIndex:zBoite, overflow:'hidden' }}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'20px 24px 16px', borderBottom:'1px solid #f0f0f0', flexShrink:0 }}>
          <h2 style={{ margin:0, fontSize:18, fontWeight:800, color:'#111', display:'flex', alignItems:'center', gap:8 }}>
            {edition
              ? <><Pencil size={17} strokeWidth={2} /> Modifier la commande N° {cmd.numero || '—'}</>
              : <><Plus size={18} strokeWidth={2.4} /> Nouvelle commande</>}
          </h2>
          <button onClick={demanderFermeture} style={{ width:34, height:34, borderRadius:'50%', border:'none', background:'#f0f0f0', cursor:'pointer', fontSize:16, color:'#666' }}>✕</button>
        </div>

        {/* minHeight:0 est indispensable : sans lui le contenu flex ne défile pas */}
        <div style={{ padding:'16px 24px', overflowY:'auto', display:'flex', flexDirection:'column', gap:16, flex:1, minHeight:0 }}>
          {/* 1. Téléphone — reconnaît le client comme sur une réservation */}
          <div>
            <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>1. Téléphone du client</p>
            <div style={{ position:'relative' }}>
              <Phone size={18} strokeWidth={2} color="#999" style={{ position:'absolute', left:14, top:'50%', transform:'translateY(-50%)', pointerEvents:'none' }} />
              <input type="tel" inputMode="numeric" value={tel} onChange={e=>saisirTel(e.target.value)} placeholder="06 43 00 49 87"
                style={{ width:'100%', height:52, border:'1.5px solid #eee', borderRadius:12, padding:'0 46px', fontSize:16, outline:'none', boxSizing:'border-box' }} />
              {clientTrouve && <CircleCheck size={20} color="#22c55e" style={{ position:'absolute', right:14, top:'50%', transform:'translateY(-50%)' }} />}
              {rechercheClient && <span style={{ position:'absolute', right:14, top:'50%', transform:'translateY(-50%)', fontSize:12, color:'#888' }}>Recherche…</span>}
            </div>
            {clientTrouve && (
              <div style={{ marginTop:8, background:'#f0fdf4', border:'1.5px solid #22c55e', borderRadius:10, padding:'10px 14px', display:'flex', alignItems:'center', gap:10 }}>
                <User size={18} strokeWidth={2} color="#16a34a" />
                <div style={{ minWidth:0 }}>
                  <span style={{ fontWeight:800, fontSize:14, color:'#111' }}>
                    {clientTrouve.entreprise || `${clientTrouve.prenom || ''} ${clientTrouve.nom || ''}`.trim() || 'Client'}
                  </span>
                  <div style={{ fontSize:12, color:'#666', marginTop:2 }}>Client déjà connu{clientTrouve.mail ? ` · ${clientTrouve.mail}` : ''}</div>
                </div>
              </div>
            )}
          </div>

          {/* Le nom fait partie du bloc client : simple si connu, fiche complète sinon */}
          <div style={{ marginTop:-4 }}>
            {nouveauClient && (
              <div style={{ display:'flex', flexDirection:'column', gap:12 }}>
                <p style={{ margin:0, fontSize:13, color:'#92400e', background:'#fffbea', border:'1.5px solid #fde68a', borderRadius:10, padding:'10px 13px', lineHeight:1.5 }}>
                  Numéro inconnu — la fiche client sera créée avec la commande.
                </p>
                <div>
                  <p style={{ fontSize:13, fontWeight:700, color:'#111', margin:'0 0 8px' }}>Genre <span style={{ color:'#dc2626' }}>*</span></p>
                  <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:8 }}>
                    {[
                      {id:'Homme', label:'M. Monsieur', activeColor:'#1d4ed8', activeBg:'#dbeafe'},
                      {id:'Femme', label:'Mme Madame', activeColor:'#be185d', activeBg:'#fce7f3'},
                      {id:'Entreprise', label:'Entreprise', activeColor:'#15803d', activeBg:'#dcfce7'},
                    ].map(g => (
                      <button key={g.id} onClick={()=>{ marquerModifie(); setGenre(g.id); setPrenom(''); setNomFamille(''); setEntreprise(''); }}
                        style={{ height:46, borderRadius:10, cursor:'pointer', fontSize:13, fontWeight:700, border:'1.5px solid',
                          borderColor: genre===g.id ? g.activeBg : '#eee',
                          background: genre===g.id ? g.activeBg : '#fff',
                          color: genre===g.id ? g.activeColor : '#666', transition:'all 0.15s' }}>{g.label}</button>
                    ))}
                  </div>
                </div>

                {genre === 'Entreprise' && (
                  <input value={entreprise} onChange={e=>{marquerModifie(); setEntreprise(e.target.value);}} placeholder="Nom de l'entreprise *"
                    style={{ width:'100%', height:50, border:'1.5px solid', borderColor: entreprise.trim() ? '#22c55e' : '#eee', borderRadius:10, padding:'0 14px', fontSize:15, outline:'none', boxSizing:'border-box' }} />
                )}

                {genre && genre !== 'Entreprise' && (
                  <div style={{ display:'flex', gap:10 }}>
                    <input value={prenom} onChange={e=>{marquerModifie(); setPrenom(e.target.value);}} placeholder="Prénom *"
                      style={{ flex:1, minWidth:0, height:50, border:'1.5px solid', borderColor: prenom.trim() ? '#22c55e' : '#eee', borderRadius:10, padding:'0 14px', fontSize:15, outline:'none' }} />
                    <input value={nomFamille} onChange={e=>{marquerModifie(); setNomFamille(e.target.value);}} placeholder="Nom *"
                      style={{ flex:1, minWidth:0, height:50, border:'1.5px solid', borderColor: nomFamille.trim() ? '#22c55e' : '#eee', borderRadius:10, padding:'0 14px', fontSize:15, outline:'none' }} />
                  </div>
                )}

                {genre && (
                  <input value={email} onChange={e=>{marquerModifie(); setEmail(e.target.value);}} type="email" placeholder="Email (optionnel)"
                    style={{ width:'100%', height:50, border:'1.5px solid #eee', borderRadius:10, padding:'0 14px', fontSize:15, outline:'none', boxSizing:'border-box' }} />
                )}
              </div>
            )}
          </div>

          {/* 2. Date de retrait */}
          <div>
            <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>2. Date de retrait</p>
            <button onClick={()=>setShowCalPicker(v=>!v)} style={{ width:'100%', height:48, borderRadius:10, border:'1.5px solid #ddd', background:'#fff', cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', padding:'0 14px' }}>
              <span style={{ display:'flex', alignItems:'center', gap:9, fontSize:14.5, fontWeight:700, color:'#111', textTransform:'capitalize' }}>
                <CalendarDays size={16} strokeWidth={2} color="#999" />
                {new Date(dateRetrait + 'T12:00:00').toLocaleDateString('fr-FR', { weekday:'long', day:'numeric', month:'long' })}
              </span>
              <ChevronDown size={16} color="#999" style={{ transform: showCalPicker ? 'rotate(180deg)' : 'none', transition:'transform 0.2s' }} />
            </button>
            {calendarJSX}
          </div>

          {/* 3. Service */}
          <div>
            <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>3. Service</p>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10 }}>
              <button onClick={()=>setSvcRetrait('midi')} style={{ height:46, borderRadius:12, cursor:'pointer', fontSize:14.5, fontWeight:700, border:`1.5px solid ${svcRetrait==='midi'?'#E8C547':'#eee'}`, background:svcRetrait==='midi'?'#fffbea':'#fff', color:'#111', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                <Sun size={17} strokeWidth={2} color={svcRetrait==='midi'?'#E8C547':'#999'} /> Midi
              </button>
              <button onClick={()=>setSvcRetrait('soir')} style={{ height:46, borderRadius:12, cursor:'pointer', fontSize:14.5, fontWeight:700, border:svcRetrait==='soir'?'none':'1.5px solid #eee', background:svcRetrait==='soir'?'#111':'#fff', color:svcRetrait==='soir'?'#E8C547':'#111', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>
                <Moon size={17} strokeWidth={2} color={svcRetrait==='soir'?'#E8C547':'#999'} /> Soir
              </button>
            </div>
          </div>

          {/* 4. Heure */}
          <div>
            <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>4. Heure</p>
            {/* Les horaires supplémentaires s'ajoutent à la même grille */}
            <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:8 }}>
              {(tousLesHoraires ? creneaux : creneauxProposes).map(h => (
                <button key={h} onClick={()=>{ marquerModifie(); setHeureRetrait(heureRetrait===h ? '' : h); }} style={{ height:46, borderRadius:10, cursor:'pointer', fontSize:14.5, fontWeight:700, border:`1.5px solid ${heureRetrait===h?'#111':'#eee'}`, background:heureRetrait===h?'#111':'#fff', color:heureRetrait===h?'#E8C547':'#111' }}>{h}</button>
              ))}
              <button onClick={()=>setTousLesHoraires(v=>!v)} style={{ height:46, borderRadius:10, cursor:'pointer', fontSize:13.5, fontWeight:700, border:'1.5px dashed #ccc', background:'#fafafa', color:'#666', display:'flex', alignItems:'center', justifyContent:'center', gap:5 }}>
                {tousLesHoraires ? 'Moins' : "Plus d'horaires"} <ChevronDown size={15} strokeWidth={2.2} style={{ transform: tousLesHoraires ? 'rotate(180deg)' : 'none' }} />
              </button>
            </div>
          </div>

          {/* Ajout d'articles */}
          <div>
            <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>5. Articles <span style={{ color:'#dc2626' }}>*</span></p>
            <div style={{ display:'flex', gap:8, alignItems:'flex-start' }}>
              <div style={{ position:'relative', flex:1 }}>
                <input value={recherche} onChange={e=>setRecherche(e.target.value)} onKeyDown={e=>{ if(e.key==='Enter'){ if(suggestions.length) ajouter(suggestions[0]); else ajouterLibre(); } }} placeholder="Rechercher un produit ou saisir un article hors carte…" style={inp(false)} />
                {suggestions.length > 0 && (
                  <div style={{ position:'absolute', top:'calc(100% + 4px)', left:0, right:0, background:'#fff', borderRadius:10, boxShadow:'0 8px 32px rgba(0,0,0,0.15)', padding:6, zIndex:20, maxHeight:230, overflowY:'auto' }}>
                    {suggestions.map(p => (
                      <button key={p.id} onClick={()=>ajouter(p)} style={{ width:'100%', padding:'10px 12px', border:'none', background:'none', textAlign:'left', cursor:'pointer', fontSize:13.5, borderRadius:6, display:'flex', justifyContent:'space-between', gap:10, color:'#111' }} onMouseEnter={e=>e.currentTarget.style.background='#f5f5f5'} onMouseLeave={e=>e.currentTarget.style.background='none'}>
                        <span>{p.nom}</span><span style={{ fontWeight:700, whiteSpace:'nowrap' }}>{fmtEuro(p.prix)}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
              {/* Ouvre toute la carte dans une modale pensée pour la tablette */}
              <button onClick={()=>{ panierAvantCatalogue.current = items; setShowCatalogue(true); }} title="Voir toute la carte"
                style={{ width:52, height:52, flexShrink:0, borderRadius:12, cursor:'pointer', fontSize:28, fontWeight:400, lineHeight:1,
                  border:'none', background:'#111', color:'#E8C547',
                  display:'flex', alignItems:'center', justifyContent:'center' }}>
                +
              </button>
            </div>
            {recherche.trim() && (
              <button onClick={ajouterLibre} style={{ ...btnSecondary, marginTop:8, height:34 }}>+ Ajouter « {recherche.trim()} » (hors carte)</button>
            )}

          </div>

          {/* Panier */}
          {items.length > 0 && (
            <div style={{ background:'#f9f9f9', borderRadius:12, padding:'8px 12px' }}>
              {items.map((it, i) => (
                <div key={i} style={{ padding:'10px 0', borderBottom: i < items.length-1 ? '1px solid #eee' : 'none' }}>
                  <div style={{ display:'flex', alignItems:'center', gap:10 }}>
                    <div style={{ display:'flex', alignItems:'center', gap:6, flexShrink:0 }}>
                      <button onClick={()=> it.quantite > 1 ? majItem(i,{quantite:it.quantite-1}) : retirer(i)} style={{ width:28, height:28, borderRadius:7, border:'1.5px solid #ddd', background:'#fff', fontSize:15, fontWeight:700, cursor:'pointer', color:'#111' }}>−</button>
                      <span style={{ minWidth:18, textAlign:'center', fontSize:14, fontWeight:800 }}>{it.quantite}</span>
                      <button onClick={()=>majItem(i,{quantite:it.quantite+1})} style={{ width:28, height:28, borderRadius:7, border:'1.5px solid #ddd', background:'#fff', fontSize:15, fontWeight:700, cursor:'pointer', color:'#111' }}>+</button>
                    </div>
                    <input value={it.nom} onChange={e=>majItem(i,{nom:e.target.value})} placeholder="Nom de l'article" style={{ flex:1, minWidth:0, height:32, border:'1.5px solid transparent', borderRadius:7, padding:'0 8px', fontSize:14, fontWeight:600, color:'#111', background:'transparent', outline:'none' }} onFocus={e=>{ e.target.style.borderColor='#ddd'; e.target.style.background='#fff'; }} onBlur={e=>{ e.target.style.borderColor='transparent'; e.target.style.background='transparent'; }} />
                    <div style={{ position:'relative', flexShrink:0 }}>
                      <input value={it.prix} onChange={e=>majItem(i,{prix:e.target.value.replace(',','.')})} inputMode="decimal" style={{ width:76, height:32, border:'1.5px solid #ddd', borderRadius:7, padding:'0 20px 0 8px', fontSize:13, textAlign:'right', outline:'none', boxSizing:'border-box' }} />
                      <span style={{ position:'absolute', right:8, top:'50%', transform:'translateY(-50%)', fontSize:12, color:'#999', pointerEvents:'none' }}>€</span>
                    </div>
                    <button onClick={()=>retirer(i)} style={{ background:'none', border:'none', cursor:'pointer', color:'#e57373', display:'flex', padding:2 }}><Trash2 size={15}/></button>
                  </div>
                  <input value={it.note || ''} onChange={e=>majItem(i,{note:e.target.value})} placeholder="Précision (sans oignons, bien cuite…)" style={{ width:'100%', height:32, border:'none', background:'transparent', fontSize:12.5, color:'#666', outline:'none', marginTop:2, fontStyle:'italic' }} />
                </div>
              ))}
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', paddingTop:10, marginTop:4, borderTop:'2px solid #111' }}>
                <span style={{ fontSize:13, fontWeight:800, color:'#111' }}>TOTAL</span>
                <span style={{ fontSize:19, fontWeight:900, color:'#111' }}>{fmtEuro(total)}</span>
              </div>
            </div>
          )}

          <div>
            <p style={{ fontSize:14, fontWeight:800, color:'#111', margin:'0 0 10px' }}>6. Notes <span style={{ fontSize:12, fontWeight:400, color:'#999' }}>(optionnel)</span></p>
            <textarea value={note} onChange={e=>{marquerModifie(); setNote(e.target.value);}} placeholder="Note générale (allergie, paiement…)" rows={3} style={{ width:'100%', border:'1.5px solid #ddd', borderRadius:7, padding:'10px 12px', fontSize:14, outline:'none', resize:'vertical', fontFamily:'inherit', boxSizing:'border-box' }} />
          </div>
        </div>

        <div style={{ padding:'14px 24px calc(18px + env(safe-area-inset-bottom, 0px))', borderTop:'1px solid #f0f0f0', flexShrink:0 }}>
          <button onClick={enregistrer} disabled={!formValide || saving} style={{ width:'100%', height:52, background: formValide && !saving ? '#E8C547' : '#f0f0f0', color: formValide && !saving ? '#111' : '#bbb', border:'none', borderRadius:14, fontSize:16, fontWeight:800, cursor: formValide && !saving ? 'pointer' : 'not-allowed' }}>
            {saving ? 'Enregistrement…' : `✓ ${edition ? 'Enregistrer les modifications' : 'Enregistrer la commande'}${items.length ? ' · ' + fmtEuro(total) : ''}`}
          </button>
        </div>
      </div>

      {/* Toute la carte, en grand, pour la prise de commande sur tablette */}
      {showCatalogue && (
        <CatalogueModal
          parCategorie={parCategorie}
          quantiteDe={quantiteDe}
          onAjouter={ajouter}
          onRetirer={retirerUn}
          nbArticles={items.reduce((n, it) => n + (Number(it.quantite) || 0), 0)}
          total={total}
          onEnregistrer={()=>setShowCatalogue(false)}
          onAbandonner={()=>{ setItems(panierAvantCatalogue.current); setShowCatalogue(false); }}
        />
      )}

      {/* Fermeture avec des modifications non enregistrées */}
      {confirmeFermeture && (
        <>
          <div onClick={()=>setConfirmeFermeture(false)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.4)', zIndex:5700 }} />
          <div style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:18, width:'min(400px, calc(100vw - 40px))', padding:'22px 24px', boxShadow:'0 32px 80px rgba(0,0,0,0.3)', zIndex:5701 }}>
            <h3 style={{ margin:'0 0 8px', fontSize:16.5, fontWeight:800, color:'#111' }}>Quitter sans enregistrer ?</h3>
            <p style={{ margin:'0 0 18px', fontSize:13.5, color:'#666', lineHeight:1.55 }}>
              {edition ? 'Les modifications apportées à cette commande seront perdues.' : 'La commande en cours de saisie sera perdue.'}
            </p>
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={()=>setConfirmeFermeture(false)} style={{ flex:1, height:46, border:'1.5px solid #ddd', borderRadius:12, background:'#fff', fontSize:14, fontWeight:700, cursor:'pointer', color:'#666' }}>Continuer la saisie</button>
              <button onClick={onClose} style={{ flex:1, height:46, border:'none', borderRadius:12, background:'#dc2626', color:'#fff', fontSize:14, fontWeight:800, cursor:'pointer' }}>Quitter</button>
            </div>
          </div>
        </>
      )}
    </>
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
  const emoji = isEmail ? <Mail size={40} style={{display:'block',margin:'0 auto'}} /> : <MessageSquare size={40} style={{display:'block',margin:'0 auto'}} />;
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
            <div style={{ marginBottom:18, filter:'drop-shadow(0 4px 8px rgba(0,0,0,0.12))' }}>{emoji}</div>
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

function MenuToggle({ value, onChange, colorOn = '#E8C547' }) {
  return (
    <div onClick={onChange} style={{ width:44, height:24, borderRadius:12, background: value ? colorOn : '#ddd', cursor:'pointer', position:'relative', transition:'background 0.2s', flexShrink:0 }}>
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

  const sheetTitle = form.type === 'plat' ? <><UtensilsCrossed size={14} style={{display:'inline',verticalAlign:'middle',marginRight:4}} />Plat du jour</> : form.type === 'dessert' ? <><UtensilsCrossed size={14} style={{display:'inline',verticalAlign:'middle',marginRight:4}} />Dessert du jour</> : <><UtensilsCrossed size={14} style={{display:'inline',verticalAlign:'middle',marginRight:4}} />Suggestion du chef</>;
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
  const [confirmForce, setConfirmForce] = useState(null); // { carte, nbCats, nbProds }
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
    const nbProds = produits.filter(p => p.carte === c.slug).length;
    const { data: cats } = await supabase.from('menu_categories').select('id').eq('carte', c.slug);
    const nbCats = cats?.length || 0;
    if (nbProds > 0 || nbCats > 0) {
      setConfirmForce({ carte: c, nbCats, nbProds });
      return;
    }
    await doDeleteCarte(c);
  }

  async function doDeleteCarte(c) {
    await supabase.from('menu_produits').delete().eq('carte', c.slug);
    await supabase.from('menu_categories').delete().eq('carte', c.slug);
    await supabase.from('menu_cartes').delete().eq('id', c.id);
    setCartes(prev => prev.filter(x => x.id !== c.id));
    setConfirmForce(null);
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
    <>
    {confirmForce && (
      <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)', zIndex:4500, display:'flex', alignItems:'center', justifyContent:'center', padding:20 }} onClick={() => setConfirmForce(null)}>
        <div style={{ background:'#fff', borderRadius:16, width:'100%', maxWidth:400, padding:'24px 20px', boxShadow:'0 8px 40px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
          <div style={{ textAlign:'center', marginBottom:12 }}><AlertCircle size={32} color="#dc2626" style={{display:'block',margin:'0 auto'}} /></div>
          <h3 style={{ margin:'0 0 10px', fontSize:16, fontWeight:800, color:'#111', textAlign:'center' }}>
            Supprimer « {confirmForce.carte.nom} » ?
          </h3>
          <p style={{ margin:'0 0 20px', fontSize:14, color:'#555', lineHeight:1.6, textAlign:'center' }}>
            Cette carte contient <strong>{confirmForce.nbCats} catégorie{confirmForce.nbCats>1?'s':''}</strong> et <strong>{confirmForce.nbProds} produit{confirmForce.nbProds>1?'s':''}</strong>.<br/>
            Tout son contenu sera définitivement perdu.
          </p>
          <div style={{ display:'flex', gap:8 }}>
            <button onClick={() => setConfirmForce(null)} style={{ ...btnSecondary, flex:1 }}>Annuler</button>
            <button onClick={() => doDeleteCarte(confirmForce.carte)} style={{ ...btnDanger, flex:1 }}>Supprimer quand même</button>
          </div>
        </div>
      </div>
    )}
    <MenuBottomSheet title="🗂 Gérer les cartes" onClose={onClose} footer={<button onClick={onClose} style={{ ...btnPrimary, width:'100%' }}>Fermer</button>}>
      <div style={{ display:'flex', gap:8, marginBottom:16 }}>
        <input value={newNom} onChange={e=>setNewNom(e.target.value)} placeholder="Nouvelle carte..." style={{ ...inp(false), flex:1, height:42 }} onKeyDown={e=>e.key==='Enter'&&addCarte()} />
        <button onClick={addCarte} disabled={adding} style={{ ...btnPrimary, height:42, whiteSpace:'nowrap' }}>+ Ajouter</button>
      </div>
      <p style={{ fontSize:11, color:'#bbb', marginBottom:12, lineHeight:1.5 }}>Le slug est généré automatiquement. Utilisé dans "carte" des produits. La suppression d'une carte supprime aussi ses catégories et produits.</p>
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
            <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={()=>{setEditingId(c.id);setEditNom(c.nom);}} style={{ background:'none', border:'none', cursor:'pointer', color:'#888', display:'flex', padding:4 }}><Pencil size={13}/></button>
            <MenuToggle value={!!c.visible} onChange={()=>toggleVisible(c)} />
            <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={e=>{e.stopPropagation();console.log('[CartesSheet] deleteCarte click', c.nom);deleteCarte(c);}} style={{ background:'none', border:'none', cursor:'pointer', color:'#e57373', display:'flex', padding:4 }}><Trash2 size={14}/></button>
          </div>
        ))}
      </div>
    </MenuBottomSheet>
    </>
  );
}



function CatsSheet({ categories: initCats, onClose, showToast, carte, produits }) {
  const [cats, setCats] = useState([...initCats]);
  const [newNom, setNewNom] = useState('');
  const [newCarte, setNewCarte] = useState(carte);
  const [editingId, setEditingId] = useState(null);
  const [editNom, setEditNom] = useState('');
  const [adding, setAdding] = useState(false);
  const [confirmDel, setConfirmDel] = useState(null); // { cat, nbProds }
  const [bannerCat, setBannerCat] = useState(null);   // catégorie dont on édite le bandeau
  const [bannerPhrase, setBannerPhrase] = useState('');
  const [bannerUploading, setBannerUploading] = useState(false);
  const [bannerSaving, setBannerSaving] = useState(false);
  const dragIdx = useRef(null);
  const dragOverIdx = useRef(null);

  function openBanner(cat) {
    setBannerCat(cat);
    setBannerPhrase(cat.banner_phrase || '');
  }

  async function uploadBanner(file) {
    if (!file || !bannerCat) return;
    if (file.size > 5 * 1024 * 1024) { showToast('❌ Image trop lourde (max 5 Mo)'); return; }
    setBannerUploading(true);
    try {
      const ext = file.name.split('.').pop().toLowerCase();
      const path = `banner-cat-${bannerCat.id}-${Date.now()}.${ext}`;
      const { error } = await supabase.storage.from('soirees-flyers').upload(path, file, { upsert: true, contentType: file.type });
      if (error) throw error;
      const { data } = supabase.storage.from('soirees-flyers').getPublicUrl(path);
      const url = data.publicUrl;
      await supabase.from('menu_categories').update({ banner_url: url }).eq('id', bannerCat.id);
      setCats(prev => prev.map(c => c.id === bannerCat.id ? { ...c, banner_url: url } : c));
      setBannerCat(prev => ({ ...prev, banner_url: url }));
      showToast('✅ Image du bandeau enregistrée');
    } catch (e) {
      logError(e.message, 'catsSheet:uploadBanner');
      showToast('❌ Échec de l\'upload', 'error');
    }
    setBannerUploading(false);
  }

  async function saveBannerPhrase() {
    if (!bannerCat || bannerSaving) return;
    setBannerSaving(true);
    await supabase.from('menu_categories').update({ banner_phrase: bannerPhrase.trim() || null }).eq('id', bannerCat.id);
    setCats(prev => prev.map(c => c.id === bannerCat.id ? { ...c, banner_phrase: bannerPhrase.trim() || null } : c));
    setBannerSaving(false);
    setBannerCat(null);
    showToast('✅ Bandeau enregistré');
  }

  async function removeBanner() {
    if (!bannerCat) return;
    await supabase.from('menu_categories').update({ banner_url: null, banner_phrase: null }).eq('id', bannerCat.id);
    setCats(prev => prev.map(c => c.id === bannerCat.id ? { ...c, banner_url: null, banner_phrase: null } : c));
    setBannerCat(null);
    showToast('Bandeau retiré');
  }

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

  async function deleteCat(cat) {
    const nbProds = (produits || []).filter(p => p.categorie_id === cat.id).length;
    if (nbProds > 0) { setConfirmDel({ cat, nbProds }); return; }
    await doDeleteCat(cat);
  }

  async function doDeleteCat(cat) {
    await supabase.from('menu_produits').delete().eq('categorie_id', cat.id);
    await supabase.from('menu_categories').delete().eq('id', cat.id);
    setCats(prev => prev.filter(c => c.id !== cat.id));
    setConfirmDel(null);
    showToast('Catégorie supprimée');
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
    <>
    {confirmDel && (
      <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)', zIndex:4500, display:'flex', alignItems:'center', justifyContent:'center', padding:20 }} onClick={() => setConfirmDel(null)}>
        <div style={{ background:'#fff', borderRadius:16, width:'100%', maxWidth:400, padding:'24px 20px', boxShadow:'0 8px 40px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
          <div style={{ textAlign:'center', marginBottom:12 }}><AlertCircle size={32} color="#dc2626" style={{display:'block',margin:'0 auto'}} /></div>
          <h3 style={{ margin:'0 0 10px', fontSize:16, fontWeight:800, color:'#111', textAlign:'center' }}>
            Supprimer « {confirmDel.cat.nom} » ?
          </h3>
          <p style={{ margin:'0 0 20px', fontSize:14, color:'#555', lineHeight:1.6, textAlign:'center' }}>
            Cette catégorie contient <strong>{confirmDel.nbProds} produit{confirmDel.nbProds>1?'s':''}</strong>.<br/>
            Tout son contenu sera définitivement perdu.
          </p>
          <div style={{ display:'flex', gap:8 }}>
            <button onClick={() => setConfirmDel(null)} style={{ ...btnSecondary, flex:1 }}>Annuler</button>
            <button onClick={() => doDeleteCat(confirmDel.cat)} style={{ ...btnDanger, flex:1 }}>Supprimer quand même</button>
          </div>
        </div>
      </div>
    )}
    {bannerCat && (
      <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)', zIndex:4500, display:'flex', alignItems:'center', justifyContent:'center', padding:20 }} onClick={() => !bannerUploading && setBannerCat(null)}>
        <div style={{ background:'#fff', borderRadius:16, width:'100%', maxWidth:440, padding:'22px 20px', boxShadow:'0 8px 40px rgba(0,0,0,0.2)' }} onClick={e => e.stopPropagation()}>
          <h3 style={{ margin:'0 0 4px', fontSize:16, fontWeight:800, color:'#111', display:'flex', alignItems:'center', gap:8 }}>
            <ImageIcon size={17}/> Bandeau photo — {bannerCat.nom}
          </h3>
          <p style={{ margin:'0 0 14px', fontSize:12.5, color:'#888', lineHeight:1.5 }}>
            Cette image s'affiche sur la carte client, juste au-dessus de la catégorie, avec une phrase d'ambiance optionnelle.
          </p>
          {bannerCat.banner_url ? (
            <div style={{ position:'relative', height:110, borderRadius:10, overflow:'hidden', marginBottom:12, background:`#111 url('${bannerCat.banner_url}') center/cover` }}>
              {bannerPhrase && <div style={{ position:'absolute', inset:0, display:'flex', alignItems:'center', justifyContent:'center', background:'rgba(20,16,11,0.45)', color:'#fff', fontStyle:'italic', fontFamily:'Georgia,serif', fontSize:16, textShadow:'0 2px 6px rgba(0,0,0,0.9)', padding:'0 16px', textAlign:'center' }}>{bannerPhrase}</div>}
            </div>
          ) : (
            <div style={{ height:70, borderRadius:10, marginBottom:12, background:'#f5f5f5', border:'1.5px dashed #ddd', display:'flex', alignItems:'center', justifyContent:'center', color:'#aaa', fontSize:13 }}>Aucune image pour l'instant</div>
          )}
          <label style={{ display:'flex', alignItems:'center', justifyContent:'center', gap:8, height:42, border:'1.5px solid #ddd', borderRadius:10, cursor: bannerUploading ? 'wait' : 'pointer', fontSize:13, fontWeight:700, color:'#333', marginBottom:12, background:'#fafafa' }}>
            <ImageIcon size={15}/> {bannerUploading ? 'Envoi en cours…' : (bannerCat.banner_url ? 'Remplacer l\'image' : 'Choisir une image')}
            <input type="file" accept="image/*" disabled={bannerUploading} style={{ display:'none' }} onChange={e => { const f = e.target.files && e.target.files[0]; if (f) uploadBanner(f); e.target.value=''; }} />
          </label>
          <label style={{ fontSize:12, fontWeight:700, color:'#666', display:'block', marginBottom:6 }}>Phrase d'ambiance (optionnelle)</label>
          <input value={bannerPhrase} onChange={e => setBannerPhrase(e.target.value.slice(0,80))} placeholder="Ex. Une douceur pour finir en beauté" style={{ width:'100%', height:42, border:'1.5px solid #ddd', borderRadius:10, padding:'0 12px', fontSize:14, boxSizing:'border-box', marginBottom:16, outline:'none' }} />
          <div style={{ display:'flex', gap:8 }}>
            {bannerCat.banner_url && <button onClick={removeBanner} style={{ ...btnDanger, flex:1 }}>Retirer</button>}
            <button onClick={() => setBannerCat(null)} style={{ ...btnSecondary, flex:1 }}>Annuler</button>
            <button onClick={saveBannerPhrase} disabled={bannerSaving} style={{ ...btnPrimary, flex:1 }}>{bannerSaving ? '…' : 'Enregistrer'}</button>
          </div>
        </div>
      </div>
    )}
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
            <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={()=>{setEditingId(cat.id);setEditNom(cat.nom);}} style={{ background:'none', border:'none', cursor:'pointer', color:'#888', display:'flex', padding:4, flexShrink:0 }}><Pencil size={13}/></button>
            <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={e=>{e.stopPropagation();openBanner(cat);}} title={cat.banner_url ? 'Bandeau photo défini — modifier' : 'Ajouter un bandeau photo au-dessus de cette catégorie'} style={{ background:'none', border:'none', cursor:'pointer', color: cat.banner_url ? '#b8860b' : '#bbb', display:'flex', padding:4, flexShrink:0 }}><ImageIcon size={14}/></button>
            <MenuToggle value={cat.visible!==false} onChange={()=>toggleVisible(cat)} />
            <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={e=>{e.stopPropagation();deleteCat(cat);}} style={{ background:'none', border:'none', cursor:'pointer', color:'#e57373', display:'flex', padding:4, flexShrink:0 }}><Trash2 size={14}/></button>
          </div>
        ))}
      </div>
    </MenuBottomSheet>
    </>
  );
}

// ── Origine des viandes ──────────────────────────────────────────────────────

function OriginesSheet({ onClose, showToast }) {
  const [titre, setTitre] = useState('Origine des viandes');
  const [editTitre, setEditTitre] = useState(false);
  const [titreDraft, setTitreDraft] = useState('');
  const [origines, setOrigines] = useState([]);
  const [loading, setLoading] = useState(true);
  const [editingId, setEditingId] = useState(null);
  const [editProduit, setEditProduit] = useState('');
  const [editOrigine, setEditOrigine] = useState('');
  const [newProduit, setNewProduit] = useState('');
  const [newOrigine, setNewOrigine] = useState('');
  const [adding, setAdding] = useState(false);
  const dragIdx = useRef(null);
  const dragOverIdx = useRef(null);

  useEffect(() => { load(); }, []);

  async function load() {
    const { data } = await supabase.from('menu_origines').select('*').order('ordre');
    if (data) setOrigines(data);
    setLoading(false);
  }

  async function saveTitre() {
    setTitre(titreDraft.trim() || titre);
    setEditTitre(false);
  }

  async function startEdit(o) {
    setEditingId(o.id);
    setEditProduit(o.produit);
    setEditOrigine(o.origine);
  }

  async function saveEdit(o) {
    if (!editProduit.trim()) { setEditingId(null); return; }
    await supabase.from('menu_origines').update({ produit: editProduit.trim(), origine: editOrigine.trim() }).eq('id', o.id);
    setOrigines(prev => prev.map(x => x.id === o.id ? { ...x, produit: editProduit.trim(), origine: editOrigine.trim() } : x));
    setEditingId(null);
    showToast('Modifié ✓');
  }

  async function deleteOrigine(o) {
    await supabase.from('menu_origines').delete().eq('id', o.id);
    setOrigines(prev => prev.filter(x => x.id !== o.id));
    showToast('Supprimé');
  }

  async function addOrigine() {
    if (!newProduit.trim()) return;
    setAdding(true);
    const ordre = origines.length > 0 ? Math.max(...origines.map(o => o.ordre || 0)) + 1 : 1;
    const { data } = await supabase.from('menu_origines').insert({ produit: newProduit.trim(), origine: newOrigine.trim(), ordre }).select().single();
    if (data) { setOrigines(prev => [...prev, data]); setNewProduit(''); setNewOrigine(''); showToast('Ajouté ✓'); }
    setAdding(false);
  }

  async function drop() {
    if (dragIdx.current === null || dragOverIdx.current === null || dragIdx.current === dragOverIdx.current) { dragIdx.current = null; dragOverIdx.current = null; return; }
    const next = [...origines];
    const [moved] = next.splice(dragIdx.current, 1);
    next.splice(dragOverIdx.current, 0, moved);
    const updated = next.map((o, i) => ({ ...o, ordre: i + 1 }));
    setOrigines(updated);
    await Promise.all(updated.map(o => supabase.from('menu_origines').update({ ordre: o.ordre }).eq('id', o.id)));
    dragIdx.current = null; dragOverIdx.current = null;
  }

  return (
    <MenuBottomSheet
      title={
        editTitre
          ? <input value={titreDraft} onChange={e => setTitreDraft(e.target.value)}
              onBlur={saveTitre} onKeyDown={e => { if (e.key==='Enter') saveTitre(); if (e.key==='Escape') setEditTitre(false); }}
              autoFocus style={{ fontSize:16, fontWeight:800, color:'#111', border:'none', borderBottom:'2px solid #E8C547', outline:'none', background:'transparent', width:'100%' }} />
          : <span style={{ cursor:'pointer', display:'flex', alignItems:'center', gap:6 }} onClick={() => { setTitreDraft(titre); setEditTitre(true); }}>
              {titre} <Pencil size={12} color="#ccc" />
            </span>
      }
      onClose={onClose}
    >
      {loading ? (
        <div style={{ textAlign:'center', padding:'40px 0', color:'#ccc', fontSize:14 }}>Chargement…</div>
      ) : (
        <div style={{ display:'flex', flexDirection:'column', gap:12 }}>
          <p style={{ margin:0, fontSize:12, color:'#bbb' }}>Cliquez sur le titre pour le renommer. Glissez pour réordonner.</p>

          {/* Liste */}
          <div style={{ border:'1px solid #eee', borderRadius:12, overflow:'hidden' }}>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr auto', padding:'8px 14px', background:'#f9f9f9', borderBottom:'1px solid #eee' }}>
              <span style={{ fontSize:10, fontWeight:700, color:'#aaa', textTransform:'uppercase', letterSpacing:1 }}>Produit</span>
              <span style={{ fontSize:10, fontWeight:700, color:'#aaa', textTransform:'uppercase', letterSpacing:1 }}>Origine</span>
              <span style={{ width:56 }} />
            </div>
            {origines.map((o, i) => (
              <div key={o.id}
                draggable onDragStart={() => { dragIdx.current = i; }} onDragEnter={() => { dragOverIdx.current = i; }} onDragEnd={drop} onDragOver={e => e.preventDefault()}
                style={{ display:'grid', gridTemplateColumns:'1fr 1fr auto', alignItems:'center', padding:'13px 14px', borderBottom:'1px solid #f0f0f0', background: i % 2 === 0 ? '#fff' : '#fafafa', cursor:'grab', userSelect:'none', transition:'background 0.1s' }}>
                {editingId === o.id ? (
                  <>
                    <input value={editProduit} onChange={e => setEditProduit(e.target.value)}
                      onKeyDown={e => { if (e.key==='Enter') saveEdit(o); if (e.key==='Escape') setEditingId(null); }}
                      style={{ height:32, border:'1.5px solid #E8C547', borderRadius:7, padding:'0 8px', fontSize:13, outline:'none', marginRight:6, background:'#fff' }} autoFocus />
                    <input value={editOrigine} onChange={e => setEditOrigine(e.target.value)}
                      onKeyDown={e => { if (e.key==='Enter') saveEdit(o); if (e.key==='Escape') setEditingId(null); }}
                      style={{ height:32, border:'1.5px solid #E8C547', borderRadius:7, padding:'0 8px', fontSize:13, outline:'none', marginRight:6, background:'#fff' }} />
                    <div style={{ display:'flex', gap:4 }}>
                      <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={() => saveEdit(o)} style={{ ...btnPrimary, height:30, fontSize:12, padding:'0 10px' }}>✓</button>
                      <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={() => setEditingId(null)} style={{ ...btnSecondary, height:30, fontSize:12, padding:'0 10px' }}>✕</button>
                    </div>
                  </>
                ) : (
                  <>
                    <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                      <span style={{ color:'#d0d0d0', fontSize:13, flexShrink:0 }}>⠿</span>
                      <span style={{ fontSize:13, fontWeight:600, color:'#111' }}>{o.produit}</span>
                    </div>
                    <span style={{ fontSize:13, color:'#555' }}>{o.origine}</span>
                    <div style={{ display:'flex', gap:2, justifyContent:'flex-end' }}>
                      <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={() => startEdit(o)} style={{ background:'none', border:'none', cursor:'pointer', color:'#bbb', display:'flex', padding:'4px 6px', borderRadius:6 }}><Pencil size={12} /></button>
                      <button draggable={false} onPointerDown={e=>e.stopPropagation()} onClick={() => deleteOrigine(o)} style={{ background:'none', border:'none', cursor:'pointer', color:'#e57373', display:'flex', padding:'4px 6px', borderRadius:6 }}><Trash2 size={12} /></button>
                    </div>
                  </>
                )}
              </div>
            ))}
            {origines.length === 0 && (
              <div style={{ padding:'24px', textAlign:'center', color:'#ccc', fontSize:13 }}>Aucune entrée. Ajoutez-en ci-dessous.</div>
            )}
          </div>

          {/* Ajouter */}
          <div style={{ display:'flex', gap:8 }}>
            <input value={newProduit} onChange={e => setNewProduit(e.target.value)} placeholder="Produit (ex: Entrecôte)" onKeyDown={e => e.key==='Enter' && addOrigine()}
              style={{ ...inp(false), flex:2 }} />
            <input value={newOrigine} onChange={e => setNewOrigine(e.target.value)} placeholder="Origine (ex: France)" onKeyDown={e => e.key==='Enter' && addOrigine()}
              style={{ ...inp(false), flex:2 }} />
            <button onClick={addOrigine} disabled={adding || !newProduit.trim()} style={{ ...btnPrimary, display:'flex', alignItems:'center', gap:6, whiteSpace:'nowrap' }}>
              <Plus size={14} /> Ajouter
            </button>
          </div>
        </div>
      )}
    </MenuBottomSheet>
  );
}

// ── Roue cadeaux — back-office ───────────────────────────────────────────────

const DEFAULT_EMAIL1_OBJET = 'Votre récompense au Grand Jeu du TED';
const DEFAULT_EMAIL1_CORPS = `Bonjour {prenom},

Félicitations ! Vous remportez : {emoji} {recompense}

Votre récompense vous attend au TED le {date}.

Pour en profiter, il vous suffit de venir accompagné(e) d'au moins 4 personnes, soit une table de 5 personnes minimum, autour d'un repas au TED.

À votre arrivée, présentez simplement cet e-mail à notre équipe : votre récompense vous sera offerte. 🥂

Une belle occasion de réunir vos proches et de célébrer cette victoire comme il se doit !

📞 Réservation par téléphone : 04 72 02 20 20
🔗 Réservation en ligne : https://ted-crm.pages.dev/reserver.html

Nous avons hâte de vous accueillir.
À très bientôt,
L'équipe du TED 🦁`;

function RecompenseSheet({ item, onClose, onSaved, showToast }) {
  const [nom, setNom] = useState(item?.nom || '');
  const [emoji, setEmoji] = useState(item?.emoji || '');
  const [probabilite, setProbabilite] = useState(item?.probabilite ?? 1);
  const [stock, setStock] = useState(item?.stock ?? 10);
  const [stockIllimite, setStockIllimite] = useState(item?.stock_illimite ?? false);
  const [actif, setActif] = useState(item?.actif ?? true);
  const [conditions, setConditions] = useState(item?.conditions || '');
  const [saving, setSaving] = useState(false);

  async function save() {
    if (!nom.trim()) { showToast('Nom requis'); return; }
    setSaving(true);
    const payload = {
      nom: nom.trim(), emoji: emoji.trim(),
      probabilite: parseFloat(probabilite) || 1,
      stock: stockIllimite ? null : parseInt(stock) || 0,
      stock_illimite: stockIllimite, actif,
      conditions: conditions.trim() || null
    };
    if (item?.id) {
      await supabase.from('roue_recompenses').update(payload).eq('id', item.id);
    } else {
      await supabase.from('roue_recompenses').insert({ ...payload, ordre: 99 });
    }
    setSaving(false);
    showToast('Récompense sauvegardée ✓');
    onSaved();
  }

  const iStyle = { width:'100%', padding:'12px 14px', borderRadius:10, border:'1.5px solid #eee', fontSize:14, fontFamily:'inherit', outline:'none', background:'#fafafa', boxSizing:'border-box' };

  return (
    <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.45)', zIndex:999, display:'flex', alignItems:'flex-end', justifyContent:'center' }}>
      <div style={{ background:'#fff', borderRadius:'20px 20px 0 0', padding:28, width:'100%', maxWidth:480, maxHeight:'90vh', overflowY:'auto' }}>
        <h3 style={{ margin:'0 0 20px', fontSize:18, fontWeight:900 }}>{item?.id ? 'Modifier' : 'Ajouter'} une récompense</h3>
        <div style={{ display:'flex', flexDirection:'column', gap:14 }}>
          <div>
            <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Nom</label>
            <input value={nom} onChange={e=>setNom(e.target.value)} placeholder="Ex: Magnum de rosé" style={{ ...iStyle, marginTop:6 }} />
          </div>
          <div>
            <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Emoji</label>
            <input value={emoji} onChange={e=>setEmoji(e.target.value)} placeholder="🍾" style={{ ...iStyle, marginTop:6 }} />
          </div>
          <div>
            <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Probabilité (poids relatif)</label>
            <input type="number" min={0.1} step={0.1} value={probabilite} onChange={e=>setProbabilite(e.target.value)} style={{ ...iStyle, marginTop:6 }} />
          </div>
          <label style={{ display:'flex', alignItems:'center', gap:10, cursor:'pointer' }}>
            <input type="checkbox" checked={stockIllimite} onChange={e=>setStockIllimite(e.target.checked)} style={{ width:18, height:18, accentColor:'#111' }} />
            <span style={{ fontWeight:600, fontSize:14 }}>Stock illimité</span>
          </label>
          {!stockIllimite && (
            <div>
              <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Stock</label>
              <input type="number" min={0} value={stock} onChange={e=>setStock(e.target.value)} style={{ ...iStyle, marginTop:6 }} />
            </div>
          )}
          <div>
            <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Conditions (optionnel)</label>
            <input value={conditions} onChange={e=>setConditions(e.target.value)} placeholder="Valable jusqu'au…" style={{ ...iStyle, marginTop:6 }} />
          </div>
          <label style={{ display:'flex', alignItems:'center', gap:10, cursor:'pointer' }}>
            <input type="checkbox" checked={actif} onChange={e=>setActif(e.target.checked)} style={{ width:18, height:18, accentColor:'#111' }} />
            <span style={{ fontWeight:600, fontSize:14 }}>Actif</span>
          </label>
        </div>
        <div style={{ display:'flex', gap:10, marginTop:24 }}>
          <button onClick={save} disabled={saving} style={{ flex:1, padding:16, borderRadius:14, border:'none', background:'#111', color:'#fff', fontSize:15, fontWeight:700, cursor:'pointer', fontFamily:'inherit', display:'flex', alignItems:'center', justifyContent:'center', gap:8 }}>{saving ? 'Sauvegarde…' : <><Save size={16} strokeWidth={2}/> Sauvegarder</>}</button>
          <button onClick={onClose} style={{ padding:'16px 20px', borderRadius:14, border:'1.5px solid #eee', background:'#fff', color:'#888', fontSize:14, fontWeight:600, cursor:'pointer', fontFamily:'inherit' }}>Annuler</button>
        </div>
      </div>
    </div>
  );
}

function RouePage({ showToast }) {
  const [loading, setLoading] = useState(true);
  const [recompenses, setRecompenses] = useState([]);
  const [gains, setGains] = useState([]);
  const [editSheet, setEditSheet] = useState(null);
  const [planModal, setPlanModal] = useState(null);
  const [planDate, setPlanDate] = useState('');
  const [planHeure, setPlanHeure] = useState('19:00');
  const [accordion, setAccordion] = useState('recompenses');
  const [filtreRec, setFiltreRec] = useState('all');
  const [filtreStatut, setFiltreStatut] = useState('all');
  const [filtrePeriode, setFiltrePeriode] = useState('mois');
  const [recherche, setRecherche] = useState('');
  const [rouеActive, setRoueActive] = useState(false);
  const [essaisMax, setEssaisMax] = useState(3);
  const [countdownSec, setCountdownSec] = useState(5);
  const [email1Delai, setEmail1Delai] = useState('1');
  const [email1Objet, setEmail1Objet] = useState(DEFAULT_EMAIL1_OBJET);
  const [email1Corps, setEmail1Corps] = useState(DEFAULT_EMAIL1_CORPS);
  const [email1Date, setEmail1Date] = useState('');
  const [email1DateFin, setEmail1DateFin] = useState('');
  const [email1DateMode, setEmail1DateMode] = useState('precise'); // 'precise' | 'periode'
  const [email1CalOpen, setEmail1CalOpen] = useState(null); // null | 'debut' | 'fin'
  const [email1CalMonth, setEmail1CalMonth] = useState(() => { const d = new Date(); return { y: d.getFullYear(), m: d.getMonth() }; });
  const [email1Message, setEmail1Message] = useState('');
  const [savingParam, setSavingParam] = useState(false);
  const [showEmail1TestModal, setShowEmail1TestModal] = useState(false);
  const [email1TestMail, setEmail1TestMail] = useState('');
  const email1CorpsRef = useRef(null);

  useEffect(() => {
    loadAll();
    const interval = setInterval(loadStats, 30000);
    return () => clearInterval(interval);
  }, []);

  async function loadAll() {
    setLoading(true);
    try {
      const [recRes, gainsRes, paramsRes] = await Promise.all([
        supabase.from('roue_recompenses').select('*').order('ordre'),
        supabase.from('roue_gains').select('*, roue_recompenses(nom,emoji)').order('date_gain', { ascending: false }).limit(500),
        supabase.from('roue_config').select('cle,valeur'),
      ]);
      if (recRes.error) console.error('[Jeux] roue_recompenses error:', recRes.error);
      if (gainsRes.error) console.error('[Jeux] roue_gains error:', gainsRes.error);
      if (paramsRes.error) console.error('[Jeux] roue_config error:', paramsRes.error);
      setRecompenses(recRes.data || []);
      setGains(gainsRes.data || []);
      const p = {};
      (paramsRes.data || []).forEach(r => { p[r.cle] = r.valeur; });
      if (p['roue_active'] !== undefined) setRoueActive(p['roue_active'] === 'true');
      if (p['roue_essais_max']) setEssaisMax(parseInt(p['roue_essais_max']));
      if (p['roue_countdown']) setCountdownSec(parseInt(p['roue_countdown']));
      if (p['roue_email1_delai']) setEmail1Delai(p['roue_email1_delai']);
      if (p['roue_email1_objet']) setEmail1Objet(p['roue_email1_objet']);
      if (p['roue_email1_corps']) setEmail1Corps(p['roue_email1_corps']);
      if (p['roue_email_date']) setEmail1Date(p['roue_email_date']);
      if (p['roue_email_date_fin']) setEmail1DateFin(p['roue_email_date_fin']);
      if (p['roue_email_date_mode']) setEmail1DateMode(p['roue_email_date_mode']);
      if (p['roue_email_message']) setEmail1Message(p['roue_email_message']);
    } finally {
      setLoading(false);
    }
  }

  async function loadStats() {
    const [recRes, gainsRes] = await Promise.all([
      supabase.from('roue_recompenses').select('*').order('ordre'),
      supabase.from('roue_gains').select('*, roue_recompenses(nom,emoji)').order('date_gain', { ascending: false }).limit(500),
    ]);
    if (!recRes.error) setRecompenses(recRes.data || []);
    if (!gainsRes.error) setGains(gainsRes.data || []);
  }

  async function saveParam(cle, valeur) {
    const { error } = await supabase.from('roue_config').upsert({ cle, valeur }, { onConflict: 'cle' });
    if (error) {
      console.error('saveParam error', cle, error);
      throw error;
    }
  }

  async function toggleRoueActive(v) {
    setRoueActive(v);
    try {
      await saveParam('roue_active', v ? 'true' : 'false');
      showToast(v ? '🟢 Jeux activé ✓' : '⚫ Jeux désactivé');
    } catch(e) {
      setRoueActive(!v); // rollback visuel
      showToast('❌ Erreur sauvegarde — vérifiez les permissions Supabase');
    }
  }

  async function saveConfigBase() {
    setSavingParam(true);
    await Promise.all([saveParam('roue_essais_max', String(essaisMax)), saveParam('roue_countdown', String(countdownSec))]);
    setSavingParam(false);
    showToast('Configuration sauvegardée ✓');
  }

  async function saveEmail1() {
    setSavingParam(true);
    await Promise.all([
      saveParam('roue_email1_delai', email1Delai),
      saveParam('roue_email1_objet', email1Objet),
      saveParam('roue_email1_corps', email1Corps),
    ]);
    setSavingParam(false);
    showToast('Email 1 sauvegardé ✓');
  }

  async function sendEmail1Test() {
    if (!email1TestMail) return;
    try {
      const res = await fetch('/api/roue-email', { method:'POST', headers: await crmAuthHeaders(),
        body: JSON.stringify({ type:'email1', to_email:email1TestMail, to_prenom:'Test', to_nom:'TED', recompense:'Magnum de rosé', emoji:'🍾' }) });
      if (res.ok) { showToast('Email de test envoyé ✓'); setShowEmail1TestModal(false); setEmail1TestMail(''); }
      else showToast('❌ Erreur envoi email');
    } catch { showToast('❌ Erreur envoi email'); }
  }

  async function handleSaveEmail1() {
    setSavingParam(true);
    try {
      await Promise.all([
        saveParam('roue_email1_delai', email1Delai),
        saveParam('roue_email1_objet', email1Objet),
        saveParam('roue_email_date', email1Date),
        saveParam('roue_email_date_fin', email1DateFin),
        saveParam('roue_email_date_mode', email1DateMode),
        saveParam('roue_email_message', email1Message),
      ]);
      showToast('✅ Modifications enregistrées');
    } catch(e) {
      console.error('handleSaveEmail1 error', e);
      showToast('❌ Erreur sauvegarde email');
    } finally {
      setSavingParam(false);
    }
  }

  async function toggleRecupere(g) {
    const v = !g.recupere;
    await supabase.from('roue_gains').update({ recupere: v }).eq('id', g.id);
    showToast(v ? '✅ Marqué récupéré' : 'Annulé');
    setGains(prev => prev.map(x => x.id === g.id ? { ...x, recupere: v } : x));
  }

  async function toggleRecompenseActif(r) {
    const v = !r.actif;
    await supabase.from('roue_recompenses').update({ actif: v }).eq('id', r.id);
    setRecompenses(prev => prev.map(x => x.id === r.id ? { ...x, actif: v } : x));
  }

  async function supprimerRecompense(r) {
    if (!window.confirm(`Supprimer "${r.nom}" ?`)) return;
    await supabase.from('roue_recompenses').delete().eq('id', r.id);
    showToast('Récompense supprimée');
    setRecompenses(prev => prev.filter(x => x.id !== r.id));
  }

  async function supprimerGain(g) {
    if (!window.confirm('Supprimer ce participant du tableau des jeux ?')) return;
    // Supprime uniquement dans roue_gains, ne touche pas clients ni reservations
    await supabase.from('roue_gains').delete().eq('id', g.id);
    setGains(prev => prev.filter(x => x.id !== g.id));
    showToast('Participant supprimé');
  }

  async function planifierVisite() {
    if (!planDate) return;
    await supabase.from('roue_gains').update({
      date_venue: planDate,
      heure_venue: planHeure || null,
    }).eq('id', planModal.id);
    showToast('Date de venue enregistrée ✓');
    setPlanModal(null);
    await loadAll();
  }

  function exportCSV() {
    const data = getFilteredGains();
    const headers = ['Date','Prénom','Nom','Téléphone','Email','Récompense','Récupéré','Date venue'];
    const rows = data.map(g => [
      g.date_gain ? new Date(g.date_gain).toLocaleDateString('fr-FR') : '',
      g.prenom||'', g.nom||'', g.tel||'', g.email||'',
      g.roue_recompenses?.nom||'',
      g.recupere ? 'Oui' : 'Non',
      g.date_venue||'',
    ]);
    const csv = [headers,...rows].map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
    const blob = new Blob(['﻿'+csv], { type:'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `roue_gains_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
  }

  function getFilteredGains() {
    let g = [...gains];
    if (filtreRec !== 'all') g = g.filter(x => x.recompense_id === filtreRec);
    const isPerdu = x => (x.roue_recompenses?.nom||'').toLowerCase().includes('perdu');
    if (filtreStatut === 'gagnants') g = g.filter(x => !isPerdu(x));
    else if (filtreStatut === 'perdants') g = g.filter(x => isPerdu(x));
    else if (filtreStatut === 'recuperes') g = g.filter(x => x.recupere);
    const now = new Date();
    if (filtrePeriode === 'today') g = g.filter(x => new Date(x.date_gain).toDateString() === now.toDateString());
    else if (filtrePeriode === 'semaine') { const d = new Date(now - 7*86400000); g = g.filter(x => new Date(x.date_gain) >= d); }
    else if (filtrePeriode === 'mois') g = g.filter(x => { const d = new Date(x.date_gain); return d.getMonth()===now.getMonth() && d.getFullYear()===now.getFullYear(); });
    if (recherche.trim()) {
      const q = recherche.toLowerCase();
      g = g.filter(x => (x.prenom||'').toLowerCase().includes(q)||(x.nom||'').toLowerCase().includes(q)||(x.email||'').toLowerCase().includes(q)||(x.tel||'').includes(q));
    }
    return g;
  }

  const isPerdu = x => (x.roue_recompenses?.nom||'').toLowerCase().includes('perdu');
  const total = gains.length;
  const gainsVrais = gains.filter(x => !isPerdu(x));
  const tauxGain = total > 0 ? Math.round(gainsVrais.length / total * 100) : 0;
  const now2 = new Date();
  const gainsMois = gainsVrais.filter(x => { const d = new Date(x.date_gain); return d.getMonth()===now2.getMonth()&&d.getFullYear()===now2.getFullYear(); }).length;
  const formComplete = gains.filter(g => g.tel && g.email).length;
  const filteredGains = getFilteredGains();

  const card = { background:'#fff', borderRadius:16, padding:'20px 24px', boxShadow:'0 2px 8px rgba(0,0,0,0.07)' };
  const iS = { width:'100%', padding:'10px 14px', borderRadius:10, border:'1.5px solid #eee', fontSize:14, fontFamily:'inherit', outline:'none', background:'#fafafa' };
  const btnN = { padding:'10px 20px', borderRadius:10, border:'none', background:'#111', color:'#fff', fontSize:14, fontWeight:700, cursor:'pointer', fontFamily:'inherit' };
  const btnG = { padding:'9px 16px', borderRadius:10, border:'1.5px solid #e0e0e0', background:'#fff', color:'#555', fontSize:13, fontWeight:600, cursor:'pointer', fontFamily:'inherit' };

  if (loading) return <div style={{ padding:40, textAlign:'center', color:'#999' }}>Chargement…</div>;

  return (
    <div style={{ padding:'24px 32px 80px', boxSizing:'border-box' }}>
      <div style={{ display:'flex', alignItems:'center', gap:12, marginBottom:24 }}>
        <Dices size={28} strokeWidth={1.8} color="#111"/>
        <h1 style={{ margin:0, fontSize:28, fontWeight:900, color:'#111' }}>Jeux</h1>
        <div style={{ marginLeft:'auto', display:'flex', alignItems:'center', gap:10 }}>
          <a href="/accueil.html?preview=roue" target="_blank" rel="noopener noreferrer" style={{ display:'flex', alignItems:'center', gap:6, fontSize:13, fontWeight:600, color:'#555', textDecoration:'none', border:'1px solid #ddd', borderRadius:8, padding:'5px 12px', background:'#fff', marginRight:8 }}>
            <ExternalLink size={13} strokeWidth={2} /> Voir le jeu
          </a>
          <span style={{ fontSize:13, fontWeight:600, color: rouеActive ? '#111' : '#aaa' }}>{rouеActive ? <><div style={{width:8,height:8,borderRadius:'50%',background:'#22c55e',display:'inline-block',marginRight:4,verticalAlign:'middle'}} />Jeux actif</> : <><div style={{width:8,height:8,borderRadius:'50%',background:'#888',display:'inline-block',marginRight:4,verticalAlign:'middle'}} />Jeux inactif</>}</span>
          <div onClick={() => toggleRoueActive(!rouеActive)} style={{ width:48, height:26, borderRadius:13, background: rouеActive ? '#E8C547' : '#ddd', position:'relative', cursor:'pointer', transition:'background .2s', flexShrink:0 }}>
            <div style={{ position:'absolute', top:3, left: rouеActive ? 25 : 3, width:20, height:20, borderRadius:'50%', background:'#fff', boxShadow:'0 1px 4px rgba(0,0,0,0.25)', transition:'left .2s' }} />
          </div>
        </div>
      </div>

      {/* Stats */}
      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit, minmax(180px, 1fr))', gap:16, marginBottom:28 }}>
        {[
          { icon:<Dices size={20} strokeWidth={2} color="#E8C547"/>, bg:'#fffbea', label:'PARTIES JOUÉES', val:total },
          { icon:<Star size={20} strokeWidth={2} color="#8b5cf6"/>, bg:'#f5f3ff', label:'TAUX DE GAIN', val:`${tauxGain} %` },
          { icon:<Trophy size={20} strokeWidth={2} color="#22c55e"/>, bg:'#f0fdf4', label:'GAGNANTS CE MOIS', val:gainsMois },
          { icon:<ClipboardList size={20} strokeWidth={2} color="#3b82f6"/>, bg:'#eff6ff', label:'FORMULAIRES COMPLETS', val:formComplete },
        ].map((s,i) => (
          <div key={i} style={{ background:'#fff', borderRadius:16, padding:'16px 20px' }}>
            <div style={{ width:36, height:36, borderRadius:10, background:s.bg, display:'flex', alignItems:'center', justifyContent:'center', marginBottom:10 }}>{s.icon}</div>
            <p style={{ fontSize:10, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:0.5, margin:'0 0 4px' }}>{s.label}</p>
            <p style={{ fontSize:22, fontWeight:900, color:'#111', margin:0 }}>{s.val}</p>
          </div>
        ))}
      </div>

      {/* Config accordion */}
      <div style={{ ...card, marginBottom:28, padding:0, overflow:'hidden' }}>
        {/* Récompenses */}
        <div>
          <button onClick={() => setAccordion(a => a==='recompenses' ? '' : 'recompenses')}
            style={{ width:'100%', padding:'18px 24px', background:'none', border:'none', display:'flex', justifyContent:'space-between', alignItems:'center', cursor:'pointer', fontSize:15, fontWeight:700, color:'#111' }}>
            <span style={{display:'flex',alignItems:'center',gap:8}}><Settings size={16} strokeWidth={2} color="#888" /> Récompenses &amp; paramètres</span>
            <ChevronDown size={18} style={{ transform: accordion==='recompenses' ? 'rotate(180deg)' : 'none', transition:'transform .2s' }} />
          </button>
          {accordion === 'recompenses' && (
            <div style={{ padding:'0 24px 24px', borderTop:'1px solid #f0f0f0' }}>
              <div style={{ display:'flex', gap:20, alignItems:'center', flexWrap:'wrap', paddingTop:16, marginBottom:20 }}>
                <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                  <span style={{ fontSize:13, color:'#666', fontWeight:600 }}>Essais max :</span>
                  <input type="number" min={1} max={10} value={essaisMax} onChange={e=>setEssaisMax(+e.target.value)} onBlur={e=>saveParam('roue_essais_max', String(e.target.value)).catch(()=>{})} style={{ ...iS, width:60 }} />
                </div>
              </div>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:12 }}>
                <span style={{ fontWeight:700, fontSize:14 }}>Récompenses</span>
                <button onClick={() => setEditSheet({})} style={{ ...btnN, padding:'7px 14px', fontSize:13, display:'flex', alignItems:'center', gap:6 }}><Plus size={13} strokeWidth={2}/> Ajouter</button>
              </div>
              <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
                {recompenses.map(r => (
                  <div key={r.id} style={{ display:'flex', alignItems:'center', gap:12, padding:'12px 16px', background:'#fafafa', borderRadius:12, border:'1.5px solid #f0f0f0' }}>
                    <span style={{ fontSize:22, minWidth:28 }}>{r.emoji||'🎁'}</span>
                    <div style={{ flex:1, minWidth:0 }}>
                      <div style={{ fontWeight:700, fontSize:14, color:'#111' }}>{r.nom}</div>
                      <div style={{ fontSize:12, color:'#888' }}>
                        Proba: {r.probabilite||'—'} &bull; {r.stock_illimite ? 'Illimité' : `Stock: ${r.stock??'—'}`}
                        {!r.stock_illimite && r.stock!=null && r.stock<5 && <span style={{ marginLeft:6, background:'#ff4444', color:'#fff', borderRadius:6, padding:'1px 6px', fontSize:10, fontWeight:700 }}>Stock bas</span>}
                      </div>
                    </div>
                    <div onClick={() => toggleRecompenseActif(r)} style={{ width:36, height:20, borderRadius:10, background: r.actif ? '#E8C547' : '#ddd', position:'relative', cursor:'pointer', flexShrink:0, transition:'background .2s' }}>
                      <div style={{ position:'absolute', top:2, left: r.actif ? 18 : 2, width:16, height:16, borderRadius:'50%', background:'#fff', boxShadow:'0 1px 3px rgba(0,0,0,0.2)', transition:'left .2s' }} />
                    </div>
                    <button onClick={() => setEditSheet(r)} style={{ ...btnG, padding:'5px 10px', fontSize:12 }}><Pencil size={13} /></button>
                    <button onClick={() => supprimerRecompense(r)} style={{ ...btnG, padding:'5px 10px', fontSize:12, color:'#cc3333' }}><Trash2 size={13} /></button>
                  </div>
                ))}
                {recompenses.length === 0 && <div style={{ color:'#bbb', fontSize:13, textAlign:'center', padding:'20px 0' }}>Aucune récompense configurée</div>}
              </div>
            </div>
          )}
        </div>

        {/* Email 1 */}
        <div style={{ borderTop:'1px solid #f0f0f0' }}>
          <button onClick={() => setAccordion(a => a==='email1' ? '' : 'email1')}
            style={{ width:'100%', padding:'18px 24px', background:'none', border:'none', display:'flex', justifyContent:'space-between', alignItems:'center', cursor:'pointer', fontSize:15, fontWeight:700, color:'#111' }}>
            <span style={{display:'flex',alignItems:'center',gap:8}}><Mail size={16} strokeWidth={2} color="#888" /> Email de confirmation du gain</span>
            <ChevronDown size={18} style={{ transform: accordion==='email1' ? 'rotate(180deg)' : 'none', transition:'transform .2s' }} />
          </button>
          {accordion === 'email1' && (() => {
            // ── helpers date ──────────────────────────────────────────────
            const fmtDate = iso => iso
              ? new Date(iso + 'T00:00:00').toLocaleDateString('fr-FR', { day:'numeric', month:'long', year:'numeric' })
              : null;
            const dateDebut = fmtDate(email1Date);
            const dateFin   = fmtDate(email1DateFin);
            const dateAffichee = email1DateMode === 'precise'
              ? (dateDebut || 'À définir par le restaurant')
              : (dateDebut && dateFin ? `Du ${dateDebut} au ${dateFin}` : dateDebut || dateFin || 'À définir par le restaurant');
            const dispoLabel = email1DateMode === 'precise'
              ? (dateDebut ? `Disponible à partir du ${dateDebut}` : 'Disponible à partir du — À définir')
              : (dateDebut && dateFin ? `Disponible du ${dateDebut} au ${dateFin}` : 'Disponible — dates à définir');

            // ── message perso pour aperçu ──────────────────────────────
            const previewMessage = email1Message
              ? email1Message.replace(/\n/g, '<br>')
              : '';

            // ── mini calendrier ────────────────────────────────────────
            const CalPicker = ({ which }) => {
              const { y, m } = email1CalMonth;
              const firstDow = new Date(y, m, 1).getDay();
              const offset = firstDow === 0 ? 6 : firstDow - 1;
              const daysInMonth = new Date(y, m + 1, 0).getDate();
              const moisFR = ['Janvier','Février','Mars','Avril','Mai','Juin','Juillet','Août','Septembre','Octobre','Novembre','Décembre'];
              const joursFR = ['Lu','Ma','Me','Je','Ve','Sa','Di'];
              const currentVal = which === 'debut' ? email1Date : email1DateFin;
              const cells = [];
              for (let i = 0; i < offset; i++) cells.push(null);
              for (let d = 1; d <= daysInMonth; d++) cells.push(d);
              return (
                <div style={{ position:'absolute', top:'calc(100% + 6px)', left:0, zIndex:1100, background:'#fff', border:'1.5px solid #eee', borderRadius:12, boxShadow:'0 8px 32px rgba(0,0,0,0.13)', padding:14, width:260 }}>
                  <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:10 }}>
                    <button onClick={()=>setEmail1CalMonth(({y,m})=>m===0?{y:y-1,m:11}:{y,m:m-1})} style={{ background:'none',border:'none',cursor:'pointer',fontSize:16,padding:'2px 8px',borderRadius:6,color:'#555' }}>‹</button>
                    <span style={{ fontWeight:700, fontSize:13, color:'#111' }}>{moisFR[m]} {y}</span>
                    <button onClick={()=>setEmail1CalMonth(({y,m})=>m===11?{y:y+1,m:0}:{y,m:m+1})} style={{ background:'none',border:'none',cursor:'pointer',fontSize:16,padding:'2px 8px',borderRadius:6,color:'#555' }}>›</button>
                  </div>
                  <div style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', gap:2, marginBottom:4 }}>
                    {joursFR.map(j=><div key={j} style={{ textAlign:'center', fontSize:10, fontWeight:700, color:'#aaa', padding:'2px 0' }}>{j}</div>)}
                  </div>
                  <div style={{ display:'grid', gridTemplateColumns:'repeat(7,1fr)', gap:2 }}>
                    {cells.map((d,i) => {
                      if (!d) return <div key={i} />;
                      const iso = `${y}-${String(m+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
                      const selected = currentVal === iso;
                      return (
                        <div key={i} onClick={()=>{ if(which==='debut') setEmail1Date(iso); else setEmail1DateFin(iso); setEmail1CalOpen(null); }}
                          style={{ textAlign:'center', padding:'5px 0', borderRadius:7, cursor:'pointer', fontSize:12, fontWeight: selected?700:400, background: selected?'#E8C547':'transparent', color: selected?'#111':'#333' }}
                          onMouseEnter={e=>{ if(!selected) e.currentTarget.style.background='#fdf6d8'; }}
                          onMouseLeave={e=>{ if(!selected) e.currentTarget.style.background='transparent'; }}
                        >{d}</div>
                      );
                    })}
                  </div>
                </div>
              );
            };

            // ── aperçu HTML (nouveau design premium) ───────────────────
            const serialPreview = 'TED-' + Math.random().toString(36).slice(2,7).toUpperCase();
            const previewHtml = `<div style="font-family:-apple-system,'Helvetica Neue',Arial,sans-serif;background:#181818;padding:24px 12px;">
  <div style="max-width:480px;margin:0 auto;">

    <!-- HEADER -->
    <div style="background:#111111;border-radius:10px 10px 0 0;padding:32px 28px 24px;text-align:center;">
      <img src="/logo-Le-TED.png" width="60" height="60" style="display:block;margin:0 auto 12px;border-radius:50%;" />
      <div style="font-weight:800;font-size:18px;letter-spacing:8px;text-transform:uppercase;color:#F0A830;">LE TED</div>
      <div style="font-size:9px;letter-spacing:3px;color:rgba(255,255,255,0.25);text-transform:uppercase;margin-top:4px;">Restaurant &amp; Club · Chassieu</div>
    </div>

    <!-- Gold bar -->
    <div style="height:3px;background:linear-gradient(90deg,#c47e10,#F0A830,#ffd278,#F0A830,#c47e10);"></div>

    <!-- WINNER HERO -->
    <div style="background:#111111;padding:32px 28px 28px;text-align:center;">
      <div style="font-weight:800;font-size:10px;letter-spacing:5px;color:#F0A830;text-transform:uppercase;margin-bottom:10px;">✦ Vous avez gagné ✦</div>
      <div style="font-size:28px;font-weight:700;color:#ffffff;margin-bottom:6px;line-height:1.15;">Prénom Nom</div>
      <div style="font-size:12px;color:rgba(255,255,255,0.45);margin-bottom:28px;">vous repart avec une récompense exclusive</div>

      <!-- TICKET -->
      <div style="border:1px solid rgba(240,168,48,0.35);border-radius:14px;overflow:hidden;margin-bottom:28px;">
        <div style="height:4px;background:linear-gradient(90deg,#c47e10,#F0A830,#ffd278,#F0A830,#c47e10);"></div>
        <div style="background:#1a1300;padding:10px 18px;display:flex;justify-content:space-between;align-items:center;">
          <div style="display:flex;align-items:center;gap:7px;">
            <img src="/logo-Le-TED.png" width="18" height="18" style="opacity:0.85;vertical-align:middle;" />
            <span style="font-size:8px;font-weight:700;letter-spacing:3px;color:#F0A830;text-transform:uppercase;">Bon Gagnant</span>
          </div>
          <span style="font-family:'Courier New',monospace;font-size:9px;color:rgba(240,168,48,0.35);">${serialPreview}</span>
        </div>
        <div style="background:#111;border-top:1px dashed rgba(240,168,48,0.2);height:1px;"></div>
        <div style="background:#111111;padding:28px 20px 22px;text-align:center;">
          <div style="font-size:9px;letter-spacing:6px;color:rgba(240,168,48,0.4);margin-bottom:14px;">✦ &nbsp; ✦ &nbsp; ✦</div>
          <div style="margin-bottom:14px;display:flex;justify-content:center;"><svg width="52" height="52" viewBox="0 0 24 24" fill="none" stroke="#F0A830" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 12 20 22 4 22 4 12"/><rect x="2" y="7" width="20" height="5"/><line x1="12" y1="22" x2="12" y2="7"/><path d="M12 7H7.5a2.5 2.5 0 0 1 0-5C11 2 12 7 12 7z"/><path d="M12 7h4.5a2.5 2.5 0 0 0 0-5C13 2 12 7 12 7z"/></svg></div>
          <div style="font-size:8px;font-weight:600;letter-spacing:4px;color:rgba(240,168,48,0.5);text-transform:uppercase;margin-bottom:8px;">Votre récompense</div>
          <div style="font-size:20px;font-weight:700;color:#ffffff;margin-bottom:18px;line-height:1.2;">La récompense gagnée apparaîtra ici</div>
          <div style="display:inline-block;background:rgba(240,168,48,0.08);border:1px solid rgba(240,168,48,0.2);border-radius:40px;padding:7px 16px;margin-bottom:16px;">
            <span style="font-size:12px;color:rgba(255,255,255,0.5);">Valable </span>
            <span style="font-size:12px;font-weight:700;color:#F0A830;">${dateAffichee}</span>
          </div>
        </div>
        <div style="background:#111;padding:6px 0;display:flex;align-items:center;">
          <div style="width:12px;height:12px;background:#181818;border-radius:0 50% 50% 0;border:1px solid rgba(240,168,48,0.2);"></div>
          <div style="flex:1;border-top:1.5px dashed rgba(240,168,48,0.2);"></div>
          <div style="width:12px;height:12px;background:#181818;border-radius:50% 0 0 50%;border:1px solid rgba(240,168,48,0.2);"></div>
        </div>
        <div style="background:#0e0e0e;padding:10px 18px;text-align:center;">
          <span style="font-size:10px;color:rgba(255,255,255,0.4);text-decoration:underline;">Conditions de retrait ci-dessous ↓</span>
        </div>
        <div style="height:3px;background:linear-gradient(90deg,#c47e10,#F0A830,#ffd278,#F0A830,#c47e10);opacity:0.6;"></div>
      </div>

      ${previewMessage ? `<div style="margin-bottom:28px;padding:14px 18px;border-left:2px solid rgba(240,168,48,0.4);text-align:left;background:rgba(255,255,255,0.03);border-radius:0 6px 6px 0;"><p style="font-style:italic;font-size:14px;color:rgba(255,255,255,0.6);margin:0 0 6px;line-height:1.7;">« ${previewMessage} »</p><p style="font-size:9px;letter-spacing:2px;color:rgba(240,168,48,0.5);text-transform:uppercase;margin:0;">L'équipe du TED</p></div>` : ''}
    </div>

    <!-- CONDITIONS -->
    <div style="background:#FDFAF5;padding:28px 24px 24px;">
      <div style="font-size:16px;font-weight:700;letter-spacing:0.5px;color:#999;text-transform:uppercase;margin-bottom:18px;">Conditions de retrait</div>
      <div style="display:flex;gap:12px;padding:12px 0;border-bottom:1px solid #f0ece4;align-items:flex-start;">
        <div style="width:34px;height:34px;background:#111;border-radius:4px;display:flex;align-items:center;justify-content:center;flex-shrink:0;"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#F0A830" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1" ry="1"/></svg></div>
        <div><div style="font-size:13px;font-weight:700;color:#111;margin-bottom:2px;">Présentation obligatoire</div><div style="font-size:11px;color:#777;line-height:1.6;">Présentez <strong style="color:#111;">cet email</strong> à votre arrivée.</div></div>
      </div>
      <div style="display:flex;gap:12px;padding:12px 0;border-bottom:1px solid #f0ece4;align-items:flex-start;">
        <div style="width:34px;height:34px;background:#111;border-radius:4px;display:flex;align-items:center;justify-content:center;flex-shrink:0;"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#F0A830" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg></div>
        <div><div style="font-size:13px;font-weight:700;color:#111;margin-bottom:2px;">Date de retrait</div><div style="font-size:11px;color:#777;line-height:1.6;">${dateAffichee}</div></div>
      </div>
      <div style="display:flex;gap:12px;padding:12px 0;align-items:flex-start;">
        <div style="width:34px;height:34px;background:#111;border-radius:4px;display:flex;align-items:center;justify-content:center;flex-shrink:0;"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#F0A830" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg></div>
        <div><div style="font-size:13px;font-weight:700;color:#111;margin-bottom:2px;">5 personnes minimum</div><div style="font-size:11px;color:#777;line-height:1.6;">Valable en groupe d'<strong style="color:#111;">au moins 5 personnes</strong>.</div></div>
      </div>
      <div style="margin-top:24px;text-align:center;">
        <a href="#" style="display:inline-block;background:#F0A830;color:#111111;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;text-decoration:none;padding:14px 36px;border-radius:4px;">Réserver ma table</a>
        <div style="font-size:11px;color:#bbb;margin-top:10px;">ou appelez-nous · <span style="color:#F0A830;font-weight:600;">04 78 90 67 80</span></div>
      </div>
      <div style="margin-top:32px;padding-top:24px;border-top:1px solid #ece8e0;text-align:center;">
        <div style="font-style:italic;font-size:15px;color:#555;line-height:1.85;margin-bottom:4px;">On vous attend avec impatience.</div>
        <div style="font-size:13px;color:#777;margin-bottom:12px;">À très bientôt,</div>
        <div style="font-family:Georgia,serif;font-size:26px;color:#F0A830;font-style:italic;">L'équipe du TED</div>
      </div>
    </div>

    <!-- FOOTER -->
    <div style="background:#0d0d0d;border-radius:0 0 10px 10px;padding:24px 28px;text-align:center;">
      <img src="/logo-Le-TED.png" width="36" height="36" style="display:block;margin:0 auto 10px;border-radius:50%;opacity:0.7;" />
      <div style="font-weight:800;font-size:11px;letter-spacing:6px;color:#F0A830;margin-bottom:12px;text-transform:uppercase;">Le Ted</div>
      <div style="font-size:10px;color:rgba(255,255,255,0.25);line-height:2;margin-bottom:8px;">28 Avenue des Frères Montgolfier, 69680 Chassieu<br><span style="color:rgba(240,168,48,0.5);">04 78 90 67 80</span> · <span style="color:rgba(240,168,48,0.5);">leted.fr</span></div>
      <div style="font-size:9px;color:rgba(255,255,255,0.12);line-height:1.7;">© 2026 Le TED — Restaurant &amp; Club — Chassieu, Lyon</div>
    </div>

  </div>
</div>`;

            return (
            <div style={{ padding:'0 24px 24px', borderTop:'1px solid #f0f0f0' }}>
              {email1CalOpen && <div onClick={()=>setEmail1CalOpen(null)} style={{ position:'fixed', inset:0, zIndex:1099 }} />}
              <div style={{ paddingTop:16, display:'flex', gap:24, flexWrap:'wrap' }}>

                {/* ── Colonne gauche : éditeur ── */}
                <div style={{ flex:'1 1 320px', display:'flex', flexDirection:'column', gap:14 }}>
                  <div>
                    <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Délai après le jeu (minutes)</label>
                    <input type="number" min={0} value={email1Delai} onChange={e=>setEmail1Delai(e.target.value)} style={{ ...iS, marginTop:6 }} />
                  </div>
                  <div>
                    <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Objet</label>
                    <input type="text" value={email1Objet} onChange={e=>setEmail1Objet(e.target.value)} style={{ ...iS, marginTop:6 }} />
                  </div>

                  {/* Date picker avec mode toggle */}
                  <div>
                    <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Date de retrait du cadeau</label>
                    {/* Toggle pills */}
                    <div style={{ display:'flex', gap:6, marginTop:8, marginBottom:10 }}>
                      {[['precise','Date précise'],['periode','Période']].map(([mode,label])=>(
                        <button key={mode} onClick={()=>setEmail1DateMode(mode)}
                          style={{ padding:'5px 14px', borderRadius:20, border:'1.5px solid', fontSize:12, fontWeight:600, cursor:'pointer',
                            borderColor: email1DateMode===mode ? '#E8C547' : '#e0e0e0',
                            background: email1DateMode===mode ? '#E8C547' : '#fff',
                            color: email1DateMode===mode ? '#111' : '#888' }}>
                          {label}
                        </button>
                      ))}
                    </div>

                    {email1DateMode === 'precise' ? (
                      <div style={{ position:'relative' }}>
                        <div onClick={()=>setEmail1CalOpen(c=>c==='debut'?null:'debut')}
                          style={{ ...iS, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', userSelect:'none' }}>
                          <span style={{ color: email1Date ? '#111' : '#aaa' }}>{dateDebut || 'Choisir une date…'}</span>
                          <CalendarDays size={14} style={{display:'inline',verticalAlign:'middle'}} />
                        </div>
                        {email1CalOpen === 'debut' && <CalPicker which="debut" />}
                      </div>
                    ) : (
                      <div style={{ display:'flex', gap:10, flexWrap:'wrap' }}>
                        <div style={{ flex:'1 1 120px', position:'relative' }}>
                          <div style={{ fontSize:11, color:'#888', fontWeight:600, marginBottom:4 }}>DU</div>
                          <div onClick={()=>setEmail1CalOpen(c=>c==='debut'?null:'debut')}
                            style={{ ...iS, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', userSelect:'none' }}>
                            <span style={{ color: email1Date ? '#111' : '#aaa', fontSize:13 }}>{dateDebut || 'Début…'}</span>
                            <CalendarDays size={13} style={{display:'inline',verticalAlign:'middle'}} />
                          </div>
                          {email1CalOpen === 'debut' && <CalPicker which="debut" />}
                        </div>
                        <div style={{ flex:'1 1 120px', position:'relative' }}>
                          <div style={{ fontSize:11, color:'#888', fontWeight:600, marginBottom:4 }}>AU</div>
                          <div onClick={()=>setEmail1CalOpen(c=>c==='fin'?null:'fin')}
                            style={{ ...iS, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'space-between', userSelect:'none' }}>
                            <span style={{ color: email1DateFin ? '#111' : '#aaa', fontSize:13 }}>{dateFin || 'Fin…'}</span>
                            <CalendarDays size={13} style={{display:'inline',verticalAlign:'middle'}} />
                          </div>
                          {email1CalOpen === 'fin' && <CalPicker which="fin" />}
                        </div>
                      </div>
                    )}
                  </div>

                  <div>
                    <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Message personnalisé <span style={{ fontWeight:400, textTransform:'none', color:'#aaa' }}>(optionnel)</span></label>
                    <textarea value={email1Message} onChange={e=>setEmail1Message(e.target.value)} rows={3} placeholder="Ex: Nous vous contacterons pour confirmer votre créneau..." style={{ ...iS, marginTop:6, resize:'none', lineHeight:1.6 }} />
                  </div>
                  <div style={{ display:'flex', justifyContent:'space-between', gap:10 }}>
                    <button onClick={() => setShowEmail1TestModal(true)} style={{ ...btnG, display:'flex', alignItems:'center', gap:6, border:'1.5px solid #E8C547' }}><Send size={13} strokeWidth={2}/> Envoyer un test</button>
                    <button onClick={handleSaveEmail1} disabled={savingParam} style={{ padding:'12px 24px', borderRadius:12, border:'none', background:'#E8C547', color:'#111', fontSize:14, fontWeight:700, cursor:savingParam?'not-allowed':'pointer', display:'flex', alignItems:'center', gap:6, opacity:savingParam?0.7:1 }}><Save size={13} strokeWidth={2}/> Enregistrer</button>
                  </div>
                </div>

                {/* ── Colonne droite : aperçu live ── */}
                <div style={{ flex:'1 1 320px' }}>
                  <p style={{ fontSize:11, fontWeight:600, color:'#999', textTransform:'uppercase', letterSpacing:'.05em', margin:'0 0 4px' }}>Aperçu du mail</p>
                  <p style={{ fontSize:11, color:'#bbb', margin:'0 0 8px' }}>Objet : <span style={{ color:'#666' }}>{email1Objet || '—'}</span></p>
                  <div
                    style={{ border:'1px solid #eee', borderRadius:12, background:'#fff', maxHeight:600, overflowY:'auto' }}
                    dangerouslySetInnerHTML={{ __html: previewHtml }}
                  />
                </div>
              </div>
            </div>
            );
          })()}
          {/* Modal test email */}
          {showEmail1TestModal && (
            <>
              <div onClick={() => setShowEmail1TestModal(false)} style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:8000 }} />
              <div onClick={e=>e.stopPropagation()} style={{ position:'fixed', top:'50%', left:'50%', transform:'translate(-50%,-50%)', background:'#fff', borderRadius:16, padding:28, width:'min(400px,calc(100vw - 48px))', zIndex:8001, boxShadow:'0 32px 80px rgba(0,0,0,0.25)' }}>
                <h3 style={{ margin:'0 0 18px', fontSize:17, fontWeight:800, color:'#111' }}>Envoyer un email de test</h3>
                <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Adresse email de test</label>
                <input type="email" value={email1TestMail} onChange={e=>setEmail1TestMail(e.target.value)} placeholder="prenom@exemple.com" style={{ ...iS, marginTop:6, marginBottom:20 }} onKeyDown={e=>e.key==='Enter'&&sendEmail1Test()} autoFocus />
                <div style={{ display:'flex', gap:10 }}>
                  <button onClick={() => { setShowEmail1TestModal(false); setEmail1TestMail(''); }} style={{ flex:1, height:48, border:'1.5px solid #eee', borderRadius:12, background:'#fff', fontSize:14, fontWeight:600, cursor:'pointer', color:'#666' }}>Annuler</button>
                  <button onClick={sendEmail1Test} style={{ flex:1, height:48, border:'none', borderRadius:12, background:'#111', fontSize:14, fontWeight:700, cursor:'pointer', color:'#fff', display:'flex', alignItems:'center', justifyContent:'center', gap:6 }}><Send size={14} strokeWidth={2}/> Envoyer</button>
                </div>
              </div>
            </>
          )}
        </div>

      </div>

      {/* Filtres */}
      <div style={{ display:'flex', gap:10, flexWrap:'wrap', marginBottom:16, alignItems:'center' }}>
        <select value={filtreRec} onChange={e=>setFiltreRec(e.target.value)} style={{ ...iS, width:'auto', background:'#fff', cursor:'pointer' }}>
          <option value="all">Toutes les récompenses</option>
          {recompenses.map(r => <option key={r.id} value={r.id}>{r.emoji} {r.nom}</option>)}
        </select>
        <select value={filtreStatut} onChange={e=>setFiltreStatut(e.target.value)} style={{ ...iS, width:'auto', background:'#fff', cursor:'pointer' }}>
          <option value="all">Tous</option>
          <option value="gagnants">Gagnants</option>
          <option value="perdants">Perdants</option>
          <option value="recuperes">Récupérés</option>
        </select>
        <select value={filtrePeriode} onChange={e=>setFiltrePeriode(e.target.value)} style={{ ...iS, width:'auto', background:'#fff', cursor:'pointer' }}>
          <option value="all">Toute la période</option>
          <option value="today">Aujourd'hui</option>
          <option value="semaine">Cette semaine</option>
          <option value="mois">Ce mois</option>
        </select>
        <input type="text" placeholder="Rechercher…" value={recherche} onChange={e=>setRecherche(e.target.value)} style={{ ...iS, width:200, background:'#fff', paddingLeft:32, backgroundImage:'none', position:'relative' }} />
        <div style={{ marginLeft:'auto' }}>
          <button onClick={exportCSV} style={{ ...btnG, display:'flex', alignItems:'center', gap:6 }}>
            <Download size={14} /> Exporter CSV
          </button>
        </div>
      </div>

      {/* Tableau */}
      <div style={{ ...card, padding:0, overflow:'hidden' }}>
        <div style={{ overflowX:'auto' }}>
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:13 }}>
            <thead>
              <tr style={{ background:'#fafafa', borderBottom:'2px solid #f0f0f0' }}>
                {['Date','Prénom','Nom','Téléphone','Email','Récompense','Statut',''].map(h => (
                  <th key={h} style={{ padding:'12px 14px', textAlign:'left', fontWeight:700, fontSize:11, color:'#888', textTransform:'uppercase', letterSpacing:'.04em', whiteSpace:'nowrap' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredGains.map(g => (
                <tr key={g.id} style={{ borderBottom:'1px solid #f5f5f5' }}>
                  <td style={{ padding:'12px 14px', whiteSpace:'nowrap', color:'#555', fontSize:12 }}>
                    {g.date_gain ? new Date(g.date_gain).toLocaleDateString('fr-FR',{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'}) : '—'}
                  </td>
                  <td style={{ padding:'12px 14px', fontWeight:600 }}>{g.prenom||'—'}</td>
                  <td style={{ padding:'12px 14px' }}>{g.nom||'—'}</td>
                  <td style={{ padding:'12px 14px', fontFamily:'monospace', fontSize:12 }}>{g.tel||'—'}</td>
                  <td style={{ padding:'12px 14px', fontSize:12, maxWidth:160, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{g.email||'—'}</td>
                  <td style={{ padding:'12px 14px', whiteSpace:'nowrap' }}>
                    <span style={{ fontSize:16 }}>{g.roue_recompenses?.emoji||''}</span>{' '}
                    <span style={{ fontWeight:600 }}>{g.roue_recompenses?.nom||'—'}</span>
                  </td>
                  <td style={{ padding:'12px 14px' }}>
                    <button onClick={() => toggleRecupere(g)} style={{ padding:'4px 10px', borderRadius:8, border:'none', fontSize:11, fontWeight:700, cursor:'pointer', background: g.recupere ? '#e8f5e9' : '#fff3e0', color: g.recupere ? '#2e7d32' : '#e65100', whiteSpace:'nowrap' }}>
                      {g.recupere ? '✓ Récupéré' : 'En attente'}
                    </button>
                  </td>
                  <td style={{ padding:'12px 14px', whiteSpace:'nowrap' }}>
                    <button
                      onClick={() => supprimerGain(g)}
                      title="Supprimer ce participant"
                      style={{ background:'none', border:'none', cursor:'pointer', color:'#999', display:'flex', alignItems:'center', padding:4, borderRadius:6, transition:'color 0.2s' }}
                      onMouseEnter={e => e.currentTarget.style.color = '#ef4444'}
                      onMouseLeave={e => e.currentTarget.style.color = '#999'}
                    ><Trash2 size={15} strokeWidth={2} /></button>
                  </td>
                </tr>
              ))}
              {filteredGains.length === 0 && (
                <tr><td colSpan={9} style={{ padding:40, textAlign:'center', color:'#bbb', fontSize:14 }}>Aucun résultat</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Modal planification */}
      {planModal && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.45)', zIndex:999, display:'flex', alignItems:'center', justifyContent:'center' }}>
          <div style={{ background:'#fff', borderRadius:20, padding:32, width:340, boxShadow:'0 20px 60px rgba(0,0,0,0.25)' }}>
            <h3 style={{ margin:'0 0 20px', fontSize:18, fontWeight:900, display:'flex', alignItems:'center', gap:8 }}><CalendarDays size={18} style={{display:'inline',verticalAlign:'middle'}} /> Date de visite</h3>
            <div style={{ fontWeight:600, fontSize:14, marginBottom:4, color:'#444' }}>{planModal.prenom} {planModal.nom}</div>
            <div style={{ fontSize:13, color:'#888', marginBottom:20 }}>{planModal.roue_recompenses?.emoji} {planModal.roue_recompenses?.nom}</div>
            <div style={{ marginBottom:14 }}>
              <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Date de venue</label>
              <input type="date" value={planDate} onChange={e=>setPlanDate(e.target.value)} style={{ ...iS, marginTop:6 }} />
            </div>
            <div style={{ marginBottom:20 }}>
              <label style={{ fontSize:12, fontWeight:600, color:'#666', textTransform:'uppercase' }}>Heure</label>
              <input type="time" value={planHeure} onChange={e=>setPlanHeure(e.target.value)} style={{ ...iS, marginTop:6 }} />
            </div>
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={planifierVisite} style={{ ...btnN, flex:1, display:'flex', alignItems:'center', justifyContent:'center', gap:6 }}><Save size={14} strokeWidth={2}/> Enregistrer</button>
              <button onClick={() => setPlanModal(null)} style={{ ...btnG, flex:1 }}>Annuler</button>
            </div>
          </div>
        </div>
      )}

      {editSheet !== null && (
        <RecompenseSheet
          item={editSheet?.id ? editSheet : null}
          onClose={() => setEditSheet(null)}
          onSaved={() => { setEditSheet(null); loadAll(); }}
          showToast={showToast}
        />
      )}
    </div>
  );
}

// ─── Menu ────────────────────────────────────────────────────────────────────

function MenuPage({ showToast }) {
  const isMobile = useIsMobile();
  const [carte, setCarte] = useState('restaurant');
  const [cartes, setCartes] = useState([{id:'restaurant',l:'Restaurant'},{id:'brasero',l:'Brasero'}]);
  const [categories, setCategories] = useState([]);
  const [produits, setProduits] = useState([]);
  const [entreeJour, setEntreeJour] = useState(null);
  const [platJour, setPlatJour] = useState(null);
  const [dessertJour, setDessertJour] = useState(null);
  const [formuleJour, setFormuleJour] = useState(null);
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
  const [showOriginesSheet, setShowOriginesSheet] = useState(false);
  const [uploadingFlyer, setUploadingFlyer] = useState(false);
  const dragSoiree = useRef(null);
  const dragOverSoiree = useRef(null);

  useEffect(() => { loadMenu(); setMenuSearch(''); setOpenCats(new Set()); }, [carte]);
  useEffect(() => { loadSoirees(); loadCartes(); }, []);

  useEffect(() => {
    return resilientChannel(supabase, 'menu-rt', (chan) => chan
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_produits' }, () => loadMenu())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_categories' }, () => loadMenu())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_plat_jour' }, () => loadMenu())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_cartes' }, () => loadCartes())
      .on('postgres_changes', { event: '*', schema: 'public', table: 'menu_soirees' }, () => loadSoirees())
    );
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

  async function uploadFlyer(file) {
    if (!file) return null;
    if (file.size > 5 * 1024 * 1024) { showToast('❌ Image trop lourde (max 5 Mo)'); return null; }
    setUploadingFlyer(true);
    try {
      const ext = file.name.split('.').pop().toLowerCase();
      const path = `flyer-${Date.now()}.${ext}`;
      const { error } = await supabase.storage.from('soirees-flyers').upload(path, file, { upsert: true, contentType: file.type });
      if (error) { showToast('❌ Erreur upload image'); return null; }
      const { data: urlData } = supabase.storage.from('soirees-flyers').getPublicUrl(path);
      return urlData?.publicUrl || null;
    } finally {
      setUploadingFlyer(false);
    }
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
      safeQuery(() => supabase.from('menu_categories').select('*').order('ordre'), { fallback: [], context: 'loadMenu:categories' }),
      safeQuery(() => supabase.from('menu_produits').select('*').order('ordre'), { fallback: [], context: 'loadMenu:produits' }),
      safeQuery(() => supabase.from('menu_plat_jour').select('*'), { fallback: [], context: 'loadMenu:platJour' })
    ]);
    setCategories(cR.data || []);
    setProduits(pR.data || []);
    let pj = jR.data || [];

    // Auto-créer les entrées manquantes pour cette carte
    const types = ['entree', 'plat', 'dessert', 'formule', 'suggestion'];
    const missing = types.filter(t => !pj.find(p => p.type === t && p.carte === carte));
    if (missing.length > 0) {
      await Promise.all(missing.map(t =>
        supabase.from('menu_plat_jour').insert({ type: t, carte, actif: false, nom: '', prix: null })
      ));
      const { data: refreshed } = await supabase.from('menu_plat_jour').select('*');
      pj = refreshed || pj;
    }

    setEntreeJour(pj.find(p => p.type === 'entree' && p.carte === carte) || null);
    setPlatJour(pj.find(p => p.type === 'plat' && p.carte === carte) || null);
    setDessertJour(pj.find(p => p.type === 'dessert' && p.carte === carte) || null);
    setFormuleJour(pj.find(p => p.type === 'formule' && p.carte === carte) || null);
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

  const CORE_CARTES = ['restaurant', 'brasero'];
  const catsFiltered = categories.filter(c => c.visible !== false && (c.carte === carte || (CORE_CARTES.includes(carte) && c.carte === 'les-deux')) && c.nom !== 'Plat du jour');
  const searchQ = menuSearch.trim();

  function produitsForCat(catId) {
    const base = produits.filter(p => p.categorie_id === catId && (p.carte === carte || (CORE_CARTES.includes(carte) && p.carte === 'les-deux'))).sort((a,b) => (a.ordre||0)-(b.ordre||0));
    if (!searchQ) return base;
    return base.filter(p => normalizeStr(p.nom||'').includes(normalizeStr(searchQ)) || normalizeStr(p.description||'').includes(normalizeStr(searchQ)));
  }

  const catsVisible = searchQ ? catsFiltered.filter(c => produitsForCat(c.id).length > 0) : catsFiltered;
  const platItems = [platJour, dessertJour, suggestionJour].filter(Boolean);

  if (loading) return <div style={{ display:'flex', alignItems:'center', justifyContent:'center', height:'60vh', color:'#888', fontSize:15 }}>Chargement du menu...</div>;

  return (
    <div style={{ height: isMobile ? undefined : '100vh', minHeight: isMobile ? '100vh' : undefined, boxSizing:'border-box', overflow: isMobile ? undefined : 'hidden', background:'#f5f5f5', display:'flex', flexDirection:'column' }}>
      <div style={{ maxWidth:900, width:'100%', margin:'0 auto', padding: isMobile ? '0 16px' : '0 20px', display:'flex', flexDirection:'column', flex: isMobile ? undefined : 1, minHeight: isMobile ? undefined : 0, overflow: isMobile ? undefined : 'hidden', boxSizing:'border-box' }}>

      {/* Header */}
      <div style={{ padding: isMobile ? '16px 0 12px' : '24px 0 16px', display:'flex', alignItems:'center', justifyContent:'space-between', flexWrap:'wrap', gap:10 }}>
        <div>
          <h1 style={{ margin:0, fontSize: isMobile ? 22 : 28, fontWeight:900, color:'#111' }}>Menu</h1>
          <p style={{ margin:'3px 0 0', fontSize:13, color:'#aaa' }}>Gérez la carte en temps réel</p>
        </div>
        <div style={{ display:'flex', alignItems:'center', gap:8, flexWrap:'wrap' }}>
          {!isMobile && (
            <>
              <button onClick={() => setShowCartesSheet(true)} style={{ height:34, padding:'0 12px', borderRadius:8, border:'1.5px solid #eee', background:'#fff', fontSize:12, fontWeight:600, cursor:'pointer', color:'#666', display:'flex', alignItems:'center', gap:6 }}>
                <LayoutGrid size={14} strokeWidth={2} color="#666"/> Gérer vos cartes
              </button>
              <button onClick={() => setShowGererCats(true)} style={{ height:34, padding:'0 12px', borderRadius:8, border:'1.5px solid #eee', background:'#fff', fontSize:12, fontWeight:600, cursor:'pointer', color:'#666', display:'flex', alignItems:'center', gap:6 }}>
                <Settings size={14} strokeWidth={2} color="#666"/> Gérer les catégories
              </button>
              <button onClick={() => setShowOriginesSheet(true)} style={{ height:34, padding:'0 12px', borderRadius:8, border:'1.5px solid #eee', background:'#fff', fontSize:12, fontWeight:600, cursor:'pointer', color:'#666', display:'flex', alignItems:'center', gap:6 }}>
                <MapPin size={14} strokeWidth={2} color="#666"/> Origine des viandes
              </button>
            </>
          )}
          {isMobile && (
            <button onClick={() => setShowGererCats(true)} style={{ height:34, padding:'0 10px', borderRadius:8, border:'1.5px solid #eee', background:'#fff', fontSize:12, fontWeight:600, cursor:'pointer', color:'#666', display:'flex', alignItems:'center', gap:5 }}>
              <Settings size={14} strokeWidth={2} color="#666"/>
            </button>
          )}
          <button onClick={() => catsFiltered.length > 0 ? setCatPickerOpen(true) : setEditProduit({ carte, disponible: true, mise_en_avant: false, badges: [], allergenes: [] })} style={{ height:36, padding: isMobile ? '0 12px' : '0 16px', borderRadius:10, border:'none', background:'#E8C547', color:'#111', fontSize:13, fontWeight:800, cursor:'pointer', display:'flex', alignItems:'center', gap:8, boxShadow:'0 2px 8px rgba(232,197,71,0.3)' }}>
            <Plus size={16} strokeWidth={2}/> {isMobile ? 'Ajouter' : 'Ajouter un produit'}
          </button>
        </div>
      </div>

      {/* Zone défilante : tout sauf l'en-tête, qui reste visible */}
      <div style={{ flex: isMobile ? undefined : 1, minHeight: isMobile ? undefined : 0, overflowY: isMobile ? undefined : 'auto', paddingRight: isMobile ? 0 : 2 }}>

      {/* Onglets + lien carte client */}
      <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:20, gap:8 }}>
        <div style={{ display:'flex', gap:6, flexWrap:'wrap' }}>
          {cartes.map(c => (
            <button key={c.id} onClick={() => setCarte(c.id)} style={{ height:36, padding:'0 16px', borderRadius:10, fontWeight:700, fontSize:13, cursor:'pointer', border:'none', background: carte===c.id ? '#E8C547' : '#fff', color: carte===c.id ? '#111' : '#666', boxShadow: carte===c.id ? '0 2px 8px rgba(232,197,71,0.25)' : '0 1px 4px rgba(0,0,0,0.06)', transition:'all 0.15s' }}>{c.l}</button>
          ))}
        </div>
        <a href="/accueil.html" target="_blank" rel="noopener noreferrer" style={{ height:34, padding:'0 12px', borderRadius:8, border:'1.5px solid #eee', background:'#fff', fontSize:12, fontWeight:600, color:'#666', display:'flex', alignItems:'center', gap:6, textDecoration:'none', boxShadow:'0 1px 4px rgba(0,0,0,0.06)', flexShrink:0 }}>
          <ExternalLink size={12} strokeWidth={2} /> Carte client
        </a>
      </div>

      {/* ── Encarts Menu du jour + Suggestion ─────────────────────────── */}
      {(() => {
        const anyActif = !!(entreeJour?.actif || platJour?.actif || dessertJour?.actif);
        const allActif = !!(entreeJour?.actif && platJour?.actif && dessertJour?.actif);

        async function saveField(item, setter, field, value) {
          setter(p => p ? ({ ...p, [field]: value }) : p);
          await supabase.from('menu_plat_jour').update({ [field]: value, updated_at: new Date().toISOString() }).eq('id', item.id);
        }

        async function toggleAll() {
          // Si au moins un est actif → tout désactiver. Sinon → tout activer.
          const v = !anyActif;
          const items = [
            { it: entreeJour,  s: setEntreeJour },
            { it: platJour,    s: setPlatJour },
            { it: dessertJour, s: setDessertJour },
          ].filter(x => x.it);
          items.forEach(({ it, s }) => s(p => ({...p, actif:v})));
          await Promise.all(items.map(({ it }) => supabase.from('menu_plat_jour').update({ actif:v, updated_at:new Date().toISOString() }).eq('id', it.id)));
        }

        async function toggleItem(item, setter) {
          const v = !item.actif;
          setter(p => ({...p, actif:v}));
          await supabase.from('menu_plat_jour').update({ actif:v, updated_at:new Date().toISOString() }).eq('id', item.id);
        }

        const rows = [
          { item: entreeJour,  setter: setEntreeJour,  label: 'Entrée',  icon: '🥗', ph: "de l'entrée" },
          { item: platJour,    setter: setPlatJour,    label: 'Plat',    icon: '🍽', iconEl: <UtensilsCrossed size={11} style={{display:'inline',verticalAlign:'middle'}} />, ph: 'du plat' },
          { item: dessertJour, setter: setDessertJour, label: 'Dessert', icon: '🍮', iconEl: <UtensilsCrossed size={11} style={{display:'inline',verticalAlign:'middle'}} />, ph: 'du dessert' },
        ];

        const inpStyle = { height:34, border:'1px solid #eee', borderRadius:8, padding:'0 10px', fontSize:13, outline:'none', background:'#fafafa', width:'100%', transition:'border-color 0.15s' };
        const descStyle = { ...inpStyle, height:30, fontSize:12, color:'#666' };
        const prixStyle = { ...inpStyle, width:84, flexShrink:0, textAlign:'right' };

        return (
          <div style={{ display:'flex', flexDirection:'column', gap:10, marginBottom:24 }}>

            {/* Encart 1 — Menu du jour */}
            <div style={{ background:'#fff', borderRadius:16, border:`2px solid ${anyActif ? '#E8C547' : '#eee'}`, overflow:'hidden', transition:'border-color 0.2s' }}>
              {/* Header global */}
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'11px 16px', borderBottom:'1px solid #f5f5f5', background: anyActif ? '#fffdf0' : '#fff' }}>
                <span style={{ fontSize:12, fontWeight:800, color:'#111', textTransform:'uppercase', letterSpacing:1, display:'flex', alignItems:'center', gap:4 }}><UtensilsCrossed size={12} /> Menu du jour</span>
                <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                  <span style={{ fontSize:10, color:'#bbb' }}>Tout</span>
                  <MenuToggle value={anyActif} onChange={toggleAll} />
                </div>
              </div>

              {/* Lignes Entrée / Plat / Dessert */}
              <div style={{ padding:'10px 16px', display:'flex', flexDirection:'column', gap:12 }}>
                {rows.map(({ item, setter, label, icon, ph }) => {
                  const active = !!item?.actif;
                  return (
                    <div key={label} style={{ borderRadius:10, border:`1px solid ${active ? '#f0e88a' : '#f0f0f0'}`, padding:'10px 12px', background: active ? '#fffef5' : '#fafafa', transition:'all 0.15s' }}>
                      {/* Ligne nom + prix + toggle */}
                      <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:7 }}>
                        <span style={{ fontSize:11, fontWeight:700, color: active ? '#888' : '#ccc', width:48, flexShrink:0 }}>{typeof icon === 'string' && (icon === '🍽' || icon === '🍮') ? <UtensilsCrossed size={11} style={{display:'inline',verticalAlign:'middle',marginRight:2}} /> : icon} {label}</span>
                        <input
                          value={item?.nom || ''}
                          onChange={e => setter(p => p ? ({...p, nom: e.target.value}) : p)}
                          onBlur={e => item && saveField(item, setter, 'nom', e.target.value)}
                          placeholder={`Nom ${ph}…`}
                          style={{ ...inpStyle, flex:1, opacity: active ? 1 : 0.5 }}
                          onFocus={e => e.target.style.borderColor='#E8C547'}
                          onBlurCapture={e => e.target.style.borderColor='#eee'}
                        />
                        <div style={{ display:'flex', alignItems:'center', border:'1px solid #eee', borderRadius:8, background:'#fafafa', overflow:'hidden', flexShrink:0, opacity: active ? 1 : 0.5 }}>
                          <input
                            value={item?.prix || ''}
                            onChange={e => setter(p => p ? ({...p, prix: e.target.value}) : p)}
                            onBlur={e => item && saveField(item, setter, 'prix', e.target.value)}
                            placeholder="0"
                            style={{ ...prixStyle, border:'none', borderRadius:0, background:'transparent', width:52 }}
                            onFocus={e => e.target.parentNode.style.borderColor='#E8C547'}
                            onBlurCapture={e => e.target.parentNode.style.borderColor='#eee'}
                          />
                          <span style={{ paddingRight:8, fontSize:13, color:'#aaa', fontWeight:600, userSelect:'none' }}>€</span>
                        </div>
                        <MenuToggle value={active} onChange={() => item && toggleItem(item, setter)} />
                      </div>
                      {/* Description */}
                      <input
                        value={item?.description || ''}
                        onChange={e => setter(p => p ? ({...p, description: e.target.value}) : p)}
                        onBlur={e => item && saveField(item, setter, 'description', e.target.value)}
                        placeholder={`Description courte (optionnel)…`}
                        style={{ ...descStyle, opacity: active ? 1 : 0.4 }}
                        onFocus={e => e.target.style.borderColor='#E8C547'}
                        onBlurCapture={e => e.target.style.borderColor='#eee'}
                      />
                    </div>
                  );
                })}

                {/* Prix formule */}
                <div style={{ borderTop:'1px solid #f0f0f0', paddingTop:10, display:'flex', alignItems:'center', gap:8 }}>
                  <span style={{ fontSize:11, fontWeight:700, color:'#E8C547', flex:1 }}>Prix formule</span>
                  <input
                    value={formuleJour?.prix || ''}
                    onChange={e => setFormuleJour(p => p ? ({...p, prix: e.target.value}) : p)}
                    onBlur={e => formuleJour && saveField(formuleJour, setFormuleJour, 'prix', e.target.value)}
                    placeholder="ex: 22 €"
                    style={{ ...prixStyle }}
                    onFocus={e => e.target.style.borderColor='#E8C547'}
                    onBlurCapture={e => e.target.style.borderColor='#eee'}
                  />
                </div>
              </div>
            </div>

            {/* Encart 2 — Suggestion du chef (compact) */}
            <div style={{ background:'#fff', borderRadius:12, border:`1.5px solid ${suggestionJour?.actif ? '#E8C547' : '#eee'}`, overflow:'hidden', transition:'border-color 0.2s' }}>
              <div style={{ display:'flex', alignItems:'center', gap:10, padding:'8px 14px' }}>
                <span style={{ fontSize:11, fontWeight:700, color:'#888', textTransform:'uppercase', letterSpacing:1, whiteSpace:'nowrap', display:'flex', alignItems:'center', gap:4 }}><UtensilsCrossed size={11} /> Suggestion du chef</span>
                <input
                  value={suggestionJour?.nom || ''}
                  onChange={e => setSuggestionJour(p => p ? ({...p, nom: e.target.value}) : p)}
                  onBlur={e => suggestionJour && saveField(suggestionJour, setSuggestionJour, 'nom', e.target.value)}
                  placeholder="Nom…"
                  style={{ ...inpStyle, flex:1 }}
                  onFocus={e => e.target.style.borderColor='#E8C547'}
                  onBlurCapture={e => e.target.style.borderColor='#eee'}
                />
                <input
                  value={suggestionJour?.description || ''}
                  onChange={e => setSuggestionJour(p => p ? ({...p, description: e.target.value}) : p)}
                  onBlur={e => suggestionJour && saveField(suggestionJour, setSuggestionJour, 'description', e.target.value)}
                  placeholder="Description…"
                  style={{ ...inpStyle, flex:1 }}
                  onFocus={e => e.target.style.borderColor='#E8C547'}
                  onBlurCapture={e => e.target.style.borderColor='#eee'}
                />
                <input
                  value={suggestionJour?.prix || ''}
                  onChange={e => setSuggestionJour(p => p ? ({...p, prix: e.target.value}) : p)}
                  onBlur={e => suggestionJour && saveField(suggestionJour, setSuggestionJour, 'prix', e.target.value)}
                  placeholder="Prix"
                  style={{ ...prixStyle }}
                  onFocus={e => e.target.style.borderColor='#E8C547'}
                  onBlurCapture={e => e.target.style.borderColor='#eee'}
                />
                <MenuToggle value={!!suggestionJour?.actif} onChange={async () => {
                  const v = !suggestionJour?.actif;
                  setSuggestionJour(p => p ? ({...p, actif:v}) : p);
                  if (suggestionJour) await supabase.from('menu_plat_jour').update({ actif:v, updated_at:new Date().toISOString() }).eq('id', suggestionJour.id);
                }} />
              </div>
            </div>

          </div>
        );
      })()}

      {/* Recherche */}
      <div style={{ position:'relative', marginBottom:20 }}>
        <span style={{ position:'absolute', left:13, top:'50%', transform:'translateY(-50%)', color:'#bbb', pointerEvents:'none' }}><Search size={14} /></span>
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

      </div>{/* fin zone défilante */}

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
          produits={produits}
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

      {showOriginesSheet && (
        <OriginesSheet onClose={() => setShowOriginesSheet(false)} showToast={showToast} />
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

            {/* Flyer upload */}
            <div>
              <label style={lbl}>Flyer de la soirée <span style={{fontWeight:400,textTransform:'none',color:'#aaa'}}>(optionnel)</span></label>
              {soireeSheet.image_url ? (
                <div style={{ display:'flex', alignItems:'center', gap:12, marginTop:8 }}>
                  <img src={soireeSheet.image_url} alt="flyer" style={{ width:80, height:80, borderRadius:10, objectFit:'cover', border:'1.5px solid #eee' }} />
                  <button onClick={() => setSoireeSheet(p=>({...p, image_url:''}))} style={{ ...btnSecondary, fontSize:12, height:32, padding:'0 12px', color:'#e57373', borderColor:'#fca5a5' }}>Supprimer</button>
                </div>
              ) : (
                <label style={{ display:'inline-flex', alignItems:'center', gap:8, marginTop:8, height:36, padding:'0 14px', borderRadius:8, border:'1.5px solid #eee', background:'#fff', fontSize:13, fontWeight:600, color:'#666', cursor: uploadingFlyer ? 'wait' : 'pointer' }}>
                  <input type="file" accept="image/jpeg,image/png,image/webp" style={{ display:'none' }} onChange={async e => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    const url = await uploadFlyer(file);
                    if (url) setSoireeSheet(p=>({...p, image_url: url}));
                  }} />
                  {uploadingFlyer ? 'Upload...' : <><ImageIcon size={14} style={{display:'inline',verticalAlign:'middle',marginRight:4}} />Choisir une image</>}
                </label>
              )}
            </div>

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
    </div>
  );
}

// ─── Page Système (backups, restauration) ────────────────────────────────────

const TABLES_BACKUP = ['clients','reservations','roue_gains','roue_recompenses','roue_config','parametres','menu_produits','menu_categories','menu_cartes','menu_soirees','menu_plat_jour','menu_origines'];

// Réglages durables, par opposition au « Statut du jour » qui ne vaut que pour
// la journée en cours.
function ParametresPage({ showToast }) {
  const [conf, setConf] = useState(null);

  async function charger() {
    const { data } = await safeQuery(
      () => supabase.from('commandes_config').select('cle,valeur'),
      { fallback: [], context: 'parametresConfig' }
    );
    const c = {};
    (data || []).forEach(r => { c[r.cle] = r.valeur; });
    setConf(c);
  }
  useEffect(() => { charger(); }, []);

  async function maj(cle, valeur) {
    setConf(c => ({ ...c, [cle]: String(valeur) }));
    const { error } = await safeQuery(
      () => supabase.from('commandes_config').upsert({ cle, valeur: String(valeur), updated_at: new Date().toISOString() }, { onConflict: 'cle' }),
      { fallback: null, context: 'majParametre' }
    );
    if (error) { showToast('Enregistrement impossible', 'error'); charger(); }
    else showToast('✅ Réglage enregistré');
  }

  const autoDefaut = conf ? conf.acceptation_auto_defaut !== 'false' : true;
  const delai = conf ? (parseInt(conf.delai_minutes) || 30) : 30;
  const horizon = conf ? (parseInt(conf.horizon_jours) || 15) : 15;

  const Bloc = ({ titre, aide, children }) => (
    <div style={{ background:'#fff', borderRadius:16, padding:'20px 24px', marginBottom:16 }}>
      <h3 style={{ margin:0, fontSize:15, fontWeight:800, color:'#111' }}>{titre}</h3>
      {aide && <p style={{ margin:'4px 0 16px', fontSize:12.5, color:'#999', lineHeight:1.6 }}>{aide}</p>}
      {children}
    </div>
  );

  return (
    <div style={{ maxWidth:900, margin:'0 auto', padding:'28px 32px 40px' }}>
      <h1 style={{ margin:0, fontSize:26, fontWeight:900, color:'#111', display:'flex', alignItems:'center', gap:10 }}>
        <Settings size={24} strokeWidth={1.9} /> Paramètres
      </h1>
      <p style={{ color:'#888', fontSize:14, margin:'6px 0 24px' }}>
        Réglages durables du CRM. Pour couper la prise de commande sur une seule journée,
        passez plutôt par « Statut du jour » dans le Click and Collect.
      </p>

      <Bloc titre="Acceptation automatique des commandes"
        aide="Quand elle est active, une commande en ligne est acceptée seule et le client reçoit aussitôt son délai. Sinon, chaque commande attend une validation à la main dans « Nouvelles commandes à traiter ».">
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', gap:16 }}>
          <span style={{ fontSize:14, fontWeight:700, color: autoDefaut ? '#15803d' : '#b91c1c' }}>
            {autoDefaut ? 'Active en permanence' : 'Désactivée — validation manuelle'}
          </span>
          <button onClick={()=>maj('acceptation_auto_defaut', autoDefaut ? 'false' : 'true')} disabled={!conf}
            style={{ height:42, padding:'0 18px', borderRadius:11, cursor: conf ? 'pointer' : 'wait', fontSize:13.5, fontWeight:800, whiteSpace:'nowrap',
              border: autoDefaut ? '1.5px solid #fca5a5' : 'none',
              background: autoDefaut ? '#fff' : '#16a34a',
              color: autoDefaut ? '#b91c1c' : '#fff' }}>
            {autoDefaut ? 'Désactiver' : 'Réactiver'}
          </button>
        </div>
      </Bloc>

      <Bloc titre="Délai annoncé au client"
        aide="Temps de préparation communiqué quand une commande est acceptée automatiquement.">
        <div style={{ display:'grid', gridTemplateColumns:'repeat(4, 1fr)', gap:8 }}>
          {DELAIS_RAPIDES.map(mn => (
            <button key={mn} onClick={()=>maj('delai_minutes', mn)} disabled={!conf}
              style={{ height:48, borderRadius:11, border: delai===mn ? '2px solid #16a34a' : '1.5px solid #ddd', background: delai===mn ? '#16a34a' : '#fff', color: delai===mn ? '#fff' : '#333', fontSize:14.5, fontWeight:800, cursor: conf ? 'pointer' : 'wait' }}>
              {fmtDelai(mn)}
            </button>
          ))}
        </div>
      </Bloc>

      <Bloc titre="Commande à l'avance"
        aide="Nombre de jours pendant lesquels un client peut réserver un retrait à partir d'aujourd'hui.">
        <div style={{ display:'flex', alignItems:'center', gap:10 }}>
          <button onClick={()=>maj('horizon_jours', Math.max(1, horizon - 1))} disabled={!conf}
            style={{ width:52, height:46, borderRadius:11, border:'1.5px solid #ddd', background:'#fff', fontSize:22, fontWeight:700, cursor: conf ? 'pointer' : 'wait', color:'#111', lineHeight:1 }}>−</button>
          <div style={{ flex:1, height:46, border:'1.5px solid #ddd', borderRadius:11, display:'flex', alignItems:'center', justifyContent:'center', fontSize:16, fontWeight:800, background:'#fafafa' }}>
            {horizon} jour{horizon > 1 ? 's' : ''}
          </div>
          <button onClick={()=>maj('horizon_jours', Math.min(90, horizon + 1))} disabled={!conf}
            style={{ width:52, height:46, borderRadius:11, border:'1.5px solid #ddd', background:'#fff', fontSize:22, fontWeight:700, cursor: conf ? 'pointer' : 'wait', color:'#111', lineHeight:1 }}>+</button>
        </div>
      </Bloc>

      {/* Sauvegardes : la page existait déjà, elle n'était simplement plus atteignable */}
      <SystemePage showToast={showToast} />
    </div>
  );
}

function SystemePage({ showToast }) {
  const [backups, setBackups] = useState([]);
  const [loadingBk, setLoadingBk] = useState(true);
  const [running, setRunning] = useState(false);
  const [kvError, setKvError] = useState(null);
  const [restoreTarget, setRestoreTarget] = useState(null);
  const [restoreTable, setRestoreTable] = useState('clients');
  const [confirmText, setConfirmText] = useState('');
  const [restoring, setRestoring] = useState(false);
  const [health, setHealth] = useState(null);
  const [errLogs, setErrLogs] = useState([]);
  const [testReport, setTestReport] = useState(null);
  const [testing, setTesting] = useState(false);

  async function runTests() {
    if (testing) return;
    setTesting(true);
    setTestReport(null);
    try {
      const res = await fetch('/api/run-tests', { headers: await authHeaders() });
      setTestReport(await res.json());
    } catch (e) {
      logError(e.message, 'systeme:runTests');
      setTestReport({ ok: false, tests: [], error: 'Suite de tests injoignable' });
    }
    setTesting(false);
  }

  useEffect(() => {
    fetch('/api/health').then(r => r.json()).then(setHealth).catch(() => setHealth({ status: 'down', components: {} }));
    safeQuery(() => supabase.from('error_logs').select('*').order('created_at', { ascending: false }).limit(50), { fallback: [], context: 'systeme:errorLogs' })
      .then(({ data }) => setErrLogs(data || []));
  }, []);

  async function authHeaders() {
    const { data: { session } } = await supabase.auth.getSession();
    return { 'Content-Type': 'application/json', Authorization: `Bearer ${session?.access_token || ''}` };
  }

  async function loadBackups() {
    setLoadingBk(true);
    try {
      const res = await fetch('/api/backups', { headers: await authHeaders() });
      const data = await res.json();
      setBackups(data.backups || []);
      setKvError(data.error || null);
    } catch (e) {
      logError(e.message, 'systeme:loadBackups');
      setKvError('Liste des backups inaccessible');
    }
    setLoadingBk(false);
  }
  useEffect(() => { loadBackups(); }, []);

  async function backupNow() {
    setRunning(true);
    try {
      const res = await fetch('/api/backup-daily', { method: 'POST', headers: await authHeaders() });
      const data = await res.json();
      if (data.ok) { showToast('✅ Backup effectué et vérifié'); loadBackups(); }
      else showToast(data.error || 'Échec du backup', 'error');
    } catch (e) {
      logError(e.message, 'systeme:backupNow');
      showToast('Échec du backup', 'error');
    }
    setRunning(false);
  }

  async function doRestore() {
    if (confirmText !== 'RESTAURER' || restoring) return;
    setRestoring(true);
    try {
      const res = await fetch('/api/backup-restore', {
        method: 'POST',
        headers: await authHeaders(),
        body: JSON.stringify({ date: restoreTarget.date, table: restoreTable, confirm: 'RESTAURER' }),
      });
      const data = await res.json();
      if (data.ok) showToast(`✅ ${data.restored}/${data.total} lignes restaurées dans ${data.table}`);
      else showToast(data.error || 'Échec de la restauration', 'error');
    } catch (e) {
      logError(e.message, 'systeme:restore');
      showToast('Échec de la restauration', 'error');
    }
    setRestoring(false);
    setRestoreTarget(null);
    setConfirmText('');
  }

  const fmtKo = (n) => n ? `${(n / 1024).toFixed(1)} Ko` : '—';
  const totalLignes = (counts) => counts ? Object.values(counts).filter(v => v >= 0).reduce((a, b) => a + b, 0) : null;

  return (
    <div style={{ padding:'32px 40px', maxWidth:980, boxSizing:'border-box' }}>
      <h1 style={{ fontSize:24, fontWeight:900, color:'#111', margin:'0 0 4px', display:'flex', alignItems:'center', gap:10 }}>
        <Settings size={26} strokeWidth={1.8} /> Système
      </h1>
      <p style={{ color:'#888', fontSize:14, margin:'0 0 28px' }}>Sauvegardes automatiques (tous les jours à 2h), restauration et état du CRM.</p>

      {/* ── Santé du système ── */}
      <div style={{ background:'#fff', borderRadius:16, padding:'24px 26px', boxShadow:'0 1px 4px rgba(0,0,0,0.06)', marginBottom:24 }}>
        <h2 style={{ fontSize:16, fontWeight:800, color:'#111', margin:'0 0 16px', display:'flex', alignItems:'center', gap:8 }}>
          <ShieldCheck size={18} /> Santé du système
          {health && (
            <span style={{ marginLeft:'auto', fontSize:12, fontWeight:800, padding:'4px 12px', borderRadius:20, background: health.status==='ok' ? '#dcfce7' : health.status==='degraded' ? '#fef3c7' : '#fee2e2', color: health.status==='ok' ? '#15803d' : health.status==='degraded' ? '#b45309' : '#b91c1c' }}>
              {health.status==='ok' ? 'Tout fonctionne' : health.status==='degraded' ? 'Dégradé' : 'Panne'}
            </span>
          )}
        </h2>
        {!health ? (
          <div style={{ color:'#999', fontSize:14 }}>Vérification…</div>
        ) : (
          <div style={{ display:'flex', flexWrap:'wrap', gap:10 }}>
            {Object.entries(health.components || {}).map(([name, c]) => (
              <div key={name} title={c.detail || ''} style={{ display:'flex', alignItems:'center', gap:8, border:'1.5px solid #f0f0f0', borderRadius:10, padding:'10px 14px', fontSize:13 }}>
                <span style={{ width:9, height:9, borderRadius:'50%', background: c.status==='ok' ? '#22c55e' : c.status==='degraded' ? '#f59e0b' : '#dc2626', boxShadow:`0 0 6px ${c.status==='ok' ? '#22c55e' : c.status==='degraded' ? '#f59e0b' : '#dc2626'}` }} />
                <span style={{ fontWeight:700, color:'#333' }}>{{ env_vars:'Variables', supabase:'Base de données', brevo:'Emails (Brevo)', backups:'Backups' }[name] || name}</span>
                {c.latency_ms != null && <span style={{ color:'#aaa', fontSize:11 }}>{c.latency_ms} ms</span>}
                {c.detail && <span style={{ color:'#b45309', fontSize:11, maxWidth:260, overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{c.detail}</span>}
              </div>
            ))}
          </div>
        )}
      </div>

      <div style={{ background:'#fff', borderRadius:16, padding:'24px 26px', boxShadow:'0 1px 4px rgba(0,0,0,0.06)', marginBottom:24 }}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:18 }}>
          <h2 style={{ fontSize:16, fontWeight:800, color:'#111', margin:0, display:'flex', alignItems:'center', gap:8 }}>
            <History size={18} /> Backups ({backups.length}/30)
          </h2>
          <button onClick={backupNow} disabled={running} style={{ height:40, padding:'0 18px', border:'none', borderRadius:10, background: running ? '#f0f0f0' : '#111', color: running ? '#aaa' : '#fff', fontSize:13, fontWeight:800, cursor: running ? 'wait' : 'pointer', display:'flex', alignItems:'center', gap:8 }}>
            <Save size={15} /> {running ? 'Backup en cours…' : 'Backup maintenant'}
          </button>
        </div>

        {kvError && (
          <div style={{ background:'#fff7ed', border:'1.5px solid #fdba74', borderRadius:10, padding:'12px 16px', fontSize:13, color:'#9a3412', marginBottom:14, display:'flex', alignItems:'center', gap:8 }}>
            <AlertCircle size={16} /> {kvError}
          </div>
        )}

        {loadingBk ? (
          <div style={{ color:'#999', fontSize:14, padding:'16px 0' }}>Chargement…</div>
        ) : backups.length === 0 ? (
          <div style={{ color:'#999', fontSize:14, padding:'16px 0' }}>Aucun backup pour l'instant. Le premier partira automatiquement cette nuit à 2h, ou cliquez sur « Backup maintenant ».</div>
        ) : (
          <table style={{ width:'100%', borderCollapse:'collapse', fontSize:13 }}>
            <thead>
              <tr>
                {['Date','Lignes','Taille','Checksum','Statut',''].map(h => (
                  <th key={h} style={{ textAlign:'left', padding:'8px 10px', color:'#999', fontWeight:700, fontSize:11, letterSpacing:0.5, textTransform:'uppercase', borderBottom:'1.5px solid #f0f0f0' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {backups.map(b => (
                <tr key={b.key}>
                  <td style={{ padding:'10px', fontWeight:700, color:'#111', borderBottom:'1px solid #f7f7f7' }}>{b.date}</td>
                  <td style={{ padding:'10px', color:'#555', borderBottom:'1px solid #f7f7f7' }}>{totalLignes(b.counts) ?? '—'}</td>
                  <td style={{ padding:'10px', color:'#555', borderBottom:'1px solid #f7f7f7' }}>{fmtKo(b.size)}</td>
                  <td style={{ padding:'10px', color:'#aaa', fontFamily:'monospace', fontSize:11, borderBottom:'1px solid #f7f7f7' }}>{b.checksum ? b.checksum.slice(0,12) + '…' : '—'}</td>
                  <td style={{ padding:'10px', borderBottom:'1px solid #f7f7f7' }}>
                    {b.errors ? (
                      <span style={{ display:'inline-flex', alignItems:'center', gap:5, color:'#c2410c', fontWeight:700 }}><AlertCircle size={13} /> {b.errors} erreur(s)</span>
                    ) : (
                      <span style={{ display:'inline-flex', alignItems:'center', gap:5, color:'#15803d', fontWeight:700 }}><CheckCircle size={13} /> Complet</span>
                    )}
                  </td>
                  <td style={{ padding:'10px', textAlign:'right', borderBottom:'1px solid #f7f7f7' }}>
                    <button onClick={() => { setRestoreTarget(b); setConfirmText(''); }} style={{ height:32, padding:'0 12px', border:'1.5px solid #ddd', borderRadius:8, background:'#fff', fontSize:12, fontWeight:700, cursor:'pointer', color:'#333', display:'inline-flex', alignItems:'center', gap:6 }}>
                      <RotateCcw size={13} /> Restaurer
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* ── Suite de tests ── */}
      <div style={{ background:'#fff', borderRadius:16, padding:'24px 26px', boxShadow:'0 1px 4px rgba(0,0,0,0.06)', marginBottom:24 }}>
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom: testReport ? 16 : 0 }}>
          <h2 style={{ fontSize:16, fontWeight:800, color:'#111', margin:0, display:'flex', alignItems:'center', gap:8 }}>
            <CircleCheck size={18} /> Tests du système
            {testReport && !testReport.error && (
              <span style={{ fontSize:12, fontWeight:800, padding:'4px 12px', borderRadius:20, background: testReport.ok ? '#dcfce7' : '#fee2e2', color: testReport.ok ? '#15803d' : '#b91c1c' }}>
                {testReport.passed}/{testReport.total} OK
              </span>
            )}
          </h2>
          <button onClick={runTests} disabled={testing} style={{ height:40, padding:'0 18px', border:'1.5px solid #ddd', borderRadius:10, background:'#fff', color: testing ? '#aaa' : '#333', fontSize:13, fontWeight:800, cursor: testing ? 'wait' : 'pointer', display:'flex', alignItems:'center', gap:8 }}>
            <RefreshCw size={15} style={testing ? { animation:'spin 1s linear infinite' } : undefined} /> {testing ? 'Tests en cours…' : 'Lancer les tests'}
          </button>
        </div>
        {testReport && (testReport.error ? (
          <div style={{ color:'#b91c1c', fontSize:13 }}>{testReport.error}</div>
        ) : (
          <div style={{ display:'flex', flexDirection:'column', gap:8 }}>
            {(testReport.tests || []).map((tst, i) => (
              <div key={i} style={{ display:'flex', alignItems:'center', gap:10, fontSize:13, padding:'8px 12px', borderRadius:8, background: tst.ok ? '#f7fdf9' : '#fef5f5', border:`1px solid ${tst.ok ? '#dcfce7' : '#fee2e2'}` }}>
                {tst.ok ? <CheckCircle size={15} color="#15803d" /> : <AlertCircle size={15} color="#b91c1c" />}
                <span style={{ fontWeight:700, color:'#333' }}>{tst.name}</span>
                <span style={{ color:'#888', fontSize:12 }}>{tst.detail}</span>
                {tst.ms != null && <span style={{ marginLeft:'auto', color:'#bbb', fontSize:11 }}>{tst.ms} ms</span>}
              </div>
            ))}
          </div>
        ))}
      </div>

      {/* ── Dernières erreurs ── */}
      <div style={{ background:'#fff', borderRadius:16, padding:'24px 26px', boxShadow:'0 1px 4px rgba(0,0,0,0.06)', marginBottom:24 }}>
        <h2 style={{ fontSize:16, fontWeight:800, color:'#111', margin:'0 0 16px', display:'flex', alignItems:'center', gap:8 }}>
          <AlertCircle size={18} /> Dernières erreurs ({errLogs.length})
        </h2>
        {errLogs.length === 0 ? (
          <div style={{ color:'#999', fontSize:14 }}>Aucune erreur enregistrée. 👌</div>
        ) : (
          <div style={{ maxHeight:320, overflowY:'auto' }}>
            <table style={{ width:'100%', borderCollapse:'collapse', fontSize:12 }}>
              <thead>
                <tr>
                  {['Date','Contexte','Message'].map(h => (
                    <th key={h} style={{ textAlign:'left', padding:'6px 10px', color:'#999', fontWeight:700, fontSize:11, letterSpacing:0.5, textTransform:'uppercase', borderBottom:'1.5px solid #f0f0f0', position:'sticky', top:0, background:'#fff' }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {errLogs.map(e => (
                  <tr key={e.id}>
                    <td style={{ padding:'8px 10px', color:'#888', whiteSpace:'nowrap', borderBottom:'1px solid #f7f7f7' }}>{new Date(e.created_at).toLocaleString('fr-FR')}</td>
                    <td style={{ padding:'8px 10px', color:'#333', fontWeight:600, borderBottom:'1px solid #f7f7f7' }}>{e.context || '—'}</td>
                    <td style={{ padding:'8px 10px', color:'#666', borderBottom:'1px solid #f7f7f7', wordBreak:'break-word' }}>{e.error_message}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {restoreTarget && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.5)', zIndex:9000, display:'flex', alignItems:'center', justifyContent:'center' }} onMouseDown={() => { if (!restoring) { setRestoreTarget(null); setConfirmText(''); } }}>
          <div style={{ background:'#fff', borderRadius:16, padding:'28px 26px', maxWidth:420, width:'92%', boxShadow:'0 20px 60px rgba(0,0,0,0.25)' }} onMouseDown={e => e.stopPropagation()}>
            <h3 style={{ margin:'0 0 6px', fontSize:17, fontWeight:800, display:'flex', alignItems:'center', gap:8 }}><RotateCcw size={18} /> Restaurer le backup du {restoreTarget.date}</h3>
            <p style={{ color:'#888', fontSize:13, lineHeight:1.6, margin:'0 0 16px' }}>
              Les lignes du backup seront réécrites dans la table choisie (fusion par identifiant).
              Les données créées depuis ne seront pas supprimées.
            </p>
            <label style={{ fontSize:12, fontWeight:700, color:'#666', display:'block', marginBottom:6 }}>Table à restaurer</label>
            <select value={restoreTable} onChange={e => setRestoreTable(e.target.value)} style={{ width:'100%', height:42, border:'1.5px solid #ddd', borderRadius:10, padding:'0 10px', fontSize:14, background:'#fff', marginBottom:16, boxSizing:'border-box' }}>
              {TABLES_BACKUP.map(t => <option key={t} value={t}>{t}{restoreTarget.counts && restoreTarget.counts[t] >= 0 ? ` (${restoreTarget.counts[t]} lignes)` : ''}</option>)}
            </select>
            <label style={{ fontSize:12, fontWeight:700, color:'#666', display:'block', marginBottom:6 }}>Tapez RESTAURER pour confirmer</label>
            <input value={confirmText} onChange={e => setConfirmText(e.target.value)} placeholder="RESTAURER" style={{ width:'100%', height:42, border:'1.5px solid #ddd', borderRadius:10, padding:'0 12px', fontSize:14, boxSizing:'border-box', marginBottom:18, outline:'none' }} />
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={() => { setRestoreTarget(null); setConfirmText(''); }} disabled={restoring} style={{ flex:1, height:44, border:'1.5px solid #ddd', borderRadius:10, background:'#fff', fontSize:14, cursor:'pointer', color:'#666' }}>Annuler</button>
              <button onClick={doRestore} disabled={confirmText !== 'RESTAURER' || restoring} style={{ flex:1, height:44, border:'none', borderRadius:10, background: confirmText === 'RESTAURER' && !restoring ? '#dc2626' : '#f0f0f0', fontSize:14, fontWeight:800, cursor: confirmText === 'RESTAURER' && !restoring ? 'pointer' : 'not-allowed', color: confirmText === 'RESTAURER' && !restoring ? '#fff' : '#bbb' }}>
                {restoring ? 'Restauration…' : 'Restaurer'}
              </button>
            </div>
          </div>
        </div>
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
  const [filtreServiceClients, setFiltreServiceClients] = useState('Tous');
  const [filtreSourceClients, setFiltreSourceClients] = useState('Tous');
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
  const [commandesData, setCommandesData] = useState([]);
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
  const [cmdNouvellesCount, setCmdNouvellesCount] = useState(0);
  const [showPlusSheet, setShowPlusSheet] = useState(false);
  const [mobileTab, setMobileTab] = useState(window.innerWidth < 768 ? 'reservations' : 'clients'); // 'clients' | 'reservations'
  const [showAddResa, setShowAddResa] = useState(false);
  const [activeView, setActiveView] = useState('reservations'); // 'reservations' | 'clients' | 'communications' | 'roue' | 'menu' | 'systeme'
  const [healthStatus, setHealthStatus] = useState(null); // null | 'ok' | 'degraded' | 'down'
  const [healthDetail, setHealthDetail] = useState('');

  useEffect(() => {
    let stop = false;
    const check = () => {
      fetch('/api/health')
        .then(r => r.json())
        .then(d => {
          if (stop) return;
          setHealthStatus(d.status || 'down');
          setHealthDetail(Object.entries(d.components || {}).filter(([, c]) => c.status !== 'ok').map(([n, c]) => `${n}: ${c.detail || c.status}`).join('\n') || 'Tous les composants fonctionnent');
        })
        .catch(() => { if (!stop) { setHealthStatus('down'); setHealthDetail('API injoignable'); } });
    };
    check();
    const iv = setInterval(check, 180000);
    return () => { stop = true; clearInterval(iv); };
  }, []);
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
    if (!user) return;
    loadCmdCount();
    return resilientChannel(supabase, 'sidebar-commandes', (chan) => chan
      .on('postgres_changes', { event: '*', schema: 'public', table: 'commandes' }, () => loadCmdCount())
    );
  }, [user]);

  useEffect(() => {
    return resilientChannel(supabase, 'nouvelles-reservations', (chan) => chan
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'reservations', filter: 'statut=eq.attente' }, async (payload) => {
        if (notifEnCoursRef.current) return;
        notifEnCoursRef.current = true;
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
          headers: await crmAuthHeaders(),
          body: JSON.stringify({
            title: 'Nouvelle réservation !',
            body: `${nom} · ${date} · ${payload.new.heure || ''} · ${payload.new.nb_personnes} pers.`
          })
        }).catch((e) => logError(e.message, 'send-push-onesignal'));
        notifEnCoursRef.current = false;
      })
    );
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

  // Commandes à traiter : compté dans le shell pour que la pastille de la
  // barre latérale reste juste, quel que soit l'onglet ouvert.
  async function loadCmdCount() {
    const { count } = await safeQuery(
      () => supabase.from('commandes').select('id', { count:'exact', head:true }).eq('statut','nouvelle'),
      { fallback: { count: 0 }, context: 'sidebar:commandesNouvelles' }
    );
    setCmdNouvellesCount(count || 0);
  }

  async function loadClients(silent = false) {
    if (!silent) setLoading(true);
    const { data, error } = await safeQuery(() => supabase.from("clients").select("*").is("deleted_at", null).order("created_at", { ascending: false }), { fallback: [], context: 'loadClients' });
    if (error) { showToast("Erreur de chargement", "error"); }
    else { setClients(data || []); }
    if (!silent) setLoading(false);
    chargerToutesStatsClients();
  }

  async function chargerToutesStatsClients() {
    const { data } = await safeQuery(() => supabase.from('reservations').select('client_id, statut, date, service'), { fallback: [], context: 'statsClients' });
    setResasData(data || []);
    // Le volet Click and Collect des fiches se sert de ces commandes.
    const { data: cmds } = await safeQuery(
      () => supabase.from('commandes').select('id,client_id,client_tel,statut,total,items,date_retrait,heure_retrait,created_at').limit(3000),
      { fallback: [], context: 'commandesClients' }
    );
    setCommandesData(cmds || []);
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
        { id:'commandes', label:'Click and Collect', icon:<ShoppingBag size={24} strokeWidth={1.8} /> },
        { id:'clients', label:'Clients', icon:<Users size={24} strokeWidth={1.8} /> },
        { id:'communications', label:'Communications', icon:<Megaphone size={24} strokeWidth={1.8} /> },
        // Onglets masqués — décommenter pour les réafficher (les pages existent toujours)
        // { id:'roue', label:'Jeux', icon:<Dices size={24} strokeWidth={1.8} /> },
        { id:'menu', label:'Menu', icon:<UtensilsCrossed size={24} strokeWidth={1.8} /> },
        // { id:'systeme', label:'Système', icon:<Settings size={24} strokeWidth={1.8} /> },
      ].map(item => {
        const nbAttenteSidebar = item.id === 'reservations' ? resaAttenteCount
          : item.id === 'commandes' ? cmdNouvellesCount : 0;
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
      {/* Pastille d'état système masquée — décommenter pour la réafficher */}
      {false && healthStatus && (
        <div title={healthDetail} style={{ display:'flex', flexDirection:'column', alignItems:'center', gap:5, padding:'10px 8px', marginBottom:2 }}>
          <span style={{ width:10, height:10, borderRadius:'50%', background: healthStatus==='ok' ? '#22c55e' : healthStatus==='degraded' ? '#f59e0b' : '#dc2626', boxShadow:`0 0 8px ${healthStatus==='ok' ? 'rgba(34,197,94,0.8)' : healthStatus==='degraded' ? 'rgba(245,158,11,0.8)' : 'rgba(220,38,38,0.9)'}` }} />
          <span style={{ fontSize:9, fontWeight:600, color:'#555' }}>{healthStatus==='ok' ? 'Système OK' : healthStatus==='degraded' ? 'Dégradé' : 'Panne'}</span>
        </div>
      )}
      <button onClick={()=>setActiveView('parametres')} style={{ width:'100%', padding:'12px 8px', border:'none', display:'flex', flexDirection:'column', alignItems:'center', gap:6, cursor:'pointer', marginBottom:4,
        borderLeft: activeView==='parametres' ? '3px solid #E8C547' : '3px solid transparent',
        background: activeView==='parametres' ? 'rgba(232,197,71,0.1)' : 'transparent',
        color: activeView==='parametres' ? '#E8C547' : '#555' }}>
        <Settings size={22} strokeWidth={1.8} />
        <span style={{ fontSize:10, fontWeight:600 }}>Paramètres</span>
      </button>
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
          margin:'0 auto 20px'
        }}>
          <Bell size={28} />
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
          <Bell size={16} style={{display:'inline',verticalAlign:'middle',marginRight:6}} /> Activer les notifications
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


  if (!isMobile && activeView === 'commandes') return (
    <>
      {sidebarDesktop}
      <div style={{ marginLeft:120, minHeight:'100vh', background:'#f5f5f5' }}>
        <CommandesPage showToast={showToast} user={user} />
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

  if (!isMobile && activeView === 'roue') return (
    <>
      {sidebarDesktop}
      <div style={{ marginLeft:120, minHeight:'100vh', background:'#f5f5f5', overflowY:'auto', boxSizing:'border-box' }}>
        <RouePage showToast={showToast} />
      </div>
    </>
  );

  if (!isMobile && activeView === 'parametres') return (
    <>
      {sidebarDesktop}
      <div style={{ marginLeft:120, minHeight:'100vh', background:'#f5f5f5', overflowY:'auto', boxSizing:'border-box' }}>
        <ParametresPage showToast={showToast} />
      </div>
      {toast && <Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)} />}
    </>
  );

  if (!isMobile && activeView === 'systeme') return (
    <>
      {sidebarDesktop}
      <div style={{ marginLeft:120, minHeight:'100vh', background:'#f5f5f5', overflowY:'auto', boxSizing:'border-box' }}>
        <SystemePage showToast={showToast} />
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
          const res = await fetch('/send-email', { method:'POST', headers: await crmAuthHeaders(), body:JSON.stringify({ to:client.mail, toName:`${client.prenom||''} ${client.nom||''}`.trim(), subject:commObjet, html:buildHtml(client) }) });
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
      const smsHeaders = await crmAuthHeaders();
      for (const id of destinatairesMobiles) {
        const client = clients.find(c => c.id === id);
        if (!client?.tel) { errors++; continue; }
        const msg = smsMessage
          .replace(/{prenom}/g, client.prenom || client.entreprise || '')
          .replace(/{nom}/g, client.nom || '')
          .replace(/{tel}/g, client.tel || '')
          .replace(/{entreprise}/g, client.entreprise || '')
          .replace(/{lien_resa}/g, 'https://ted-crm.pages.dev/reserver.html');
        try {
          // Envoi via la function protégée (la clé Brevo ne quitte jamais le serveur)
          const res = await fetch('/send-sms', {
            method: 'POST',
            headers: smsHeaders,
            body: JSON.stringify({ to: client.tel, message: msg, type: 'marketing' })
          });
          const data = await res.json().catch(() => ({}));
          if (!res.ok || !data.success) { logError(`SMS ${res.status}: ${data.error || ''}`, 'doSendSms'); errors++; }
          else { success++; }
        } catch(err) { errors++; logError(err.message, 'doSendSms'); }
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
                        {filtreServices?.size>0 ? [...filtreServices].map(s=>s==='midi'?'Midi':'Soir').join(', ') : 'Tous les services'}
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
                          {[{id:'midi',label:<><Sun size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Midi</>},{id:'soir',label:<><Moon size={13} style={{display:'inline',verticalAlign:'middle',marginRight:4}}/> Soir</>}].map(s=>{
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
                      <p style={{fontSize:16,fontWeight:800,color:'#111',margin:'0 0 2px'}}>{commType==='email'?<><Mail size={16} style={{display:'inline',verticalAlign:'middle',marginRight:4}} />Email</>:<><MessageSquare size={16} style={{display:'inline',verticalAlign:'middle',marginRight:4}} />SMS</>}</p>
                      <p style={{fontSize:12,color:'#666',margin:0}}>{commType==='sms'?`~${(selectedComm.length*0.04).toFixed(2)}€ estimés`:'Envoi gratuit'}</p>
                    </div>
                    <div style={{flex:2,background:'#f9f9f9',borderRadius:12,padding:'14px 16px'}}>
                      <p style={{fontSize:11,fontWeight:700,color:'#999',textTransform:'uppercase',letterSpacing:1,margin:'0 0 8px'}}>Filtres actifs</p>
                      <div style={{display:'flex',flexWrap:'wrap',gap:6}}>
                        {filtreGenresComm.size>0 && [...filtreGenresComm].map(g=><span key={g} style={{background:'#111',color:'#E8C547',borderRadius:20,padding:'3px 10px',fontSize:12,fontWeight:700}}>{g}s</span>)}
                        {filtreAbsentsMois>0 && <span style={{background:'#111',color:'#E8C547',borderRadius:20,padding:'3px 10px',fontSize:12,fontWeight:700}}>Absents {filtreAbsentsMois}m</span>}
                        {filtreJours?.size>0 && [...filtreJours].map(j=><span key={j} style={{background:'#111',color:'#E8C547',borderRadius:20,padding:'3px 10px',fontSize:12,fontWeight:700}}>{j}</span>)}
                        {filtreServices?.size>0 && [...filtreServices].map(s=><span key={s} style={{background:'#111',color:'#E8C547',borderRadius:20,padding:'3px 10px',fontSize:12,fontWeight:700,display:'inline-flex',alignItems:'center',gap:4}}>{s==='midi'?<><Sun size={12}/> Midi</>:<><Moon size={12}/> Soir</>}</span>)}
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
          <div style={{ width:42, height:42, borderRadius:10, background:'#E8C547', display:'flex', alignItems:'center', justifyContent:'center', flexShrink:0 }}><CalendarDays size={20} /></div>
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
                <button onClick={demanderPermissionNotif} style={{ background:'#E8C547', color:'#111', border:'none', borderRadius:8, width:36, height:36, display:'flex', alignItems:'center', justifyContent:'center', cursor:'pointer' }}><Bell size={18} /></button>
              )}
              <button onClick={()=>setShowConfirmDeconnexion(true)} style={{ background:'rgba(255,255,255,0.1)', border:'none', borderRadius:8, width:36, height:36, display:'flex', alignItems:'center', justifyContent:'center', cursor:'pointer', color:'#fff' }}><LockKeyhole size={16} /></button>
            </div>
          </div>
          {/* Onglets + Recherche — uniquement sur l'onglet Clients */}
          {mobileTab === 'clients' && (
            <>
              <div style={{ padding:'8px 12px 6px', background:'#f5f5f5' }}>
                <div style={{ position:'relative', marginBottom:8 }}>
                  <span style={{ position:'absolute', left:12, top:'50%', transform:'translateY(-50%)', color:'#bbb' }}><Search size={14} /></span>
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

      {isMobile && mobileTab === 'menu' && (
        <div style={{ paddingTop:56, paddingBottom:'calc(80px + env(safe-area-inset-bottom, 16px))', background:'#f5f5f5', minHeight:'100vh', overflowX:'hidden' }}>
          <MenuPage showToast={showToast} />
        </div>
      )}

      {/* ═══ MOBILE CARDS ═══ */}
      {isMobile && mobileTab === 'clients' && (
        <div style={{ paddingTop:146, paddingBottom:'calc(90px + env(safe-area-inset-bottom, 16px))', background:'#f5f5f5', minHeight:'100vh' }}>
          {pageClients.length === 0 && (
            <div style={{ textAlign:'center', padding:'4rem 2rem' }}>
              <div style={{ marginBottom:12 }}><Search size={48} color="#ddd" style={{display:'block',margin:'0 auto'}} /></div>
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
          .filter(c => filtreServiceClients==='Tous' || c.service_favori===filtreServiceClients)
          .filter(c => filtreSourceClients==='Tous' || (filtreSourceClients==='Grand Jeux du TED' ? c.source==='Grand Jeux du TED' : filtreSourceClients==='Réservation' ? c.source==='reservation' : c.source==='manuel' || !c.source))
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
        <div style={{ height:'100vh', boxSizing:'border-box', overflow:'hidden', background:'#f5f5f5', display:'flex', flexDirection:'column' }}>

          {/* 1. HEADER */}
          <div style={{padding:'16px 20px 10px', flexShrink:0}}>
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
          <div style={{ zIndex:100, background:'#f5f5f5', padding:'10px 20px 14px', borderBottom:'1px solid #eee', boxShadow:'0 2px 12px rgba(0,0,0,0.06)', flexShrink:0 }}>
            <div style={{display:'flex', alignItems:'center', gap:10, flexWrap:'wrap'}}>
              <div style={{position:'relative', flex:1}}>
                <Search size={16} strokeWidth={2} color="#999" style={{position:'absolute', left:16, top:'50%', transform:'translateY(-50%)', pointerEvents:'none'}}/>
                <input placeholder="Rechercher un client..." value={rechercheClients} onChange={e=>setRechercheClients(e.target.value)}
                  style={{width:'100%', height:36, minWidth:150, border:'1.5px solid #eee', borderRadius:10, padding:'0 16px 0 44px', fontSize:13, outline:'none', background:'#fff', boxSizing:'border-box'}}
                  onFocus={e=>e.target.style.borderColor='#E8C547'} onBlur={e=>e.target.style.borderColor='#eee'}/>
                {rechercheClients && <button onClick={()=>setRechercheClients('')} style={{position:'absolute', right:14, top:'50%', transform:'translateY(-50%)', background:'none', border:'none', cursor:'pointer', color:'#aaa', fontSize:16, padding:0}}>✕</button>}
              </div>
              <div style={{display:'flex', gap:6, flexWrap:'wrap'}}>
                {[{id:'Tous',label:'Tous'},{id:'Homme',label:'Hommes'},{id:'Femme',label:'Femmes'},{id:'Entreprise',label:'Entreprises'}].map(f=>(
                  <button key={f.id} onClick={()=>setFiltreGenreClients(f.id)} style={{height:36, padding:'0 14px', borderRadius:10, cursor:'pointer', fontSize:13, fontWeight:700, border:'none', background: filtreGenreClients===f.id?'#111':'#fff', color: filtreGenreClients===f.id?'#fff':'#666', boxShadow: filtreGenreClients===f.id?'none':'0 1px 4px rgba(0,0,0,0.06)'}}>{f.label}</button>
                ))}
              </div>
              <button onClick={()=>setModalAdd(true)} style={{height:36, padding:'0 16px', borderRadius:10, border:'none', background:'#E8C547', color:'#111', fontSize:13, fontWeight:800, cursor:'pointer', display:'flex', alignItems:'center', gap:8, flexShrink:0, boxShadow:'0 2px 8px rgba(232,197,71,0.3)'}}>
                <Plus size={16} strokeWidth={2}/> Nouveau client
              </button>
            </div>
          </div>

          {/* 3. STATS + LISTE — seule zone qui défile */}
          <div style={{ padding:'16px 20px 20px', flex:1, minHeight:0, overflowY:'auto' }}>

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
            {/* Menu */}
            <button onClick={()=>setMobileTab('menu')} style={{ flex:1, display:'flex', flexDirection:'column', alignItems:'center', gap:5, border:'none', background:'none', cursor:'pointer', color: mobileTab==='menu' ? '#111' : '#aaa', paddingBottom:4 }}>
              <UtensilsCrossed size={28} strokeWidth={1.8} />
              <span style={{ fontSize:12, fontWeight: mobileTab==='menu' ? 700 : 500 }}>Menu</span>
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
              <button onPointerDown={()=>{ setModalEdit(mobileAction); setMobileAction(null); }} style={{ width:'100%', padding:'13px 16px', background:'none', border:'none', borderBottom:'1px solid #f5f5f5', fontSize:14, fontWeight:600, color:'#1d4ed8', cursor:'pointer', textAlign:'left', display:'flex', alignItems:'center', gap:10 }}><Pencil size={14} style={{display:'inline',verticalAlign:'middle'}} /> Modifier</button>
              <button onPointerDown={()=>{ setModalDelete(mobileAction); setMobileAction(null); }} style={{ width:'100%', padding:'13px 16px', background:'none', border:'none', fontSize:14, fontWeight:600, color:'#dc2626', cursor:'pointer', textAlign:'left', display:'flex', alignItems:'center', gap:10 }}><Trash2 size={14} style={{display:'inline',verticalAlign:'middle'}} /> Supprimer</button>
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
              {c.tel && <div style={{ fontSize:14, color:'#333', display:'flex', alignItems:'center', gap:6 }}><Phone size={14} style={{display:'inline',verticalAlign:'middle'}} /> <a href={`tel:${c.tel}`} style={{ color:'#111', textDecoration:'none', fontWeight:600 }}>{c.tel}</a></div>}
              {c.mail && <div style={{ fontSize:13, color:'#3b82f6', display:'flex', alignItems:'center', gap:6 }}><Mail size={14} style={{display:'inline',verticalAlign:'middle'}} /> <a href={`mailto:${c.mail}`} style={{ color:'#3b82f6', textDecoration:'none' }}>{c.mail}</a></div>}
              {c.tel && (
                <div style={{ display:'flex', gap:8, marginTop:4, marginBottom:4 }}>
                  <a href={`sms:${c.tel}`} style={{ flex:1, height:44, background:'#fff', border:'1.5px solid #ddd', borderRadius:10, fontSize:13, fontWeight:700, cursor:'pointer', color:'#111', display:'flex', alignItems:'center', justifyContent:'center', textDecoration:'none', gap:6 }}><MessageSquare size={14} /> SMS</a>
                  <a href={`tel:${c.tel}`} style={{ flex:1, height:44, background:'#111', border:'none', borderRadius:10, fontSize:13, fontWeight:700, cursor:'pointer', color:'#fff', display:'flex', alignItems:'center', justifyContent:'center', textDecoration:'none', gap:6 }}><Phone size={14} /> Appeler</a>
                </div>
              )}
              {c.created_at && <div style={{ fontSize:12, color:'#999', display:'flex', alignItems:'center', gap:4 }}><ClipboardList size={12} style={{display:'inline',verticalAlign:'middle'}} /> Client depuis le {formatDate(c.created_at)}</div>}
              {(c.service_favori || c.jour_favori || (c.source && c.source !== 'manuel')) && (
                <div style={{ display:'flex', flexWrap:'wrap', gap:6, marginTop:2 }}>
                  {c.service_favori && (
                    <span style={{ display:'inline-flex', alignItems:'center', gap:4, fontSize:12, fontWeight:700, padding:'4px 10px', borderRadius:20, background: c.service_favori==='midi' ? '#fffbea' : '#1e1b4b', color: c.service_favori==='midi' ? '#92400e' : '#c7d2fe', border: c.service_favori==='midi' ? '1.5px solid #fde68a' : '1.5px solid #4338ca' }}>
                      {c.service_favori === 'midi' ? <><Clock size={12} style={{display:'inline',verticalAlign:'middle',marginRight:3}} />Midi</> : <><Moon size={12} style={{display:'inline',verticalAlign:'middle',marginRight:3}} />Soir</>}
                    </span>
                  )}
                  {c.jour_favori && (
                    <span style={{ display:'inline-flex', alignItems:'center', gap:4, fontSize:12, fontWeight:700, padding:'4px 10px', borderRadius:20, background:'#f0fdf4', color:'#166534', border:'1.5px solid #bbf7d0' }}>
                      <CalendarDays size={12} style={{display:'inline',verticalAlign:'middle',marginRight:3}} />{c.jour_favori}
                    </span>
                  )}
                  {c.source === 'Grand Jeux du TED' && (
                    <span style={{ display:'inline-flex', alignItems:'center', gap:4, fontSize:12, fontWeight:700, padding:'4px 10px', borderRadius:20, background:'#E8C547', color:'#111', border:'1.5px solid #d4a800' }}>
                      <Dices size={12} style={{display:'inline',verticalAlign:'middle',marginRight:3}} />Grand Jeu du TED
                    </span>
                  )}
                </div>
              )}
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
                      <p style={{ fontSize:11, fontWeight:700, color:'#999', textTransform:'uppercase', letterSpacing:1, margin:0, display:'flex', alignItems:'center', gap:4 }}><Trophy size={11} /> Jours favoris</p>
                      <span style={{ fontSize:10, color:'#bbb' }}>↻ {periodeLabel}</span>
                    </div>
                    {topJoursClient.length === 0
                      ? <p style={{ fontSize:12, color:'#bbb', margin:0 }}>Pas de données sur cette période</p>
                      : topJoursClient.map(([label, count], i) => (
                        <div key={i} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:4 }}>
                          <span style={{ fontSize:13, color:'#444', display:'flex', alignItems:'center', gap:4 }}><Award size={13} color={i===0?'#FFD700':i===1?'#C0C0C0':'#CD7F32'} /> {label}</span>
                          <span style={{ fontSize:13, fontWeight:700, color:'#111' }}>{count} résa</span>
                        </div>
                      ))
                    }
                  </div>
                );
              })()}
              <BlocClickCollect stats={statsClickCollect(commandesData, c)} compact />
            </div>
        );
        const ficheFooter = (
          <div style={{ display:'flex', flexDirection:'column', gap:8, width:'100%' }}>
            {!ficheClientReadOnly && (
              <div style={{ display:'flex', gap:8 }}>
                <button onClick={()=>{ fermerFiche(); setModalDelete(c); }} style={{ flex:1, height:44, border:'1.5px solid #dc2626', borderRadius:10, background:'#fef2f2', color:'#dc2626', fontSize:14, fontWeight:700, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:6 }}><Trash2 size={14} /> Supprimer le client</button>
                <button onClick={()=>{ setModalEdit(c); }} style={{ flex:2, height:44, border:'none', borderRadius:10, background:'#E8C547', color:'#111', fontSize:14, fontWeight:700, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:6 }}><Pencil size={14} /> Modifier le client</button>
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

              <BlocClickCollect stats={statsClickCollect(commandesData, c)} />

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
                              {service==='Midi' ? <Sun size={12} style={{display:'inline',verticalAlign:'middle',marginRight:2}} /> : <Moon size={12} style={{display:'inline',verticalAlign:'middle',marginRight:2}} />} {service} · {count} résa
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
              <h2 style={{margin:0,fontSize:20,fontWeight:800,color:'#111',display:'flex',alignItems:'center',gap:8}}><Trophy size={20} /> Classement clients</h2>
              <button onClick={()=>setShowTopClients(false)} style={{width:36,height:36,borderRadius:'50%',border:'none',background:'#f0f0f0',cursor:'pointer',fontSize:18,color:'#666',display:'flex',alignItems:'center',justifyContent:'center'}}>✕</button>
            </div>
            <div style={{flex:1,overflowY:'auto',padding:'16px 28px'}}>
              {clients
                .map(c=>({...c, nb:resasData.filter(r=>r.client_id===c.id&&r.statut!=='absente'&&r.statut!=='annulee'&&r.statut!=='refusee').length}))
                .filter(c=>c.nb>0)
                .sort((a,b)=>b.nb-a.nb)
                .slice(0,50)
                .map((c,i)=>{
                  const medals=[<Award size={18} color="#FFD700" key="or"/>,<Award size={18} color="#C0C0C0" key="arg"/>,<Award size={18} color="#CD7F32" key="bro"/>];
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
