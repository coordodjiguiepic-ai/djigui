<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CPS Performance DJIGUI — Google Sheets</title>
<style>
body{margin:0;font-family:Arial,sans-serif;background:#f4f7f5;color:#17362b}
.app{display:flex;min-height:100vh}
aside{width:240px;background:#104d39;color:#fff;padding:18px}
.brand{font-weight:bold;font-size:18px;margin-bottom:25px}
nav button{display:block;width:100%;padding:11px;margin:3px 0;border:0;border-radius:7px;background:transparent;color:#fff;text-align:left;cursor:pointer}
nav button.active{background:#23765a}
main{flex:1;padding:24px}
.head{display:flex;justify-content:space-between;margin-bottom:20px;align-items:center;flex-wrap:wrap;gap:10px}
h1{margin:0;font-size:24px}
h2{font-size:17px}
.grid{display:grid;gap:14px}
.kpis{grid-template-columns:repeat(4,1fr)}
.card{background:white;border:1px solid #dfe8e3;border-radius:12px;padding:16px}
.num{font-size:28px;font-weight:bold}
.good{color:#14845a}
.bad{color:#c23b32}
.muted{font-size:12px;color:#718078}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{padding:9px;border-bottom:1px solid #e5ece8;text-align:left}
th{font-size:10px;color:#718078;text-transform:uppercase}
.panel{display:none}
.active{display:block}
.notice{padding:11px;background:#edf7f2;border-left:4px solid #16855d;margin-bottom:15px}
.btn{padding:8px 12px;border:1px solid #d5e0db;background:white;border-radius:7px;cursor:pointer}
.primary{background:#147c56;color:#fff}
.form{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
input,select{padding:8px;border:1px solid #d8e2de;border-radius:7px;box-sizing:border-box}
input.editable-input{width:65px;padding:4px;text-align:center}
.filters-bar{display:flex;gap:8px;flex-wrap:wrap;background:white;padding:12px;border-radius:10px;border:1px solid #dfe8e3;margin-bottom:15px;align-items:center}
.login-box{background:#fff;border:1px solid #dfe8e3;padding:15px;border-radius:10px;margin-bottom:15px;display:flex;gap:15px;align-items:center;flex-wrap:wrap}
#syncStatus{font-size:12px;font-weight:bold;margin-left:10px}
@media(max-width:750px){aside{width:190px}.kpis{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<div class="app">
<aside>
<div class="brand">DJIGUI VIH/SIDA<br><small>Oct 2026 — Sep 2027</small></div>
<nav>
<button class="active" onclick="show('dash',this)">▣ Tableau de bord</button>
<button onclick="show('entry',this)">✎ Saisie CPS & Agents</button>
<button onclick="show('ranking',this)">🏆 Classements</button>
<button onclick="show('admin',this)">⚙ Matrices Admin</button>
</nav>
</aside>
<main>
<div class="head">
  <div>
    <h1 id="title">Tableau de bord</h1>
    <span class="muted">Suivi consolidé & Indicateurs opérationnels (Google Sheets)</span>
    <span id="syncStatus" class="muted"></span>
  </div>
  <div>
    <button class="btn primary" onclick="exportExcel()">📥 Exporter vers Excel</button>
    <button class="btn primary" onclick="loadDataFromSheets()" style="background:#104d39;font-weight:bold;margin-left:10px">🔄 Recharger depuis Google Sheets</button>
  </div>
</div>

<!-- DASHBOARD -->
<section id="dash" class="panel active">
<div class="filters-bar">
  <div>
    <label class="muted">Fréquence d'affichage</label>
    <select id="timeFilterType" onchange="toggleTimeSelectors(); updateAll();" style="font-weight:bold">
      <option value="month">Mensuel (Mois)</option>
      <option value="week">Hebdomadaire (Semaine)</option>
      <option value="quarter">Trimestriel</option>
      <option value="semester">Semestriel</option>
      <option value="year">Annuel (12 Mois)</option>
    </select>
  </div>
  <div>
    <label class="muted">Période</label>
    <select id="periodValueSelect" onchange="updateAll()" style="font-weight:bold;background:#edf7f2;color:#104d39"></select>
  </div>
  <div>
    <label class="muted">Filtrer par CPS</label>
    <select id="siteSelect" onchange="updateAll()" style="font-weight:bold;background:#edf7f2;color:#104d39"></select>
  </div>
</div>

<div class="notice"><b>Visibilité publique :</b> Le tableau reste intact, visible et filtrable pour tous les utilisateurs.</div>

<div class="grid kpis">
  <div class="card"><div class="muted">Taux global</div><div id="kpiGlobalRate" class="num good">0%</div></div>
  <div class="card"><div class="muted">Cible totale</div><div id="kpiTarget" class="num">0</div></div>
  <div class="card"><div class="muted">Réalisation</div><div id="kpiRealized" class="num">0</div></div>
  <div class="card"><div class="muted">GAP restant</div><div id="kpiGap" class="num bad">0</div></div>
</div><br>

<div class="card" style="margin-bottom:20px">
  <h2>Performance CPS (Indicateurs Cliniques/Communautaires)</h2>
  <table>
    <thead><tr><th>Indicateur</th><th>Cible Période</th><th>Réalisé Période</th><th>GAP</th><th>Taux</th></tr></thead>
    <tbody id="dashTableBody"></tbody>
  </table>
</div>

<div class="card">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:10px">
    <h2>Performance Agents de Saisie (Indicateurs Spécifiques)</h2>
    <div>
      <label class="muted">Filtrer Agent de Saisie :</label>
      <select id="dashAgentFilter" onchange="updateAll()" style="font-weight:bold;background:#edf7f2;color:#104d39"></select>
    </div>
  </div>
  <table>
    <thead><tr><th>Indicateur Agent</th><th>Cible Période</th><th>Réalisé Période</th><th>GAP</th><th>Taux</th></tr></thead>
    <tbody id="dashAgentTableBody"></tbody>
  </table>
</div>
</section>

<!-- SAISIE -->
<section id="entry" class="panel">
<div class="login-box">
  <div>
    <label class="muted">Filtre 1 : Rôle & Entité</label>
    <select id="authRole" style="font-weight:bold" onchange="populateAuthDropdowns(); checkAuth();">
      <option value="CPS">CPS</option>
      <option value="ASD">Agent de Saisie (ASD)</option>
    </select>
  </div>
  <div>
    <label class="muted">Filtre 2 : Site (Implantation)</label>
    <select id="authSite" style="font-weight:bold" onchange="populateAuthNames(); checkAuth();"><option value="">-- Choisir un site --</option></select>
  </div>
  <div>
    <label class="muted">Filtre 3 : Nom (Individuel)</label>
    <select id="authAgent" style="font-weight:bold" onchange="checkAuth()"><option value="">-- Choisir un nom --</option></select>
  </div>
  <div id="authStatus" style="font-weight:bold;margin-top:15px" class="bad">Veuillez configurer l'ensemble des filtres d'accès.</div>
</div>

<div class="card" id="saisieCard" style="opacity:.4;pointer-events:none">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;flex-wrap:wrap;gap:10px">
    <h2 id="saisieCardTitle">Saisie</h2>
    <div style="display:flex;align-items:center;gap:10px">
      <div>
        <label class="muted">Mois de saisie :</label>
        <select id="entryMonthSelect" onchange="updateAll(); checkAuth();" style="font-weight:bold"></select>
      </div>
      <div>
        <button class="btn primary" onclick="sauvegardeSauvage()" style="background:#104d39;font-weight:bold;margin-top:16px">💾 Sauvegarder dans Google Sheets</button>
      </div>
    </div>
  </div>
  <div id="lockNotice" class="notice" style="display:none;background:#fdecea;border-color:#c23b32;color:#c23b32"><b>Période verrouillée :</b> L'administrateur a verrouillé ce mois. La modification des données est désactivée.</div>

  <div id="saisieCpsContainer">
    <h3>Indicateurs CPS</h3>
    <table>
      <thead><tr><th>Indicateur</th><th>Cible Mensuelle</th><th>S1</th><th>S2</th><th>S3</th><th>S4</th><th>Total</th><th>Taux</th></tr></thead>
      <tbody id="entryTableBody"></tbody>
    </table>
  </div>

  <div id="saisieAgentContainer" style="display:none">
    <h3>Indicateurs Agents de Saisie (ASD)</h3>
    <table>
      <thead><tr><th>Indicateur Agent</th><th>Cible Mensuelle</th><th>S1</th><th>S2</th><th>S3</th><th>S4</th><th>Total</th><th>Taux</th></tr></thead>
      <tbody id="entryAgentTableBody"></tbody>
    </table>
  </div>
</div>
</section>

<!-- CLASSEMENTS -->
<section id="ranking" class="panel">
<div class="card" style="margin-bottom:20px;background:#fafcfb">
  <h2>Paramètres Globaux du Classement</h2>
  <div class="form" style="margin-top:10px">
    <div>
      <label class="muted">Pourcentage de référence cible (%)</label>
      <input type="number" id="globalRefPercentage" value="80" min="0" max="100" onchange="updateAll()" style="width:80px;font-weight:bold">
    </div>
    <div>
      <label class="muted">Mois de Notation Administrative à afficher</label>
      <select id="rankingMonthSelect" onchange="updateAll()" style="font-weight:bold"></select>
    </div>
  </div>
</div>

<div class="grid" style="gap:20px">
  <div class="card">
    <h2>Classement des CPS</h2>
    <p class="muted">Basé sur le nombre d'indicateurs atteints par rapport au pourcentage de référence, plus les notes administratives de qualité des données et de suivi.</p>
    <table>
      <thead><tr><th>Rang</th><th>CPS</th><th>Indicateurs Atteints</th><th>Qualité Données (-3 à +3)</th><th>Qualité Suivi (-2 à +2)</th><th>Score Total</th></tr></thead>
      <tbody id="rankingCpsBody"></tbody>
    </table>
  </div>
  <div class="card">
    <h2>Classement des Agents de Saisie (ASD)</h2>
    <p class="muted">Basé sur l'atteinte des indicateurs spécifiques par rapport au pourcentage de référence et la notation administrative.</p>
    <table>
      <thead><tr><th>Rang</th><th>Agent (ASD)</th><th>Site Rattachement</th><th>Indicateurs Atteints</th><th>Qualité Données (-3 à +3)</th><th>Qualité Suivi (-2 à +2)</th><th>Score Total</th></tr></thead>
      <tbody id="rankingAgentsBody"></tbody>
    </table>
  </div>
</div>
</section>

<!-- ADMIN -->
<section id="admin" class="panel">
<div id="adminLoginWrapper" class="card" style="max-width:450px;margin:40px auto;text-align:center">
  <h2>🔐 Accès Restreint - Administrateur</h2>
  <p class="muted">Veuillez saisir votre nom de profil et votre mot de passe administrateur.</p>
  <div style="display:flex;flex-direction:column;gap:12px;margin-top:20px">
    <input type="text" id="adminUserField" placeholder="Nom de profil" style="width:100%">
    <input type="password" id="adminPassField" placeholder="Mot de passe administrateur" style="width:100%">
    <button class="btn primary" onclick="verifyAdminAccess()" style="width:100%;padding:10px">Se connecter à l'administration</button>
    <div id="adminLoginError" class="bad" style="font-weight:bold;display:none">Identifiants incorrects. Accès refusé.</div>
  </div>
</div>

<div id="adminContentContainer" style="display:none" class="grid">
  <div class="card">
    <h2>1. Matrice Cibles CPS & Association des Agents</h2>
    <table>
      <thead><tr id="matrixCpsHeader"></tr></thead>
      <tbody id="matrixCpsBody"></tbody>
    </table><br>
    <div class="form">
      <input type="text" id="newCpsName" placeholder="Nom du CPS">
      <input type="text" id="newCpsAgent" placeholder="Agent responsable">
      <button class="btn primary" onclick="addCps()">Ajouter CPS</button>
    </div>
  </div>

  <div class="card">
    <h2>2. Matrice Cibles Agents de Saisie (ASD)</h2>
    <table>
      <thead><tr id="matrixAgentHeader"></tr></thead>
      <tbody id="matrixAgentBody"></tbody>
    </table><br>
    <div class="form">
      <input type="text" id="newAgentNameInput" placeholder="Nom de l'agent">
      <input type="text" id="newAgentSiteInput" placeholder="Site assigné">
      <button class="btn primary" onclick="addAgentEntry()">Ajouter Agent</button>
    </div>
  </div>

  <div class="card">
    <h2>3. Gestion des Indicateurs (Matrice Admin Séparée)</h2>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px">
      <div>
        <h3>Indicateurs Modifiables CPS</h3>
        <table>
          <thead><tr><th>Code</th><th>Libellé</th><th>Action</th></tr></thead>
          <tbody id="adminIndCpsBody"></tbody>
        </table><br>
        <div class="form">
          <input type="text" id="indCpsCode" placeholder="Code">
          <input type="text" id="indCpsLabel" placeholder="Libellé">
          <button class="btn primary" onclick="addIndCps()">Ajouter CPS</button>
        </div>
      </div>
      <div>
        <h3>Indicateurs Modifiables Agents de Saisie (ASD)</h3>
        <table>
          <thead><tr><th>Code</th><th>Libellé</th><th>Action</th></tr></thead>
          <tbody id="adminIndAgentBody"></tbody>
        </table><br>
        <div class="form">
          <input type="text" id="indAgentCode" placeholder="Code">
          <input type="text" id="indAgentLabel" placeholder="Libellé">
          <button class="btn primary" onclick="addIndAgent()">Ajouter ASD</button>
        </div>
      </div>
    </div>
  </div>

  <div class="card">
    <h2>4. Verrouillage des Périodes de Saisie</h2>
    <p class="muted">Sélectionnez les mois à verrouiller pour empêcher toute modification par les agents après la date limite.</p>
    <table>
      <thead><tr><th>Mois</th><th>Statut Actuel</th><th>Action Admin</th></tr></thead>
      <tbody id="adminLockBody"></tbody>
    </table>
  </div>

  <div class="card">
    <h2>5. Notation Administrative (Mensuelle)</h2>
    <div class="form" style="margin-bottom:15px">
      <div>
        <label class="muted">Mois de notation :</label>
        <select id="adminNoteMonthSelect" onchange="renderAdminMatrices()" style="font-weight:bold"></select>
      </div>
    </div>
    <table>
      <thead><tr><th>Entité / Nom</th><th>Type</th><th>Qualité Données (-3 à +3)</th><th>Qualité Suivi (-2 à +2)</th></tr></thead>
      <tbody id="adminNotesBody"></tbody>
    </table>
  </div>

  <div class="card">
    <h2>6. Historique des Modifications (Audit Trail - Cibles)</h2>
    <p class="muted">Trace des modifications de cibles effectuées par l'administrateur dans la matrice.</p>
    <table>
      <thead><tr><th>Date & Heure</th><th>Cible / Élément</th><th>Indicateur</th><th>Ancienne Valeur</th><th>Nouvelle Valeur</th></tr></thead>
      <tbody id="auditTrailBody"></tbody>
    </table>
  </div>
</div>
</section>

<script>
// =====================================================
// 1. CONFIGURATION GOOGLE SHEETS  (à modifier)
// =====================================================
// Collez ici l'URL de l'application web Apps Script (se termine par /exec)
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbynn6vDY05WqG96mdL-bC56EGrGY4Q9xEtVMijXHOzL1wdIlyLcOSycqRBqBT5ky82aog/exec";

// =====================================================
// 2. COMMUNICATION AVEC GOOGLE SHEETS
// =====================================================
let _queue = Promise.resolve();   // les écritures partent une par une, dans l'ordre

function setSync(msg, cls) {
  const el = document.getElementById('syncStatus');
  if (el) { el.textContent = msg; el.className = cls || 'muted'; }
}
function urlOk() { return SCRIPT_URL.indexOf('https://script.google.com/') === 0; }

// rows = [{ sheet:'cps_data', data:{...} }, ...]  -> true si succès
function sheetsSave(rows) {
  _queue = _queue.then(async function () {
    if (!urlOk()) { setSync('⚠ URL Google Sheets non configurée', 'bad'); return false; }
    setSync('⏳ Enregistrement…');
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        // Pas d'en-tête Content-Type => requête "simple", pas de blocage CORS
        const res = await fetch(SCRIPT_URL, { method: 'POST', body: JSON.stringify({ action: 'upsert', rows: rows }) });
        const json = await res.json();
        if (!json.ok) throw new Error(json.error || 'Réponse invalide');
        setSync('✅ Enregistré dans Google Sheets (' + new Date().toLocaleTimeString() + ')', 'good');
        return true;
      } catch (e) {
        console.error('Écriture Sheets (essai ' + attempt + ') :', e);
      }
    }
    setSync('❌ Échec de l\'enregistrement — vérifiez la connexion', 'bad');
    return false;
  });
  return _queue;
}

// =====================================================
// 3. DONNÉES INITIALES
// =====================================================
const ADMIN_CREDENTIALS = [{ name: "coordodjiguiepic", pass: "DJIGUI2001@" }];
let isAdminAuthenticated = false;

let structures = [
  { name: "BOUGOUNI 1", agent: "DRISSA MARIKO" },
  { name: "BOUGOUNI 2", agent: "ROKIA SINYOKO" },
  { name: "GARALO 2", agent: "MODIBO TRAORE" },
  { name: "MANANKORO", agent: "SEYDOU SIDIBE" },
  { name: "KOUMANTOU", agent: "MADY DIAKITE" },
  { name: "SANSO", agent: "Mariam GUINDO" },
  { name: "NIENA", agent: "HAWA KONE" },
  { name: "KOLONDIEBA", agent: "Lobou COULIBALY" },
  { name: "FAKOLA", agent: "Mamoutou TOGOLA" },
  { name: "KADIANA", agent: "ADIARRA TOGOLA" },
  { name: "YANFOLILA", agent: "NGOLO COULIBALY" },
  { name: "KALANA", agent: "BATOGOMA DIARRA" },
  { name: "SELINGUE", agent: "HAROUNA DOUMBIA" },
  { name: "SIEKOROLE", agent: "Baba FALL" }
];

let agentsSaisieList = [
  { name: "MADINE COULIBALY", site: "YANFOLILA" },
  { name: "OUMAR DOUMBIA", site: "BOUGOUNI" },
  { name: "BABA DIARRA", site: "KOLONDIEBA" }
];

let indicatorsCps = [
  { code: "HTS_SELF", label: "Kit distribué (HTS_SELF)" },
  { code: "SELF_REACT", label: "ADVIH réactif" },
  { code: "SELF_CONF", label: "ADVIH confirmé" },
  { code: "SELF_LINKED", label: "ADVIH réactif lié au TARV" },
  { code: "VL_SAMPLE", label: "CV prélevée" },
  { code: "INDEX", label: "index testé" },
  { code: "TX_NEW", label: "Positif mis sous ARV" }
];

let indicatorsAgent = [
  { code: "CV_DOC", label: "Résultats CV documentés" },
  { code: "PDV_RET", label: "Perdus de vue retrouvés" },
  { code: "DIST_COMM", label: "Distribution communautaire" }
];

const monthsList = ["2026-10","2026-11","2026-12","2027-01","2027-02","2027-03","2027-04","2027-05","2027-06","2027-07","2027-08","2027-09"];

let database = { cps: {}, agents: {}, evals: {}, locks: {}, audit: [] };

function ensure(o, k) { if (!o[k]) o[k] = {}; return o[k]; }

// Crée les valeurs par défaut pour tout ce qui n'existe pas encore (peut être rappelée sans danger)
function seedLocal() {
  monthsList.forEach(m => {
    if (database.locks[m] === undefined) database.locks[m] = false;
    structures.forEach(st => {
      const mo = ensure(ensure(database.cps, st.name), m);
      ensure(database.evals, st.name);
      if (!database.evals[st.name][m]) database.evals[st.name][m] = { dataQuality: 0, trackingQuality: 0 };
      indicatorsCps.forEach(ind => { if (!mo[ind.code]) mo[ind.code] = { target: 40, s1: 0, s2: 0, s3: 0, s4: 0 }; });
    });
    agentsSaisieList.forEach(ag => {
      const mo = ensure(ensure(database.agents, ag.name), m);
      ensure(database.evals, ag.name);
      if (!database.evals[ag.name][m]) database.evals[ag.name][m] = { dataQuality: 0, trackingQuality: 0 };
      indicatorsAgent.forEach(ind => { if (!mo[ind.code]) mo[ind.code] = { target: 20, s1: 0, s2: 0, s3: 0, s4: 0 }; });
    });
  });
}

// =====================================================
// 4. CHARGEMENT DEPUIS GOOGLE SHEETS
// =====================================================
async function loadDataFromSheets() {
  if (!urlOk()) {
    setSync('⚠ URL Google Sheets non configurée (voir SCRIPT_URL)', 'bad');
    seedLocal(); updateAll();
    return;
  }
  setSync('⏳ Chargement…');
  try {
    const res = await fetch(SCRIPT_URL + '?action=all&t=' + Date.now());
    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'Réponse invalide');
    const d = json.data;
    const num = x => Number(x) || 0;

    // a) listes (CPS, agents, indicateurs) sauvegardées par l'admin
    (d.config || []).forEach(r => {
      try {
        const v = JSON.parse(r.json);
        if (!Array.isArray(v)) return;
        if (r.key === 'structures') structures = v;
        if (r.key === 'agentsSaisieList') agentsSaisieList = v;
        if (r.key === 'indicatorsCps') indicatorsCps = v;
        if (r.key === 'indicatorsAgent') indicatorsAgent = v;
      } catch (e) { console.warn('config illisible', r.key); }
    });

    // b) valeurs par défaut, puis c) on superpose les données du Sheet
    seedLocal();
    (d.cps_data || []).forEach(r => {
      if (!r.site || !r.month || !r.code) return;
      ensure(ensure(database.cps, r.site), r.month)[r.code] = { target: num(r.target), s1: num(r.s1), s2: num(r.s2), s3: num(r.s3), s4: num(r.s4) };
    });
    (d.agents_data || []).forEach(r => {
      if (!r.agent || !r.month || !r.code) return;
      ensure(ensure(database.agents, r.agent), r.month)[r.code] = { target: num(r.target), s1: num(r.s1), s2: num(r.s2), s3: num(r.s3), s4: num(r.s4) };
    });
    (d.evals || []).forEach(r => {
      if (!r.entity || !r.month) return;
      ensure(database.evals, r.entity)[r.month] = { dataQuality: num(r.dataQuality), trackingQuality: num(r.trackingQuality) };
    });
    (d.locks || []).forEach(r => { if (r.month) database.locks[r.month] = num(r.locked) === 1; });
    database.audit = d.audit || [];

    setSync('✅ Données chargées (' + new Date().toLocaleTimeString() + ')', 'good');
    updateAll();
  } catch (err) {
    console.error('Chargement Sheets :', err);
    setSync('❌ Chargement impossible — vérifiez l\'URL et le déploiement', 'bad');
    seedLocal(); updateAll();
    alert("Impossible de charger les données depuis Google Sheets.\nVérifiez que le déploiement est en accès « Tout le monde » et que l'URL se termine par /exec.");
  }
}

// =====================================================
// 5. ÉCRITURE DES DONNÉES
// =====================================================
function cpsRow(site, month, code) {
  const c = database.cps[site][month][code];
  return { sheet: 'cps_data', data: { site, month, code, target: c.target, s1: c.s1, s2: c.s2, s3: c.s3, s4: c.s4, updatedAt: new Date().toISOString() } };
}
function agentRow(agent, month, code) {
  const c = database.agents[agent][month][code];
  return { sheet: 'agents_data', data: { agent, month, code, target: c.target, s1: c.s1, s2: c.s2, s3: c.s3, s4: c.s4, updatedAt: new Date().toISOString() } };
}
function evalRow(entity, month) {
  const e = database.evals[entity][month];
  return { sheet: 'evals', data: { entity, month, dataQuality: e.dataQuality, trackingQuality: e.trackingQuality, updatedAt: new Date().toISOString() } };
}
function saveConfig() {
  return sheetsSave([
    { sheet: 'config', data: { key: 'structures', json: JSON.stringify(structures) } },
    { sheet: 'config', data: { key: 'agentsSaisieList', json: JSON.stringify(agentsSaisieList) } },
    { sheet: 'config', data: { key: 'indicatorsCps', json: JSON.stringify(indicatorsCps) } },
    { sheet: 'config', data: { key: 'indicatorsAgent', json: JSON.stringify(indicatorsAgent) } }
  ]);
}

function updateCpsVal(site, month, code, wk, val) {
  const c = ensure(ensure(database.cps, site), month);
  if (!c[code]) c[code] = { target: 40, s1: 0, s2: 0, s3: 0, s4: 0 };
  c[code][wk] = parseInt(val) || 0;
  updateAll();
  sheetsSave([cpsRow(site, month, code)]);
}

function updateAgentVal(agent, month, code, wk, val) {
  const c = ensure(ensure(database.agents, agent), month);
  if (!c[code]) c[code] = { target: 20, s1: 0, s2: 0, s3: 0, s4: 0 };
  c[code][wk] = parseInt(val) || 0;
  updateAll();
  sheetsSave([agentRow(agent, month, code)]);
}

async function sauvegardeSauvage() {
  const role = document.getElementById('authRole').value;
  const site = document.getElementById('authSite').value;
  const who = document.getElementById('authAgent').value;
  const month = document.getElementById('entryMonthSelect').value;
  if (!site || !who) { alert("⚠️ Veuillez d'abord sélectionner votre rôle, site et nom pour sauvegarder."); return; }

  let rows = [];
  if (role === 'CPS' && database.cps[site] && database.cps[site][month]) {
    rows = Object.keys(database.cps[site][month]).map(code => cpsRow(site, month, code));
  } else if (role === 'ASD' && database.agents[who] && database.agents[who][month]) {
    rows = Object.keys(database.agents[who][month]).map(code => agentRow(who, month, code));
  }
  if (!rows.length) { alert("Aucune donnée à sauvegarder."); return; }
  const ok = await sheetsSave(rows);
  alert(ok ? "✅ Données sauvegardées avec succès dans Google Sheets !" : "❌ Erreur lors de la sauvegarde vers Google Sheets.");
}

// =====================================================
// 6. INTERFACE
// =====================================================
function verifyAdminAccess() {
  const nameInput = document.getElementById('adminUserField').value.trim();
  const passInput = document.getElementById('adminPassField').value.trim();
  const errorDiv = document.getElementById('adminLoginError');
  const match = ADMIN_CREDENTIALS.find(a => a.name.toLowerCase() === nameInput.toLowerCase() && a.pass === passInput);
  if (match) {
    isAdminAuthenticated = true;
    errorDiv.style.display = 'none';
    document.getElementById('adminLoginWrapper').style.display = 'none';
    document.getElementById('adminContentContainer').style.display = 'grid';
    renderAdminMatrices();
  } else {
    errorDiv.style.display = 'block';
  }
}

function show(id, btn) {
  document.querySelectorAll('.panel').forEach(x => x.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('nav button').forEach(x => x.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('title').textContent = btn.textContent.replace(/^[^A-Za-zÀ-ÿ]+/, '').trim();
  if (id === 'admin' && !isAdminAuthenticated) {
    document.getElementById('adminLoginWrapper').style.display = 'block';
    document.getElementById('adminContentContainer').style.display = 'none';
  }
  updateAll();
}

function fillSelect(id, html) {
  const el = document.getElementById(id), cur = el.value;
  el.innerHTML = html;
  if (cur && Array.prototype.some.call(el.options, o => o.value === cur)) el.value = cur;
}

function populateDropdowns() {
  fillSelect('siteSelect', '<option value="CONSOLIDE">-- Tous les CPS (Consolidé) --</option>' + structures.map(s => `<option value="${s.name}">CPS ${s.name}</option>`).join(''));
  fillSelect('dashAgentFilter', '<option value="CONSOLIDE">-- Tous les Agents (Consolidé) --</option>' + agentsSaisieList.map(a => `<option value="${a.name}">${a.name} (${a.site})</option>`).join(''));
  const mHtml = monthsList.map(m => `<option value="${m}">${m}</option>`).join('');
  fillSelect('entryMonthSelect', mHtml);
  fillSelect('rankingMonthSelect', mHtml);
  fillSelect('adminNoteMonthSelect', mHtml);
}

function populateAuthDropdowns() {
  const role = document.getElementById('authRole').value;
  let sites = [];
  if (role === 'CPS') sites = structures.map(s => s.name);
  else agentsSaisieList.forEach(a => { if (sites.indexOf(a.site) === -1) sites.push(a.site); });
  fillSelect('authSite', '<option value="">-- Choisir un site --</option>' + sites.map(s => `<option value="${s}">${s}</option>`).join(''));
  populateAuthNames();
}

function populateAuthNames() {
  const role = document.getElementById('authRole').value;
  const site = document.getElementById('authSite').value;
  let names = [];
  if (role === 'CPS') structures.forEach(s => { if (!site || s.name === site) names.push(s.agent); });
  else agentsSaisieList.forEach(a => { if (!site || a.site === site) names.push(a.name); });
  fillSelect('authAgent', '<option value="">-- Choisir un nom --</option>' + names.map(n => `<option value="${n}">${n}</option>`).join(''));
  checkAuth();
}

function checkAuth() {
  const role = document.getElementById('authRole').value;
  const site = document.getElementById('authSite').value;
  const agent = document.getElementById('authAgent').value;
  const status = document.getElementById('authStatus');
  const card = document.getElementById('saisieCard');
  const entryMonth = document.getElementById('entryMonthSelect').value;

  const deny = (msg) => { status.className = 'bad'; status.textContent = msg; card.style.opacity = '0.4'; card.style.pointerEvents = 'none'; };
  if (!site || !agent) return deny("Veuillez configurer l'ensemble des filtres d'accès.");

  const isValid = role === 'CPS'
    ? !!structures.find(s => s.name === site && s.agent === agent)
    : !!agentsSaisieList.find(a => a.site === site && a.name === agent);
  if (!isValid) return deny("Accès refusé : Le nom choisi n'est pas attribué à ce site pour ce profil.");

  status.className = 'good';
  status.textContent = `Accès autorisé (${role}) pour ${agent} sur ${site}.`;
  card.style.opacity = '1'; card.style.pointerEvents = 'auto';
  document.getElementById('saisieCardTitle').textContent = role === 'CPS' ? `Saisie CPS : ${site} (${agent})` : `Saisie Agent de Saisie : ${agent} (${site})`;
  document.getElementById('saisieCpsContainer').style.display = role === 'CPS' ? 'block' : 'none';
  document.getElementById('saisieAgentContainer').style.display = role === 'CPS' ? 'none' : 'block';

  const locked = !!database.locks[entryMonth];
  document.getElementById('lockNotice').style.display = locked ? 'block' : 'none';
  card.querySelectorAll('input').forEach(inp => inp.disabled = locked);
}

function toggleTimeSelectors() {
  const type = document.getElementById('timeFilterType').value;
  let html = '';
  if (type === 'month') monthsList.forEach(m => { html += `<option value="${m}">Mois : ${m}</option>`; });
  else if (type === 'week') monthsList.forEach(m => { ['S1','S2','S3','S4'].forEach(w => { html += `<option value="${m}_${w}">${m} - Semaine ${w}</option>`; }); });
  else if (type === 'quarter') html = `<option value="Q1">Q1 (Oct-Déc 2026)</option><option value="Q2">Q2 (Jan-Mar 2027)</option><option value="Q3">Q3 (Avr-Juin 2027)</option><option value="Q4">Q4 (Juil-Sep 2027)</option>`;
  else if (type === 'semester') html = `<option value="S1_FY">S1 (Oct 2026 - Mars 2027)</option><option value="S2_FY">S2 (Avr 2027 - Sep 2027)</option>`;
  else html = `<option value="FULL_YEAR">Année Complète</option>`;
  document.getElementById('periodValueSelect').innerHTML = html;
}

function aggregate(listOfNames, store, indicators, targetMonths, filterType, periodVal) {
  const agg = {};
  indicators.forEach(ind => { agg[ind.code] = { label: ind.label, target: 0, realized: 0 }; });
  listOfNames.forEach(name => {
    if (!store[name]) return;
    targetMonths.forEach(m => {
      const mData = store[name][m];
      if (!mData) return;
      for (const code in mData) {
        if (!agg[code]) continue;
        const d = mData[code];
        agg[code].target += d.target * (filterType === 'week' ? 0.25 : 1);
        if (filterType === 'week') agg[code].realized += (d[periodVal.split('_')[1].toLowerCase()] || 0);
        else agg[code].realized += d.s1 + d.s2 + d.s3 + d.s4;
      }
    });
  });
  return agg;
}

function aggRowsHtml(agg, totals) {
  let html = '';
  for (const code in agg) {
    const it = agg[code], gap = it.target - it.realized;
    const rate = it.target > 0 ? Math.round((it.realized / it.target) * 100) : 0;
    totals.t += it.target; totals.r += it.realized;
    html += `<tr><td><b>${code}</b><br><span class="muted">${it.label}</span></td><td>${Math.round(it.target)}</td><td>${it.realized}</td><td class="${gap > 0 ? 'bad' : 'good'}">${gap > 0 ? Math.round(gap) : 0}</td><td><b>${rate}%</b></td></tr>`;
  }
  return html;
}

function entryRowsHtml(indicators, data, defTarget, locked, fn, who, month) {
  let html = '';
  indicators.forEach(ind => {
    let d = data[ind.code];
    if (!d) { d = { target: defTarget, s1: 0, s2: 0, s3: 0, s4: 0 }; data[ind.code] = d; }
    const r = d.s1 + d.s2 + d.s3 + d.s4;
    const rate = d.target > 0 ? Math.round((r / d.target) * 100) : 0;
    html += `<tr><td><b>${ind.code}</b><br><span class="muted">${ind.label}</span></td><td><b>${d.target}</b></td>`;
    ['s1','s2','s3','s4'].forEach(w => {
      html += `<td><input type="number" min="0" class="editable-input" value="${d[w]}" ${locked ? 'disabled' : ''} onchange="${fn}('${who}','${month}','${ind.code}','${w}',this.value)"></td>`;
    });
    html += `<td><b>${r}</b></td><td><span class="${rate >= 100 ? 'good' : 'bad'}">${rate}%</span></td></tr>`;
  });
  return html;
}

function updateAll() {
  populateDropdowns();
  if (document.getElementById('authSite').options.length <= 1) populateAuthDropdowns();
  if (document.getElementById('periodValueSelect').options.length === 0) toggleTimeSelectors();

  const filterType = document.getElementById('timeFilterType').value;
  const periodVal = document.getElementById('periodValueSelect').value;
  const currentSite = document.getElementById('siteSelect').value;
  const currentAgentFilter = document.getElementById('dashAgentFilter').value;
  const entryMonth = document.getElementById('entryMonthSelect').value;

  let targetMonths = [];
  if (filterType === 'month') targetMonths = [periodVal];
  else if (filterType === 'week') targetMonths = [periodVal.substring(0, 7)];
  else if (filterType === 'quarter') {
    const q = { Q1: [0,3], Q2: [3,6], Q3: [6,9], Q4: [9,12] }[periodVal];
    if (q) targetMonths = monthsList.slice(q[0], q[1]);
  } else if (filterType === 'semester') targetMonths = periodVal === 'S1_FY' ? monthsList.slice(0, 6) : monthsList.slice(6, 12);
  else targetMonths = monthsList;

  const sitesToProcess = currentSite === 'CONSOLIDE' ? structures.map(s => s.name) : [currentSite];
  const agentsToProcess = agentsSaisieList.filter(ag =>
    (currentSite === 'CONSOLIDE' || ag.site === currentSite) && (currentAgentFilter === 'CONSOLIDE' || ag.name === currentAgentFilter)
  ).map(a => a.name);

  const totals = { t: 0, r: 0 };
  document.getElementById('dashTableBody').innerHTML = aggRowsHtml(aggregate(sitesToProcess, database.cps, indicatorsCps, targetMonths, filterType, periodVal), totals);
  document.getElementById('dashAgentTableBody').innerHTML = aggRowsHtml(aggregate(agentsToProcess, database.agents, indicatorsAgent, targetMonths, filterType, periodVal), totals);

  const gGap = totals.t - totals.r;
  document.getElementById('kpiGlobalRate').textContent = (totals.t > 0 ? Math.round((totals.r / totals.t) * 100) : 0) + '%';
  document.getElementById('kpiTarget').textContent = Math.round(totals.t);
  document.getElementById('kpiRealized').textContent = totals.r;
  document.getElementById('kpiGap').textContent = gGap > 0 ? Math.round(gGap) : 0;

  const authRole = document.getElementById('authRole').value;
  const authSite = document.getElementById('authSite').value;
  const authAgent = document.getElementById('authAgent').value;
  const locked = !!database.locks[entryMonth];
  if (authSite && authAgent) {
    if (authRole === 'CPS' && database.cps[authSite] && database.cps[authSite][entryMonth]) {
      document.getElementById('entryTableBody').innerHTML = entryRowsHtml(indicatorsCps, database.cps[authSite][entryMonth], 40, locked, 'updateCpsVal', authSite, entryMonth);
    } else if (authRole === 'ASD' && database.agents[authAgent] && database.agents[authAgent][entryMonth]) {
      document.getElementById('entryAgentTableBody').innerHTML = entryRowsHtml(indicatorsAgent, database.agents[authAgent][entryMonth], 20, locked, 'updateAgentVal', authAgent, entryMonth);
    }
  }

  renderRankings();
  if (isAdminAuthenticated) renderAdminMatrices();
}

function countReached(data, refPct) {
  let n = 0;
  for (const code in (data || {})) {
    const d = data[code], r = d.s1 + d.s2 + d.s3 + d.s4;
    if ((d.target > 0 ? (r / d.target) * 100 : 0) >= refPct) n++;
  }
  return n;
}

function renderRankings() {
  const refPct = parseFloat(document.getElementById('globalRefPercentage').value) || 80;
  const rankMonth = document.getElementById('rankingMonthSelect').value;
  const evalOf = n => (database.evals[n] && database.evals[n][rankMonth]) || { dataQuality: 0, trackingQuality: 0 };

  const cps = structures.map(st => {
    const reached = countReached(database.cps[st.name] && database.cps[st.name][rankMonth], refPct), ev = evalOf(st.name);
    return { name: st.name, reached, dq: ev.dataQuality, tq: ev.trackingQuality, total: reached + ev.dataQuality + ev.trackingQuality };
  }).sort((a, b) => b.total - a.total);
  document.getElementById('rankingCpsBody').innerHTML = cps.map((it, i) =>
    `<tr><td><b>#${i + 1}</b></td><td>${it.name}</td><td><b>${it.reached}</b></td><td>${it.dq}</td><td>${it.tq}</td><td><b>${it.total} pts</b></td></tr>`).join('');

  const ags = agentsSaisieList.map(ag => {
    const reached = countReached(database.agents[ag.name] && database.agents[ag.name][rankMonth], refPct), ev = evalOf(ag.name);
    return { name: ag.name, site: ag.site, reached, dq: ev.dataQuality, tq: ev.trackingQuality, total: reached + ev.dataQuality + ev.trackingQuality };
  }).sort((a, b) => b.total - a.total);
  document.getElementById('rankingAgentsBody').innerHTML = ags.map((it, i) =>
    `<tr><td><b>#${i + 1}</b></td><td>${it.name}</td><td>${it.site}</td><td><b>${it.reached}</b></td><td>${it.dq}</td><td>${it.tq}</td><td><b>${it.total} pts</b></td></tr>`).join('');
}

function renderAdminMatrices() {
  if (!isAdminAuthenticated) return;

  document.getElementById('matrixCpsHeader').innerHTML = '<th>CPS</th><th>Agent Associé</th>' + indicatorsCps.map(i => `<th>Cible : ${i.code}</th>`).join('') + '<th>Action</th>';
  document.getElementById('matrixCpsBody').innerHTML = structures.map((st, i) => {
    let r = `<tr><td><b>${st.name}</b></td><td>${st.agent}</td>`;
    indicatorsCps.forEach(ind => {
      const c = database.cps[st.name] && database.cps[st.name][monthsList[0]] && database.cps[st.name][monthsList[0]][ind.code];
      r += `<td><input type="number" class="editable-input" value="${c ? c.target : 40}" onchange="updateMatrixCpsTarget('${st.name}','${ind.code}',this.value)"></td>`;
    });
    return r + `<td><button class="btn" style="color:#c23b32;padding:3px 6px" onclick="removeCps(${i})">Supprimer</button></td></tr>`;
  }).join('');

  document.getElementById('matrixAgentHeader').innerHTML = '<th>Agent de Saisie</th><th>Site Affecté</th>' + indicatorsAgent.map(i => `<th>Cible : ${i.code}</th>`).join('') + '<th>Action</th>';
  document.getElementById('matrixAgentBody').innerHTML = agentsSaisieList.map((ag, i) => {
    let r = `<tr><td><b>${ag.name}</b></td><td>${ag.site}</td>`;
    indicatorsAgent.forEach(ind => {
      const c = database.agents[ag.name] && database.agents[ag.name][monthsList[0]] && database.agents[ag.name][monthsList[0]][ind.code];
      r += `<td><input type="number" class="editable-input" value="${c ? c.target : 20}" onchange="updateMatrixAgentTarget('${ag.name}','${ind.code}',this.value)"></td>`;
    });
    return r + `<td><button class="btn" style="color:#c23b32;padding:3px 6px" onclick="removeAgentEntry(${i})">Supprimer</button></td></tr>`;
  }).join('');

  const indRow = (ind, fn, i) => `<tr><td><b>${ind.code}</b></td><td>${ind.label}</td><td><button class="btn" style="color:#c23b32" onclick="${fn}(${i})">Supprimer</button></td></tr>`;
  document.getElementById('adminIndCpsBody').innerHTML = indicatorsCps.map((x, i) => indRow(x, 'removeIndCps', i)).join('');
  document.getElementById('adminIndAgentBody').innerHTML = indicatorsAgent.map((x, i) => indRow(x, 'removeIndAgent', i)).join('');

  document.getElementById('adminLockBody').innerHTML = monthsList.map(m => {
    const l = !!database.locks[m];
    return `<tr><td><b>${m}</b></td><td><span class="${l ? 'bad' : 'good'}">${l ? 'Verrouillé' : 'Ouvert'}</span></td><td><button class="btn ${l ? '' : 'primary'}" onclick="toggleLock('${m}')">${l ? 'Déverrouiller' : 'Verrouiller'}</button></td></tr>`;
  }).join('');

  const adminMonth = document.getElementById('adminNoteMonthSelect').value || monthsList[0];
  const noteRow = (label, type, name) => {
    ensure(database.evals, name);
    if (!database.evals[name][adminMonth]) database.evals[name][adminMonth] = { dataQuality: 0, trackingQuality: 0 };
    const ev = database.evals[name][adminMonth];
    return `<tr><td><b>${label}</b></td><td>${type}</td>
      <td><input type="number" min="-3" max="3" class="editable-input" value="${ev.dataQuality}" onchange="updateAdminEval('${name}','${adminMonth}','dataQuality',this.value)"></td>
      <td><input type="number" min="-2" max="2" class="editable-input" value="${ev.trackingQuality}" onchange="updateAdminEval('${name}','${adminMonth}','trackingQuality',this.value)"></td></tr>`;
  };
  document.getElementById('adminNotesBody').innerHTML =
    structures.map(st => noteRow('CPS : ' + st.name, 'CPS', st.name)).join('') +
    agentsSaisieList.map(ag => noteRow('Agent : ' + ag.name, 'ASD', ag.name)).join('');

  document.getElementById('auditTrailBody').innerHTML = database.audit.length === 0
    ? '<tr><td colspan="5" class="muted">Aucune modification enregistrée pour l\'instant.</td></tr>'
    : database.audit.map(a => `<tr><td>${a.timestamp}</td><td><b>${a.targetName}</b></td><td>${a.indicator}</td><td>${a.oldVal}</td><td><b>${a.newVal}</b></td></tr>`).join('');
}

// =====================================================
// 7. ACTIONS ADMIN
// =====================================================
function auditEntry(targetName, indicator, oldVal, newVal) {
  return { timestamp: new Date().toLocaleString(), timestampDesc: new Date().toISOString(), targetName, indicator, oldVal, newVal };
}

function updateMatrixCpsTarget(site, code, val) {
  const t = parseInt(val) || 0;
  const oldVal = database.cps[site][monthsList[0]][code].target;
  monthsList.forEach(m => { database.cps[site][m][code].target = t; });
  const a = auditEntry('CPS : ' + site, code, oldVal, t);
  database.audit.unshift(a);
  const rows = monthsList.map(m => cpsRow(site, m, code));
  rows.push({ sheet: 'audit', data: a });
  sheetsSave(rows);
  updateAll();
}

function updateMatrixAgentTarget(agent, code, val) {
  const t = parseInt(val) || 0;
  const oldVal = database.agents[agent][monthsList[0]][code].target;
  monthsList.forEach(m => { database.agents[agent][m][code].target = t; });
  const a = auditEntry('Agent : ' + agent, code, oldVal, t);
  database.audit.unshift(a);
  const rows = monthsList.map(m => agentRow(agent, m, code));
  rows.push({ sheet: 'audit', data: a });
  sheetsSave(rows);
  updateAll();
}

function toggleLock(m) {
  database.locks[m] = !database.locks[m];
  sheetsSave([{ sheet: 'locks', data: { month: m, locked: database.locks[m] ? 1 : 0 } }]);
  updateAll();
  checkAuth();
}

function updateAdminEval(entity, month, field, val) {
  let v = parseInt(val) || 0;
  if (field === 'dataQuality') v = Math.max(-3, Math.min(3, v));
  if (field === 'trackingQuality') v = Math.max(-2, Math.min(2, v));
  database.evals[entity][month][field] = v;
  sheetsSave([evalRow(entity, month)]);
  renderAdminMatrices();
  renderRankings();
}

function addCps() {
  const name = document.getElementById('newCpsName').value.trim().toUpperCase();
  const agent = document.getElementById('newCpsAgent').value.trim();
  if (!name || structures.some(s => s.name === name)) return;
  structures.push({ name, agent });
  seedLocal();
  document.getElementById('newCpsName').value = '';
  document.getElementById('newCpsAgent').value = '';
  saveConfig();
  populateAuthDropdowns();
  updateAll();
}

function removeCps(i) {
  if (!confirm('Supprimer ce CPS de la liste ? (les données restent dans le Google Sheet)')) return;
  structures.splice(i, 1);
  saveConfig();
  populateAuthDropdowns();
  updateAll();
}

function addAgentEntry() {
  const name = document.getElementById('newAgentNameInput').value.trim();
  const site = document.getElementById('newAgentSiteInput').value.trim().toUpperCase();
  if (!name || agentsSaisieList.some(a => a.name === name)) return;
  agentsSaisieList.push({ name, site });
  seedLocal();
  document.getElementById('newAgentNameInput').value = '';
  document.getElementById('newAgentSiteInput').value = '';
  saveConfig();
  populateAuthDropdowns();
  updateAll();
}

function removeAgentEntry(i) {
  if (!confirm('Supprimer cet agent de la liste ? (les données restent dans le Google Sheet)')) return;
  agentsSaisieList.splice(i, 1);
  saveConfig();
  populateAuthDropdowns();
  updateAll();
}

function addIndCps() {
  const code = document.getElementById('indCpsCode').value.trim().toUpperCase();
  const label = document.getElementById('indCpsLabel').value.trim();
  if (!code || indicatorsCps.some(x => x.code === code)) return;
  indicatorsCps.push({ code, label });
  seedLocal();
  document.getElementById('indCpsCode').value = '';
  document.getElementById('indCpsLabel').value = '';
  saveConfig();
  updateAll();
}

function removeIndCps(i) {
  if (!confirm('Supprimer cet indicateur ?')) return;
  indicatorsCps.splice(i, 1);
  saveConfig();
  updateAll();
}

function addIndAgent() {
  const code = document.getElementById('indAgentCode').value.trim().toUpperCase();
  const label = document.getElementById('indAgentLabel').value.trim();
  if (!code || indicatorsAgent.some(x => x.code === code)) return;
  indicatorsAgent.push({ code, label });
  seedLocal();
  document.getElementById('indAgentCode').value = '';
  document.getElementById('indAgentLabel').value = '';
  saveConfig();
  updateAll();
}

function removeIndAgent(i) {
  if (!confirm('Supprimer cet indicateur ?')) return;
  indicatorsAgent.splice(i, 1);
  saveConfig();
  updateAll();
}

function exportExcel() {
  alert("Fonctionnalité d'export Excel à implémenter avec une bibliothèque comme SheetJS. En attendant, vos données sont directement lisibles dans le Google Sheet (menu Fichier > Télécharger > Excel).");
}

// =====================================================
// 8. DÉMARRAGE
// =====================================================
seedLocal();
toggleTimeSelectors();
updateAll();
loadDataFromSheets();
</script>
</body>
</html>
