const COL={pkg:'ID_BALICKU',id:'ID_POLOZKY',name:'NAZEV_POLOZKY_KRATKY',inactive:'NEAKTIVNI',pm:'PM'};
const OLD_MARKS=['ZZZ','!!!'];
let DATA=[];          // [{__i, ID_BALICKU, ID_POLOZKY, NAZEV_POLOZKY_KRATKY, NEAKTIVNI, PM, __status}]
let CATALOG=[];       // [{id, name}] zdroj pro výběr náhrad
let CATALOG_CUSTOM=false;
let CURRENT_PM=null;
let CURRENT_COUNTRY='sk';
let CURRENT_USER='';
let CURRENT_USER_NAME='';
let CURRENT_PKG=null;
let LAST_UNDO=null;
let LOG=[];           // [{cas, pm, akce, balicek, stara_id, stary_nazev, nova_id, novy_nazev, propagace}]
let MODIFIED=new Map();// název balíčku -> Set(PM), kteří balíček změnili (pro ERP import = kompletní stav)
let DELETED_PKGS=new Map();// název balíčku -> {pm, pocet} smazané celé balíčky (ruční řešení v ERP)
let PKG_ORIGINAL=new Map();// aktuální název balíčku -> původní název před přejmenováním
function markModified(pkg,pm){ if(!MODIFIED.has(pkg)) MODIFIED.set(pkg,new Set()); MODIFIED.get(pkg).add(pm); }
function modifiedPkgsForPm(pm){ return [...MODIFIED.entries()].filter(([n,set])=>set.has(pm)&&!DELETED_PKGS.has(n)).map(([n])=>n); }
let WB=null, SHEET=null, FILENAME='balicky_sk.xlsx';
let modalCtx=null, picked=null;

const isOld = (s)=> OLD_MARKS.some(m=> (s||'').toUpperCase().includes(m.toUpperCase()));
const esc = (s)=> String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const now = ()=> new Date().toLocaleString('cs-CZ');
const API_BASE='/api';

function sessionStorageKey(){
  return `session:${CURRENT_COUNTRY}:${CURRENT_USER}:${CURRENT_PM||'no-pm'}`;
}
async function apiJson(path,opt){
  try{
    const r=await fetch(path,opt);
    if(!r.ok) return null;
    return await r.json();
  }catch(e){ return null; }
}
function cloneModifiedState(src){ return new Map([...src.entries()].map(([k,s])=>[k,new Set([...s])])); }
function cloneDeletedState(src){ return new Map([...src.entries()].map(([k,v])=>[k,{...v}])); }
function clonePkgOriginal(src){ return new Map([...src.entries()]); }
function snapshotMeta(){ return {logLen:LOG.length, modified:cloneModifiedState(MODIFIED), deleted:cloneDeletedState(DELETED_PKGS), pkgOriginal:clonePkgOriginal(PKG_ORIGINAL)}; }
function restoreMeta(meta){
  LOG=LOG.slice(0,meta.logLen);
  MODIFIED=cloneModifiedState(meta.modified);
  DELETED_PKGS=cloneDeletedState(meta.deleted);
  PKG_ORIGINAL=clonePkgOriginal(meta.pkgOriginal||new Map());
}
function originalPkgName(name){ return PKG_ORIGINAL.get(name)||name; }
function setUndoAction(fn,label){
  LAST_UNDO={fn,label:label||''};
  const b=document.getElementById('hUndo');
  if(b){ b.classList.remove('hidden'); b.title=label?('Zrušit: '+label):'Zrušit poslední změnu'; }
}
function clearUndoAction(){
  LAST_UNDO=null;
  const b=document.getElementById('hUndo');
  if(b) b.classList.add('hidden');
}
function undoLastChange(){
  if(!LAST_UNDO){ alert('Není co vrátit.'); return; }
  const fn=LAST_UNDO.fn;
  clearUndoAction();
  fn();
}
function showHelp(){
  alert(
`Jak aplikaci používat:
1) Přihlaste se loginem (prijmeni.jmeno) a 6místným PINem.
2) Vyberte PM vlastníka balíčků.
3) V každém balíčku nahraďte staré položky (ZZZ/!!!) přes Zaměnit nebo návrhy.
4) Pokud potřebujete, přidejte do balíčku novou položku tlačítkem "➕ Přidat položku".
5) Pokud uděláte chybu, použijte "Vrátit" u konkrétního řádku.
6) Po dokončení exportujte soubory pro Agendu.

Tlačítka exportu:
- Soubor pro Agendu = export jen právě otevřeného balíčku.
- Import do Agendy (ZIP po balíčcích) = export všech upravených balíčků najednou.
- Přehled (XLSX) = kontrolní tabulka pro člověka (není import pro Agendu).
- Uložit/Načíst rozpracované = ruční záloha a obnova (sekce Pokročilé).

Benefity:
- Když měníte starou položku, můžete stejnou záměnu propagovat i do dalších balíčků, kde se ta položka vyskytuje.
- Tím výrazně zrychlíte práci a snížíte riziko, že na některý balíček zapomenete.

Import do Agendy:
Jakmile máte importní soubor, můžete provést import balíčku do Agendy.
V zásadě máte 2 možnosti.
1) Najdete v Agendě balíček, který chcete upravovat, smažete z něj staré položky a následně provedete import balíčku.
   Importem se přidají nové položky, ty co jsou nezměněné, v balíčku zůstanou beze změny.
2) Naimportujete balíček jako nový, tedy založíte nový balíček (Název, Popis), provedete jeho import z připraveného souboru
   a starý původní balíček následně smažete.

Omezení:
- Zrušit poslední změnu vrací jen poslední akci (1 krok zpět).
- Sdílení práce je oddělené podle uživatele; jiný uživatel nevidí vaše rozpracované změny.`
  );
}
function restoreAuthFromStorage(){
  CURRENT_USER=localStorage.getItem('balicky_login')||'';
  CURRENT_USER_NAME=localStorage.getItem('balicky_user_name')||'';
}
function logoutUser(){
  clearUndoAction();
  CURRENT_USER='';
  CURRENT_USER_NAME='';
  CURRENT_PM=null;
  CURRENT_PKG=null;
  DATA=[]; CATALOG=[]; LOG=[];
  MODIFIED=new Map();
  DELETED_PKGS=new Map();
  PKG_ORIGINAL=new Map();
  localStorage.removeItem('balicky_login');
  localStorage.removeItem('balicky_user_name');
  document.getElementById('hPm').textContent='';
  document.getElementById('hSaved').textContent='';
  document.getElementById('hExport').classList.add('hidden');
  document.getElementById('hReset').classList.add('hidden');
  screenLoad();
}
function saveAuthToStorage(login,name){
  CURRENT_USER=String(login||'').trim().toLowerCase();
  CURRENT_USER_NAME=String(name||'').trim();
  localStorage.setItem('balicky_login',CURRENT_USER);
  localStorage.setItem('balicky_user_name',CURRENT_USER_NAME);
}
async function loginUser(login,pin){
  const r=await apiJson(`${API_BASE}/auth/login`,{
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body:JSON.stringify({login,pin})
  });
  if(!r||!r.ok||!r.user) return false;
  const full=[r.user.lastName,r.user.firstName].filter(Boolean).join(' ').trim();
  saveAuthToStorage(r.user.login,full||r.user.login);
  return true;
}
async function loadServerData(){
  const d=await apiJson(`${API_BASE}/data/${CURRENT_COUNTRY}`);
  if(!d) return false;
  if(Array.isArray(d.catalog) && d.catalog.length){
    buildCatalogFromKatalog(d.catalog);
  }else if(!CATALOG_CUSTOM && window.KATALOG && window.KATALOG.length){
    buildCatalogFromKatalog(window.KATALOG);
  }
  if(Array.isArray(d.packages) && d.packages.length){
    loadBalickyData(d.packages);
    return true;
  }
  return false;
}
async function loadSavedSessionForPm(){
  if(!CURRENT_PM) return false;
  const remote=await apiJson(`${API_BASE}/session?country=${encodeURIComponent(CURRENT_COUNTRY)}&user=${encodeURIComponent(CURRENT_USER)}&pm=${encodeURIComponent(CURRENT_PM)}`);
  if(remote && remote.session && remote.session.data && remote.session.data.length){
    restoreSession(remote.session);
    return true;
  }
  const local=await idbGet(sessionStorageKey());
  if(local && local.data && local.data.length){
    restoreSession(local);
    return true;
  }
  return false;
}

/* ---------- podobnost názvů (návrhy alternativ) ---------- */
let IDX=null; // {items:[{id,name,toks}], df:Map, inv:Map, N}
function normTxt(s){
  return String(s==null?'':s).normalize('NFKD').replace(/[̀-ͯ]/g,'')
    .toLowerCase().replace(/zzz-?|!!!|!!|!/g,' ').replace(/[^a-z0-9]+/g,' ').trim();
}
function tokenize(s){ return normTxt(s).split(' ').filter(t=>t && (t.length>=2 || /[0-9]/.test(t))); }
const isNum=t=>/[0-9]/.test(t);
const normId=v=>{ const s=String(v==null?'':v).trim().replace(/^"+|"+$/g,'').trim(); return s.replace(/^0+/,'')||(s?'0':''); };

// metadata z katalogu (Skup.výrobků, Výrobce, PM, Podskupina) pro dohledání u staré položky
let META_BY_ID=new Map();     // normId -> {group,vyr,pm,psk}
let HAS_GROUP=false;          // katalog obsahuje sloupec Skup.výrobků

// slovník TYPŮ výrobku – jemné dočištění uvnitř skupiny + záloha, když skupina chybí
const TYPES=[
  ['EPS',/\beps\b/],['XPS',/\bxps\b/],['PIR',/\bpir\b/],['PUR',/\bpur\b/],
  ['MW',/\bmw\b|miner|kamenn|\bvata\b|\bvlna\b|rockwool|nobasil|orsil|isover|ursa/],
  ['NATER',/penetr|nater|\bnat\b|barva|\bemail\b|\blak\b|lazur/],
  ['OMIETKA',/omietk|omitk/],
  ['STIERKA',/stierk/],
  ['SPAROVACKA',/sparov|skarov/],
  ['MALTA',/malt/],
  ['LEPIDLO',/lepid/],
  ['TMEL',/tmel|silikon|akryl/],
  ['KRYTINA',/skridl|\btaska\b|krytin|sindel|bobrovka|hrebenac|hrebanc/],
  ['DOSKA',/sadrokarton|\bgkb\b|\bgkf\b|\bdoska\b|\bdosky\b|\bplatn/],
  ['FOLIA',/\bfoli|paropriepust|parozabran|dpa\b/],
  ['PAS',/\bpas\b|asfalt/],
  ['LISTA',/\blista\b|\blisty\b|profil|kefov/],
  ['PLOSINA',/plosin|stupac|\bstep\b/],
  ['SPOJOVACI',/skrutk|klinec|klince|\bvrut|hmozdink|kotv|prichytk|priponk/],
  ['PALETA',/palet/]
];
function typeOf(nn){ for(const [c,re] of TYPES) if(re.test(nn)) return c; return ''; }
function buildIndex(){
  const items=CATALOG.map(c=>{
    const toks=tokenize(c.name);
    return {id:c.id,name:c.name,toks,
      brand:toks[0]||'',                       // značka = první slovo názvu
      type:typeOf(normTxt(c.name)),            // typ výrobku (slovník)
      group:c.group||'', vyr:c.vyr||'', pm:c.pm||'', psk:c.psk||'',
      nums:new Set(toks.filter(isNum))};       // čísla (tloušťka/pevnost/model/rozměr)
  });
  const df=new Map(), inv=new Map();
  items.forEach((it,idx)=>{ new Set(it.toks).forEach(t=>{
    df.set(t,(df.get(t)||0)+1);
    if(!inv.has(t)) inv.set(t,[]); inv.get(t).push(idx);
  }); });
  IDX={items,df,inv,N:items.length||1};
}
// ohodnocení kandidátů: shoda názvu (IDF) + bonusy za výrobce/PM/podskupinu/čísla
function scoreCand(idx,baseScore,o){
  const it=IDX.items[idx]; let s=baseScore;
  if(o.vyr && it.vyr===o.vyr) s+=6;          // stejný výrobce
  if(o.pm && it.pm===o.pm) s+=4;             // stejný PM
  if(o.psk && it.psk===o.psk) s+=3;          // stejná podskupina
  if(o.brand && it.brand===o.brand) s+=3;    // stejná značka (první slovo)
  o.nums.forEach(t=>{ if(it.nums.has(t)) s+=4; }); // shoda tloušťky/pevnosti/rozměru/modelu
  if(o.type && it.type===o.type) s+=3;       // stejný typ
  return s;
}
function suggest(name,n,oldId){
  if(!IDX||!IDX.N) return [];
  const meta=META_BY_ID.get(normId(oldId))||{};
  const o={ brand:(tokenize(name)[0])||'', type:typeOf(normTxt(name)), nums:new Set(tokenize(name).filter(isNum)),
            group:meta.group||'', vyr:meta.vyr||'', pm:meta.pm||'', psk:meta.psk||'' };
  const qToks=[...new Set(tokenize(name))];
  const cap=IDX.N*0.2;
  const score=new Map();
  qToks.forEach(t=>{
    const dfc=IDX.df.get(t)||0; if(!dfc||dfc>cap) return;
    const w=Math.log((IDX.N+1)/(dfc+1));
    IDX.inv.get(t).forEach(idx=>score.set(idx,(score.get(idx)||0)+w));
  });
  const all=[...score.entries()];
  const sortTop=arr=>arr.sort((a,b)=>b[1]-a[1]).map(([idx])=>({id:IDX.items[idx].id,name:IDX.items[idx].name}));
  if(HAS_GROUP && o.group){
    // PRIMÁRNĚ: stejná Skup.výrobků (vyloučit prokazatelně jiný typ)
    let grp=all.filter(([idx])=>IDX.items[idx].group===o.group)
               .filter(([idx])=>{ const t=IDX.items[idx].type; return !(o.type&&t&&t!==o.type); })
               .map(([idx,s])=>[idx,scoreCand(idx,s,o)]);
    grp=sortTop(grp);
    if(grp.length>=n) return grp.slice(0,n);
    // DOPLNĚNÍ: mimo skupinu, stejný typ a (výrobce nebo značka), s penalizací
    const have=new Set(grp.map(x=>x.id));
    let ext=all.filter(([idx])=>{ const it=IDX.items[idx];
        return it.group!==o.group && o.type && it.type===o.type && ((o.vyr&&it.vyr===o.vyr)||(o.brand&&it.brand===o.brand)); })
      .map(([idx,s])=>[idx,scoreCand(idx,s,o)-100]);
    ext=sortTop(ext).filter(x=>!have.has(x.id));
    return grp.concat(ext).slice(0,n);
  }
  // ZÁLOHA bez skupiny: typ (vyloučit konflikt) + preference značky
  let cand=all.filter(([idx])=>{ const t=IDX.items[idx].type; return !(o.type&&t&&t!==o.type); });
  if(o.brand){ const sb=cand.filter(([idx])=>IDX.items[idx].brand===o.brand); if(sb.length) cand=sb; }
  return sortTop(cand.map(([idx,s])=>[idx,scoreCand(idx,s,o)])).slice(0,n);
}
let SUG={}; // __i staré položky -> [ {id,name}, ... ]

/* ---------- ukládání rozpracované práce ---------- */
const DB_NAME='balicky_app', STORE='kv';
function idb(){ return new Promise((res,rej)=>{ const r=indexedDB.open(DB_NAME,1);
  r.onupgradeneeded=()=>{ if(!r.result.objectStoreNames.contains(STORE)) r.result.createObjectStore(STORE); };
  r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
async function idbSet(k,v){ try{ const db=await idb(); return new Promise((res,rej)=>{ const t=db.transaction(STORE,'readwrite'); t.objectStore(STORE).put(v,k); t.oncomplete=()=>res(); t.onerror=()=>rej(t.error); }); }catch(e){ console.warn('idbSet',e); } }
async function idbGet(k){ try{ const db=await idb(); return new Promise(res=>{ const rq=db.transaction(STORE,'readonly').objectStore(STORE).get(k); rq.onsuccess=()=>res(rq.result); rq.onerror=()=>res(undefined); }); }catch(e){ return undefined; } }
async function idbDel(k){ try{ const db=await idb(); db.transaction(STORE,'readwrite').objectStore(STORE).delete(k); }catch(e){} }

function serializeSession(){
  return {v:1, savedAt:new Date().toISOString(), pm:CURRENT_PM, country:CURRENT_COUNTRY, user:CURRENT_USER, user_name:CURRENT_USER_NAME, filename:FILENAME,
    data:DATA, log:LOG,
    modified:[...MODIFIED.entries()].map(([k,s])=>[k,[...s]]),
    deleted:[...DELETED_PKGS.entries()],
    pkgOriginal:[...PKG_ORIGINAL.entries()]};
}
function restoreSession(s){
  DATA=(s.data||[]).map((r,i)=>({__i:(r.__i!=null?r.__i:i),__status:r.__status||'',...r}));
  LOG=s.log||[]; FILENAME=s.filename||FILENAME; CURRENT_PM=s.pm||null;
  CURRENT_COUNTRY=(s.country||CURRENT_COUNTRY||'sk').toLowerCase();
  CURRENT_USER=(s.user||CURRENT_USER||'').toLowerCase();
  CURRENT_USER_NAME=s.user_name||CURRENT_USER_NAME||CURRENT_USER;
  MODIFIED=new Map((s.modified||[]).map(([k,a])=>[k,new Set(a)]));
  DELETED_PKGS=new Map(s.deleted||[]);
  PKG_ORIGINAL=new Map(s.pkgOriginal||[]);
}
let saveT=null;
function scheduleSave(){ clearTimeout(saveT); saveT=setTimeout(doSave,800); }
async function doSave(){
  if(!DATA.length) return;
  const session=serializeSession();
  await idbSet(sessionStorageKey(),session);
  if(CURRENT_PM){
    await apiJson(`${API_BASE}/session`,{
      method:'PUT',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({country:CURRENT_COUNTRY,user:CURRENT_USER,pm:CURRENT_PM,session})
    });
  }
  const e=document.getElementById('hSaved');
  if(e) e.textContent='✓ uloženo '+new Date().toLocaleTimeString('cs-CZ');
}

async function saveCatalogCache(){
  try{ await idbSet('catalog',{has_group:HAS_GROUP,custom:CATALOG_CUSTOM,
    items:CATALOG.map(c=>({id:c.id,name:c.name,vyr:c.vyr||'',pm:c.pm||'',group:c.group||'',psk:c.psk||''})),
    meta:[...META_BY_ID.entries()]}); }catch(e){ console.warn('catalog cache',e); }
}
async function restoreCatalogCache(){
  const c=await idbGet('catalog'); if(!c||!c.items||!c.items.length) return false;
  HAS_GROUP=!!c.has_group; CATALOG_CUSTOM=!!c.custom;
  CATALOG=c.items; CATALOG.forEach(x=>x._n=normTxt(x.name)+' '+normTxt(x.id));
  META_BY_ID=new Map(c.meta||[]); buildIndex(); return true;
}
async function clearSaved(){
  await idbDel(sessionStorageKey());
  if(CURRENT_PM){
    await fetch(`${API_BASE}/session?country=${encodeURIComponent(CURRENT_COUNTRY)}&user=${encodeURIComponent(CURRENT_USER)}&pm=${encodeURIComponent(CURRENT_PM)}`,{method:'DELETE'});
  }
}

// přenosný soubor (mezi PC přes OneDrive)
function saveSessionFile(){
  if(!DATA.length){ alert('Není co uložit.'); return; }
  const blob=new Blob([JSON.stringify(serializeSession())],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download='rozpracovane_balicky_'+(CURRENT_PM?safeName(CURRENT_PM)+'_':'')+stamp()+'.json'; a.click(); URL.revokeObjectURL(a.href);
}
function loadSessionFile(ev){
  const f=ev.target.files[0]; if(!f) return; const r=new FileReader();
  r.onload=e=>{ try{ const s=JSON.parse(e.target.result); restoreSession(s); afterRestore();
    alert('Rozpracovaná práce načtena (PM '+CURRENT_PM+').'); }catch(err){ alert('Soubor se nepodařilo načíst.'); } };
  r.readAsText(f);
}
function afterRestore(){
  document.getElementById('hPm').textContent='PM vlastník: '+(CURRENT_PM||'')+' · '+CURRENT_COUNTRY.toUpperCase()+' · uživatel: '+(CURRENT_USER_NAME||CURRENT_USER);
  document.getElementById('hExport').classList.add('hidden');
  document.getElementById('hReset').classList.remove('hidden');
  clearUndoAction();
  scheduleSave(); screenPkgs();
}

/* ---------- načtení ---------- */
function screenLoad(){
  clearUndoAction();
  document.getElementById('app').innerHTML = `
    <div class="center">
      <h2>Aktualizace produktových balíčků</h2>
      <p>Data balíčků se načítají automaticky ze serveru. Přihlaste se uživatelem a PINem.</p>
      <div class="card" style="text-align:left;margin-bottom:14px">
        <label>Země</label>
        <select id="countrySel">
          <option value="sk">Slovensko (SK)</option>
          <option value="cz">Česko (CZ)</option>
        </select>
        <label style="margin-top:10px">Login (prijmeni.jmeno)</label>
        <input type="text" id="loginInput" placeholder="napr. novak.jan">
        <label style="margin-top:10px">PIN (6 číslic)</label>
        <input type="password" id="pinInput" placeholder="••••••" maxlength="6" inputmode="numeric" pattern="[0-9]*">
        <div style="margin-top:10px;display:flex;gap:8px">
          <button class="btn" id="loginBtn">Přihlásit a načíst</button>
        </div>
      </div>
    </div>`;
  const countrySel=document.getElementById('countrySel');
  const loginInput=document.getElementById('loginInput');
  const pinInput=document.getElementById('pinInput');
  const loginBtn=document.getElementById('loginBtn');
  countrySel.value=(localStorage.getItem('balicky_country')||'sk').toLowerCase();
  restoreAuthFromStorage();
  loginInput.value=CURRENT_USER||'';
  const applyCountry=()=>{
    CURRENT_COUNTRY=countrySel.value.toLowerCase();
    localStorage.setItem('balicky_country',CURRENT_COUNTRY);
  };
  const loadAfterLogin=async()=>{
    const ok=await loadServerData();
    if(ok){ screenPm(); return; }
    alert('Serverova data se nepodarilo nacist. Zkuste to prosim znovu.');
  };
  loginBtn.onclick=async()=>{
    applyCountry();
    const login=(loginInput.value||'').trim().toLowerCase();
    const pin=(pinInput.value||'').trim();
    if(!login){ alert('Zadejte login.'); return; }
    if(!/^\d{6}$/.test(pin)){ alert('PIN musi mit 6 cislic.'); return; }
    const ok=await loginUser(login,pin);
    if(!ok){ alert('Neplatny login nebo PIN.'); return; }
    pinInput.value='';
    await loadAfterLogin();
  };
  // tiche nacteni pro rychly start
  (async()=>{
    applyCountry();
    if(CURRENT_USER){
      if(!CATALOG_CUSTOM && window.KATALOG && window.KATALOG.length) buildCatalogFromKatalog(window.KATALOG);
      await loadAfterLogin();
    }
  })();
}
// automatické načtení balíčků z balicky.js (window.BALICKY = [[ID_BALICKU,ID_POLOZKY,NAZEV,NEAKTIVNI,PM],...])
function loadBalickyData(B){
  PKG_ORIGINAL=new Map();
  DATA=B.map((r,i)=>({__i:i,__status:'',
    [COL.pkg]:r[0],[COL.id]:r[1],[COL.name]:r[2],[COL.inactive]:r[3],[COL.pm]:r[4]}));
  FILENAME='balicky_'+CURRENT_COUNTRY+'.js';
  if(!CATALOG_CUSTOM) buildCatalog();
}

function readFile(f){
  FILENAME=f.name;
  const r=new FileReader();
  r.onload=e=>{
    WB=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
    SHEET=WB.SheetNames[0];
    const rows=XLSX.utils.sheet_to_json(WB.Sheets[SHEET],{defval:''});
    if(!rows.length||!(COL.pkg in rows[0])){ alert('Soubor nemá očekávané sloupce ('+COL.pkg+', '+COL.id+', '+COL.name+'…).'); return; }
    PKG_ORIGINAL=new Map();
    DATA=rows.map((r,i)=>({__i:i,__status:'',...r}));
    buildCatalog();
    screenPm();
  };
  r.readAsArrayBuffer(f);
}

function buildCatalog(){
  if(CATALOG_CUSTOM) return;
  const seen=new Map();
  DATA.forEach(r=>{
    if(r[COL.inactive]==1||isOld(r[COL.name])) return;
    const id=r[COL.id];
    if(!seen.has(id)) seen.set(id,{id,name:r[COL.name],pm:r[COL.pm]||'',vyr:'',group:'',psk:''});
  });
  CATALOG=[...seen.values()];
  CATALOG.forEach(c=>c._n=normTxt(c.name)+' '+normTxt(c.id));
  buildIndex();
}

/* ---------- výběr PM ---------- */
function screenPm(){
  const pms=[...new Set(DATA.map(r=>r[COL.pm]).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'cs'));
  document.getElementById('hPm').textContent='Země: '+CURRENT_COUNTRY.toUpperCase()+' · uživatel: '+(CURRENT_USER_NAME||CURRENT_USER);
  document.getElementById('hReset').classList.remove('hidden');
  document.getElementById('hExport').classList.add('hidden');
  document.getElementById('app').innerHTML=`
    <div class="center">
      <h2>Vyberte vlastníka balíčků (PM)</h2>
      <p>Přihlášený uživatel může upravovat balíčky libovolného PM vlastníka.</p>
      <label>Vlastník balíčků (PM)</label>
      <select id="pmSel">${pms.map(p=>`<option>${esc(p)}</option>`).join('')}</select>
      <div style="margin-top:18px"><button class="btn" onclick="pickPm()">Pokračovat →</button></div>
    </div>`;
}
async function pickPm(){
  CURRENT_PM=document.getElementById('pmSel').value;
  document.getElementById('hPm').textContent='PM vlastník: '+CURRENT_PM+' · '+CURRENT_COUNTRY.toUpperCase()+' · uživatel: '+(CURRENT_USER_NAME||CURRENT_USER);
  document.getElementById('hExport').classList.add('hidden');
  document.getElementById('hReset').classList.remove('hidden');
  clearUndoAction();
  const restored=await loadSavedSessionForPm();
  if(!restored) screenPkgs();
  else afterRestore();
  scheduleSave();
}

/* ---------- seznam balíčků ---------- */
// balíčky, kde má aktuální PM aspoň jednu položku; položky = CELÝ balíček (napříč PM),
// protože ERP nahrazuje celý balíček a náhled musí odpovídat importu
function pkgsOfPm(){
  const mine=new Set();
  DATA.forEach(r=>{ if(r[COL.pm]===CURRENT_PM) mine.add(r[COL.pkg]); });
  const map=new Map();
  DATA.forEach(r=>{ if(!mine.has(r[COL.pkg]))return; const k=r[COL.pkg]; if(!map.has(k))map.set(k,[]); map.get(k).push(r); });
  return map;
}
function screenPkgs(){
  const map=pkgsOfPm();
  const list=[...map.entries()]
    .filter(([name])=>!DELETED_PKGS.has(name))
    .map(([name,items])=>({name,items,
      shared:new Set(items.map(r=>r[COL.pm])).size>1,
      oldN:items.filter(r=>isOld(r[COL.name])&&r.__status!=='done'&&r.__status!=='deleted').length}))
    .sort((a,b)=>b.oldN-a.oldN||a.name.localeCompare(b.name,'cs'));
  const totalOld=list.reduce((s,p)=>s+p.oldN,0);
  const delMine=[...DELETED_PKGS.entries()].filter(([,v])=>v.pm===CURRENT_PM).length;
  const modMine=modifiedPkgsForPm(CURRENT_PM).length;
  let myItems=0,otherItems=0;
  list.forEach(p=>p.items.forEach(r=>{ if(r.__status==='deleted')return; if(r[COL.pm]===CURRENT_PM)myItems++; else otherItems++; }));
  document.getElementById('app').innerHTML=`
    <div class="card">
      <div class="stat">
        <div><b>${list.length}</b> balíčků</div>
        <div><b>${myItems}</b> položek (můj PM)</div>
        <div><b>${otherItems}</b> položek (jiný PM)</div>
        <div><b>${totalOld}</b> starých položek k řešení</div>
        <div><b>${modMine}</b> upravených balíčků</div>
        <div><b>${delMine}</b> smazaných balíčků</div>
      </div>
    </div>
    ${CATALOG_CUSTOM
      ? `<div style="color:var(--ok);font-size:12px;margin-bottom:8px">✓ Katalog načten: <b>${CATALOG.length.toLocaleString('cs-CZ')}</b> položek${HAS_GROUP?' · návrhy podle Skup.výrobků + výrobce/PM':' · návrhy podle slovníku typů'}</div>`
      : '<div class="note" style="background:#fff7e6;border-color:#ffd98a;color:#7a4a00">⚠ Katalog položek (<b>katalog.js</b>) se nenačetl – musí ležet ve stejné složce jako tato aplikace a být plně stažený (ne jen zástupce z OneDrive). Návrhy zatím vybírám jen z aktivních položek v balíčcích.</div>'}
    <div class="toolbar">
      <input type="text" class="grow" id="pkgFilter" placeholder="Filtrovat balíčky…" oninput="filterPkgs()">
      <button class="btn sec sm" onclick="screenDeleted()">🗑 Smazané balíčky (${delMine})</button>
    </div>
    <details class="card" style="margin-bottom:14px;padding:10px 12px">
      <summary style="cursor:pointer;font-weight:600">Pokročilé</summary>
      <div class="toolbar" style="margin-top:10px">
        <button class="btn sec sm" onclick="exportXlsx()" title="Stáhnout aktualizovaný XLSX se změnami a logem">💾 Uložit změny</button>
        <button class="btn sec sm" onclick="saveSessionFile()" title="Uložit rozpracovanou práci jako soubor (přenos mezi PC)">💾 Uložit rozpracované</button>
        <button class="btn sec sm" onclick="document.getElementById('sessFile').click()" title="Načíst rozpracovanou práci ze souboru">📂 Načíst rozpracované</button>
        <input type="file" id="sessFile" accept=".json" class="hidden" onchange="loadSessionFile(event)">
      </div>
    </details>
    <div class="toolbar">
      <span style="color:var(--muted);font-size:13px">Import do Agendy (upravené balíčky):</span>
      <button class="btn sm" onclick="erpExportZip()">📤 Import do Agendy (ZIP po balíčcích)</button>
      <button class="btn sm sec" onclick="exportOverview()">📄 Přehled (XLSX)</button>
    </div>
    <div id="pkgList">${list.map(p=>pkgRow(p)).join('')||'<div class="empty">Žádné balíčky.</div>'}</div>`;
}
function pkgRow(p){
  const mod=MODIFIED.has(p.name);
  const original=originalPkgName(p.name);
  const stateTag = mod
    ? (p.oldN
      ? '<span class="tag old">rozpracováno</span>'
      : '<span class="tag done">upraveno</span>')
    : '';
  const nm=esc(p.name).replace(/'/g,"\\'");
  const renamedNote=original!==p.name?`<span class="meta">původně: ${esc(original)}</span>`:'';
  return `<div class="pkg" data-name="${esc(p.name).toLowerCase()}">
    <span class="name" style="cursor:pointer" onclick="openPkg('${nm}')">${esc(p.name)}</span>
    ${renamedNote}
    <span class="meta">${p.items.filter(r=>r.__status!=='deleted').length} položek</span>
    ${p.shared?'<span class="tag" style="background:#e7e0ff;color:#4b2db3" title="Balíček sdílený více PM">sdílený</span>':''}
    ${stateTag}
    <span class="badge ${p.oldN?'':'zero'}" onclick="openPkg('${nm}')" style="cursor:pointer">${p.oldN?p.oldN+' starých':'hotovo'}</span>
    <button class="btn sm sec" title="Přejmenovat balíček" onclick="renamePkgPrompt('${nm}')">✏</button>
    <button class="btn sm sec" title="Smazat celý balíček" onclick="deletePkg('${nm}')">🗑</button>
  </div>`;
}
function filterPkgs(){ const q=document.getElementById('pkgFilter').value.toLowerCase(); document.querySelectorAll('#pkgList .pkg').forEach(e=>{e.style.display=e.dataset.name.includes(q)?'':'none';}); }

/* ---------- detail balíčku ---------- */
function openPkg(name){ CURRENT_PKG=name; screenPkg(); }
function screenPkg(){
  // CELÝ balíček (napříč PM) – odpovídá tomu, co se vyexportuje do ERP
  const items=DATA.filter(r=>r[COL.pkg]===CURRENT_PKG);
  const otherPms=[...new Set(items.map(r=>r[COL.pm]).filter(p=>p!==CURRENT_PM))];
  SUG={};
  const rows=items.map(r=>{
    const old=isOld(r[COL.name]), done=r.__status==='done', del=r.__status==='deleted';
    const foreign=r[COL.pm]!==CURRENT_PM;
    const cls=del?'':(done?'done':(old?'old':''));
    let tag='';
    if(del) tag='<span class="tag" style="background:#fde2e2;color:#b3001a">smazáno</span>';
    else if(done) tag='<span class="tag done">vyřešeno</span>';
    else if(old) tag='<span class="tag old">stará</span>';
    if(foreign&&!del) tag+=` <span class="tag" style="background:#e7e0ff;color:#4b2db3" title="${esc(r[COL.pm])}">jiný PM</span>`;
    let act='', sug='';
    if(del){
      act=`<div class="row-actions"><button class="btn sm sec" onclick="undoRow(${r.__i})">Vrátit</button></div>`;
    } else if(done){
      act = r.__undo
        ? `<div class="row-actions"><button class="btn sm sec" onclick="undoDone(${r.__i})">Vrátit</button></div>`
        : '';
    } else if(old&&!done){
      act=`<div class="row-actions">
        <button class="btn sm" onclick="openModal('replace',${r.__i})">Zaměnit</button>
        <button class="btn sm sec" onclick="openModal('add',${r.__i})">Přidat novou</button>
        <button class="btn sm sec" style="color:#b3001a" onclick="deleteRow(${r.__i})">Smazat</button></div>`;
      const s=suggest(r[COL.name],3,r[COL.id]); SUG[r.__i]=s;
      if(s.length) sug=`<div class="sugwrap"><span class="suglbl">Navrhované:</span>`+
        s.map((c,k)=>`<span class="chip" title="Zaměnit za tuto" onclick="quickReplace(${r.__i},${k})"><span class="cid">${esc(c.id)}</span>${esc(c.name)}</span>`).join('')+`</div>`;
    } else if(!done){
      act=`<div class="row-actions"><button class="btn sm sec" style="color:#b3001a" onclick="deleteRow(${r.__i})">Smazat</button></div>`;
    }
    const nameStyle=del?'style="text-decoration:line-through;color:var(--muted)"':(foreign?'style="color:var(--muted)"':'');
    return `<tr class="${cls}">
      <td class="mono">${esc(r[COL.id])}</td>
      <td ${nameStyle}>${esc(r[COL.name])} ${tag}${sug}</td>
      <td>${act}</td></tr>`;
  }).join('');
  const shareNote = otherPms.length
    ? `<div class="note">Tento balíček je <b>sdílený</b> i s: ${esc(otherPms.join(', '))}. Zobrazené jsou <b>všechny</b> položky balíčku, protože Agenda nahrazuje celý balíček – přesně to se vyexportuje do importu. Cizí položky můžete nechat beze změny.</div>`
    : '';
  const origName=originalPkgName(CURRENT_PKG);
  const renamedInfo=origName!==CURRENT_PKG?`<div class="note" style="margin-top:8px">Původní název: <b>${esc(origName)}</b> → Nový název: <b>${esc(CURRENT_PKG)}</b></div>`:'';
  document.getElementById('app').innerHTML=`
    <div class="crumb" style="display:flex;align-items:center;gap:10px">
      <span style="flex:1"><a onclick="screenPkgs()">← Balíčky</a> / ${esc(CURRENT_PKG)}</span>
      <button class="btn sm sec" onclick="renamePkgPrompt('${esc(CURRENT_PKG).replace(/'/g,"\\'")}')">✏ Přejmenovat balíček</button>
      <button class="btn sm" onclick="openAddNew()" title="Přidat do balíčku úplně novou položku z katalogu">➕ Přidat položku</button>
      <button class="btn sm sec" onclick="erpExportOnePkg()" title="Stáhnout importní .txt pro Agendu jen pro tento balíček">📤 Soubor pro Agendu</button>
    </div>
    ${renamedInfo}
    ${shareNote}
    <div class="card">
      <table>
        <thead><tr><th style="width:130px">ID položky</th><th>Název</th><th style="width:240px"></th></tr></thead>
        <tbody>${rows||'<tr><td colspan=3 class="empty">Žádné položky.</td></tr>'}</tbody>
      </table>
    </div>`;
}

/* ---------- modal výběru náhrady ---------- */
function openModal(mode,i){
  const r=DATA.find(x=>x.__i===i);
  modalCtx={mode,i,old:r}; picked=null;
  document.getElementById('mTitle').textContent = mode==='replace'?'Zaměnit za aktuální položku':'Přidat aktuální položku';
  document.getElementById('mSub').innerHTML = `Stará položka: <b>${esc(r[COL.name])}</b> <span class="mono">(${esc(r[COL.id])})</span>`;
  document.getElementById('mPropagate').parentElement.style.display = mode==='replace'?'flex':'none';
  document.getElementById('mSearch').value='';
  document.getElementById('mConfirm').disabled=true;
  renderResults();
  document.getElementById('overlay').classList.add('on');
  document.getElementById('mSearch').focus();
}
function openAddNew(){
  modalCtx={mode:'addNew',pkg:CURRENT_PKG}; picked=null;
  document.getElementById('mTitle').textContent='Přidat novou položku do balíčku';
  document.getElementById('mSub').innerHTML=`Balíček: <b>${esc(CURRENT_PKG)}</b> — najděte položku podle názvu nebo ID.`;
  document.getElementById('mPropagate').parentElement.style.display='none';
  document.getElementById('mSearch').value='';
  document.getElementById('mConfirm').disabled=true;
  renderResults();
  document.getElementById('overlay').classList.add('on');
  document.getElementById('mSearch').focus();
}
function closeModal(){ document.getElementById('overlay').classList.remove('on'); }
function renderResults(){
  const q=normTxt(document.getElementById('mSearch').value);
  let res, hdr='';
  if(!q){
    const hasOld = modalCtx && modalCtx.old;
    res = hasOld ? suggest(modalCtx.old[COL.name],50,modalCtx.old[COL.id]) : CATALOG.slice(0,50);
    hdr = res.length?`<div style="padding:6px 12px;font-size:12px;color:var(--muted);background:#fafafa">${hasOld?'Navrhované podle podobnosti — nebo hledejte fulltextem výše':'Začněte psát pro vyhledání položky (níže ukázka katalogu)'}</div>`:'';
  } else {
    const terms=q.split(' ').filter(Boolean);
    res=CATALOG.filter(c=>terms.every(t=>c._n.includes(t))).slice(0,200);
  }
  document.getElementById('mResults').innerHTML = hdr + (res.length? res.map(c=>
    `<div class="res ${picked&&picked.id===c.id?'sel':''}" onclick="pickRes(${JSON.stringify({id:c.id,name:c.name}).replace(/"/g,'&quot;')})">
      <span class="rid">${esc(c.id)}</span><span>${esc(c.name)}</span></div>`).join('')
    : '<div class="empty">Nic nenalezeno.</div>');
}
function pickRes(c){ picked=c; document.getElementById('mConfirm').disabled=false; renderResults(); }
function applyReplace(old,pick,propagate){
  const oldId=old[COL.id];
  const targets = propagate
    ? DATA.filter(r=>r[COL.id]===oldId && isOld(r[COL.name]) && r.__status!=='done' && r.__status!=='deleted')
    : [old];
  const before=targets.map(r=>({__i:r.__i,id:r[COL.id],name:r[COL.name],inactive:r[COL.inactive],status:r.__status}));
  const meta=snapshotMeta();
  targets.forEach(r=>{
    const prev={id:r[COL.id],name:r[COL.name],inactive:r[COL.inactive],status:r.__status};
    LOG.push({cas:now(),pm:CURRENT_PM,akce:'záměna',balicek:r[COL.pkg],
      stara_id:r[COL.id],stary_nazev:r[COL.name],nova_id:pick.id,novy_nazev:pick.name,
      propagace:(propagate&&r!==old)?'ano':''});
    r[COL.id]=pick.id; r[COL.name]=pick.name; r[COL.inactive]=0; r.__status='done';
    r.__undo={type:'replace',prev};
    markModified(r[COL.pkg],r[COL.pm]);
  });
  setUndoAction(()=>{
    before.forEach(b=>{
      const r=DATA.find(x=>x.__i===b.__i); if(!r) return;
      r[COL.id]=b.id; r[COL.name]=b.name; r[COL.inactive]=b.inactive; r.__status=b.status;
    });
    restoreMeta(meta);
    scheduleSave(); screenPkg();
  },'poslední záměnu');
  scheduleSave();
  return targets.length;
}
function applyAdd(old,pick){
  const meta=snapshotMeta();
  const prev={__i:old.__i,status:old.__status};
  const newRow={__i:DATA.length,__status:'done',
    [COL.pkg]:old[COL.pkg],[COL.id]:pick.id,[COL.name]:pick.name,[COL.inactive]:0,[COL.pm]:CURRENT_PM};
  newRow.__undo={type:'removeRow'};
  DATA.push(newRow);
  old.__status='done';
  old.__undo={type:'add',addedI:newRow.__i,prevStatus:prev.status};
  markModified(old[COL.pkg],CURRENT_PM);
  LOG.push({cas:now(),pm:CURRENT_PM,akce:'přidání',balicek:old[COL.pkg],
    stara_id:old[COL.id],stary_nazev:old[COL.name],nova_id:pick.id,novy_nazev:pick.name,propagace:''});
  setUndoAction(()=>{
    DATA=DATA.filter(r=>r.__i!==newRow.__i);
    const rr=DATA.find(x=>x.__i===prev.__i); if(rr) rr.__status=prev.status;
    restoreMeta(meta);
    scheduleSave(); screenPkg();
  },'poslední přidání');
  scheduleSave();
}
function applyAddNew(pick){
  const meta=snapshotMeta();
  const newRow={__i:DATA.length,__status:'done',
    [COL.pkg]:modalCtx.pkg,[COL.id]:pick.id,[COL.name]:pick.name,[COL.inactive]:0,[COL.pm]:CURRENT_PM};
  newRow.__undo={type:'removeRow'};
  DATA.push(newRow);
  markModified(modalCtx.pkg,CURRENT_PM);
  LOG.push({cas:now(),pm:CURRENT_PM,akce:'přidání nové',balicek:modalCtx.pkg,
    stara_id:'',stary_nazev:'',nova_id:pick.id,novy_nazev:pick.name,propagace:''});
  setUndoAction(()=>{
    DATA=DATA.filter(r=>r.__i!==newRow.__i);
    restoreMeta(meta);
    scheduleSave(); screenPkg();
  },'poslední přidání');
  scheduleSave();
}
function confirmPick(){
  if(!picked) return;
  const {mode,old}=modalCtx;
  if(mode==='replace') applyReplace(old,picked,document.getElementById('mPropagate').checked);
  else if(mode==='addNew') applyAddNew(picked);
  else applyAdd(old,picked);
  closeModal(); screenPkg();
}
function undoDone(i){
  const r=DATA.find(x=>x.__i===i); if(!r||!r.__undo) return;
  const u=r.__undo;
  if(u.type==='replace' && u.prev){
    r[COL.id]=u.prev.id; r[COL.name]=u.prev.name; r[COL.inactive]=u.prev.inactive; r.__status=u.prev.status||'';
    delete r.__undo;
    LOG.push({cas:now(),pm:CURRENT_PM,akce:'vrácení záměny',balicek:r[COL.pkg],
      stara_id:'',stary_nazev:'',nova_id:'',novy_nazev:'',propagace:''});
  } else if(u.type==='add'){
    DATA=DATA.filter(x=>x.__i!==u.addedI);
    r.__status=u.prevStatus||''; delete r.__undo;
    LOG.push({cas:now(),pm:CURRENT_PM,akce:'vrácení přidání',balicek:r[COL.pkg],
      stara_id:'',stary_nazev:'',nova_id:'',novy_nazev:'',propagace:''});
  } else if(u.type==='removeRow'){
    const pkg=r[COL.pkg];
    DATA=DATA.filter(x=>x.__i!==i);
    LOG.push({cas:now(),pm:CURRENT_PM,akce:'odebrání přidané položky',balicek:pkg,
      stara_id:'',stary_nazev:'',nova_id:'',novy_nazev:'',propagace:''});
  }
  clearUndoAction();
  scheduleSave(); screenPkg();
}
function renamePkgPrompt(name){
  const oldName=String(name||'').trim();
  if(!oldName) return;
  const next=prompt('Nový název balíčku:',oldName);
  if(next==null) return;
  const newName=String(next).trim();
  if(!newName || newName===oldName) return;
  const exists=DATA.some(r=>r[COL.pkg]===newName);
  if(exists){ alert('Balíček s tímto názvem už existuje.'); return; }
  const meta=snapshotMeta();
  const rows=DATA.filter(r=>r[COL.pkg]===oldName);
  if(!rows.length){ alert('Balíček nebyl nalezen.'); return; }
  rows.forEach(r=>{ r[COL.pkg]=newName; });
  const original=originalPkgName(oldName);
  PKG_ORIGINAL.delete(oldName);
  if(original!==newName) PKG_ORIGINAL.set(newName,original);
  if(MODIFIED.has(oldName)){
    MODIFIED.set(newName,MODIFIED.get(oldName));
    MODIFIED.delete(oldName);
  }
  if(DELETED_PKGS.has(oldName)){
    DELETED_PKGS.set(newName,DELETED_PKGS.get(oldName));
    DELETED_PKGS.delete(oldName);
  }
  if(CURRENT_PKG===oldName) CURRENT_PKG=newName;
  markModified(newName,CURRENT_PM);
  LOG.push({cas:now(),pm:CURRENT_PM,akce:'přejmenování balíčku',balicek:newName,
    stara_id:'',stary_nazev:oldName,nova_id:'',novy_nazev:newName,propagace:''});
  setUndoAction(()=>{
    DATA.filter(r=>r[COL.pkg]===newName).forEach(r=>{ r[COL.pkg]=oldName; });
    restoreMeta(meta);
    CURRENT_PKG=(CURRENT_PKG===newName?oldName:CURRENT_PKG);
    scheduleSave();
    if(CURRENT_PKG===oldName) screenPkg(); else screenPkgs();
  },'přejmenování balíčku');
  scheduleSave();
  if(CURRENT_PKG===newName) screenPkg(); else screenPkgs();
}
// rychlá záměna kliknutím na navržený štítek
function quickReplace(i,k){
  const r=DATA.find(x=>x.__i===i); const pick=(SUG[i]||[])[k];
  if(!r||!pick) return;
  const cnt=DATA.filter(x=>x[COL.id]===r[COL.id]&&isOld(x[COL.name])&&x.__status!=='done'&&x.__status!=='deleted').length;
  let prop=true;
  if(cnt>1) prop=confirm('Stejná stará položka je i v '+(cnt-1)+' dalších balíčcích.\n\nOK = nahradit i tam, Zrušit = jen v tomto balíčku.');
  applyReplace(r,pick,prop);
  screenPkg();
}

/* ---------- mazání položek a balíčků ---------- */
function deleteRow(i){
  const r=DATA.find(x=>x.__i===i); if(!r)return;
  const meta=snapshotMeta();
  const prev=r.__status;
  r.__status='deleted'; markModified(r[COL.pkg],r[COL.pm]);
  delete r.__undo;
  LOG.push({cas:now(),pm:CURRENT_PM,akce:'smazání položky',balicek:r[COL.pkg],
    stara_id:r[COL.id],stary_nazev:r[COL.name],nova_id:'',novy_nazev:'',propagace:''});
  setUndoAction(()=>{
    const rr=DATA.find(x=>x.__i===i); if(rr) rr.__status=prev;
    restoreMeta(meta);
    scheduleSave(); screenPkg();
  },'smazání položky');
  scheduleSave(); screenPkg();
}
function undoRow(i){
  const r=DATA.find(x=>x.__i===i); if(!r)return;
  const meta=snapshotMeta();
  const prev=r.__status;
  r.__status='';
  delete r.__undo;
  LOG.push({cas:now(),pm:CURRENT_PM,akce:'vrácení smazání',balicek:r[COL.pkg],
    stara_id:r[COL.id],stary_nazev:r[COL.name],nova_id:'',novy_nazev:'',propagace:''});
  setUndoAction(()=>{
    const rr=DATA.find(x=>x.__i===i); if(rr) rr.__status=prev;
    restoreMeta(meta);
    scheduleSave(); screenPkg();
  },'vrácení smazání');
  scheduleSave(); screenPkg();
}
function deletePkg(name){
  const items=DATA.filter(r=>r[COL.pkg]===name);   // celý balíček (napříč PM)
  const shared=new Set(items.map(r=>r[COL.pm])).size>1;
  const meta=snapshotMeta();
  const before=items.map(r=>({__i:r.__i,status:r.__status}));
  let msg='Smazat CELÝ balíček "'+name+'" ('+items.length+' položek)?\n\nBalíček se přesune do přehledu smazaných – v Agendě ho pak ručně zneaktivníte/smažete.';
  if(shared) msg+='\n\nPOZOR: balíček je sdílený i s jinými PM – smaže se celý.';
  if(!confirm(msg)) return;
  items.forEach(r=>{ r.__status='deleted'; });
  DELETED_PKGS.set(name,{pm:CURRENT_PM,pocet:items.length});
  MODIFIED.delete(name); // smazané balíčky se neposílají do běžného importu
  LOG.push({cas:now(),pm:CURRENT_PM,akce:'smazání balíčku',balicek:name,
    stara_id:'',stary_nazev:items.length+' položek',nova_id:'',novy_nazev:'',propagace:''});
  setUndoAction(()=>{
    before.forEach(b=>{ const r=DATA.find(x=>x.__i===b.__i); if(r) r.__status=b.status; });
    restoreMeta(meta);
    scheduleSave(); screenPkgs();
  },'smazání balíčku');
  scheduleSave(); screenPkgs();
}
function undoPkg(name){
  const info=DELETED_PKGS.get(name); if(!info)return;
  const meta=snapshotMeta();
  const before=DATA.filter(r=>r[COL.pkg]===name).map(r=>({__i:r.__i,status:r.__status}));
  DATA.filter(r=>r[COL.pkg]===name).forEach(r=>{ if(r.__status==='deleted') r.__status=''; });
  DELETED_PKGS.delete(name);
  LOG.push({cas:now(),pm:CURRENT_PM,akce:'vrácení balíčku',balicek:name,
    stara_id:'',stary_nazev:'',nova_id:'',novy_nazev:'',propagace:''});
  setUndoAction(()=>{
    before.forEach(b=>{ const r=DATA.find(x=>x.__i===b.__i); if(r) r.__status=b.status; });
    restoreMeta(meta);
    scheduleSave(); screenDeleted();
  },'vrácení balíčku');
  scheduleSave(); screenDeleted();
}

/* ---------- přehled smazaných balíčků ---------- */
function screenDeleted(){
  const list=[...DELETED_PKGS.entries()].filter(([,v])=>v.pm===CURRENT_PM);
  const rows=list.map(([name,v])=>`<tr>
    <td>${esc(name)}</td><td class="mono">${v.pocet}</td>
    <td class="row-actions"><button class="btn sm sec" onclick="undoPkg('${esc(name).replace(/'/g,"\\'")}')">Obnovit</button></td></tr>`).join('');
  document.getElementById('app').innerHTML=`
    <div class="crumb"><a onclick="screenPkgs()">← Balíčky</a> / Smazané balíčky</div>
    <div class="note">Tyto balíčky jsou určené ke smazání. Import do Agendy je <b>neobsahuje</b> – v Agendě je zneaktivněte/smažte ručně podle tohoto seznamu.</div>
    <div class="toolbar">
      <button class="btn sm" onclick="exportDeleted('xlsx')" ${list.length?'':'disabled'}>📄 Seznam (XLSX)</button>
      <button class="btn sm sec" onclick="exportDeleted('csv')" ${list.length?'':'disabled'}>📄 Seznam (CSV)</button>
    </div>
    <div class="card"><table>
      <thead><tr><th>Balíček</th><th>Počet položek</th><th></th></tr></thead>
      <tbody>${rows||'<tr><td colspan=3 class="empty">Zatím žádné smazané balíčky.</td></tr>'}</tbody>
    </table></div>`;
}
function exportDeleted(fmt){
  const rows=[...DELETED_PKGS.entries()].filter(([,v])=>v.pm===CURRENT_PM)
    .map(([name,v])=>({ID_BALICKU:name,PM:v.pm,POCET_POLOZEK:v.pocet,AKCE_V_AGENDE:'zneaktivnit / smazat ručně'}));
  if(!rows.length){alert('Žádné smazané balíčky.');return;}
  if(fmt==='xlsx'){
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'smazane_balicky');
    XLSX.writeFile(wb,`smazane_balicky_${stamp()}.xlsx`);
  } else {
    downloadCsv(['ID_BALICKU','PM','POCET_POLOZEK','AKCE_V_AGENDE'],rows,`smazane_balicky_${stamp()}.csv`);
  }
}

/* ---------- katalog ---------- */
const cleanId=v=>String(v==null?'':v).trim().replace(/^"+|"+$/g,'').trim();
// automatické načtení z katalog.js (window.KATALOG = [[id,name,vyrobce,pm,skupina,podskupina],...])
function buildCatalogFromKatalog(K){
  if(!K||!K.length) return false;
  HAS_GROUP=K.some(r=>r[4]);
  META_BY_ID=new Map();
  K.forEach(r=>{ const id=normId(r[0]); if(id) META_BY_ID.set(id,{group:r[4]||'',vyr:r[2]||'',pm:r[3]||'',psk:r[5]||''}); });
  const seen=new Map();
  K.forEach(r=>{ const id=cleanId(r[0]), name=String(r[1]||''); if(id===''||isOld(name)) return;
    if(!seen.has(id)) seen.set(id,{id,name,vyr:r[2]||'',pm:r[3]||'',group:r[4]||'',psk:r[5]||''}); });
  CATALOG=[...seen.values()];
  CATALOG.forEach(c=>c._n=normTxt(c.name)+' '+normTxt(c.id));
  CATALOG_CUSTOM=true; buildIndex(); return true;
}
function loadCatalog(ev){
  const f=ev.target.files[0]; if(!f)return;
  const r=new FileReader();
  r.onload=e=>{
    const wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
    if(!rows.length){alert('Prázdný katalog.');return;}
    const keys=Object.keys(rows[0]);
    const find=(...res)=>keys.find(k=>res.some(re=>re.test(k)));
    const idK =find(/čís|cis|kod|kód|polož|polozk/i,/\bid\b/i)||keys[0];
    const nmK =find(/popis|nazev|název|name/i)||keys[1]||keys[0];
    const vyrK=find(/výrobc|vyrobc|manufact|brand/i);
    const pmK =find(/^pm$|\bpm\b/i);
    const grpK=find(/skup\.?\s*výr|skup\.?\s*vyr|skupin|komodit|sortiment|group|class/i);
    const pskK=find(/podskup/i);
    HAS_GROUP=!!grpK;
    // metadata z PLNÉHO katalogu (vč. starých) – pro dohledání u staré položky podle ID
    META_BY_ID=new Map();
    rows.forEach(x=>{ const id=normId(x[idK]); if(!id)return;
      META_BY_ID.set(id,{group:grpK?String(x[grpK]||'').trim():'',vyr:vyrK?String(x[vyrK]||'').trim():'',
        pm:pmK?String(x[pmK]||'').trim():'',psk:pskK?String(x[pskK]||'').trim():''}); });
    const all=rows.map(x=>({id:cleanId(x[idK]),name:String(x[nmK]||''),
      vyr:vyrK?String(x[vyrK]||'').trim():'',pm:pmK?String(x[pmK]||'').trim():'',
      group:grpK?String(x[grpK]||'').trim():'',psk:pskK?String(x[pskK]||'').trim():''})).filter(c=>c.id!=='');
    const active=all.filter(c=>!isOld(c.name));            // do návrhů jen aktuální položky
    const seen=new Map(); active.forEach(c=>{ if(!seen.has(c.id)) seen.set(c.id,c); });
    CATALOG=[...seen.values()];
    CATALOG.forEach(c=>c._n=normTxt(c.name)+' '+normTxt(c.id));
    CATALOG_CUSTOM=true;
    buildIndex();
    saveCatalogCache();
    alert('Katalog načten: '+CATALOG.length+' aktuálních položek (z '+all.length+' řádků).'
      +'\nSloupce → ID: '+idK+' | název: '+nmK+(vyrK?' | výrobce: '+vyrK:'')+(pmK?' | PM: '+pmK:'')+(grpK?' | skupina: '+grpK:'')
      +'\nVyřazeno '+(all.length-active.length)+' starých (ZZZ/!!!).'
      +'\nFiltr: '+(grpK?('Skup.výrobků „'+grpK+'" + doplnění podle typu/výrobce'):'sloupec skupiny nenalezen – použiju slovník typů'));
    screenPkgs();
  };
  r.readAsArrayBuffer(f);
}

/* ---------- export ---------- */
function rowToObj(r){const o={};o[COL.pkg]=r[COL.pkg];o[COL.id]=r[COL.id];o[COL.name]=r[COL.name];o[COL.inactive]=r[COL.inactive];o[COL.pm]=r[COL.pm];return o;}
function downloadCsv(cols,rows,fname){
  const sep=';';
  const q=v=>{v=String(v==null?'':v); return /[";\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v;};
  const lines=[cols.join(sep)].concat(rows.map(r=>cols.map(c=>q(r[c])).join(sep)));
  const blob=new Blob(['﻿'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=fname;a.click();URL.revokeObjectURL(a.href);
}
function exportXlsx(){
  if(!LOG.length && !confirm('Zatím nejsou žádné změny. Přesto uložit kopii?')) return;
  // aktualizovaný stav: bez smazaných položek a bez smazaných balíčků
  const out=DATA.filter(r=>r.__status!=='deleted' && !DELETED_PKGS.has(r[COL.pkg])).map(rowToObj);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(out),'balicky_sk');
  const logSheet=XLSX.utils.json_to_sheet(LOG.length?LOG:[{cas:'',pm:'',akce:'',balicek:'',stara_id:'',stary_nazev:'',nova_id:'',novy_nazev:'',propagace:''}]);
  XLSX.utils.book_append_sheet(wb,logSheet,'log_zmen');
  const d=new Date(),pad=n=>String(n).padStart(2,'0');
  const stamp=`${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
  XLSX.writeFile(wb,`balicky_sk_aktualizovany_${stamp}.xlsx`);
}

/* ---------- ERP import (přesný formát: TXT, TAB, CP1250, sloupce ID + POPIS) ---------- */
const CP1250_MAP={160:160,164:164,166:166,167:167,168:168,169:169,171:171,172:172,173:173,174:174,176:176,177:177,180:180,181:181,182:182,183:183,184:184,187:187,193:193,194:194,196:196,199:199,201:201,203:203,205:205,206:206,211:211,212:212,214:214,215:215,218:218,220:220,221:221,223:223,225:225,226:226,228:228,231:231,233:233,235:235,237:237,238:238,243:243,244:244,246:246,247:247,250:250,252:252,253:253,258:195,259:227,260:165,261:185,262:198,263:230,268:200,269:232,270:207,271:239,272:208,273:240,280:202,281:234,282:204,283:236,313:197,314:229,317:188,318:190,321:163,322:179,323:209,324:241,327:210,328:242,336:213,337:245,340:192,341:224,344:216,345:248,346:140,347:156,350:170,351:186,352:138,353:154,354:222,355:254,356:141,357:157,366:217,367:249,368:219,369:251,377:143,378:159,379:175,380:191,381:142,382:158,711:161,728:162,729:255,731:178,733:189,8211:150,8212:151,8216:145,8217:146,8218:130,8220:147,8221:148,8222:132,8224:134,8225:135,8226:149,8230:133,8240:137,8249:139,8250:155,8364:128,8482:153};
function cp1250Bytes(str){ const out=new Uint8Array(str.length); for(let i=0;i<str.length;i++){ const c=str.charCodeAt(i); out[i]=c<128?c:(CP1250_MAP[c]!=null?CP1250_MAP[c]:0x3F); } return out; }
function downloadBytes(bytes,fname,mime){ const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([bytes],{type:mime||'application/octet-stream'})); a.download=fname; a.click(); URL.revokeObjectURL(a.href); }
// jeden importní soubor = seznam položek balíčku: "ID<TAB>POPIS", CRLF
function erpTxt(items){ return ['ID\tPOPIS'].concat(items.map(r=>r.id+'\t'+r.name)).join('\r\n')+'\r\n'; }
function itemsOfPkg(name){ return DATA.filter(r=>r[COL.pkg]===name && r.__status!=='deleted').map(r=>({id:r[COL.id],name:r[COL.name]})); }

function modifiedPkgNames(){ return modifiedPkgsForPm(CURRENT_PM); }
function pkgPms(name){ return [...new Set(DATA.filter(r=>r[COL.pkg]===name).map(r=>r[COL.pm]))]; }
function otherPmsOf(name){ return pkgPms(name).filter(p=>p!==CURRENT_PM); }
function isShared(name){ return pkgPms(name).length>1; }
function confirmShared(){
  const sh=modifiedPkgNames().filter(isShared);
  if(!sh.length) return true;
  return confirm('POZOR – mezi upravenými balíčky jsou SDÍLENÉ s jinými PM:\n\n'+
    sh.map(n=>'• '+n+'  (sdíleno s: '+otherPmsOf(n).join(', ')+')').join('\n')+
    '\n\nAgenda přepíše celý balíček, takže import obsahuje i položky ostatních PM. Doporučuji to s nimi zkoordinovat.\n\nPokračovat v exportu?');
}
function safeName(s){return String(s).replace(/[\\/:*?"<>|]/g,'_').slice(0,80);}
function stamp(){const d=new Date(),p=n=>String(n).padStart(2,'0');return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;}

// ZIP s importními TXT (jeden na balíček) + soupis
async function erpExportZip(){
  const names=modifiedPkgNames();
  if(!names.length){ alert('Zatím jste neupravil žádný balíček.'); return; }
  if(!confirmShared()) return;
  if(typeof JSZip==='undefined'){ alert('Knihovna pro ZIP se nenačetla (chybí internet).'); return; }
  const zip=new JSZip(); const pad=String(names.length).length;
  const index=names.map((name,i)=>({
    poradi:String(i+1).padStart(pad,'0'),
    SOUBOR:String(i+1).padStart(pad,'0')+'_'+safeName(name)+'.txt',
    ID_BALICKU:name, PUVODNI_NAZEV_BALICKU:originalPkgName(name), POCET_POLOZEK:itemsOfPkg(name).length,
    SDILENY:isShared(name)?'ANO':'', SDILENO_S:otherPmsOf(name).join(', '), PM:CURRENT_PM }));
  zip.file('_SEZNAM_balicku.csv',
    '﻿'+['poradi;SOUBOR;ID_BALICKU;PUVODNI_NAZEV_BALICKU;POCET_POLOZEK;SDILENY;SDILENO_S;PM']
      .concat(index.map(r=>[r.poradi,r.SOUBOR,r.ID_BALICKU,r.PUVODNI_NAZEV_BALICKU,r.POCET_POLOZEK,r.SDILENY,'"'+r.SDILENO_S+'"',r.PM].join(';'))).join('\r\n'));
  names.forEach((name,i)=>{
    const fn=String(i+1).padStart(pad,'0')+'_'+safeName(name)+'.txt';
    zip.file(fn,cp1250Bytes(erpTxt(itemsOfPkg(name))));   // TXT, TAB, CP1250
  });
  const blob=await zip.generateAsync({type:'blob'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=`Agenda_import_${safeName(CURRENT_PM)}_${stamp()}.zip`; a.click(); URL.revokeObjectURL(a.href);
}
// jeden TXT pro právě otevřený balíček (rychlý import jednoho balíčku)
function erpExportOnePkg(){
  if(!CURRENT_PKG){ return; }
  downloadBytes(cp1250Bytes(erpTxt(itemsOfPkg(CURRENT_PKG))), safeName(CURRENT_PKG)+'.txt','text/plain');
}
// přehled upravených balíčků pro vlastní kontrolu (XLSX, ne pro ERP)
function exportOverview(){
  const names=modifiedPkgNames();
  if(!names.length){ alert('Zatím jste neupravil žádný balíček.'); return; }
  const rows=names.flatMap(n=>DATA.filter(r=>r[COL.pkg]===n && r.__status!=='deleted')
    .map(r=>({ID_BALICKU:r[COL.pkg],PUVODNI_NAZEV_BALICKU:originalPkgName(r[COL.pkg]),ID_POLOZKY:r[COL.id],NAZEV:r[COL.name],PM:r[COL.pm]})));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'prehled');
  XLSX.writeFile(wb,`prehled_upravenych_${safeName(CURRENT_PM)}_${stamp()}.xlsx`);
}

screenLoad();
