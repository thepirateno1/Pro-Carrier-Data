<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>FMCSA Fresh MC Lead Extractor — Pro</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#000;--surface:#0d1117;--surface2:#161b22;--border:#21262d;--border2:#30363d;--text:#e6edf3;--muted:#8b949e;--dim:#6e7681;--blue:#58a6ff;--accent:#1f6feb;--purple:#8957e5;--red:#f85149;--green:#3fb950}
body{font-family:'Segoe UI',-apple-system,Arial,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;display:flex;flex-direction:column;line-height:1.5}
.hidden{display:none!important}
.hero{background:linear-gradient(180deg,#0a0e1a 0%,#000 100%);border-bottom:1px solid var(--border);padding:30px 20px 24px;text-align:center;position:relative;overflow:hidden}
.hero::before{content:'';position:absolute;top:-120px;left:50%;transform:translateX(-50%);width:600px;height:240px;background:radial-gradient(ellipse,rgba(31,111,235,.18) 0%,transparent 70%);pointer-events:none}
.hero-badge{display:inline-block;background:rgba(31,111,235,.12);border:1px solid rgba(88,166,255,.35);color:var(--blue);font-size:11px;font-weight:700;letter-spacing:1.5px;padding:4px 14px;border-radius:20px;margin-bottom:12px;text-transform:uppercase}
.hero h1{font-size:clamp(22px,5vw,34px);font-weight:800;background:linear-gradient(90deg,#fff,#79b8ff);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}
.hero p{font-size:14px;color:var(--muted);margin-top:8px;max-width:680px;margin:8px auto 0}
.hero-tags{display:flex;justify-content:center;gap:8px;flex-wrap:wrap;margin-top:14px}
.hero-tag{font-size:12px;color:var(--muted);background:var(--surface2);border:1px solid var(--border2);padding:4px 12px;border-radius:14px}
.wrap{max-width:1240px;margin:24px auto;padding:0 14px;width:100%;flex:1}
.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,.4)}
.card-title{font-size:15px;font-weight:700;color:#fff;margin-bottom:14px;display:flex;align-items:center;gap:8px}
.card-title::before{content:'';width:3px;height:16px;background:linear-gradient(180deg,var(--blue),var(--purple));border-radius:2px}
label{font-size:11px;font-weight:600;color:var(--muted);display:block;margin-bottom:5px;text-transform:uppercase;letter-spacing:.5px}
input,select{width:100%;padding:10px 12px;background:#010409;border:1px solid var(--border2);border-radius:8px;color:var(--text);font-size:14px}
input:focus,select:focus{outline:none;border-color:var(--blue);box-shadow:0 0 0 3px rgba(31,111,235,.15)}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}
.checks{display:flex;gap:16px;flex-wrap:wrap;margin-top:12px;font-size:13px;color:#c9d1d9}
.checks label{display:inline-flex;align-items:center;gap:6px;text-transform:none;letter-spacing:0;font-size:13px;color:#c9d1d9;margin:0;font-weight:500}
.checks input{width:auto}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;padding:11px 22px;border:none;border-radius:9px;font-size:14px;font-weight:700;cursor:pointer;transition:transform .1s,box-shadow .15s}
.btn:active{transform:scale(.97)}
.btn-primary{background:linear-gradient(135deg,#1f6feb,#388bfd);color:#fff;box-shadow:0 4px 14px rgba(31,111,235,.35)}
.btn-primary:hover{box-shadow:0 6px 20px rgba(31,111,235,.5)}
.btn-green{background:linear-gradient(135deg,#1a7f37,#2ea043);color:#fff;box-shadow:0 4px 14px rgba(46,160,67,.3)}
.btn-row{display:flex;gap:12px;flex-wrap:wrap;margin-top:16px;align-items:center}
.errbox{color:var(--red);font-size:13.5px;margin-top:12px;font-weight:600}
.statusbar{position:sticky;top:8px;z-index:50;text-align:center;font-size:13.5px;font-weight:700;color:var(--blue);background:rgba(13,17,23,.95);border:1px solid rgba(88,166,255,.4);border-radius:20px;padding:8px 18px;margin:0 auto 12px;max-width:520px;box-shadow:0 4px 14px rgba(0,0,0,.5)}
.tabs{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}
.tab{flex:1;min-width:150px;padding:12px 14px;background:var(--surface2);border:1px solid var(--border2);border-radius:10px;color:var(--muted);font-size:14px;font-weight:700;cursor:pointer;text-align:center;transition:all .15s}
.tab.active{background:linear-gradient(135deg,#1f6feb,#388bfd);color:#fff;border-color:var(--blue);box-shadow:0 4px 14px rgba(31,111,235,.35)}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:14px}
.stat{background:var(--surface2);border:1px solid var(--border2);border-radius:10px;padding:12px;text-align:center}
.stat .num{font-size:22px;font-weight:800;color:var(--blue)}
.stat .lbl{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
.tblwrap{overflow-x:auto;border:1px solid var(--border2);border-radius:10px;-webkit-overflow-scrolling:touch}
table{width:100%;border-collapse:collapse;font-size:13px;min-width:1100px}
th{background:var(--surface2);color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.5px;padding:10px 8px;text-align:left;position:sticky;top:0}
td{padding:9px 8px;border-top:1px solid var(--border);white-space:nowrap;max-width:220px;overflow:hidden;text-overflow:ellipsis}
tr:hover td{background:rgba(88,166,255,.05)}
a.dlink{color:var(--blue);text-decoration:none}
a.dlink:hover{text-decoration:underline}
.pill{display:inline-block;padding:2px 9px;border-radius:12px;font-size:10.5px;font-weight:700;letter-spacing:.3px}
.pill.ok{background:rgba(46,160,67,.15);color:var(--green);border:1px solid rgba(46,160,67,.4)}
.pill.no{background:rgba(248,81,73,.12);color:var(--red);border:1px solid rgba(248,81,73,.4)}
.copybtn{background:none;border:none;cursor:pointer;font-size:13px;margin-left:4px;padding:2px}
.seg{display:inline-flex;background:var(--surface2);border:1px solid var(--border2);border-radius:10px;padding:4px;gap:4px;margin-bottom:14px}
.seg button{padding:9px 22px;border:none;border-radius:7px;background:transparent;color:var(--muted);font-size:13.5px;font-weight:700;cursor:pointer}
.seg button.on{background:linear-gradient(135deg,#6f42c1,#8957e5);color:#fff;box-shadow:0 2px 10px rgba(137,87,229,.4)}
.gate{max-width:460px;margin:40px auto;text-align:center}
.gate .lockicon{font-size:44px;margin-bottom:10px}
.licinput{letter-spacing:3px;text-align:center;font-size:16px;font-weight:700}
.licmsg{margin-top:12px;font-size:13.5px;font-weight:600;min-height:20px}
.licmsg.ok{color:var(--green)}.licmsg.bad{color:var(--red)}
.note{font-size:12px;color:var(--dim);margin-top:10px}
footer{border-top:1px solid var(--border);background:#05070c;padding:30px 20px;text-align:center;margin-top:24px}
.aslogo{width:58px;height:58px;margin:0 auto 12px;border-radius:16px;background:linear-gradient(135deg,#1f6feb,#8957e5);display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:800;color:#fff;box-shadow:0 6px 20px rgba(31,111,235,.4);letter-spacing:1px}
.fname{font-size:17px;font-weight:800;color:#fff}
.frole{font-size:12.5px;color:var(--blue);font-weight:600;letter-spacing:1.5px;text-transform:uppercase;margin-top:3px}
.fnote{font-size:12px;color:var(--dim);margin-top:10px}
</style>
</head>
<body>

<!-- ===== LICENSE GATE ===== -->
<div id="licGate" class="wrap">
  <div class="card gate">
    <div class="lockicon">🔒</div>
    <div class="card-title" style="justify-content:center">FMCSA Fresh MC Lead Extractor</div>
    <p style="font-size:13.5px;color:var(--muted);margin-bottom:18px">Enter your product key to activate, or start the free 3-run trial.</p>
    <label>Product Key</label>
    <input id="licKey" class="licinput" type="password" placeholder="XXXXXXXXXX" autocomplete="off">
    <div class="btn-row" style="justify-content:center">
      <button id="licBtn" class="btn btn-primary" onclick="activate()">Activate</button>
      <button id="trialBtn" class="btn btn-green" onclick="startTrial()">Start 3-Run Trial</button>
    </div>
    <div id="licMsg" class="licmsg"></div>
    <div id="trialInfo" class="note"></div>
  </div>
</div>

<!-- ===== APP ===== -->
<div id="app" class="wrap hidden">
  <div id="statusBox" class="statusbar hidden"></div>

  <div class="tabs">
    <div id="tabDate" class="tab active" onclick="switchTab('date')">📅 Date Range</div>
    <div id="tabMc" class="tab" onclick="switchTab('mc')">🔢 MC / USDOT Lookup</div>
    <div id="tabName" class="tab" onclick="switchTab('name')">🔎 Name Search</div>
  </div>

  <!-- TAB 1: DATE RANGE -->
  <div id="panelDate">
    <div class="card">
      <div class="card-title">Extract Fresh Carriers by Registration Date</div>
      <div class="grid">
        <div><label>Start Date (dd/mm/yyyy)</label><input id="startDate" placeholder="01/08/2026"></div>
        <div><label>End Date (dd/mm/yyyy)</label><input id="endDate" placeholder="31/08/2026"></div>
        <div><label>State</label><select id="stateSel"><option value="ALL">ALL States</option></select></div>
        <div><label>Authority</label><select id="authoritySel"><option value="ANY">Any</option><option value="AUTH">Authorized Only</option><option value="NON">Non-Authorized Only</option></select></div>
        <div><label>Entity Type</label><select id="entitySel"><option value="ALL">All Types</option><option value="C">Carrier</option><option value="B">Broker</option><option value="S">Shipper</option><option value="F">Freight Forwarder</option></select></div>
        <div><label>Max Fleet Size (Units)</label><input id="maxUnits" placeholder="e.g. 5" inputmode="numeric"></div>
      </div>
      <div class="checks">
        <label><input type="checkbox" id="interstateOnly"> Interstate only</label>
        <label><input type="checkbox" id="phoneOnly"> With phone only</label>
        <label><input type="checkbox" id="emailOnly"> With email only</label>
      </div>
      <div class="btn-row">
        <button class="btn btn-primary" onclick="runExtract()">⚡ Extract Leads</button>
        <button class="btn btn-green" onclick="downloadExcel()">📥 Download Excel (All Rows)</button>
      </div>
      <div id="apiErr" class="errbox"></div>
    </div>
    <div id="results" class="hidden">
      <div class="card">
        <div class="card-title">Results</div>
        <div class="stats">
          <div class="stat"><div class="num" id="statTotal">0</div><div class="lbl">Total</div></div>
          <div class="stat"><div class="num" id="statPhone">0</div><div class="lbl">With Phone</div></div>
          <div class="stat"><div class="num" id="statEmail">0</div><div class="lbl">With Email</div></div>
          <div class="stat"><div class="num" id="statAuth">0</div><div class="lbl">Authorized</div></div>
          <div class="stat"><div class="num" id="statStates">0</div><div class="lbl">States</div></div>
        </div>
        <div class="note" style="margin-bottom:8px">Showing first 300 rows on screen — Excel export contains ALL rows.</div>
        <div id="tableWrap" class="tblwrap"></div>
      </div>
    </div>
  </div>

  <!-- TAB 2: MC / USDOT -->
  <div id="panelMc" class="hidden">
    <div class="card">
      <div class="card-title">Lookup Carriers by MC or USDOT Number</div>
      <div class="seg">
        <button id="segMc" class="on" onclick="setNumType('mc')">MC Number</button>
        <button id="segDot" onclick="setNumType('dot')">USDOT</button>
      </div>
      <div class="grid">
        <div><label id="lblFrom">MC Number From</label><input id="mcFrom" placeholder="e.g. 1790000" inputmode="numeric"></div>
        <div><label id="lblTo">MC To (optional — range)</label><input id="mcTo" placeholder="e.g. 1791000" inputmode="numeric"></div>
        <div><label>Authority</label><select id="mcAuthoritySel"><option value="ANY">Any</option><option value="AUTH">Authorized Only</option><option value="NON">Non-Authorized Only</option></select></div>
        <div><label>Entity Type</label><select id="mcEntitySel"><option value="ALL">All Types</option><option value="C">Carrier</option><option value="B">Broker</option><option value="S">Shipper</option><option value="F">Freight Forwarder</option></select></div>
      </div>
      <div class="btn-row">
        <button class="btn btn-primary" onclick="runMcLookup()">🔍 Look Up</button>
        <button class="btn btn-green" onclick="downloadExcel()">📥 Download Excel (All Rows)</button>
      </div>
      <div id="mcErr" class="errbox"></div>
      <div class="note">Tip: MC lookup uses the primary docket (MC-). A few carriers may carry their MC in a secondary docket.</div>
    </div>
    <div id="mcResults" class="hidden">
      <div class="card">
        <div class="card-title">Results</div>
        <div class="stats">
          <div class="stat"><div class="num" id="mcStatTotal">0</div><div class="lbl">Total</div></div>
          <div class="stat"><div class="num" id="mcStatPhone">0</div><div class="lbl">With Phone</div></div>
          <div class="stat"><div class="num" id="mcStatEmail">0</div><div class="lbl">With Email</div></div>
          <div class="stat"><div class="num" id="mcStatAuth">0</div><div class="lbl">Authorized</div></div>
          <div class="stat"><div class="num" id="mcStatRange" style="font-size:15px">—</div><div class="lbl">Number Range</div></div>
        </div>
        <div class="note" style="margin-bottom:8px">Showing first 300 rows on screen — Excel export contains ALL rows.</div>
        <div id="mcTableWrap" class="tblwrap"></div>
      </div>
    </div>
  </div>

  <!-- TAB 3: NAME SEARCH -->
  <div id="panelName" class="hidden">
    <div class="card">
      <div class="card-title">Search Carriers by Company Name</div>
      <div class="grid">
        <div><label>Company Name Contains</label><input id="nameQuery" placeholder="e.g. TRUCKING"></div>
        <div><label>Registered After (dd/mm/yyyy)</label><input id="nameDate" placeholder="01/01/2026"></div>
        <div><label>State</label><select id="nameStateSel"><option value="ALL">ALL States</option></select></div>
        <div><label>Authority</label><select id="nameAuthoritySel"><option value="ANY">Any</option><option value="AUTH">Authorized Only</option><option value="NON">Non-Authorized Only</option></select></div>
        <div><label>Entity Type</label><select id="nameEntitySel"><option value="ALL">All Types</option><option value="C">Carrier</option><option value="B">Broker</option><option value="S">Shipper</option><option value="F">Freight Forwarder</option></select></div>
      </div>
      <div class="btn-row">
        <button class="btn btn-primary" onclick="runNameSearch()">🔎 Search</button>
        <button class="btn btn-green" onclick="downloadExcel()">📥 Download Excel (All Rows)</button>
      </div>
      <div id="nameErr" class="errbox"></div>
    </div>
    <div id="nameResults" class="hidden">
      <div class="card">
        <div class="card-title">Results</div>
        <div class="stats">
          <div class="stat"><div class="num" id="nmStatTotal">0</div><div class="lbl">Total</div></div>
          <div class="stat"><div class="num" id="nmStatPhone">0</div><div class="lbl">With Phone</div></div>
          <div class="stat"><div class="num" id="nmStatEmail">0</div><div class="lbl">With Email</div></div>
          <div class="stat"><div class="num" id="nmStatAuth">0</div><div class="lbl">Authorized</div></div>
        </div>
        <div class="note" style="margin-bottom:8px">Showing first 300 rows on screen — Excel export contains ALL rows.</div>
        <div id="nameTableWrap" class="tblwrap"></div>
      </div>
    </div>
  </div>
</div>

<!-- ===== FOOTER ===== -->
<footer>
  <div class="aslogo">AS</div>
  <div class="fname">Created by Mr. Adnan Shah</div>
  <div class="frole">Dispatch Lead Solutions</div>
  <div class="fnote">FMCSA Fresh MC Lead Extractor Pro • Data source: FMCSA Company Census (data.transportation.gov)</div>
</footer>

<script>
/* ==================== CONSTANTS ==================== */
var STATES=['ALL','AL','AK','AZ','AR','CA','CO','CT','DE','DC','FL','GA','HI','ID','IL','IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY'];
var SELECT_FIELDS='dot_number,legal_name,dba_name,phy_street,phy_city,phy_state,phy_zip,phone,fax,email_address,add_date,carrier_operation,power_units,carship,docket1prefix,docket1,docket1_status_code';
var TRIAL_RUNS=3, MASTER_KEYS=['JACK-7424-ADNAN-SHAH'], KEY_PREFIX='JACK-7424-ADNAN', SECRET='ADNAN-SHAH-2026-FMCSA';
var allRows=[], numType='mc', licensed=false;

/* ==================== HELPERS ==================== */
function el(id){return document.getElementById(id);}
function sleep(ms){return new Promise(function(r){setTimeout(r,ms);});}
function setStatus(t){var s=el('statusBox');if(t){s.textContent=t;s.classList.remove('hidden');window.scrollTo({top:0,behavior:'smooth'});}else{s.classList.add('hidden');}}
function parseDate(s){
  s=(s||'').trim();
  var m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if(m){var dd=+m[1],mm=+m[2],yy=m[3];if(mm<1||mm>12||dd<1||dd>31)throw new Error('Invalid date: '+s);return yy+String(mm).padStart(2,'0')+String(dd).padStart(2,'0');}
  m=s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if(m){return m[1]+String(+m[2]).padStart(2,'0')+String(+m[3]).padStart(2,'0');}
  throw new Error('Invalid date format. Use dd/mm/yyyy');
}
function addDays(ymd,n){var y=+ymd.slice(0,4),m=+ymd.slice(4,6),d=+ymd.slice(6,8);var dt=new Date(Date.UTC(y,m-1,d+n));return dt.getUTCFullYear()+String(dt.getUTCMonth()+1).padStart(2,'0')+String(dt.getUTCDate()).padStart(2,'0');}
function fmtDDMMYYYY(ymd){return ymd.slice(6,8)+'/'+ymd.slice(4,6)+'/'+ymd.slice(0,4);}
function daysOld(ymd){if(!ymd||ymd.length!==8)return '';var y=+ymd.slice(0,4),m=+ymd.slice(4,6),d=+ymd.slice(6,8);var then=new Date(Date.UTC(y,m-1,d)),now=new Date();var nowUtc=Date.UTC(now.getFullYear(),now.getMonth(),now.getDate());return Math.max(0,Math.floor((nowUtc-then.getTime())/86400000));}
function fmtPhone(p){var d=(p||'').replace(/\D/g,'');if(d.length===10)return '('+d.slice(0,3)+') '+d.slice(3,6)+'-'+d.slice(6);return p||'';}
function entityLabel(cs){
  if(!cs)return '';
  var out=[];if(cs.indexOf('C')>-1)out.push('Carrier');if(cs.indexOf('B')>-1)out.push('Broker');if(cs.indexOf('S')>-1)out.push('Shipper');if(cs.indexOf('F')>-1)out.push('Fwd');
  return out.join('; ');
}

/* ==================== API FETCH (429 retry + backoff) ==================== */
async function fetchJson(url,attempt){
  attempt=attempt||0;
  var res;
  try{res=await fetch(url);}catch(e){if(attempt<3){await sleep(800*Math.pow(2,attempt));return fetchJson(url,attempt+1);}throw new Error('Network error — check your internet connection.');}
  if(res.status===429){if(attempt<4){await sleep(2000*Math.pow(2,attempt));return fetchJson(url,attempt+1);}throw new Error('Rate limited by FMCSA API (429). Wait 1 minute and try again.');}
  if(res.status===400){var t=await res.text();throw new Error('API error 400 (bad query): '+t.slice(0,200));}
  if(!res.ok){throw new Error('API error '+res.status+' — try again in a moment.');}
  return res.json();
}

/* ==================== LICENSE (SHA-256 + HMAC) ==================== */
function sha256(ascii){
  function rr(v,c){return (v>>>c)|(v<<(32-c));}
  var mathPow=Math.pow,maxWord=mathPow(2,32),result='',words=[],asciiBitLength=ascii.length*8;
  var hash=sha256.h=sha256.h||[],k=sha256.k=sha256.k||[],primeCounter=k.length;
  var isComposite={};
  for(var candidate=2;primeCounter<64;candidate++){
    if(!isComposite[candidate]){
      for(var i=0;i<313;i+=candidate){isComposite[i]=candidate;}
      hash[primeCounter]=(mathPow(candidate,.5)*maxWord)|0;
      k[primeCounter++]=(mathPow(candidate,1/3)*maxWord)|0;
    }
  }
  ascii+='\x80';
  while(ascii.length%64-56)ascii+='\x00';
  for(i=0;i<ascii.length;i++){
    var j=ascii.charCodeAt(i);
    if(j>>8)return '';
    words[i>>2]|=j<<((3-i)%4)*8;
  }
  words[words.length]=(asciiBitLength/maxWord)|0;
  words[words.length]=asciiBitLength;
  for(j=0;j<words.length;){
    var w=words.slice(j,j+=16),oldHash=hash;
    hash=hash.slice(0,8);
    for(i=0;i<64;i++){
      var w15=w[i-15],w2=w[i-2];
      var a=hash[0],e=hash[4];
      var temp1=hash[7]+(rr(e,6)^rr(e,11)^rr(e,25))+((e&hash[5])^((~e)&hash[6]))+k[i]+(w[i]=(i<16)?w[i]:(w[i-16]+(rr(w15,7)^rr(w15,18)^(w15>>>3))+w[i-7]+(rr(w2,17)^rr(w2,19)^(w2>>>10))+w[i-2])|0);
      var temp2=(rr(a,2)^rr(a,13)^rr(a,22))+((a&hash[1])^(a&hash[2])^(hash[1]&hash[2]));
      hash=[(temp1+temp2)|0].concat(hash);
      hash[4]=(hash[4]+temp1)|0;
    }
    for(i=0;i<8;i++){hash[i]=(hash[i]+oldHash[i])|0;}
  }
  for(i=0;i<8;i++){for(j=3;j+1;j--){var b=(hash[i]>>(j*8))&255;result+=((b<16)?0:'')+b.toString(16);}}
  return result;
}
function hmacSha256(key,msg){
  if(key.length>64)key=sha256(key);
  var okey='',ikey='';
  for(var i=0;i<64;i++){
    var c=key.charCodeAt(i)||0;
    okey+=String.fromCharCode(0x5e^c);
    ikey+=String.fromCharCode(0x36^c);
  }
  return sha256(okey+sha256(ikey+msg));
}
function lsGet(k){try{return localStorage.getItem(k);}catch(e){return null;}}
function lsSet(k,v){try{localStorage.setItem(k,v);}catch(e){}}
function lsRm(k){try{localStorage.removeItem(k);}catch(e){}}
function licChecksum(v){
  var s='';try{s=JSON.stringify(v);}catch(e){return '';}
  return hmacSha256(SECRET,s).slice(0,16);
}
function loadTrial(){
  var raw=lsGet('fmcsa_trial');
  if(!raw)return {runs:TRIAL_RUNS,active:false};
  try{
    var parts=raw.split('|');var obj=JSON.parse(parts[0]);
    if(parts[1]!==licChecksum(obj)){lsRm('fmcsa_trial');return {runs:TRIAL_RUNS,active:false};}
    return obj;
  }catch(e){lsRm('fmcsa_trial');return {runs:TRIAL_RUNS,active:false};}
}
function saveTrial(t){lsSet('fmcsa_trial',JSON.stringify(t)+'|'+licChecksum(t));}
function isKeyValid(k){
  k=(k||'').trim().toUpperCase();
  if(MASTER_KEYS.indexOf(k)>-1)return true;
  var m=k.match(/^JACK-7424-([A-Z0-9]{4})-([A-Z0-9]{4})$/);
  if(!m)return false;
  var sig=hmacSha256(SECRET,m[1]+':'+m[2]).slice(0,8).toUpperCase();
  return sig===m[2];
}
function unlock(msg,isTrial){
  licensed=true;
  el('licGate').classList.add('hidden');
  el('app').classList.remove('hidden');
  if(msg)setStatus(msg),setTimeout(function(){setStatus('');},2500);
  initDefaults();
}
function lockScreen(reason){
  licensed=false;
  el('app').classList.add('hidden');
  el('licGate').classList.remove('hidden');
  var info=el('trialInfo'),m=el('licMsg');
  m.className='licmsg bad';
  m.textContent=reason||'';
  var t=loadTrial();
  info.textContent=t.runs>0?('Trial runs remaining: '+t.runs+' of '+TRIAL_RUNS):'Trial finished — enter product key to continue.';
}
function activate(){
  var k=el('licKey').value.trim(),m=el('licMsg');
  if(!k){m.className='licmsg bad';m.textContent='Please enter a product key.';return;}
  if(!isKeyValid(k)){m.className='licmsg bad';m.textContent='Invalid product key. Check and try again.';return;}
  lsSet('fmcsa_lic','1');
  saveTrial({runs:TRIAL_RUNS,active:false});
  m.className='licmsg ok';m.textContent='Activated successfully!';
  setTimeout(function(){unlock();},400);
}
function startTrial(){
  var t=loadTrial();
  if(t.runs<=0&&t.active){lockScreen('Trial already used on this device. Enter your product key.');return;}
  t.active=true;if(t.runs>TRIAL_RUNS)t.runs=TRIAL_RUNS;
  saveTrial(t);unlock();
}
function consumeRun(){
  var t=loadTrial();
  if(lsGet('fmcsa_lic')==='1')return true;
  if(!t.active){lockScreen('Please start the trial or enter a product key.');return false;}
  if(t.runs<=0){lockScreen('Trial finished ('+TRIAL_RUNS+' runs used). Enter your product key to continue.');return false;}
  t.runs--;saveTrial(t);
  return true;
}

/* ==================== TABS / NUMBER TYPE ==================== */
function switchTab(tab){
  ['Date','Mc','Name'].forEach(function(t){
    el('tab'+t).classList.toggle('active',t.toLowerCase()===tab);
  });
  el('panelDate').classList.toggle('hidden',tab!=='date');
  el('panelMc').classList.toggle('hidden',tab!=='mc');
  el('panelName').classList.toggle('hidden',tab!=='name');
}
function setNumType(t){
  numType=t;
  el('segDot').className=(t==='dot')?'on':'';
  el('segMc').className=(t==='mc')?'on':'';
  var label=(t==='mc')?'MC Number':'USDOT';
  el('lblFrom').textContent=label+' From';
  el('lblTo').textContent=label+' To (optional — range)';
}

/* ==================== SHARED FILTERS ==================== */
function authWhere(selId){
  var a=el(selId).value;
  if(a==='AUTH')return "docket1_status_code = 'A'";
  if(a==='NON')return "(docket1_status_code is null or docket1_status_code != 'A')";
  return null;
}
function entityWhere(selId){
  var e=el(selId).value;
  if(e==='ALL')return null;
  return "carship like '%"+e+"%'";
}

/* ==================== TAB 1: DATE RANGE ==================== */
async function runExtract(){
  if(!consumeRun())return;
  var errBox=el('apiErr');errBox.textContent='';
  el('results').classList.add('hidden');
  var where;
  try{
    var start=parseDate(el('startDate').value);
    var endExcl=addDays(parseDate(el('endDate').value),1);
    if(start>=endExcl)throw new Error('Start date must be before end date.');
    var parts=["add_date >= '"+start+"'","add_date < '"+endExcl+"'"];
    var st=el('stateSel').value;
    if(st!=='ALL')parts.push("phy_state = '"+st+"'");
    if(el('interstateOnly').checked)parts.push("carrier_operation = 'A'");
    if(el('phoneOnly').checked)parts.push('phone is not null');
    var mu=el('maxUnits').value.trim();
    if(mu&&/^\d+$/.test(mu))parts.push('power_units <= '+parseInt(mu,10));
    var a=authWhere('authoritySel');if(a)parts.push(a);
    var e=entityWhere('entitySel');if(e)parts.push(e);
    where=parts.join(' and ');
  }catch(ex){errBox.textContent=ex.message;return;}
  setStatus('⏳ Fetching newly registered carriers...');
  try{
    allRows=[];var offset=0;
    var base='https://data.transportation.gov/resource/az4n-8mr2.json?$select='+encodeURIComponent(SELECT_FIELDS)+'&$where='+encodeURIComponent(where)+'&$order=add_date%20DESC';
    while(true){
      var rows=await fetchJson(base+'&$limit=50000&$offset='+offset);
      allRows=allRows.concat(rows);
      if(rows.length<50000)break;
      offset+=rows.length;
      await sleep(300);
    }
  }catch(ex){setStatus('');errBox.textContent=ex.message;return;}
  if(!allRows.length){setStatus('');errBox.textContent='No records found for this date range. Check the dates or filters.';return;}
  if(el('emailOnly').checked){
    allRows=allRows.filter(function(r){return r.email_address;});
    if(!allRows.length){setStatus('');errBox.textContent='No carriers with email found (email coverage is ~55%). Try removing the "With email only" filter.';return;}
  }
  setStatus('');
  el('statTotal').textContent=allRows.length.toLocaleString();
  el('statPhone').textContent=allRows.filter(function(r){return r.phone;}).length.toLocaleString();
  el('statEmail').textContent=allRows.filter(function(r){return r.email_address;}).length.toLocaleString();
  el('statAuth').textContent=allRows.filter(function(r){return r.docket1_status_code==='A';}).length.toLocaleString();
  el('statStates').textContent=new Set(allRows.map(function(r){return r.phy_state;})).size;
  el('tableWrap').innerHTML=buildTableHtml(allRows);
  el('results').classList.remove('hidden');
}

/* ==================== TAB 2: MC / USDOT ==================== */
async function runMcLookup(){
  if(!consumeRun())return;
  var errBox=el('mcErr');errBox.textContent='';
  el('mcResults').classList.add('hidden');
  el('mcTableWrap').innerHTML='';
  var fromVal=el('mcFrom').value.trim().replace(/[^0-9]/g,'');
  var toVal=el('mcTo').value.trim().replace(/[^0-9]/g,'');
  if(!fromVal){errBox.textContent='Please enter an MC or USDOT number (or a range).';return;}
  var field=(numType==='mc')?'docket1':'dot_number';
  var where,label;
  if(toVal){
    if(parseInt(fromVal,10)>parseInt(toVal,10)){errBox.textContent='"From" must be smaller than or equal to "To".';return;}
    where=field+" between '"+fromVal+"' and '"+toVal+"'";
    label=numType.toUpperCase()+' range '+fromVal+' – '+toVal;
  }else{
    where=field+" = '"+fromVal+"'";
    label=numType.toUpperCase()+' '+fromVal;
  }
  if(numType==='mc')where+=" and docket1prefix = 'MC'";
  var a=authWhere('mcAuthoritySel');if(a)where+=' and '+a;
  var e=entityWhere('mcEntitySel');if(e)where+=' and '+e;
  setStatus('⏳ Fetching '+label+'...');
  try{
    allRows=[];var offset=0;
    var url='https://data.transportation.gov/resource/az4n-8mr2.json?$select='+encodeURIComponent(SELECT_FIELDS)+'&$where='+encodeURIComponent(where)+'&$order='+field;
    while(true){
      var rows=await fetchJson(url+'&$limit=50000&$offset='+offset);
      allRows=allRows.concat(rows);
      if(rows.length<50000)break;
      offset+=rows.length;
      await sleep(300);
    }
    var lo=parseInt(fromVal,10),hi=toVal?parseInt(toVal,10):null;
    allRows=allRows.filter(function(r){
      var n=parseInt(r[field],10);
      if(isNaN(n))return false;
      return (n>=lo)&&(hi===null||n<=hi);
    });
    if(!allRows.length){setStatus('');errBox.textContent='No carriers found for this '+numType.toUpperCase()+' / range. Check the number and try again.';return;}
    setStatus('');
    el('mcStatTotal').textContent=allRows.length.toLocaleString();
    el('mcStatPhone').textContent=allRows.filter(function(r){return r.phone;}).length.toLocaleString();
    el('mcStatEmail').textContent=allRows.filter(function(r){return r.email_address;}).length.toLocaleString();
    el('mcStatAuth').textContent=allRows.filter(function(r){return r.docket1_status_code==='A';}).length.toLocaleString();
    var nums=allRows.map(function(r){return parseInt(r[field],10)||0;});
    var mn=Math.min.apply(null,nums),mx=Math.max.apply(null,nums);
    el('mcStatRange').textContent=(mn===mx)?String(mn):mn+' – '+mx;
    el('mcTableWrap').innerHTML=buildTableHtml(allRows);
    el('mcResults').classList.remove('hidden');
  }catch(ex){setStatus('');errBox.textContent=ex.message;}
}

/* ==================== TAB 3: NAME SEARCH ==================== */
async function runNameSearch(){
  if(!consumeRun())return;
  var errBox=el('nameErr');errBox.textContent='';
  el('nameResults').classList.add('hidden');
  el('nameTableWrap').innerHTML='';
  var q=el('nameQuery').value.trim().replace(/'/g,'');
  if(!q){errBox.textContent='Please enter a company name to search.';return;}
  var parts=["upper(legal_name) like '%"+q.toUpperCase()+"%'"];
  try{
    var after=el('nameDate').value.trim();
    if(after)parts.push("add_date >= '"+parseDate(after)+"'");
  }catch(e){errBox.textContent=e.message;return;}
  var st=el('nameStateSel').value;
  if(st!=='ALL')parts.push("phy_state = '"+st+"'");
  var a=authWhere('nameAuthoritySel');if(a)parts.push(a);
  var en=entityWhere('nameEntitySel');if(en)parts.push(en);
  var where=parts.join(' and ');
  setStatus('⏳ Searching companies matching "'+q+'"...');
  try{
    allRows=[];var offset=0;
    var base='https://data.transportation.gov/resource/az4n-8mr2.json?$select='+encodeURIComponent(SELECT_FIELDS)+'&$where='+encodeURIComponent(where)+'&$order=add_date%20DESC';
    while(true){
      var rows=await fetchJson(base+'&$limit=50000&$offset='+offset);
      allRows=allRows.concat(rows);
      if(rows.length<50000)break;
      offset+=rows.length;
      await sleep(300);
    }
  }catch(ex){setStatus('');errBox.textContent=ex.message;return;}
  if(!allRows.length){setStatus('');errBox.textContent='No companies found matching "'+q+'". Try a shorter or different word.';return;}
  setStatus('');
  el('nmStatTotal').textContent=allRows.length.toLocaleString();
  el('nmStatPhone').textContent=allRows.filter(function(r){return r.phone;}).length.toLocaleString();
  el('nmStatEmail').textContent=allRows.filter(function(r){return r.email_address;}).length.toLocaleString();
  el('nmStatAuth').textContent=allRows.filter(function(r){return r.docket1_status_code==='A';}).length.toLocaleString();
  el('nameTableWrap').innerHTML=buildTableHtml(allRows);
  el('nameResults').classList.remove('hidden');
}

/* ==================== TABLE + EXCEL ==================== */
function buildTableHtml(rows){
  var headers=['USDOT','MC Number','Legal Name','DBA Name','Address','City','State','ZIP','Phone / Call','WhatsApp','Email','Add Date','Age (days)','Entity Type','Authority','Units'];
  var html='<table><tr>'+headers.map(function(h){return '<th>'+h+'</th>';}).join('')+'</tr>';
  var shown=rows.slice(0,300);
  for(var i=0;i<shown.length;i++){
    var r=shown[i];
    var auth=r.docket1_status_code==='A';
    var tel=(r.phone||'').replace(/\D/g,'');
    var wa=(tel.length===10)?'https://wa.me/1'+tel:'';
    var em=r.email_address||'';
    var mc=(r.docket1prefix==='MC'&&r.docket1)?('MC-'+r.docket1):(r.docket1?r.docket1prefix+'-'+r.docket1:'');
    html+='<tr>'
      +'<td>'+(r.dot_number||'')+'</td>'
      +'<td>'+mc+'</td>'
      +'<td>'+(r.legal_name||'')+'</td>'
      +'<td>'+(r.dba_name||'')+'</td>'
      +'<td>'+(r.phy_street||'')+'</td>'
      +'<td>'+(r.phy_city||'')+'</td>'
      +'<td>'+(r.phy_state||'')+'</td>'
      +'<td>'+(r.phy_zip||'')+'</td>'
      +'<td>'+(r.phone?'<a class="dlink" href="tel:'+tel+'">'+fmtPhone(r.phone)+'</a>':'')+'</td>'
      +'<td>'+(wa?'<a class="dlink" href="'+wa+'" target="_blank">💬 Chat</a>':'')+'</td>'
      +'<td>'+(em?'<a class="dlink" href="mailto:'+em+'">'+em+'</a><button class="copybtn" onclick="copyEmail(this,\''+em.replace(/'/g,"\\'")+'\')">📋</button>':'')+'</td>'
      +'<td>'+(r.add_date||'')+'</td>'
      +'<td>'+daysOld(r.add_date)+'</td>'
      +'<td>'+entityLabel(r.carship)+'</td>'
      +'<td><span class="pill '+(auth?'ok':'no')+'">'+(auth?'AUTHORIZED':'NON-AUTH')+'</span></td>'
      +'<td>'+(r.power_units||'')+'</td></tr>';
  }
  html+='</table>';
  return html;
}
function copyEmail(btn,email){
  function done(){btn.textContent='✅';setTimeout(function(){btn.textContent='📋';},1200);}
  if(navigator.clipboard&&navigator.clipboard.writeText){
    navigator.clipboard.writeText(email).then(done).catch(function(){fallbackCopy(email);done();});
  }else{fallbackCopy(email);done();}
}
function fallbackCopy(text){
  var ta=document.createElement('textarea');ta.value=text;
  document.body.appendChild(ta);ta.select();
  try{document.execCommand('copy');}catch(e){}
  document.body.removeChild(ta);
}
function downloadExcel(){
  if(!allRows.length)return;
  var mapOrder=[
    ['dot_number','USDOT Number'],['docket1prefix','MC Prefix'],['docket1','MC Number'],
    ['docket1_status_code','Authority (A=Active)'],['legal_name','Legal Name'],['dba_name','DBA Name'],
    ['phy_street','Address (Street)'],['phy_city','City'],['phy_state','State'],['phy_zip','ZIP Code'],
    ['phone','Phone'],['fax','Fax'],['email_address','Email'],['add_date','Add Date (YYYYMMDD)'],
    ['carrier_operation','Operation (A=Interstate)'],['power_units','Power Units'],['carship','Entity Codes']
  ];
  var data=allRows.map(function(r){
    var o={};
    mapOrder.forEach(function(p){o[p[1]]=r[p[0]]||'';});
    return o;
  });
  var ws=XLSX.utils.json_to_sheet(data);
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Fresh Carriers');
  XLSX.writeFile(wb,'fresh_mc_carriers.xlsx');
}

/* ==================== INIT ==================== */
function initDefaults(){
  ['stateSel','nameStateSel'].forEach(function(id){
    var sel=el(id);
    STATES.forEach(function(s){
      var o=document.createElement('option');o.value=s;o.textContent=s;sel.appendChild(o);
    });
  });
  var today=new Date();
  var endYmd=today.getUTCFullYear()+String(today.getUTCMonth()+1).padStart(2,'0')+String(today.getUTCDate()).padStart(2,'0');
  el('startDate').value=fmtDDMMYYYY(addDays(endYmd,-30));
  el('endDate').value=fmtDDMMYYYY(endYmd);
  setNumType('mc');
}
(function init(){
  if(lsGet('fmcsa_lic')==='1'){unlock();return;}
  var t=loadTrial();
  if(t.active&&t.runs>0){
    el('trialInfo').textContent='Trial runs remaining: '+t.runs+' of '+TRIAL_RUNS+'. Click "Start 3-Run Trial" to continue.';
  }else if(t.active&&t.runs<=0){
    el('trialInfo').textContent='Trial finished — enter product key to continue.';
  }
})();
</script>
</body>
</html>
