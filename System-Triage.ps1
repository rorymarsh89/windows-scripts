Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process -Force

# change window size to fit textt
# change window color

$pshost = Get-Host
$pswindow = $pshost.UI.RawUI

$newBufferSize = $pswindow.BufferSize
$newBufferSize.Width = 170
$newBufferSize.Height = 3000
$pswindow.BufferSize = $newBufferSize

$newWindowSize = $pswindow.WindowSize
$newWindowSize.Width = 170
$newWindowSize.Height = 50
$pswindow.WindowSize = $newWindowSize


$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"

Clear-Host

$Host.UI.RawUI.WindowTitle = "PCHH Triage"

# checks if script is running as admin
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "============================================" -ForegroundColor Red
    Write-Host "-- Script must be run as an Administrator --" -ForegroundColor Red
    Write-Host "-- Right-Click Start -> Terminal(Admin)   --" -ForegroundColor Red
    Write-Host "============================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Press any key to exit the script.." -ForegroundColor Yellow
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Stop-Process -Id $PID -Force
}

Write-Host ""

# Variable setup
$random = Get-Random -Minimum 1 -Maximum 5000
$minidump = "$env:SystemRoot\minidump"
$source = "$env:SystemRoot\minidump\*.dmp"

$desktop = [Environment]::GetFolderPath("Desktop")

$File = "$desktop\PCHH-Triage"
$infofile = "$File\specs-programs.txt"

$ziptar = "$File\PCHH-Triage_$random.zip"

$sys_eventlog_path = "$File\system_eventlogs.evtx"

$scriptVersion = "1.0"
$lookbackDays = 365   # match reliability history's ~1 year span; System log is size-capped anyway
$reliability_csv_path = "$File\reliability.csv"
$reliability_html_path = "$File\triage-report.html"

# Embedded HTML viewer (reliability + specs + system events, data injected at runtime)
$viewerTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PCHH Triage — System Report</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Albert+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#0b0c0f; --panel:#151720; --panel2:#1e212b; --line:#2c313d;
  --text:#f5f7fa; --dim:#aab0bd; --faint:#767e8c;
  --err:#ff6b6b; --warn:#ffc069; --ok:#3ddc97; --info:#6aa7ff;
}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:'Albert Sans',sans-serif;font-size:16px;min-height:100vh}
#appShell{display:flex;background:var(--bg);padding:40px 16px 40px 40px;gap:16px;min-height:100vh;width:100%}
.mono{font-family:'IBM Plex Mono',monospace}
#sidebar{width:230px;flex:0 0 230px;background:var(--panel2);border:1px solid var(--line);border-radius:16px;display:flex;flex-direction:column;padding:36px 16px 22px}
#brand{font-size:19px;font-weight:600;letter-spacing:.01em;padding:0 10px 18px}
#content{flex:1;min-width:0}
@media (max-width:900px){
  #appShell{flex-direction:column;padding:16px}
  #sidebar{width:auto;flex:0 0 auto;flex-direction:row;flex-wrap:wrap;align-items:center;padding:14px}
  #brand{padding:0 14px 0 0}
  #tabs{flex-direction:row;flex-wrap:wrap;gap:6px 14px;flex:1}
  .nav-group{display:contents}
  .nav-group-title{display:none}
  #sideFoot{width:100%;order:99;flex-direction:row;justify-content:space-between;padding-top:10px}
}
#tabs{display:flex;flex-direction:column;gap:2px;flex:1;overflow-y:auto}
.nav-group{margin-bottom:14px}
.nav-group-title{display:flex;align-items:center;justify-content:space-between;color:var(--text);font-size:18px;text-transform:none;letter-spacing:0;font-weight:600;padding:8px 10px;cursor:pointer;border-radius:6px;user-select:none}
.nav-group-title:hover{color:var(--text)}
.nav-group-title.static{cursor:default}
.nav-group-title.static:hover{color:var(--dim)}
.nav-group-title .chev{font-size:10px;transition:transform .15s;color:var(--faint)}
.nav-group.collapsed .nav-group-title .chev{transform:rotate(-90deg)}
.nav-group.collapsed .nav-group-items{display:none}
.nav-group-items{display:flex;flex-direction:column;gap:2px}
#sideFoot{margin-top:auto;padding-top:14px;border-top:1px solid var(--line);display:flex;flex-direction:column;gap:4px}
#sideFoot span{color:var(--faint);font-size:12px}
.tab{display:block;width:100%;text-align:left;background:none;border:none;color:var(--dim);font-family:inherit;font-size:16px;font-weight:500;padding:9px 10px 9px 20px;cursor:pointer;border-radius:9px}
.tab:hover{color:var(--text);background:color-mix(in srgb,var(--panel2) 60%,transparent)}
.tab.on{color:var(--text);background:var(--panel2)}
#summary{padding:0;display:flex;flex-direction:column;gap:6px;font-size:15.5px;line-height:1.55}


#summary .sline{color:var(--dim)}
#summary .sline b{color:inherit;font-weight:500}
#summary .slabel{color:var(--dim)}
#summary .sline{color:var(--text)}
.summary-kv{grid-template-columns:165px 1fr;margin-bottom:4px}
.summary-kv dt{color:var(--dim)}
.summary-kv dd b{font-weight:500}
#summary .notes-head{color:var(--faint);font-size:14px;text-transform:uppercase;letter-spacing:.08em;font-weight:500;margin-top:12px}
.notes{margin:4px 0 0 2px;padding-left:18px}
.notes li{margin:3px 0;color:var(--text)}
#summary .g{color:var(--ok)}
#summary .r{color:var(--err)}
#summary .y{color:var(--warn)}
.view{display:none}
body.tab-summary #summaryView{display:block}
body.tab-rel #relView{display:block}
body.tab-sys #sysView{display:block}
body.tab-drives #drivesView{display:block}
body.tab-gpu #gpuView{display:block}
body.tab-memory #memoryView{display:block}
body.tab-net #netView{display:block}
body.tab-security #securityView{display:block}
body.tab-processes #processesView{display:block}
body.tab-apps #appsView{display:block}
body.tab-updates #updatesView{display:block}
body.tab-extensions #extensionsView{display:block}
body.tab-dumps #dumpsView{display:block}
#pageTitle{padding:36px 36px 0;font-size:40px;font-weight:700;letter-spacing:-.01em;color:var(--text);max-width:1160px}
#pageTitleSub{color:var(--info);font-weight:600}
#summaryView,#sysView,#drivesView,#netView,#securityView,#appsView,#dumpsView,#memoryView,#gpuView,#processesView,#extensionsView,#updatesView{padding:20px 36px 64px;max-width:1160px}
.sys-ok{color:var(--ok);padding:24px 0;font-size:16px}
.sys-note{color:var(--faint);font-size:13px;margin-bottom:14px}
.spec-section{margin-bottom:40px}
.spec-section h2{font-size:14px;font-weight:600;color:var(--faint);text-transform:uppercase;letter-spacing:.08em;padding:4px 0 14px;border-bottom:1px solid var(--line);margin-bottom:20px}
.kv{display:grid;grid-template-columns:210px 1fr;gap:7px 16px;font-size:15px}
.kv dt{color:var(--dim)}
.kv dd{word-break:break-word}
.kv dd.flag-off{color:var(--warn)}
.drive-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(340px,1fr));gap:18px}
.drive{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:22px}
.drive h3{font-size:17px;font-weight:600;margin-bottom:4px}
.drive .sub{color:var(--dim);font-size:14.5px;margin-bottom:16px}
.drive .meter{height:6px;background:var(--panel2);border-radius:3px;overflow:hidden;margin-bottom:6px}
.drive .meter div{height:100%;background:var(--info)}
.drive .meter.low div{background:var(--warn)}
.drive .use{color:var(--dim);font-size:14px}
.drive.smart-bad{border-color:var(--err)}
.smart-kv{grid-template-columns:1fr auto;font-size:14.5px;gap:7px 12px}
.smart-kv dt{color:var(--dim)}
.smart-kv dd{text-align:right;font-family:'IBM Plex Mono',monospace}
.proc-head{display:grid;grid-template-columns:1fr 110px 110px;color:var(--faint);font-size:13px;text-transform:uppercase;letter-spacing:.06em;padding:6px 4px;border-bottom:1px solid var(--line);margin-top:8px}
.proc-row{display:grid;grid-template-columns:1fr 110px 110px;padding:5px 4px;border-bottom:1px solid color-mix(in srgb,var(--line) 40%,transparent);font-size:14.5px}
.proc-row span:nth-child(2),.proc-row span:nth-child(3),.proc-head span:nth-child(2),.proc-head span:nth-child(3){text-align:right}
.pager{display:flex;gap:12px;align-items:center;margin-top:12px}
.pg-btn{background:var(--panel);border:1px solid var(--line);border-radius:6px;color:var(--dim);font-family:inherit;font-size:14px;padding:6px 14px;cursor:pointer}
.pg-btn:hover:not(:disabled){color:var(--text);border-color:var(--dim)}
.pg-btn:disabled{opacity:.35;cursor:default}
.pg-info{color:var(--faint);font-size:13.5px}
.sorth{cursor:pointer;user-select:none}
.sorth:hover{color:var(--text)}
#procSearch{background:var(--panel);border:1px solid var(--line);border-radius:6px;color:var(--text);padding:8px 12px;font-size:14px;font-family:inherit;width:260px;margin-bottom:6px}
#procSearch:focus{outline:none;border-color:var(--dim)}
#progSearch{background:var(--panel);border:1px solid var(--line);border-radius:6px;color:var(--text);padding:8px 12px;font-size:14px;font-family:inherit;width:260px;margin-bottom:10px}
#hfSearch{background:var(--panel);border:1px solid var(--line);border-radius:6px;color:var(--text);padding:8px 12px;font-size:14px;font-family:inherit;width:260px;margin-bottom:10px}
#hfSearch:focus{outline:none;border-color:var(--dim)}
#progSearch:focus{outline:none;border-color:var(--dim)}
#progList{columns:3;column-gap:24px;font-size:14.5px;line-height:1.9;color:var(--dim)}
#progList div{break-inside:avoid}
@media (max-width:900px){#progList{columns:2}}
@media (max-width:600px){#progList{columns:1}.kv{grid-template-columns:1fr;gap:0}.kv dt{margin-top:8px}}
h1{font-size:24px;font-weight:600;letter-spacing:.01em}
#range{color:var(--dim);font-size:14px}
#drop{display:block;margin:0 0 16px;font-size:12px;color:var(--faint);border:1px dashed var(--line);border-radius:8px;padding:8px 10px;cursor:pointer;text-align:center}
#drop:hover{color:var(--dim);border-color:var(--dim)}
body.dragging #drop{color:var(--info);border-color:var(--info)}

/* timeline */
#timeline{padding:18px 0 6px}
#tlHead{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:8px}
#tlRange{color:var(--text);font-size:15px;font-weight:500}
#tlHint{color:var(--faint);font-size:13.5px}
#tl-inner{display:flex;align-items:stretch;gap:10px}
#tl-main{flex:1;min-width:0}
.tl-nav{background:var(--panel);border:1px solid var(--line);border-radius:8px;color:var(--dim);font-size:20px;width:34px;cursor:pointer;font-family:inherit;align-self:stretch}
.tl-nav:hover:not(:disabled){color:var(--text);border-color:var(--dim)}
.tl-nav:disabled{opacity:.3;cursor:default}
#bars{display:flex;align-items:flex-end;gap:6px;height:72px;border-bottom:1px solid var(--line);padding-bottom:1px}
.bar{flex:1;display:flex;flex-direction:column;justify-content:flex-end;gap:2px;cursor:pointer;min-width:4px;border-radius:3px 3px 0 0;position:relative}
.bar div{width:100%}
.bar .seg-err{background:var(--err)}
.bar .seg-warn{background:var(--warn)}
.bar .seg-ok{background:#2f3542}
.bar.clean .seg-ok{background:color-mix(in srgb,var(--ok) 45%,#2f3542)}
.bar:hover .seg-ok,.bar.active .seg-ok{background:#3d4554}
.bar.clean:hover .seg-ok,.bar.clean.active .seg-ok{background:color-mix(in srgb,var(--ok) 60%,#2f3542)}
.bar.active{outline:1px solid var(--dim);outline-offset:1px}
#axis{display:flex;gap:6px;margin-top:6px}
.axis-lab{flex:1;text-align:center;color:var(--faint);font-size:13px;white-space:nowrap;overflow:hidden}
.axis-lab.active{color:var(--text)}

/* controls */
#controls{padding:12px 0;display:flex;gap:8px;flex-wrap:wrap;align-items:center;border-bottom:1px solid var(--line)}
.chip{background:var(--panel);border:1px solid var(--line);border-radius:20px;padding:8px 16px;font-size:14px;color:var(--dim);cursor:pointer;font-family:inherit}
.chip .n{color:var(--faint);margin-left:4px}
.chip.on{color:var(--text);border-color:var(--dim)}
.chip.on.c-err{color:var(--err);border-color:var(--err)}
.chip.on.c-warn{color:var(--warn);border-color:var(--warn)}
.chip.on.c-ok{color:var(--ok);border-color:var(--ok)}
.chip.on.c-info{color:var(--info);border-color:var(--info)}
#search{background:var(--panel);border:1px solid var(--line);border-radius:6px;color:var(--text);padding:9px 13px;font-size:14px;font-family:inherit;width:220px;margin-left:auto}
#search:focus{outline:none;border-color:var(--dim)}
#clearDay{display:none;font-size:12px;color:var(--info);cursor:pointer;background:none;border:none;font-family:inherit}

/* rows */
#list{padding:8px 0 48px}
.day-head{color:var(--text);font-size:16px;font-weight:500;padding:22px 0 8px;border-bottom:1px solid var(--line);margin-bottom:4px}
.sev-head{font-size:15px;font-weight:500;padding:12px 0 5px 4px}
.sev-err{color:var(--err)}.sev-warn{color:var(--warn)}.sev-info{color:var(--info)}
.row{display:grid;grid-template-columns:58px 12px 1fr;gap:10px;padding:9px 8px;border-radius:6px;cursor:pointer;align-items:baseline}
.row:hover{background:var(--panel)}
.row.open{background:var(--panel2)}
.time{color:var(--faint);font-size:14px}
.dot{width:8px;height:8px;border-radius:50%;align-self:center}
.d-err{background:var(--err)}.d-warn{background:var(--warn)}.d-ok{background:var(--ok)}.d-info{background:var(--info)}
.title{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.row.open .title{white-space:normal}
.src{color:var(--dim);font-size:14px;margin-left:10px}
.msg{grid-column:3;color:var(--dim);font-size:14px;line-height:1.5;padding:6px 0 2px;white-space:pre-wrap;display:none;word-break:break-word}
.row.open .msg{display:block}
#empty{color:var(--faint);padding:40px 0;text-align:center;display:none}
@media (max-width:600px){
  #timeline,#controls,#list{padding-left:14px;padding-right:14px}
  #search{width:100%;margin-left:0}
  .row{grid-template-columns:44px 10px 1fr}
  #summaryView,#sysView,#drivesView,#netView,#securityView,#appsView,#dumpsView,#memoryView,#gpuView,#processesView,#extensionsView,#updatesView{padding:24px 16px 48px}
  #pageTitle{font-size:28px;padding:24px 16px 0}
}
</style>
</head>
<body class="tab-summary">
<div id="appShell">
<aside id="sidebar">
  <div id="brand">PCHH Triage</div>
  <label id="drop">Open another CSV<input type="file" accept=".csv" hidden></label>
  <nav id="tabs">
    <div class="nav-group">
      <div class="nav-group-title static"><span>Overview</span></div>
      <div class="nav-group-items">
        <button class="tab on" data-tab="summary">System Summary</button>
      </div>
    </div>
    <div class="nav-group">
      <div class="nav-group-title"><span>Diagnostics</span><span class="chev">&#9660;</span></div>
      <div class="nav-group-items">
        <button class="tab" data-tab="rel">Reliability History</button>
        <button class="tab" data-tab="sys">Event Viewer</button>
        <button class="tab" data-tab="dumps" id="dumpsTab" style="display:none">Memory Dumps</button>
      </div>
    </div>
    <div class="nav-group">
      <div class="nav-group-title"><span>Hardware</span><span class="chev">&#9660;</span></div>
      <div class="nav-group-items">
        <button class="tab" data-tab="drives">Drives</button>
        <button class="tab" data-tab="gpu">GPU and Display(s)</button>
        <button class="tab" data-tab="memory">Memory (RAM)</button>
        <button class="tab" data-tab="net">Network</button>
      </div>
    </div>
    <div class="nav-group">
      <div class="nav-group-title"><span>System</span><span class="chev">&#9660;</span></div>
      <div class="nav-group-items">
        <button class="tab" data-tab="security">Security</button>
        <button class="tab" data-tab="processes">Running Processes</button>
        <button class="tab" data-tab="apps">Installed Apps</button>
        <button class="tab" data-tab="updates">Windows Updates</button>
        <button class="tab" data-tab="extensions">Browser Extensions</button>
      </div>
    </div>
  </nav>
  <div id="sideFoot"><span id="pageFoot"></span></div>
</aside>
<main id="content">

<h1 id="pageTitle">PCHH Triage <span id="pageTitleSub">- System Summary</span></h1>

<div id="summaryView" class="view">
  <div class="spec-section"><h2>System Specs</h2><div id="summary"></div><div id="specsContent"></div></div>
  <div class="spec-section"><h2>General Notes</h2><div id="notesBody"></div></div>
</div>

<div id="relView" class="view">
<div id="timeline">
  <div id="tlHead"><span id="tlRange" class="mono"></span><span id="tlHint">Click a bar to see that day's events. Use the arrows to move back a week.</span></div>
  <div id="tl-inner">
    <button id="tlPrev" class="tl-nav" title="Earlier">&#8249;</button>
    <div id="tl-main"><div id="bars"></div><div id="axis"></div></div>
    <button id="tlNext" class="tl-nav" title="Later">&#8250;</button>
  </div>
</div>

<div id="controls">
  <button class="chip c-err" data-cat="err">&#10060;&#65038; Critical events<span class="n"></span></button>
  <button class="chip c-warn" data-cat="warn">&#9888;&#65038; Warnings<span class="n"></span></button>
  <button class="chip c-info" data-cat="info">&#8505;&#65038; Informational events<span class="n"></span></button>
  <button id="clearDay"></button>
  <input id="search" type="text" placeholder="Search product or message…">
</div>

<div id="list"></div>
<div id="empty">No events match.</div>
</div>

<div id="sysView" class="view"></div>
<div id="drivesView" class="view"></div>
<div id="gpuView" class="view"></div>
<div id="memoryView" class="view"></div>
<div id="netView" class="view"></div>
<div id="securityView" class="view"></div>
<div id="processesView" class="view"></div>
<div id="appsView" class="view"></div>
<div id="updatesView" class="view"></div>
<div id="extensionsView" class="view"></div>
<div id="dumpsView" class="view"></div>

</main>
</div>

<script>
const RAW = /*__DATA__*/[];
const SPECS = /*__SPECS__*/"";
const DUMPS = /*__DUMPS__*/[];
const SYSEVT = /*__SYSEVT__*/[];
const SMART = /*__SMART__*/[];
const DIRTY = /*__DIRTY__*/[];
const DISKLAYOUT = /*__DISKLAYOUT__*/[];
const RAM = /*__RAM__*/[];
const GPUS = /*__GPUS__*/[];
const HAGS = /*__HAGS__*/null;
const WUHISTORY = /*__WUHISTORY__*/[];
const WINUPDATE = /*__WINUPDATE__*/null;
const MONS = /*__MONS__*/[];
const DISPLAYS = /*__DISPLAYS__*/[];
const PROCS = /*__PROCS__*/[];
const MEMUSE = /*__MEMUSE__*/null;
const NET = /*__NET__*/null;
const SECURITY = /*__SECURITY__*/null;
const HOTFIXES = /*__HOTFIXES__*/[];
const DEVERR = /*__DEVERR__*/[];
const AUDIO = /*__AUDIO__*/null;
const VER = /*__VER__*/"";
const GEN = /*__GEN__*/"";

// --- parsing / classification ---
function parseDate(s){
  // DD/MM/YYYY HH:MM:SS or MM/DD/YYYY — detect: first field >12 means DD first
  const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})[ ,]+(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)?/i);
  if(!m) return null;
  let [,a,b,y,h,mi,se,ap] = m;
  a=+a;b=+b;h=+h;
  let day=a, mon=b;
  if(a<=12 && b>12){ day=b; mon=a; }
  if(ap){ if(/pm/i.test(ap)&&h<12)h+=12; if(/am/i.test(ap)&&h===12)h=0; }
  return new Date(+y, mon-1, day, h, +mi, +se);
}
function classify(r){
  const src=r.s, msg=(r.m||'').toLowerCase();
  if(src==='Application Error'||src==='Windows Error Reporting'||/bugcheck/i.test(src)) return 'err';
  if(src==='EventLog') return /unexpected/.test(msg)?'err':'warn';
  if(/fail|error status: 1|not.*success/i.test(msg) && !/status: 0/.test(msg)) return 'warn';
  return 'info';
}
const CATNAMES={err:'Critical events',warn:'Warnings',info:'Informational events'};

let events=[], state={cats:new Set(['err','warn','info']), q:'', day:null, tlEnd:null};
const TL_WIN=7;

function load(raw){
  events = raw.map(r=>{
    const d=parseDate(r.t);
    return {...r, d, cat:classify(r), dayKey:d?d.toISOString().slice(0,10):'?'};
  }).filter(e=>e.d).sort((a,b)=>b.d-a.d);
  state.day=null;
  state.tlEnd=null;
  render();
}

function fmtDay(k){const d=new Date(k);return d.toLocaleDateString('en-GB',{weekday:'short',day:'numeric',month:'short'});}
function fmtTime(d){return d.toLocaleTimeString('en-GB',{hour:'2-digit',minute:'2-digit'});}

function render(){
  // counts per category (unfiltered by cat, filtered by search+day)
  const base = events.filter(e=>
    (!state.day||e.dayKey===state.day) &&
    (!state.q || (e.p+' '+e.m+' '+e.s).toLowerCase().includes(state.q)));
  document.querySelectorAll('.chip[data-cat]').forEach(c=>{
    const cat=c.dataset.cat;
    c.querySelector('.n').textContent=base.filter(e=>e.cat===cat).length;
    c.classList.toggle('on',state.cats.has(cat));
  });

  const shown = base.filter(e=>state.cats.has(e.cat));


  // timeline: continuous calendar days, windowed to 7 with scroll
  const allDays=[];
  if(events.length){
    const lo=new Date(events[events.length-1].dayKey), hi=new Date(events[0].dayKey);
    for(let d=new Date(lo); d<=hi; d.setDate(d.getDate()+1)) allDays.push(d.toISOString().slice(0,10));
  }
  if(state.tlEnd===null||state.tlEnd>allDays.length-1) state.tlEnd=allDays.length-1;
  if(state.tlEnd<Math.min(TL_WIN,allDays.length)-1) state.tlEnd=Math.min(TL_WIN,allDays.length)-1;
  const winStart=Math.max(0,state.tlEnd-TL_WIN+1);
  const days=allDays.slice(winStart,state.tlEnd+1);
  const byDay={};
  allDays.forEach(k=>byDay[k]={err:0,warn:0,rest:0});
  events.forEach(e=>{
    if(state.q && !(e.p+' '+e.m+' '+e.s).toLowerCase().includes(state.q)) return;
    const b=byDay[e.dayKey]; if(!b) return;
    if(e.cat==='err')b.err++; else if(e.cat==='warn')b.warn++; else b.rest++;
  });
  const max=Math.max(1,...days.map(k=>byDay[k].err+byDay[k].warn+byDay[k].rest));
  document.getElementById('tlPrev').disabled = winStart===0;
  document.getElementById('tlNext').disabled = state.tlEnd>=allDays.length-1;
  const bars=document.getElementById('bars');
  bars.innerHTML='';
  days.forEach(k=>{
    const b=byDay[k], tot=b.err+b.warn+b.rest;
    const bar=document.createElement('div');
    bar.className='bar'+(state.day===k?' active':'')+((b.err+b.warn)===0?' clean':'');
    bar.title=fmtDay(k)+' — '+tot+' event'+(tot===1?'':'s')+(b.err?' ('+b.err+' critical)':'');
    if(!tot){const s=document.createElement('div');s.className='seg-ok';s.style.height='3px';s.style.opacity='.45';bar.appendChild(s);}
    const h=x=>Math.round(x/max*64);
    if(b.rest){const s=document.createElement('div');s.className='seg-ok';s.style.height=Math.max(tot?3:0,h(b.rest))+'px';bar.appendChild(s);}
    if(b.warn){const s=document.createElement('div');s.className='seg-warn';s.style.height=Math.max(8,h(b.warn))+'px';bar.appendChild(s);}
    if(b.err){const s=document.createElement('div');s.className='seg-err';s.style.height=Math.max(8,h(b.err))+'px';bar.appendChild(s);}
    bar.onclick=()=>{state.day=state.day===k?null:k;render();};
    bars.appendChild(bar);
  });
  const axis=document.getElementById('axis');
  axis.innerHTML=days.map(k=>{
    const d=new Date(k);
    return '<span class="axis-lab'+(state.day===k?' active':'')+'">'+d.toLocaleDateString('en-GB',{weekday:'short',day:'numeric'})+'</span>';
  }).join('');
  const rEl=document.getElementById('tlRange');
  if(days.length){
    const a=new Date(days[0]),b2=new Date(days[days.length-1]);
    rEl.textContent=a.toLocaleDateString('en-GB',{day:'numeric',month:'short'})+' \u2013 '+b2.toLocaleDateString('en-GB',{day:'numeric',month:'short',year:'numeric'});
  }

  const cd=document.getElementById('clearDay');
  cd.style.display=state.day?'inline':'none';
  cd.textContent=state.day?('✕ '+fmtDay(state.day)):'';

  // list grouped by day
  const list=document.getElementById('list');
  list.innerHTML='';
  const dayGroups=new Map();
  shown.forEach(e=>{
    if(!dayGroups.has(e.dayKey))dayGroups.set(e.dayKey,{err:[],warn:[],info:[]});
    dayGroups.get(e.dayKey)[e.cat].push(e);
  });
  const ICONS={err:'\u274C\uFE0E',warn:'\u26A0\uFE0E',info:'\u2139\uFE0E'};
  dayGroups.forEach((groups,dayKey)=>{
    const h=document.createElement('div');h.className='day-head';
    h.textContent='Reliability details for: '+fmtDay(dayKey);
    list.appendChild(h);
    ['err','warn','info'].forEach(cat=>{
      const evs=groups[cat];
      if(!evs.length)return;
      const sh=document.createElement('div');sh.className='sev-head sev-'+cat;
      sh.textContent=ICONS[cat]+' '+CATNAMES[cat]+(evs.length>1?' ('+evs.length+')':'');
      list.appendChild(sh);
      evs.forEach(e=>{
        const row=document.createElement('div');row.className='row';
        row.innerHTML='<span class="time mono">'+fmtTime(e.d)+'</span>'+
          '<span class="dot d-'+e.cat+'"></span>'+
          '<span class="title">'+esc(e.p||'(unnamed)')+'<span class="src">'+summary(e)+'</span></span>'+
          '<div class="msg mono">'+esc(e.m)+'</div>';
        row.onclick=()=>row.classList.toggle('open');
        list.appendChild(row);
      });
    });
  });
  document.getElementById('empty').style.display=shown.length?'none':'block';
}
function summary(e){
  const msg=(e.m||'').toLowerCase();
  if(e.cat==='err'){
    if(/faulting application/.test(msg))return 'Stopped working';
    if(/unexpected/.test(msg))return 'Windows was not properly shut down';
    return 'Critical event';
  }
  if(e.s==='Microsoft-Windows-WindowsUpdateClient')
    return /success/.test(msg)?'Successful Windows Update':'Windows Update';
  if(e.s==='MsiInstaller'){
    if(/installed the product/.test(msg))return 'Successful application installation';
    if(/removed the product/.test(msg))return 'Successful application removal';
    if(/reconfigured/.test(msg))return 'Successful application reconfiguration';
    return 'Application event';
  }
  return esc(e.s);
}
function esc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}

document.querySelectorAll('.chip[data-cat]').forEach(c=>c.onclick=()=>{
  const cat=c.dataset.cat;
  state.cats.has(cat)?state.cats.delete(cat):state.cats.add(cat);
  render();
});
document.getElementById('clearDay').onclick=()=>{state.day=null;render();};
document.getElementById('tlPrev').onclick=()=>{state.tlEnd=Math.max(TL_WIN-1,state.tlEnd-TL_WIN);render();};
document.getElementById('tlNext').onclick=()=>{state.tlEnd=state.tlEnd+TL_WIN;render();};
document.getElementById('search').oninput=e=>{state.q=e.target.value.toLowerCase();render();};

// CSV loading (drop or picker) for future exports
function parseCSV(text){
  const rows=[];let cur=[''],inQ=false,i=0;
  for(;i<text.length;i++){
    const c=text[i];
    if(inQ){
      if(c==='"'){ if(text[i+1]==='"'){cur[cur.length-1]+='"';i++;} else inQ=false; }
      else cur[cur.length-1]+=c;
    } else {
      if(c==='"')inQ=true;
      else if(c===',')cur.push('');
      else if(c==='\n'||c==='\r'){ if(cur.length>1||cur[0]!==''){rows.push(cur);cur=[''];} }
      else cur[cur.length-1]+=c;
    }
  }
  if(cur.length>1||cur[0]!=='')rows.push(cur);
  const head=rows.shift().map(h=>h.replace(/^\ufeff/,''));
  const ix=n=>head.findIndex(h=>h.toLowerCase()===n);
  const [t,s,e,p,m]=['timegenerated','sourcename','eventidentifier','productname','message'].map(ix);
  return rows.map(r=>({t:r[t],s:r[s],e:r[e],p:r[p],m:r[m]}));
}
function handleFile(f){
  const rd=new FileReader();
  rd.onload=()=>{try{load(parseCSV(rd.result));}catch(err){alert('Could not parse that CSV: '+err.message);}};
  rd.readAsText(f);
}
document.querySelector('#drop input').onchange=e=>e.target.files[0]&&handleFile(e.target.files[0]);
['dragover','dragenter'].forEach(ev=>document.addEventListener(ev,e=>{e.preventDefault();document.body.classList.add('dragging');}));
['dragleave','drop'].forEach(ev=>document.addEventListener(ev,e=>{e.preventDefault();document.body.classList.remove('dragging');}));
document.addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)handleFile(f);});

// --- specs parsing & rendering ---
function parseSpecs(text){
  const out={info:[],drives:[],programs:[]};
  if(!text||!text.trim())return out;
  const norm=text.replace(/\r/g,'');
  const [head, rest] = splitOnce(norm, /^Drive Information:\s*$/m);
  const [driveTxt, progTxt] = splitOnce(rest||'', /^Programs Installed:\s*$/m);
  head.split('\n').forEach(l=>{
    const m=l.match(/^([^:]+):\s?(.*)$/);
    if(m&&m[2]!=='')out.info.push([m[1].trim(),m[2].trim()]);
  });
  let cur=null;
  (driveTxt||'').split('\n').forEach(l=>{
    const m=l.match(/^([^:]+):\s?(.*)$/);
    if(!m)return;
    const k=m[1].trim(),v=m[2].trim();
    if(k==='Drive Label'){cur={};out.drives.push(cur);}
    if(cur)cur[k]=v;
  });
  (progTxt||'').split('\n').forEach(l=>{
    const t=l.trim();
    if(!t||t==='DisplayName'||/^-+$/.test(t))return;
    out.programs.push(t);
  });
  out.programs.sort((a,b)=>a.localeCompare(b,undefined,{sensitivity:'base'}));
  return out;
}
function splitOnce(text,re){
  const m=text.match(re);
  if(!m)return[text,''];
  return[text.slice(0,m.index),text.slice(m.index+m[0].length)];
}
function renderSpecs(){
  const sp=parseSpecs(SPECS);
  const v=document.getElementById('specsContent');
  if(!sp.info.length&&!sp.drives.length&&!sp.programs.length){
    v.innerHTML='';
    return;
  }
  let h='';
  if(sp.info.length){
    h+='<div class="spec-section" style="border-top:1px solid var(--line);padding-top:18px"><dl class="kv">';
    const SHOWN=['OS','OS Version','Build','System Uptime','CPU Name','GPU','Motherboard','Motherboard Manufacturer','BIOS Date','Ram Capacity','RAM Speed'];
    sp.info.filter(([k])=>!SHOWN.includes(k)).forEach(([k,val])=>{
      const off=/^(Secure Boot State|TPM Status)$/.test(k)&&/Disabled/i.test(val);
      h+='<dt>'+esc(k)+'</dt><dd'+(off?' class="flag-off"':'')+'>'+esc(val)+'</dd>';
    });
    h+='</dl></div>';
  }
  if(AUDIO&&AUDIO.playback){
    h+='<div class="spec-section"><h2>Audio</h2><dl class="kv"><dt>Default playback device</dt><dd>'+esc(AUDIO.playback)+'</dd></dl></div>';
  }
  if(DEVERR.length){
    h+='<div class="spec-section"><h2>Device Manager errors ('+DEVERR.length+')</h2><dl class="kv">';
    DEVERR.forEach(e=>{h+='<dt>'+esc(e.name)+'</dt><dd style="color:var(--err)">Error code '+esc(e.code)+'</dd>';});
    h+='</dl></div>';
  }
  let dh='';
  if(sp.drives.length){
    dh+='<div class="spec-section"><h2>Drives ('+sp.drives.length+')</h2><div class="drive-grid">';
    sp.drives.forEach(d=>{
      const total=parseFloat(d['Total Size (GB)'])||0, free=parseFloat(d['Free Space (GB)'])||0;
      const pctFree=total?Math.round(free/total*100):0, pctUsed=100-pctFree;
      const name=d['Drive Name']&&d['Drive Name']!=='No Name Found'?d['Drive Name']:'';
      dh+='<div class="drive"><h3>'+esc(d['Drive Label']||'?')+(name?' <span style="color:var(--dim);font-weight:400">'+esc(name)+'</span>':'')+
        (d['Windows Drive']==='True'?' <span style="color:var(--info);font-size:11px">Windows</span>':'')+'</h3>'+
        '<div class="sub">'+esc(d['Drive Type']||'Unknown')+' · '+esc(d['Drive Status']||'Unknown')+'</div>'+
        (DIRTY.some(v=>String(d['Drive Label']||'').toUpperCase().startsWith(v.toUpperCase()))?'<div style="color:var(--warn);font-size:13.5px;margin-bottom:6px">Dirty bit set</div>':'')+
        '<div class="meter'+(pctFree<15?' low':'')+'"><div style="width:'+pctUsed+'%"></div></div>'+
        '<div class="use mono">'+free.toFixed(0)+' GB free of '+total.toFixed(0)+' GB ('+pctFree+'%)</div></div>';
    });
    dh+='</div></div>';
  }
  if(DISKLAYOUT.length){
    dh+='<div class="spec-section"><h2>Disk layout</h2>';
    const TYPE_COLOR={'EFI System Partition':'var(--info)','Recovery':'var(--warn)','Recovery (MBR)':'var(--warn)','Microsoft Reserved':'var(--faint)','Data':'var(--ok)','System':'var(--dim)'};
    DISKLAYOUT.forEach(dk=>{
      const total=dk.partitions.reduce((a,p)=>a+p.sizeGB,0)||dk.sizeGB||1;
      dh+='<div class="drive" style="margin-bottom:14px"><h3>Disk '+dk.disk+' <span style="color:var(--dim);font-weight:400">'+esc(dk.style||'Unknown')+' \u00b7 '+dk.sizeGB+' GB</span></h3>';
      dh+='<div style="display:flex;height:22px;border-radius:6px;overflow:hidden;margin:10px 0;background:var(--panel2)">';
      dk.partitions.forEach(p=>{
        const pct=Math.max(1.5,(p.sizeGB/total*100));
        const col=TYPE_COLOR[p.type]||'var(--dim)';
        dh+='<div title="'+esc(p.type)+(p.letter?' ('+esc(p.letter)+')':'')+' \u2014 '+p.sizeGB+' GB" style="width:'+pct+'%;background:'+col+';border-right:1px solid var(--panel)"></div>';
      });
      dh+='</div><dl class="kv smart-kv">';
      dk.partitions.forEach(p=>{
        dh+='<dt><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:'+(TYPE_COLOR[p.type]||'var(--dim)')+';margin-right:6px"></span>'+esc(p.type)+(p.letter?' ('+esc(p.letter)+')':'')+'</dt><dd>'+p.sizeGB+' GB</dd>';
      });
      dh+='</dl></div>';
    });
    dh+='</div>';
  }
  if(SMART.length){
    dh+='<div class="spec-section"><h2>SMART data</h2><div class="drive-grid">';
    SMART.forEach(d=>{
      const bad=(d.health&&d.health!=='Healthy')||(+d.reu>0)||(+d.weu>0)||(+d.rl>0)||(+d.pend>0)||(+d.unc>0)||(+d.crc>0)||d.pf==='1';
      let rows='';
      const add=(l,val)=>{if(val!=='')rows+='<dt>'+l+'</dt><dd>'+esc(val)+'</dd>';};
      add('Health',d.health+(d.op&&d.op!=='OK'?' ('+d.op+')':''));
      add('Temperature',d.temp?d.temp+'\u00b0C'+(d.tmax?' (max '+d.tmax+'\u00b0C)':''):'');
      add('Power-on hours',d.hours);
      add('Wear',d.wear?d.wear+'%':'');
      add('Read errors (uncorrected)',d.reu);
      add('Read errors (corrected)',d.rec);
      add('Write errors (uncorrected)',d.weu);
      add('Write errors (corrected)',d.wec);
      add('Reallocated sectors',d.rl);
      add('Pending sectors',d.pend);
      add('Uncorrectable sectors',d.unc);
      add('UltraDMA CRC errors',d.crc);
      add('Command timeouts',d.cto);
      if(d.pf==='1')rows+='<dt style="color:var(--err)">Failure predicted</dt><dd style="color:var(--err)">Yes (drive self-report)</dd>';
      dh+='<div class="drive'+(bad?' smart-bad':'')+'"><h3>Disk '+esc(d.disk)+'<br><span style="color:var(--dim);font-weight:400;font-size:14px">'+esc(d.name)+'</span></h3>'+
        '<div class="sub">'+esc(d.media||'Unknown')+(d.bus?' \u00b7 '+esc(d.bus):'')+'</div>'+
        '<dl class="kv smart-kv">'+rows+'</dl></div>';
    });
    dh+='</div></div>';
    const alerts=[];
    SMART.forEach(d=>{
      const probs=smartProbs(d);
      if(probs.length)alerts.push('<li><span class="r" style="color:var(--err)">Disk '+esc(d.disk)+' ('+esc(d.name)+'): '+esc(probs.join(', '))+'</span></li>');
    });
    DIRTY.forEach(v=>alerts.push('<li><span style="color:var(--warn)">Volume '+esc(v)+' has its dirty bit set</span></li>'));
    dh+='<div class="spec-section" style="margin-top:26px"><h2>SMART alerts</h2>'+
      (alerts.length?'<ul class="notes">'+alerts.join('')+'</ul>'
       :'<div style="color:var(--ok)">\u2713 No SMART alerts \u2014 all disks report Healthy with no uncorrected errors.</div>')+'</div>';
  }
  v.innerHTML=h;
  document.getElementById('drivesView').innerHTML=dh||'<div class="spec-section"><h2>Drives</h2><div style="color:var(--faint)">No drive data embedded.</div></div>';
  renderAppsList(sp.programs);
  renderProcesses();
  renderExtensions();
  renderUpdates();
}
const PS_={q:'',page:1,key:'mem',dir:-1}, PG_={q:'',page:1};
let PROGS_ALL=[];
function renderProcesses(){
  const v=document.getElementById('processesView');
  if(!PROCS.length){
    v.innerHTML='<div class="spec-section"><h2>Running Processes</h2><div style="color:var(--faint)">No process data embedded.</div></div>';
    return;
  }
  v.innerHTML='<div class="spec-section"><h2>Running processes ('+PROCS.length+')</h2>'+
    '<input id="procSearch" type="text" placeholder="Filter processes\u2026">'+
    '<div class="proc-head">'+
    '<span class="sorth" data-key="name">Process<span class="arrow"></span></span>'+
    '<span class="sorth" data-key="cnt">Instances<span class="arrow"></span></span>'+
    '<span class="sorth" data-key="mem">Memory<span class="arrow"></span></span></div>'+
    '<div id="procList"></div><div class="pager" id="procPager"></div></div>';
  const pf=document.getElementById('procSearch');
  if(pf)pf.oninput=e=>{PS_.q=e.target.value.toLowerCase();PS_.page=1;renderProcList();};
  document.querySelectorAll('.sorth').forEach(hd=>hd.onclick=()=>{
    const k=hd.dataset.key;
    if(PS_.key===k)PS_.dir=-PS_.dir; else {PS_.key=k;PS_.dir=k==='name'?1:-1;}
    PS_.page=1;renderProcList();
  });
  renderProcList();
}
function renderAppsList(programs){
  PROGS_ALL=programs||[];
  const v=document.getElementById('appsView');
  v.innerHTML=PROGS_ALL.length?
    '<div class="spec-section"><h2>Installed programs ('+PROGS_ALL.length+')</h2>'+
    '<input id="progSearch" type="text" placeholder="Filter programs\u2026">'+
    '<div id="progList"></div><div class="pager" id="progPager"></div></div>'
    :'<div class="spec-section"><h2>Installed Apps</h2><div style="color:var(--faint)">No data embedded.</div></div>';
  const ps=document.getElementById('progSearch');
  if(ps)ps.oninput=e=>{PG_.q=e.target.value.toLowerCase();PG_.page=1;renderProgList();};
  renderProgList();
}
const WH_={page:1};
function renderUpdates(){
  const v=document.getElementById('updatesView');
  let h='';
  const w=WINUPDATE;
  if(w){
    h+='<div class="spec-section"><h2>Windows Update status</h2><dl class="kv">';
    h+='<dt>Pending Reboot?</dt><dd style="color:'+(w.pendingReboot?'var(--warn)':'var(--ok)')+'">'+(w.pendingReboot?'Yes':'No')+'</dd>';
    if(w.serviceStatus)h+='<dt>Windows Update service</dt><dd style="color:'+(w.serviceStatus==='Running'?'var(--ok)':'var(--warn)')+'">'+esc(w.serviceStatus)+'</dd>';
    h+='</dl></div>';
  }
  if(WUHISTORY.length){
    const failCount=WUHISTORY.filter(u=>u.result==='Failed'||u.result==='Cancelled').length;
    h+='<div class="spec-section"><h2>Recent update history ('+WUHISTORY.length+')</h2>';
    if(failCount)h+='<div style="color:var(--warn);font-size:14px;margin-bottom:12px">'+failCount+' update'+(failCount>1?'s':'')+' did not complete successfully</div>';
    h+='<dl class="kv" id="wuHistList"></dl><div class="pager" id="wuHistPager"></div></div>';
  }
  if(HOTFIXES.length){
    h+='<div class="spec-section"><h2>Installed updates ('+HOTFIXES.length+')</h2>'+
      '<input id="hfSearch" type="text" placeholder="Filter updates\u2026">'+
      '<dl class="kv" id="hfList"></dl><div class="pager" id="hfPager"></div></div>';
  }
  v.innerHTML=h||'<div class="spec-section"><h2>Windows Updates</h2><div style="color:var(--faint)">No update data embedded.</div></div>';
  const hf=document.getElementById('hfSearch');
  if(hf)hf.oninput=e=>{HF_.q=e.target.value.toLowerCase();HF_.page=1;renderHfList();};
  renderHfList();
  renderWuHistList();
}
function renderWuHistList(){
  const el=document.getElementById('wuHistList');if(!el)return;
  const SZ=15,pages=Math.max(1,Math.ceil(WUHISTORY.length/SZ));
  if(WH_.page>pages)WH_.page=pages;
  const slice=WUHISTORY.slice((WH_.page-1)*SZ,WH_.page*SZ);
  el.innerHTML=slice.map(u=>{
    const col=u.result==='Succeeded'?'var(--ok)':(u.result==='Failed'||u.result==='Cancelled')?'var(--err)':'var(--warn)';
    return '<dt>'+esc(u.date)+'</dt><dd>'+esc(u.title)+' <span style="color:'+col+'">('+esc(u.result)+')</span></dd>';
  }).join('')||'<dd style="color:var(--faint)">No update history.</dd>';
  pager(document.getElementById('wuHistPager'),WH_.page,pages,WUHISTORY.length,slice.length,g=>{WH_.page+=g;renderWuHistList();});
}
function renderExtensions(){
  const v=document.getElementById('extensionsView');
  if(!SECURITY||!SECURITY.extensions||!SECURITY.extensions.length){
    v.innerHTML='<div class="spec-section"><h2>Browser Extensions</h2><div style="color:var(--faint)">No browser extension data embedded.</div></div>';
    return;
  }
  const byBrowser={};
  SECURITY.extensions.forEach(e=>{(byBrowser[e.browser]=byBrowser[e.browser]||[]).push(e.name);});
  let h='<div class="spec-section"><h2>Browser extensions ('+SECURITY.extensions.length+')</h2><div class="drive-grid">';
  Object.keys(byBrowser).sort().forEach(b=>{
    const names=[...new Set(byBrowser[b])].sort((a,c)=>a.localeCompare(c,undefined,{sensitivity:'base'}));
    h+='<div class="drive"><h3>'+esc(b)+'</h3>'+
      '<div class="sub">'+names.length+' extension'+(names.length>1?'s':'')+'</div>'+
      '<div style="color:var(--dim);font-size:14px;line-height:1.8">'+names.map(esc).join('<br>')+'</div></div>';
  });
  h+='</div></div>';
  v.innerHTML=h;
}
function pager(el,page,pages,total,shown,onGo){
  if(!el)return;
  if(pages<=1){el.innerHTML=total?'<span class="pg-info">'+shown+' of '+total+'</span>':'';return;}
  el.innerHTML='<button class="pg-btn" '+(page<=1?'disabled':'')+' data-go="-1">\u2039 Prev</button>'+
    '<span class="pg-info">Page '+page+' of '+pages+' \u00b7 '+total+' items</span>'+
    '<button class="pg-btn" '+(page>=pages?'disabled':'')+' data-go="1">Next \u203a</button>';
  el.querySelectorAll('.pg-btn').forEach(b=>b.onclick=()=>onGo(+b.dataset.go));
}
function renderProcList(){
  const el=document.getElementById('procList');if(!el)return;
  let rows=PROCS.filter(p=>!PS_.q||p.name.toLowerCase().includes(PS_.q));
  const k=PS_.key,d=PS_.dir;
  rows=rows.slice().sort((a,b)=>k==='name'?d*a.name.localeCompare(b.name,undefined,{sensitivity:'base'}):d*((+a[k]||0)-(+b[k]||0)));
  const SZ=50,pages=Math.max(1,Math.ceil(rows.length/SZ));
  if(PS_.page>pages)PS_.page=pages;
  const slice=rows.slice((PS_.page-1)*SZ,PS_.page*SZ);
  el.innerHTML=slice.map(p=>'<div class="proc-row"><span>'+esc(p.name)+'</span><span>'+esc(String(p.cnt))+'</span><span class="mono">'+esc(String(p.mem))+' MB</span></div>').join('')||'<div style="color:var(--faint);padding:10px 4px">No matches.</div>';
  document.querySelectorAll('.sorth').forEach(hd=>{
    hd.querySelector('.arrow').textContent=hd.dataset.key===PS_.key?(PS_.dir>0?' \u25b2':' \u25bc'):'';
  });
  pager(document.getElementById('procPager'),PS_.page,pages,rows.length,slice.length,g=>{PS_.page+=g;renderProcList();});
}
function renderProgList(){
  const el=document.getElementById('progList');if(!el)return;
  const rows=PROGS_ALL.filter(p=>!PG_.q||p.toLowerCase().includes(PG_.q));
  const SZ=60,pages=Math.max(1,Math.ceil(rows.length/SZ));
  if(PG_.page>pages)PG_.page=pages;
  const slice=rows.slice((PG_.page-1)*SZ,PG_.page*SZ);
  el.innerHTML=slice.map(p=>'<div>'+esc(p)+'</div>').join('')||'<div style="color:var(--faint)">No matches.</div>';
  pager(document.getElementById('progPager'),PG_.page,pages,rows.length,slice.length,g=>{PG_.page+=g;renderProgList();});
}
const HF_={q:'',page:1};
function renderHfList(){
  const el=document.getElementById('hfList');if(!el)return;
  const rows=HOTFIXES.filter(h=>!HF_.q||(h.id+' '+h.desc).toLowerCase().includes(HF_.q));
  const SZ=15,pages=Math.max(1,Math.ceil(rows.length/SZ));
  if(HF_.page>pages)HF_.page=pages;
  const slice=rows.slice((HF_.page-1)*SZ,HF_.page*SZ);
  el.innerHTML=slice.map(h=>'<dt>'+esc(h.date||'')+'</dt><dd>'+esc(h.id)+(h.desc?' <span style="color:var(--faint)">'+esc(h.desc)+'</span>':'')+'</dd>').join('')||'<dd style="color:var(--faint)">No matches.</dd>';
  pager(document.getElementById('hfPager'),HF_.page,pages,rows.length,slice.length,g=>{HF_.page+=g;renderHfList();});
}
function smartProbs(d){
  const probs=[];
  if(d.pf==='1')probs.push('drive predicts its own failure');
  if(d.health&&d.health!=='Healthy')probs.push('health: '+d.health);
  if(+d.rl>0)probs.push(d.rl+' reallocated sectors');
  if(+d.pend>0)probs.push(d.pend+' pending sectors');
  if(+d.unc>0)probs.push(d.unc+' uncorrectable sectors');
  if(+d.crc>0)probs.push(d.crc+' UltraDMA CRC errors');
  if(+d.reu>0)probs.push(d.reu+' uncorrected read errors');
  if(+d.weu>0)probs.push(d.weu+' uncorrected write errors');
  return probs;
}
function friendlyMedia(m){
  const map={ '802.3':'Ethernet (802.3)', 'Native 802.11':'Wi-Fi (802.11)', '802.11':'Wi-Fi (802.11)', 'Bluetooth':'Bluetooth', 'WiMax':'WiMAX', 'Unspecified':'' };
  return map.hasOwnProperty(m)?map[m]:m;
}
function friendlyDriver(gpuName,ver,radeon){
  if(!ver)return '';
  const n=(gpuName||'').toLowerCase();
  if(/nvidia|geforce|quadro|rtx|gtx/.test(n)){
    const digits=ver.replace(/\./g,'');
    if(digits.length>=5){
      const five=digits.slice(-5);
      return five.slice(0,3)+'.'+five.slice(3)+' <span style="color:var(--faint)">('+ver+')</span>';
    }
  }
  if(/\bintel\b|\barc\b|iris|uhd/.test(n)){
    const parts=ver.split('.');
    if(parts.length>=4)return parts[2]+'.'+parts[3]+' <span style="color:var(--faint)">('+ver+')</span>';
  }
  if(/amd|radeon/.test(n)&&radeon){
    return radeon+' <span style="color:var(--faint)">('+ver+')</span>';
  }
  return ver;
}
function specVal(info,key){const f=info.find(([k])=>k===key);return f?f[1]:null;}
function renderSummary(){
  const sp=parseSpecs(SPECS);
  const el=document.getElementById('summary');
  const pairs=[];
  const os=specVal(sp.info,'OS'), build=specVal(sp.info,'Build'), up=specVal(sp.info,'System Uptime');
  const WINVER={ '26200':'25H2','26100':'24H2','22631':'23H2','22621':'22H2','22000':'21H2','19045':'22H2','19044':'21H2' };
  if(os){
    const bMajor=build?build.split('.')[0]:'';
    const fv=WINVER[bMajor];
    pairs.push(['OS', esc(os.replace('Microsoft ',''))+(fv?' '+fv:'')+(build?' <span style="color:var(--faint)">(build '+esc(build)+')</span>':'')]);
  }
  if(up)pairs.push(['System uptime', esc(up.replace(/ days?/,'d').replace(/ hours?/,'h').replace(/ minutes?/,'m').replace(/,/g,''))]);
  const cpu=specVal(sp.info,'CPU Name');
  if(cpu)pairs.push(['CPU', esc(cpu.trim())]);
  if(GPUS.length||DISPLAYS.length){
    const gpuNames=[...new Set(GPUS.length?GPUS.map(g=>g.name):DISPLAYS.map(d=>d.gpu))];
    pairs.push(['GPU'+(gpuNames.length>1?'s':''), gpuNames.map(esc).join(', ')+' <span style="color:var(--faint)">(see Hardware \u203a GPU)</span>']);
  } else {
    const gpu=specVal(sp.info,'GPU');
    if(gpu)pairs.push(['GPU', esc(gpu)+' <span style="color:var(--faint)">(see Hardware \u203a GPU)</span>']);
  }
  const mb=specVal(sp.info,'Motherboard'), mbMfr=specVal(sp.info,'Motherboard Manufacturer');
  if(mb)pairs.push(['Motherboard', esc(((mbMfr||'').replace(/ASUSTeK COMPUTER INC\./i,'ASUS').replace(/Micro-Star International.*/i,'MSI').replace(/Gigabyte Technology.*/i,'Gigabyte')+' '+mb).trim())]);
  const bdate=specVal(sp.info,'BIOS Date');
  if(bdate)pairs.push(['BIOS date', esc(bdate.replace(/\s+\d{1,2}:\d{2}(:\d{2})?(\s*[AP]M)?$/i,''))]);
  if(RAM.length){
    const totGB=RAM.reduce((a,x)=>a+(+x.cap||0),0);
    const conf=[...new Set(RAM.map(m=>m.conf).filter(Boolean))].join('/');
    pairs.push(['Memory', totGB+' GB total'+(conf?' @ '+esc(conf)+' MT/s':'')+' <span style="color:var(--faint)">(see Hardware \u203a Memory)</span>']);
  } else {
    const rc=specVal(sp.info,'Ram Capacity');
    if(rc)pairs.push(['Memory', esc(rc)]);
  }
  if(MEMUSE&&MEMUSE.pt){
    const pct=Math.round(MEMUSE.pu/MEMUSE.pt*100);
    pairs.push(['Memory used', MEMUSE.pu.toFixed(1)+' / '+MEMUSE.pt.toFixed(1)+' GB ('+pct+'%)'+
      ' <span style="color:var(--faint)">at time of capture</span>']);
    if(MEMUSE.ct)pairs.push(['Commit charge', MEMUSE.cu.toFixed(1)+' / '+MEMUSE.ct.toFixed(1)+' GB ('+Math.round(MEMUSE.cu/MEMUSE.ct*100)+'%)']);
  }
  const crashes=events.filter(e=>e.cat==='err'&&/faulting application/i.test(e.m)).length;
  const shutdowns=events.filter(e=>/unexpected/i.test(e.m)&&e.s==='EventLog').length;
  const notes=[];
  notes.push(crashes?'<span class="r"><b>'+crashes+'</b> Application crash'+(crashes>1?'es':'')+'</span>':'<span class="g">No application crashes</span>');
  // Unexpected shutdowns: reliability history (6008-derived) and Kernel-Power 41 record the
  // same incident. Report one merged line, using the larger count if they disagree.
  const kp41ev=SYSEVT.filter(r=>String(r.id)==='41');
  const shutdownCount=Math.max(shutdowns,kp41ev.length);
  if(shutdownCount){
    const BC_NAMES={ '278':'0x116 VIDEO_TDR_FAILURE','279':'0x117 VIDEO_TDR_TIMEOUT_DETECTED','281':'0x119 VIDEO_SCHEDULER_INTERNAL_ERROR','321':'0x141 VIDEO_ENGINE_TIMEOUT_DETECTED','322':'0x142 VIDEO_TDR_APPLICATION_BLOCKED' };
    const bcs=[...new Set(kp41ev.map(r=>String(r.bc||'')).filter(b=>b&&b!=='0'))];
    let detail;
    if(bcs.length){
      detail='bugcheck '+bcs.map(b=>BC_NAMES[b]||('code '+b)).join(', ');
    }else if(kp41ev.length){
      detail='no bugcheck \u2014 power loss, hard reset or hang';
    }else{
      detail='reliability history only \u2014 outside event log window';
    }
    notes.push('<span class="r"><b>'+shutdownCount+'</b> Unexpected shutdown'+(shutdownCount>1?'s':'')+'</span> <span style="color:var(--faint)">('+esc(detail)+')</span>');
  }else{
    notes.push('<span class="g">No unexpected shutdowns</span>');
  }
  if(DUMPS.length)notes.push('<span class="y"><b>'+DUMPS.length+'</b> Memory dump'+(DUMPS.length>1?'s':'')+' collected</span> <span style="color:var(--faint)">(in zip)</span>');
  const wheaFatal=SYSEVT.filter(r=>/WHEA/i.test(r.prov)&&['18','46'].includes(String(r.id))).length;
  if(wheaFatal)notes.push('<span class="r"><b>'+wheaFatal+'</b> Fatal hardware error'+(wheaFatal>1?'s':'')+' (WHEA)</span>');
  SMART.forEach(d=>{
    const probs=smartProbs(d);
    if(probs.length)notes.push('<span class="r">Disk '+esc(d.disk)+' ('+esc(d.name)+'): '+esc(probs.join(', '))+'</span>');
  });
  DIRTY.forEach(v=>notes.push('<span class="y">Volume '+esc(v)+' has its dirty bit set</span>'));
  if(DEVERR.length)notes.push('<span class="y"><b>'+DEVERR.length+'</b> device'+(DEVERR.length>1?'s':'')+' showing errors in Device Manager</span>');
  const sysDisk=DISKLAYOUT.find(dk=>dk.partitions.some(p=>p.letter==='C:'));
  if(sysDisk&&sysDisk.style&&sysDisk.style.toUpperCase()==='MBR')notes.push('<span class="y">System disk uses MBR partitioning (Secure Boot requires GPT)</span>');
  if(WINUPDATE&&WINUPDATE.pendingReboot)notes.push('<span class="y">System has a pending reboot (Windows Update or servicing)</span>');
  if(WINUPDATE&&WINUPDATE.serviceStatus&&WINUPDATE.serviceStatus!=='Running')notes.push('<span class="y">Windows Update service is '+esc(WINUPDATE.serviceStatus)+'</span>');
  const wuFails=WUHISTORY.filter(u=>u.result==='Failed'||u.result==='Cancelled').length;
  if(wuFails)notes.push('<span class="y"><b>'+wuFails+'</b> Windows Update'+(wuFails>1?'s':'')+' did not complete successfully (see Windows Updates tab)</span>');
  if(RAM.length){
    const slow=RAM.filter(m=>m.rated&&m.conf&&+m.conf<+m.rated);
    if(slow.length)notes.push('<span class="y">RAM configured at '+esc(slow[0].conf)+' MT/s, rated '+esc(slow[0].rated)+' MT/s</span>');
  }
  // Known software flags: anti-cheat/kernel drivers, OC & monitoring tools, RGB/peripheral suites, bloatware/PUPs
  const SOFT_FLAGS=[
    {re:/riot vanguard/i,        label:'Riot Vanguard',              grp:'ac'},
    {re:/easy anti-?cheat/i,     label:'Easy Anti-Cheat',            grp:'ac'},
    {re:/battleye/i,             label:'BattlEye',                   grp:'ac'},
    {re:/faceit anti-?cheat|faceit ac/i, label:'FACEIT AC',          grp:'ac'},
    {re:/msi afterburner/i,      label:'MSI Afterburner',            grp:'oc'},
    {re:/rtss|rivatuner/i,       label:'RTSS (RivaTuner Statistics)',grp:'oc'},
    {re:/intel.*extreme tuning|intel\(r\) xtu/i, label:'Intel XTU', grp:'oc'},
    {re:/ryzen master/i,         label:'AMD Ryzen Master',           grp:'oc'},
    {re:/corsair icue/i,         label:'Corsair iCUE',               grp:'periph'},
    {re:/razer synapse/i,        label:'Razer Synapse',              grp:'periph'},
    {re:/(logitech|logi) g ?hub/i, label:'Logitech G HUB',           grp:'periph'},
    {re:/armoury crate/i,        label:'ASUS Armoury Crate',         grp:'periph'},
    {re:/mystic light/i,         label:'MSI Mystic Light',           grp:'periph'},
    {re:/aura sync/i,            label:'ASUS Aura Sync',             grp:'periph'},
    {re:/mcafee/i,               label:'McAfee',                     grp:'bloat'},
    {re:/norton (360|security)/i,label:'Norton 360',                 grp:'bloat'},
    {re:/wildtangent/i,          label:'WildTangent Games',          grp:'bloat'},
    {re:/advanced systemcare|driver booster|iobit/i, label:'IObit utilities', grp:'bloat'},
    {re:/reimage|restoro/i,      label:'Restoro/Reimage',            grp:'bloat'},
    {re:/pc cleaner pro|mycleanpc|pc healthboost|systweak/i, label:'PC "cleaner" utility', grp:'bloat'},
    {re:/driverfix|smart driver care|driver updater/i, label:'Third-party driver updater', grp:'bloat'},
    {re:/nzxt cam/i,             label:'NZXT CAM',                   grp:'periph'},
    {re:/msi dragon center|dragon center/i, label:'MSI Dragon Center', grp:'periph'},
    {re:/nahimic/i,              label:'Nahimic Audio',              grp:'audio'},
    {re:/nvidia geforce experience/i, label:'GeForce Experience',    grp:'audio'},
    {re:/xbox game bar|gaming services/i, label:'Xbox Game Bar',     grp:'audio'},
    {re:/streamlabs/i,           label:'Streamlabs OBS',             grp:'audio'},
    {re:/hola vpn/i,             label:'Hola VPN',                   grp:'net'},
    {re:/killer network|killer control center/i, label:'Killer Network Manager', grp:'net'},
  ];
  const GRP_NAME={ac:'Anti-cheat / kernel driver',oc:'Overclock / monitoring tool',periph:'RGB / peripheral suite',audio:'Audio / overlay software',net:'Network software',bloat:'Potential bloatware/PUP'};
  const foundSoft={};
  (sp.programs||[]).forEach(p=>{
    SOFT_FLAGS.forEach(f=>{ if(f.re.test(p)){ (foundSoft[f.grp]=foundSoft[f.grp]||new Set()).add(f.label); } });
  });
  const avStr=specVal(sp.info,'Antivirus');
  if(avStr){
    const avList=avStr.split(',').map(s=>s.trim()).filter(Boolean);
    if(avList.length>1)notes.push('<span class="y">Multiple real-time antivirus products active: '+esc(avList.join(', '))+'</span>');
  }
  Object.keys(foundSoft).forEach(grp=>{
    const items=[...foundSoft[grp]].sort().join(', ');
    notes.push('<span class="'+(grp==='bloat'?'y':'')+'"><span class="slabel">'+GRP_NAME[grp]+':</span> '+esc(items)+'</span>');
  });

  if(SECURITY){
    if(SECURITY.defender&&SECURITY.defender.rtp!=='True')notes.push('<span class="r">Windows Defender real-time protection is disabled</span>');
    if(SECURITY.firewall&&SECURITY.firewall.some(f=>f.enabled!=='True')){
      const off=SECURITY.firewall.filter(f=>f.enabled!=='True').map(f=>f.profile);
      notes.push('<span class="r">Firewall disabled on: '+esc(off.join(', '))+'</span>');
    }
    if(SECURITY.stalledServices&&SECURITY.stalledServices.length)notes.push('<span class="y"><b>'+SECURITY.stalledServices.length+'</b> Automatic service'+(SECURITY.stalledServices.length>1?'s':'')+' not running (see Security tab)</span>');
    if(SECURITY.threats&&SECURITY.threats.length)notes.push('<span class="r"><b>'+SECURITY.threats.length+'</b> threat detection'+(SECURITY.threats.length>1?'s':'')+' recorded by Windows Defender</span>');
    if(SECURITY.exclFlags&&SECURITY.exclFlags.length)notes.push('<span class="y"><b>'+SECURITY.exclFlags.length+'</b> risky Defender exclusion'+(SECURITY.exclFlags.length>1?'s':'')+' (see Security tab)</span>');
    if(SECURITY.hostsFlags&&SECURITY.hostsFlags.length)notes.push('<span class="y">Hosts file redirects a known update/security domain (see Security tab)</span>');
    if(SECURITY.startupFlags&&SECURITY.startupFlags.length)notes.push('<span class="y"><b>'+SECURITY.startupFlags.length+'</b> flagged startup entr'+(SECURITY.startupFlags.length>1?'ies':'y')+' (see Security tab)</span>');
  }
  const gpuDrvRe=/nvlddmkm|amdwddmg|amdkmdag|atikmdag/i;
  const tdrEvents=SYSEVT.filter(r=>String(r.id)==='4101'||gpuDrvRe.test(r.prov)||gpuDrvRe.test(r.msg||''));
  if(tdrEvents.length){
    const drv=[...new Set(tdrEvents.map(r=>{const m2=(r.prov+' '+(r.msg||'')).match(gpuDrvRe);return m2?m2[0].toLowerCase():null;}).filter(Boolean))];
    notes.push('<span class="r"><b>'+tdrEvents.length+'</b> display driver timeout/reset event'+(tdrEvents.length>1?'s':'')+(drv.length?' ('+esc(drv.join(', '))+')':'')+'</span>');
  }
  const lke=RAW.filter(r=>/LiveKernelEvent/i.test(r.m||'')).length;
  if(lke)notes.push('<span class="r"><b>'+lke+'</b> LiveKernelEvent record'+(lke>1?'s':'')+' in reliability history</span>');
  if(NET&&NET.wifi&&NET.wifi.signal){
    const sig=parseInt(NET.wifi.signal)||0;
    if(sig&&sig<50)notes.push('<span class="y">Wi-Fi signal at '+sig+'%'+(NET.wifi.band?' on '+esc(NET.wifi.band):'')+'</span>');
  }
  if(MEMUSE&&MEMUSE.ct&&MEMUSE.cu/MEMUSE.ct>0.9)notes.push('<span class="y">Commit charge at '+Math.round(MEMUSE.cu/MEMUSE.ct*100)+'% of limit at time of capture</span>');
  const nEl=document.getElementById('notesBody');
  el.innerHTML=pairs.length?'<dl class="kv summary-kv">'+pairs.map(([k,v])=>'<dt>'+k+'</dt><dd>'+v+'</dd>').join('')+'</dl>':'';
  nEl.innerHTML=notes.length?'<ul class="notes">'+notes.map(n=>'<li>'+n+'</li>').join('')+'</ul>':'';
}
function sysCat(lvl){return lvl<=2?'err':lvl===3?'warn':'info';}
function renderSys(){
  const v=document.getElementById('sysView');
  let h='';
  if(!SYSEVT.length){
    h+='<div class="sys-ok">\u2713 No notable system events found in the collection window.</div>';
    v.innerHTML=h;return;
  }
  const evs=SYSEVT.map(r=>({...r,dt:parseDate(r.t)})).filter(r=>r.dt).sort((a,b)=>b.dt-a.dt);
  let lastDay=null;
  evs.forEach(e=>{
    const dk=e.dt.toISOString().slice(0,10);
    if(dk!==lastDay){lastDay=dk;h+='<div class="day-head">'+fmtDay(dk)+'</div>';}
    const cat=sysCat(e.lvl);
    let title=esc(e.prov)+' '+esc(e.id);
    if(e.bc&&e.bc!=='0')title+=' <span class="r" style="color:var(--err)">\u2014 Bugcheck 0x'+esc(parseInt(e.bc).toString(16).toUpperCase())+'</span>';
    if(e.cnt)title+=' \u00d7'+e.cnt;
    h+='<div class="row"><span class="time mono">'+fmtTime(e.dt)+'</span>'+
      '<span class="dot d-'+cat+'"></span>'+
      '<span class="title">'+title+'</span>'+
      '<div class="msg mono">'+esc(e.msg||'')+'</div></div>';
  });
  v.innerHTML=h;
  v.querySelectorAll('.row').forEach(r=>r.onclick=()=>r.classList.toggle('open'));
}
function renderSecurity(){
  const v=document.getElementById('securityView');
  if(!SECURITY){
    v.innerHTML='<div class="spec-section"><h2>Security</h2><div style="color:var(--faint)">No security data embedded.</div></div>';
    return;
  }
  let h='';
  const d=SECURITY.defender;
  if(d){
    h+='<div class="spec-section"><h2>Windows Defender</h2><dl class="kv">';
    h+='<dt>Real-time protection</dt><dd style="color:'+(d.rtp==='True'?'var(--ok)':'var(--err)')+'">'+(d.rtp==='True'?'Enabled':'Disabled')+'</dd>';
    if(d.lastQuick)h+='<dt>Last quick scan</dt><dd>'+esc(d.lastQuick)+'</dd>';
    if(d.lastFull)h+='<dt>Last full scan</dt><dd>'+esc(d.lastFull)+'</dd>';
    if(d.sigAge)h+='<dt>Signature age</dt><dd>'+esc(d.sigAge)+' day'+(d.sigAge==='1'?'':'s')+'</dd>';
    h+='</dl></div>';
  }
  if(SECURITY.firewall&&SECURITY.firewall.length){
    h+='<div class="spec-section"><h2>Firewall</h2><dl class="kv">';
    SECURITY.firewall.forEach(f=>{h+='<dt>'+esc(f.profile)+'</dt><dd style="color:'+(f.enabled==='True'?'var(--ok)':'var(--err)')+'">'+(f.enabled==='True'?'Enabled':'Disabled')+'</dd>';});
    h+='</dl></div>';
  }
  if(SECURITY.threats&&SECURITY.threats.length){
    h+='<div class="spec-section"><h2>Threat detections ('+SECURITY.threats.length+')</h2><dl class="kv">';
    SECURITY.threats.forEach(t=>{h+='<dt>'+esc(t.time)+'</dt><dd>'+esc(t.name)+(t.act==='True'?' <span style="color:var(--ok)">(action successful)</span>':' <span style="color:var(--err)">(action failed)</span>')+'</dd>';});
    h+='</dl></div>';
  } else if(d) {
    h+='<div class="spec-section"><h2>Threat detections</h2><div style="color:var(--ok)">\u2713 No threats recorded by Windows Defender.</div></div>';
  }
  if(SECURITY.exclFlags&&SECURITY.exclFlags.length){
    h+='<div class="spec-section"><h2>Defender exclusion alerts</h2><ul class="notes">'+SECURITY.exclFlags.map(f=>'<li><span class="y">'+esc(f)+'</span></li>').join('')+'</ul></div>';
  }
  if(SECURITY.exclusions&&SECURITY.exclusions.length){
    h+='<div class="spec-section"><h2>Defender Exclusions ('+SECURITY.exclusions.length+')</h2><div style="color:var(--dim);font-size:14px;line-height:1.8">'+SECURITY.exclusions.map(esc).join('<br>')+'</div></div>';
  }
  if(SECURITY.hostsFlags&&SECURITY.hostsFlags.length){
    h+='<div class="spec-section"><h2>Hosts file</h2><div style="color:var(--dim);font-size:14px;margin-bottom:8px">'+SECURITY.hostsCustom+' custom entr'+(SECURITY.hostsCustom===1?'y':'ies')+' found.</div><ul class="notes">'+SECURITY.hostsFlags.map(f=>'<li><span class="y">'+esc(f)+'</span></li>').join('')+'</ul></div>';
  } else if(typeof SECURITY.hostsCustom==='number'){
    h+='<div class="spec-section"><h2>Hosts file</h2><div style="color:var(--dim);font-size:14px">'+SECURITY.hostsCustom+' custom entr'+(SECURITY.hostsCustom===1?'y':'ies')+' found, none flagged.</div></div>';
  }
  if(SECURITY.startupFlags&&SECURITY.startupFlags.length){
    h+='<div class="spec-section"><h2>Startup entries flagged</h2><ul class="notes">'+SECURITY.startupFlags.map(f=>'<li><span class="y">'+esc(f)+'</span></li>').join('')+'</ul></div>';
  }
  if(SECURITY.stalledServices&&SECURITY.stalledServices.length){
    h+='<div class="spec-section"><h2>Automatic services not running ('+SECURITY.stalledServices.length+')</h2><dl class="kv">';
    SECURITY.stalledServices.forEach(s=>{h+='<dt>'+esc(s.name)+'</dt><dd style="color:var(--warn)">'+esc(s.state)+'</dd>';});
    h+='</dl></div>';
  }
  v.innerHTML=h||'<div class="spec-section"><h2>Security</h2><div style="color:var(--faint)">No security data embedded.</div></div>';
}
function renderGPU(){
  const v=document.getElementById('gpuView');
  if(!GPUS.length && !DISPLAYS.length){
    v.innerHTML='<div class="spec-section"><h2>GPU and Display(s)</h2><div style="color:var(--faint)">No GPU data embedded.</div></div>';
    return;
  }
  const byGpu={};
  DISPLAYS.forEach(d=>{(byGpu[d.gpu]=byGpu[d.gpu]||[]).push(d);});
  const drvByName={},radByName={},vramByName={};
  GPUS.forEach(g=>{if(g.name){drvByName[g.name]=g.drv;radByName[g.name]=g.radeon||'';vramByName[g.name]=g.vram||0;}});
  const gnames=Object.keys(byGpu).length?Object.keys(byGpu):GPUS.map(g=>g.name);

  let h='<div class="spec-section"><h2>Graphics adapters ('+gnames.length+')</h2>';
  if(HAGS)h+='<div style="color:var(--dim);font-size:14px;margin-bottom:16px">Hardware-accelerated GPU Scheduling: <b style="color:var(--text)">'+esc(HAGS)+'</b></div>';
  h+='<div class="drive-grid">';
  gnames.forEach(g=>{
    const drv=drvByName[g]?friendlyDriver(g,esc(drvByName[g]),radByName[g]?esc(radByName[g]):''):'';
    const vram=vramByName[g];
    const displays=byGpu[g]||[];
    h+='<div class="drive"><h3>'+esc(g)+'</h3>'+
      (drv?'<div class="sub">Driver '+drv+(vram?' \u00b7 '+vram+' GB VRAM':'')+'</div>':(vram?'<div class="sub">'+vram+' GB VRAM</div>':''));
    const dispRows=displays.filter(d=>d.mon||d.mode);
    const fallbackMons=(!Object.keys(byGpu).length && MONS.length)?MONS:[];
    if(dispRows.length||fallbackMons.length){
      h+='<div class="sev-head" style="color:var(--dim);padding-top:4px">Connected Display'+((dispRows.length+fallbackMons.length)>1?'s':'')+'</div><dl class="kv smart-kv">';
      dispRows.forEach(d=>{h+='<dt>'+esc(d.mon||'Display')+'</dt><dd>'+esc(d.mode||'')+'</dd>';});
      fallbackMons.forEach(m=>{h+='<dt>Connected Display</dt><dd>'+esc(m)+'</dd>';});
      h+='</dl>';
    }
    h+='</div>';
  });
  h+='</div></div>';
  v.innerHTML=h;
}
function renderMemory(){
  const v=document.getElementById('memoryView');
  if(!RAM.length){
    v.innerHTML='<div class="spec-section"><h2>Memory modules</h2><div style="color:var(--faint)">No memory module data embedded.</div></div>';
    return;
  }
  let h='<div class="spec-section"><h2>Memory modules ('+RAM.length+')</h2><div class="drive-grid">';
  RAM.forEach(m=>{
    h+='<div class="drive"><h3>'+esc(m.slot)+'</h3>'+
      '<div class="sub">'+esc(m.mfr||'')+'</div>'+
      '<dl class="kv smart-kv">'+
      '<dt>Part number</dt><dd>'+esc(m.pn||'?')+'</dd>'+
      '<dt>Capacity</dt><dd>'+esc(m.cap)+' GB</dd>'+
      (m.rated?'<dt>Rated speed</dt><dd>'+esc(m.rated)+' MT/s</dd>':'')+
      (m.conf?'<dt>Configured speed</dt><dd>'+esc(m.conf)+' MT/s</dd>':'')+
      '</dl></div>';
  });
  h+='</div></div>';
  if(MEMUSE&&MEMUSE.pt){
    const pct=Math.round(MEMUSE.pu/MEMUSE.pt*100);
    h+='<div class="spec-section"><h2>Memory usage at capture</h2><dl class="kv">'+
      '<dt>Physical memory used</dt><dd>'+MEMUSE.pu.toFixed(1)+' / '+MEMUSE.pt.toFixed(1)+' GB ('+pct+'%)</dd>';
    if(MEMUSE.ct)h+='<dt>Commit charge</dt><dd>'+MEMUSE.cu.toFixed(1)+' / '+MEMUSE.ct.toFixed(1)+' GB ('+Math.round(MEMUSE.cu/MEMUSE.ct*100)+'%)</dd>';
    h+='</dl></div>';
  }
  v.innerHTML=h;
}
function renderNet(){
  const v=document.getElementById('netView');
  if(!NET||(!NET.adapters||!NET.adapters.length)&&!NET.wifi){
    v.innerHTML='<div class="spec-section"><h2>Network adapters</h2><div style="color:var(--faint)">No network data embedded.</div></div>';
    return;
  }
  let h='';
  if(NET.adapters&&NET.adapters.length){
    h+='<div class="spec-section"><h2>Network adapters ('+NET.adapters.length+')</h2><div class="drive-grid">';
    NET.adapters.forEach(a=>{
      const up=/^up$/i.test(a.status);
      const stCol=up?'var(--ok)':/disconnect/i.test(a.status)?'var(--warn)':'var(--faint)';
      h+='<div class="drive"><h3>'+esc(a.name)+'</h3>'+
        '<div class="sub">'+esc(a.desc||'')+'</div>'+
        '<dl class="kv smart-kv">'+
        '<dt>Status</dt><dd style="color:'+stCol+'">'+esc(a.status)+'</dd>'+
        (up&&a.speed?'<dt>Link speed</dt><dd>'+esc(a.speed)+'</dd>':'')+
        (a.media?'<dt>Media</dt><dd>'+esc(friendlyMedia(a.media))+'</dd>':'')+
        '</dl></div>';
    });
    h+='</div></div>';
  }
  if(NET.vpns&&NET.vpns.length){
    h+='<div class="spec-section"><h2>VPN / virtual adapters ('+NET.vpns.length+')</h2><div class="drive-grid">';
    NET.vpns.forEach(a=>{
      const up=/^up$/i.test(a.status);
      h+='<div class="drive"><h3>'+esc(a.name)+'</h3>'+
        '<div class="sub">'+esc(a.desc||'')+'</div>'+
        '<dl class="kv smart-kv"><dt>Status</dt><dd style="color:'+(up?'var(--ok)':'var(--faint)')+'">'+esc(a.status)+'</dd></dl></div>';
    });
    h+='</div></div>';
  }
  if(NET.wifi&&NET.wifi.signal){
    const w=NET.wifi;
    const sig=parseInt(w.signal)||0;
    h+='<div class="spec-section"><h2>Wi-Fi connection</h2>'+
      '<div class="drive" style="max-width:420px">'+
      '<div class="meter'+(sig<50?' low':'')+'" style="margin-bottom:8px"><div style="width:'+sig+'%"></div></div>'+
      '<dl class="kv smart-kv">'+
      '<dt>Signal</dt><dd>'+esc(w.signal)+'</dd>'+
      (w.band?'<dt>Band</dt><dd>'+esc(w.band)+'</dd>':'')+
      (w.channel?'<dt>Channel</dt><dd>'+esc(w.channel)+'</dd>':'')+
      (w.radio?'<dt>Radio type</dt><dd>'+esc(w.radio)+'</dd>':'')+
      (w.rx?'<dt>Receive rate</dt><dd>'+esc(w.rx)+' Mbps</dd>':'')+
      (w.tx?'<dt>Transmit rate</dt><dd>'+esc(w.tx)+' Mbps</dd>':'')+
      (w.auth?'<dt>Authentication</dt><dd>'+esc(w.auth)+'</dd>':'')+
      '</dl></div>'+
      '<div style="color:var(--faint);font-size:13.5px;margin-top:10px">SSID, BSSID and IP details are intentionally not collected.</div></div>';
  }
  v.innerHTML=h;
}
function renderDumps(){
  if(!DUMPS.length)return;
  document.getElementById('dumpsTab').style.display='';
  let h='<div class="spec-section"><h2>Memory dumps ('+DUMPS.length+')</h2><dl class="kv">';
  DUMPS.forEach(d=>{h+='<dt class="mono">'+esc(d.n)+'</dt><dd>'+esc(d.d)+' \u00b7 '+esc(d.z)+'</dd>';});
  h+='</dl><div style="color:var(--faint);font-size:13px;margin-top:12px">The .dmp files are included in the zip.</div></div>';
  document.getElementById('dumpsView').innerHTML=h;
}
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('on',x===t));
  document.body.className='tab-'+t.dataset.tab;
  document.getElementById('pageTitleSub').textContent='- '+t.textContent;
});
document.querySelectorAll('.nav-group-title:not(.static)').forEach(g=>g.onclick=()=>{
  g.closest('.nav-group').classList.toggle('collapsed');
});
renderSpecs();
load(RAW);
renderSummary();
renderSys();
renderDumps();
renderNet();
renderGPU();
renderMemory();
renderSecurity();
document.getElementById('pageFoot').textContent=(GEN?'Generated '+GEN+' · ':'')+'PCHH Triage'+(VER?' v'+VER:'')+' · Author: Rory (ctrl.alt.repeat)';
</script>
</body>
</html>
'@


$dmpfound = $false

$errors = @{
    fileCreate  = $false
    Compress    = $false
    event       = $false
    reliability = $false
}

$null = New-Module {
    function Invoke-WithoutProgress {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)] [scriptblock] $ScriptBlock
        )

        $prevProgressPreference = $global:ProgressPreference
        $global:ProgressPreference = 'SilentlyContinue'

        try {
            . $ScriptBlock
        }
        finally {
            $global:ProgressPreference = $prevProgressPreference
        }
    }
}

function cmark {
    return [char]0x2705
}

function xmark {
    return [char]0x274C
}

function dmpcheck {
    Clear-Host 
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor DarkGreen
    Write-Host "         PCHH Triage v1.0 - 23/07/2026            " -ForegroundColor Green
    Write-Host "       Developed by Rory (ctrl.alt.repeat)		  " -ForegroundColor DarkGray
    Write-Host "==================================================" -ForegroundColor DarkGreen
    Write-Host ""
    Write-Host "This collects crash logs, specs and diagnostics into" -ForegroundColor Gray
    Write-Host "a single zip on your Desktop. This can take time, please be patient." -ForegroundColor Gray
    Write-Host ""
    Write-Host "[1/4] Collecting system specs.." -ForegroundColor Blue

    $limit = (Get-Date).AddDays(-60)

    Get-ChildItem -Path $env:systemroot -Filter "MEMORY.dmp" -File | Remove-Item  -Force -Recurse -ErrorAction SilentlyContinue > $null 2>&1

    if (Test-Path $minidump) {
        Get-ChildItem -Path $source -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.LastWriteTime -lt $limit } | Remove-Item -Force -ErrorAction SilentlyContinue > $null 2>&1

        if (Test-Path $source) {
            $dmpfound = $true
        }
    }
    
    filecreation
}

function filecreation {
    Remove-Item -Path "$File\*" -Force -Recurse -ErrorAction SilentlyContinue > $null 2>&1

    try {
        New-Item -Path $File -ItemType Directory -Force | Out-Null
        New-Item -Path $infofile -ItemType File -Force | Out-Null
    }
    catch {
        $errors.fileCreate = $true
    }

    fileadd
}

# Grabbing specs & info
function fileadd {

    $secCompat = $false
    $cpu = Get-WmiObject Win32_Processor
    $cpuName = $cpu | Select-Object -ExpandProperty Name
    $cpuSpeed = $cpu | Select-Object -ExpandProperty MaxClockSpeed
    $gpu = Get-WmiObject Win32_VideoController | Select-Object -ExpandProperty Name

    if ((Get-Tpm).TpmEnabled -eq "True") {
        $tpmEnabled = "Enabled"
    }
    else {
        $tpmEnabled = "Disabled"
    }

    $tpmSpecParts = "$((Get-CimInstance -Namespace "root\CIMV2\Security\MicrosoftTPM" -ClassName Win32_TPM).SpecVersion)" -split ',' | ForEach-Object { $_.Trim() }
    $tpmVersion = "$($tpmSpecParts[0])"
    if ($tpmVersion -and $tpmVersion -notmatch '\.') { $tpmVersion = "$tpmVersion.0" }
    if ($tpmSpecParts.Count -ge 3 -and $tpmSpecParts[2]) { $tpmVersion = "$tpmVersion (rev $($tpmSpecParts[2]))" }

    $motherboardModel = Get-WmiObject Win32_BaseBoard | Select-Object -ExpandProperty Product
    $motherboardMfr = Get-WmiObject Win32_BaseBoard | Select-Object -ExpandProperty Manufacturer
    $bios = Get-WmiObject Win32_BIOS
    $biosVersion = $bios | Select-Object -ExpandProperty SMBIOSBIOSVersion
    $biosDate = $bios | Select-Object -ExpandProperty ReleaseDate
    $os = Get-WmiObject Win32_OperatingSystem
    $osName = $os | Select-Object -ExpandProperty Caption
    $osVersion = $os | Select-Object -ExpandProperty Version
    $bootDevice = $os | Select-Object -ExpandProperty BootDevice
    $systemDirectory = $env:SystemDrive
    $secureBoot = try { Confirm-SecureBootUEFI } catch { $secCompat = $true }
    $fastboot = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name HiberbootEnabled).HiberbootEnabled

    $buildNumber = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
    $ubr = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").UBR
    $build = "$buildNumber.$ubr"


    $osInstallDate = try { ([System.Management.ManagementDateTimeConverter]::ToDateTime($os.InstallDate)).ToString("dd/MM/yyyy") } catch { "" }
    $cpuCores = ($cpu | Select-Object -ExpandProperty NumberOfCores) -join "+"
    $cpuThreads = ($cpu | Select-Object -ExpandProperty ThreadCount) -join "+"
    $avNames = try { (Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct -ErrorAction Stop | Select-Object -ExpandProperty displayName) -join ", " } catch { "" }
    $uacEnabled = try { if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -ErrorAction Stop).EnableLUA -eq 1) { "Enabled" } else { "Disabled" } } catch { "" }
    $powerPlan = try { if ((powercfg /getactivescheme) -match '\((.+)\)\s*$') { $Matches[1] } else { "" } } catch { "" }

    $lboottime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $uptime = (Get-Date) - $lboottime

    $pgfile = Get-WmiObject -Query "SELECT * FROM Win32_PageFileUsage"
    $pgfilesize = $pgfile.AllocatedBaseSize

    $installedMemory = Get-WmiObject Win32_ComputerSystem | Select-Object -ExpandProperty TotalPhysicalMemory
    $ramSpeed = Get-WmiObject Win32_PhysicalMemory | Select-Object -ExpandProperty Speed

    $secureBootState = if ($secureBoot -match "True") { "Enabled" } elseif ($secureBoot -match "False") { "Disabled" } elseif ($secCompat -eq "$true") { "Not Supported" }
    $fastbootState = if ($fastboot -eq "1") { "Enabled" } else { "Disabled" }

    specs "CPU Name: $cpuName"
    specs "CPU Speed (MHz): $cpuSpeed"
    specs "GPU: $gpu"
    specs "`nTPM Status: $tpmEnabled"
    if ($tpmEnabled -eq "Enabled") {
        specs "TPM Version: $tpmVersion"
    }
    specs "`nMotherboard Manufacturer: $motherboardMfr"
    specs "Motherboard: $motherboardModel"
    specs "BIOS Version: $biosVersion"
    specs "BIOS Date: $([System.Management.ManagementDateTimeConverter]::ToDateTime($biosDate))"
    specs "`nOS: $osName"
    specs "OS Version: $osVersion"
    specs "System Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes"
    specs "Build: $build"
    specs "Page File Size: $pgfilesize MB"
    specs "Boot Device: $bootDevice"
    specs "System Directory: $systemDirectory\"
    specs "Secure Boot State: $secureBootState"
    specs "Fast Boot State: $fastbootState"
    specs "CPU Cores/Threads: ${cpuCores}C / ${cpuThreads}T"
    if ($osInstallDate) { specs "Windows Install Date: $osInstallDate" }
    if ($avNames) { specs "Antivirus: $avNames" }
    if ($uacEnabled) { specs "UAC: $uacEnabled" }
    if ($powerPlan) { specs "Active Power Plan: $powerPlan" }
    specs "`nRam Capacity: $([math]::Round($installedMemory/1GB)) GB"
    specs "RAM Speed: $ramSpeed MT/s"

    $drives = Get-WmiObject Win32_LogicalDisk | ForEach-Object {
        $logicalDisk = $_
        $windowsDrive = $logicalDisk.DeviceID.TrimEnd(':')

        $partition = Get-Partition | Where-Object { $_.DriveLetter -eq $windowsDrive }

        $diskNumber = if ($partition) {
            $partition.DiskNumber
        }
        else {
            $null
        }

        $disk = if ($null -ne $diskNumber) {
            Get-Disk -Number $diskNumber
        }

        $physicalDisk = if ($disk) {
            Get-PhysicalDisk | Where-Object { $_.DeviceId -eq $diskNumber }
        }

        $driveType = if ($physicalDisk) { $physicalDisk.MediaType } else { 'Unknown' }
        $operationalStatus = if ($physicalDisk) { $physicalDisk.OperationalStatus } else { 'Unknown' }
        $healthStatus = if ($physicalDisk) { $physicalDisk.HealthStatus } else { 'Unknown' }

        $totalSizeGB = if ($logicalDisk.Size) { [math]::Round($logicalDisk.Size / 1GB, 2) } else { 0 }
        $freeSpaceGB = if ($logicalDisk.FreeSpace) { [math]::Round($logicalDisk.FreeSpace / 1GB, 2) } else { 0 }
        $percentageFree = if ($totalSizeGB -ne 0) {
            [math]::Round(($freeSpaceGB / $totalSizeGB) * 100, 2)
        }
        else {
            'N/A'
        }

        [PSCustomObject]@{
            'Drive Label'         = $logicalDisk.DeviceID + '\'
            'Drive Name'          = if (-not [string]::IsNullOrEmpty($logicalDisk.VolumeName)) { $logicalDisk.VolumeName } else { 'No Name Found' }
            'Drive Status'        = "$operationalStatus, $healthStatus"
            'Windows Drive'       = ($logicalDisk.DeviceID -eq "$env:SystemDrive")
            'Drive ID'            = if ($null -ne $diskNumber) { $diskNumber } else { 'Unknown' }
            'Drive Type'          = $driveType
            'Total Size (GB)'     = $totalSizeGB
            'Free Space (GB)'     = $freeSpaceGB
            'Percentage Free (%)' = $percentageFree
        }
    }

    specs "`n`nDrive Information:`n`n"

    foreach ($drive in $drives) {
        specs "Drive Label: $($drive.'Drive Label')"
        specs "Drive Name: $($drive.'Drive Name')"
        specs "Drive Status: $($drive.'Drive Status')"
        specs "Windows Drive: $($drive.'Windows Drive')"
        specs "Drive ID: $($drive.'Drive ID')"
        specs "Drive Type: $($drive.'Drive Type')"
        specs "Total Size (GB): $($drive.'Total Size (GB)')"
        specs "Free Space (GB): $($drive.'Free Space (GB)')"
        specs "Percentage Free (%): $($drive.'Percentage Free (%)')`n"
    }



    $installedPrograms = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Select-Object DisplayName |
    Where-Object { $null -ne $_.DisplayName }

    $programs = $installedPrograms | Out-String

    specs "`n`nPrograms Installed:`n $programs"
    
    Write-Host -NoNewline -ForegroundColor Green "$(cmark)"
    Write-Host " System specs collected"

    eventlogexport
}


function specs {
    param (
        [string]$value
    )
    Add-Content -Path $infofile -Value "$value"
}


function eventlogexport {
    Write-Host ""
    Write-Host "[2/4] Exporting Windows event logs.." -ForegroundColor Blue

    $startTime = (Get-Date).AddDays(-$lookbackDays).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss")

    try {
        wevtutil epl System $sys_eventlog_path /q:"*[System[TimeCreated[@SystemTime>='$startTime']]]"
    }
    catch {
        $errors.event = $true
        functionerror
    }

    Write-Host -NoNewline -ForegroundColor Green "$(cmark)"
    Write-Host " Event logs exported"

    reliabilityexport
}
# Curated System event log entries (crash / hardware / storage / GPU / service failures)
function Get-CuratedSystemEvents {
    $allow = @(
        @{ P = '*Kernel-Power';               I = 41, 137, 142 },
        @{ P = '*WHEA-Logger';                I = 17, 18, 19, 46, 47 },
        @{ P = 'disk';                        I = 7, 51, 153, 154, 157 },
        @{ P = '*stor*';                      I = 129 },
        @{ P = '*Ntfs*';                      I = 55 },
        @{ P = 'volmgr';                      I = 161 },
        @{ P = 'Display';                     I = 4101 },
        @{ P = 'nvlddmkm';                    I = 13, 14 },
        @{ P = 'Service Control Manager';     I = 7034 },
        @{ P = '*MemoryDiagnostics-Results';  I = 1102 },
        @{ P = 'EventLog';                    I = 6008 },
        @{ P = '*WER-SystemErrorReporting';   I = 1001 }
    )
    $ids = @($allow | ForEach-Object { $_.I } | Select-Object -Unique)
    $since = (Get-Date).AddDays(-$lookbackDays)

    # Windows limits FilterHashtable to 23 event IDs per query - chunk the list
    $raw = @()
    for ($i = 0; $i -lt $ids.Count; $i += 20) {
        $chunk = $ids[$i..([Math]::Min($i + 19, $ids.Count - 1))]
        $raw += @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; StartTime = $since; Id = $chunk } -ErrorAction SilentlyContinue)
    }

    $matched = @($raw | Where-Object {
        $ev = $_
        $allow | Where-Object { $ev.ProviderName -like $_.P -and $_.I -contains $ev.Id } | Select-Object -First 1
    })

    # WHEA 17 (corrected PCIe) can flood - summarise to a single record
    $whea17 = @($matched | Where-Object { $_.ProviderName -like '*WHEA-Logger' -and $_.Id -eq 17 })
    $keep   = @($matched | Where-Object { -not ($_.ProviderName -like '*WHEA-Logger' -and $_.Id -eq 17) })

    $out = @($keep | Select-Object -First 400 | ForEach-Object {
        $bc = ''
        if ($_.ProviderName -like '*Kernel-Power' -and $_.Id -eq 41) {
            try {
                $x = [xml]$_.ToXml()
                $bc = "$(($x.Event.EventData.Data | Where-Object { $_.Name -eq 'BugcheckCode' }).'#text')"
            } catch { }
        }
        [PSCustomObject]@{
            t    = $_.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss")
            prov = ($_.ProviderName -replace '^Microsoft-Windows-', '')
            id   = "$($_.Id)"
            lvl  = [int]$_.Level
            bc   = $bc
            msg  = "$($_.Message)"
        }
    })

    if ($whea17.Count -gt 0) {
        $latest = $whea17 | Sort-Object TimeCreated -Descending | Select-Object -First 1
        $out += [PSCustomObject]@{
            t    = $latest.TimeCreated.ToString("dd/MM/yyyy HH:mm:ss")
            prov = 'WHEA-Logger'
            id   = '17'
            lvl  = 3
            bc   = ''
            cnt  = $whea17.Count
            msg  = "$($whea17.Count) corrected PCIe hardware error(s) recorded in the last $lookbackDays days (summarised)."
        }
    }

    return $out
}

# Exports reliability history + system specs and builds an interactive HTML viewer
function reliabilityexport {
    Write-Host ""
    Write-Host "[3/4] Collecting diagnostics.." -ForegroundColor Blue

    try {
        Write-Host "      - Reliability history" -ForegroundColor DarkGray
        $recs = Get-CimInstance Win32_ReliabilityRecords -ErrorAction Stop | ForEach-Object {
            [PSCustomObject]@{
                t = $_.TimeGenerated.ToString("dd/MM/yyyy HH:mm:ss")
                s = $_.SourceName
                e = "$($_.EventIdentifier)"
                p = $_.ProductName
                m = $_.Message
            }
        }

        # CSV copy for sharing
        $recs | Export-Csv $reliability_csv_path -NoTypeInformation -Encoding UTF8

        # Curated system events for the viewer
        Write-Host "      - Notable system events" -ForegroundColor DarkGray
        $sysEvents = @(Get-CuratedSystemEvents)

        Write-Host "      - Drive S.M.A.R.T data" -ForegroundColor DarkGray
        # Raw ATA SMART attributes (SATA drives; NVMe reports via reliability counters instead)
        $rawSmart = @{}
        $predictFail = @{}
        try {
            $ddMap = @{}
            Get-CimInstance Win32_DiskDrive -ErrorAction Stop | ForEach-Object { $ddMap["$($_.PNPDeviceID)".ToUpper()] = "$($_.Index)" }
            $fpd = @(Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictData -ErrorAction Stop)
            foreach ($f in $fpd) {
                $pnp = ("$($f.InstanceName)" -replace '_\d+$', '').ToUpper()
                if (-not $ddMap.ContainsKey($pnp)) { continue }
                $attrs = @{}
                $bytes = $f.VendorSpecific
                for ($i = 0; $i -lt 30; $i++) {
                    $o = 2 + ($i * 12)
                    if ($o + 11 -ge $bytes.Count) { break }
                    $id = [int]$bytes[$o]
                    if ($id -eq 0) { continue }
                    $rawv = [uint64]0
                    for ($j = 0; $j -lt 6; $j++) { $rawv += ([uint64]$bytes[$o + 5 + $j]) -shl (8 * $j) }
                    $attrs[$id] = $rawv
                }
                $rawSmart[$ddMap[$pnp]] = $attrs
            }
            $fps = @(Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue)
            foreach ($f in $fps) {
                $pnp = ("$($f.InstanceName)" -replace '_\d+$', '').ToUpper()
                if ($ddMap.ContainsKey($pnp) -and $f.PredictFailure) { $predictFail[$ddMap[$pnp]] = $true }
            }
        } catch { }

        # SMART / drive reliability data (admin required; some drives report partial data)
        $smart = @()
        try {
            $smart = @(Get-PhysicalDisk -ErrorAction Stop | ForEach-Object {
                $pd = $_
                $rc = $null
                try { $rc = $pd | Get-StorageReliabilityCounter -ErrorAction Stop } catch { }
                $ra = $rawSmart["$($pd.DeviceId)"]
                [PSCustomObject]@{
                    name   = "$($pd.FriendlyName)"
                    disk   = "$($pd.DeviceId)"
                    media  = "$($pd.MediaType)"
                    bus    = "$($pd.BusType)"
                    health = "$($pd.HealthStatus)"
                    op     = "$($pd.OperationalStatus)"
                    temp   = if ($null -ne $rc.Temperature -and $rc.Temperature -gt 0) { "$($rc.Temperature)" } else { "" }
                    tmax   = if ($null -ne $rc.TemperatureMax -and $rc.TemperatureMax -gt 0) { "$($rc.TemperatureMax)" } else { "" }
                    hours  = if ($null -ne $rc.PowerOnHours) { "$($rc.PowerOnHours)" } else { "" }
                    wear   = if ($null -ne $rc.Wear) { "$($rc.Wear)" } else { "" }
                    reu    = if ($null -ne $rc.ReadErrorsUncorrected) { "$($rc.ReadErrorsUncorrected)" } else { "" }
                    rec    = if ($null -ne $rc.ReadErrorsCorrected) { "$($rc.ReadErrorsCorrected)" } else { "" }
                    weu    = if ($null -ne $rc.WriteErrorsUncorrected) { "$($rc.WriteErrorsUncorrected)" } else { "" }
                    wec    = if ($null -ne $rc.WriteErrorsCorrected) { "$($rc.WriteErrorsCorrected)" } else { "" }
                    rl     = if ($ra -and $ra.ContainsKey(5))   { "$($ra[5])" }   else { "" }
                    cto    = if ($ra -and $ra.ContainsKey(188)) { "$($ra[188])" } else { "" }
                    pend   = if ($ra -and $ra.ContainsKey(197)) { "$($ra[197])" } else { "" }
                    unc    = if ($ra -and $ra.ContainsKey(198)) { "$($ra[198])" } else { "" }
                    crc    = if ($ra -and $ra.ContainsKey(199)) { "$($ra[199])" } else { "" }
                    pf     = if ($predictFail["$($pd.DeviceId)"]) { "1" } else { "" }
                }
            })
        } catch { }

        # Dirty bit per fixed volume
        $dirtyVols = @()
        try {
            Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop | ForEach-Object {
                $dl = $_.DeviceID
                $q = fsutil dirty query $dl 2>$null
                if ("$q" -match 'is Dirty') { $dirtyVols += "$dl" }
            }
        } catch { }

        # Disk layout: partition style (GPT/MBR) and partition -> drive-letter chain per physical disk.
        # Useful for Secure Boot troubleshooting (requires GPT) and spotting a missing/damaged ESP.
        $diskLayout = @()
        try {
            Get-Disk -ErrorAction Stop | ForEach-Object {
                $disk = $_
                $parts = @(Get-Partition -DiskNumber $disk.Number -ErrorAction SilentlyContinue | Sort-Object PartitionNumber | ForEach-Object {
                    $p = $_
                    $typeLabel = switch -Regex ("$($p.GptType)$($p.MbrType)") {
                        'c12a7328-f81f-11d2-ba4b-00a0c93ec93b' { "EFI System Partition"; break }
                        'e3c9e316-0b5c-4db8-817d-f92df00215ae' { "Microsoft Reserved"; break }
                        'de94bba4-06d1-4d40-a16a-bfd50179d6ac' { "Recovery"; break }
                        '^27$'                                 { "Recovery (MBR)"; break }
                        default { if ($p.DriveLetter) { "Data" } else { "System" } }
                    }
                    [PSCustomObject]@{
                        num    = $p.PartitionNumber
                        type   = $typeLabel
                        sizeGB = [math]::Round($p.Size / 1GB, 2)
                        letter = if ($p.DriveLetter) { "$($p.DriveLetter):" } else { "" }
                    }
                })
                $diskLayout += [PSCustomObject]@{
                    disk       = $disk.Number
                    style      = "$($disk.PartitionStyle)"
                    sizeGB     = [math]::Round($disk.Size / 1GB, 1)
                    partitions = $parts
                }
            }
        } catch { }

        # Per-stick RAM info (slots, part numbers, rated vs configured speed)
        $ram = @()
        try {
            $ram = @(Get-CimInstance Win32_PhysicalMemory -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{
                    slot  = "$($_.DeviceLocator)"
                    mfr   = "$($_.Manufacturer)".Trim()
                    pn    = "$($_.PartNumber)".Trim()
                    cap   = "$([math]::Round($_.Capacity / 1GB))"
                    rated = if ($_.Speed) { "$($_.Speed)" } else { "" }
                    conf  = if ($_.ConfiguredClockSpeed) { "$($_.ConfiguredClockSpeed)" } else { "" }
                }
            })
        } catch { }

        # GPU adapters (name, driver, current mode) and monitor models
        $radeonVer = ""
        try {
            $radeonVer = "$((Get-ItemProperty 'HKLM:\SOFTWARE\AMD\CN' -ErrorAction Stop).RadeonSoftwareVersion)"
        } catch { }
        # Accurate VRAM per adapter - Win32_VideoController's AdapterRAM is a 32-bit field that
        # overflows/wraps on cards with >4GB VRAM (a known, widely-reported Windows bug). The real
        # value lives in the driver's registry key as a 64-bit QWORD.
        $vramByKey = @{}
        try {
            $classRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}'
            Get-ChildItem $classRoot -ErrorAction Stop | Where-Object { $_.PSChildName -match '^\d{4}$' } | ForEach-Object {
                $qw = (Get-ItemProperty -Path $_.PSPath -Name 'HardwareInformation.qwMemorySize' -ErrorAction SilentlyContinue).'HardwareInformation.qwMemorySize'
                $drvDesc = (Get-ItemProperty -Path $_.PSPath -Name 'DriverDesc' -ErrorAction SilentlyContinue).DriverDesc
                if ($qw -and $drvDesc) { $vramByKey[$drvDesc] = $qw }
            }
        } catch { }

        Write-Host "      - Hotfixes, Device Manager and audio devices" -ForegroundColor DarkGray
        $hotfixes = @()
        try {
            $hotfixes = @(Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending | ForEach-Object {
                [PSCustomObject]@{
                    id   = "$($_.HotFixID)"
                    desc = "$($_.Description)"
                    date = if ($_.InstalledOn) { $_.InstalledOn.ToString("dd/MM/yyyy") } else { "" }
                }
            })
        } catch { }

        Write-Host "      - Windows Update history and pending reboot status (can take a few seconds)" -ForegroundColor DarkGray
        # Pending reboot: several independent flags across Windows can indicate this; any one being set means yes
        $pendingReboot = $false
        try {
            $rebootChecks = @(
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
                'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
            )
            foreach ($rc in $rebootChecks) { if (Test-Path $rc) { $pendingReboot = $true } }
            if (-not $pendingReboot) {
                $pfro = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue).PendingFileRenameOperations
                if ($pfro) { $pendingReboot = $true }
            }
        } catch { }

        # Windows Update service status
        $wuServiceStatus = ""
        try { $wuServiceStatus = "$((Get-Service -Name wuauserv -ErrorAction Stop).Status)" } catch { }

        # Recent Windows Update history, including FAILED/pending attempts that Get-HotFix cannot show
        $wuHistory = @()
        try {
            $session = New-Object -ComObject Microsoft.Update.Session
            $searcher = $session.CreateUpdateSearcher()
            $historyCount = $searcher.GetTotalHistoryCount()
            if ($historyCount -gt 0) {
                $resultMap = @{ 1 = "In progress"; 2 = "Succeeded"; 3 = "Succeeded with errors"; 4 = "Failed"; 5 = "Cancelled" }
                $wuHistory = @($searcher.QueryHistory(0, [Math]::Min($historyCount, 40)) | ForEach-Object {
                    [PSCustomObject]@{
                        title  = "$($_.Title)"
                        date   = if ($_.Date) { $_.Date.ToString("dd/MM/yyyy HH:mm") } else { "" }
                        result = if ($resultMap.ContainsKey([int]$_.ResultCode)) { $resultMap[[int]$_.ResultCode] } else { "Unknown" }
                    }
                } | Sort-Object date -Descending)
            }
        } catch { }

        $devErrors = @()
        try {
            $devErrors = @(Get-CimInstance Win32_PNPEntity -ErrorAction Stop | Where-Object { $_.ConfigManagerErrorCode -ne 0 } | ForEach-Object {
                [PSCustomObject]@{ name = "$($_.Name)"; code = "$($_.ConfigManagerErrorCode)" }
            })
        } catch { }

        $audio = $null
        try {
            $playback = (Get-CimInstance Win32_SoundDevice -ErrorAction Stop | Where-Object { $_.Status -eq 'OK' } | Select-Object -First 1).Name
            $audio = [PSCustomObject]@{ playback = "$playback" }
        } catch { }

        # Hardware-accelerated GPU Scheduling (system-wide setting, not per-adapter)
        $hagsEnabled = $null
        try {
            $hw = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\GraphicsDrivers' -Name 'HwSchMode' -ErrorAction Stop).HwSchMode
            $hagsEnabled = if ($hw -eq 2) { "Enabled" } else { "Disabled" }
        } catch { }

        $gpus = @()
        try {
            $gpus = @(Get-CimInstance Win32_VideoController -ErrorAction Stop | ForEach-Object {
                $vram = if ($vramByKey.ContainsKey($_.Name)) { $vramByKey[$_.Name] } elseif ($_.AdapterRAM) { $_.AdapterRAM } else { 0 }
                [PSCustomObject]@{
                    name   = "$($_.Name)"
                    drv    = "$($_.DriverVersion)"
                    radeon = if ($_.Name -match 'AMD|Radeon') { $radeonVer } else { "" }
                    hres   = if ($_.CurrentHorizontalResolution) { [int]$_.CurrentHorizontalResolution } else { 0 }
                    vres   = if ($_.CurrentVerticalResolution) { [int]$_.CurrentVerticalResolution } else { 0 }
                    hz     = if ($_.CurrentRefreshRate) { [int]$_.CurrentRefreshRate } else { 0 }
                    vram   = if ($vram) { [math]::Round($vram / 1GB, 1) } else { 0 }
                }
            })
        } catch { }
        $mons = @()
        try {
            $mons = @(Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction Stop | ForEach-Object {
                if ($_.UserFriendlyName) {
                    ([System.Text.Encoding]::ASCII.GetString(($_.UserFriendlyName | Where-Object { $_ -ne 0 }))).Trim()
                }
            } | Where-Object { $_ })
        } catch { }

        Write-Host "      - GPU and display info via DXDIAG (this can take up to 30 seconds)" -ForegroundColor DarkGray
        # Per-output display -> GPU mapping via dxdiag (waits up to 30s; falls back to WMI data above)
        $displays = @()
        try {
            $dxPath = "$env:TEMP\pchh_dxdiag.xml"
            Remove-Item $dxPath -Force -ErrorAction SilentlyContinue
            Start-Process dxdiag -ArgumentList "/whql:off", "/x", "`"$dxPath`"" -WindowStyle Hidden
            for ($i = 0; $i -lt 30 -and -not (Test-Path $dxPath); $i++) { Start-Sleep -Seconds 1 }
            Start-Sleep -Seconds 1
            if (Test-Path $dxPath) {
                [xml]$dx = Get-Content $dxPath -Raw
                $displays = @($dx.DxDiag.DisplayDevices.DisplayDevice | ForEach-Object {
                    $mon = "$($_.MonitorName)"
                    if (-not $mon) { $mon = "$($_.MonitorModel)" }
                    [PSCustomObject]@{
                        gpu  = "$($_.CardName)"
                        mon  = $mon.Trim()
                        mode = $(
                            $raw = "$($_.CurrentMode)".Trim()
                            if ($raw -match '^(.*?)\s*\((\d+) bit\)\s*\((\d+Hz)\)\s*$') { "$($Matches[1]) ($($Matches[3]), $($Matches[2])-bit)" } else { $raw }
                        )
                    }
                } | Where-Object { $_.gpu })
                Remove-Item $dxPath -Force -ErrorAction SilentlyContinue
            }
        } catch { }

        # Running processes grouped by name (top 150 by memory)
        $procs = @()
        try {
            $procs = @(Get-Process -ErrorAction Stop | Group-Object ProcessName | ForEach-Object {
                [PSCustomObject]@{
                    name = $_.Name
                    cnt  = $_.Count
                    mem  = [math]::Round((($_.Group | Measure-Object WorkingSet64 -Sum).Sum) / 1MB)
                }
            } | Sort-Object mem -Descending)
        } catch { }

        Write-Host "      - Security (Defender status, exclusions, hosts file, startup, browser extensions)" -ForegroundColor DarkGray
        $security = $null
        try {
            # Defender status + scan history
            $mpStatus = $null
            try { $mpStatus = Get-MpComputerStatus -ErrorAction Stop } catch { }
            $defender = if ($mpStatus) {
                [PSCustomObject]@{
                    rtp        = "$($mpStatus.RealTimeProtectionEnabled)"
                    lastQuick  = if ($mpStatus.QuickScanEndTime) { $mpStatus.QuickScanEndTime.ToString("dd/MM/yyyy HH:mm") } else { "" }
                    lastFull   = if ($mpStatus.FullScanEndTime) { $mpStatus.FullScanEndTime.ToString("dd/MM/yyyy HH:mm") } else { "" }
                    sigAge     = "$($mpStatus.AntivirusSignatureAge)"
                    sigVersion = "$($mpStatus.AntivirusSignatureVersion)"
                }
            } else { $null }

            $threats = @()
            try {
                $threats = @(Get-MpThreatDetection -ErrorAction Stop | Select-Object -First 25 | ForEach-Object {
                    [PSCustomObject]@{
                        name = "$($_.ThreatName)"
                        time = $_.InitialDetectionTime.ToString("dd/MM/yyyy HH:mm")
                        act  = "$($_.ActionSuccess)"
                    }
                })
            } catch { }

            # Exclusions with dangerous-pattern flagging (paths genericized to strip username)
            $genericize = { param($p) if ("$p") { "$p" -replace [regex]::Escape("$env:USERPROFILE"), "%USERPROFILE%" -replace 'C:\\Users\\[^\\]+', "C:\Users\<user>" } else { "$p" } }
            $exclusions = @()
            $exclFlags = @()
            try {
                $mpPref = Get-MpPreference -ErrorAction Stop
                foreach ($p in $mpPref.ExclusionPath) {
                    $g = & $genericize $p
                    $exclusions += "Path: $g"
                    if ($p -match '^[A-Za-z]:\\?$') { $exclFlags += "Entire drive excluded: $g" }
                    elseif ($p -match '\\(Temp|AppData\\Roaming)\\?$') { $exclFlags += "Broad system folder excluded: $g" }
                }
                foreach ($e in $mpPref.ExclusionExtension) {
                    $exclusions += "Extension: .$e"
                    if ($e -match '^(exe|dll|scr|bat|ps1)$') { $exclFlags += "Executable file type excluded: .$e" }
                }
                foreach ($pr in $mpPref.ExclusionProcess) { $exclusions += "Process: $(& $genericize $pr)" }
            } catch { }

            # Hosts file: count custom entries, flag known-domain redirects
            $hostsCustom = 0
            $hostsFlags = @()
            try {
                $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
                $watchDomains = 'windowsupdate\.microsoft\.com|update\.microsoft\.com|\.microsoft\.com$|malwarebytes\.com|windowsdefender|virustotal\.com|avast\.com|kaspersky\.com|mcafee\.com|norton\.com'
                Get-Content $hostsPath -ErrorAction Stop | ForEach-Object {
                    $line = $_.Trim()
                    if ($line -and -not $line.StartsWith('#')) {
                        $hostsCustom++
                        if ($line -match $watchDomains -and $line -notmatch '^\s*(0\.0\.0\.0|127\.0\.0\.1)\s') {
                            $hostsFlags += $line
                        } elseif ($line -match $watchDomains) {
                            $hostsFlags += "$line (redirected to loopback/null - likely intentional block)"
                        }
                    }
                }
            } catch { }

            # Suspicious startup entries: no publisher/signature, or launching from Temp/AppData with odd naming
            $startupFlags = @()
            try {
                $runKeys = @(
                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
                    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
                )
                foreach ($rk in $runKeys) {
                    if (Test-Path $rk) {
                        $props = Get-ItemProperty -Path $rk -ErrorAction SilentlyContinue
                        $props.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | ForEach-Object {
                            $val = "$($_.Value)"
                            $exePath = ($val -replace '^"?([^"]+\.exe)"?.*$', '$1')
                            $suspicious = $false
                            $reason = ""
                            if ($exePath -match '\\(Temp|AppData\\Local\\Temp)\\') { $suspicious = $true; $reason = "runs from Temp folder" }
                            elseif (Test-Path $exePath -ErrorAction SilentlyContinue) {
                                try {
                                    $sig = Get-AuthenticodeSignature -FilePath $exePath -ErrorAction Stop
                                    if ($sig.Status -ne 'Valid') { $suspicious = $true; $reason = "unsigned or invalid signature" }
                                } catch { }
                            }
                            if ($suspicious) { $startupFlags += "$($_.Name): $reason" }
                        }
                    }
                }
            } catch { }

            # Firewall status per profile
            $firewall = @()
            try {
                $firewall = @(Get-NetFirewallProfile -ErrorAction Stop | ForEach-Object {
                    [PSCustomObject]@{ profile = "$($_.Name)"; enabled = "$($_.Enabled)" }
                })
            } catch { }

            # Scheduled Tasks: user-created, non-Microsoft, enabled - flag unsigned/Temp-run actions
            try {
                $tasks = Get-ScheduledTask -ErrorAction Stop | Where-Object {
                    $_.State -ne 'Disabled' -and $_.TaskPath -notmatch '\\Microsoft\\' -and $_.TaskPath -notmatch '\\Windows\\'
                }
                foreach ($t in $tasks) {
                    $act = ($t.Actions | Where-Object { $_.Execute } | Select-Object -First 1).Execute
                    if (-not $act) { continue }
                    $exePath = $act -replace '^"?([^"]+)"?.*$', '$1'
                    $suspicious = $false
                    $reason = ""
                    if ($exePath -match '\\(Temp|AppData\\Local\\Temp)\\') { $suspicious = $true; $reason = "runs from Temp folder" }
                    elseif (Test-Path $exePath -ErrorAction SilentlyContinue) {
                        try {
                            $sig = Get-AuthenticodeSignature -FilePath $exePath -ErrorAction Stop
                            if ($sig.Status -ne 'Valid') { $suspicious = $true; $reason = "unsigned or invalid signature" }
                        } catch { }
                    }
                    if ($suspicious) { $startupFlags += "Scheduled task '$($t.TaskName)': $reason" }
                }
            } catch { }

            # Startup folder shortcuts (both all-users and current user)
            try {
                $startupDirs = @("$env:ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp", "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\StartUp")
                foreach ($sd in $startupDirs) {
                    if (-not (Test-Path $sd)) { continue }
                    Get-ChildItem -Path $sd -Filter "*.lnk" -ErrorAction SilentlyContinue | ForEach-Object {
                        try {
                            $sh = New-Object -ComObject WScript.Shell
                            $target = $sh.CreateShortcut($_.FullName).TargetPath
                            if ($target -match '\\(Temp|AppData\\Local\\Temp)\\') {
                                $startupFlags += "Startup shortcut '$($_.BaseName)': runs from Temp folder"
                            }
                        } catch { }
                    }
                }
            } catch { }

            # Services set to Automatic that are not Running
            $stalledServices = @()
            try {
                $stalledServices = @(Get-CimInstance Win32_Service -Filter "StartMode='Auto' AND State!='Running'" -ErrorAction Stop | ForEach-Object {
                    [PSCustomObject]@{ name = "$($_.DisplayName)"; state = "$($_.State)" }
                } | Select-Object -First 20)
            } catch { }

            # Browser extensions: Chrome + Edge, all profiles, name only (no IDs, no sync data)
            $extensions = @()
            try {
                $browserRoots = @(
                    @{ browser = "Chrome";    root = "$env:LOCALAPPDATA\Google\Chrome\User Data" },
                    @{ browser = "Edge";      root = "$env:LOCALAPPDATA\Microsoft\Edge\User Data" },
                    @{ browser = "Brave";     root = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data" },
                    @{ browser = "Opera";     root = "$env:APPDATA\Opera Software\Opera Stable" },
                    @{ browser = "Opera GX";  root = "$env:APPDATA\Opera Software\Opera GX Stable" },
                    @{ browser = "Vivaldi";   root = "$env:LOCALAPPDATA\Vivaldi\User Data" }
                )
                foreach ($b in $browserRoots) {
                    if (-not (Test-Path $b.root)) { continue }
                    $profiles = @(Get-ChildItem -Path $b.root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'Default' -or $_.Name -match '^Profile \d+$' })
                    if ($profiles.Count -eq 0 -and (Test-Path (Join-Path $b.root "Extensions"))) {
                        # Opera / Opera GX keep the Extensions folder directly under the root (no Default subfolder)
                        $profiles = @([PSCustomObject]@{ FullName = $b.root; Name = "Default" })
                    }
                    foreach ($prof in $profiles) {
                        $extDir = Join-Path $prof.FullName "Extensions"
                        if (-not (Test-Path $extDir)) { continue }

                        # Cross-reference against the browser's own extension state, not just what's
                        # on disk - Chromium doesn't always clean up an extension's folder immediately
                        # after uninstall/disable, which can make a removed extension look "installed".
                        # state: 0 = disabled, 1 = enabled. Only IDs present here with state=1 count.
                        $enabledIds = $null
                        foreach ($prefFile in @('Secure Preferences', 'Preferences')) {
                            $prefPath = Join-Path $prof.FullName $prefFile
                            if (-not (Test-Path $prefPath)) { continue }
                            try {
                                $prefs = Get-Content $prefPath -Raw -ErrorAction Stop | ConvertFrom-Json
                                if ($prefs.extensions -and $prefs.extensions.settings) {
                                    $enabledIds = @($prefs.extensions.settings.PSObject.Properties | Where-Object { $_.Value.state -eq 1 } | ForEach-Object { $_.Name })
                                    break
                                }
                            } catch { }
                        }

                        Get-ChildItem -Path $extDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                            $extId = $_.Name
                            if ($enabledIds -and $enabledIds -notcontains $extId) { return }
                            $verDir = Get-ChildItem -Path $_.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
                            if (-not $verDir) { return }
                            $manifestPath = Join-Path $verDir.FullName "manifest.json"
                            if (-not (Test-Path $manifestPath)) { return }
                            try {
                                $manifest = Get-Content $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
                                $name = "$($manifest.name)"
                                if ($name -match '^__MSG_(.+)__$') {
                                    $key = $Matches[1]
                                    $locale = if ($manifest.default_locale) { $manifest.default_locale } else { "en" }
                                    $msgPath = Join-Path $verDir.FullName "_locales\$locale\messages.json"
                                    if (Test-Path $msgPath) {
                                        try {
                                            $msgs = Get-Content $msgPath -Raw -ErrorAction Stop | ConvertFrom-Json
                                            if ($msgs.$key.message) { $name = "$($msgs.$key.message)" }
                                        } catch { }
                                    }
                                }
                                if ($name -and $name -notmatch '^__MSG_') {
                                    $extensions += [PSCustomObject]@{ browser = $b.browser; profile = $prof.Name; name = $name }
                                }
                            } catch { }
                        }
                    }
                }
            } catch { }

            # Firefox: extensions.json per profile (different storage format to Chromium)
            try {
                $ffRoot = "$env:APPDATA\Mozilla\Firefox\Profiles"
                if (Test-Path $ffRoot) {
                    Get-ChildItem -Path $ffRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '\.default' } | ForEach-Object {
                        $extJsonPath = Join-Path $_.FullName "extensions.json"
                        if (-not (Test-Path $extJsonPath)) { return }
                        try {
                            $extData = Get-Content $extJsonPath -Raw -ErrorAction Stop | ConvertFrom-Json
                            foreach ($addon in $extData.addons) {
                                if ($addon.type -ne 'extension' -or $addon.active -ne $true) { continue }
                                $name = if ($addon.defaultLocale -and $addon.defaultLocale.name) { "$($addon.defaultLocale.name)" } else { "$($addon.id)" }
                                if ($name) { $extensions += [PSCustomObject]@{ browser = "Firefox"; profile = $_.Name; name = $name } }
                            }
                        } catch { }
                    }
                }
            } catch { }

            $security = [PSCustomObject]@{
                defender         = $defender
                threats          = $threats
                exclusions       = $exclusions
                exclFlags        = $exclFlags
                hostsCustom      = $hostsCustom
                hostsFlags       = $hostsFlags
                startupFlags     = $startupFlags
                extensions       = $extensions
                firewall         = $firewall
                stalledServices  = $stalledServices
            }
        } catch { }

        Write-Host "      - Network adapters, Memory and Running processes" -ForegroundColor DarkGray
        # Network adapters (no IPs, MACs or SSIDs collected)
        $net = $null
        try {
            $adapters = @(Get-NetAdapter -Physical -ErrorAction Stop | ForEach-Object {
                [PSCustomObject]@{
                    name   = "$($_.Name)"
                    desc   = "$($_.InterfaceDescription)"
                    status = "$($_.Status)"
                    speed  = "$($_.LinkSpeed)"
                    media  = "$($_.PhysicalMediaType)"
                }
            })
            $vpns = @(Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
                -not $_.Physical -and (
                    $_.Status -eq 'Up' -or
                    "$($_.InterfaceDescription) $($_.Name)" -match 'TAP|Wintun|WireGuard|OpenVPN|Tailscale|Nord|ExpressVPN|Proton|Surfshark|Mullvad|ZeroTier|Hamachi|Radmin|VPN'
                ) -and "$($_.InterfaceDescription)" -notmatch 'WAN Miniport|Bluetooth|Loopback|Kernel Debug'
            } | ForEach-Object {
                [PSCustomObject]@{
                    name   = "$($_.Name)"
                    desc   = "$($_.InterfaceDescription)"
                    status = "$($_.Status)"
                }
            })
            $wifi = $null
            $wl = netsh wlan show interfaces 2>$null
            if ($wl) {
                $wf = @{}
                foreach ($line in $wl) {
                    if ($line -match '^\s*(Radio type|Band|Channel|Signal|Authentication|Receive rate \(Mbps\)|Transmit rate \(Mbps\))\s*:\s*(.+)$') {
                        $wf[$Matches[1]] = $Matches[2].Trim()
                    }
                }
                if ($wf['Signal']) {
                    $wifi = [PSCustomObject]@{
                        signal  = "$($wf['Signal'])"
                        band    = "$($wf['Band'])"
                        channel = "$($wf['Channel'])"
                        radio   = "$($wf['Radio type'])"
                        auth    = "$($wf['Authentication'])"
                        rx      = "$($wf['Receive rate (Mbps)'])"
                        tx      = "$($wf['Transmit rate (Mbps)'])"
                    }
                }
            }
            $net = [PSCustomObject]@{ adapters = $adapters; vpns = $vpns; wifi = $wifi }
        } catch { }

        # Memory usage at time of capture (physical + commit charge)
        $memuse = $null
        try {
            $osm = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
            $memuse = [PSCustomObject]@{
                pt = [math]::Round($osm.TotalVisibleMemorySize / 1MB, 1)
                pu = [math]::Round(($osm.TotalVisibleMemorySize - $osm.FreePhysicalMemory) / 1MB, 1)
                ct = [math]::Round($osm.TotalVirtualMemorySize / 1MB, 1)
                cu = [math]::Round(($osm.TotalVirtualMemorySize - $osm.FreeVirtualMemory) / 1MB, 1)
            }
        } catch { }

        # Minidump info for the viewer
        $dumps = @()
        if ($dmpfound) {
            $dumps = @(Get-ChildItem -Path $source -ErrorAction SilentlyContinue | ForEach-Object {
                [PSCustomObject]@{
                    n = $_.Name
                    d = $_.LastWriteTime.ToString("dd/MM/yyyy HH:mm")
                    z = "{0:N1} MB" -f ($_.Length / 1MB)
                }
            })
        }

        # JSON payloads ("</" escaped so text cannot close the script tag)
        $json      = (ConvertTo-Json @($recs) -Compress -Depth 3).Replace('</', '<\/')
        $sysJson   = if ($sysEvents.Count -gt 0) { (ConvertTo-Json @($sysEvents) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $dumpsJson = if ($dumps.Count -gt 0) { (ConvertTo-Json @($dumps) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $gpusJson = if ($gpus.Count -gt 0) { (ConvertTo-Json @($gpus) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $hagsJson = if ($hagsEnabled) { "`"$hagsEnabled`"" } else { 'null' }
        $monsJson = if ($mons.Count -gt 0) { (ConvertTo-Json @($mons) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $displaysJson = if ($displays.Count -gt 0) { (ConvertTo-Json @($displays) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $procsJson = if ($procs.Count -gt 0) { (ConvertTo-Json @($procs) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $netJson = if ($net) { (ConvertTo-Json $net -Compress -Depth 4).Replace('</', '<\/') } else { 'null' }
        $securityJson = if ($security) { (ConvertTo-Json $security -Compress -Depth 5).Replace('</', '<\/') } else { 'null' }
        $hotfixesJson = if ($hotfixes.Count -gt 0) { (ConvertTo-Json @($hotfixes) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $wuHistoryJson = if ($wuHistory.Count -gt 0) { (ConvertTo-Json @($wuHistory) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $winUpdateInfo = [PSCustomObject]@{ pendingReboot = $pendingReboot; serviceStatus = $wuServiceStatus }
        $winUpdateJson = (ConvertTo-Json $winUpdateInfo -Compress).Replace('</', '<\/')
        $devErrorsJson = if ($devErrors.Count -gt 0) { (ConvertTo-Json @($devErrors) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $audioJson = if ($audio) { (ConvertTo-Json $audio -Compress).Replace('</', '<\/') } else { 'null' }
        $memuseJson = if ($memuse) { (ConvertTo-Json $memuse -Compress).Replace('</', '<\/') } else { 'null' }
        $ramJson = if ($ram.Count -gt 0) { (ConvertTo-Json @($ram) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $smartJson = if ($smart.Count -gt 0) { (ConvertTo-Json @($smart) -Compress -Depth 3).Replace('</', '<\/') } else { '[]' }
        $dirtyJson = if ($dirtyVols.Count -gt 0) { (ConvertTo-Json @($dirtyVols) -Compress).Replace('</', '<\/') } else { '[]' }
        $diskLayoutJson = if ($diskLayout.Count -gt 0) { (ConvertTo-Json @($diskLayout) -Compress -Depth 4).Replace('</', '<\/') } else { '[]' }
        $specsRaw = Get-Content -Path $infofile -Raw -ErrorAction SilentlyContinue
        if ($null -eq $specsRaw) { $specsRaw = "" }
        $specsJson = (ConvertTo-Json "$specsRaw" -Compress).Replace('</', '<\/')

        $genStamp = (Get-Date).ToString("dd/MM/yyyy HH:mm")
        $viewerHtml = $viewerTemplate.Replace('/*__VER__*/""', "`"$scriptVersion`"").Replace('/*__GEN__*/""', "`"$genStamp`"").Replace('/*__DATA__*/[]', $json).Replace('/*__SPECS__*/""', $specsJson).Replace('/*__DUMPS__*/[]', $dumpsJson).Replace('/*__SYSEVT__*/[]', $sysJson).Replace('/*__SMART__*/[]', $smartJson).Replace('/*__DIRTY__*/[]', $dirtyJson).Replace('/*__DISKLAYOUT__*/[]', $diskLayoutJson).Replace('/*__RAM__*/[]', $ramJson).Replace('/*__GPUS__*/[]', $gpusJson).Replace('/*__HAGS__*/null', $hagsJson).Replace('/*__MONS__*/[]', $monsJson).Replace('/*__DISPLAYS__*/[]', $displaysJson).Replace('/*__PROCS__*/[]', $procsJson).Replace('/*__MEMUSE__*/null', $memuseJson).Replace('/*__NET__*/null', $netJson).Replace('/*__SECURITY__*/null', $securityJson).Replace('/*__HOTFIXES__*/[]', $hotfixesJson).Replace('/*__WUHISTORY__*/[]', $wuHistoryJson).Replace('/*__WINUPDATE__*/null', $winUpdateJson).Replace('/*__DEVERR__*/[]', $devErrorsJson).Replace('/*__AUDIO__*/null', $audioJson)
        Set-Content -Path $reliability_html_path -Value $viewerHtml -Encoding UTF8
    }
    catch {
        $errors.reliability = $true
        functionerror
    }

    Write-Host -NoNewline -ForegroundColor Green "$(cmark)"
    Write-Host " Diagnostics collected"

    compression
}



# Compresses files
function compression {
    Write-Host ""
    Write-Host "[4/4] Compressing everything into one zip.." -ForegroundColor Blue

    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl" -Name "DisplayParameters" -Value 1 -Type DWord -Force | Out-Null

    $filesToCompress = @($infofile, $sys_eventlog_path, $reliability_csv_path, $reliability_html_path)

    if ($dmpfound) {
        $filesToCompress += Get-ChildItem -Path $source
    }

    try {
        Invoke-WithoutProgress {
            Compress-Archive -Path $filesToCompress -CompressionLevel Optimal -DestinationPath $ziptar -Force | Out-Null
        }
    }
    catch {


        Write-Host ""
        Write-Host "     Unable to compress files..." -ForegroundColor Red
        Write-Host "     Re-run the script to attempt to fix the issue." -ForegroundColor Red
        Write-Host ""

        $errors.Compress = $true
        functionerror
    }

    Remove-Item -Path $infofile, $sys_eventlog_path, $reliability_csv_path, $reliability_html_path -Force -Recurse -ErrorAction SilentlyContinue > $null 2>&1

    Write-Host -NoNewline -ForegroundColor Green "$(cmark)"
    Write-Host " Zip created"

    eof
}

function eof {
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor DarkGreen
    Write-Host "  DONE - your report is ready to share" -ForegroundColor Green
    Write-Host "==================================================" -ForegroundColor DarkGreen
    Write-Host ""
    Write-Host -NoNewline "  Zip file:   " -ForegroundColor Gray
    Write-Host "$ziptar"
    Write-Host ""
    Write-Host "  The zip is already on your clipboard -" -ForegroundColor Gray
    Write-Host "  just press Ctrl+V in Discord to attach it." -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Want to see the report yourself? Open the zip and" -ForegroundColor Gray
    Write-Host "  double-click triage-report.html - it opens in your browser." -ForegroundColor Gray
    Start-Process explorer.exe -ArgumentList $File
    $eofcomplete = $true

    endmessage
}

function functionerror {
    Write-Host -NoNewline -ForegroundColor Red "$(xmark)"

    if ($errors.Compress -eq "true") {
        Write-Host " There was an error during compression.."
    }
    elseif ($errors.event -eq "true") {
        Write-Host "There was an error while grabbing the event logs.."
    }
    elseif ($errors.fileCreate -eq "true") {
        Write-Host "There was an error while creating files.."
    }
    elseif ($errors.reliability -eq "true") {
        Write-Host " There was an error while grabbing reliability history.."
    }

    Write-Host -NoNewline -ForegroundColor White "Error:"
    Write-Host " $_" -ForegroundColor Red

    Remove-Item -Path $infofile, $sys_eventlog_path, $reliability_csv_path, $reliability_html_path -Force -Recurse -ErrorAction SilentlyContinue > $null 2>&1

    endmessage
}

function endmessage {
    Write-Host ""
    Write-Host "Press any key to exit.."

    if ($eofcomplete) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Clipboard]::SetFileDropList([System.Collections.Specialized.StringCollection]@($ziptar))
    }
        
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        Stop-Process -Id $PID -Force
}

dmpcheck

