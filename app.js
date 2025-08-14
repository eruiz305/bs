const fmtUSD=new Intl.NumberFormat(undefined,{style:"currency",currency:"USD"});
const toISO=(d)=>d?new Date(d).toISOString().slice(0,10):"";
const parseDate=(v)=>{if(!v)return null;if(v instanceof Date)return v;if(typeof v==="number"){const e=new Date(Date.UTC(1899,11,30));return new Date(e.getTime()+v*86400000)}const d=new Date(v);return isNaN(d)?null:d};
const eom=(date)=>new Date(date.getFullYear(),date.getMonth()+1,0);
const addMonths=(date,n)=>new Date(date.getFullYear(),date.getMonth()+n,date.getDate());
const round2=(n)=>Math.round((Number(n)+Number.EPSILON)*100)/100;
const uid=()=>Math.random().toString(36).slice(2,10);
const filterUsableCols=(cols)=>cols.filter(c=>!c.toLowerCase().endsWith("ky"));
const guessMap=(cols)=>{const lc=cols.map(c=>c.toLowerCase());const find=(...names)=>{for(let n of names){const i=lc.indexOf(n.toLowerCase());if(i>=0)return cols[i]}for(let n of names){const i=lc.findIndex(c=>c.includes(n.toLowerCase()));if(i>=0)return cols[i]}return""};return{amount:find("accamt","amount","gross amount"),vendor:find("co_nam","vendor"),invoiceNumber:find("invnum","voucher","invoice"),invoiceDate:find("invdat","invoice date"),description:find("invdsc","description","memo"),accamtdsc:find("accamtdsc","line desc"),seg2:find("segnumtwo","seg2","segment 2","account"),seg3:find("segnumthr","seg3","segment 3","department"),seg4:find("segnumfou","seg4","segment 4","location")}};
function downloadTXT(filename,rows){if(!rows?.length)return;const txt=rows.map(r=>r.map(x=>x==null?"":String(x)).join("\t")).join("\n");const blob=new Blob([txt],{type:"text/plain;charset=utf-8;"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=filename.endsWith(".txt")?filename:filename+".txt";a.click()}
function downloadBlob(name,content){const blob=new Blob([content],{type:"text/plain;charset=utf-8;"});const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;a.click()}
const el=(id)=>document.getElementById(id),setText=(id,txt)=>{el(id).textContent=txt};

const STORAGE="so-bs-amort-v5";
let profile={first:"",last:"",email:""};
let wbSheets={},wbNames=[];
let apRows=[],apCols=[],mapCols={},groups=[],detailGroup=null,summaryTB=[],detailTB=[];
let acctSheet="",acctCols={seg2:"",seg3:"",seg4:"",desc:"",active:""},acctMap=new Map(),acctList=[];
let actionsByKey={},items=[],reclassItems=[];

let periodEnd="",fiscalYY="",actualMM="",seqStart="01",journalTitle="Standard Amortization Entry";
let defaults={fyy:"",amm:"",amemo:"{{vendor}} {{invnum}} amortization ({{start}}–{{end}})",jnltitle:"Standard Amortization Entry",groups:[]};
let mode="amort";

function save(){localStorage.setItem(STORAGE,JSON.stringify({profile,actionsByKey,items,reclassItems,periodEnd,fiscalYY,actualMM,seqStart,journalTitle,acctSheet,acctCols,defaults}))}
function load(){try{const s=localStorage.getItem(STORAGE);if(s){const o=JSON.parse(s);profile=o.profile||profile;actionsByKey=o.actionsByKey||{};items=o.items||[];reclassItems=o.reclassItems||[];periodEnd=o.periodEnd||"";fiscalYY=o.fiscalYY||"";actualMM=o.actualMM||"";seqStart=o.seqStart||"01";journalTitle=o.journalTitle||journalTitle;acctSheet=o.acctSheet||"";acctCols=o.acctCols||acctCols;defaults=o.defaults||defaults;if(!defaults.groups)defaults.groups=[]}}catch{}}


function renderProfile(){const name=profile.first&&profile.last?`${profile.first} ${profile.last}`:"Not set";setText("user-line",`${name} · ${profile.email||""}`);setText("user-badge",name?`${name}`:"Not signed in");el("login").style.display=(profile.email? "none":"flex")}
function openProfile(){el("u-first").value=profile.first||"";el("u-last").value=profile.last||"";el("u-email").value=profile.email||"";el("login").style.display="flex"}
function applyProfile(){profile.first=el("u-first").value.trim();profile.last=el("u-last").value.trim();profile.email=el("u-email").value.trim();save();renderProfile()}

function setMode(m){
  mode=m;
  el("amort-view").style.display=m==="amort"?"block":"none";
  el("activity-view").style.display=m==="activity"?"block":"none";
  el("settings-view").style.display=m==="settings"?"block":"none";
  el("recon-view").style.display=m==="recon"?"block":"none";
  el("tab-amort").classList.toggle("active",m==="amort");
  el("tab-activity").classList.toggle("active",m==="activity");
  el("tab-settings").classList.toggle("active",m==="settings");
  el("tab-recon").classList.toggle("active",m==="recon");
}

function renderMapUI(){
  const fields=[["amount","Amount"],["vendor","Vendor"],["invoiceNumber","Invoice #"],["invoiceDate","Invoice Date"],["description","Header Desc"],["accamtdsc","Line Desc (accamtdsc)"],["seg2","Seg2"],["seg3","Seg3"],["seg4","Seg4"]];
  const container=el("map-area");container.innerHTML="";
  fields.forEach(([k,label])=>{const wrap=document.createElement("div");const lab=document.createElement("label");lab.textContent=label;wrap.appendChild(lab);const sel=document.createElement("select");const opt0=document.createElement("option");opt0.value="";opt0.disabled=true;opt0.selected=!mapCols[k];opt0.textContent="Select";sel.appendChild(opt0);apCols.forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;o.selected=mapCols[k]===c;sel.appendChild(o)});sel.addEventListener("change",()=>{mapCols[k]=sel.value;save();renderGroups()});wrap.appendChild(sel);container.appendChild(wrap)})
}

function mappedRows(){if(!apRows.length)return[];const m=mapCols||{};return apRows.map((r,idx)=>({id:(r[m.invoiceNumber]?String(r[m.invoiceNumber]):"NO-INV-"+idx)+"-"+idx,amount:Number(r[m.amount]||0)||0,vendor:String(r[m.vendor]||""),invoiceNumber:String(r[m.invoiceNumber]||""),invoiceDate:parseDate(r[m.invoiceDate]),description:String(r[m.description]||""),accamtdsc:String(r[m.accamtdsc]||""),seg2:String(r[m.seg2]||""),seg3:String(r[m.seg3]||""),seg4:String(r[m.seg4]||""),raw:r}))}
function buildGroups(){const rows=mappedRows();const by=new Map();rows.forEach(r=>{const key=r.invoiceNumber||"NO-INV";if(!by.has(key))by.set(key,{key,vendor:r.vendor,invoiceNumber:r.invoiceNumber,date:r.invoiceDate,rows:[],amount:0});const g=by.get(key);g.rows.push(r);g.amount=round2(g.amount+(r.amount||0));if(!g.vendor&&r.vendor)g.vendor=r.vendor;if(!g.date&&r.invoiceDate)g.date=r.invoiceDate;if(!actionsByKey[g.key])actionsByKey[g.key]={amortize:false,reclass:false}});groups=Array.from(by.values())}
function renderGroups(){buildGroups();const tbody=el("ap-groups");tbody.innerHTML="";groups.forEach(g=>{const tr=document.createElement("tr");tr.style.cursor="pointer";tr.addEventListener("dblclick",()=>openDetail(g));const td0=document.createElement("td");const chk=document.createElement("input");chk.type="checkbox";chk.checked=!!actionsByKey[g.key]?.amortize;chk.addEventListener("change",()=>{actionsByKey[g.key]={...(actionsByKey[g.key]||{}),amortize:chk.checked};renderSelectedTotal();updateActionButtons();save()});td0.appendChild(chk);tr.appendChild(td0);const td1=document.createElement("td");td1.textContent=g.invoiceNumber;td1.style.fontWeight="600";tr.appendChild(td1);const td2=document.createElement("td");td2.textContent=g.vendor;td2.title=g.vendor;td2.style.maxWidth="280px";td2.style.overflow="hidden";td2.style.textOverflow="ellipsis";td2.style.whiteSpace="nowrap";tr.appendChild(td2);const td3=document.createElement("td");td3.textContent=g.date?toISO(g.date):"";tr.appendChild(td3);const td4=document.createElement("td");td4.className="num";td4.textContent=fmtUSD.format(g.amount||0);tr.appendChild(td4);const td5=document.createElement("td");const pill=document.createElement("span");pill.className="pill";const r0=g.rows[0]||{};pill.textContent=(r0.seg2||"")+"-"+(r0.seg3||"")+"-"+(r0.seg4||"");td5.appendChild(pill);tr.appendChild(td5);const td6=document.createElement("td");const lines=document.createElement("span");lines.className="pill";lines.textContent=String(g.rows.length);lines.title="Double-click row to view invoice lines";td6.appendChild(lines);tr.appendChild(td6);const td7=document.createElement("td");const c1=document.createElement("input");c1.type="checkbox";c1.checked=!!actionsByKey[g.key]?.reclass;c1.addEventListener("change",()=>{actionsByKey[g.key]={...(actionsByKey[g.key]||{}),reclass:c1.checked};updateActionButtons();save()});td7.appendChild(c1);tr.appendChild(td7);const td8=document.createElement("td");const c2=document.createElement("input");c2.type="checkbox";c2.checked=!!actionsByKey[g.key]?.amortize;c2.addEventListener("change",()=>{actionsByKey[g.key]={...(actionsByKey[g.key]||{}),amortize:c2.checked};renderSelectedTotal();updateActionButtons();save()});td8.appendChild(c2);tr.appendChild(td8);tbody.appendChild(tr)});renderSelectedTotal();updateActionButtons();renderTBCheck()}
function openDetail(g){detailGroup=g;el("dlg-title").textContent=`Invoice ${g.invoiceNumber||""} — ${g.vendor||""}`;const body=el("dlg-body");body.innerHTML="";const t=document.createElement("table");t.innerHTML=`<thead><tr><th>Invoice Date</th><th>Header Desc</th><th>Line Desc (accamtdsc)</th><th>Acct Combo</th><th class="num">Amount</th></tr></thead>`;const tb=document.createElement("tbody");g.rows.forEach(r=>{const tr=document.createElement("tr");tr.innerHTML=`<td>${r.invoiceDate?toISO(r.invoiceDate):""}</td><td title="${r.description}">${r.description}</td><td title="${r.accamtdsc}">${r.accamtdsc}</td><td>${(r.seg2||"")}-${(r.seg3||"")}-${(r.seg4||"")}</td><td class="num">${fmtUSD.format(r.amount)}</td>`;tb.appendChild(tr)});t.appendChild(tb);body.appendChild(t);el("dlg").style.display="flex"}

function renderSelectedTotal(){const total=groups.filter(g=>actionsByKey[g.key]?.amortize).reduce((a,b)=>a+(b.amount||0),0);setText("sel-total",fmtUSD.format(total))}
function updateActionButtons(){const anyAm=groups.some(g=>actionsByKey[g.key]?.amortize);const anyRc=groups.some(g=>actionsByKey[g.key]?.reclass);el("add-to-amort").disabled=!anyAm;el("add-to-reclass").disabled=!anyRc}

function addSelectedToAmort(){const picks=groups.filter(g=>actionsByKey[g.key]?.amortize);const newItems=picks.map(g=>({id:uid(),source:{type:"AP",vendor:g.vendor,invoiceNumber:g.invoiceNumber,amount:g.amount,invoiceDate:g.date||new Date(),description:g.rows[0]?.description||"",seg2:g.rows[0]?.seg2||"",seg3:g.rows[0]?.seg3||"",seg4:g.rows[0]?.seg4||"",lines:g.rows},amort:{enabled:true,method:"straight",months:12,startDate:g.date||new Date(),postOn:"EOM",expSeg2:"",expSeg3:"",expSeg4:"",memoTemplate:defaults.amemo},asset:{seg2:g.rows[0]?.seg2||"",seg3:g.rows[0]?.seg3||"",seg4:g.rows[0]?.seg4||""},schedule:[]}));items=items.concat(newItems);save();renderItems()}
function addSelectedToReclass(){const picks=groups.filter(g=>actionsByKey[g.key]?.reclass);const newItems=picks.map(g=>({id:uid(),vendor:g.vendor,invoiceNumber:g.invoiceNumber,amount:g.amount,fromSeg2:g.rows[0]?.seg2||"",fromSeg3:g.rows[0]?.seg3||"",fromSeg4:g.rows[0]?.seg4||"",toSeg2:"",toSeg3:"",toSeg4:"",memo:`Reclass ${g.vendor||""} ${g.invoiceNumber||""}`}));reclassItems=reclassItems.concat(newItems);save();renderReclass()}

function buildSchedule(it){const src=it.source,a=it.amort;if(!a.enabled||!a.months)return[];const base=Number(src.amount)||0;const sDate=a.startDate?new Date(a.startDate):new Date();const first=a.postOn==="EOM"?eom(sDate):sDate;const exp=(a.expSeg2||"")+"-"+(a.expSeg3||"")+"-"+(a.expSeg4||"");const ast=(it.asset.seg2||src.seg2||"")+"-"+(it.asset.seg3||src.seg3||"")+"-"+(it.asset.seg4||src.seg4||"");const rows=[];const memo=(st,en)=> (a.memoTemplate||defaults.amemo).replace("{{vendor}}",src.vendor||"").replace("{{invnum}}",src.invoiceNumber||"").replace("{{start}}",toISO(st)).replace("{{end}}",toISO(en));if(a.method==="straight"){const per=round2(base/a.months);let acc=0;for(let i=0;i<a.months;i++){const d=a.postOn==="EOM"?eom(addMonths(first,i)):addMonths(first,i);const amt=i===a.months-1?round2(base-acc):per;acc=round2(acc+amt);rows.push({date:d,amount:amt,debitCombo:exp,creditCombo:ast,memo:memo(first,a.postOn==="EOM"?eom(addMonths(first,a.months-1)):addMonths(first,a.months-1))})}}else{const dim=new Date(first.getFullYear(),first.getMonth()+1,0).getDate();const firstDays=dim-first.getDate()+1;const daily=base/(firstDays+(a.months-1)*30);let acc=0;for(let i=0;i<a.months;i++){const d=a.postOn==="EOM"?eom(addMonths(first,i)):addMonths(first,i);const days=i===0?firstDays:30;const amt=i===a.months-1?round2(base-acc):round2(daily*days);acc=round2(acc+amt);rows.push({date:d,amount:amt,debitCombo:exp,creditCombo:ast,memo:memo(first,a.postOn==="EOM"?eom(addMonths(first,a.months-1)):addMonths(first,a.months-1))})}}return rows}

function acctKey(a,b,c){return `${String(a||"")}-${String(b||"")}-${String(c||"")}`}
function acctLookup(a,b,c){return acctMap.get(acctKey(a,b,c))||null}
function parseCombos(str){return (str||"").split(",").map(s=>s.trim()).filter(Boolean)}

function renderItems(){
  const host=el("items");host.innerHTML="";
  const anyMissing=items.some(x=>x.amort?.enabled&&(!x.schedule||!x.schedule.length));
  el("missing").style.display=anyMissing?"inline":"none";

  items.forEach(it=>{
    const card=document.createElement("div");card.className="card";
    const locked=Object.keys(it.reconciled||{}).some(p=>p<=closingPeriod);
    const h2=document.createElement("h2");h2.style.display="flex";h2.style.justifyContent="space-between";h2.style.alignItems="center";h2.innerHTML=`<span>${it.source.vendor} — ${it.source.invoiceNumber}</span>`;
    const rm=document.createElement("button");rm.className="ghost";rm.textContent="Remove";rm.onclick=()=>{if(locked){alert("Transaction is reconciled and cannot be edited");return;}items=items.filter(x=>x.id!==it.id);save();renderItems();};h2.appendChild(rm);card.appendChild(h2);

    const c=document.createElement("div");c.className="content";
    const meta=document.createElement("div");meta.className="small";meta.style.marginBottom="8px";meta.textContent=`${toISO(it.source.invoiceDate)} · ${fmtUSD.format(it.source.amount)} · ${(it.source.seg2||"")}-${(it.source.seg3||"")}-${(it.source.seg4||"")}`;c.appendChild(meta);
    const grid=document.createElement("div");grid.className="grid3";

    const months=document.createElement("div");months.innerHTML=`<label>Months</label>`;const monthsIn=document.createElement("input");monthsIn.type="number";monthsIn.min=1;monthsIn.value=it.amort.months;monthsIn.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');monthsIn.value=it.amort.months;return;}it.amort.months=Number(monthsIn.value||0);save();};months.appendChild(monthsIn);grid.appendChild(months);
    const start=document.createElement("div");start.innerHTML=`<label>Start Date</label>`;const startIn=document.createElement("input");startIn.type="date";startIn.value=toISO(it.amort.startDate);startIn.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');startIn.value=toISO(it.amort.startDate);return;}it.amort.startDate=new Date(startIn.value);save();};start.appendChild(startIn);grid.appendChild(start);
    const post=document.createElement("div");post.innerHTML=`<label>Post On</label>`;const postSel=document.createElement("select");postSel.innerHTML=`<option value="EOM">End of Month</option><option value="SameDay">Same Day</option>`;postSel.value=it.amort.postOn;postSel.onchange=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');postSel.value=it.amort.postOn;return;}it.amort.postOn=postSel.value;save();};post.appendChild(postSel);grid.appendChild(post);
    const method=document.createElement("div");method.innerHTML=`<label>Method</label>`;const methodSel=document.createElement("select");methodSel.innerHTML=`<option value="straight">Straight-line</option><option value="prorata">Prorata</option>`;methodSel.value=it.amort.method;methodSel.onchange=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');methodSel.value=it.amort.method;return;}it.amort.method=methodSel.value;save();};method.appendChild(methodSel);grid.appendChild(method);

    const exp2=document.createElement("div");exp2.innerHTML=`<label>Expense Account</label>`;const exp2in=document.createElement("input");exp2in.type="text";exp2in.value=it.amort.expSeg2||"";exp2in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');exp2in.value=it.amort.expSeg2||'';return;}it.amort.expSeg2=exp2in.value;save();mark();};exp2.appendChild(exp2in);grid.appendChild(exp2);
    const exp3=document.createElement("div");exp3.innerHTML=`<label>Department</label>`;const exp3in=document.createElement("input");exp3in.type="text";exp3in.value=it.amort.expSeg3||"";exp3in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');exp3in.value=it.amort.expSeg3||'';return;}it.amort.expSeg3=exp3in.value;save();mark();};exp3.appendChild(exp3in);grid.appendChild(exp3);
    const exp4=document.createElement("div");exp4.innerHTML=`<label>Location</label>`;const exp4in=document.createElement("input");exp4in.type="text";exp4in.value=it.amort.expSeg4||"";exp4in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');exp4in.value=it.amort.expSeg4||'';return;}it.amort.expSeg4=exp4in.value;save();mark();};exp4.appendChild(exp4in);grid.appendChild(exp4);

    const expStatus=document.createElement("div");expStatus.className="status";expStatus.innerHTML=`<span id="expIcon">✖</span><input id="expName" type="text" readonly placeholder="Account name" />`;grid.appendChild(expStatus);

    const hr=document.createElement("div");hr.style.gridColumn="1/-1";hr.style.borderTop="1px solid #e5e7eb";hr.style.marginTop="8px";grid.appendChild(hr);

    const as2=document.createElement("div");as2.innerHTML=`<label>Asset Account</label>`;const as2in=document.createElement("input");as2in.type="text";as2in.value=it.asset.seg2||"";as2in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');as2in.value=it.asset.seg2||'';return;}it.asset.seg2=as2in.value;save();mark();};as2.appendChild(as2in);grid.appendChild(as2);
    const as3=document.createElement("div");as3.innerHTML=`<label>Department</label>`;const as3in=document.createElement("input");as3in.type="text";as3in.value=it.asset.seg3||"";as3in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');as3in.value=it.asset.seg3||'';return;}it.asset.seg3=as3in.value;save();mark();};as3.appendChild(as3in);grid.appendChild(as3);
    const as4=document.createElement("div");as4.innerHTML=`<label>Location</label>`;const as4in=document.createElement("input");as4in.type="text";as4in.value=it.asset.seg4||"";as4in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');as4in.value=it.asset.seg4||'';return;}it.asset.seg4=as4in.value;save();mark();};as4.appendChild(as4in);grid.appendChild(as4);

    const astStatus=document.createElement("div");astStatus.className="status";astStatus.innerHTML=`<span id="astIcon">✖</span><input id="astName" type="text" readonly placeholder="Account name" />`;grid.appendChild(astStatus);

    const memo=document.createElement("div");memo.style.gridColumn="1/-1";memo.innerHTML=`<label>Memo Template</label>`;const memoIn=document.createElement("input");memoIn.type="text";memoIn.value=it.amort.memoTemplate||defaults.amemo;memoIn.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');memoIn.value=it.amort.memoTemplate||defaults.amemo;return;}it.amort.memoTemplate=memoIn.value;save();};memo.appendChild(memoIn);grid.appendChild(memo);

    const btns=document.createElement("div");btns.style.gridColumn="1/-1";
    const build=document.createElement("button");build.className="secondary";build.textContent="Build Schedule";
    build.onclick=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');return;}it.schedule=buildSchedule(it);save();renderItems();};
    const val=document.createElement("button");val.className="secondary";val.style.marginLeft="8px";val.textContent="Validate";
    val.onclick=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');return;}mark(true);};
    btns.appendChild(build);btns.appendChild(val);grid.appendChild(btns);

    const scNote=document.createElement("div");scNote.style.gridColumn="1/-1";c.appendChild(grid);

    function mark(){
      const exp=acctLookup(it.amort.expSeg2,it.amort.expSeg3,it.amort.expSeg4);
      const ast=acctLookup(it.asset.seg2,it.asset.seg3,it.asset.seg4);
      const expIcon=expStatus.querySelector("#expIcon"), expName=expStatus.querySelector("#expName");
      const astIcon=astStatus.querySelector("#astIcon"), astName=astStatus.querySelector("#astName");
      if(exp){expIcon.textContent="✔";expIcon.className="ok";expName.value=exp}else{expIcon.textContent="✖";expIcon.className="bad";expName.value=""}
      if(ast){astIcon.textContent="✔";astIcon.className="ok";astName.value=ast}else{astIcon.textContent="✖";astIcon.className="bad";astName.value=""}
      scNote.innerHTML="";
    }
    mark();

    if(it.schedule?.length){
      const sc=document.createElement("div");sc.className="scroll";sc.style.marginTop="12px";sc.style.maxHeight="200px";
      const tbl=document.createElement("table");
      tbl.innerHTML=`<thead><tr><th>Date</th><th class="num">Amount</th><th>Dr (Expense)</th><th>Cr (Asset)</th><th>Memo</th></tr></thead>`;
      const tb=document.createElement("tbody");
      it.schedule.forEach(r=>{
        const tr=document.createElement("tr");
        tr.innerHTML=`<td>${toISO(r.date)}</td><td class="num">${fmtUSD.format(r.amount)}</td><td>${r.debitCombo}</td><td>${r.creditCombo}</td><td title="${r.memo}" style="max-width:320px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r.memo}</td>`;
        tb.appendChild(tr);
      });
      tbl.appendChild(tb);sc.appendChild(tbl);c.appendChild(sc);
    }

    card.appendChild(c);
    if(locked){
      const ov=document.createElement('div');
      ov.style.position='absolute';ov.style.inset='0';
      ov.style.background='rgba(255,255,255,0.6)';
      ov.style.display='flex';ov.style.alignItems='center';ov.style.justifyContent='center';
      ov.style.fontWeight='700';ov.textContent='Reconciled';
      ov.onclick=()=>alert('Transaction is reconciled and cannot be edited');
      card.style.position='relative';card.appendChild(ov);
    }
    host.appendChild(card);
  });
  if(mode==='recon') renderReconciliation();
}

function renderReclass(){const host=el("reclass-items");host.innerHTML="";el("reclass-empty").style.display=reclassItems.length?"none":"block";reclassItems.forEach(it=>{const card=document.createElement("div");card.className="card";const locked=Object.keys(it.reconciled||{}).some(p=>p<=closingPeriod);const h2=document.createElement("h2");h2.style.display="flex";h2.style.justifyContent="space-between";h2.style.alignItems="center";h2.innerHTML=`<span>${it.vendor} — ${it.invoiceNumber}</span>`;const rm=document.createElement("button");rm.className="ghost";rm.textContent="Remove";rm.onclick=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');return;}reclassItems=reclassItems.filter(x=>x.id!==it.id);save();renderReclass()};h2.appendChild(rm);card.appendChild(h2);const c=document.createElement("div");c.className="content";const meta=document.createElement("div");meta.className="small";meta.style.marginBottom="8px";meta.textContent=`From ${(it.fromSeg2||"")}-${(it.fromSeg3||"")}-${(it.fromSeg4||"")} · Amount ${fmtUSD.format(it.amount)}`;c.appendChild(meta);const grid=document.createElement("div");grid.className="grid3";const t2=document.createElement("div");t2.innerHTML=`<label>To Account</label>`;const t2in=document.createElement("input");t2in.type="text";t2in.value=it.toSeg2||"";t2in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');t2in.value=it.toSeg2||'';return;}it.toSeg2=t2in.value;save();mark()};t2.appendChild(t2in);grid.appendChild(t2);const t3=document.createElement("div");t3.innerHTML=`<label>Department</label>`;const t3in=document.createElement("input");t3in.type="text";t3in.value=it.toSeg3||"";t3in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');t3in.value=it.toSeg3||'';return;}it.toSeg3=t3in.value;save();mark()};t3.appendChild(t3in);grid.appendChild(t3);const t4=document.createElement("div");t4.innerHTML=`<label>Location</label>`;const t4in=document.createElement("input");t4in.type="text";t4in.value=it.toSeg4||"";t4in.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');t4in.value=it.toSeg4||'';return;}it.toSeg4=t4in.value;save();mark()};t4.appendChild(t4in);grid.appendChild(t4);const memo=document.createElement("div");memo.style.gridColumn="1/-1";memo.innerHTML=`<label>Memo</label>`;const memoIn=document.createElement("input");memoIn.type="text";memoIn.value=it.memo||"";memoIn.oninput=()=>{if(locked){alert('Transaction is reconciled and cannot be edited');memoIn.value=it.memo||'';return;}it.memo=memoIn.value;save()};memo.appendChild(memoIn);grid.appendChild(memo);const note=document.createElement("div");note.style.gridColumn="1/-1";c.appendChild(grid);c.appendChild(note);function mark(){[t2in,t3in,t4in].forEach(x=>x.classList.remove("invalid"));const desc=acctLookup(it.toSeg2,it.toSeg3,it.toSeg4);note.innerHTML=desc?`<div class="valid-note">Target: ${desc}</div>`:(it.toSeg2||it.toSeg3||it.toSeg4)?`<div class="invalid-note">Target combo not found</div>`:""}mark();card.appendChild(c);if(locked){const ov=document.createElement('div');ov.style.position='absolute';ov.style.inset='0';ov.style.background='rgba(255,255,255,0.6)';ov.style.display='flex';ov.style.alignItems='center';ov.style.justifyContent='center';ov.style.fontWeight='700';ov.textContent='Reconciled';ov.onclick=()=>alert('Transaction is reconciled and cannot be edited');card.style.position='relative';card.appendChild(ov);}host.appendChild(card)})}

function renderReconciliation(){const host=el('recon-table');if(!host)return;const period=el('recon-period').value||closingPeriod||"";const t=document.createElement('table');t.innerHTML='<thead><tr><th>Type</th><th>Info</th><th>Account</th><th>Reconciled</th></tr></thead>';const tb=document.createElement('tbody');items.forEach(it=>{const combo=`${it.asset.seg2||it.source.seg2||""}-${it.asset.seg3||it.source.seg3||""}-${it.asset.seg4||it.source.seg4||""}`;const tr=document.createElement('tr');tr.innerHTML=`<td>Amort</td><td>${it.source.vendor} ${it.source.invoiceNumber}</td><td>${combo}</td>`;const td=document.createElement('td');const chk=document.createElement('input');chk.type='checkbox';chk.checked=!!(it.reconciled&&it.reconciled[period]);chk.onchange=()=>{it.reconciled=it.reconciled||{};if(chk.checked)it.reconciled[period]=true;else delete it.reconciled[period];save();renderItems();renderReclass();};td.appendChild(chk);tr.appendChild(td);tb.appendChild(tr)});reclassItems.forEach(it=>{const combo=`${it.fromSeg2||""}-${it.fromSeg3||""}-${it.fromSeg4||""}`;const tr=document.createElement('tr');tr.innerHTML=`<td>Reclass</td><td>${it.vendor} ${it.invoiceNumber}</td><td>${combo}</td>`;const td=document.createElement('td');const chk=document.createElement('input');chk.type='checkbox';chk.checked=!!(it.reconciled&&it.reconciled[period]);chk.onchange=()=>{it.reconciled=it.reconciled||{};if(chk.checked)it.reconciled[period]=true;else delete it.reconciled[period];save();renderItems();renderReclass();};td.appendChild(chk);tr.appendChild(td);tb.appendChild(tr)});t.appendChild(tb);host.innerHTML='';host.appendChild(t)}

function renderPeriods(){const tb=el('period-table')?.querySelector('tbody');if(!tb)return;tb.innerHTML='';periods.forEach((p,i)=>{const tr=document.createElement('tr');const td0=document.createElement('td');const in0=document.createElement('input');in0.type='month';in0.value=p.period||'';in0.oninput=()=>{p.period=in0.value;save();};td0.appendChild(in0);tr.appendChild(td0);const td1=document.createElement('td');const in1=document.createElement('input');in1.type='date';in1.value=p.begin||'';in1.oninput=()=>{p.begin=in1.value;save();};td1.appendChild(in1);tr.appendChild(td1);const td2=document.createElement('td');const in2=document.createElement('input');in2.type='date';in2.value=p.end||'';in2.oninput=()=>{p.end=in2.value;save();};td2.appendChild(in2);tr.appendChild(td2);const td3=document.createElement('td');const btn=document.createElement('button');btn.className='ghost';btn.textContent='✖';btn.onclick=()=>{periods.splice(i,1);save();renderPeriods();};td3.appendChild(btn);tr.appendChild(td3);tb.appendChild(tr);})}

function apModuleTotal(){return groups.reduce((a,g)=>a+(g.amount||0),0)}
function apNetTB(){if(!summaryTB.length)return 0;const rows=summaryTB.filter(r=>{const mod=(r["modsrc"]??r["Modsrc"]??r["MODSRC"]??"").toString().toUpperCase();const j=(r["jnlsrc"]??r["Jnlsrc"]??r["JNLSRC"]??"").toString().toUpperCase();return mod.includes("AP")&&j.includes("BCHCLS")});const dKey=Object.keys(summaryTB[0]||{}).find(k=>k.toLowerCase().includes("debit"));const cKey=Object.keys(summaryTB[0]||{}).find(k=>k.toLowerCase().includes("credit"));return round2(rows.reduce((a,r)=>a+(Number(dKey?r[dKey]:0)||0)-(Number(cKey?r[cKey]:0)||0),0))}
function renderTBCheck(){setText("ap-total",fmtUSD.format(apModuleTotal()));const net=apNetTB();setText("tb-total",fmtUSD.format(net));const diff=Math.abs(apModuleTotal()-Math.abs(net));setText("tb-diff",fmtUSD.format(diff));const ok=diff<0.01;const b=el("ap-check");b.textContent=ok?"Match":"Mismatch";b.className=ok?"badge-ok":"badge-bad";b.style.display="inline-block";const host=el("summarytb");host.innerHTML="";if(summaryTB.length){const all=Object.keys(summaryTB[0]);const cols=all.filter(k=>!k.toLowerCase().endsWith("ky")&&k.toLowerCase()!=="acctky");const t=document.createElement("table");t.innerHTML="<thead><tr>"+cols.map(k=>`<th>${k}</th>`).join("")+"</tr></thead>";const tb=document.createElement("tbody");summaryTB.forEach(r=>{const tr=document.createElement("tr");tr.innerHTML=cols.map(k=>`<td>${String(r[k]??"")}</td>`).join("");tb.appendChild(tr)});t.appendChild(tb);host.appendChild(t);host.style.display="block"}else host.style.display="none"}

function exportTXT(){
  const trndat=toISO(eom(periodEnd?new Date(periodEnd):new Date()));
  const ref=periodEnd?new Date(periodEnd):new Date();
  let seq=Number(seqStart||1);
  const out=[["Trndat","Jnlsrc","Jnlidn","Jnldsc","Modsrc","Co_num","SegnumT","SegnumT","SegnumF","SegnumF","DebAmt","CrdAmt","lngdsc"]];
  const baseCols={jnlsrc:"GL",modsrc:"GL",co_num:"01",segBlank:""};
  const fy=(fiscalYY||defaults.fyy||"00").slice(-2);
  const mm=(actualMM||defaults.amm||"00").slice(-2);
  const title=journalTitle||defaults.jnltitle||"";
  const jeid=()=>fy+mm+String(seq++).padStart(2,"0").slice(-2);

  reclassItems.forEach(it=>{
    const je=jeid();const base=[trndat,baseCols.jnlsrc,je,title,baseCols.modsrc,baseCols.co_num];
    const memo=it.memo||`Reclass ${it.vendor||""} ${it.invoiceNumber||""}`;
    out.push([...base,it.toSeg2||"",it.toSeg3||"",it.toSeg4||"",baseCols.segBlank,round2(it.amount),0,memo]);
    out.push([...base,it.fromSeg2||"",it.fromSeg3||"",it.fromSeg4||"",baseCols.segBlank,0,round2(it.amount),memo]);
  });

  const month=ref.getMonth(),year=ref.getFullYear();
  items.forEach(it=>{
    const lines=(it.schedule||[]).filter(r=>{const d=new Date(r.date);return d.getMonth()===month&&d.getFullYear()===year});
    if(!lines.length)return;const je=jeid();const base=[trndat,baseCols.jnlsrc,je,title,baseCols.modsrc,baseCols.co_num];
    lines.forEach(r=>{
      out.push([...base,it.amort.expSeg2||"",it.amort.expSeg3||"",it.amort.expSeg4||"",baseCols.segBlank,round2(r.amount),0,r.memo]);
      out.push([...base,it.asset.seg2||it.source.seg2||"",it.asset.seg3||it.source.seg3||"",it.asset.seg4||it.source.seg4||"",baseCols.segBlank,0,round2(r.amount),r.memo]);
    });
  });
  downloadTXT("AFS_Amort_"+trndat+".txt",out);
}

function showDetected(cols){const d=el("detected");d.textContent="Detected columns (SQL keys removed): "+cols.join(", ");d.style.display="block"}

function fillSelectOptions(sel, options, selected){sel.innerHTML="";const opt0=document.createElement("option");opt0.value="";opt0.textContent="(auto)";sel.appendChild(opt0);options.forEach(c=>{const o=document.createElement("option");o.value=c;o.textContent=c;o.selected=selected===c;sel.appendChild(o)})}

function sortAcctKeysNumerically(list){
  const parse=(k)=>k.split("-").map(x=>({raw:x,num:parseInt(String(x).replace(/\D/g,""))}));
  return list.sort((a,b)=>{
    const pa=parse(a.key), pb=parse(b.key);
    for(let i=0;i<3;i++){
      const na=isNaN(pa[i].num)?Number.MAX_SAFE_INTEGER:pa[i].num;
      const nb=isNaN(pb[i].num)?Number.MAX_SAFE_INTEGER:pb[i].num;
      if(na!==nb) return na-nb;
      if(pa[i].raw!==pb[i].raw) return pa[i].raw.localeCompare(pb[i].raw);
    }
    return a.key.localeCompare(b.key);
  });
}

function buildAcctIndex(){
  acctMap=new Map();acctList=[];
  if(!acctSheet||!wbSheets[acctSheet]?.length){setText("acct-stats","0 accounts loaded");return}
  const rows=wbSheets[acctSheet];
  rows.forEach(r=>{
    const a=String(r[acctCols.seg2]??"").trim(),b=String(r[acctCols.seg3]??"").trim(),c=String(r[acctCols.seg4]??"").trim();
    const desc=String(r[acctCols.desc]??"").trim();
    const activeCol=acctCols.active||"";
    const active=activeCol? String(r[activeCol]??"").toUpperCase()!=="N":true;
    if(a||b||c){acctMap.set(`${a}-${b}-${c}`,desc+(active?"":" (inactive)"));acctList.push({key:`${a}-${b}-${c}`,desc})}
  });
  acctList=sortAcctKeysNumerically(acctList);
  setText("acct-stats",`${acctMap.size} accounts loaded`);
  buildActivitySelect();
}

function loadWorkbook(file){
  const reader=new FileReader();
  reader.onload=(e)=>{
    const data=new Uint8Array(e.target.result);
    const wb=XLSX.read(data,{type:"array"});
    wbNames=wb.SheetNames.slice();
    wbSheets=Object.fromEntries(wb.SheetNames.map(n=>[n,XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:null})]));

    const apName=wb.SheetNames.find(n=>n.toLowerCase().includes("module")&&n.toLowerCase().includes("ap"))||wb.SheetNames.find(n=>n.toLowerCase().includes("module"))||wb.SheetNames[0];
    const ap=wbSheets[apName]||[];
    apCols=filterUsableCols(Object.keys(ap[0]||{}));
    apRows=ap;mapCols=guessMap(apCols);el("fname").textContent=file.name+" → "+apName;showDetected(apCols);renderMapUI();

    const sumName=wb.SheetNames.find(n=>n.toLowerCase().includes("summary")&&n.toLowerCase().includes("tb"));summaryTB=sumName?wbSheets[sumName]:[];
    const detName=wb.SheetNames.find(n=>n.toLowerCase().includes("detail")&&n.toLowerCase().includes("tb"));detailTB=detName?wbSheets[detName]:[];

    const preferred=wb.SheetNames.find(n=>/module\s*-\s*gl\s*\(accounts\)/i.test(n))||wb.SheetNames.find(n=>/account|gl/i.test(n));
    acctSheet=preferred||acctSheet||"";
    const sSel=el("acct-sheet");fillSelectOptions(sSel,wbNames,acctSheet);acctSheet=sSel.value||acctSheet;

    const cols=filterUsableCols(Object.keys((wbSheets[acctSheet]||[{}])[0]||{}));
    const pick=(name,fallback)=>cols.find(c=>c.toLowerCase()===name)||cols.find(c=>c.toLowerCase().includes(name.replace(/seg/,"segnum")))||fallback||"";
    acctCols.seg2=pick("segnumtwo",acctCols.seg2);
    acctCols.seg3=pick("segnumthr",acctCols.seg3);
    acctCols.seg4=pick("segnumfou",acctCols.seg4);
    acctCols.desc=(cols.includes("ovrdsc")?"ovrdsc":acctCols.desc||cols.find(c=>/desc/i.test(c))||"");
    acctCols.active=(cols.includes("accsts")?"accsts":acctCols.active||"");
    ["acct-seg2","acct-seg3","acct-seg4","acct-desc","acct-active"].forEach((id,i)=>fillSelectOptions(el(id),cols,acctCols[["seg2","seg3","seg4","desc","active"][i]]||""));

    buildAcctIndex();renderGroups();renderReclass();buildActivityDefaults();
  };
  reader.readAsArrayBuffer(file);
}


function buildActivitySelect(){
  const sel=el("act-select");sel.innerHTML="";
  const selected=parseCombos(el("act-search").value);
  acctList.slice(0,4000).forEach(a=>{
    const o=document.createElement("option");
    o.value=a.key;o.textContent=`${a.key} — ${a.desc}`;
    o.selected=selected.includes(a.key);
    sel.appendChild(o);
  });
}

function renderActivity(){
  const ref=el("act-period").value||toISO(new Date());
  const sel=el("act-select");
  let combos=parseCombos(el("act-search").value);
  if(!combos.length) combos=Array.from(sel.selectedOptions).map(o=>o.value);
  combos=[...new Set(combos)];
  Array.from(sel.options).forEach(o=>o.selected=combos.includes(o.value));
  el("act-search").value=combos.join(", ");
  const sum=el("act-summary"),host=el("act-table");sum.innerHTML="";host.innerHTML="";
  if(!combos.length){sum.innerHTML='<span class="small">Enter or select an account combo.</span>';return}
  const months=actMonths(ref);
  const calcMap=Object.fromEntries(months.map(m=>[monthKey(m),{dr:0,cr:0}]));
  const tbMap=detailTB.length?Object.fromEntries(months.map(m=>[monthKey(m),{dr:0,cr:0}])):null;
  combos.forEach(c=>{
    const calc=actCalcForCombo(c,ref);months.forEach(m=>{const k=monthKey(m);calcMap[k].dr+=calc.map[k].dr;calcMap[k].cr+=calc.map[k].cr});
    if(tbMap){const tb=actTBForCombo(c,ref);if(tb)months.forEach(m=>{const k=monthKey(m);tbMap[k].dr+=tb.map[k].dr;tbMap[k].cr+=tb.map[k].cr})}
  });
  const hdr=`<tr><th>Row</th>${months.map(m=>`<th class="num">${m.toLocaleString(undefined,{month:"short"})} Dr</th><th class="num">${m.toLocaleString(undefined,{month:"short"})} Cr</th>`).join("")}<th class="num">Total Dr</th><th class="num">Total Cr</th></tr>`;
  const rowFrom=(label,mp)=>`<tr><td>${label}</td>${months.map(m=>{const k=monthKey(m);const v=mp[k]||{dr:0,cr:0};return `<td class="num">${fmtUSD.format(v.dr)}</td><td class="num">${fmtUSD.format(v.cr)}</td>`}).join("")}<td class="num">${fmtUSD.format(months.reduce((a,m)=>a+(mp[monthKey(m)]?.dr||0),0))}</td><td class="num">${fmtUSD.format(months.reduce((a,m)=>a+(mp[monthKey(m)]?.cr||0),0))}</td></tr>`;
  const t=document.createElement("table");t.innerHTML=`<thead>${hdr}</thead><tbody>${rowFrom("Calculated",calcMap)}${tbMap?rowFrom("TB",tbMap):""}</tbody>`;host.appendChild(t);
  sum.innerHTML=combos.map(c=>{const desc=acctLookup(...c.split("-"));return `<div class=\"row\"><span class=\"pill\">${c}</span><span class=\"small\">${desc||""}</span></div>`}).join("");
}

function actMonths(ref){const arr=[];const base=ref?new Date(ref):new Date();for(let i=0;i<12;i++){const d=new Date(base.getFullYear(),base.getMonth()-11+i,1);arr.push(new Date(d.getFullYear(),d.getMonth(),1))}return arr}
function monthKey(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")}
function actCalcForCombo(combo,refDate){const combos=Array.isArray(combo)?combo:[combo];const set=new Set(combos);const months=actMonths(refDate);const map=Object.fromEntries(months.map(m=>[monthKey(m),{dr:0,cr:0}]));items.forEach(it=>{(it.schedule||[]).forEach(r=>{const d=new Date(r.date);const k=monthKey(new Date(d.getFullYear(),d.getMonth(),1));if(!map[k])return;if(set.has(r.debitCombo)) map[k].dr+=Number(r.amount)||0;if(set.has(r.creditCombo)) map[k].cr+=Number(r.amount)||0})});reclassItems.forEach(j=>{const k=monthKey(new Date(refDate||new Date()));if(set.has(j.toSeg2+"-"+j.toSeg3+"-"+j.toSeg4)) map[k].dr+=Number(j.amount)||0;if(set.has(j.fromSeg2+"-"+j.fromSeg3+"-"+j.fromSeg4)) map[k].cr+=Number(j.amount)||0});return {months,map}}
function actTBForCombo(combo,refDate){if(!detailTB.length)return null;const seg2k=Object.keys(detailTB[0]).find(k=>k.toLowerCase().includes("seg2")||k.toLowerCase().includes("segnumtwo"));const seg3k=Object.keys(detailTB[0]).find(k=>k.toLowerCase().includes("seg3")||k.toLowerCase().includes("segnumthr"));const seg4k=Object.keys(detailTB[0]).find(k=>k.toLowerCase().includes("seg4")||k.toLowerCase().includes("segnumfou"));const datek=Object.keys(detailTB[0]).find(k=>/date|trndat|trandat|posting/i.test(k));const debitk=Object.keys(detailTB[0]).find(k=>/debit|debitamt|debamt/i.test(k));const creditk=Object.keys(detailTB[0]).find(k=>/credit|crdamt|cramt/i.test(k));if(!seg2k||!seg3k||!seg4k||!datek||!(debitk||creditk))return null;const combos=Array.isArray(combo)?combo:[combo];const set=new Set(combos);const months=actMonths(refDate);const map=Object.fromEntries(months.map(m=>[monthKey(m),{dr:0,cr:0}]));detailTB.forEach(r=>{const k=[r[seg2k],r[seg3k],r[seg4k]].map(x=>String(x||"")).join("-");if(!set.has(k))return;const d=parseDate(r[datek]);if(!d)return;const mk=monthKey(new Date(d.getFullYear(),d.getMonth(),1));if(!map[mk])return;const dr=Number((debitk&&r[debitk])||0)||0;const cr=Number((creditk&&r[creditk])||0)||0;map[mk].dr+=dr;map[mk].cr+=cr});return {months,map}}


function buildActivityDefaults(){el("act-period").value=periodEnd||toISO(new Date());buildGroupFilter();buildActivitySelect();renderActivity()}


function handleImportSchedule(f){const r=new FileReader();r.onload=e=>{const rows=String(e.target.result).split(/\r?\n/).map(l=>l.split(/,|\t/));const hdr=rows.shift().map(x=>x.trim().toLowerCase());const idx=(n)=>hdr.indexOf(n);rows.forEach(c=>{if(!c.length)return;const o={date:c[idx("date")],amount:Number(c[idx("amount")]||0),debitCombo:`${c[idx("debitseg2")]||""}-${c[idx("debitseg3")]||""}-${c[idx("debitseg4")]||""}`,creditCombo:`${c[idx("creditseg2")]||""}-${c[idx("creditseg3")]||""}-${c[idx("creditseg4")]||""}`,memo:c[idx("memo")]||""};items.push({id:uid(),source:{type:"Import",vendor:"Imported",invoiceNumber:"",amount:o.amount,invoiceDate:parseDate(o.date)||new Date(),description:o.memo,seg2:o.creditCombo.split("-")[0],seg3:o.creditCombo.split("-")[1],seg4:o.creditCombo.split("-")[2],lines:[]},amort:{enabled:false,method:"straight",months:1,startDate:parseDate(o.date)||new Date(),postOn:"EOM",expSeg2:o.debitCombo.split("-")[0],expSeg3:o.debitCombo.split("-")[1],expSeg4:o.debitCombo.split("-")[2],memoTemplate:o.memo},asset:{seg2:o.creditCombo.split("-")[0],seg3:o.creditCombo.split("-")[1],seg4:o.creditCombo.split("-")[2]},schedule:[{date:parseDate(o.date)||new Date(),amount:o.amount,debitCombo:o.debitCombo,creditCombo:o.creditCombo,memo:o.memo}]})});save();renderItems();renderActivity()};r.readAsText(f)}
function handleImportJE(f){const r=new FileReader();r.onload=e=>{const rows=String(e.target.result).split(/\r?\n/).map(l=>l.split(/\t|,/));const hdr=rows.shift().map(x=>x.trim().toLowerCase());const g=(n)=>hdr.indexOf(n);rows.forEach(c=>{if(c.length<13)return;const deb=Number(c[g("debamt")]||0)||0;const crd=Number(c[g("crdamt")]||0)||0;const s2=c[g("segnumt")]||c[g("segnumt")];const s3=c[g("segnumt")]||c[g("segnumt")];const s4=c[g("segnumf")]||c[g("segnumf")];const memo=c[g("lngdsc")]||"";if(deb>0){reclassItems.push({id:uid(),vendor:"Imported JE",invoiceNumber:"",amount:deb,fromSeg2:"",fromSeg3:"",fromSeg4:"",toSeg2:s2,toSeg3:s3,toSeg4:s4,memo})}if(crd>0){reclassItems.push({id:uid(),vendor:"Imported JE",invoiceNumber:"",amount:crd,fromSeg2:s2,fromSeg3:s3,fromSeg4:s4,toSeg2:"",toSeg3:"",toSeg4:"",memo})}});save();renderReclass();renderActivity()};r.readAsText(f)}

/* Tabs & events */
el("dlg-close").onclick=()=>{el("dlg").style.display="none";detailGroup=null};
el("file").addEventListener("change",(e)=>{const f=e.target.files?.[0];if(f)loadWorkbook(f)});
el("chk-all").addEventListener("change",(e)=>{const c=e.currentTarget.checked;groups.forEach(g=>{actionsByKey[g.key]={...(actionsByKey[g.key]||{}),amortize:c}});renderGroups()});
el("add-to-amort").addEventListener("click",addSelectedToAmort);
el("add-to-reclass").addEventListener("click",addSelectedToReclass);
el("rebuild-all").addEventListener("click",()=>{items=items.map(it=>({...it,schedule:buildSchedule(it)}));save();renderItems()});
el("clear-items").addEventListener("click",()=>{items=[];save();renderItems()});
el("export").addEventListener("click",exportTXT);
["periodEnd","fiscalYY","actualMM","seqStart","jnlTitle"].forEach(id=>el(id).addEventListener("input",e=>{if(id==="periodEnd")periodEnd=e.target.value;if(id==="fiscalYY")fiscalYY=e.target.value.replace(/[^0-9]/g,"").slice(-2);if(id==="actualMM")actualMM=e.target.value.replace(/[^0-9]/g,"").slice(-2);if(id==="seqStart")seqStart=(e.target.value.replace(/[^0-9]/g,"").slice(-2)||"01");if(id==="jnlTitle")journalTitle=e.target.value;el(id).value=(id==="seqStart"?seqStart:(id==="fiscalYY"?fiscalYY:(id==="actualMM"?actualMM:e.target.value)));save()}));

el("acct-sheet").addEventListener("change",e=>{acctSheet=e.target.value;const cols=filterUsableCols(Object.keys((wbSheets[acctSheet]||[{}])[0]||{}));["acct-seg2","acct-seg3","acct-seg4","acct-desc","acct-active"].forEach((id,i)=>fillSelectOptions(el(id),cols,acctCols[["seg2","seg3","seg4","desc","active"][i]]||""))});
["acct-seg2","acct-seg3","acct-seg4","acct-desc","acct-active"].forEach((id,i)=>el(id).addEventListener("change",e=>{acctCols[["seg2","seg3","seg4","desc","active"][i]]=e.target.value;save()}));
el("acct-build").addEventListener("click",()=>{buildAcctIndex();save()});


el("tab-amort").onclick=()=>setMode("amort");
el("tab-activity").onclick=()=>{setMode("activity");buildActivityDefaults()};
el("tab-settings").onclick=()=>{setMode("settings");renderSettings()};
el("act-refresh").onclick=renderActivity;

el("act-select").addEventListener("change",()=>{
  const sel=el("act-select");
  const vals=Array.from(sel.selectedOptions).map(o=>o.value);
  el("act-search").value=vals.join(", ");
  renderActivity();
});

el("act-search").addEventListener("input",renderActivity);
el("act-period").addEventListener("input",renderActivity);


el("imp-sched").addEventListener("change",e=>{const f=e.target.files?.[0];if(f)handleImportSchedule(f)});
el("imp-je").addEventListener("change",e=>{const f=e.target.files?.[0];if(f)handleImportJE(f)});

// Sample template downloads
el("dl-sched").addEventListener("click",()=>downloadBlob(
  "schedule_template.csv",
  "date,amount,debitSeg2,debitSeg3,debitSeg4,creditSeg2,creditSeg3,creditSeg4,memo\n" +
  "2025-07-31,1250,61500,000,01,11415,000,02,Amortization July 2025\n"
));
el("dl-je").addEventListener("click",()=>downloadBlob(
  "je_template.txt",
  "Trndat\tJnlsrc\tJnlidn\tJnldsc\tModsrc\tCo_num\tSegnumT\tSegnumT\tSegnumF\tSegnumF\tDebAmt\tCrdAmt\tlngdsc\n" +
  "2025-07-31\tGL\t250701\tSample JE\tGL\t01\t15060\t000\t02\t\t1000\t0\tSample debit line\n" +
  "2025-07-31\tGL\t250701\tSample JE\tGL\t01\t11415\t000\t02\t\t0\t1000\tSample credit line\n"
));


/* Settings */
function renderGroupTable(){
  const tbody=el("grp-rows");
  tbody.innerHTML="";
  (defaults.groups||[]).forEach((g,i)=>{
    const tr=document.createElement("tr");
    tr.innerHTML=`<td><input data-idx="${i}" data-k="group" type="text" value="${g.group||""}"></td>`+
    `<td><input data-idx="${i}" data-k="seg2" type="text" value="${g.seg2||""}"></td>`+
    `<td><input data-idx="${i}" data-k="seg3" type="text" value="${g.seg3||""}"></td>`+
    `<td><input data-idx="${i}" data-k="seg4" type="text" value="${g.seg4||""}"></td>`+
    `<td><button data-idx="${i}" class="grp-del">✖</button></td>`;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll("input").forEach(inp=>{
    inp.addEventListener("input",e=>{
      const {idx,k}=e.target.dataset;defaults.groups[idx][k]=e.target.value.trim();
    });
  });
  tbody.querySelectorAll(".grp-del").forEach(btn=>{
    btn.addEventListener("click",e=>{const {idx}=e.target.dataset;defaults.groups.splice(idx,1);renderGroupTable()});
  });
}


function renderSettings(){
  el("set-fyy").value=defaults.fyy||fiscalYY||"";
  el("set-amm").value=defaults.amm||actualMM||"";
  el("set-amemo").value=defaults.amemo||"";
  el("set-jnltitle").value=defaults.jnltitle||journalTitle||"";

  renderGroupTable();
}
el("set-save").onclick=()=>{
  defaults.fyy=el("set-fyy").value.replace(/[^0-9]/g,"").slice(-2);
  defaults.amm=el("set-amm").value.replace(/[^0-9]/g,"").slice(-2);
  defaults.amemo=el("set-amemo").value||defaults.amemo;
  defaults.jnltitle=el("set-jnltitle").value||defaults.jnltitle;
  if(!fiscalYY)fiscalYY=defaults.fyy; if(!actualMM)actualMM=defaults.amm; if(!journalTitle)journalTitle=defaults.jnltitle;
  el("fiscalYY").value=fiscalYY; el("actualMM").value=actualMM; el("jnlTitle").value=journalTitle;
  save();
  buildGroupFilter();buildActivitySelect();
};
el("grp-add").onclick=()=>{defaults.groups.push({group:"",seg2:"",seg3:"",seg4:""});renderGroupTable()};
=======


el("edit-user").onclick=openProfile;el("force-login").onclick=openProfile;el("u-save").onclick=()=>{applyProfile();el("login").style.display="none"};

load();renderProfile();setMode("amort");
if(defaults.fyy&&!fiscalYY)fiscalYY=defaults.fyy;
if(defaults.amm&&!actualMM)actualMM=defaults.amm;
if(defaults.jnltitle&&!journalTitle)journalTitle=defaults.jnltitle;
if(!periodEnd&&closingPeriod)periodEnd=closingPeriodEnd();
el("recon-period").value=closingPeriod||"";
if(periodEnd)el("periodEnd").value=periodEnd;if(fiscalYY)el("fiscalYY").value=fiscalYY;if(actualMM)el("actualMM").value=actualMM;if(seqStart)el("seqStart").value=seqStart;if(journalTitle)el("jnlTitle").value=journalTitle;
renderItems();renderReclass();
