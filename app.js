/* Canvassing – web app statica.
   Liste Excel lette nel browser, geocodifica OpenStreetMap (o Google se c'è la chiave),
   suddivisione bilanciata tra le coppie e ordine di visita ottimizzato.
   Nessun dato clienti viene salvato su server: resta nel telefono o dentro il link condiviso. */
(function(){
"use strict";

/* ---------- utilità ---------- */
const $=s=>document.querySelector(s);
const esc=s=>String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const LS={
  get(k,d){try{const v=localStorage.getItem("cv:"+k);return v==null?d:JSON.parse(v)}catch(e){return d}},
  set(k,v){try{localStorage.setItem("cv:"+k,JSON.stringify(v));return true}catch(e){return false}},
  del(k){try{localStorage.removeItem("cv:"+k)}catch(e){}}
};
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
function toast(t,ms){const el=$("#toast");el.textContent=t;el.classList.remove("hidden");clearTimeout(toast._t);toast._t=setTimeout(()=>el.classList.add("hidden"),ms||2600)}
function km(a,b){const R=6371,r=Math.PI/180,dLa=(b.lat-a.lat)*r,dLo=(b.lon-a.lon)*r,x=Math.sin(dLa/2)**2+Math.cos(a.lat*r)*Math.cos(b.lat*r)*Math.sin(dLo/2)**2;return 2*R*Math.asin(Math.sqrt(x))}
const COLORS=["#c2410c","#1d4ed8","#15803d","#7e22ce","#be123c","#0e7490"];
const CODES="ABCDEF";
const ESITI=["","Interessato","Tessera / ordine fatto","Da richiamare","Titolare assente","Non interessato","Chiuso / cessato"];
const uid=()=>Date.now().toString(36)+Math.random().toString(36).slice(2,6);

/* ---------- stato ---------- */
let settings=Object.assign({city:"Rovereto",k:3,meet:"Parcheggio Mart, Via Sticcotta",ret:true,gkey:""},LS.get("settings",{}));
let files=LS.get("files",[]);          // [{id,name,label,rows:[client]}]
let plan=LS.get("plan",null);          // piano completo (responsabile)
let geoCache=LS.get("geo",{});         // query -> {lat,lon,prec}
let shared=null;                       // piano ricevuto via link (venditori)
let curTeam=0, selStop=null, moveMode=null;

/* ---------- normalizzazione indirizzi ---------- */
const SMALL=new Set(["di","de","del","della","delle","dei","degli","dello","d","al","alla","alle","ai","agli","e","a","in","da","dal","dalla","sul","sulla","per","con","san","la","le","lo","il","i","gli"]);
function titleCase(s){
  return s.toLowerCase().replace(/[a-zà-ü]+('|’)?/gi,(w,ap,i)=>{
    const base=w.replace(/['’]$/,"");
    if(i>0&&SMALL.has(base)&&base!=="san")return w;
    return w.charAt(0).toUpperCase()+w.slice(1);
  });
}
const ABBR=[[/\bC\.?\s*SO\b\.?/i,"Corso"],[/\bV\.?\s*LE\b\.?/i,"Viale"],[/\bP\.?\s*Z+A\b\.?/i,"Piazza"],[/\bP\.?\s*LE\b\.?/i,"Piazzale"],[/\bL\.?\s*GO\b\.?/i,"Largo"],[/\bLOC\.?(?=\s)/i,"Località"],[/\bFRAZ\.?(?=\s)/i,"Frazione"],[/\bV\.\s*(?=[A-Z])/i,"Via "],[/\bCS\.?\s*(?=[A-Z])/i,"Corso "]];
function cleanAddr(a){
  let s=String(a||"").replace(/\s+/g," ").trim();
  if(!s)return "";
  s=s.replace(/\bC\/O\b.*$/i,"").replace(/\bS\.?N\.?C\.?\s*$/i,"").replace(/\bSNC\b/i,"").replace(/,?\s*\d{5}\s.*$/,"");
  ABBR.forEach(([r,t])=>{s=s.replace(r,t)});
  // civico: "5/7"->"5", "6/B"->"6B", "82 D"->"82D"
  s=s.replace(/(\d+)\s*\/\s*([A-Z])\b/i,"$1$2").replace(/(\d+)\s*\/\s*\d+/,"$1").replace(/(\d+)\s+([A-Z])$/i,"$1$2").replace(/\s*-\s*$/,"");
  s=titleCase(s).replace(/(\d+)([a-z])\b/g,(m,n,l)=>n+l.toUpperCase()).replace(/\b([a-z])\./g,(m,l)=>l.toUpperCase()+".");
  return s.trim();
}
function cityFrom(v){
  const s=String(v||"").trim(); if(!s)return {city:"",cap:""};
  const m=s.match(/(\d{5})?\s*[-–,]?\s*(.*)$/);
  return {cap:(m&&m[1])||"",city:titleCase(((m&&m[2])||"").trim())};
}

/* ---------- lettura Excel ---------- */
const COLS={
  rs:/ragione|denominaz|^nome|cliente$|nominativo|azienda/i,
  ins:/insegna|nome commerciale/i,
  piva:/iva|vat|cod.*fisc/i,
  cat:/ateco|^ctg$|categor|settore|tipolog|attivit/i,
  addr:/indirizzo|^via$|address|strada/i,
  civ:/civico|^n\.?\s*civ|^num/i,
  city:/comune|citt|cap.?com|localit/i,
  cap:/^cap$/i,
  cid:/^cu ?id$|codice cliente|^id$|cod\.? ?cli/i,
  trend:/segment|trend|stato|note/i
};
function loadXLSX(){
  if(window.XLSX)return Promise.resolve();
  return new Promise((ok,ko)=>{const s=document.createElement("script");s.src="vendor/xlsx.core.min.js";s.onload=ok;s.onerror=()=>ko(new Error("Impossibile caricare il lettore Excel"));document.head.appendChild(s)});
}
function guessLabel(name){
  const n=name.toUpperCase();
  if(/NBC/.test(n))return "NBC";
  if(/PROSP/.test(n))return "PRO";
  return n.replace(/\.[A-Z]+$/,"").replace(/[^A-Z0-9]/g,"").slice(0,6)||"LISTA";
}
function parseSheet(wb){
  for(const sn of wb.SheetNames){
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:"",raw:false});
    for(let h=0;h<Math.min(rows.length,15);h++){
      const hdr=rows[h].map(x=>String(x).trim());
      const map={};
      hdr.forEach((t,i)=>{if(!t)return;for(const k in COLS){if(map[k]==null&&COLS[k].test(t)){if(k==="cap"&&/com/i.test(t))continue;map[k]=i;break}}});
      if(map.addr!=null&&(map.rs!=null||map.ins!=null)){
        const out=[];
        for(const r of rows.slice(h+1)){
          const g=k=>map[k]==null?"":String(r[map[k]]==null?"":r[map[k]]).trim();
          let addr=g("addr"); if(!addr)continue;
          if(g("civ"))addr+=" "+g("civ");
          const cc=cityFrom(g("city")); if(!cc.cap&&g("cap"))cc.cap=g("cap");
          let rs=g("rs").replace(/^\d+_GESTOR_/i,"");
          out.push({rs:rs||g("ins"),ins:(g("ins")&&g("ins")!==rs)?g("ins"):"",piva:g("piva").replace(/^IT/i,""),cat:g("cat").replace(/^\d+\s*-\s*/,""),addr,city:cc.city,cap:cc.cap,trend:g("trend"),cid:g("cid")});
        }
        return {sheet:sn,header:hdr,map,rows:out};
      }
    }
  }
  return null;
}
async function onFiles(list){
  try{await loadXLSX()}catch(e){toast(e.message);return}
  for(const f of list){
    try{
      const buf=await f.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array"});
      const res=parseSheet(wb);
      if(!res){files.push({id:uid(),name:f.name,label:guessLabel(f.name),rows:[],err:"Non trovo le colonne Ragione Sociale/Insegna e Indirizzo"});continue}
      files.push({id:uid(),name:f.name,label:guessLabel(f.name),rows:res.rows,cols:Object.keys(res.map).map(k=>res.header[res.map[k]])});
    }catch(e){files.push({id:uid(),name:f.name,label:"?",rows:[],err:"File non leggibile: "+e.message})}
  }
  saveFiles();renderFiles();
}
function saveFiles(){if(!LS.set("files",files))toast("Memoria del telefono piena: le liste non sono state salvate")}
function renderFiles(){
  const box=$("#files");box.innerHTML="";
  files.forEach(f=>{
    const d=document.createElement("div");d.className="filerow";
    d.innerHTML=`<div class="nm"><b>${esc(f.name)}</b><span class="small ${f.err?"prec-zona":"muted"}">${f.err?esc(f.err):f.rows.length+" clienti · colonne: "+esc((f.cols||[]).join(", "))}</span></div>
      <input value="${esc(f.label)}" maxlength="6" aria-label="Etichetta lista"><button aria-label="Rimuovi">×</button>`;
    d.querySelector("input").onchange=e=>{f.label=e.target.value.toUpperCase().trim()||"LISTA";saveFiles()};
    d.querySelector("button").onclick=()=>{files=files.filter(x=>x!==f);saveFiles();renderFiles()};
    box.appendChild(d);
  });
  const n=files.reduce((s,f)=>s+f.rows.length,0);
  $("#b-calc").disabled=!n;
  $("#b-calc").textContent=n?`Calcola percorsi (${n} clienti)`:"Calcola percorsi";
}

/* ---------- geocodifica ---------- */
let gReady=null;
function loadGoogle(){
  if(!settings.gkey)return Promise.resolve(false);
  if(gReady)return gReady;
  gReady=new Promise(ok=>{
    window.__gmok=()=>ok(true);
    window.gm_authFailure=()=>{toast("Chiave Google non valida: uso OpenStreetMap",4000);settings.gkey="";ok(false)};
    const s=document.createElement("script");
    s.src="https://maps.googleapis.com/maps/api/js?key="+encodeURIComponent(settings.gkey)+"&callback=__gmok&loading=async&language=it&region=IT";
    s.onerror=()=>ok(false);document.head.appendChild(s);
    setTimeout(()=>ok(!!(window.google&&google.maps)),12000);
  });
  return gReady;
}
let nomQ=Promise.resolve();
function nomSlot(){const p=nomQ.then(()=>sleep(1100));const ready=nomQ;nomQ=p;return ready}
async function nominatim(params){
  await nomSlot();
  const u="https://nominatim.openstreetmap.org/search?format=jsonv2&addressdetails=1&limit=3&countrycodes=it&accept-language=it&"+new URLSearchParams(params);
  const r=await fetch(u);if(!r.ok)throw new Error("nominatim "+r.status);return r.json();
}
async function photon(q,bias){
  let u="https://photon.komoot.io/api/?limit=6&q="+encodeURIComponent(q);
  if(bias)u+="&lat="+bias.lat+"&lon="+bias.lon;
  const r=await fetch(u);if(!r.ok)throw new Error("photon "+r.status);return r.json();
}
function streetKey(s){return s.toLowerCase().replace(/\s+\d+\s*[a-z]?\s*$/,"").replace(/\b(via|viale|corso|piazza|piazzale|largo|vicolo|strada|località|scala|salita|contrada)\b/g,"").replace(/[^a-zà-ü]/g,"")}
/* nome commerciale "pulito" per la ricerca: toglie forma societaria e titolare */
function cleanName(n){
  let s=String(n||"").replace(/^\d+_GESTOR_/i,"").replace(/["“”]/g," ");
  s=s.split(/\s(?:DI|DEI|DELLA|DEL)\s+[A-Z][A-Z']+\s+[A-Z][A-Z']+.*$/)[0];
  s=s.replace(/\b(S\.?\s?R\.?\s?L\.?S?|S\.?\s?N\.?\s?C\.?|S\.?\s?A\.?\s?S\.?|S\.?\s?P\.?\s?A\.?|SOCIETA'?\s.*|SOC\.?\s.*|SOCIET[AÀ]\s.*|IMPRESA SOCIALE|COOPERATIVA.*|& C\.?|E C\.)\s*$/i,"");
  s=s.replace(/\b(S\.?\s?R\.?\s?L\.?S?|S\.?\s?N\.?\s?C\.?|S\.?\s?A\.?\s?S\.?)\b\.?/ig,"").replace(/[\s,.-]+$/,"").replace(/\s+/g," ").trim();
  return s.length>=3?titleCase(s):"";
}
const nameToks=x=>String(x||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").split(/[^a-z0-9]+/).filter(w=>w.length>2&&!/^(bar|the|del|della|dei|di|ristorante|pizzeria|hotel|albergo|caffe|cafe|caffetteria|osteria|trattoria|gelateria|pasticceria|panificio|birreria|enoteca|pub|market|alimentari|macelleria|bistrot|locanda|srl|snc|sas|spa|via|and|con)$/.test(w));
function nameMatch(want,got,city){const cs=nameToks(city);const a=nameToks(want).filter(w=>!cs.includes(w)),b=nameToks(got);if(!a.length||!b.length)return false;const hit=a.filter(w=>b.includes(w)).length;return hit>=Math.min(2,a.length)||(hit>=1&&a.length<=2)}
const sameCity=(p,city,center)=>{const c=(p.city||p.town||p.village||p.municipality||"").toLowerCase();return (c&&c===city.toLowerCase())||(!c&&center&&km(p._pt,center)<8)};
/* cerca l'attività per nome (OpenStreetMap / Google) nella stessa città */
async function findByName(names,city,center){
  for(const nm of names){
    if(!nm)continue;
    if(await loadGoogle()){
      try{const {results}=await new google.maps.Geocoder().geocode({address:nm+", "+city+", Italia",region:"it"});
        const r=(results||[]).find(r=>(r.types||[]).some(t=>t==="establishment"||t==="point_of_interest"||t==="food"));
        if(r){const l=r.geometry.location,pt={lat:l.lat(),lon:l.lng()};if(!center||km(pt,center)<12)return {...pt,name:nm,where:r.formatted_address.split(",").slice(0,2).join(",")}}}catch(e){}
    }
    try{
      const j=await photon(nm+" "+city,center);
      for(const f of j.features||[]){
        const p=f.properties;p._pt={lat:f.geometry.coordinates[1],lon:f.geometry.coordinates[0]};
        if(!p.name||!nameMatch(nm,p.name,city))continue;
        if(!sameCity(p,city,center))continue;
        return {...p._pt,name:p.name,where:[p.street,p.housenumber].filter(Boolean).join(" ")||p.district||city,street:p.street||""};
      }
    }catch(e){}
    try{
      const j=await nominatim({q:nm+", "+city});
      for(const x of j){const pt={lat:+x.lat,lon:+x.lon};const a=x.address||{};
        const nmx=x.name||(x.display_name||"").split(",")[0];
        if(!nameMatch(nm,nmx,city))continue;
        const c=(a.city||a.town||a.village||a.municipality||"").toLowerCase();
        if(c!==city.toLowerCase()&&!(center&&km(pt,center)<8))continue;
        return {...pt,name:nmx,where:[a.road,a.house_number].filter(Boolean).join(" ")||city,street:a.road||""};}
    }catch(e){}
  }
  return null;
}
async function geocodeOne(addr,city,cap,center,names){
  const q=addr+", "+(cap?cap+" ":"")+city;
  const key=q.toLowerCase()+"|"+(names||[]).join("|").toLowerCase();
  if(geoCache[key])return geoCache[key];
  let res=null;
  const g=addr?await loadGoogle():false;
  if(g){
    try{
      const geo=new google.maps.Geocoder();
      const {results}=await geo.geocode({address:q+", Italia",region:"it"});
      const r=results&&results[0];
      if(r){const l=r.geometry.location;const t=r.types||[];
        const p={lat:l.lat(),lon:l.lng()};
        if(!center||km(p,center)<25)res={...p,prec:t.includes("street_address")||t.includes("premise")||t.includes("establishment")?"civico":t.includes("route")?"via":"zona"};}
    }catch(e){}
  }
  if(!res&&addr){
    const sk=streetKey(addr), hn=((addr.match(/\s(\d+[A-Z]?)\s*$/i)||["",""])[1]).toLowerCase();
    try{
      const j=await photon(q,center);
      const feats=(j.features||[]).map(f=>({p:f.properties,lat:f.geometry.coordinates[1],lon:f.geometry.coordinates[0]})).filter(f=>!center||km(f,center)<25);
      const toks=x=>x.toLowerCase().replace(/\s+\d+\s*[a-z]?\s*$/,"").split(/[^a-zà-ü]+/).filter(w=>w.length>3&&!/^(via|viale|corso|piazza|piazzale|largo|vicolo|strada|località|scala|salita|contrada|della|delle|degli)$/.test(w));
      const qt=toks(addr),last=qt[qt.length-1];
      const same=f=>{const n=f.p.street||f.p.name||"";if(sk&&streetKey(n)===sk)return true;const rt=toks(n);return !!last&&rt.includes(last)&&(qt.length<2||qt.filter(w=>rt.includes(w)).length>=Math.min(2,rt.length))};
      let f=feats.find(f=>same(f)&&f.p.housenumber&&String(f.p.housenumber).toLowerCase()===hn)||feats.find(f=>same(f)&&f.p.housenumber);
      if(f)res={lat:f.lat,lon:f.lon,prec:String(f.p.housenumber).toLowerCase()===hn?"civico":"via"};
      else if((f=feats.find(same)))res={lat:f.lat,lon:f.lon,prec:"via"};
    }catch(e){}
  }
  if(!res&&addr){
    try{
      const j=await nominatim({street:addr,city,postalcode:cap||""});
      const f=j.find(x=>!center||km({lat:+x.lat,lon:+x.lon},center)<25);
      if(f)res={lat:+f.lat,lon:+f.lon,prec:f.address&&f.address.house_number?"civico":"via"};
    }catch(e){}
  }
  if(!res){ // ultima prova: solo la via
    const street=addr.replace(/\s+\d+\s*[A-Z]?\s*$/i,"");
    if(street&&street!==addr){
      try{const j=await nominatim({street,city});const f=j.find(x=>!center||km({lat:+x.lat,lon:+x.lon},center)<25);if(f)res={lat:+f.lat,lon:+f.lon,prec:"via"}}catch(e){}
    }
  }
  // indirizzo non trovato o solo la via: cerco l'attività per nome nella stessa città
  if((!res||res.prec!=="civico")&&names&&names.length){
    const hit=await findByName(names,city,center);
    if(hit){
      const sk=streetKey(addr),hs=streetKey(hit.street||hit.where||"");
      const sameStreet=!!sk&&!!hs&&(sk===hs||sk.includes(hs)||hs.includes(sk));
      if(sameStreet)res={lat:hit.lat,lon:hit.lon,prec:"civico",web:"Posizione trovata cercando \u00ab"+hit.name+"\u00bb: stessa via della lista"};
      else if(!res)res={lat:hit.lat,lon:hit.lon,prec:"nome",web:"Indirizzo della lista non trovato. Trovata l'attivit\u00e0 \u00ab"+hit.name+"\u00bb in "+hit.where+", "+city+": verifica sul posto"};
    }
  }
  if(res){geoCache[key]=res;LS.set("geo",geoCache)}
  return res;
}
async function geocodeCity(city){
  const key="#city:"+city.toLowerCase();
  if(geoCache[key])return geoCache[key];
  let res=null;
  if(await loadGoogle()){try{const {results}=await new google.maps.Geocoder().geocode({address:city+", Italia"});const l=results[0].geometry.location;res={lat:l.lat(),lon:l.lng()}}catch(e){}}
  if(!res){try{const j=await nominatim({city});if(j[0])res={lat:+j[0].lat,lon:+j[0].lon}}catch(e){}}
  if(!res){try{const j=await photon(city+", Italia");const f=j.features&&j.features[0];if(f)res={lat:f.geometry.coordinates[1],lon:f.geometry.coordinates[0]}}catch(e){}}
  if(res){geoCache[key]=res;LS.set("geo",geoCache)}
  return res;
}
async function geocodePlace(text,city,center){
  const key="#place:"+(text+"|"+city).toLowerCase();
  if(geoCache[key])return geoCache[key];
  let res=null;
  if(await loadGoogle()){try{const {results}=await new google.maps.Geocoder().geocode({address:text+", "+city+", Italia"});const l=results[0].geometry.location;res={lat:l.lat(),lon:l.lng()}}catch(e){}}
  if(!res){try{const j=await photon(text+", "+city,center);const f=(j.features||[]).find(f=>!center||km({lat:f.geometry.coordinates[1],lon:f.geometry.coordinates[0]},center)<25);if(f)res={lat:f.geometry.coordinates[1],lon:f.geometry.coordinates[0]}}catch(e){}}
  if(!res){try{const j=await nominatim({q:text+", "+city});const f=j.find(x=>!center||km({lat:+x.lat,lon:+x.lon},center)<25);if(f)res={lat:+f.lat,lon:+f.lon}}catch(e){}}
  if(res){geoCache[key]=res;LS.set("geo",geoCache)}
  return res;
}

/* ---------- ottimizzazione ---------- */
function routeLen(start,pts,ret){let d=0,p=start;for(const q of pts){d+=km(p,q);p=q}if(ret)d+=km(p,start);return d}
function orderRoute(start,pts,ret){
  if(pts.length<2)return pts.slice();
  // nearest neighbour
  let rest=pts.slice(),out=[],p=start;
  while(rest.length){let bi=0,bd=1e9;rest.forEach((q,i)=>{const d=km(p,q);if(d<bd){bd=d;bi=i}});p=rest.splice(bi,1)[0];out.push(p)}
  // 2-opt + or-opt
  const all=[start,...out];
  const n=all.length;
  const D=(i,j)=>km(all[i],all[j]);
  let improved=true,guard=0;
  while(improved&&guard++<200){
    improved=false;
    for(let i=1;i<n-1;i++){
      for(let j=i+1;j<n;j++){
        const a=i-1,b=i,c=j,d=j+1<n?j+1:(ret?0:-1);
        const before=D(a,b)+(d>=0?D(c,d):0), after=D(a,c)+(d>=0?D(b,d):0);
        if(after<before-1e-9){const seg=all.slice(i,j+1).reverse();all.splice(i,seg.length,...seg);improved=true}
      }
    }
    // or-opt: sposta una tappa in un'altra posizione
    for(let i=1;i<n;i++){
      const cur=routeLen(all[0],all.slice(1),ret);
      const node=all.splice(i,1)[0];let best=-1,bl=cur-1e-9;
      for(let j=1;j<=all.length;j++){all.splice(j,0,node);const l=routeLen(all[0],all.slice(1),ret);if(l<bl){bl=l;best=j}all.splice(j,1)}
      all.splice(best>0?best:i,0,node);if(best>0&&best!==i)improved=true;
    }
  }
  return all.slice(1);
}
function balancedClusters(pts,k,start,ret){
  // pts: [{lat,lon,w}] ; zone bilanciate per numero di clienti (w), compatte, poi ordinate
  if(k<=1)return [pts.map((_,i)=>i)];
  const total=pts.reduce((s,p)=>s+p.w,0),cap=Math.ceil(total/k)+1;
  let best=null,bestCost=1e18;
  const tries=Math.min(24,4+pts.length);
  for(let t=0;t<tries;t++){
    // k-means++ seed (deterministico per tentativo)
    let seed=t*9301+49297;const rnd=()=>((seed=(seed*9301+49297)%233280)/233280);
    const cs=[pts[Math.floor(rnd()*pts.length)]];
    while(cs.length<k){const ds=pts.map(p=>Math.min(...cs.map(c=>km(p,c)))**2);const s=ds.reduce((a,b)=>a+b,0);let r=rnd()*s,i=0;while(i<ds.length-1&&(r-=ds[i])>0)i++;cs.push(pts[i])}
    let C=cs.map(c=>({lat:c.lat,lon:c.lon})),asg=[];
    for(let it=0;it<25;it++){
      const pairs=[];pts.forEach((p,i)=>C.forEach((c,j)=>pairs.push([km(p,c),i,j])));pairs.sort((a,b)=>a[0]-b[0]);
      const load=Array(k).fill(0),na=Array(pts.length).fill(-1);
      for(const [,i,j] of pairs){if(na[i]>=0)continue;if(load[j]+pts[i].w>cap&&load.some((l,jj)=>l+pts[i].w<=cap))continue;na[i]=j;load[j]+=pts[i].w}
      const same=asg.length&&na.every((v,i)=>v===asg[i]);asg=na;
      C=C.map((c,j)=>{const m=pts.filter((_,i)=>asg[i]===j);if(!m.length)return c;return {lat:m.reduce((s,p)=>s+p.lat,0)/m.length,lon:m.reduce((s,p)=>s+p.lon,0)/m.length}});
      if(same)break;
    }
    const groups=Array.from({length:k},(_,j)=>pts.map((_,i)=>i).filter(i=>asg[i]===j));
    const cost=groups.reduce((s,g)=>s+routeLen(start,orderRoute(start,g.map(i=>pts[i]),ret),ret),0);
    // penalizza squilibrio
    const loads=groups.map(g=>g.reduce((s,i)=>s+pts[i].w,0)),imb=Math.max(...loads)-Math.min(...loads);
    const c2=cost*(1+imb*0.03);
    if(c2<bestCost){bestCost=c2;best=groups}
  }
  return best;
}
function buildStops(clients){
  // raggruppa clienti allo stesso indirizzo (entro 25 m)
  const stops=[];
  clients.forEach(c=>{
    let s=stops.find(s=>km(s,c)<0.025);
    if(!s){s={lat:c.lat,lon:c.lon,prec:c.prec,addr:c.addrClean||c.addr,acc:[]};if(c.web)s.web=c.web;stops.push(s)}
    s.acc.push(c.id);
    if(c.prec==="civico")s.prec="civico";
  });
  return stops;
}
function makeTeams(stops,k,start,ret){
  const pts=stops.map(s=>({lat:s.lat,lon:s.lon,w:s.acc.length,s}));
  const groups=balancedClusters(pts,k,start,ret);
  // ordina le zone in senso orario intorno al ritrovo per lettere stabili
  const ang=g=>{const la=g.reduce((s,i)=>s+pts[i].lat,0)/g.length,lo=g.reduce((s,i)=>s+pts[i].lon,0)/g.length;return Math.atan2(la-start.lat,lo-start.lon)};
  groups.sort((a,b)=>ang(b)-ang(a));
  return groups.map((g,j)=>{
    const ordered=orderRoute(start,g.map(i=>pts[i]),ret).map(p=>p.s);
    return {code:CODES[j],color:COLORS[j%COLORS.length],stops:ordered};
  });
}
function reorderTeam(T,start,ret){T.stops=orderRoute(start,T.stops,ret);T.geom=null}

/* ---------- calcolo ---------- */
async function calculate(){
  const btn=$("#b-calc"),pr=$("#progress"),msg=$("#calc-msg");
  settings.city=$("#s-city").value.trim()||"Rovereto";settings.k=+$("#s-k").value;settings.meet=$("#s-meet").value.trim();settings.ret=$("#s-ret").checked;LS.set("settings",settings);
  btn.disabled=true;pr.classList.remove("hidden");msg.innerHTML="";
  const setP=(i,n,t)=>{pr.querySelector("i").style.width=(n?Math.round(i/n*100):0)+"%";pr.querySelector("span").textContent=t};
  try{
    setP(0,1,"Cerco "+settings.city+"…");
    const center=await geocodeCity(settings.city);
    if(!center)throw new Error("Non riesco a trovare il comune \""+settings.city+"\". Controlla il nome o la connessione.");
    let meet=null;
    if(settings.meetGps&&settings.meet==="La mia posizione")meet=settings.meetGps;
    else if(settings.meet){setP(0,1,"Cerco il punto di ritrovo…");const p=await geocodePlace(settings.meet,settings.city,center);if(p)meet={...p,name:settings.meet}}
    if(!meet){meet={...center,name:"Centro di "+settings.city};msg.innerHTML+=`<div class="msg warn">Punto di ritrovo non trovato: uso il centro di ${esc(settings.city)}. Puoi spostarlo dalla mappa.</div>`}
    // clienti
    const clients=[];const seen=new Set();const dups=[];const usedIds=new Set();
    files.forEach(f=>f.rows.forEach((r,i)=>{
      const ac=cleanAddr(r.addr);const city=r.city||settings.city;
      const dk=(r.rs+"|"+ac).toLowerCase();if(seen.has(dk)){dups.push(r.rs+" – "+r.addr+" ("+f.label+")");return}seen.add(dk);
      let id=f.label+"-"+String(r.cid||i+1).replace(/[^\w.]/g,"");const base=id;let n=2;while(usedIds.has(id))id=base+"_"+(n++);usedIds.add(id);
      clients.push({id,src:f.label,rs:r.rs,ins:r.ins,piva:r.piva,cat:r.cat,addr:r.addr,addrClean:ac,city,cap:r.cap,trend:r.trend});
    }));
    const miss=[];
    let next=0,doneN=0;const t0=Date.now();
    const worker=async()=>{while(next<clients.length){const c=clients[next++];
      const names=[...new Set([c.ins,cleanName(c.rs)].map(x=>cleanName(x)).filter(Boolean))];
      let g=null;try{g=c.addrClean?await geocodeOne(c.addrClean,c.city,c.cap,center,names):(names.length?await geocodeOne("",c.city,c.cap,center,names):null)}catch(e){}
      if(g){c.lat=g.lat;c.lon=g.lon;c.prec=g.prec;if(g.web)c.web=g.web}else miss.push(c);
      doneN++;const left=doneN>4?Math.round((Date.now()-t0)/doneN*(clients.length-doneN)/1000):0;
      setP(doneN,clients.length,`Trovo gli indirizzi: ${doneN} di ${clients.length}`+(left>20?` · circa ${left<90?left+" s":Math.round(left/60)+" min"}`:""))}};
    setP(0,clients.length,`Trovo gli indirizzi: 0 di ${clients.length}`);
    await Promise.all([worker(),worker(),worker()]);
    miss.sort((a,b)=>clients.indexOf(a)-clients.indexOf(b));
    setP(1,1,"Divido le zone e calcolo i percorsi…");await sleep(30);
    const ok=clients.filter(c=>c.lat!=null);
    if(!ok.length)throw new Error("Non sono riuscito a trovare nessun indirizzo. Controlla la connessione e il comune.");
    const stops=buildStops(ok);
    const k=Math.min(settings.k,stops.length);
    const teams=makeTeams(stops,k,meet,settings.ret);
    const byId={};clients.forEach(c=>{byId[c.id]=c});
    plan={sync:settings.sync||"",id:uid(),date:new Date().toISOString().slice(0,10),city:settings.city,ret:settings.ret,meet,clients:byId,teams,missing:miss.map(c=>c.id)};
    LS.set("plan",plan)||toast("Memoria piena: il piano non è stato salvato");
    pr.classList.add("hidden");
    msg.innerHTML+=`<div class="msg ok">Fatto: ${ok.length} clienti in ${stops.length} tappe, divisi in ${k} coppie.</div>`+
      (dups.length?`<div class="msg warn"><b>${dups.length} righe doppie</b> (stesso nome e indirizzo), contate una volta sola:<br>${dups.map(esc).join("<br>")}</div>`:"")+
      (miss.length?`<div class="msg warn"><b>${miss.length} indirizzi non trovati</b> (li trovi in fondo alla mappa, puoi posizionarli a mano):<br>${miss.map(c=>esc(c.rs+" – "+c.addr)).join("<br>")}</div>`:"");
    curTeam=0;selStop=null;
    fetchGeoms();startSync();
    setTimeout(()=>go("mappa"),miss.length?1800:600);
  }catch(e){pr.classList.add("hidden");msg.innerHTML=`<div class="msg err">${esc(e.message)}</div>`}
  btn.disabled=false;
}

/* ---------- percorso stradale (OSRM) ---------- */
async function fetchGeom(T,start,ret){
  const pts=[start,...T.stops];if(ret)pts.push(start);
  try{
    const u="https://router.project-osrm.org/route/v1/driving/"+pts.map(p=>p.lon.toFixed(6)+","+p.lat.toFixed(6)).join(";")+"?overview=false";
    const r=await fetch(u);const j=await r.json();
    if(j.code==="Ok"){T.geom=1;T.km=+(j.routes[0].distance/1000).toFixed(1);return true}
  }catch(e){}
  T.geom=null;T.km=+(routeLen(start,T.stops,ret)*1.3).toFixed(1);return false;
}
async function fetchGeoms(){
  const P=current();if(!P)return;
  for(const T of P.teams){if(!T.geom)await fetchGeom(T,P.meet,P.ret)}
  if(P===plan)LS.set("plan",plan);
  drawMap();renderPanel();
}

/* ---------- condivisione ---------- */
function teamPayload(P,ti){
  const T=P.teams[ti];
  return {v:1,s:P.sync||"",id:P.id,d:P.date,c:P.city,r:P.ret?1:0,m:[+P.meet.lat.toFixed(6),+P.meet.lon.toFixed(6),P.meet.name],
    t:[T.code,T.color,T.stops.map(s=>[+s.lat.toFixed(5),+s.lon.toFixed(5),s.prec==="civico"?1:s.prec==="nome"?2:0,s.acc.map(id=>{const c=P.clients[id];return [c.id.replace(/^[A-Z]+-/,""),c.src,c.rs,c.ins&&c.ins!==c.rs?c.ins:"",c.piva||"",(c.cat||"").slice(0,24),c.addr,""]}),s.prec==="nome"?(s.web||"").replace("Indirizzo della lista non trovato. Trovata l'attività","Trovata").replace(": verifica sul posto",""):""])]};
}
/* codifica solo con lettere, cifre, "-" e "_": WhatsApp e altre app non spezzano il link */
function b64u(bytes){let s="";for(let i=0;i<bytes.length;i++)s+=String.fromCharCode(bytes[i]);return btoa(s).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,"")}
function unb64u(t){t=t.replace(/-/g,"+").replace(/_/g,"/");while(t.length%4)t+="=";const s=atob(t),u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u}
function shareUrl(P,ti){return location.origin+location.pathname+"#z="+b64u(LZString.compressToUint8Array(JSON.stringify(teamPayload(P,ti))))}
function readHash(){
  const m=location.hash.match(/#([tz])=(.+)$/);if(!m)return location.hash.length>3&&/^#[tz]/.test(location.hash)?{error:true}:null;
  try{
    const raw=m[1]==="z"?LZString.decompressFromUint8Array(unb64u(decodeURIComponent(m[2]))):LZString.decompressFromEncodedURIComponent(m[2]);
    const o=JSON.parse(raw);if(!o||!o.t||!Array.isArray(o.t[2]))throw 0;
    const num=x=>{const n=+x;if(!isFinite(n))throw 0;return n};
    const str=x=>x==null?"":String(x);
    const clients={};
    const T={code:str(o.t[0]).replace(/[^A-Z0-9]/gi,"").slice(0,3)||"A",color:/^#[0-9a-f]{6}$/i.test(o.t[1])?o.t[1]:"#1d4ed8",stops:o.t[2].map(s=>({lat:num(s[0]),lon:num(s[1]),prec:s[2]===1?"civico":s[2]===2?"nome":"via",web:s[4]||"",acc:s[3].map(a=>{a=a.map(str);const cid=String(a[0]).startsWith(a[1]+"-")?a[0]:a[1]+"-"+a[0];a[0]=cid;clients[cid]={id:cid,src:a[1],rs:a[2],ins:a[3],piva:a[4],cat:a[5],addr:a[6],trend:a[7]};return a[0]})}))};
    T.stops.forEach(s=>{s.addr=cleanAddr(clients[s.acc[0]].addr)});
    const sync=/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(o.s||"")?o.s:"";
    const date=/^\d{4}-\d{2}-\d{2}$/.test(o.d)?o.d:new Date().toISOString().slice(0,10);
    return {sync,id:str(o.id).replace(/[^\w]/g,"")||"x",date,city:str(o.c),ret:!!o.r,meet:{lat:num(o.m[0]),lon:num(o.m[1]),name:str(o.m[2])},clients,teams:[T],shared:true};
  }catch(e){return {error:true}}
}
function current(){return shared||plan}

/* ---------- esiti (per telefono) ---------- */
function eKey(P){return "esiti:"+P.id}
function getE(P){
  const loc=LS.get(eKey(P),{}),rem=LS.get("remote:"+P.id,{});
  const out=Object.assign({},loc);
  for(const id in rem){const r=rem[id];if(!out[id]||(r.ts||0)>(out[id].ts||0)){if(r.esito||r.nota)out[id]=r;else delete out[id]}}
  return out;
}
function setE(P,id,patch){
  const all=LS.get(eKey(P),{});
  all[id]=Object.assign(all[id]||{},patch,{at:new Date().toTimeString().slice(0,5),ts:Date.now()});
  const e=all[id];LS.set(eKey(P),all);
  if(P.sync)queueSync(P,id,e);
}
/* ---------- sincronizzazione con Foglio Google ---------- */
function renderSheetLink(){const el=$("#sheet-link");if(!el)return;el.innerHTML=settings.sheetUrl?`<a class="btn small" href="${esc(settings.sheetUrl)}" target="_blank" rel="noopener">Apri il foglio degli esiti</a>`:""}
const DEV=LS.get("dev",null)||(()=>{const d=/iPhone/.test(navigator.userAgent)?"iPhone-"+uid().slice(-4):"tel-"+uid().slice(-4);LS.set("dev",d);return d})();
function teamOf(P,id){const T=P.teams.find(T=>T.stops.some(s=>s.acc.includes(id)));return T?T.code:""}
function queueSync(P,id,e){
  const c=P.clients[id]||{};
  const q=LS.get("q",[]).filter(x=>!(x.plan===P.id&&x.id===id));
  q.push({url:P.sync,plan:P.id,date:P.date,team:teamOf(P,id),id,rs:c.rs,ins:c.ins,addr:c.addr,src:c.src,esito:e.esito||"",nota:e.nota||"",at:e.at,ts:e.ts,dev:DEV});
  LS.set("q",q);flushSync();
}
let flushing=false;
async function flushSync(){
  if(flushing)return;const q=LS.get("q",[]);if(!q.length){syncState("ok");return}
  flushing=true;
  const byUrl={};q.forEach(x=>{(byUrl[x.url]=byUrl[x.url]||[]).push(x)});
  let left=[];
  for(const url in byUrl){
    try{await fetch(url,{method:"POST",mode:"no-cors",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(byUrl[url].map(({url,...r})=>r))})}
    catch(e){left=left.concat(byUrl[url])}
  }
  const now=LS.get("q",[]);const sent=new Set(q.filter(x=>!left.includes(x)).map(x=>x.plan+"|"+x.id+"|"+x.ts));
  LS.set("q",now.filter(x=>!sent.has(x.plan+"|"+x.id+"|"+x.ts)));
  flushing=false;syncState(left.length?"off":"ok");
}
let lastPull=null;
async function pullSync(){
  const P=current();if(!P||!P.sync)return;
  try{
    const u=P.sync+(P.sync.includes("?")?"&":"?")+"plan="+encodeURIComponent(P.id)+(P.shared?"&team="+encodeURIComponent(P.teams[0].code):"");
    const r=await fetch(u);const j=await r.json();if(!j||!j.rows)throw 0;
    const rem={};j.rows.forEach(x=>{rem[x.id]={esito:x.esito,nota:x.nota,at:x.at,ts:x.ts,team:x.team}});
    const before=JSON.stringify(LS.get("remote:"+P.id,{}));LS.set("remote:"+P.id,rem);
    lastPull=new Date().toTimeString().slice(0,5);syncState("ok");
    if(before!==JSON.stringify(rem)&&!document.activeElement.classList.contains("note")){drawMap();renderPanel()}
    else{const el=$("#sync-at");if(el)el.textContent=lastPull}
  }catch(e){syncState("off")}
}
function syncState(st){const el=$("#sync-dot");if(el){el.className="sync "+st;el.title=st==="ok"?"Esiti sincronizzati":"In attesa di connessione"}}
let syncStarted=false;
function startSync(){
  if(!current()||!current().sync)return;
  flushSync();pullSync();
  if(syncStarted)return;syncStarted=true;
  setInterval(()=>{if(!document.hidden){flushSync();pullSync()}},30000);
  window.addEventListener("online",()=>{flushSync();pullSync()});
  document.addEventListener("visibilitychange",()=>{if(!document.hidden){flushSync();pullSync()}});
}
function summary(P,T){
  const E=getE(P);let done=[],todo=[];
  T.stops.forEach((s,i)=>s.acc.forEach(id=>{const c=P.clients[id],e=E[id];const line=`${i+1}. ${c.rs}${c.ins?" ("+c.ins+")":""} – ${c.addr}`;if(e&&e.esito)done.push(line+` → ${e.esito}${e.nota?" · "+e.nota:""} [${e.at}]`);else todo.push(line+(e&&e.nota?" · "+e.nota:""))}));
  return `Canvassing ${P.city} ${P.date} – Coppia ${T.code}\nVisitati ${done.length}/${done.length+todo.length}\n\n`+(done.length?"FATTI\n"+done.join("\n")+"\n\n":"")+(todo.length?"DA FARE\n"+todo.join("\n"):"");
}

/* ---------- link navigazione ---------- */
function appleNav(p,label){return "https://maps.apple.com/?daddr="+p.lat.toFixed(6)+","+p.lon.toFixed(6)+"&dirflg="+(LS.get("walk",false)?"w":"d")+(label?"&q="+encodeURIComponent(label):"")}
function gLegs(P,T){
  const pts=[{lat:P.meet.lat,lon:P.meet.lon,n:0},...T.stops.map((s,i)=>({lat:s.lat,lon:s.lon,n:i+1}))];
  if(P.ret)pts.push({lat:P.meet.lat,lon:P.meet.lon,n:"R"});
  const legs=[];const STEP=9;
  for(let i=0;i<pts.length-1;i+=STEP-1){
    const seg=pts.slice(i,i+STEP);if(seg.length<2)break;
    const ll=p=>p.lat.toFixed(6)+","+p.lon.toFixed(6);
    const u="https://www.google.com/maps/dir/?api=1&travelmode=driving&origin="+ll(seg[0])+"&destination="+ll(seg[seg.length-1])+(seg.length>2?"&waypoints="+encodeURIComponent(seg.slice(1,-1).map(ll).join("|")):"");
    const lab=p=>p.n===0?"ritrovo":p.n==="R"?"rientro":"tappa "+p.n;
    legs.push({u,t:lab(seg[0])+" → "+lab(seg[seg.length-1])});
  }
  // percorso intero: formato /maps/dir/punto/punto/... (Google Maps può tagliare oltre le ~10 fermate nell'app)
  if(pts.length>2){
    const all="https://www.google.com/maps/dir/"+pts.map(p=>p.lat.toFixed(6)+","+p.lon.toFixed(6)).join("/")+"/data=!4m2!4m1!3e0";
    legs.unshift({u:all,t:"Percorso intero ("+T.stops.length+" tappe)",full:true});
  }
  return legs;
}

/* ---------- mappa (Leaflet o Google) ---------- */
let MAP=null;
const Leaf={
  init(){
    this.m=L.map("map",{zoomControl:false,attributionControl:true,tapTolerance:20,bounceAtZoomLimits:false,zoomSnap:0.25,zoomDelta:0.5,wheelPxPerZoomLevel:90}).setView([45.89,11.04],13);
    L.control.zoom({position:"bottomright"}).addTo(this.m);
    L.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png",{maxZoom:19,attribution:"© OpenStreetMap",keepBuffer:4,updateWhenIdle:false,detectRetina:false,crossOrigin:true}).addTo(this.m);
    this.layer=L.layerGroup().addTo(this.m);
    this.m.on("click",e=>MAP.onClick&&MAP.onClick({lat:e.latlng.lat,lon:e.latlng.lng}));
  },
  clear(){this.layer.clearLayers()},
  pin(p,html,cls,color,onclick,z){const ic=L.divIcon({className:"",html:`<div class="pin ${cls||""}" style="background:${color}">${html}</div>`,iconSize:[28,28],iconAnchor:[14,14]});const mk=L.marker([p.lat,p.lon],{icon:ic,zIndexOffset:z||0}).addTo(this.layer);if(onclick)mk.on("click",onclick);return mk},
  line(coords,color,dash){L.polyline(coords,{color,weight:dash?3:5,opacity:dash?.6:.75,dashArray:dash?"6 8":null}).addTo(this.layer)},
  fit(pts){if(!pts.length)return;this.m.invalidateSize();this.m.fitBounds(L.latLngBounds(pts.map(p=>[p.lat,p.lon])),{padding:[30,30],maxZoom:17})},
  pan(p,z){this.m.setView([p.lat,p.lon],Math.max(z||16,this.m.getZoom()))},
  me(p,acc){if(!this.meMk){this.meMk=L.marker([p.lat,p.lon],{icon:L.divIcon({className:"",html:'<div class="me"></div>',iconSize:[18,18],iconAnchor:[9,9]}),zIndexOffset:2000}).addTo(this.m);this.meC=L.circle([p.lat,p.lon],{radius:acc||20,color:"#1a73e8",weight:1,fillOpacity:.08}).addTo(this.m)}else{this.meMk.setLatLng([p.lat,p.lon]);this.meC.setLatLng([p.lat,p.lon]).setRadius(acc||20)}},
  resize(){this.m.invalidateSize();setTimeout(()=>this.m.invalidateSize(),80)}
};
const Goog={
  init(){
    this.m=new google.maps.Map($("#map"),{center:{lat:45.89,lng:11.04},zoom:13,disableDefaultUI:true,zoomControl:true,clickableIcons:false,gestureHandling:"greedy"});
    this.objs=[];this.m.addListener("click",e=>MAP.onClick&&MAP.onClick({lat:e.latLng.lat(),lon:e.latLng.lng()}));
  },
  clear(){this.objs.forEach(o=>o.setMap(null));this.objs=[]},
  pin(p,html,cls,color,onclick,z){
    const sel=/sel/.test(cls||""),done=/done/.test(cls||""),meet=/meet/.test(cls||"");const sz=sel?36:28;
    const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${sz}" height="${sz}"><${meet?`rect x="2" y="2" width="${sz-4}" height="${sz-4}" rx="7"`:`circle cx="${sz/2}" cy="${sz/2}" r="${sz/2-2}"`} fill="${meet?"#111":color}" fill-opacity="${done?.45:1}" stroke="${sel?"#f2b705":"#fff"}" stroke-width="${sel?4:2}"/><text x="50%" y="54%" dominant-baseline="middle" text-anchor="middle" font-family="Arial" font-weight="700" font-size="13" fill="#fff">${html}</text></svg>`;
    const mk=new google.maps.Marker({position:{lat:p.lat,lng:p.lon},map:this.m,zIndex:(z||0)+10,icon:{url:"data:image/svg+xml;charset=UTF-8,"+encodeURIComponent(svg),anchor:new google.maps.Point(sz/2,sz/2)}});
    if(onclick)mk.addListener("click",onclick);this.objs.push(mk);return mk;
  },
  line(coords,color,dash){const o=new google.maps.Polyline({path:coords.map(c=>({lat:c[0],lng:c[1]})),map:this.m,strokeColor:color,strokeOpacity:dash?0:.75,strokeWeight:dash?3:5,icons:dash?[{icon:{path:"M 0,-1 0,1",strokeOpacity:.7,scale:3},offset:"0",repeat:"12px"}]:null});this.objs.push(o)},
  fit(pts){if(!pts.length)return;const b=new google.maps.LatLngBounds();pts.forEach(p=>b.extend({lat:p.lat,lng:p.lon}));this.m.fitBounds(b,30);google.maps.event.addListenerOnce(this.m,"idle",()=>{if(this.m.getZoom()>17)this.m.setZoom(17)})},
  pan(p,z){this.m.panTo({lat:p.lat,lng:p.lon});if(this.m.getZoom()<(z||16))this.m.setZoom(z||16)},
  me(p,acc){if(!this.meMk){this.meMk=new google.maps.Marker({map:this.m,zIndex:9999,icon:{path:google.maps.SymbolPath.CIRCLE,scale:8,fillColor:"#1a73e8",fillOpacity:1,strokeColor:"#fff",strokeWeight:3}})}this.meMk.setPosition({lat:p.lat,lng:p.lon})},
  resize(){}
};
async function ensureMap(){
  if(MAP)return;
  const g=await loadGoogle();
  MAP=g?Goog:Leaf;MAP.init();
  $("#mapnote").textContent=g?"":"";
}
let fitted=false;
function drawMap(){
  const P=current();if(!MAP||!P)return;
  MAP.clear();
  const E=getE(P);
  const showAll=!P.shared&&curTeam===-1;
  const teams=showAll?P.teams:[P.teams[curTeam]].filter(Boolean);
  const pts=[P.meet];
  teams.forEach(T=>{
    T.stops.forEach((s,i)=>{
      const done=s.acc.every(id=>E[id]&&E[id].esito);
      const sel=selStop&&selStop.t===T.code&&selStop.i===i;
      MAP.pin(s,showAll?T.code+(i+1):String(i+1),(done?"done ":"")+(sel?"sel":""),T.color,()=>{selectStop(P.teams.indexOf(T),i,true)},sel?1000:0);
      pts.push(s);
    });
  });
  MAP.pin(P.meet,"P","meet","#111",()=>toast("Ritrovo: "+P.meet.name),500);
  if(!fitted||drawMap.refit){MAP.fit(pts);fitted=true;drawMap.refit=false}
}

/* ---------- posizione live ---------- */
let watchId=null,mePos=null,follow=false;
function startGps(center){
  if(!navigator.geolocation){toast("Posizione non disponibile su questo browser");return}
  if(watchId==null){
    watchId=navigator.geolocation.watchPosition(p=>{mePos={lat:p.coords.latitude,lon:p.coords.longitude};MAP&&MAP.me(mePos,p.coords.accuracy);if(follow){MAP.pan(mePos,16);follow=false}renderNext()},
      e=>{toast(e.code===1?"Permesso posizione negato: attivalo in Impostazioni › Safari › Posizione":"Posizione non disponibile");watchId=null},
      {enableHighAccuracy:true,maximumAge:5000,timeout:20000});
  }
  if(center){if(mePos)MAP.pan(mePos,16);else follow=true}
}

/* ---------- pannello ---------- */
function renderPairbar(){
  const P=current(),bar=$("#pairbar");bar.innerHTML="";
  if(!P){bar.classList.add("hidden");return}
  bar.classList.toggle("hidden",!!P.shared);
  if(P.shared)return;
  const mk=(lab,i,c)=>{const b=document.createElement("button");b.textContent=lab;b.style.setProperty("--c",c);if(i===curTeam)b.classList.add("on");b.onclick=()=>{curTeam=i;selStop=null;drawMap.refit=true;drawMap();renderPairbar();renderPanel()};bar.appendChild(b)};
  P.teams.forEach((T,i)=>mk("Coppia "+T.code+" · "+T.stops.reduce((s,x)=>s+x.acc.length,0),i,T.color));
  mk("Tutte",-1,"#0b2d5b");
}
function nextStop(P,T){
  const E=getE(P);const i=T.stops.findIndex(s=>!s.acc.every(id=>E[id]&&E[id].esito));return i;
}
function renderNext(){
  const box=$("#nextbox");const P=current();if(!box||!P||curTeam<0)return;
  const T=P.teams[curTeam];const i=nextStop(P,T);
  if(i<0){box.innerHTML=`<h3>Giro completato</h3><div>Tutti i clienti hanno un esito. Manda il riepilogo con il pulsante qui sopra.</div>`;return}
  const s=T.stops[i];const d=mePos?km(mePos,s):null;
  const names=s.acc.map(id=>P.clients[id].ins||P.clients[id].rs).join(", ");
  box.innerHTML=`<h3>Prossima tappa · ${i+1}${d!=null?` · ${d<1?Math.round(d*1000)+" m":d.toFixed(1)+" km"} da te`:""}</h3><div><b>${esc(s.addr)}</b></div><div class="small muted">${esc(names)}</div>
   <div class="row"><a class="btn nav small" href="${appleNav(s,names)}" target="_blank" rel="noopener">Naviga</a><button class="btn small" data-go="${i}">Vedi sulla mappa</button></div>`;
  box.querySelector("[data-go]").onclick=()=>selectStop(curTeam,i,true);
}
function renderPanel(){
  const P=current(),panel=$("#panel");
  if(!P){panel.innerHTML=`<div class="empty">Nessun percorso ancora.<br>Vai su <b>Liste</b>, carica gli Excel e premi <b>Calcola percorsi</b>.</div>`;return}
  if(curTeam===-1)return renderAll(P,panel);
  const T=P.teams[curTeam];if(!T){curTeam=0;return renderPanel()}
  const E=getE(P);
  const nAcc=T.stops.reduce((s,x)=>s+x.acc.length,0),nDone=T.stops.reduce((s,x)=>s+x.acc.filter(id=>E[id]&&E[id].esito).length,0);
  let h=`<div class="sum"><span style="color:${T.color};font-weight:800">Coppia ${T.code}</span><span><b id="cnt-done">${nDone}/${nAcc}</b> visitati</span><span class="pbar" aria-hidden="true"><i id="pbar" style="width:${nAcc?Math.round(nDone/nAcc*100):0}%;background:${T.color}"></i></span><span><b>${T.stops.length}</b> tappe</span><span><b>${T.km!=null?T.km:"…"}</b> km${T.geom?"":" circa"}</span>${T.km!=null?`<span>≈ <b>${durTxt(T,nAcc,LS.get("walk",false))}</b></span>`:""}<span>Ritrovo: <b>${esc(P.meet.name)}</b></span>${P.sync?`<span><i id="sync-dot" class="sync ok"></i>Esiti condivisi${lastPull?` · <span id="sync-at">${lastPull}</span>`:""}</span>`:""}</div>
  <div class="acts">
    <a class="btn small" href="${appleNav(P.meet,P.meet.name)}" target="_blank" rel="noopener">Naviga al ritrovo</a>
    <button class="btn small" id="b-wa">Riepilogo su WhatsApp</button>
    <label class="chk small"><input type="checkbox" id="s-walk" ${LS.get("walk",false)?"checked":""}> Naviga a piedi</label>
    <label class="chk small"><input type="checkbox" id="s-hide" ${LS.get("hide",false)?"checked":""}> Nascondi visitati</label>
    ${!P.shared&&P.sync&&settings.sheetUrl?`<a class="btn small" href="${esc(settings.sheetUrl)}" target="_blank" rel="noopener">Foglio esiti</a>`:""}
    ${P.shared?"":`<button class="btn small primary" id="b-share">Invia link alla coppia ${T.code}</button>`}
  </div>
  ${!P.shared&&P.missing&&P.missing.length?`<div class="msg warn" style="margin-bottom:10px"><b>${P.missing.length} clienti senza posizione</b> non sono in nessuna coppia: li trovi in fondo alla pagina per posizionarli sulla mappa.</div>`:""}
  <div id="nextbox" class="next"></div>
  <div class="legs small muted" style="margin-bottom:10px"><div style="margin-bottom:6px">Percorso su Google Maps:</div>${gLegs(P,T).map(l=>`<a class="btn small${l.full?" primary":""}" href="${l.u}" target="_blank" rel="noopener">${esc(l.t)}</a>`).join("")}${T.stops.length>9?`<div class="small muted" style="margin-top:4px">Con più di 10 fermate l'app Google Maps può mostrare solo le prime: in quel caso usa gli spezzoni.</div>`:""}</div>`;
  T.stops.forEach((s,i)=>{
    const names=s.acc.map(id=>P.clients[id].ins||P.clients[id].rs).join(", ");
    const sel=selStop&&selStop.t===T.code&&selStop.i===i;
    h+=`<article class="stop${sel?" sel":""}${stopDone(E,s)?" done":""}" id="st-${i}"><div class="hd"><div class="num" style="background:${T.color}">${i+1}</div><div><div class="addr">${esc(s.addr)}</div>
      <div class="meta">${s.prec==="nome"?`<span class="prec-zona">${esc(s.web)}</span>`:s.web?`<span class="prec-ok">${esc(s.web)}</span>`:s.prec==="civico"?"":s.prec==="manuale"?"Posizione impostata a mano":`<span class="prec-via">Posizione indicativa sulla via: cerca il civico</span>`}</div></div></div>`;
    s.acc.forEach(id=>{const c=P.clients[id],e=E[id]||{};
      h+=`<div class="acc"><div class="rs"><span class="tag ${esc(c.src)}">${esc(c.src)}</span>${esc(c.ins||c.rs)}</div>
        <div class="meta">${c.ins?esc(c.rs)+" · ":""}${c.piva?"P.IVA "+esc(c.piva)+" · ":""}${esc(c.cat||"")}${c.trend?" · "+esc(c.trend):""}</div>
        <div class="esiti" data-id="${esc(id)}">${ESITI.slice(1).map(x=>`<button data-e="${esc(x)}" class="${e.esito===x?"on":""}" aria-pressed="${e.esito===x}">${esc(x)}</button>`).join("")}</div>
        <input class="note" data-id="${esc(id)}" placeholder="Note (referente, telefono, richiamare il…)" value="${esc(e.nota||"")}"></div>`});
    h+=`<div class="ft"><a class="btn nav small" href="${appleNav(s,names)}" target="_blank" rel="noopener">Naviga</a><button class="btn small" data-map="${i}">Mappa</button></div>`;
    if(!P.shared)h+=`<div class="mgr">Sposta in <select data-move="${i}"><option value="">—</option>${P.teams.map((t,j)=>j===curTeam?"":`<option value="${j}">Coppia ${t.code}</option>`).join("")}</select><button class="btn small ghost" data-pos="${i}">Correggi posizione</button></div>`;
    h+=`</article>`;
  });
  if(!P.shared&&P.missing&&P.missing.length){
    h+=`<div class="card"><h2>Indirizzi non trovati (${P.missing.length})</h2><p class="muted small">Tocca "Posiziona", poi tocca il punto giusto sulla mappa: il cliente viene aggiunto alla coppia selezionata.</p><table class="cl">${P.missing.map(id=>{const c=P.clients[id];return `<tr><td><b>${esc(c.rs)}</b><br>${esc(c.addr)}</td><td><button class="btn small" data-place="${esc(id)}">Posiziona</button></td></tr>`}).join("")}</table></div>`;
  }
  panel.innerHTML=h;
  renderNext();
  panel.querySelectorAll(".esiti").forEach(row=>row.onclick=ev=>{const b=ev.target.closest("button");if(!b)return;const id=row.dataset.id;const cur=(getE(P)[id]||{}).esito;const val=cur===b.dataset.e?"":b.dataset.e;setE(P,id,{esito:val});
    row.querySelectorAll("button").forEach(x=>{x.classList.toggle("on",x.dataset.e===val);x.setAttribute("aria-pressed",x.dataset.e===val)});
    if(navigator.vibrate)navigator.vibrate(10);
    refreshProgress(P,T,val?id:null);});
  panel.querySelectorAll("input.note").forEach(n=>{const save=()=>{clearTimeout(n._t);const v=n.value.trim();if(((getE(P)[n.dataset.id]||{}).nota||"")!==v)setE(P,n.dataset.id,{nota:v})};
    n.oninput=()=>{clearTimeout(n._t);n._t=setTimeout(save,700)};n.onchange=save;n.onblur=save;n.onkeydown=e=>{if(e.key==="Enter")n.blur()}});
  panel.querySelectorAll("[data-map]").forEach(b=>b.onclick=()=>selectStop(curTeam,+b.dataset.map,true));
  $("#s-walk").onchange=e=>{LS.set("walk",e.target.checked);renderPanel()};
  $("#s-hide").onchange=e=>{LS.set("hide",e.target.checked);applyHide(P,T)};applyHide(P,T);
  $("#b-wa").onclick=()=>{const txt=summary(P,T);window.open("https://wa.me/?text="+encodeURIComponent(txt),"_blank")};
  if(!P.shared){
    $("#b-share").onclick=()=>shareSheet(P,curTeam);
    panel.querySelectorAll("[data-move]").forEach(sl=>sl.onchange=()=>{const j=+sl.value;if(isNaN(j)||sl.value==="")return;const s=T.stops.splice(+sl.dataset.move,1)[0];P.teams[j].stops.push(s);reorderTeam(T,P.meet,P.ret);reorderTeam(P.teams[j],P.meet,P.ret);savePlan();toast("Spostato nella coppia "+P.teams[j].code);selStop=null;fetchGeoms()});
    panel.querySelectorAll("[data-pos]").forEach(b=>b.onclick=()=>{moveMode={kind:"stop",t:curTeam,i:+b.dataset.pos};expandMap(true);toast("Tocca sulla mappa il punto esatto",4000)});
    panel.querySelectorAll("[data-place]").forEach(b=>b.onclick=()=>{moveMode={kind:"missing",id:b.dataset.place};expandMap(true);toast("Tocca sulla mappa dove si trova il cliente",4000)});
  }
}
function renderAll(P,panel){
  let h=`<div class="sum"><span>Ritrovo: <b>${esc(P.meet.name)}</b></span><span>${P.date}</span></div>`+(P.missing&&P.missing.length?`<div class="msg warn" style="margin-bottom:10px"><b>${P.missing.length} clienti senza posizione</b>: apri una coppia e scorri in fondo per posizionarli.</div>`:"");
  P.teams.forEach((T,i)=>{const n=T.stops.reduce((s,x)=>s+x.acc.length,0);
    h+=`<div class="card"><h2 style="color:${T.color}">Coppia ${T.code}</h2><div class="sum"><span><b>${n}</b> clienti</span><span><b>${T.stops.length}</b> tappe</span><span><b>${T.km!=null?T.km:"…"}</b> km</span></div>
      <div class="row"><button class="btn small primary" data-sh="${i}">Invia link</button><button class="btn small" data-open="${i}">Apri</button></div></div>`});
  h+=`<div class="card"><h2>Modifica</h2><p class="muted small">Ricalcola le zone da zero, ad esempio dopo aver spostato il ritrovo o cambiato il numero di coppie.</p><button class="btn small" id="b-recalc">Ricalcola zone e percorsi</button> <button class="btn small ghost" id="b-meetmove">Sposta ritrovo sulla mappa</button></div>`;
  panel.innerHTML=h;
  panel.querySelectorAll("[data-sh]").forEach(b=>b.onclick=()=>shareSheet(P,+b.dataset.sh));
  panel.querySelectorAll("[data-open]").forEach(b=>b.onclick=()=>{curTeam=+b.dataset.open;drawMap.refit=true;drawMap();renderPairbar();renderPanel()});
  $("#b-recalc").onclick=()=>{const ids=P.teams.flatMap(T=>T.stops);const k=Math.min(settings.k,ids.length);P.teams=makeTeams(ids,k,P.meet,P.ret);savePlan();drawMap.refit=true;fetchGeoms();renderPairbar();toast("Zone ricalcolate")};
  $("#b-meetmove").onclick=()=>{moveMode={kind:"meet"};expandMap(true);toast("Tocca sulla mappa il nuovo punto di ritrovo",4000)};
}
/* stima: spostamento (a piedi 4,5 km/h, in auto 25 km/h in città + 5 min a tappa per parcheggio) + 12 min a cliente */
function durTxt(T,nAcc,walk){const m=Math.round((walk?T.km/4.5*60:T.km/25*60+T.stops.length*5)+nAcc*12);return m<60?m+" min":Math.floor(m/60)+" h"+(m%60?" "+String(m%60).padStart(2,"0"):"")}
function stopDone(E,s){return s.acc.every(id=>E[id]&&E[id].esito)}
function refreshProgress(P,T,justId){
  const E=getE(P);
  const pb=$("#pbar");const nAcc=T.stops.reduce((s,x)=>s+x.acc.length,0),nDone=T.stops.reduce((s,x)=>s+x.acc.filter(id=>E[id]&&E[id].esito).length,0);
  const c=$("#cnt-done");if(c)c.textContent=nDone+"/"+nAcc;if(pb)pb.style.width=(nAcc?Math.round(nDone/nAcc*100):0)+"%";
  if(justId){const si=T.stops.findIndex(s=>s.acc.includes(justId));if(si>=0&&stopDone(E,T.stops[si])){const ni=nextStop(P,T);
    if(ni<0)toast("Giro completato! Manda il riepilogo su WhatsApp",4000);else if(ni!==si)toast("Tappa "+(si+1)+" fatta · prossima: "+(ni+1)+" – "+T.stops[ni].addr,3200)}}
  T.stops.forEach((s,i)=>{const el=$("#st-"+i);if(el)el.classList.toggle("done",stopDone(E,s))});
  renderNext();drawMap();
}
function applyHide(P,T){const on=LS.get("hide",false);document.body.classList.toggle("hide-done",on)}
function savePlan(){if(plan)LS.set("plan",plan)}
function selectStop(ti,i,scroll){
  const P=current();let changed=false;if(curTeam!==ti&&!(P.shared)){curTeam=ti;changed=true;renderPairbar()}
  const T=P.teams[ti];selStop={t:T.code,i};
  drawMap();
  if(!changed&&$("#st-"+i)){document.querySelectorAll(".stop.sel").forEach(x=>x.classList.remove("sel"));$("#st-"+i).classList.add("sel")}else renderPanel();
  MAP.pan(T.stops[i],16);
  if(scroll){const el=$("#st-"+i);if(el)el.scrollIntoView({behavior:"smooth",block:"center"})}
}
function shareSheet(P,ti){
  const T=P.teams[ti];const url=shareUrl(P,ti);
  const msg=`Canvassing ${P.city} ${P.date} – Coppia ${T.code}: ${T.stops.reduce((s,x)=>s+x.acc.length,0)} clienti. Apri il link (in Safari) per mappa, tappe e navigazione:\n${url}`;
  const sh=$("#sheet"),inn=sh.querySelector(".sheet-in");
  inn.innerHTML=`<h3>Link per la coppia ${T.code}</h3><p class="small muted">Il link contiene i clienti di questa coppia (nomi, indirizzi, P.IVA): chi lo riceve li vede. Non passa da nessun server. Gli esiti restano sul telefono di chi li segna: a fine giro la coppia ti manda il riepilogo su WhatsApp.</p>
    <a class="btn primary" href="https://wa.me/?text=${encodeURIComponent(msg)}" target="_blank" rel="noopener">Invia su WhatsApp</a>
    ${navigator.share?'<button class="btn" id="sh-native">Condividi…</button>':""}
    <button class="btn" id="sh-copy">Copia link</button><button class="btn ghost" id="sh-close">Chiudi</button>`;
  sh.classList.remove("hidden");
  const close=()=>sh.classList.add("hidden");
  sh.onclick=e=>{if(e.target===sh)close()};$("#sh-close").onclick=close;
  $("#sh-copy").onclick=async()=>{try{await navigator.clipboard.writeText(url);toast("Link copiato")}catch(e){prompt("Copia il link:",url)}};
  if(navigator.share)$("#sh-native").onclick=()=>navigator.share({title:"Canvassing coppia "+T.code,text:msg.split("\n")[0],url}).catch(()=>{});
}
function expandMap(on){const w=$("#mapwrap");w.classList.toggle("full",on);MAP&&MAP.resize();setTimeout(()=>MAP&&MAP.resize(),320)}

function onMapClick(p){
  if(!moveMode)return;
  const P=current();if(!P||P.shared){moveMode=null;return}
  if(moveMode.kind==="stop"){const s=P.teams[moveMode.t].stops[moveMode.i];s.lat=p.lat;s.lon=p.lon;s.prec="manuale";reorderTeam(P.teams[moveMode.t],P.meet,P.ret);toast("Posizione aggiornata e percorso ricalcolato")}
  else if(moveMode.kind==="missing"){const id=moveMode.id;const c=P.clients[id];c.lat=p.lat;c.lon=p.lon;c.prec="manuale";
    const ti=curTeam<0?0:curTeam;P.teams[ti].stops.push({lat:p.lat,lon:p.lon,prec:"manuale",addr:cleanAddr(c.addr),acc:[id]});P.missing=P.missing.filter(x=>x!==id);reorderTeam(P.teams[ti],P.meet,P.ret);toast("Aggiunto alla coppia "+P.teams[ti].code)}
  else if(moveMode.kind==="meet"){P.meet={lat:p.lat,lon:p.lon,name:"Ritrovo (posizionato a mano)"};P.teams.forEach(T=>reorderTeam(T,P.meet,P.ret));toast("Ritrovo spostato, percorsi ricalcolati")}
  moveMode=null;expandMap(false);savePlan();fetchGeoms();
}

/* ---------- navigazione tra pagine ---------- */
function renderTabs(){
  const tabs=$("#tabs");tabs.innerHTML="";
  if(shared)return;
  [["liste","Liste"],["mappa","Mappa e percorsi"]].forEach(([k,t])=>{const b=document.createElement("button");b.textContent=t;b.dataset.k=k;b.onclick=()=>go(k);tabs.appendChild(b)});
}
async function go(pg){
  document.querySelectorAll(".page").forEach(p=>p.classList.toggle("hidden",p.id!=="pg-"+pg));
  document.querySelectorAll("#tabs button").forEach(b=>b.classList.toggle("on",b.dataset.k===pg));
  if(!shared)LS.set("page",pg);
  if(pg==="mappa"){
    await ensureMap();MAP.onClick=onMapClick;MAP.resize();
    drawMap.refit=true;drawMap();renderPairbar();renderPanel();
    const P=current();if(P&&P.teams.some(T=>T.km==null))fetchGeoms();
  }
  window.scrollTo(0,0);
}

/* ---------- avvio ---------- */
function initSettingsUI(){
  $("#s-city").value=settings.city;$("#s-k").value=settings.k;$("#s-meet").value=settings.meet;$("#s-ret").checked=settings.ret;$("#s-gkey").value=settings.gkey||"";$("#s-sync").value=settings.sync||"";renderSheetLink();if(location.hash==="#impostazioni"){$("details").open=true;setTimeout(()=>$("#s-sync").scrollIntoView({block:"center"}),200)}
  ["#s-city","#s-k","#s-meet","#s-ret"].forEach(s=>$(s).onchange=()=>{settings.city=$("#s-city").value.trim();settings.k=+$("#s-k").value;settings.meet=$("#s-meet").value.trim();settings.ret=$("#s-ret").checked;LS.set("settings",settings)});
  $("#file").onchange=e=>{onFiles([...e.target.files]);e.target.value=""};
  $("#b-calc").onclick=calculate;
  $("#b-meet-gps").onclick=()=>{navigator.geolocation&&navigator.geolocation.getCurrentPosition(p=>{settings.meetGps={lat:p.coords.latitude,lon:p.coords.longitude,name:"Ritrovo (posizione GPS)"};settings.meet="La mia posizione";$("#s-meet").value=settings.meet;LS.set("settings",settings);$("#meet-info").textContent="Ritrovo impostato sulla tua posizione attuale."},()=>toast("Posizione non disponibile"),{enableHighAccuracy:true,timeout:15000})};
  $("#b-gkey").onclick=()=>{settings.gkey=$("#s-gkey").value.trim();LS.set("settings",settings);toast("Chiave salvata: ricarico l'app");setTimeout(()=>location.reload(),900)};
  $("#b-sync").onclick=async()=>{
    const v=$("#s-sync").value.trim();
    if(v&&!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/.test(v)){toast("Il link deve iniziare con https://script.google.com/macros/s/ e finire con /exec",5000);return}
    settings.sync=v;LS.set("settings",settings);
    if(plan){plan.sync=v;savePlan()}
    if(!v){toast("Sincronizzazione disattivata");return}
    try{const r=await fetch(v);const j=await r.json();if(j&&j.sheet){settings.sheetUrl=j.sheet;LS.set("settings",settings);renderSheetLink()}toast(j&&j.ok?"Collegamento al foglio riuscito. Rimanda i link alle coppie.":"Il foglio non risponde come previsto",5000)}catch(e){toast("Non riesco a raggiungere il foglio: controlla che il deployment sia per \"Chiunque\"",6000)}
    startSync();
  };
  $("#b-gkey-del").onclick=()=>{settings.gkey="";$("#s-gkey").value="";LS.set("settings",settings);setTimeout(()=>location.reload(),300)};
  $("#b-reset").onclick=()=>{if(!confirm("Cancello liste, percorsi ed esiti salvati su questo telefono?"))return;["files","plan","geo","q"].forEach(LS.del);try{Object.keys(localStorage).filter(k=>k.startsWith("cv:esiti:")||k.startsWith("cv:remote:")).forEach(k=>localStorage.removeItem(k))}catch(e){}location.reload()};
}
$("#b-locate").onclick=()=>startGps(true);
$("#b-expand").onclick=()=>expandMap(!$("#mapwrap").classList.contains("full"));

function boot(){
  shared=readHash();
  if(shared&&!shared.error)LS.set("lastShared",{h:location.hash,at:Date.now()});
  else if(!shared&&!plan&&!files.length){const ls=LS.get("lastShared",null);if(ls&&Date.now()-ls.at<3*864e5&&/^#[tz]=/.test(ls.h)){history.replaceState(null,"",location.pathname+ls.h);shared=readHash();if(shared&&shared.error)shared=null}}
  if(shared&&shared.error){document.body.innerHTML=`<div class="page"><div class="card"><h2>Link incompleto</h2><p>Il link della coppia si è interrotto durante l'invio, quindi non contiene il percorso. Chiedi al responsabile di rimandarlo con il tasto <b>Copia link</b>, incollandolo da solo nel messaggio.</p></div></div>`;return}
  renderTabs();
  if(shared){
    const T=shared.teams[0];
    $("#title").textContent=`Coppia ${T.code} · ${shared.city} ${shared.date.split("-").reverse().join("/")}`;
    $("#top").style.background=T.color;
    curTeam=0;go("mappa");startGps(false);startSync();
    return;
  }
  initSettingsUI();renderFiles();
  go(location.hash!=="#impostazioni"&&plan&&LS.get("page","liste")==="mappa"?"mappa":"liste");
}
window.addEventListener("hashchange",()=>location.reload());
/* salva la nota in corso se Safari va in background o si chiude */
window.addEventListener("pagehide",()=>{const a=document.activeElement;if(a&&a.classList&&a.classList.contains("note"))a.blur()});
document.addEventListener("visibilitychange",()=>{if(document.hidden){const a=document.activeElement;if(a&&a.classList&&a.classList.contains("note"))a.blur()}});
function setHH(){document.documentElement.style.setProperty("--hh",$("#top").offsetHeight+"px")}
window.addEventListener("resize",setHH);
boot();setHH();if(!shared)startSync();
/* funziona anche con poco segnale; i dati del telefono non vengono cancellati da Safari */
if("serviceWorker" in navigator&&(location.protocol==="https:"||location.hostname==="localhost"))navigator.serviceWorker.register("sw.js").catch(()=>{});
try{navigator.storage&&navigator.storage.persist&&navigator.storage.persist()}catch(e){}
window.addEventListener("offline",()=>toast("Sei offline: gli esiti restano salvati sul telefono",3500));
})();
