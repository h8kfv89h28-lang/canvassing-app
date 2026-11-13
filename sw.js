/* Canvassing – service worker: l'app si apre anche con poco segnale.
   Pagina: prima la rete (versione aggiornata), poi la copia salvata.
   File con ?v=N: dalla copia salvata. Mappa (tile OSM): copia salvata delle zone già viste. */
const V="cv-v9", TILES="cv-tiles";
const SHELL=["./","index.html","app.css?v=9","app.js?v=9","vendor/leaflet.css","vendor/leaflet.js","vendor/lz-string.min.js","vendor/xlsx.core.min.js","icon.png","manifest.json"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()))});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==V&&k!==TILES).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
const timeout=(p,ms)=>Promise.race([p,new Promise((_,ko)=>setTimeout(()=>ko(new Error("timeout")),ms))]);
async function trimTiles(c){const ks=await c.keys();for(let i=0;i<ks.length-600;i++)await c.delete(ks[i])}
self.addEventListener("fetch",e=>{
  const r=e.request;if(r.method!=="GET")return;
  const u=new URL(r.url);
  if(u.origin===location.origin){
    if(r.mode==="navigate"){
      e.respondWith(timeout(fetch(r),6000).then(res=>{if(res.ok&&/\/(index\.html)?$/.test(u.pathname)){const cp=res.clone();caches.open(V).then(c=>c.put("./",cp))}return res})
        .catch(()=>caches.match(/\/(index\.html)?$/.test(u.pathname)?"./":r,{ignoreSearch:true}).then(m=>m||caches.match("./"))));
      return;
    }
    if(u.search.includes("v=")){e.respondWith(caches.match(r).then(m=>m||fetch(r).then(res=>{if(res.ok){const cp=res.clone();caches.open(V).then(c=>c.put(r,cp))}return res})));return}
    e.respondWith(caches.match(r).then(m=>{const net=fetch(r).then(res=>{if(res.ok){const cp=res.clone();caches.open(V).then(c=>c.put(r,cp))}return res}).catch(()=>m);return m||net}));
    return;
  }
  if(u.hostname==="tile.openstreetmap.org"){
    e.respondWith(caches.open(TILES).then(c=>c.match(r).then(m=>{
      const net=fetch(r).then(res=>{if(res.ok){c.put(r,res.clone());trimTiles(c)}return res});
      return m?(net.catch(()=>{}),m):net;
    })));
  }
});
