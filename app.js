/* Return Scanner: scan a parcel, record its opening, save the video and a small record to OneDrive. */
(function(){
"use strict";
const CFG=window.SCANNER_CONFIG||{};
const TEST=(location.hostname==="localhost"&&window.SC_TEST)||null;
const GRAPH=TEST?TEST.graph:"https://graph.microsoft.com/v1.0";
const SCOPES=["Files.ReadWrite.All"];
const MYNTRA=/^(MYER|MYEC|MYEP)\d{8,}$/, FLIPKART=/^FMP[A-Z]\d{8,}$/, TYPED=/^[A-Z0-9-]{6,40}$/;
const ISSUES=["Wrong item (my own product)","Other seller's product","Used, damaged or stained","Missing part","Tag or label removed","Short or empty box"];
const CHUNK=327680*16; /* 5 MiB; Graph needs multiples of 320 KiB */
const MIN_REC_SECONDS=3;
const $=s=>document.querySelector(s), app=$("#app");
const esc=s=>String(s==null?"":s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const store={get(k,d){try{const v=localStorage.getItem(k);return v==null?d:JSON.parse(v);}catch(e){return d;}},set(k,v){try{localStorage.setItem(k,JSON.stringify(v));}catch(e){}}};
const S={view:"home",who:store.get("who",""),index:store.get("index",null),done:store.get("done",[]),queue:[],barcode:"",lines:[],rec:null,authError:"",pumping:false,typed:""};

/* ---------- sign-in ---------- */
let pca=null, account=null;
async function initAuth(){
  if(TEST) return;
  if(!CFG.clientId||!CFG.tenantId) throw new Error("This app isn't set up yet: the Microsoft app IDs are missing from config.js.");
  pca=new msal.PublicClientApplication({auth:{clientId:CFG.clientId,authority:"https://login.microsoftonline.com/"+CFG.tenantId,redirectUri:location.origin+location.pathname},cache:{cacheLocation:"localStorage"}});
  await pca.initialize();
  const r=await pca.handleRedirectPromise();
  account=(r&&r.account)||pca.getAllAccounts()[0]||null;
}
function signIn(){ return pca.loginRedirect({scopes:SCOPES}); }
async function token(){
  if(TEST) return "test";
  if(!account) throw Object.assign(new Error("Not signed in"),{auth:true});
  try{ return (await pca.acquireTokenSilent({scopes:SCOPES,account})).accessToken; }
  catch(e){ throw Object.assign(new Error("Microsoft sign-in has expired"),{auth:true}); }
}
async function g(path,opts){
  const o=opts||{}; const headers=Object.assign({Authorization:"Bearer "+await token()},o.headers||{});
  return fetch(GRAPH+path,Object.assign({},o,{headers}));
}

/* ---------- OneDrive folder ---------- */
let folder=null;
async function getFolder(){
  if(folder) return folder;
  const cached=store.get("folder",null); if(cached&&cached.path===CFG.folderPath) return folder=cached;
  const r=await g("/me/drive/root:/"+CFG.folderPath.split("/").map(encodeURIComponent).join("/")+":?$select=id,name,parentReference");
  if(!r.ok) throw new Error("Could not open the OneDrive folder ("+r.status+")");
  const j=await r.json(); folder={path:CFG.folderPath,drive:j.parentReference.driveId,id:j.id}; store.set("folder",folder); return folder;
}
const itemPath=(f,p)=>"/drives/"+f.drive+"/items/"+f.id+":/"+p.split("/").map(encodeURIComponent).join("/")+":";
async function loadIndex(){
  try{
    const f=await getFolder(); const r=await g(itemPath(f,"_index/returns_index.json")+"?$select=id,@microsoft.graph.downloadUrl");
    if(!r.ok) return; const j=await r.json(); const u=j["@microsoft.graph.downloadUrl"]; if(!u) return;
    const d=await fetch(u); if(!d.ok) return;
    S.index=await d.json(); store.set("index",S.index); if(S.view==="home") render();
  }catch(e){ if(e.auth){ S.authError=e.message; if(S.view==="home") render(); } }
}

/* ---------- saved queue, kept on the device until OneDrive confirms ---------- */
let dbp=null;
function db(){ return dbp||(dbp=new Promise((res,rej)=>{ const q=indexedDB.open("return-scanner",1); q.onupgradeneeded=()=>q.result.createObjectStore("queue",{keyPath:"id"}); q.onsuccess=()=>res(q.result); q.onerror=()=>rej(q.error); })); }
async function tx(mode,fn){ const d=await db(); return new Promise((res,rej)=>{ const t=d.transaction("queue",mode); const r=fn(t.objectStore("queue")); t.oncomplete=()=>res(r&&r.result); t.onerror=()=>rej(t.error); t.onabort=()=>rej(t.error); }); }
const qPut=i=>tx("readwrite",s=>s.put(i)), qDel=id=>tx("readwrite",s=>s.delete(id)), qAll=()=>tx("readonly",s=>s.getAll());
async function refreshQueue(){ S.queue=(await qAll()).sort((a,b)=>a.at<b.at?-1:1); }

async function uploadVideo(item,f){
  /* Pick a free name so nothing is overwritten: BARCODE.mp4, then BARCODE_2.mp4 */
  if(!item.file){
    for(let n=1;;n++){ const name=item.barcode+(n===1?"":"_"+n)+"."+item.ext; const r=await g(itemPath(f,name)+"?$select=id"); if(r.status===404){ item.file=name; break; } if(!r.ok) throw new Error("OneDrive answered "+r.status); if(n>50) throw new Error("Too many videos for this barcode"); }
    await qPut(item);
  }
  if(!item.uploadUrl){
    const r=await g(itemPath(f,item.file)+"/createUploadSession",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({item:{"@microsoft.graph.conflictBehavior":"fail"}})});
    if(r.status===409){ item.file=null; await qPut(item); return uploadVideo(item,f); }
    if(!r.ok) throw new Error("Could not start the upload ("+r.status+")");
    item.uploadUrl=(await r.json()).uploadUrl; item.offset=0; await qPut(item);
  }else{
    const r=await fetch(item.uploadUrl);
    if(r.status===404||r.status===410){ item.uploadUrl=null; item.offset=0; await qPut(item); return uploadVideo(item,f); }
    if(r.ok){ const j=await r.json(); if(j.nextExpectedRanges&&j.nextExpectedRanges[0]) item.offset=parseInt(j.nextExpectedRanges[0],10); }
  }
  const size=item.blob.size;
  while(true){
    const start=item.offset||0, end=Math.min(start+CHUNK,size);
    const r=await fetch(item.uploadUrl,{method:"PUT",headers:{"Content-Range":"bytes "+start+"-"+(end-1)+"/"+size},body:item.blob.slice(start,end)});
    if(r.status===200||r.status===201) return r.json();
    if(r.status===202){ const j=await r.json(); item.offset=parseInt(j.nextExpectedRanges[0],10); item.pct=Math.round(item.offset/size*100); await qPut(item); await paintQueue(); continue; }
    if(r.status===404||r.status===410){ item.uploadUrl=null; item.offset=0; await qPut(item); return uploadVideo(item,f); }
    throw new Error("Upload stopped ("+r.status+")");
  }
}
async function uploadItem(item){
  const f=await getFolder();
  if(!item.driveItem){ const d=await uploadVideo(item,f); item.driveItem={webUrl:d.webUrl,size:d.size}; await qPut(item); }
  const rec=Object.assign({},item.record,{file:item.file,webUrl:item.driveItem.webUrl,size:item.driveItem.size});
  const r=await g(itemPath(f,"_records/"+item.file+".json")+"/content",{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(rec,null,1)});
  if(!r.ok) throw new Error("Could not save the record ("+r.status+")");
  S.done=[{barcode:item.barcode,file:item.file,result:item.record.result,at:item.at}].concat(S.done).slice(0,20); store.set("done",S.done);
  await qDel(item.id);
}
async function pump(){
  if(S.pumping) return; S.pumping=true;
  try{
    let delay=0;
    while(true){
      await refreshQueue(); await paintQueue();
      const item=S.queue.find(i=>!i.waitUntil||i.waitUntil<=Date.now());
      if(!item){ if(!S.queue.length) break; await sleep(4000); continue; }
      try{ item.error=""; await qPut(item); await paintQueue(); await uploadItem(item); delay=0; S.authError=""; }
      catch(e){ if(e.auth){ S.authError=e.message; if(S.view==="home") render(); break; } delay=Math.min((delay||15000)*2,300000); item.error=navigator.onLine?e.message:"Waiting for internet"; item.waitUntil=Date.now()+delay; await qPut(item); }
    }
  }finally{ S.pumping=false; await refreshQueue(); await paintQueue(); }
}
window.addEventListener("online",()=>{ qAll().then(a=>Promise.all(a.map(i=>{ i.waitUntil=0; return qPut(i); }))).then(pump); });

/* ---------- barcode reading ---------- */
let wasmReady=false;
function prepZX(){ if(wasmReady) return; wasmReady=true; ZXingWASM.prepareZXingModule({overrides:{locateFile:(p,s)=>p.endsWith(".wasm")?new URL("vendor/"+p,location.href).href:s+p}}); }
async function readCanvas(c){ prepZX(); const d=c.getContext("2d",{willReadFrequently:true}).getImageData(0,0,c.width,c.height); const res=await ZXingWASM.readBarcodes(d,{formats:["Code128"],tryHarder:true,tryRotate:true,maxNumberOfSymbols:4}); return res.filter(r=>r.isValid).map(r=>r.text.trim().toUpperCase()); }
let stream=null, canvas=document.createElement("canvas");
function stopStream(){ if(stream){ stream.getTracks().forEach(t=>{ t.onended=null; t.stop(); }); stream=null; } }
function camError(e){ return e&&e.name==="NotAllowedError"?"Camera access is blocked. Allow it in your browser settings, then try again.":e&&e.name==="NotFoundError"?"No camera was found on this device.":"The camera could not start: "+(e&&e.message||e); }
async function startScan(){
  S.view="scan"; S.scanMsg="Point the box at the long barcode under MYER, MYEC or MYEP."; render();
  try{ stream=await navigator.mediaDevices.getUserMedia({audio:false,video:{facingMode:{ideal:"environment"},width:{ideal:1920},height:{ideal:1080}}}); }
  catch(e){ S.view="home"; toast(camError(e)); render(); return; }
  const v=$("#cam"); if(!v){ stopStream(); return; } v.srcObject=stream; await v.play().catch(()=>{});
  const track=stream.getVideoTracks()[0], caps=track.getCapabilities?track.getCapabilities():{};
  if(caps.torch){ const b=$("#torch"); if(b) b.hidden=false; }
  if(caps.focusMode&&caps.focusMode.includes("continuous")) track.applyConstraints({advanced:[{focusMode:"continuous"}]}).catch(()=>{});
  let scale=1;
  while(S.view==="scan"&&stream){
    const w=v.videoWidth,h=v.videoHeight;
    if(w&&h){
      const cw=Math.round(w*.9),ch=Math.round(h*.6); canvas.width=Math.round(cw*scale); canvas.height=Math.round(ch*scale);
      const x=canvas.getContext("2d",{willReadFrequently:true}); x.imageSmoothingQuality="high"; x.drawImage(v,(w-cw)/2,(h-ch)/2,cw,ch,0,0,canvas.width,canvas.height);
      let hits=[]; try{ hits=await readCanvas(canvas); }catch(e){}
      const hit=hits.find(t=>MYNTRA.test(t))||hits.find(t=>FLIPKART.test(t));
      if(hit){ if(navigator.vibrate) navigator.vibrate(120); stopStream(); return chooseBarcode(hit); }
      if(hits.length&&S.view==="scan"){ const m=$("#scanMsg"); if(m) m.textContent="That's a bag barcode. Aim at the one under MYER, MYEC or MYEP."; }
      scale=scale===1?1.6:1;
    }
    await sleep(120);
  }
}
function toggleTorch(){ if(!stream) return; const t=stream.getVideoTracks()[0]; S.torch=!S.torch; t.applyConstraints({advanced:[{torch:S.torch}]}).catch(()=>{}); }

/* ---------- parcel lookup ---------- */
function chooseBarcode(code){
  S.barcode=code; S.lines=(S.index&&S.index.barcodes&&S.index.barcodes[code])||[]; S.view="confirm"; render();
}
function daysSince(s){ if(!s) return null; const d=new Date(String(s).slice(0,10)+"T00:00:00"); return Math.floor((Date.now()-d.getTime())/864e5); }
function parcelNotes(){
  const out=[], code=S.barcode, L=S.lines, limit=CFG.claimLimitDays||7;
  if(FLIPKART.test(code)) out.push(["bad","This isn't a Myntra parcel. A video can be saved, but the ledger won't change."]);
  else if(!L.length) out.push(["warn",S.index?"Not found in the ledger yet. This can happen if the report is a day behind. You can still record, and it will be matched later.":"The order list hasn't loaded yet. You can still record, and it will be matched later."]);
  const stages=new Set(L.map(l=>l.stage));
  if(stages.has("raised")) out.push(["warn","A claim is already raised for this parcel. A new video will be added, but the ledger status won't change."]);
  if(stages.has("ok")||stages.has("accepted")||stages.has("rejected")) out.push(["warn","This parcel is already closed in the ledger. A new video will be added, but the status won't change."]);
  if(stages.has("wrong")) out.push(["warn","This parcel is already marked wrong and waiting for a claim."]);
  if(S.done.some(d=>d.barcode===code)||S.queue.some(q=>q.barcode===code)) out.push(["warn","A video for this barcode was already recorded on this device. A new one will be saved as a separate file."]);
  const back=L.map(l=>l.backAt).filter(Boolean)[0];
  if(back){ const left=limit-daysSince(back); out.push(left<0?["bad","Past Myntra's "+limit+"-day claim limit by "+(-left)+" day"+(left===-1?"":"s")+". Open it on video anyway."]:left===0?["bad","Last day to raise a claim for this parcel."]:left<=3?["warn",left+" day"+(left===1?"":"s")+" left to raise a claim."]:["good",left+" days left to raise a claim."]); }
  return out;
}

/* ---------- recording ---------- */
function pickMime(){ const list=["video/mp4;codecs=avc1.42E01E,mp4a.40.2","video/mp4","video/webm;codecs=vp9,opus","video/webm;codecs=vp8,opus","video/webm"]; return list.find(m=>window.MediaRecorder&&MediaRecorder.isTypeSupported(m))||""; }
async function openRecordStream(){
  const video={facingMode:{ideal:"environment"},width:{ideal:1280},height:{ideal:720}};
  try{ return await navigator.mediaDevices.getUserMedia({video,audio:true}); }
  catch(e){ if(e.name==="NotAllowedError"||e.name==="NotFoundError"){ return navigator.mediaDevices.getUserMedia({video,audio:false}); } throw e; }
}
let tick=null, wake=null;
async function startRecording(){
  if(!window.MediaRecorder||!pickMime()){ toast("This browser can't record video. Update it, or use Chrome or Safari."); return; }
  S.view="record"; S.rec={chunks:[],started:0,result:null,interrupted:false}; render();
  try{ stream=await openRecordStream(); }
  catch(e){ S.view="confirm"; S.rec=null; toast(camError(e)); render(); return; }
  const v=$("#cam"); v.srcObject=stream; await v.play().catch(()=>{});
  const mime=pickMime(); let rec;
  try{ rec=new MediaRecorder(stream,{mimeType:mime,videoBitsPerSecond:2500000}); }catch(e){ rec=new MediaRecorder(stream); }
  S.rec.recorder=rec; S.rec.mime=rec.mimeType||mime;
  rec.ondataavailable=e=>{ if(e.data&&e.data.size) S.rec.chunks.push(e.data); };
  rec.onstop=finishRecording;
  stream.getVideoTracks()[0].onended=()=>{ if(S.rec&&!S.rec.result&&rec.state!=="inactive"){ S.rec.interrupted=true; rec.stop(); } };
  rec.start(3000); S.rec.started=Date.now();
  try{ if(navigator.wakeLock) wake=await navigator.wakeLock.request("screen"); }catch(e){}
  tick=setInterval(()=>{ const s=Math.floor((Date.now()-S.rec.started)/1000), t=$("#timer"); if(t) t.textContent=Math.floor(s/60)+":"+String(s%60).padStart(2,"0"); document.querySelectorAll("[data-result]").forEach(b=>{ b.disabled=s<MIN_REC_SECONDS; }); },500);
}
function stopRecording(result){ const r=S.rec; if(!r||!r.recorder||r.recorder.state==="inactive") return; r.result=result; r.recorder.stop(); }
function cancelRecording(){ if(!confirm("Discard this video?")) return; const r=S.rec; S.rec=null; cleanRec(); if(r&&r.recorder&&r.recorder.state!=="inactive"){ r.recorder.onstop=null; r.recorder.stop(); } S.view="confirm"; render(); }
function cleanRec(){ clearInterval(tick); tick=null; stopStream(); if(wake){ wake.release().catch(()=>{}); wake=null; } }
function finishRecording(){
  const r=S.rec; if(!r) return; cleanRec();
  r.seconds=Math.round((Date.now()-r.started)/1000); r.blob=new Blob(r.chunks,{type:(r.mime||"video/webm").split(";")[0]}); r.chunks=null; r.ext=(r.mime||"").includes("mp4")?"mp4":"webm";
  const entries=S.lines.length?S.lines:[{l:"",sku:"",size:"",o:""}];
  r.items=entries.map(x=>({l:x.l,o:x.o,sku:x.sku,size:x.size,result:r.result||"",issue:""}));
  if(r.result==="ok"&&r.items.length===1){ return saveResult(); }
  S.view="items"; render();
}
async function saveResult(){
  const r=S.rec, overall=r.items.some(i=>i.result==="wrong")?"wrong":"ok", at=new Date().toISOString();
  const record={v:1,barcode:S.barcode,result:overall,by:S.who,at,seconds:r.seconds,mime:r.mime,interrupted:!!r.interrupted,device:navigator.userAgent.slice(0,120),items:r.items.map(i=>({line:i.l,order:i.o,sku:i.sku,result:i.result,issue:i.result==="wrong"?i.issue:""}))};
  const item={id:at+"-"+Math.random().toString(36).slice(2,8),barcode:S.barcode,ext:r.ext,blob:r.blob,at,record,file:null,offset:0};
  try{ await qPut(item); }
  catch(e){ toast("Could not save the video on this device. Free up storage and record again."); return; }
  S.rec=null; S.view="home"; toast(overall==="ok"?"Saved as received OK. Uploading now.":"Saved as a wrong return. Uploading now."); await refreshQueue(); render(); pump();
}

/* ---------- screens ---------- */
let toastT=null;
function toast(m){ let t=$(".toast"); if(!t){ t=document.createElement("div"); t.className="toast"; t.setAttribute("role","status"); document.body.appendChild(t); } t.textContent=m; clearTimeout(toastT); toastT=setTimeout(()=>t.remove(),5000); }
function notesHtml(n){ return n.map(([k,t])=>`<div class="note ${k}">${esc(t)}</div>`).join(""); }
function queueHtml(){
  if(!S.queue.length) return "";
  return `<h2>Waiting to upload: keep this page open</h2><div class="card">${S.queue.map(i=>`<div class="q"><span class="mono">${esc(i.barcode)}</span><span class="muted">${i.error?esc(i.error):i.pct?i.pct+"%":"In line"}</span></div>`).join("")}</div>`;
}
async function paintQueue(){ const q=$("#queue"); if(q) q.innerHTML=queueHtml(); }
function render(){
  const v=S.view;
  if(v==="scan"){ app.innerHTML=`<div class="cam"><video id="cam" playsinline muted autoplay></video><div class="guide"></div><div class="top"><div id="scanMsg">${esc(S.scanMsg)}</div></div><div class="bottom"><div class="row"><button data-a="scanCancel">Cancel</button><button id="torch" data-a="torch" hidden>Light</button></div></div></div>`; return; }
  if(v==="record"){ app.innerHTML=`<div class="cam"><video id="cam" playsinline muted autoplay></video><div class="top"><div class="big"><span class="dot"></span><span id="timer">0:00</span></div><div class="mono">${esc(S.barcode)}</div><div>Show the label first, then open the parcel.</div></div><div class="bottom"><div class="row" style="margin-bottom:10px"><button class="good tall" data-a="stop" data-result="ok" disabled>Received OK</button><button class="bad tall" data-a="stop" data-result="wrong" disabled>Wrong return</button></div><button class="wide" data-a="recCancel">Discard video</button></div></div>`; return; }
  if(v==="confirm"){
    app.innerHTML=`<h1>Confirm parcel</h1><div class="card"><div class="big mono">${esc(S.barcode)}</div>${S.lines.map(l=>`<p class="muted" style="margin:8px 0 0">${esc(l.sku)}${l.size?" · Size "+esc(l.size):""}<br>Order ${esc(l.o)}${l.kind?" · "+esc(l.kind):""}</p>`).join("")}</div>${notesHtml(parcelNotes())}<button class="primary wide tall" data-a="record">Start video</button><div style="height:10px"></div><button class="wide" data-a="home">Scan a different parcel</button>`; return;
  }
  if(v==="items"){
    const r=S.rec, ready=r.items.every(i=>i.result==="ok"||(i.result==="wrong"&&i.issue));
    app.innerHTML=`<h1>Result for ${esc(S.barcode)}</h1>${r.interrupted?`<div class="note warn">The recording was interrupted, so check the video covers the opening. If it doesn't, save this and record again.</div>`:""}${r.items.map((it,n)=>`<div class="card"><p>${it.sku?esc(it.sku)+(it.size?" · Size "+esc(it.size):""):"This parcel"}</p><div class="row" style="margin-bottom:${it.result==="wrong"?"10px":"0"}"><button class="${it.result==="ok"?"good":""}" data-a="set" data-n="${n}" data-v="ok" aria-pressed="${it.result==="ok"}">Received OK</button><button class="${it.result==="wrong"?"bad":""}" data-a="set" data-n="${n}" data-v="wrong" aria-pressed="${it.result==="wrong"}">Wrong</button></div>${it.result==="wrong"?`<label for="iss${n}">What's wrong?</label><select id="iss${n}" data-issue="${n}"><option value="">Choose one</option>${ISSUES.map(x=>`<option${it.issue===x?" selected":""}>${esc(x)}</option>`).join("")}</select>`:""}</div>`).join("")}<button class="primary wide tall" data-a="save"${ready?"":" disabled"}>Save</button>`; return;
  }
  const idx=S.index&&S.index.at?new Date(S.index.at):null;
  if(!TEST&&!account){ app.innerHTML=`<h1>Return Scanner</h1><div class="card"><p>Sign in with the Store Chikankari Microsoft account to save videos to OneDrive.</p><button class="primary wide" data-a="signin">Sign in</button></div>`; return; }
  app.innerHTML=`<h1>Return Scanner</h1>${S.authError?`<div class="note bad">${esc(S.authError)}. Your videos are safe on this device. <button data-a="signin" style="margin-top:8px">Sign in again</button></div>`:""}
  <div class="card"><label for="who">Who is opening parcels?</label><select id="who"><option value="">Choose your name</option>${(CFG.people||[]).map(p=>`<option${S.who===p?" selected":""}>${esc(p)}</option>`).join("")}</select></div>
  <button class="primary wide tall" data-a="scan"${S.who?"":" disabled"}>Scan a parcel</button>
  <h2>Barcode won't scan?</h2><form id="typed" class="row"><input id="code" autocapitalize="characters" autocomplete="off" spellcheck="false" placeholder="MYER1073035663" aria-label="Type the barcode number" value="${esc(S.typed)}"><button type="submit" style="flex:0 0 auto"${S.who?"":" disabled"}>Continue</button></form>
  <div id="queue">${queueHtml()}</div>
  ${S.done.length?`<h2>Recently saved</h2><div class="card">${S.done.slice(0,8).map(d=>`<div class="q"><span class="mono">${esc(d.file)}</span><span class="muted">${d.result==="ok"?"Received OK":"Wrong return"}</span></div>`).join("")}</div>`:""}
  <p class="muted" style="margin-top:14px">${idx?"Order list updated "+esc(idx.toLocaleString())+".":"Order list not loaded yet."} Saved videos reach the ledger within a few hours.</p>`;
}
document.addEventListener("click",e=>{
  const b=e.target.closest("[data-a]"); if(!b||b.disabled) return; const a=b.dataset.a;
  if(a==="signin") signIn();
  else if(a==="scan") startScan();
  else if(a==="scanCancel"){ stopStream(); S.view="home"; render(); }
  else if(a==="torch") toggleTorch();
  else if(a==="home"){ S.view="home"; render(); }
  else if(a==="record") startRecording();
  else if(a==="stop") stopRecording(b.dataset.result);
  else if(a==="recCancel") cancelRecording();
  else if(a==="set"){ const it=S.rec.items[+b.dataset.n]; it.result=b.dataset.v; if(it.result==="ok") it.issue=""; render(); }
  else if(a==="save") saveResult();
});
document.addEventListener("change",e=>{
  if(e.target.id==="who"){ S.who=e.target.value; store.set("who",S.who); render(); }
  if(e.target.dataset&&e.target.dataset.issue!=null){ S.rec.items[+e.target.dataset.issue].issue=e.target.value; render(); }
});
document.addEventListener("input",e=>{ if(e.target.id==="code") S.typed=e.target.value; });
document.addEventListener("submit",e=>{ if(e.target.id!=="typed") return; e.preventDefault(); const c=(S.typed||"").trim().toUpperCase(); if(!TYPED.test(c)){ toast("Type the number under the barcode, such as MYER1073035663."); return; } S.typed=""; chooseBarcode(c); });
window.addEventListener("beforeunload",e=>{ if(S.view==="record"){ e.preventDefault(); e.returnValue=""; } });

(async function start(){
  try{ await initAuth(); }catch(e){ app.innerHTML=`<div class="note bad">${esc(e.message||e)}</div>`; return; }
  try{ if(navigator.storage&&navigator.storage.persist) navigator.storage.persist(); await refreshQueue(); }catch(e){ toast("This browser is blocking device storage, so videos can't be kept safely. Leave private mode."); }
  render();
  if(TEST||account){ loadIndex(); pump(); }
})();
if(TEST) window.SC={S,readCanvas,MYNTRA,FLIPKART,qAll,pump,render};
})();
