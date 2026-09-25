const models={
  nu_gestion:{label:"Location nue — gestion",filename:"CONTRAT DE LOCATION NU GESTION.docx"},
  nu_hors:{label:"Location nue — hors gestion",filename:"CONTRAT DE LOCATION NU HORS GESTION.docx"},
  meuble_gestion:{label:"Location meublée — gestion",filename:"CONTRAT DE LOCATION MEUBLE GESTION.docx"},
  meuble_hors:{label:"Location meublée — hors gestion",filename:"CONTRAT DE LOCATION MEUBLE HORS GESTION.docx"}
};
const $=id=>document.getElementById(id);
function currentKey(){const t=document.querySelector('input[name="type"]:checked').value,g=document.querySelector('input[name="gestion"]:checked').value;return t+"_"+g}
function dbOpen(){return new Promise((res,rej)=>{const q=indexedDB.open("action-immo-location-models",1);q.onupgradeneeded=()=>q.result.createObjectStore("files");q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error)})}
async function putFile(key,file){const db=await dbOpen(),bytes=new Uint8Array(await file.arrayBuffer());return new Promise((res,rej)=>{const tx=db.transaction("files","readwrite");tx.objectStore("files").put({name:file.name,bytes},"model_"+key);tx.oncomplete=res;tx.onerror=()=>rej(tx.error)})}
async function getFile(key){try{const db=await dbOpen();return await new Promise((res,rej)=>{const tx=db.transaction("files","readonly"),q=tx.objectStore("files").get("model_"+key);q.onsuccess=()=>res(q.result||null);q.onerror=()=>rej(q.error)})}catch{return null}}
async function render(){
  const key=currentKey(),m=models[key];$("selection").textContent="Contrat sélectionné : "+m.label;
  let html="";
  for(const [k,v] of Object.entries(models)){
    const f=await getFile(k);
    html+=`<div class="templateRow"><div><strong>${v.label}</strong><small>${f?'<span class="ready">Modèle chargé : '+f.name+'</span>':'<span class="missing">Modèle à charger</span>'}</small></div><div><button data-load="${k}">${f?'Remplacer':'Charger le modèle'}</button><input type="file" accept=".docx" id="file-${k}"></div></div>`;
  }
  $("templates").innerHTML=html;
  document.querySelectorAll("[data-load]").forEach(b=>b.onclick=()=>document.getElementById("file-"+b.dataset.load).click());
  for(const k of Object.keys(models)){const inp=document.getElementById("file-"+k);inp.onchange=async()=>{if(!inp.files[0])return;await putFile(k,inp.files[0]);$("status").textContent="Modèle enregistré localement";await render()}}
}
document.querySelectorAll('input[name="type"],input[name="gestion"]').forEach(x=>x.onchange=render);
$("downloadBtn").onclick=async()=>{const key=currentKey(),f=await getFile(key);if(!f){alert("Charge d’abord ce modèle Word dans la liste ci-dessous.");return}const blob=new Blob([f.bytes instanceof Uint8Array?f.bytes:new Uint8Array(f.bytes)],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}),a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=models[key].filename;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)};
render();