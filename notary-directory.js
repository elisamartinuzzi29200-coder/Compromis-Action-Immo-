(function(){
  "use strict";

  const DIRECTORY_VERSION="2026-10-13-v1";
  const ENDPOINT="https://nqzriajxhkiujlnldayx.supabase.co/functions/v1/notary-directory";
  const APP_HEADER="compromis-notaires-v1";
  const originalRender=render;
  const runtime=new Map();

  function getRuntime(key){
    if(!runtime.has(key))runtime.set(key,{controller:null,timer:null,results:[],lastQuery:""});
    return runtime.get(key);
  }
  function savedSelection(key){
    const root=data._notaryDirectory&&typeof data._notaryDirectory==="object"?data._notaryDirectory:{};
    return root[key]&&typeof root[key]==="object"?root[key]:null;
  }
  function persistSelection(key,value){
    if(!data._notaryDirectory||typeof data._notaryDirectory!=="object")data._notaryDirectory={};
    if(value)data._notaryDirectory[key]=value;else delete data._notaryDirectory[key];
  }
  function formatNotary(item){
    const lines=[];
    if(item.name)lines.push("Maître "+item.name);
    if(item.office)lines.push("Étude : "+item.office);
    if(item.address)lines.push(item.address);
    if(item.phone)lines.push("Tél. : "+item.phone);
    if(item.email)lines.push("E-mail : "+item.email);
    return lines.join("\n");
  }
  function fieldTitle(key){return key==="notaireAssistant"?"Notaire assistant":"Notaire désigné"}
  function setStatus(card,message,type){
    const node=card.querySelector("[data-notary-status]");
    if(!node)return;
    node.textContent=message||"";
    node.className="notaryDirectoryStatus"+(type?" "+type:"");
  }
  function resultHtml(item,index){
    return `<article class="notaryResult">
      <div class="notaryResultMain">
        <strong>Maître ${esc(item.name||"")}</strong>
        ${item.office?`<span>${esc(item.office)}</span>`:""}
        ${item.address?`<small>${esc(item.address)}</small>`:""}
        <div class="notaryContactLine">
          ${item.phone?`<span>☎ ${esc(item.phone)}</span>`:""}
          ${item.email?`<span>✉ ${esc(item.email)}</span>`:""}
        </div>
      </div>
      <div class="notaryResultActions">
        ${item.profileUrl?`<a href="${esc(item.profileUrl)}" target="_blank" rel="noopener">Fiche officielle</a>`:""}
        <button type="button" class="primary" data-notary-select="${index}">Utiliser ces coordonnées</button>
      </div>
    </article>`;
  }
  function selectedHtml(item){
    if(!item)return "";
    return `<div class="notarySelected">
      <div><span>COORDONNÉES SÉLECTIONNÉES</span><strong>Maître ${esc(item.name||"")}</strong><small>${esc(item.office||item.address||"")}</small></div>
      <button type="button" data-notary-clear>Effacer la sélection</button>
    </div>`;
  }
  function renderResults(card,key,items){
    const state=getRuntime(key);state.results=items||[];
    const list=card.querySelector("[data-notary-results]");
    if(!list)return;
    if(!state.results.length){
      list.innerHTML=`<div class="notaryEmpty"><strong>Aucun résultat trouvé</strong><p>Vérifie l’orthographe du nom et la ville, ou saisis davantage de lettres.</p><a href="https://www.notaires.fr/fr/directory" target="_blank" rel="noopener">Ouvrir l’annuaire officiel</a></div>`;
      return;
    }
    list.innerHTML=state.results.map(resultHtml).join("");
    list.querySelectorAll("[data-notary-select]").forEach(button=>button.onclick=()=>selectResult(card,key,Number(button.dataset.notarySelect)));
  }
  function selectResult(card,key,index){
    const item=getRuntime(key).results[index];if(!item)return;
    const value=formatNotary(item);
    data[key]=value;
    persistSelection(key,{...item,selectedAt:new Date().toISOString()});
    const textarea=document.querySelector(`[data-key="${key}"]`);if(textarea)textarea.value=value;
    saveDossier("✓ Coordonnées du notaire ajoutées");
    setStatus(card,"Coordonnées ajoutées au dossier.","success");
    render();
  }
  async function searchDirectory(card,key,manual){
    const name=card.querySelector("[data-notary-name]").value.trim();
    const city=card.querySelector("[data-notary-city]").value.trim();
    if(name.length<2&&city.length<2){
      if(manual)setStatus(card,"Saisis au moins le nom du notaire ou une ville.","error");
      return;
    }
    const state=getRuntime(key),query=(name+"|"+city).toLocaleLowerCase("fr");
    if(!manual&&state.lastQuery===query)return;
    state.lastQuery=query;
    if(state.controller)state.controller.abort();
    state.controller=new AbortController();
    setStatus(card,"Recherche dans l’annuaire officiel…","loading");
    card.classList.add("is-loading");
    try{
      const response=await fetch(ENDPOINT,{method:"POST",signal:state.controller.signal,headers:{"Content-Type":"application/json","x-action-immo-app":APP_HEADER},body:JSON.stringify({name,city})});
      let payload={};try{payload=await response.json()}catch{}
      if(!response.ok)throw Error(payload.error||"Recherche indisponible.");
      const items=Array.isArray(payload.results)?payload.results:[];
      renderResults(card,key,items);
      setStatus(card,items.length?`${items.length} proposition${items.length>1?"s":""} trouvée${items.length>1?"s":""}.`:"Aucun notaire trouvé.",items.length?"success":"");
    }catch(error){
      if(error&&error.name==="AbortError")return;
      setStatus(card,"Impossible de consulter l’annuaire : "+(error&&error.message?error.message:error),"error");
      const list=card.querySelector("[data-notary-results]");if(list)list.innerHTML="";
    }finally{card.classList.remove("is-loading")}
  }
  function scheduleSearch(card,key){
    const state=getRuntime(key);clearTimeout(state.timer);
    const name=card.querySelector("[data-notary-name]").value.trim(),city=card.querySelector("[data-notary-city]").value.trim();
    if(name.length<3||city.length<2)return;
    state.timer=setTimeout(()=>searchDirectory(card,key,false),650);
  }
  function addDirectoryBlock(key){
    const textarea=document.querySelector(`[data-key="${key}"]`);if(!textarea)return;
    const field=textarea.closest(".field");if(!field||field.previousElementSibling&&field.previousElementSibling.matches(`[data-notary-directory="${key}"]`))return;
    const selected=savedSelection(key);
    const card=document.createElement("section");
    card.className="notaryDirectoryCard";
    card.dataset.notaryDirectory=key;
    card.innerHTML=`
      <div class="notaryDirectoryHead">
        <div><span>ANNUAIRE OFFICIEL</span><h3>${fieldTitle(key)}</h3><p>Indique le nom et la ville. Les coordonnées de l’étude seront ajoutées automatiquement dans le champ ci-dessous.</p></div>
        <a href="https://www.notaires.fr/fr/directory" target="_blank" rel="noopener">Notaires de France ↗</a>
      </div>
      ${selectedHtml(selected)}
      <div class="notarySearchGrid">
        <label>Nom du notaire<input type="text" data-notary-name value="${esc(selected&&selected.name||"")}" placeholder="Ex. MEROUR"></label>
        <label>Ville ou code postal<input type="text" data-notary-city value="${esc(selected&&selected.city||"")}" placeholder="Ex. Brest ou 29200"></label>
        <button type="button" class="primary" data-notary-search>Rechercher</button>
      </div>
      <p class="notaryDirectoryHint">Les résultats proviennent de l’annuaire officiel des Notaires de France. Vérifie la fiche proposée avant de l’utiliser.</p>
      <div class="notaryDirectoryStatus" data-notary-status aria-live="polite"></div>
      <div class="notaryResults" data-notary-results></div>`;
    field.parentNode.insertBefore(card,field);
    const nameInput=card.querySelector("[data-notary-name]"),cityInput=card.querySelector("[data-notary-city]");
    nameInput.addEventListener("input",()=>scheduleSearch(card,key));
    cityInput.addEventListener("input",()=>scheduleSearch(card,key));
    nameInput.addEventListener("keydown",event=>{if(event.key==="Enter"){event.preventDefault();searchDirectory(card,key,true)}});
    cityInput.addEventListener("keydown",event=>{if(event.key==="Enter"){event.preventDefault();searchDirectory(card,key,true)}});
    card.querySelector("[data-notary-search]").onclick=()=>searchDirectory(card,key,true);
    const clear=card.querySelector("[data-notary-clear]");if(clear)clear.onclick=()=>{persistSelection(key,null);data[key]="";saveDossier("✓ Coordonnées du notaire effacées");render()};
  }
  function addDirectoryBlocks(){
    if(viewMode!=="editor"||step!==9)return;
    addDirectoryBlock("notaire");
    addDirectoryBlock("notaireAssistant");
  }
  render=function(){const value=originalRender.apply(this,arguments);setTimeout(addDirectoryBlocks,0);return value};
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",()=>setTimeout(addDirectoryBlocks,0));else setTimeout(addDirectoryBlocks,0);
  window.__notaryDirectory={version:DIRECTORY_VERSION,search:(name,city)=>fetch(ENDPOINT,{method:"POST",headers:{"Content-Type":"application/json","x-action-immo-app":APP_HEADER},body:JSON.stringify({name,city})}).then(r=>r.json())};
})();