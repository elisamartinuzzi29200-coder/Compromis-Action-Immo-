const sections=["Parties","Situation et désignation","Déclarations du vendeur","Copropriété","Diagnostics","Prix et jouissance","Conditions particulières","Financement","Conditions suspensives","Réalisation / notaires","Négociation / séquestre","Bordereau documents","Récapitulatif"];
let step=0,templateBytes=null,viewMode="dashboard",currentId=null;
const blankPerson=()=>({type:"physique",nom:"",prenoms:"",naissance:"",lieuNaissance:"",profession:"",situation:"Célibataire",unionDate:"",unionLieu:"",adresse:"",primo:false,societe:"",siret:"",siegeSocial:""});
const initial=()=>({vendeurs:[blankPerson()],acquereurs:[blankPerson()],adresseBien:"",typeBien:"ancien",designation:"",origineVendeur:"",origineActe:"",occupation:"libre",carrez:"",carrezDate:"",metreur:"",syndic:"",construction:"",assainissement:"",repartitionAssainissement:"",erp:false,erpDate:"",erpTech:false,erpNat:false,erpSis:false,erpMinier:false,erpSismique:false,sinistre:"",parasitaire:false,parasitaireDate:"",plomb:false,plombDate:"",plombResultat:"",amiante:false,amianteDate:"",amiantePriv:false,amianteComm:false,gaz:false,gazDate:"",electricite:false,electriciteDate:"",dpe:false,dpeDate:"",audit:false,auditDate:"",prixBien:"",prixMeubles:"",jouissance:"",autresConditions:"",fraisActe:"",honoraires:"",financementMode:"avec",deniers:"",prets:"",relais:"",empruntsCours:"",ressources:"",montantPrets:"",tauxMax:"",dureePret:"",chargesMax:"",banques:"",sansPretMention:"",conditionDuree:"",conditionDate:"",autresSuspensives:"",delaiActe:"",dateActe:"",notaire:"",notaireAssistant:"",clausePenale:"",honorairesAcq:"",mandatNo:"",mandatDate:"",sequestre:false,sequestreNom:"",sequestreMontant:"",sequestreRef:"",bordereau:{}});
let data=initial();
function loadDossiers(){try{return JSON.parse(localStorage.getItem("ai-compromis-dossiers")||"[]")}catch{return []}}
function storeDossiers(list){localStorage.setItem("ai-compromis-dossiers",JSON.stringify(list))}
function dossierTitle(d){
  const partyName=x=>x.type==="morale"?(x.societe||"").trim():(x.nom+" "+x.prenoms).trim();
  const v=(d.vendeurs||[]).map(partyName).filter(Boolean).join(" / ");
  const a=(d.acquereurs||[]).map(partyName).filter(Boolean).join(" / ");
  return [v&&"Vendeur : "+v,a&&"Acquéreur : "+a,d.adresseBien].filter(Boolean).join(" — ")||"Dossier sans nom";
}
function newDossier(){
  currentId="d_"+Date.now()+"_"+Math.random().toString(36).slice(2,8);
  data=initial();step=0;viewMode="editor";render();
}
function saveDossier(){
  estimateActeFees();
  if(!currentId)currentId="d_"+Date.now()+"_"+Math.random().toString(36).slice(2,8);
  const list=loadDossiers(),idx=list.findIndex(x=>x.id===currentId);
  const item={id:currentId,title:dossierTitle(data),updatedAt:new Date().toISOString(),data:JSON.parse(JSON.stringify(data))};
  if(idx>=0)list[idx]=item;else list.unshift(item);
  storeDossiers(list);
  $("status").textContent="Dossier sauvegardé";
}
function openDossier(id){
  const item=loadDossiers().find(x=>x.id===id);if(!item)return;
  currentId=id;data={...initial(),...item.data};step=0;viewMode="editor";render();
}
function duplicateDossier(id){
  const item=loadDossiers().find(x=>x.id===id);if(!item)return;
  currentId="d_"+Date.now()+"_"+Math.random().toString(36).slice(2,8);
  data={...initial(),...JSON.parse(JSON.stringify(item.data))};step=0;viewMode="editor";saveDossier();render();
}
function deleteDossier(id){
  if(!confirm("Supprimer définitivement ce dossier de ce navigateur ?"))return;
  storeDossiers(loadDossiers().filter(x=>x.id!==id));renderDashboard();
}
try{
  const legacy=localStorage.getItem("ai-compromis-exact");
  if(legacy&&loadDossiers().length===0){
    const old={...initial(),...JSON.parse(legacy)},id="d_migration_"+Date.now();
    storeDossiers([{id,title:dossierTitle(old),updatedAt:new Date().toISOString(),data:old}]);
    localStorage.removeItem("ai-compromis-exact");
  }
}catch{}
const $=id=>document.getElementById(id),esc=s=>String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]));

let pendingImport=null;
function importButton(type,label,kind,index){
  const parts=[type];
  if(kind!==undefined&&kind!=="")parts.push(kind);
  if(index!==undefined&&index!=="")parts.push(String(index));
  return '<button type="button" class="importBtn" data-import="'+esc(parts.join(":"))+'">Importer '+esc(label)+'</button>';
}
function normalizeDateFr(v){
  const m=String(v||"").match(/\b(\d{1,2})[\\/.\-](\d{1,2})[\\/.\-](\d{2,4})\b/);
  if(!m)return "";
  let y=m[3];if(y.length===2)y=(Number(y)>40?"19":"20")+y;
  return y.padStart(4,"0")+"-"+m[2].padStart(2,"0")+"-"+m[1].padStart(2,"0");
}
function cleanDocText(s){return String(s||"").replace(/\u00ad/g,"").replace(/[ \t]+/g," ").replace(/\n{3,}/g,"\n\n").trim()}
function lineValue(text,labelRe){
  const lines=cleanDocText(text).split(/\n+/).map(x=>x.trim()).filter(Boolean);
  for(let i=0;i<lines.length;i++){
    const m=lines[i].match(labelRe);if(!m)continue;
    const same=lines[i].slice((m.index||0)+m[0].length).replace(/^[\s:;\-]+/,"").trim();
    if(same)return same;
    if(lines[i+1])return lines[i+1];
  }
  return "";
}
function sectionValue(text,startRe,endRe,maxLen){
  const t=cleanDocText(text),m=t.match(startRe);if(!m)return "";
  const pos=(m.index||0)+m[0].length,rest=t.slice(pos);
  const e=endRe?rest.search(endRe):-1;
  return cleanDocText((e>=0?rest.slice(0,e):rest.slice(0,maxLen||5000)).trim());
}
function extractSiret(s){
  const m=String(s||"").replace(/[.\-]/g," ").match(/\b(?:SIRET\s*:?\s*)?(\d{3}\s?\d{3}\s?\d{3}\s?\d{5})\b/i);
  return m?m[1].replace(/\s/g,""):"";
}
function extractPersonFromId(text){
  const t=cleanDocText(text);
  let nom=lineValue(t,/^(?:NOM|Nom)\b\s*:?\s*/i);
  let prenoms=lineValue(t,/^(?:PR[ÉE]NOMS?|Pr[ée]nom(?:s)?)\b\s*:?\s*/i);
  let naissance=lineValue(t,/(?:N[ÉE]\(E\)? LE|DATE DE NAISSANCE|Date de naissance)\s*:?\s*/i);
  let lieu=lineValue(t,/(?:LIEU DE NAISSANCE|Lieu de naissance)\s*:?\s*/i);
  if(!naissance){const d=t.match(/\b(\d{1,2}[\\/.\-]\d{1,2}[\\/.\-]\d{4})\b/);if(d)naissance=d[1]}
  nom=nom.replace(/[<>]/g," ").replace(/\s{2,}.*/,"").trim();
  prenoms=prenoms.replace(/[<>]/g," ").replace(/\s{2,}.*/,"").trim();
  lieu=lieu.replace(/\s{2,}.*/,"").trim();
  return {nom,prenoms,naissance:normalizeDateFr(naissance),lieuNaissance:lieu};
}
function extractCompanyFromKbis(text){
  const t=cleanDocText(text);
  let societe=lineValue(t,/(?:D[ÉE]NOMINATION|Dénomination|RAISON SOCIALE)\s*:?\s*/i);
  let siege=lineValue(t,/(?:SI[ÈE]GE SOCIAL|Adresse du si[èe]ge|Adresse de l'entreprise)\s*:?\s*/i);
  const siret=extractSiret(t);
  if(!societe){const lines=t.split(/\n/).map(x=>x.trim()).filter(Boolean);societe=lines.find(x=>/^(SCI|SARL|SAS|SASU|EURL|SA|SNC)\b/i.test(x))||""}
  return {societe,siret,siegeSocial:siege};
}
function extractTitleProperty(text){
  const t=cleanDocText(text);
  let designation=sectionValue(t,/(?:D[ÉE]SIGNATION|DESIGNATION)\b\s*:?\s*/i,/(?:ORIGINE DE PROPRI[ÉE]T[ÉE]|EFFET RELATIF|PROPRI[ÉE]T[ÉE][ -]JOUISSANCE|CHARGES ET CONDITIONS|SERVITUDES|URBANISME)\b/i,12000);
  designation=designation.replace(/^\s*[:.\-]+/,"").trim();
  let date="";const dm=t.match(/(?:en date du|reçu le|acte[^\n]{0,80}?du)\s*(\d{1,2}[\\/.\-]\d{1,2}[\\/.\-]\d{4})/i);if(dm)date=dm[1];
  let notaire="";const nm=t.match(/(?:Ma[iî]tre|Me)\s+([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÖØ-öø-ÿ' \-]{2,})(?:,|\n|\s+notaire)/);if(nm)notaire=("Maître "+nm[1]).replace(/\s+/g," ").trim();
  const origine=[date&&("Acte du "+date),notaire].filter(Boolean).join(" — ");
  return {designation,origineActe:origine};
}
function extractCarrez(text){
  const t=cleanDocText(text);
  let carrez="";const m=t.match(/(?:surface|superficie)[^.\n]{0,120}?(\d+(?:[.,]\d+)?)\s*m[²2]/i)||t.match(/(\d+(?:[.,]\d+)?)\s*m[²2][^\n]{0,60}(?:loi carrez|carrez)/i);
  if(m)carrez=m[1].replace(",",".")+" m²";
  let carrezDate="";const dm=t.match(/(?:date (?:du )?(?:mesurage|diagnostic|rapport)|[ée]tabli le|r[ée]alis[ée] le)\s*:?\s*(\d{1,2}[\\/.\-]\d{1,2}[\\/.\-]\d{4})/i);if(dm)carrezDate=normalizeDateFr(dm[1]);
  const metreur=lineValue(t,/(?:Op[ée]rateur|Diagnostiqueur|Cabinet|Technicien)\s*:?\s*/i);
  return {carrez,carrezDate,metreur};
}
function extractDiagnostic(text,type){
  const t=cleanDocText(text),out={};
  const dm=t.match(/(?:date (?:du )?(?:diagnostic|rapport)|[ée]tabli le|r[ée]alis[ée] le|effectu[ée] le)\s*:?\s*(\d{1,2}[\\/.\-]\d{1,2}[\\/.\-]\d{4})/i);
  const date=dm?normalizeDateFr(dm[1]):"";
  if(type==="erp"){out.erp=true;out.erpDate=date;out.erpTech=/risques? technologiques?/i.test(t);out.erpNat=/risques? naturels?/i.test(t);out.erpMinier=/risques? miniers?/i.test(t);out.erpSismique=/sismique|sismicit[ée]/i.test(t);out.erpSis=/secteur d'information sur les sols|\bSIS\b/i.test(t)}
  if(type==="parasitaire"){out.parasitaire=true;out.parasitaireDate=date}
  if(type==="plomb"){out.plomb=true;out.plombDate=date;if(/absence.*plomb|aucune unit[ée].*plomb/i.test(t))out.plombResultat="absence";else if(/sup[ée]rieur.*seuil|classe 3/i.test(t))out.plombResultat="sup";else if(/inf[ée]rieur.*seuil|classe [12]/i.test(t))out.plombResultat="inf"}
  if(type==="amiante"){out.amiante=true;out.amianteDate=date;out.amiantePriv=/parties privatives/i.test(t);out.amianteComm=/parties communes|\bDTA\b/i.test(t)}
  if(type==="gaz"){out.gaz=true;out.gazDate=date}
  if(type==="electricite"){out.electricite=true;out.electriciteDate=date}
  if(type==="dpe"){out.dpe=true;out.dpeDate=date}
  if(type==="audit"){out.audit=true;out.auditDate=date}
  return out;
}
async function extractPdfText(bytes){
  if(!window.pdfjsLib)throw Error("Le lecteur PDF n’est pas chargé. Recharge la page.");
  const pdf=await pdfjsLib.getDocument({data:bytes}).promise,parts=[];
  for(let n=1;n<=Math.min(pdf.numPages,30);n++){
    const page=await pdf.getPage(n),tc=await page.getTextContent();
    const items=tc.items.filter(x=>x.str&&x.str.trim()).map(x=>({s:x.str.trim(),x:x.transform[4],y:x.transform[5]}));
    items.sort((a,b)=>Math.abs(b.y-a.y)>2?b.y-a.y:a.x-b.x);
    const lines=[];let cur=[],last=null;
    for(const it of items){if(last===null||Math.abs(it.y-last)<=2){cur.push(it);last=last===null?it.y:(last+it.y)/2}else{lines.push(cur.sort((a,b)=>a.x-b.x).map(z=>z.s).join(" "));cur=[it];last=it.y}}
    if(cur.length)lines.push(cur.sort((a,b)=>a.x-b.x).map(z=>z.s).join(" "));
    parts.push(lines.join("\n"));
  }
  return parts.join("\n");
}
async function readImportedDocument(file){
  if(/\.pdf$/i.test(file.name)||file.type==="application/pdf"){
    const bytes=new Uint8Array(await file.arrayBuffer()),text=await extractPdfText(bytes);
    if(text.replace(/\s/g,"").length>=100)return text;
    throw Error("Ce PDF semble scanné. Envoie plutôt une image JPG/PNG du document pour la reconnaissance.");
  }
  if(!window.Tesseract)throw Error("Le module de lecture d’image n’est pas chargé. Recharge la page.");
  $("status").textContent="Lecture du document…";
  const res=await Tesseract.recognize(file,"fra",{logger:m=>{if(m.status==="recognizing text")$("status").textContent="Lecture du document… "+Math.round((m.progress||0)*100)+"%"}});
  return res.data.text||"";
}
function applyImportResult(meta,result){
  const parts=meta.split(":"),type=parts[0],kind=parts[1],index=parts[2];
  if(type==="vendeur"||type==="acquereur"){
    const arr=type==="vendeur"?"vendeurs":"acquereurs",i=Number(index)||0,p=data[arr][i]||blankPerson();
    Object.assign(p,kind==="kbis"?{type:"morale"}:{type:"physique"},result);data[arr][i]=p;
  }else Object.assign(data,result);
  saveDossier();render();
}
async function handleDocumentImport(file,meta){
  if(!file)return;
  try{
    $("status").textContent="Analyse de "+file.name+"…";
    const text=await readImportedDocument(file),parts=meta.split(":"),type=parts[0],kind=parts[1];
    let result;
    if(type==="vendeur"||type==="acquereur")result=kind==="kbis"?extractCompanyFromKbis(text):extractPersonFromId(text);
    else if(type==="titre")result=extractTitleProperty(text);
    else if(type==="carrez")result=extractCarrez(text);
    else result=extractDiagnostic(text,type);
    const useful=Object.entries(result).filter(([k,v])=>v!==""&&v!==false&&v!=null);
    if(!useful.length){alert("Aucune information suffisamment fiable n’a été détectée.");$("status").textContent="";return}
    const preview=useful.map(([k,v])=>k+" : "+v).join("\n");
    if(confirm("Informations détectées :\n\n"+preview+"\n\nLes appliquer au dossier ?"))applyImportResult(meta,result);
    $("status").textContent="Document analysé";
  }catch(e){$("status").textContent="";alert("Impossible de lire ce document : "+e.message)}
  finally{$("documentImportInput").value=""}
}
const F=(l,k,t="text",full=false)=>`<div class="field ${full?"full":""}"><label>${l}</label><input type="${t}" data-key="${k}" value="${esc(data[k])}"></div>`;
const T=(l,k)=>`<div class="field full"><label>${l}</label><textarea data-key="${k}">${esc(data[k])}</textarea></div>`;
const C=(l,k)=>`<label class="check"><input type="checkbox" data-key="${k}" ${data[k]?"checked":""}>${l}</label>`;
function personHtml(kind){return data[kind].map((p,i)=>`<div class="box"><div class="boxhead"><strong>${kind==="vendeurs"?"Vendeur":"Acquéreur"} ${i+1}</strong>${data[kind].length>1?`<button data-remove="${kind}:${i}">Retirer</button>`:""}</div>
<div class="box"><strong>Type de partie</strong>
<label class="check"><input type="radio" name="partytype-${kind}-${i}" data-party-type="${kind}:${i}" value="physique" ${(p.type||"physique")==="physique"?"checked":""}>Personne physique</label>
<label class="check"><input type="radio" name="partytype-${kind}-${i}" data-party-type="${kind}:${i}" value="morale" ${p.type==="morale"?"checked":""}>Personne morale</label>
</div>
${p.type==="morale"?`<div class="grid">
<div class="field"><label>Société</label><input data-person="${kind}:${i}:societe" value="${esc(p.societe)}"></div>
<div class="field"><label>SIRET</label><input data-person="${kind}:${i}:siret" value="${esc(p.siret)}"></div>
<div class="field full"><label>Siège social</label><input data-person="${kind}:${i}:siegeSocial" value="${esc(p.siegeSocial)}"></div>
</div>`:`<div class="grid">
<div class="field"><label>Nom</label><input data-person="${kind}:${i}:nom" value="${esc(p.nom)}"></div><div class="field"><label>Prénom(s)</label><input data-person="${kind}:${i}:prenoms" value="${esc(p.prenoms)}"></div>
<div class="field"><label>Date de naissance</label><input type="date" data-person="${kind}:${i}:naissance" value="${esc(p.naissance)}"></div><div class="field"><label>Lieu de naissance</label><input data-person="${kind}:${i}:lieuNaissance" value="${esc(p.lieuNaissance)}"></div>
<div class="field"><label>Profession</label><input data-person="${kind}:${i}:profession" value="${esc(p.profession)}"></div><div class="field"><label>Situation maritale</label><select data-person="${kind}:${i}:situation">${["Célibataire","Marié(e)","Pacsé(e)","Divorcé(e)","Veuf/Veuve"].map(x=>`<option ${p.situation===x?"selected":""}>${x}</option>`).join("")}</select></div>
${["Marié(e)","Pacsé(e)"].includes(p.situation)?`<div class="field"><label>Date mariage / PACS</label><input type="date" data-person="${kind}:${i}:unionDate" value="${esc(p.unionDate)}"></div><div class="field"><label>Lieu mariage / PACS</label><input data-person="${kind}:${i}:unionLieu" value="${esc(p.unionLieu)}"></div>`:""}
<div class="field full"><label>Adresse</label><input data-person="${kind}:${i}:adresse" value="${esc(p.adresse)}"></div>${kind==="acquereurs"?`<label class="check"><input type="checkbox" data-person-check="${kind}:${i}:primo" ${p.primo?"checked":""}>Primo-accédant</label>`:""}</div>`}
</div>`).join("")+`<button data-add="${kind}">Ajouter un ${kind==="vendeurs"?"vendeur":"acquéreur"}</button>`}
function parseMoney(v){return Number(String(v||"").replace(/\\s/g,"").replace(/€/g,"").replace(",",".").replace(/[^0-9.-]/g,""))||0}
function numberToFrench(n){n=Math.round(n);if(n===0)return "zéro";const units=["","un","deux","trois","quatre","cinq","six","sept","huit","neuf","dix","onze","douze","treize","quatorze","quinze","seize"];const under100=x=>{if(x<17)return units[x];if(x<20)return "dix-"+units[x-10];const tens=Math.floor(x/10),u=x%10;if(tens===7)return "soixante-"+under100(10+u);if(tens===9)return "quatre-vingt-"+under100(10+u);const names={2:"vingt",3:"trente",4:"quarante",5:"cinquante",6:"soixante",8:"quatre-vingt"};let s=names[tens]||"";if(u===1&&tens!==8)s+=" et un";else if(u)s+="-"+units[u];if(tens===8&&u===0)s+="s";return s};const under1000=x=>{if(x<100)return under100(x);const h=Math.floor(x/100),r=x%100;let s=(h===1?"cent":units[h]+" cent");if(r===0&&h>1)s+="s";return r?s+" "+under100(r):s};const parts=[];let billions=Math.floor(n/1e9);n%=1e9;let millions=Math.floor(n/1e6);n%=1e6;let thousands=Math.floor(n/1000);n%=1000;if(billions)parts.push((billions===1?"un":under1000(billions))+" milliard"+(billions>1?"s":""));if(millions)parts.push((millions===1?"un":under1000(millions))+" million"+(millions>1?"s":""));if(thousands)parts.push((thousands===1?"":under1000(thousands)+" ")+"mille");if(n)parts.push(under1000(n));return parts.join(" ")}
function isPrimoAcquereur(){
  const physiques=(data.acquereurs||[]).filter(p=>(p.type||"physique")==="physique");
  return physiques.length>0&&physiques.every(p=>p.primo===true);
}
function departmentFromAddress(){
  const m=String(data.adresseBien||"").match(/\b(\d{5})\b/);
  if(!m)return "";
  const cp=m[1];
  if(cp.startsWith("97")||cp.startsWith("98"))return cp.slice(0,3);
  return cp.slice(0,2);
}
function estimateActeFees(){
  const price=parseMoney(data.prixBien);
  if(!price){data.fraisActe="";return ""}
  const type=data.typeBien||"ancien";
  const dept=departmentFromAddress();
  const primo=isPrimoAcquereur();
  let rate;
  if(type==="neuf"){
    rate=0.025;
  }else{
    // Estimation pratique : 7 à 8 % dans l'ancien.
    // Pour le Finistère (29), on conserve une estimation autour de 7,5 % ;
    // pour les départements ayant relevé les DMTO, autour de 8 %.
    rate=(dept==="29"||primo)?0.075:0.08;
  }
  const amount=Math.round(price*rate/100)*100;
  data.fraisActe=amount.toLocaleString("fr-FR")+" €";
  return data.fraisActe;
}
function penaltyText(){const total=parseMoney(data.prixBien)+parseMoney(data.prixMeubles);const p=total*0.10;data.clausePenale=p?Math.round(p).toLocaleString("fr-FR")+" € ("+numberToFrench(p)+" euros)":"";return data.clausePenale}
function renderDashboard(){
  viewMode="dashboard";currentId=null;$("nav").innerHTML="";
  const list=loadDossiers().sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt)));
  $("content").innerHTML=`<div class="dashboardHead"><div><h2>MES DOSSIERS</h2><p class="hint">Les dossiers sont enregistrés uniquement dans ce navigateur.</p></div><button class="primary" id="dashNew">+ Nouveau dossier</button></div>
  ${list.length?`<div class="dossierList">${list.map(x=>`<article class="dossierCard"><div><strong>${esc(x.title)}</strong><small>Dernière modification : ${new Date(x.updatedAt).toLocaleString("fr-FR")}</small></div><div class="dossierActions"><button data-open="${x.id}">Ouvrir</button><button data-duplicate="${x.id}">Dupliquer</button><button class="dangerBtn" data-delete="${x.id}">Supprimer</button></div></article>`).join("")}</div>`:`<div class="emptyState"><strong>Aucun dossier enregistré</strong><p>Crée ton premier compromis. Tu pourras ensuite revenir ici pour le rouvrir, le dupliquer ou le supprimer.</p></div>`}`;
  $("prev").style.display="none";$("next").style.display="none";
  $("dashNew").onclick=newDossier;
  document.querySelectorAll("[data-open]").forEach(x=>x.onclick=()=>openDossier(x.dataset.open));
  document.querySelectorAll("[data-duplicate]").forEach(x=>x.onclick=()=>duplicateDossier(x.dataset.duplicate));
  document.querySelectorAll("[data-delete]").forEach(x=>x.onclick=()=>deleteDossier(x.dataset.delete));
}
const bordereauDocs=[
["TITRE / PROPRIÉTÉ",["Plan cadastral","Bornage / procès-verbal de bornage","Plan de division","Servitudes / conventions privées"]],
["COPROPRIÉTÉ",["Règlement de copropriété","État descriptif de division","Modificatifs au règlement de copropriété","Carnet d’entretien","Fiche synthétique de copropriété","Diagnostic technique global (DTG)","Projet de plan pluriannuel de travaux (PPPT)","Plan pluriannuel de travaux (PPT)","Procès-verbaux des 3 dernières assemblées générales","Appels de fonds","Dernier relevé de charges","Budget prévisionnel","État des impayés / fonds travaux","Pré-état daté","État daté","Coordonnées du syndic","Attestation loi Carrez"]],
["DIAGNOSTICS",["DPE","Audit énergétique","État des risques et pollutions (ERP)","Diagnostic amiante parties privatives","Dossier amiante parties communes","Constat de risque d’exposition au plomb (CREP)","Diagnostic gaz","Diagnostic électricité","Diagnostic termites","État parasitaire","Diagnostic assainissement non collectif","Contrôle de raccordement assainissement collectif","Information mérule","Diagnostic bruit / nuisances sonores aériennes"]],
["URBANISME / TRAVAUX",["Certificat d’urbanisme","Note de renseignements d’urbanisme","Déclaration préalable","Permis de construire","Permis d’aménager","Permis de démolir","Déclaration d’achèvement et de conformité","Attestation de non-contestation de conformité","Autorisation de copropriété pour travaux","Factures de travaux","Garanties décennales","Assurance dommages-ouvrage","Plans / notices techniques"]],
["LOCATION / OCCUPATION",["Bail en cours","État des lieux d’entrée","Dernière quittance de loyer","Dépôt de garantie","Congé délivré au locataire","Congé reçu du locataire","Avenants au bail","Inventaire mobilier","État locatif","Attestation d’assurance locataire"]],
["FISCALITÉ / CHARGES",["Dernier avis de taxe foncière","Factures eau","Factures électricité","Factures gaz","Factures chauffage","Justificatifs abonnements / contrats d’entretien","Factures ordures ménagères / TEOM"]],
["ÉQUIPEMENTS / ENTRETIEN",["Contrat entretien chaudière","Dernière facture entretien chaudière","Certificat ramonage","Contrat entretien pompe à chaleur","Contrat entretien piscine","Contrat entretien assainissement","Notice équipements","Garanties électroménager / équipements inclus","Factures équipements récents"]],
["ASSURANCES / SINISTRES",["Attestation assurance habitation","Déclaration de sinistre","Rapport d’expertise sinistre","Justificatif indemnisation assurance","Arrêté de catastrophe naturelle / technologique","Dossier dégât des eaux / incendie / fissures"]],
["MAISON / TERRAIN / ASSAINISSEMENT",["Plan du terrain","Plan de maison","Étude de sol","Étude géotechnique","Contrôle assainissement","Contrat vidange / entretien fosse","Autorisation puits / forage","Informations cuve fioul / gaz","Attestation conformité installation"]],
["VENTE / AGENCE / NOTAIRE",["Coordonnées du notaire vendeur","Coordonnées du notaire acquéreur","RIB pour séquestre / remboursement","Justificatif de séquestre"]],
["FINANCEMENT ACQUÉREUR",["Simulation de financement","Accord de principe bancaire","Attestation de courtier","Justificatif d’apport personnel","Plan de financement","Justificatif prêt relais"]],
["AUTRES ANNEXES",["Photographies annexées","Inventaire du mobilier","Liste des éléments inclus dans la vente","Liste des éléments exclus de la vente","Clés / badges / télécommandes - inventaire","Documents techniques divers","Correspondances utiles","Autre document"]]
];
function bordereauHtml(){return bordereauDocs.map(([cat,docs])=>`<div class="section">${cat}</div><div class="grid">${docs.map(d=>`<label class="check"><input type="checkbox" data-doc="${esc(d)}" ${data.bordereau&&data.bordereau[d]?"checked":""}>${d}</label>`).join("")}</div>`).join("")}
function render(){if(viewMode==="dashboard"){renderDashboard();return}$("prev").style.display="";$("next").style.display="";$("nav").innerHTML=sections.map((s,i)=>`<button data-step="${i}" class="${i===step?"active":""}">${i+1}. ${s}</button>`).join("");let h=`<h2>${sections[step]}</h2><p class="hint">Uniquement les zones à renseigner du modèle Word fourni. Le texte juridique du document n’est pas réécrit.</p>`;
if(step===0)h+=`<div class="section">VENDEUR(S)</div>${personHtml("vendeurs")}<div class="section">ACQUÉREUR(S)</div>${personHtml("acquereurs")}`;
if(step===1)h+=`<div class="box"><strong>Type de bien</strong><label class="check"><input type="radio" name="typeBien" value="ancien" ${data.typeBien==="ancien"?"checked":""}>Ancien</label><label class="check"><input type="radio" name="typeBien" value="neuf" ${data.typeBien==="neuf"?"checked":""}>Neuf</label></div><div class="grid">${F("Adresse du bien","adresseBien")}${T("Désignation complète du bien","designation")}${F("Le vendeur a acquis l’immeuble de","origineVendeur")}${F("Acte, Date, Notaire","origineActe")}</div>`;
if(step===2)h+=`<div class="box"><strong>État d’occupation</strong><label class="check"><input type="radio" name="occupation" value="libre" ${data.occupation==="libre"?"checked":""}>Libre de toute location, occupation, réquisition ou encombrement</label><label class="check"><input type="radio" name="occupation" value="loue" ${data.occupation==="loue"?"checked":""}>Loué selon l’état locatif annexé</label></div>`;
if(step===3)h+=`<div class="grid">${F("Superficie loi Carrez","carrez")}${F("Date du métrage","carrezDate","date")}${F("Métrage réalisé par","metreur")}${F("Syndic de copropriété","syndic")}</div>`;
if(step===4)h+=`<div class="grid">${F("Année de construction de l’immeuble","construction","number")}<div class="field"><label>Assainissement</label><select data-key="assainissement"><option value=""></option><option value="collectif_ok" ${data.assainissement==="collectif_ok"?"selected":""}>Collectif - raccordé</option><option value="collectif_ko" ${data.assainissement==="collectif_ko"?"selected":""}>Collectif - non/mal raccordé</option><option value="non_collectif" ${data.assainissement==="non_collectif"?"selected":""}>Non collectif</option></select></div>${F("Répartition du coût de raccordement","repartitionAssainissement")}${C("État des risques et pollution","erp")}${F("ERP établi le","erpDate","date")}${C("Risques technologiques","erpTech")}${C("Risques naturels","erpNat")}${C("Zone sismique","erpSismique")}${C("Risques miniers","erpMinier")}${C("Secteur d’information sur les sols","erpSis")}<div class="field"><label>Sinistre indemnisé</label><select data-key="sinistre"><option value=""></option><option value="non" ${data.sinistre==="non"?"selected":""}>Non</option><option value="oui" ${data.sinistre==="oui"?"selected":""}>Oui</option></select></div>${C("Diagnostic parasitaire / termites","parasitaire")}${F("Parasitaire établi le","parasitaireDate","date")}${C("Constat plomb","plomb")}${F("Plomb établi le","plombDate","date")}<div class="field"><label>Résultat plomb</label><select data-key="plombResultat"><option value=""></option><option value="absence" ${data.plombResultat==="absence"?"selected":""}>Absence</option><option value="sup" ${data.plombResultat==="sup"?"selected":""}>Présence supérieure aux seuils</option><option value="inf" ${data.plombResultat==="inf"?"selected":""}>Présence inférieure aux seuils</option></select></div>${C("Amiante","amiante")}${F("Amiante établi le","amianteDate","date")}${C("Amiante parties privatives","amiantePriv")}${C("Amiante parties communes","amianteComm")}${C("Gaz","gaz")}${F("Gaz établi le","gazDate","date")}${C("Électricité","electricite")}${F("Électricité établie le","electriciteDate","date")}${C("DPE","dpe")}${F("DPE établi le","dpeDate","date")}${C("Audit énergétique","audit")}${F("Audit établi le","auditDate","date")}</div>`;
if(step===5)h+=`<div class="grid">${F("Prix du bien immeuble","prixBien")}${F("Prix des éléments meubles","prixMeubles")}<div class="field"><label>Jouissance</label><select data-key="jouissance"><option value=""></option><option value="possession" ${data.jouissance==="possession"?"selected":""}>Prise de possession réelle</option><option value="loyers" ${data.jouissance==="loyers"?"selected":""}>Perception des loyers</option></select></div></div>`;
if(step===6)h+=`<div class="grid">${T("Autre(s) condition(s) particulière(s)","autresConditions")}</div>`;
if(step===7)h+=`<div class="box"><strong>Mode de financement</strong><label class="check"><input type="radio" name="financementMode" value="avec" ${data.financementMode==="avec"?"checked":""}>Avec prêt</label><label class="check"><input type="radio" name="financementMode" value="sans" ${data.financementMode==="sans"?"checked":""}>Sans prêt</label></div><div class="grid">${(()=>{const v=estimateActeFees();return `<div class="field"><label>Provision pour frais d’acte estimés automatiquement</label><input value="${esc(v)}" readonly><small>Calcul selon prix immeuble, ancien/neuf, adresse/département et statut primo-accédant.</small></div>`})()}${F("Honoraires d’agence","honoraires")}${F("Deniers personnels","deniers")}${data.financementMode==="avec"?F("Prêt(s) bancaire(s)","prets")+F("Prêt(s) relais","relais")+F("Emprunt(s) en cours","empruntsCours")+F("Ressources nettes mensuelles","ressources")+F("Montant global des prêts sollicités","montantPrets")+F("Taux d’intérêts maximum","tauxMax")+F("Durée du prêt","dureePret")+F("Charges mensuelles maximum","chargesMax")+T("Organisme(s) financier(s) sollicité(s)","banques"):T("Déclaration manuscrite - acquisition sans prêt","sansPretMention")}</div>`;
if(step===8)h+=`<div class="grid">${data.financementMode==="avec"?F("Durée de la condition suspensive de prêt","conditionDuree")+F("Date d’échéance","conditionDate","date"):""}${T("Autre(s) condition(s) suspensive(s)","autresSuspensives")}</div>`;
if(step===9){const penal=penaltyText();h+=`<div class="grid">${F("Signature de l’acte authentique dans un délai de","delaiActe")}${F("Ou à la date du","dateActe","date")}${T("Notaire désigné : Maître / ville / téléphone / mail","notaire")}${T("Notaire assistant : Maître / ville / téléphone / mail","notaireAssistant")}<div class="field"><label>Clause pénale (10 % du prix principal)</label><input value="${esc(penal)}" readonly></div></div>`}
if(step===10)h+=`<div class="grid">${F("Honoraires à la charge de l’acquéreur","honorairesAcq")}${F("Mandat n°","mandatNo")}${F("Mandat en date du","mandatDate","date")}${C("Existence d’un séquestre","sequestre")}${data.sequestre?F("Séquestre choisi","sequestreNom")+F("Montant du séquestre","sequestreMontant")+F("Référence séquestre","sequestreRef"):""}</div>`;
if(step===11)h+=`<div class="notice">Coche les documents effectivement remis avec le compromis. Le bordereau sera ajouté à la suite du document Word.</div>${bordereauHtml()}`;
if(step===12){const label=x=>x.type==="morale"?(x.societe||"").trim():(x.nom+" "+x.prenoms).trim(),s=data.vendeurs.map(label).filter(Boolean).join(", "),a=data.acquereurs.map(label).filter(Boolean).join(", ");h+=`<div class="notice">Le fichier généré est le modèle Word chargé, avec les zones vides et cases à cocher remplies. La charte, les textes de loi et la pagination restent ceux du modèle.</div><div class="preview">Vendeur(s) : ${esc(s)}\nAcquéreur(s) : ${esc(a)}\nBien : ${esc(data.adresseBien)}\nPrix : ${esc(data.prixBien)}</div>`}
$("content").innerHTML=h;bind()}
function bind(){document.querySelectorAll("[data-step]").forEach(x=>x.onclick=()=>{step=+x.dataset.step;render()});document.querySelectorAll("[data-key]").forEach(x=>{x.oninput=x.onchange=()=>data[x.dataset.key]=x.type==="checkbox"?x.checked:x.value});document.querySelectorAll("[data-person]").forEach(x=>x.onchange=x.oninput=()=>{const[k,i,f]=x.dataset.person.split(":");data[k][+i][f]=x.value;if(f==="situation")render()});document.querySelectorAll("[data-party-type]").forEach(x=>x.onchange=()=>{const[k,i]=x.dataset.partyType.split(":");data[k][+i].type=x.value;render()});document.querySelectorAll("[data-person-check]").forEach(x=>x.onchange=()=>{const[k,i,f]=x.dataset.personCheck.split(":");data[k][+i][f]=x.checked});document.querySelectorAll("[data-doc]").forEach(x=>x.onchange=()=>{data.bordereau=data.bordereau||{};data.bordereau[x.dataset.doc]=x.checked});document.querySelectorAll("[data-add]").forEach(x=>x.onclick=()=>{data[x.dataset.add].push(blankPerson());render()});document.querySelectorAll("[data-remove]").forEach(x=>x.onclick=()=>{const[k,i]=x.dataset.remove.split(":");data[k].splice(+i,1);render()});document.querySelectorAll("input[name=occupation]").forEach(x=>x.onchange=()=>data.occupation=x.value);document.querySelectorAll("input[name=typeBien]").forEach(x=>x.onchange=()=>{data.typeBien=x.value;estimateActeFees();render()});document.querySelectorAll("input[name=financementMode]").forEach(x=>x.onchange=()=>{data.financementMode=x.value;render()});const seq=document.querySelector('[data-key="sequestre"]');if(seq)seq.onchange=()=>{data.sequestre=seq.checked;render()}}
function dbOpen(){return new Promise((res,rej)=>{const q=indexedDB.open("compromis-action-immo",1);q.onupgradeneeded=()=>q.result.createObjectStore("files");q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error)})}
async function saveTemplate(bytes,name){const db=await dbOpen();return new Promise((res,rej)=>{const tx=db.transaction("files","readwrite");tx.objectStore("files").put({bytes,name},"template");tx.oncomplete=res;tx.onerror=()=>rej(tx.error)})}
async function loadTemplate(){try{const db=await dbOpen();return await new Promise((res,rej)=>{const tx=db.transaction("files","readonly"),q=tx.objectStore("files").get("template");q.onsuccess=()=>res(q.result||null);q.onerror=()=>rej(q.error)})}catch{return null}}
$("templateInput").onchange=async()=>{const f=$("templateInput").files[0];if(!f)return;templateBytes=new Uint8Array(await f.arrayBuffer());await saveTemplate(templateBytes,f.name);$("status").textContent="Modèle chargé : "+f.name};
$("templateBtn").onclick=()=>$("templateInput").click();
$("saveBtn").onclick=()=>{if(viewMode==="dashboard")return;saveDossier()};
$("dashboardBtn").onclick=()=>{viewMode="dashboard";renderDashboard()};$("newBtn").onclick=newDossier;
$("generateBtn").onclick=async()=>{try{if(viewMode==="dashboard"){alert("Ouvre d’abord un dossier.");return}if(!templateBytes){alert("Charge d’abord le fichier COMPROMIS DE VENTE(2).docx. Il sera mémorisé dans ce navigateur ensuite.");return}estimateActeFees();penaltyText();const blob=await buildCompromis(templateBytes,data),a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="COMPROMIS_DE_VENTE_REMPLI.docx";a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}catch(e){alert("Impossible de générer le document : "+e.message)}};
$("prev").onclick=()=>{if(step>0){step--;render()}};
$("next").onclick=()=>{if(step<sections.length-1){step++;render()}};
(async()=>{const t=await loadTemplate();if(t){templateBytes=t.bytes instanceof Uint8Array?t.bytes:new Uint8Array(t.bytes);$("status").textContent="Modèle Word mémorisé : "+t.name}else $("status").textContent="Charge le modèle Word une seule fois";renderDashboard()})();