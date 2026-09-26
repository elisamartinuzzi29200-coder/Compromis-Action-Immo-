const sections=["Parties","Situation et désignation","Déclarations du vendeur","Copropriété","Diagnostics","Prix et jouissance","Conditions particulières","Financement","Conditions suspensives","Réalisation / notaires","Négociation / séquestre","Bordereau documents","Récapitulatif"];
let step=0,sruStep=0,templateBytes=null,viewMode="dashboard",currentId=null;
const blankPerson=()=>({type:"physique",nom:"",prenoms:"",naissance:"",lieuNaissance:"",profession:"",situation:"Célibataire",unionDate:"",unionLieu:"",adresse:"",primo:false,societe:"",siret:"",siegeSocial:""});
const initial=()=>({sru:{acquereurIndex:0,civilites:{},signatureDates:{},dateNotification:"",lieu:"Brest",adresseBien:"",designationCourte:"",modeEnvoi:"LRAR",adresseRetour:"ACTION IMMOBILIÈRE - 6 rue La Bruyère - 29200 BREST"},vendeurs:[blankPerson()],acquereurs:[blankPerson()],adresseBien:"",typeBien:"ancien",designation:"",origineVendeur:"",origineActe:"",occupation:"libre",carrez:"",carrezDate:"",metreur:"",syndic:"",construction:"",assainissement:"",repartitionAssainissement:"",erp:false,erpDate:"",erpTech:false,erpNat:false,erpSis:false,erpMinier:false,erpSismique:false,sinistre:"",parasitaire:false,parasitaireDate:"",plomb:false,plombDate:"",plombResultat:"",amiante:false,amianteDate:"",amianteResultat:"",amiantePriv:false,amianteComm:false,gaz:false,gazDate:"",electricite:false,electriciteDate:"",dpe:false,dpeDate:"",audit:false,auditDate:"",prixBien:"",prixMeubles:"",jouissance:"",autresConditions:"",fraisActe:"",honoraires:"",financementMode:"avec",deniers:"",prets:"",relais:"",empruntsCours:"",ressources:"",montantPrets:"",tauxMax:"",dureePret:"",chargesMax:"",banques:"",sansPretMention:"",conditionDuree:"",conditionDate:"",autresSuspensives:"",delaiActe:"",dateActe:"",notaire:"",notaireAssistant:"",clausePenale:"",honorairesAcq:"",mandatNo:"",mandatDate:"",sequestre:false,sequestreNom:"",sequestreMontant:"",sequestreRef:"",bordereau:{}});
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
  const s=String(v||"").trim();
  let m=s.match(/\b(\d{1,2})[\\/.\-\s]+(\d{1,2})[\\/.\-\s]+(\d{2,4})\b/);
  if(m){
    let y=m[3];if(y.length===2)y=(Number(y)>40?"19":"20")+y;
    return y.padStart(4,"0")+"-"+m[2].padStart(2,"0")+"-"+m[1].padStart(2,"0");
  }
  const months={janvier:1,"février":2,fevrier:2,mars:3,avril:4,mai:5,juin:6,juillet:7,"août":8,aout:8,septembre:9,octobre:10,novembre:11,"décembre":12,decembre:12};
  m=s.toLowerCase().match(/\b(\d{1,2})\s+(janvier|f[ée]vrier|mars|avril|mai|juin|juillet|ao[uû]t|septembre|octobre|novembre|d[ée]cembre)\s+(\d{4})\b/i);
  if(m){
    const key=m[2].toLowerCase(),mo=months[key]||months[key.normalize("NFD").replace(/[\u0300-\u036f]/g,"")];
    if(mo)return m[3]+"-"+String(mo).padStart(2,"0")+"-"+m[1].padStart(2,"0");
  }
  return "";
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
  const lines=t.split(/\n+/).map(x=>x.trim()).filter(Boolean);
  const cleanMrz=x=>String(x||"").toUpperCase().replace(/[ «‹›]/g,"<").replace(/[^A-Z0-9<]/g,"");
  const mrzLines=lines.map(cleanMrz).filter(x=>x.length>=30);
  let mrz1=mrzLines.find(x=>/^P<[A-Z]{3}/.test(x))||"";
  let mrz2="";
  if(mrz1){const idx=mrzLines.indexOf(mrz1);mrz2=mrzLines.slice(idx+1).find(x=>x.length>=40)||""}

  let nom="",prenoms="",naissance="",lieu="",adresse="";
  if(mrz1){
    const m=mrz1.match(/^P<[A-Z]{3}([^<]+(?:<[^<]+)*)<<(.+)$/);
    if(m){
      nom=m[1].replace(/<+/g," ").trim();
      prenoms=m[2].replace(/<+/g," ").trim();
    }
  }
  if(mrz2){
    const z=mrz2.replace(/O/g,"0");
    const m=z.match(/^[A-Z0-9<]{9}[0-9][A-Z]{3}([0-9]{6})[0-9]/);
    if(m){
      const yy=Number(m[1].slice(0,2)),mm=m[1].slice(2,4),dd=m[1].slice(4,6);
      const currentYY=new Date().getFullYear()%100,century=yy>currentYY?"19":"20";
      naissance=dd+" "+mm+" "+century+String(yy).padStart(2,"0");
    }
  }

  const after=(re,span=2)=>{
    for(let i=0;i<lines.length;i++){
      const m=lines[i].match(re);if(!m)continue;
      const same=lines[i].slice((m.index||0)+m[0].length).replace(/^[\s:;\-()]+/,"").trim();
      if(same)return same;
      const vals=[];
      for(let j=1;j<=span&&lines[i+j];j++){
        if(/^(?:Nationalit|Sexe|Taille|Couleur|Date |Lieu |Autorit|Passeport|Domicile|Type|Code du pays)/i.test(lines[i+j]))break;
        vals.push(lines[i+j]);
      }
      return vals.join(" ").trim();
    }
    return "";
  };

  if(!nom)nom=after(/^(?:Nom\s*\/\s*Surname(?:\s*\(\d+\))?|NOM|SURNAME)\s*:?\s*/i,1);
  if(!prenoms)prenoms=after(/^(?:Pr[ée]noms?\s*\/\s*Given names?(?:\s*\(\d+\))?|PR[ÉE]NOMS?|GIVEN NAMES?)\s*:?\s*/i,1);
  if(!naissance)naissance=after(/^(?:Date de naissance\s*\/\s*Date of birth(?:\s*\(\d+\))?|DATE DE NAISSANCE|DATE OF BIRTH|N[ÉE]\(E\)? LE)\s*:?\s*/i,1);
  lieu=after(/^(?:Lieu de naissance\s*\/\s*Place of birth(?:\s*\(\d+\))?|LIEU DE NAISSANCE|PLACE OF BIRTH)\s*:?\s*/i,1);
  adresse=after(/^(?:Domicile\s*\/\s*Residence(?:\s*\(\d+\))?|DOMICILE|ADRESSE)\s*:?\s*/i,3);

  if(!naissance){
    const m=t.match(/Date de naissance\s*\/\s*Date of birth(?:\s*\(\d+\))?[^0-9]{0,40}(\d{1,2}[\s\/\.\-]+\d{1,2}[\s\/\.\-]+\d{2,4})/i);
    if(m)naissance=m[1];
  }
  if(!lieu){
    const m=t.match(/Lieu de naissance\s*\/\s*Place of birth(?:\s*\(\d+\))?\s*[:\-]?\s*([A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ýa-zà-öø-ÿ' .\-]{2,80})/i);
    if(m)lieu=m[1];
  }
  if(!adresse){
    const m=t.match(/Domicile\s*\/\s*Residence(?:\s*\(\d+\))?\s*[:\-]?\s*([^\n]{5,100}(?:\n[^\n]{2,60}){0,2})/i);
    if(m)adresse=m[1];
  }

  const cleanName=v=>String(v||"").replace(/[<>]/g," ").replace(/\s{2,}/g," ").trim();
  nom=cleanName(nom).replace(/\b(?:Nationalit|Française|Francaise).*$/i,"").trim();
  prenoms=cleanName(prenoms).replace(/\b(?:Nationalit|Sexe|Taille).*$/i,"").trim();
  lieu=String(lieu||"").replace(/\b(?:Date de d[ée]livrance|Domicile|Autorit[ée]).*$/i,"").replace(/\s{2,}/g," ").trim();
  adresse=String(adresse||"").replace(/\b(?:Date d['’]expiration|Autorit[ée]).*$/i,"").replace(/\s{2,}/g," ").trim();

  return {nom,prenoms,naissance:normalizeDateFr(naissance),lieuNaissance:lieu,adresse};
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
  const lines=t.split(/\n+/).map(x=>x.trim()).filter(Boolean);

  // Exact designation block: from DESIGNATION through the listed lots, stopping before buyer/price clauses.
  let designation=sectionValue(
    t,
    /(?:^|\n)\s*(?:D[ÉE]SIGNATION|DESIGNATION)\s*:?\s*/im,
    /(?:^|\n)\s*(?:L['’]ACQU[ÉE]REUR|PRIX|PROPRI[ÉE]T[ÉE][ -]JOUISSANCE|ORIGINE DE PROPRI[ÉE]T[ÉE]|EFFET RELATIF)\b/im,
    20000
  );
  designation=designation.replace(/^\s*[:.\-]+/,"").trim();

  // Typical notarial wording: "situé à BREST (FINISTERE) 29200 28 rue du Professeur Langevin".
  let adresseBien="";
  let m=t.match(/situ[ée](?:e)?\s+[àa]\s+([A-ZÀ-ÖØ-Ýa-zà-öø-ÿ' -]+)(?:\s*\([^)]*\))?\s+(\d{5})\s+(\d{1,4}\s+(?:rue|avenue|boulevard|place|impasse|all[ée]e|route|chemin|quai|cours|square|lotissement)\s+[^\n.;]{2,120})/i);
  if(m)adresseBien=(m[3]+" "+m[2]+" "+m[1]).replace(/\s+/g," ").trim();
  if(!adresseBien){
    m=t.match(/\b(\d{1,4}\s+(?:rue|avenue|boulevard|place|impasse|all[ée]e|route|chemin|quai|cours|square|lotissement)\s+[^\n,;]{2,100})[, ]+(\d{5})\s+([A-ZÀ-ÖØ-Ýa-zà-öø-ÿ' -]{2,80})/i);
    if(m)adresseBien=(m[1]+" "+m[2]+" "+m[3]).replace(/\s+/g," ").trim();
  }

  // Sellers in attestations: text between "Par :" and "Au profit de :".
  let origineVendeur="";
  const sellerBlock=sectionValue(t,/(?:^|\n)\s*Par\s*:\s*/im,/(?:^|\n)\s*Au profit de\s*:\s*/im,5000);
  if(sellerBlock){
    const names=[];
    const re=/(?:Monsieur|Madame|Mademoiselle|M\.|Mme)\s+([^,\n]{3,120})/gi;
    for(const nm of sellerBlock.matchAll(re)){
      let v=nm[1].replace(/\s+/g," ").trim();
      v=v.replace(/\s+(?:demeurant|né|née|profession|électronicien|agent|sans profession).*$/i,"").trim();
      if(v&&!names.includes(v))names.push(v);
    }
    origineVendeur=names.join(" / ");
  }
  if(!origineVendeur){
    const m=t.match(/(?:VENDEUR(?:S)?|LE VENDEUR|DE LA PART DE)\s*:?\s*(?:Monsieur|Madame|Mademoiselle|M\.|Mme)?\s*([^\n,]{3,100})/i);
    if(m)origineVendeur=m[1].replace(/\s+/g," ").trim();
  }

  // Date of the sale/act, including long French dates such as "4 octobre 2006".
  let date="";
  const datePatterns=[
    /(?:avoir re[çc]u|a re[çc]u|re[çc]u)\s+le\s+(\d{1,2}\s+(?:janvier|f[ée]vrier|mars|avril|mai|juin|juillet|ao[uû]t|septembre|octobre|novembre|d[ée]cembre)\s+\d{4})/i,
    /(?:acte|vente)[^\n]{0,100}?(?:en date du|le)\s+(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{4})/i,
    /(?:en date du|re[çc]u le|sign[ée] le)\s*(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{4})/i
  ];
  for(const re of datePatterns){const dm=t.match(re);if(dm){date=dm[1];break}}

  let notaire="";
  const notaryPatterns=[
    /JE SOUSSIGN[ÉE]?\s+(?:Ma[iî]tre|Me)\s+([A-ZÀ-ÖØ-Ýa-zà-öø-ÿ' \-]{2,80}?)\s+Notaire/i,
    /(?:Ma[iî]tre|Me)\s+([A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ýa-zà-öø-ÿ' \-]{2,80}?)(?:,|\n|\s+notaire\b)/i
  ];
  for(const re of notaryPatterns){const nm=t.match(re);if(nm){notaire=("Maître "+nm[1]).replace(/\s+/g," ").trim();break}}

  const normalizedActDate=normalizeDateFr(date);
  const origineActe=[normalizedActDate?("Acte du "+normalizedActDate):date?("Acte du "+date):"",notaire].filter(Boolean).join(" — ");
  return {adresseBien,designation,origineVendeur,origineActe};
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
  const lines=t.split(/\n+/).map(x=>x.trim()).filter(Boolean);

  function scoreDateCandidate(raw,context){
    const d=normalizeDateFr(raw);if(!d)return null;
    let score=0;const c=String(context||"");
    if(/date (?:du )?(?:diagnostic|rapport|contr[oô]le|visite|inspection)/i.test(c))score+=12;
    if(/[ée]tabli le|r[ée]alis[ée] le|effectu[ée] le|visit[ée] le|inspect[ée] le|date de visite|date de mission/i.test(c))score+=12;
    if(/diagnostic de performance [ée]nerg[ée]tique|\bDPE\b/i.test(c)&&type==="dpe")score+=8;
    if(/amiante/i.test(c)&&type==="amiante")score+=8;
    if(/gaz/i.test(c)&&type==="gaz")score+=8;
    if(/[ée]lectricit[ée]/i.test(c)&&type==="electricite")score+=8;
    if(/plomb|CREP/i.test(c)&&type==="plomb")score+=8;
    if(/termites?|parasitaire/i.test(c)&&type==="parasitaire")score+=8;
    if(/risques? et pollution|\bERP\b/i.test(c)&&type==="erp")score+=8;
    if(/audit [ée]nerg[ée]tique/i.test(c)&&type==="audit")score+=8;
    if(/validit[ée]|valable jusqu|expiration|fin de validit[ée]/i.test(c))score-=15;
    if(/construction|permis|facture|commande|impression/i.test(c))score-=8;
    return {d,score};
  }

  const candidates=[];
  for(let i=0;i<lines.length;i++){
    const ctx=[lines[i-1]||"",lines[i],lines[i+1]||""].join(" ");
    for(const m of ctx.matchAll(/\b(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{4})\b/g)){
      const s=scoreDateCandidate(m[1],ctx);if(s)candidates.push(s);
    }
  }
  const labeled=t.match(/(?:date (?:du )?(?:diagnostic|rapport|contr[oô]le|visite|inspection)|[ée]tabli le|r[ée]alis[ée] le|effectu[ée] le|visit[ée] le|inspect[ée] le|date de visite|date de mission)\s*:?\s*(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{4})/i);
  if(labeled)candidates.push({d:normalizeDateFr(labeled[1]),score:30});
  candidates.sort((a,b)=>b.score-a.score);
  const date=(candidates[0]&&candidates[0].score>0)?candidates[0].d:"";

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

function extractDiagnosticBundle(text){
  const t=cleanDocText(text);
  const defs=[
    ["erp",/(?:ÉTAT DES RISQUES(?: ET POLLUTIONS?)?|ETAT DES RISQUES(?: ET POLLUTIONS?)?|\bERP\b)/ig,"erpDate"],
    ["parasitaire",/(?:DIAGNOSTIC|ÉTAT|ETAT)[^\n]{0,80}(?:TERMITES?|PARASITAIRE)|\bTERMITES?\b/ig,"parasitaireDate"],
    ["plomb",/(?:CONSTAT DE RISQUE D['’]EXPOSITION AU PLOMB|\bCREP\b|DIAGNOSTIC[^\n]{0,50}PLOMB)/ig,"plombDate"],
    ["amiante",/(?:DIAGNOSTIC|REPÉRAGE|REPERAGE|ÉTAT|ETAT)[^\n]{0,80}AMIANTE|\bAMIANTE\b/ig,"amianteDate"],
    ["gaz",/(?:ÉTAT|ETAT|DIAGNOSTIC)[^\n]{0,100}(?:INSTALLATION INTÉRIEURE DE GAZ|INSTALLATION INTERIEURE DE GAZ|\bGAZ\b)/ig,"gazDate"],
    ["electricite",/(?:ÉTAT|ETAT|DIAGNOSTIC)[^\n]{0,100}(?:INSTALLATION INTÉRIEURE D['’]ÉLECTRICITÉ|INSTALLATION INTERIEURE D['’]ELECTRICITE|ÉLECTRICITÉ|ELECTRICITE)/ig,"electriciteDate"],
    ["dpe",/(?:DIAGNOSTIC DE PERFORMANCE ÉNERGÉTIQUE|DIAGNOSTIC DE PERFORMANCE ENERGETIQUE|\bDPE\b)/ig,"dpeDate"],
    ["audit",/(?:AUDIT ÉNERGÉTIQUE|AUDIT ENERGETIQUE)/ig,"auditDate"]
  ];
  const merged={},summary=[];
  for(const [type,re,dateKey] of defs){
    const matches=[...t.matchAll(re)];
    if(!matches.length)continue;
    let best=null,bestScore=-999;
    for(const m of matches){
      const start=Math.max(0,(m.index||0)-350),end=Math.min(t.length,(m.index||0)+7000);
      const seg=t.slice(start,end);
      const r=extractDiagnostic(seg,type);
      let score=0;
      if(r[dateKey])score+=50;
      if(/date (?:du )?(?:diagnostic|rapport|visite)|établi le|réalisé le|réalisée le|effectué le|date de visite|date de mission/i.test(seg))score+=15;
      if(type==="plomb"&&r.plombResultat)score+=5;
      if(type==="amiante"&&(r.amiantePriv||r.amianteComm))score+=5;
      if(score>bestScore){bestScore=score;best=r}
    }
    if(best){
      Object.assign(merged,best);
      const labels={erp:"ERP",parasitaire:"Termites / parasitaire",plomb:"Plomb",amiante:"Amiante",gaz:"Gaz",electricite:"Électricité",dpe:"DPE",audit:"Audit énergétique"};
      summary.push(labels[type]+(best[dateKey]?" — "+best[dateKey]:" — date non détectée"));
    }
  }
  return {data:merged,summary};
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
async function loadImageFile(file){
  const url=URL.createObjectURL(file);
  try{return await new Promise((res,rej)=>{const im=new Image();im.onload=()=>res(im);im.onerror=rej;im.src=url})}
  finally{setTimeout(()=>URL.revokeObjectURL(url),0)}
}
function canvasFromImage(img,crop,scale){
  crop=crop||{x:0,y:0,w:img.width,h:img.height};scale=scale||1;
  const canvas=document.createElement("canvas");
  canvas.width=Math.max(1,Math.round(crop.w*scale));canvas.height=Math.max(1,Math.round(crop.h*scale));
  const ctx=canvas.getContext("2d",{willReadFrequently:true});
  ctx.drawImage(img,crop.x,crop.y,crop.w,crop.h,0,0,canvas.width,canvas.height);
  return canvas;
}
function enhanceCanvas(canvas,strong){
  const ctx=canvas.getContext("2d",{willReadFrequently:true}),im=ctx.getImageData(0,0,canvas.width,canvas.height),d=im.data;
  for(let i=0;i<d.length;i+=4){
    const g=0.299*d[i]+0.587*d[i+1]+0.114*d[i+2];
    const v=strong?(g<150?0:g>225?255:Math.max(0,Math.min(255,(g-135)*1.8+128))):(g<115?0:g>238?255:Math.max(0,Math.min(255,(g-120)*1.35+120)));
    d[i]=d[i+1]=d[i+2]=v;
  }
  ctx.putImageData(im,0,0);return canvas;
}
async function ocrConfigured(image,label,lang,params){
  if(!window.Tesseract)throw Error("Le module de lecture d’image n’est pas chargé. Recharge la page.");
  const worker=await Tesseract.createWorker(lang||"fra+eng",1,{logger:m=>{if(m.status==="recognizing text")$("status").textContent=(label||"Lecture du document")+"… "+Math.round((m.progress||0)*100)+"%"}});
  try{
    if(params)await worker.setParameters(params);
    const r=await worker.recognize(image);return r.data.text||"";
  }finally{await worker.terminate()}
}
async function ocrIdentityImage(file){
  const img=await loadImageFile(file);
  const full=enhanceCanvas(canvasFromImage(img,null,Math.max(1.5,Math.min(2.6,2600/Math.max(img.width,img.height)))),false);
  const fullText=await ocrConfigured(full,"Lecture identité","fra+eng",{tessedit_pageseg_mode:"11"});
  const y=Math.floor(img.height*0.68),crop={x:0,y,w:img.width,h:img.height-y};
  const mrz=enhanceCanvas(canvasFromImage(img,crop,Math.max(2,Math.min(3.4,3200/img.width))),true);
  const mrzText=await ocrConfigured(mrz,"Lecture zone MRZ","eng",{tessedit_pageseg_mode:"6",tessedit_char_whitelist:"ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<"});
  return fullText+"\n"+mrzText;
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
async function ocrPdfPages(bytes,meta){
  const pdf=await pdfjsLib.getDocument({data:bytes}).promise,parts=[];
  const identity=/^(vendeur|acquereur):cni/.test(meta||""),title=(meta||"")==="titre";
  const max=identity?1:title?Math.min(pdf.numPages,3):Math.min(pdf.numPages,8);
  for(let n=1;n<=max;n++){
    const page=await pdf.getPage(n),viewport=page.getViewport({scale:title?2.9:2.4}),canvas=document.createElement("canvas"),ctx=canvas.getContext("2d",{willReadFrequently:true});
    canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);
    await page.render({canvasContext:ctx,viewport}).promise;
    parts.push(await ocrConfigured(canvas,(title?"Lecture titre page ":"Lecture PDF page ")+n,"fra+eng",{tessedit_pageseg_mode:title?"6":"11"}));
    if(identity){
      const y=Math.floor(canvas.height*0.68),mrz=document.createElement("canvas"),mctx=mrz.getContext("2d",{willReadFrequently:true});
      mrz.width=canvas.width;mrz.height=canvas.height-y;mctx.drawImage(canvas,0,y,canvas.width,canvas.height-y,0,0,mrz.width,mrz.height);enhanceCanvas(mrz,true);
      parts.push(await ocrConfigured(mrz,"Lecture zone MRZ","eng",{tessedit_pageseg_mode:"6",tessedit_char_whitelist:"ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<"}));
    }
  }
  return parts.join("\n");
}
async function readImportedDocument(files,meta){
  const list=Array.from(files||[]),parts=[],identity=/^(vendeur|acquereur):cni/.test(meta||"");
  for(let fi=0;fi<list.length;fi++){
    const file=list[fi];$("status").textContent="Lecture "+(fi+1)+"/"+list.length+"…";
    if(/\.pdf$/i.test(file.name)||file.type==="application/pdf"){
      const bytes=new Uint8Array(await file.arrayBuffer());
      let text=await extractPdfText(bytes);
      if(text.replace(/\s/g,"").length<100||identity||(meta||"")==="titre")text=await ocrPdfPages(bytes,meta);
      parts.push(text);
    }else if(identity){
      parts.push(await ocrIdentityImage(file));
    }else{
      const img=await loadImageFile(file),canvas=enhanceCanvas(canvasFromImage(img,null,Math.max(1.6,Math.min(2.8,2800/Math.max(img.width,img.height)))),false);
      parts.push(await ocrConfigured(canvas,"Lecture document","fra+eng",{tessedit_pageseg_mode:"6"}));
    }
  }
  return parts.join("\n");
}
function applyImportResult(meta,result){
  const parts=meta.split(":"),type=parts[0],kind=parts[1],index=parts[2];
  if(type==="vendeur"||type==="acquereur"){
    const arr=type==="vendeur"?"vendeurs":"acquereurs",i=Number(index)||0,p=data[arr][i]||blankPerson();
    Object.assign(p,kind==="kbis"?{type:"morale"}:{type:"physique"},result);data[arr][i]=p;
  }else Object.assign(data,result);
  saveDossier();render();
}
async function handleDocumentImport(files,meta){
  const list=Array.from(files||[]);if(!list.length)return;
  try{
    $("status").textContent="Analyse locale de "+list.map(x=>x.name).join(", ")+"…";
    const text=await readImportedDocument(list,meta),parts=meta.split(":"),type=parts[0],kind=parts[1];
    let result;
    if(type==="vendeur"||type==="acquereur")result=kind==="kbis"?extractCompanyFromKbis(text):extractPersonFromId(text);
    else if(type==="titre")result=extractTitleProperty(text);
    else if(type==="carrez")result=extractCarrez(text);
    else if(type==="diagbundle"){
      const bundle=extractDiagnosticBundle(text);result=bundle.data;
      if(bundle.summary.length)alert("Diagnostics repérés :\n\n"+bundle.summary.join("\n"));
    }else result=extractDiagnostic(text,type);

    const useful=Object.entries(result).filter(([k,v])=>v!==""&&v!==false&&v!=null);
    if(!useful.length){
      alert("Aucune information suffisamment fiable n’a été détectée.");
      $("status").textContent="";
      return;
    }

    const labels={
      nom:"Nom",prenoms:"Prénom(s)",naissance:"Date de naissance",lieuNaissance:"Lieu de naissance",adresse:"Adresse",
      societe:"Société",siret:"SIRET",siegeSocial:"Siège social",
      adresseBien:"Adresse du bien",designation:"Désignation",origineVendeur:"Acquis de",origineActe:"Date / notaire",
      carrez:"Superficie Carrez",carrezDate:"Date Carrez",metreur:"Diagnostiqueur"
    };
    const preview=useful.map(([k,v])=>(labels[k]||k)+" : "+v).join("\n");
    if(confirm("Informations détectées :\n\n"+preview+"\n\nLes appliquer au dossier ?"))applyImportResult(meta,result);
    $("status").textContent="Document analysé localement";
  }catch(e){
    $("status").textContent="";
    alert("Impossible de lire ce document : "+e.message);
  }finally{
    $("documentImportInput").value="";
  }
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
</div></div>`).join("")+`<button data-add="${kind}">Ajouter un ${kind==="vendeurs"?"vendeur":"acquéreur"}</button>`}
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
  viewMode="dashboard";currentId=null;$("generateBtn").textContent="Générer le Word";$("templateBtn").style.display="";$("nav").innerHTML="";
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
function ensureSru(){
  const defaults={acquereurIndex:0,civilites:{},signatureDates:{},dateNotification:"",lieu:"Brest",adresseBien:"",designationCourte:"",modeEnvoi:"LRAR",adresseRetour:"ACTION IMMOBILIÈRE - 6 rue La Bruyère - 29200 BREST"};
  data.sru={...defaults,...(data.sru||{})};
  if(!data.sru.civilites||typeof data.sru.civilites!=="object")data.sru.civilites={};
  if(!data.sru.signatureDates||typeof data.sru.signatureDates!=="object")data.sru.signatureDates={};
  if(!data.sru.adresseBien)data.sru.adresseBien=data.adresseBien||"";
  if(!data.sru.designationCourte)data.sru.designationCourte=data.designation||data.adresseBien||"";
  if(!data.sru.adresseRetour)data.sru.adresseRetour="ACTION IMMOBILIÈRE - 6 rue La Bruyère - 29200 BREST";
  if(!data.sru.modeEnvoi)data.sru.modeEnvoi="LRAR";
  if(data.sru.dateCompromis&&!Object.keys(data.sru.signatureDates).length){
    (data.vendeurs||[]).forEach((_,i)=>data.sru.signatureDates["vendeur-"+i]=data.sru.dateCompromis);
    (data.acquereurs||[]).forEach((_,i)=>data.sru.signatureDates["acquereur-"+i]=data.sru.dateCompromis);
  }
  return data.sru;
}
function sruPartyName(p){return p.type==="morale"?(p.societe||""):[p.prenoms,p.nom].filter(Boolean).join(" ")}
function sruSellerNames(){return (data.vendeurs||[]).map(sruPartyName).filter(Boolean).join(" / ")}
function fmtDateFr(v){if(!v)return "";const m=String(v).match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?m[3]+"/"+m[2]+"/"+m[1]:v}
function sruCivilite(s,index,buyer){
  if(buyer&&buyer.type==="morale")return "";
  return (s.civilites&&s.civilites[String(index)])||"";
}
function sruDisplayBuyer(s,index,buyer){
  const civ=sruCivilite(s,index,buyer);
  return [civ,sruPartyName(buyer)].filter(Boolean).join(" ");
}
function sruModeLabel(mode){return mode==="LRE"?"Lettre recommandée électronique (LRE)":"Lettre recommandée avec accusé de réception (LRAR)"}
function sruSignatoryRows(s){
  const rows=[];
  (data.vendeurs||[]).forEach((p,i)=>rows.push({key:"vendeur-"+i,label:"Vendeur "+(i+1),name:sruPartyName(p),date:(s.signatureDates||{})["vendeur-"+i]||""}));
  (data.acquereurs||[]).forEach((p,i)=>rows.push({key:"acquereur-"+i,label:"Acquéreur "+(i+1),name:sruPartyName(p),date:(s.signatureDates||{})["acquereur-"+i]||""}));
  return rows;
}
function renderSru(){
  ensureSru();viewMode="sru";
  const s=data.sru,buyers=data.acquereurs||[],buyerIndex=Math.min(Number(s.acquereurIndex)||0,Math.max(0,buyers.length-1)),buyer=buyers[buyerIndex]||blankPerson();
  const steps=["Destinataire","Avant-contrat","Notification","Vérification"];
  $("nav").innerHTML=steps.map((x,i)=>`<button data-sru-step="${i}" class="${i===sruStep?"active":""}">${i+1}. ${x}</button>`).join("");
  let h=`<div class="sruTop"><div><span class="sruEyebrow">ASSISTANT SRU</span><h2>Notification du délai de rétractation</h2><p class="hint">Les données déjà saisies dans le compromis sont reprises automatiquement. Vérifie seulement les éléments propres à la notification.</p></div><div class="progressPill">Étape ${sruStep+1} / 4</div></div><div class="sruProgress"><span style="width:${(sruStep+1)*25}%"></span></div>`;
  if(sruStep===0){
    h+=`<div class="sectionCard"><h3>À qui adresse-t-on cette notification ?</h3><p class="hint">Une notification est préparée pour un acquéreur à la fois.</p><div class="choiceCards">${buyers.map((p,i)=>`<label class="choiceCard ${buyerIndex===i?"selected":""}"><input type="radio" name="sruBuyer" value="${i}" ${buyerIndex===i?"checked":""}><span><strong>${esc(sruPartyName(p)||("Acquéreur "+(i+1)))}</strong><small>${esc(p.adresse||"Adresse non renseignée")}</small></span></label>`).join("")||'<div class="notice">Aucun acquéreur n’est encore renseigné dans le compromis.</div>'}</div>
    ${buyer&&buyer.type!=="morale"?`<div class="section sruMiniSection">Civilité à afficher sur la notification</div><div class="civiliteChoice">
      <label class="choiceCard ${sruCivilite(s,buyerIndex,buyer)==="M."?"selected":""}"><input type="radio" name="sruCivilite" value="M." ${sruCivilite(s,buyerIndex,buyer)==="M."?"checked":""}><span><strong>M.</strong><small>Monsieur</small></span></label>
      <label class="choiceCard ${sruCivilite(s,buyerIndex,buyer)==="Mme"?"selected":""}"><input type="radio" name="sruCivilite" value="Mme" ${sruCivilite(s,buyerIndex,buyer)==="Mme"?"checked":""}><span><strong>Mme</strong><small>Madame</small></span></label>
    </div>`:""}</div>`;
  }
  if(sruStep===1){
    const sigRows=sruSignatoryRows(s);
    h+=`<div class="sectionCard"><h3>Avant-contrat concerné</h3><div class="grid">
      <div class="field"><label>Lieu de signature</label><input data-sru="lieu" value="${esc(s.lieu)}"></div>
      <div class="field"><label>Adresse du bien</label><input data-sru="adresseBien" value="${esc(s.adresseBien||data.adresseBien)}"><small>Reprise automatiquement depuis le compromis.</small></div>
      <div class="field full"><label>Vendeur(s)</label><input value="${esc(sruSellerNames())}" readonly></div>
      <div class="field full"><label>Bien concerné / désignation</label><textarea data-sru="designationCourte">${esc(s.designationCourte||data.designation||data.adresseBien)}</textarea></div>
    </div>
    <div class="section sruMiniSection">Dates de signature des parties</div>
    <p class="hint">Renseigne la date propre à chaque signataire si les signatures n’ont pas eu lieu le même jour.</p>
    <div class="signatureDates">${sigRows.map(r=>`<div class="signatureDateRow"><div><strong>${esc(r.name||r.label)}</strong><small>${r.label}</small></div><input type="date" data-sru-sign="${r.key}" value="${esc(r.date)}"></div>`).join("")}</div>
    </div>`;
  }
  if(sruStep===2){
    h+=`<div class="sectionCard"><h3>Préparer l’envoi</h3><div class="grid">
      <div class="field"><label>Date de la notification / de l’envoi</label><input type="date" data-sru="dateNotification" value="${esc(s.dateNotification)}"></div>
      <div class="field"><label>Mode d’envoi prévu</label><select data-sru="modeEnvoi"><option value="LRAR" ${s.modeEnvoi==="LRAR"?"selected":""}>LRAR - Lettre recommandée avec AR</option><option value="LRE" ${s.modeEnvoi==="LRE"?"selected":""}>LRE - Lettre recommandée électronique</option></select></div>
      <div class="field full"><label>Adresse de retour du coupon de rétractation</label><input data-sru="adresseRetour" value="${esc(s.adresseRetour)}"></div>
    </div><div class="notice">Le mode choisi sera repris tel quel dans la notification et dans le coupon de rétractation.</div></div>`;
  }
  if(sruStep===3){
    const sig=sruSignatoryRows(s).filter(r=>r.date).map(r=>esc(r.name||r.label)+" : "+esc(fmtDateFr(r.date))).join("<br>");
    h+=`<div class="sectionCard"><h3>Vérification avant génération</h3>
      <div class="reviewGrid">
        <div><span>Acquéreur</span><strong>${esc(sruDisplayBuyer(s,buyerIndex,buyer)||"Non renseigné")}</strong></div>
        <div><span>Adresse acquéreur</span><strong>${esc(buyer.adresse||"Non renseignée")}</strong></div>
        <div><span>Mode d’envoi</span><strong>${esc(sruModeLabel(s.modeEnvoi))}</strong></div>
        <div><span>Date d’envoi</span><strong>${esc(fmtDateFr(s.dateNotification)||"Non renseignée")}</strong></div>
        <div class="wide"><span>Adresse du bien</span><strong>${esc(s.adresseBien||data.adresseBien||"Non renseignée")}</strong></div>
        <div class="wide"><span>Dates de signature</span><strong>${sig||"Non renseignées"}</strong></div>
        <div class="wide"><span>Vendeur(s)</span><strong>${esc(sruSellerNames()||"Non renseigné")}</strong></div>
      </div>
      <button type="button" class="primary bigAction" id="generateSruInline">Générer la notification SRU</button>
      <p class="hint">Le PDF généré comprend la notification, le coupon de rétractation et le tampon Action Immobilière prérempli.</p>
    </div>`;
  }
  $("content").innerHTML=h;
  $("prev").style.display="";$("next").style.display="";
  $("prev").disabled=sruStep===0;$("next").disabled=sruStep===3;
  $("generateBtn").textContent="Générer la notification SRU en PDF";
  $("templateBtn").style.display="none";
  document.querySelectorAll("[data-sru-step]").forEach(x=>x.onclick=()=>{sruStep=Number(x.dataset.sruStep);renderSru()});
  document.querySelectorAll("[data-sru]").forEach(x=>x.oninput=x.onchange=()=>{data.sru[x.dataset.sru]=x.value;saveDossier()});
  document.querySelectorAll("[data-sru-sign]").forEach(x=>x.onchange=()=>{data.sru.signatureDates[x.dataset.sruSign]=x.value;saveDossier()});
  document.querySelectorAll("input[name=sruBuyer]").forEach(x=>x.onchange=()=>{data.sru.acquereurIndex=Number(x.value);saveDossier();renderSru()});
  document.querySelectorAll("input[name=sruCivilite]").forEach(x=>x.onchange=()=>{data.sru.civilites[String(buyerIndex)]=x.value;saveDossier();renderSru()});
  const g=$("generateSruInline");if(g)g.onclick=generateSruDocument;
}
async function generateSruDocument(){
  try{
    ensureSru();
    const index=Number(data.sru.acquereurIndex)||0,buyer=(data.acquereurs||[])[index];
    if(!buyer){alert("Renseigne d’abord un acquéreur.");return}
    if(buyer.type!=="morale"&&!sruCivilite(data.sru,index,buyer)){alert("Choisis M. ou Mme pour l’acquéreur destinataire.");sruStep=0;renderSru();return}
    const dates=Object.values(data.sru.signatureDates||{}).filter(Boolean);
    if(!dates.length){alert("Renseigne au moins une date de signature de l’avant-contrat.");sruStep=1;renderSru();return}
    const pdf=await buildSruPdf(data,data.sru,buyer,index);
    const clean=(sruPartyName(buyer)||"acquereur").replace(/[^A-Za-zÀ-ÿ0-9]+/g,"_").replace(/^_+|_+$/g,"");
    pdf.save("NOTIFICATION_SRU_"+clean+".pdf");
  }catch(e){alert("Impossible de générer la notification SRU : "+e.message)}
}
function render(){if(viewMode==="dashboard"){renderDashboard();return}if(viewMode==="sru"){renderSru();return}$("generateBtn").textContent="Générer le Word";$("templateBtn").style.display="";$("prev").style.display="";$("next").style.display="";$("nav").innerHTML=sections.map((s,i)=>`<button data-step="${i}" class="${i===step?"active":""}">${i+1}. ${s}</button>`).join("");let h=`<h2>${sections[step]}</h2><p class="hint">Uniquement les zones à renseigner du modèle Word fourni. Le texte juridique du document n’est pas réécrit.</p>`;
if(step===0)h+=`<div class="section">VENDEUR(S)</div>${personHtml("vendeurs")}<div class="section">ACQUÉREUR(S)</div>${personHtml("acquereurs")}`;
if(step===1)h+=`<div class="box"><strong>Type de bien</strong><label class="check"><input type="radio" name="typeBien" value="ancien" ${data.typeBien==="ancien"?"checked":""}>Ancien</label><label class="check"><input type="radio" name="typeBien" value="neuf" ${data.typeBien==="neuf"?"checked":""}>Neuf</label></div><div class="grid">${F("Adresse du bien","adresseBien")}${T("Désignation complète du bien","designation")}${F("Le vendeur a acquis l’immeuble de","origineVendeur")}${F("Acte, Date, Notaire","origineActe")}</div>`;
if(step===2)h+=`<div class="box"><strong>État d’occupation</strong><label class="check"><input type="radio" name="occupation" value="libre" ${data.occupation==="libre"?"checked":""}>Libre de toute location, occupation, réquisition ou encombrement</label><label class="check"><input type="radio" name="occupation" value="loue" ${data.occupation==="loue"?"checked":""}>Loué selon l’état locatif annexé</label></div>`;
if(step===3)h+=`<div class="box"><strong>Diagnostic loi Carrez</strong><p class="hint">Importe le rapport de mesurage pour préremplir superficie, date et diagnostiqueur.</p>${importButton("carrez","le diagnostic Carrez")}</div><div class="grid">${F("Superficie loi Carrez","carrez")}${F("Date du métrage","carrezDate","date")}${F("Métrage réalisé par","metreur")}${F("Syndic de copropriété","syndic")}</div>`;
if(step===4)h+=`<div class="box"><strong>Importer les diagnostics</strong><p class="hint">Tu peux importer un DDT complet en une seule fois : l’application repère chaque diagnostic présent, coche la rubrique correspondante et cherche sa date. Tu peux aussi importer les rapports séparément.</p><div class="importRow">${importButton("diagbundle","le rapport complet DDT")}</div><div class="importGrid">${importButton("erp","ERP")}${importButton("parasitaire","termites / parasitaire")}${importButton("plomb","plomb")}${importButton("amiante","amiante")}${importButton("gaz","gaz")}${importButton("electricite","électricité")}${importButton("dpe","DPE")}${importButton("audit","audit énergétique")}</div></div><div class="grid">${F("Année de construction de l’immeuble","construction","number")}<div class="field"><label>Assainissement</label><select data-key="assainissement"><option value=""></option><option value="collectif_ok" ${data.assainissement==="collectif_ok"?"selected":""}>Collectif - raccordé</option><option value="collectif_ko" ${data.assainissement==="collectif_ko"?"selected":""}>Collectif - non/mal raccordé</option><option value="non_collectif" ${data.assainissement==="non_collectif"?"selected":""}>Non collectif</option></select></div>${F("Répartition du coût de raccordement","repartitionAssainissement")}${C("État des risques et pollution","erp")}${F("ERP établi le","erpDate","date")}${C("Risques technologiques","erpTech")}${C("Risques naturels","erpNat")}${C("Zone sismique","erpSismique")}${C("Risques miniers","erpMinier")}${C("Secteur d’information sur les sols","erpSis")}<div class="field"><label>Sinistre indemnisé</label><select data-key="sinistre"><option value=""></option><option value="non" ${data.sinistre==="non"?"selected":""}>Non</option><option value="oui" ${data.sinistre==="oui"?"selected":""}>Oui</option></select></div>${C("Diagnostic parasitaire / termites","parasitaire")}${F("Parasitaire établi le","parasitaireDate","date")}${C("Constat plomb","plomb")}${F("Plomb établi le","plombDate","date")}<div class="field"><label>Résultat plomb</label><select data-key="plombResultat"><option value=""></option><option value="absence" ${data.plombResultat==="absence"?"selected":""}>Absence</option><option value="sup" ${data.plombResultat==="sup"?"selected":""}>Présence supérieure aux seuils</option><option value="inf" ${data.plombResultat==="inf"?"selected":""}>Présence inférieure aux seuils</option></select></div>${C("Amiante","amiante")}${F("Amiante établi le","amianteDate","date")}<div class="field"><label>Résultat amiante</label><select data-key="amianteResultat"><option value=""></option><option value="absence" ${data.amianteResultat==="absence"?"selected":""}>Absence</option><option value="presence" ${data.amianteResultat==="presence"?"selected":""}>Présence</option></select></div>${C("Amiante parties privatives","amiantePriv")}${C("Amiante parties communes","amianteComm")}${C("Gaz","gaz")}${F("Gaz établi le","gazDate","date")}${C("Électricité","electricite")}${F("Électricité établie le","electriciteDate","date")}${C("DPE","dpe")}${F("DPE établi le","dpeDate","date")}${C("Audit énergétique","audit")}${F("Audit établi le","auditDate","date")}</div>`;
if(step===5)h+=`<div class="grid">${F("Prix du bien immeuble","prixBien")}${F("Prix des éléments meubles","prixMeubles")}<div class="field"><label>Jouissance</label><select data-key="jouissance"><option value=""></option><option value="possession" ${data.jouissance==="possession"?"selected":""}>Prise de possession réelle</option><option value="loyers" ${data.jouissance==="loyers"?"selected":""}>Perception des loyers</option></select></div></div>`;
if(step===6)h+=`<div class="grid">${T("Autre(s) condition(s) particulière(s)","autresConditions")}</div>`;
if(step===7)h+=`<div class="box"><strong>Mode de financement</strong><label class="check"><input type="radio" name="financementMode" value="avec" ${data.financementMode==="avec"?"checked":""}>Avec prêt</label><label class="check"><input type="radio" name="financementMode" value="sans" ${data.financementMode==="sans"?"checked":""}>Sans prêt</label></div><div class="grid">${(()=>{const v=estimateActeFees();return `<div class="field"><label>Provision pour frais d’acte estimés automatiquement</label><input value="${esc(v)}" readonly><small>Calcul selon prix immeuble, ancien/neuf, adresse/département et statut primo-accédant.</small></div>`})()}${F("Honoraires d’agence","honoraires")}${F("Deniers personnels","deniers")}${data.financementMode==="avec"?F("Prêt(s) bancaire(s)","prets")+F("Prêt(s) relais","relais")+F("Emprunt(s) en cours","empruntsCours")+F("Ressources nettes mensuelles","ressources")+F("Montant global des prêts sollicités","montantPrets")+F("Taux d’intérêts maximum","tauxMax")+F("Durée du prêt","dureePret")+F("Charges mensuelles maximum","chargesMax")+T("Organisme(s) financier(s) sollicité(s)","banques"):T("Déclaration manuscrite - acquisition sans prêt","sansPretMention")}</div>`;
if(step===8)h+=`<div class="grid">${data.financementMode==="avec"?F("Durée de la condition suspensive de prêt","conditionDuree")+F("Date d’échéance","conditionDate","date"):""}${T("Autre(s) condition(s) suspensive(s)","autresSuspensives")}</div>`;
if(step===9){const penal=penaltyText();h+=`<div class="grid">${F("Signature de l’acte authentique dans un délai de","delaiActe")}${F("Ou à la date du","dateActe","date")}${T("Notaire désigné : Maître / ville / téléphone / mail","notaire")}${T("Notaire assistant : Maître / ville / téléphone / mail","notaireAssistant")}<div class="field"><label>Clause pénale (10 % du prix principal)</label><input value="${esc(penal)}" readonly></div></div>`}
if(step===10)h+=`<div class="grid">${F("Honoraires à la charge de l’acquéreur","honorairesAcq")}${F("Mandat n°","mandatNo")}${F("Mandat en date du","mandatDate","date")}${C("Existence d’un séquestre","sequestre")}${data.sequestre?F("Séquestre choisi","sequestreNom")+F("Montant du séquestre","sequestreMontant")+F("Référence séquestre","sequestreRef"):""}</div>`;
if(step===11)h+=`<div class="notice">Coche les documents effectivement remis avec le compromis. Le bordereau sera ajouté à la suite du document Word.</div>${bordereauHtml()}`;
if(step===12){const label=x=>x.type==="morale"?(x.societe||"").trim():(x.nom+" "+x.prenoms).trim(),s=data.vendeurs.map(label).filter(Boolean).join(", "),a=data.acquereurs.map(label).filter(Boolean).join(", ");h+=`<div class="notice">Le fichier généré est le modèle Word chargé, avec les zones vides et cases à cocher remplies. La charte, les textes de loi et la pagination restent ceux du modèle.</div><div class="preview">Vendeur(s) : ${esc(s)}\nAcquéreur(s) : ${esc(a)}\nBien : ${esc(data.adresseBien)}\nPrix : ${esc(data.prixBien)}</div>`}
$("content").innerHTML=h;bind()}
function bind(){document.querySelectorAll("[data-import]").forEach(x=>x.onclick=()=>{pendingImport=x.dataset.import;$("documentImportInput").click()});document.querySelectorAll("[data-step]").forEach(x=>x.onclick=()=>{step=+x.dataset.step;render()});document.querySelectorAll("[data-key]").forEach(x=>{x.oninput=x.onchange=()=>data[x.dataset.key]=x.type==="checkbox"?x.checked:x.value});document.querySelectorAll("[data-person]").forEach(x=>x.onchange=x.oninput=()=>{const[k,i,f]=x.dataset.person.split(":");data[k][+i][f]=x.value;if(f==="situation")render()});document.querySelectorAll("[data-party-type]").forEach(x=>x.onchange=()=>{const[k,i]=x.dataset.partyType.split(":");data[k][+i].type=x.value;render()});document.querySelectorAll("[data-person-check]").forEach(x=>x.onchange=()=>{const[k,i,f]=x.dataset.personCheck.split(":");data[k][+i][f]=x.checked});document.querySelectorAll("[data-doc]").forEach(x=>x.onchange=()=>{data.bordereau=data.bordereau||{};data.bordereau[x.dataset.doc]=x.checked});document.querySelectorAll("[data-add]").forEach(x=>x.onclick=()=>{data[x.dataset.add].push(blankPerson());render()});document.querySelectorAll("[data-remove]").forEach(x=>x.onclick=()=>{const[k,i]=x.dataset.remove.split(":");data[k].splice(+i,1);render()});document.querySelectorAll("input[name=occupation]").forEach(x=>x.onchange=()=>data.occupation=x.value);document.querySelectorAll("input[name=typeBien]").forEach(x=>x.onchange=()=>{data.typeBien=x.value;estimateActeFees();render()});document.querySelectorAll("input[name=financementMode]").forEach(x=>x.onchange=()=>{data.financementMode=x.value;render()});const seq=document.querySelector('[data-key="sequestre"]');if(seq)seq.onchange=()=>{data.sequestre=seq.checked;render()}}
function dbOpen(){return new Promise((res,rej)=>{const q=indexedDB.open("compromis-action-immo",1);q.onupgradeneeded=()=>q.result.createObjectStore("files");q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error)})}
async function saveTemplate(bytes,name){const db=await dbOpen();return new Promise((res,rej)=>{const tx=db.transaction("files","readwrite");tx.objectStore("files").put({bytes,name},"template");tx.oncomplete=res;tx.onerror=()=>rej(tx.error)})}
async function loadTemplate(){try{const db=await dbOpen();return await new Promise((res,rej)=>{const tx=db.transaction("files","readonly"),q=tx.objectStore("files").get("template");q.onsuccess=()=>res(q.result||null);q.onerror=()=>rej(q.error)})}catch{return null}}
$("templateInput").onchange=async()=>{const f=$("templateInput").files[0];if(!f)return;templateBytes=new Uint8Array(await f.arrayBuffer());await saveTemplate(templateBytes,f.name);$("status").textContent="Modèle chargé : "+f.name};
$("documentImportInput").onchange=async()=>{const files=$("documentImportInput").files;if(!files||!files.length||!pendingImport)return;const meta=pendingImport;pendingImport=null;await handleDocumentImport(files,meta)};
$("templateBtn").onclick=()=>$("templateInput").click();
$("saveBtn").onclick=()=>{if(viewMode==="dashboard")return;saveDossier()};
$("sruBtn").onclick=()=>{if(viewMode==="dashboard"||!currentId){alert("Ouvre d’abord un dossier compromis.");return}saveDossier();sruStep=0;viewMode="sru";renderSru()};
$("dashboardBtn").onclick=()=>{viewMode="dashboard";renderDashboard()};$("newBtn").onclick=newDossier;
$("generateBtn").onclick=async()=>{try{if(viewMode==="sru"){await generateSruDocument();return}if(viewMode==="dashboard"){alert("Ouvre d’abord un dossier.");return}if(!templateBytes){alert("Charge d’abord le fichier COMPROMIS DE VENTE(2).docx. Il sera mémorisé dans ce navigateur ensuite.");return}estimateActeFees();penaltyText();const blob=await buildCompromis(templateBytes,data),a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="COMPROMIS_DE_VENTE_REMPLI.docx";a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}catch(e){alert("Impossible de générer le document : "+e.message)}};
$("prev").onclick=()=>{if(viewMode==="sru"){if(sruStep>0){sruStep--;renderSru()}return}if(step>0){step--;render()}};
$("next").onclick=()=>{if(viewMode==="sru"){if(sruStep<3){sruStep++;renderSru()}return}if(step<sections.length-1){step++;render()}};
(async()=>{const t=await loadTemplate();if(t){templateBytes=t.bytes instanceof Uint8Array?t.bytes:new Uint8Array(t.bytes);$("status").textContent="Modèle Word mémorisé : "+t.name}else $("status").textContent="Charge le modèle Word une seule fois";renderDashboard()})();