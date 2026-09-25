const te=new TextEncoder(),td=new TextDecoder();
const u16=(b,o)=>b[o]|b[o+1]<<8,u32=(b,o)=>(b[o]|b[o+1]<<8|b[o+2]<<16|b[o+3]<<24)>>>0,p16=v=>[v&255,v>>>8&255],p32=v=>[v&255,v>>>8&255,v>>>16&255,v>>>24&255];
function crc32(d){let c=0xffffffff;for(const x of d){c^=x;for(let k=0;k<8;k++)c=c>>>1^(0xedb88320&-(c&1))}return(c^0xffffffff)>>>0}
async function inflate(d){const ds=new DecompressionStream("deflate-raw");return new Uint8Array(await new Response(new Blob([d]).stream().pipeThrough(ds)).arrayBuffer())}
async function unzip(b){let e=-1;for(let i=b.length-22;i>=Math.max(0,b.length-65557);i--)if(u32(b,i)===0x06054b50){e=i;break}if(e<0)throw Error("DOCX invalide");const n=u16(b,e+10);let p=u32(b,e+16),f=[];for(let i=0;i<n;i++){if(u32(b,p)!==0x02014b50)throw Error("Archive DOCX invalide");const m=u16(b,p+10),cs=u32(b,p+20),nl=u16(b,p+28),el=u16(b,p+30),cl=u16(b,p+32),lo=u32(b,p+42),name=td.decode(b.slice(p+46,p+46+nl)),ln=u16(b,lo+26),le=u16(b,lo+28),st=lo+30+ln+le,comp=b.slice(st,st+cs);f.push({name,data:m===0?comp:m===8?await inflate(comp):(()=>{throw Error("Compression non prise en charge")})()});p+=46+nl+el+cl}return f}
function zip(files){let locals=[],centrals=[],off=0;for(const f of files){const n=te.encode(f.name),crc=crc32(f.data),lh=new Uint8Array([...p32(0x04034b50),...p16(20),...p16(0),...p16(0),...p16(0),...p16(0),...p32(crc),...p32(f.data.length),...p32(f.data.length),...p16(n.length),...p16(0)]),l=new Uint8Array(lh.length+n.length+f.data.length);l.set(lh);l.set(n,lh.length);l.set(f.data,lh.length+n.length);locals.push(l);const ch=new Uint8Array([...p32(0x02014b50),...p16(20),...p16(20),...p16(0),...p16(0),...p16(0),...p16(0),...p32(crc),...p32(f.data.length),...p32(f.data.length),...p16(n.length),...p16(0),...p16(0),...p16(0),...p16(0),...p32(0),...p32(off)]),c=new Uint8Array(ch.length+n.length);c.set(ch);c.set(n,ch.length);centrals.push(c);off+=l.length}const cs=centrals.reduce((s,x)=>s+x.length,0),end=new Uint8Array([...p32(0x06054b50),...p16(0),...p16(0),...p16(files.length),...p16(files.length),...p32(cs),...p32(off),...p16(0)]),out=new Uint8Array(off+cs+end.length);let q=0;for(const x of locals){out.set(x,q);q+=x.length}for(const x of centrals){out.set(x,q);q+=x.length}out.set(end,q);return out}
function setBox(doc,name,occ,value,prepend=false){const n="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",w="http://schemas.openxmlformats.org/wordprocessingml/2006/main",d=[...doc.getElementsByTagNameNS(n,"docPr")].filter(x=>x.getAttribute("name")===name)[occ];if(!d)return;let a=d.parentElement;while(a&&a.localName!=="anchor")a=a.parentElement;if(!a)return;const b=a.getElementsByTagNameNS(w,"txbxContent")[0];if(!b)return;let p;if(prepend){p=doc.createElementNS(w,"w:p");b.insertBefore(p,b.firstChild)}else{p=b.getElementsByTagNameNS(w,"p")[0]||doc.createElementNS(w,"w:p");if(!p.parentNode)b.appendChild(p);[...p.childNodes].forEach(x=>{if(x.localName!=="pPr")p.removeChild(x)})}const r=doc.createElementNS(w,"w:r");String(value||"").split("\n").forEach((s,i)=>{if(i)r.appendChild(doc.createElementNS(w,"w:br"));const t=doc.createElementNS(w,"w:t");t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");t.textContent=s;r.appendChild(t)});p.appendChild(r)}
function setCheck(doc,text,on,n=0,c=0){const w="http://schemas.openxmlformats.org/wordprocessingml/2006/main",p=[...doc.getElementsByTagNameNS(w,"p")].filter(x=>(x.textContent||"").includes(text))[n];if(!p)return;let i=0;for(const t of p.getElementsByTagNameNS(w,"t"))if(/[☐☒]/.test(t.textContent||"")){if(i===c){t.textContent=(t.textContent||"").replace(/[☐☒]/,on?"☒":"☐");return}i++}}
function setCell(doc,ti,ri,v){const w="http://schemas.openxmlformats.org/wordprocessingml/2006/main",t=doc.getElementsByTagNameNS(w,"tbl")[ti],r=t&&t.getElementsByTagNameNS(w,"tr")[ri],c=r&&r.getElementsByTagNameNS(w,"tc")[1];if(!c)return;let p=c.getElementsByTagNameNS(w,"p")[0]||doc.createElementNS(w,"w:p");if(!p.parentNode)c.appendChild(p);[...p.childNodes].forEach(x=>{if(x.localName!=="pPr")p.removeChild(x)});let rr=doc.createElementNS(w,"w:r"),tt=doc.createElementNS(w,"w:t");tt.textContent=v||"";rr.appendChild(tt);p.appendChild(rr)}
const partyLabel=p=>p.type==="morale"?(p.societe||"").trim():(p.nom+" "+p.prenoms).trim();
const personText=(p,isBuyer=false)=>{
  if(p.type==="morale"){
    return `SOCIÉTÉ ${p.societe||"___"}, SIRET ${p.siret||"___"}, siège social ${p.siegeSocial||"___"}.`;
  }
  return `${p.nom.toUpperCase()} ${p.prenoms}, né(e) le ${p.naissance||"___"} à ${p.lieuNaissance||"___"}, ${p.profession||"___"}, demeurant ${p.adresse||"___"}, situation : ${p.situation}${["Marié(e)","Pacsé(e)"].includes(p.situation)?` le ${p.unionDate||"___"} à ${p.unionLieu||"___"}`:""}${isBuyer&&p.primo?", primo-accédant":""}.`;
};
function removeBodyRange(doc,startText,endText){
  const w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const body=doc.getElementsByTagNameNS(w,"body")[0];
  if(!body)return;
  const nodes=[...body.childNodes];
  let start=-1,end=-1;
  for(let i=0;i<nodes.length;i++){
    const n=nodes[i];
    if(n.nodeType!==1)continue;
    const txt=(n.textContent||"").replace(/\s+/g," ").trim();
    if(start<0&&txt.includes(startText)){start=i;continue}
    if(start>=0&&txt.includes(endText)){end=i;break}
  }
  if(start<0||end<0||end<=start)return;
  for(let i=end-1;i>=start;i--)body.removeChild(nodes[i]);
}
function appendBordereau(doc,data){
  const w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const body=doc.getElementsByTagNameNS(w,"body")[0]; if(!body)return;
  const addP=(text,bold=false,size=22,center=false)=>{
    const p=doc.createElementNS(w,"w:p");
    const pPr=doc.createElementNS(w,"w:pPr");
    if(center){const jc=doc.createElementNS(w,"w:jc");jc.setAttributeNS(w,"w:val","center");pPr.appendChild(jc)}
    p.appendChild(pPr);
    const r=doc.createElementNS(w,"w:r"),rPr=doc.createElementNS(w,"w:rPr");
    if(bold)rPr.appendChild(doc.createElementNS(w,"w:b"));
    const sz=doc.createElementNS(w,"w:sz");sz.setAttributeNS(w,"w:val",String(size));rPr.appendChild(sz);r.appendChild(rPr);
    const t=doc.createElementNS(w,"w:t");t.textContent=text;r.appendChild(t);p.appendChild(r);body.insertBefore(p,body.lastElementChild);
  };
  const br=doc.createElementNS(w,"w:p"),r=doc.createElementNS(w,"w:r"),b=doc.createElementNS(w,"w:br");b.setAttributeNS(w,"w:type","page");r.appendChild(b);br.appendChild(r);body.insertBefore(br,body.lastElementChild);
  addP("BORDEREAU DE REMISE DES DOCUMENTS",true,30,true);
  addP("Documents remis à l’acquéreur / annexés au compromis",false,20,true);
  const cats=[
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
  let count=0;
  for(const [cat,docs] of cats){
    const selected=docs.filter(d=>data.bordereau&&data.bordereau[d]);
    if(!selected.length)continue;
    addP(cat,true,22,false);
    for(const d of selected){addP("☒ "+d,false,20,false);count++}
  }
  if(!count)addP("Aucun document sélectionné.",false,20,false);
}
async function buildCompromis(templateBytes,data){const files=await unzip(templateBytes),xf=files.find(f=>f.name==="word/document.xml");if(!xf)throw Error("Modèle Word incomplet");const doc=new DOMParser().parseFromString(td.decode(xf.data),"application/xml"),sn=data.vendeurs.map(partyLabel).filter(Boolean).join(" / "),an=data.acquereurs.map(partyLabel).filter(Boolean).join(" / ");
setBox(doc,"Zone de texte 218",0,sn);setBox(doc,"Zone de texte 219",0,an);setBox(doc,"Zone de texte 6",0,data.vendeurs.map(p=>personText(p,false)).join("\n")+"\nD’une part,");setBox(doc,"Zone de texte 7",0,data.acquereurs.map(p=>personText(p,true)).join("\n")+"\nD’autre part.");setBox(doc,"Zone de texte 2",0,data.adresseBien);setBox(doc,"Zone de texte 8",0,data.designation,true);setBox(doc,"Zone de texte 2",1,data.origineVendeur);setBox(doc,"Zone de texte 2",2,data.origineActe);
setCheck(doc,"Libre de toute location",data.occupation==="libre");setCheck(doc,"Loué selon l’état locatif",data.occupation==="loue");
setBox(doc,"Zone de texte 5",0,data.carrez);setBox(doc,"Zone de texte 13",0,data.carrezDate);setBox(doc,"Zone de texte 14",0,data.metreur);setBox(doc,"Zone de texte 15",0,data.syndic);setBox(doc,"Zone de texte 21",0,data.construction);
setCheck(doc,"zone d’assainissement collectif",data.assainissement==="collectif_ok"||data.assainissement==="collectif_ko");setCheck(doc,"raccordé au tout à l’égout",data.assainissement==="collectif_ok");setCheck(doc,"mal raccordé au tout à l’égout",data.assainissement==="collectif_ko");setCheck(doc,"zone d’assainissement non collectif",data.assainissement==="non_collectif");setBox(doc,"Zone de texte 19",0,data.repartitionAssainissement);
setCheck(doc,"ÉTAT DES RISQUES ET POLLUTION",data.erp);setBox(doc,"Zone de texte 22",0,data.erpDate);setCheck(doc,"risques technologiques",data.erpTech);setCheck(doc,"risques naturels prévisibles",data.erpNat);setCheck(doc,"Sismique",data.erpSismique);setCheck(doc,"risques miniers",data.erpMinier);setCheck(doc,"Secteur d'Information sur les Sols",data.erpSis);setCheck(doc,"n'a pas fait objet de sinistre",data.sinistre==="non");setCheck(doc,"a fait l'objet d'un sinistre",data.sinistre==="oui");
setCheck(doc,"DIAGNOSTIC PARASITAIRE",data.parasitaire);setBox(doc,"Zone de texte 30",0,data.parasitaireDate);setCheck(doc,"CONSTAT DE RISQUES",data.plomb);setBox(doc,"Zone de texte 31",0,data.plombDate);setCheck(doc,"A l’absence de plomb",data.plombResultat==="absence");setCheck(doc,"concentration supérieure",data.plombResultat==="sup");setCheck(doc,"concentration inférieure",data.plombResultat==="inf");setCheck(doc,"DIAGNOSTIC AMIANTE",data.amiante);setBox(doc,"Zone de texte 199",0,data.amianteDate);setCheck(doc,"Dans les parties privatives",data.amiantePriv);setCheck(doc,"Dans les parties communes",data.amianteComm);setCheck(doc,"DIAGNOSTIC GAZ",data.gaz);setBox(doc,"Zone de texte 204",0,data.gazDate);setCheck(doc,"DIAGNOSTIC ÉLECTRIQUE",data.electricite);setBox(doc,"Zone de texte 207",0,data.electriciteDate);setCheck(doc,"DIAGNOSTIC DE PERFORMANCE",data.dpe);setBox(doc,"Zone de texte 209",0,data.dpeDate);setCheck(doc,"AUDIT ÉNERGETIQUE",data.audit);setBox(doc,"Zone de texte 210",0,data.auditDate);
setBox(doc,"Zone de texte 213",0,data.prixBien);setBox(doc,"Zone de texte 214",0,data.prixMeubles);setCheck(doc,"prise de possession réelle",data.jouissance==="possession");setCheck(doc,"perception des loyers",data.jouissance==="loyers");setBox(doc,"Zone de texte 220",0,data.autresConditions);
setBox(doc,"Zone de texte 221",0,data.banques);setBox(doc,"Zone de texte 2",3,data.sansPretMention);setBox(doc,"Zone de texte 3",0,data.autresSuspensives);setBox(doc,"Zone de texte 4",0,data.delaiActe);setBox(doc,"Zone de texte 4",1,data.dateActe);setBox(doc,"Zone de texte 5",1,data.notaire);setBox(doc,"Zone de texte 5",2,data.notaireAssistant);setBox(doc,"Zone de texte 6",1,data.clausePenale);setBox(doc,"Zone de texte 7",1,data.honorairesAcq);setBox(doc,"Zone de texte 7",2,data.mandatDate);setBox(doc,"Zone de texte 7",3,data.mandatNo);
setCheck(doc,"Existence d’un séquestre",data.sequestre,0,0);setCheck(doc,"Existence d’un séquestre",!data.sequestre,0,1);setBox(doc,"Zone de texte 8",1,data.sequestre?data.sequestreNom:"");setBox(doc,"Zone de texte 8",2,data.sequestre?data.sequestreMontant:"");setBox(doc,"Zone de texte 8",3,data.sequestre?data.sequestreRef:"");
const num=x=>Number(String(x).replace(/[^0-9,.-]/g,"").replace(",","."))||0,total=num(data.prixBien)+num(data.prixMeubles)+num(data.fraisActe)+num(data.honoraires);[data.prixBien,data.prixMeubles,data.fraisActe,data.honoraires,String(total)].forEach((v,i)=>setCell(doc,0,i,v));[data.deniers,data.prets,data.relais].forEach((v,i)=>setCell(doc,1,i,v));[data.empruntsCours,data.ressources].forEach((v,i)=>setCell(doc,2,i,v));[data.montantPrets,data.tauxMax,data.dureePret,data.chargesMax].forEach((v,i)=>setCell(doc,3,i,v));[data.prixBien,data.prixMeubles,data.fraisActe,data.honoraires,String(total)].forEach((v,i)=>setCell(doc,4,i,v));[data.conditionDuree,data.conditionDate].forEach((v,i)=>setCell(doc,5,i,v));
if(data.financementMode==="avec"){
  removeBodyRange(doc,"FINANCEMENT DE L'ACQUISITION SANS","CONDITION SUSPENSIVE RELATIVE AU FINANCEMENT");
}else{
  removeBodyRange(doc,"FINANCEMENT DE L'ACQUISITION A L'AIDE","FINANCEMENT DE L'ACQUISITION SANS");
  removeBodyRange(doc,"CONDITION SUSPENSIVE RELATIVE AU FINANCEMENT","CONDITIONS SUSPENSIVES");
}
appendBordereau(doc,data);xf.data=te.encode(new XMLSerializer().serializeToString(doc));return new Blob([zip(files)],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"})}