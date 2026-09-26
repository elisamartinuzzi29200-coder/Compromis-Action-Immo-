const te=new TextEncoder(),td=new TextDecoder();
const u16=(b,o)=>b[o]|b[o+1]<<8,u32=(b,o)=>(b[o]|b[o+1]<<8|b[o+2]<<16|b[o+3]<<24)>>>0,p16=v=>[v&255,v>>>8&255],p32=v=>[v&255,v>>>8&255,v>>>16&255,v>>>24&255];
function crc32(d){let c=0xffffffff;for(const x of d){c^=x;for(let k=0;k<8;k++)c=c>>>1^(0xedb88320&-(c&1))}return(c^0xffffffff)>>>0}
async function inflate(d){const ds=new DecompressionStream("deflate-raw");return new Uint8Array(await new Response(new Blob([d]).stream().pipeThrough(ds)).arrayBuffer())}
async function unzip(b){let e=-1;for(let i=b.length-22;i>=Math.max(0,b.length-65557);i--)if(u32(b,i)===0x06054b50){e=i;break}if(e<0)throw Error("DOCX invalide");const n=u16(b,e+10);let p=u32(b,e+16),f=[];for(let i=0;i<n;i++){if(u32(b,p)!==0x02014b50)throw Error("Archive DOCX invalide");const m=u16(b,p+10),cs=u32(b,p+20),nl=u16(b,p+28),el=u16(b,p+30),cl=u16(b,p+32),lo=u32(b,p+42),name=td.decode(b.slice(p+46,p+46+nl)),ln=u16(b,lo+26),le=u16(b,lo+28),st=lo+30+ln+le,comp=b.slice(st,st+cs);f.push({name,data:m===0?comp:m===8?await inflate(comp):(()=>{throw Error("Compression non prise en charge")})()});p+=46+nl+el+cl}return f}
function zip(files){let locals=[],centrals=[],off=0;for(const f of files){const n=te.encode(f.name),crc=crc32(f.data),lh=new Uint8Array([...p32(0x04034b50),...p16(20),...p16(0),...p16(0),...p16(0),...p16(0),...p32(crc),...p32(f.data.length),...p32(f.data.length),...p16(n.length),...p16(0)]),l=new Uint8Array(lh.length+n.length+f.data.length);l.set(lh);l.set(n,lh.length);l.set(f.data,lh.length+n.length);locals.push(l);const ch=new Uint8Array([...p32(0x02014b50),...p16(20),...p16(20),...p16(0),...p16(0),...p16(0),...p16(0),...p32(crc),...p32(f.data.length),...p32(f.data.length),...p16(n.length),...p16(0),...p16(0),...p16(0),...p16(0),...p32(0),...p32(off)]),c=new Uint8Array(ch.length+n.length);c.set(ch);c.set(n,ch.length);centrals.push(c);off+=l.length}const cs=centrals.reduce((s,x)=>s+x.length,0),end=new Uint8Array([...p32(0x06054b50),...p16(0),...p16(0),...p16(files.length),...p16(files.length),...p32(cs),...p32(off),...p16(0)]),out=new Uint8Array(off+cs+end.length);let q=0;for(const x of locals){out.set(x,q);q+=x.length}for(const x of centrals){out.set(x,q);q+=x.length}out.set(end,q);return out}
function setBox(doc,name,occ,value,prepend=false){const n="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",w="http://schemas.openxmlformats.org/wordprocessingml/2006/main",d=[...doc.getElementsByTagNameNS(n,"docPr")].filter(x=>x.getAttribute("name")===name)[occ];if(!d)return;let a=d.parentElement;while(a&&a.localName!=="anchor")a=a.parentElement;if(!a)return;const b=a.getElementsByTagNameNS(w,"txbxContent")[0];if(!b)return;let p;if(prepend){p=doc.createElementNS(w,"w:p");b.insertBefore(p,b.firstChild)}else{p=b.getElementsByTagNameNS(w,"p")[0]||doc.createElementNS(w,"w:p");if(!p.parentNode)b.appendChild(p);[...p.childNodes].forEach(x=>{if(x.localName!=="pPr")p.removeChild(x)})}const r=doc.createElementNS(w,"w:r");String(value||"").split("\n").forEach((s,i)=>{if(i)r.appendChild(doc.createElementNS(w,"w:br"));const t=doc.createElementNS(w,"w:t");t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");t.textContent=s;r.appendChild(t)});p.appendChild(r)}
function setCheck(doc,text,on,n=0,c=0){const w="http://schemas.openxmlformats.org/wordprocessingml/2006/main",p=[...doc.getElementsByTagNameNS(w,"p")].filter(x=>(x.textContent||"").includes(text))[n];if(!p)return;let i=0;for(const t of p.getElementsByTagNameNS(w,"t"))if(/[☐☒]/.test(t.textContent||"")){if(i===c){t.textContent=(t.textContent||"").replace(/[☐☒]/,on?"☒":"☐");return}i++}}
function setAmianteResult(doc,value){
  const w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const ps=[...doc.getElementsByTagNameNS(w,"p")];
  const norm=s=>String(s||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  const anchor=ps.findIndex(p=>norm(p.textContent).includes("diagnostic amiante"));
  if(anchor<0)return;
  const range=ps.slice(anchor,Math.min(ps.length,anchor+8));
  let foundAbs=false,foundPre=false;
  for(const p of range){
    const txt=norm(p.textContent);
    const boxes=[...p.getElementsByTagNameNS(w,"t")].filter(t=>/[☐☒]/.test(t.textContent||""));
    if(!boxes.length)continue;
    const hasAbs=/\babsence\b/.test(txt),hasPre=/\bpresence\b/.test(txt);
    if(hasAbs&&hasPre&&boxes.length>=2){
      const aPos=txt.indexOf("absence"),pPos=txt.indexOf("presence");
      const order=aPos<=pPos?["absence","presence"]:["presence","absence"];
      boxes[0].textContent=(boxes[0].textContent||"").replace(/[☐☒]/,value===order[0]?"☒":"☐");
      boxes[1].textContent=(boxes[1].textContent||"").replace(/[☐☒]/,value===order[1]?"☒":"☐");
      foundAbs=foundPre=true;
      continue;
    }
    if(hasAbs&&!foundAbs){
      boxes[0].textContent=(boxes[0].textContent||"").replace(/[☐☒]/,value==="absence"?"☒":"☐");foundAbs=true;
    }
    if(hasPre&&!foundPre){
      boxes[0].textContent=(boxes[0].textContent||"").replace(/[☐☒]/,value==="presence"?"☒":"☐");foundPre=true;
    }
  }
}
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
setCheck(doc,"DIAGNOSTIC PARASITAIRE",data.parasitaire);setBox(doc,"Zone de texte 30",0,data.parasitaireDate);setCheck(doc,"CONSTAT DE RISQUES",data.plomb);setBox(doc,"Zone de texte 31",0,data.plombDate);setCheck(doc,"A l’absence de plomb",data.plombResultat==="absence");setCheck(doc,"concentration supérieure",data.plombResultat==="sup");setCheck(doc,"concentration inférieure",data.plombResultat==="inf");setCheck(doc,"DIAGNOSTIC AMIANTE",data.amiante);setBox(doc,"Zone de texte 199",0,data.amianteDate);setAmianteResult(doc,data.amianteResultat||"");setCheck(doc,"Dans les parties privatives",data.amiantePriv);setCheck(doc,"Dans les parties communes",data.amianteComm);setCheck(doc,"DIAGNOSTIC GAZ",data.gaz);setBox(doc,"Zone de texte 204",0,data.gazDate);setCheck(doc,"DIAGNOSTIC ÉLECTRIQUE",data.electricite);setBox(doc,"Zone de texte 207",0,data.electriciteDate);setCheck(doc,"DIAGNOSTIC DE PERFORMANCE",data.dpe);setBox(doc,"Zone de texte 209",0,data.dpeDate);setCheck(doc,"AUDIT ÉNERGETIQUE",data.audit);setBox(doc,"Zone de texte 210",0,data.auditDate);
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
function sruXmlEscape(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")}
function sruDateFr(v){const m=String(v||"").match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?m[3]+"/"+m[2]+"/"+m[1]:String(v||"")}
function sruName(p){return p&&p.type==="morale"?(p.societe||""):[p&&p.prenoms,p&&p.nom].filter(Boolean).join(" ")}
function sruModeLabelDoc(mode){return mode==="LRE"?"Lettre recommandée électronique (LRE)":"Lettre recommandée avec accusé de réception (LRAR)"}
function sruP(text,opt={}){
  const align=opt.align?'<w:jc w:val="'+opt.align+'"/>':'';
  const spacing='<w:spacing w:before="'+(opt.before||0)+'" w:after="'+(opt.after==null?110:opt.after)+'" w:line="'+(opt.line||275)+'" w:lineRule="auto"/>';
  const shade=opt.shade?'<w:shd w:val="clear" w:color="auto" w:fill="'+opt.shade+'"/>':'';
  const border=opt.border?'<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="6" w:color="'+opt.border+'"/></w:pBdr>':'';
  const rPr='<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>'+(opt.bold?'<w:b/>':'')+(opt.color?'<w:color w:val="'+opt.color+'"/>':'')+'<w:sz w:val="'+(opt.size||20)+'"/></w:rPr>';
  return '<w:p><w:pPr>'+align+spacing+shade+border+'</w:pPr><w:r>'+rPr+'<w:t xml:space="preserve">'+sruXmlEscape(text)+'</w:t></w:r></w:p>';
}
function sruTc(content,width,opt={}){
  const fill=opt.fill||"FFFFFF",border=opt.border||"DDE3E7";
  return '<w:tc><w:tcPr><w:tcW w:w="'+width+'" w:type="dxa"/><w:vAlign w:val="center"/><w:tcMar><w:top w:w="135" w:type="dxa"/><w:left w:w="160" w:type="dxa"/><w:bottom w:w="135" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar><w:shd w:val="clear" w:color="auto" w:fill="'+fill+'"/><w:tcBorders><w:top w:val="single" w:sz="5" w:color="'+border+'"/><w:left w:val="single" w:sz="5" w:color="'+border+'"/><w:bottom w:val="single" w:sz="5" w:color="'+border+'"/><w:right w:val="single" w:sz="5" w:color="'+border+'"/></w:tcBorders></w:tcPr>'+content+'</w:tc>';
}
function sruInfoCell(label,value,width,fill="F6F8F9"){
  return sruTc(sruP(label,{size:14,bold:true,color:"73808C",after:28})+sruP(value||"—",{size:18,bold:true,color:"344252",after:10}),width,{fill});
}
function sruTableRow(cells){return '<w:tr>'+cells.join("")+'</w:tr>'}
function sruTable(rows,width="0"){
  return '<w:tbl><w:tblPr><w:tblW w:w="'+width+'" w:type="'+(width==="0"?"auto":"dxa")+'"/><w:tblCellMar><w:top w:w="35" w:type="dxa"/><w:left w:w="35" w:type="dxa"/><w:bottom w:w="35" w:type="dxa"/><w:right w:w="35" w:type="dxa"/></w:tblCellMar></w:tblPr>'+rows.join("")+'</w:tbl>';
}
function sruLogoDrawing(){
  return '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="80"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2050000" cy="1452000"/><wp:docPr id="1" name="Logo Action Immobilière"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo-action-immo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:link="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2050000" cy="1452000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
}
async function sruLogoPng(){
  const svg='<svg xmlns="http://www.w3.org/2000/svg" width="842" height="596" viewBox="0 0 842 596"><rect width="842" height="596" fill="white"/><path d="M213 382 A220 220 0 1 1 629 382" fill="none" stroke="#0A9C9B" stroke-width="16"/><rect x="304" y="210" width="134" height="172" fill="#0A9C9B"/><rect x="493" y="255" width="46" height="127" fill="#0A9C9B"/><circle cx="304" cy="344" r="38" fill="#3F4A59"/><circle cx="516" cy="187" r="38" fill="#3F4A59"/><text x="421" y="475" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="60" font-weight="700" letter-spacing="2" fill="#3F4A59">ACTION IMMOBILIÈRE</text><text x="421" y="516" text-anchor="middle" font-family="Arial, Helvetica, sans-serif" font-size="26" letter-spacing="1" fill="#0A9C9B">VENTE • LOCATION • GESTION</text></svg>';
  const blob=new Blob([svg],{type:"image/svg+xml;charset=utf-8"}),url=URL.createObjectURL(blob),img=new Image();
  await new Promise((res,rej)=>{img.onload=res;img.onerror=()=>rej(new Error("Logo impossible à générer"));img.src=url});
  const canvas=document.createElement("canvas");canvas.width=842;canvas.height=596;const ctx=canvas.getContext("2d");ctx.drawImage(img,0,0);URL.revokeObjectURL(url);
  const png=await new Promise((res,rej)=>canvas.toBlob(b=>b?res(b):rej(new Error("Logo impossible à convertir")),"image/png"));
  return new Uint8Array(await png.arrayBuffer());
}
function sruSignatureTable(data,sru){
  const rows=[];
  (data.vendeurs||[]).forEach((p,i)=>rows.push(["Vendeur",sruName(p),(sru.signatureDates||{})["vendeur-"+i]||""]));
  (data.acquereurs||[]).forEach((p,i)=>rows.push(["Acquéreur",sruName(p),(sru.signatureDates||{})["acquereur-"+i]||""]));
  const head=sruTableRow([
    sruTc(sruP("QUALITÉ",{size:14,bold:true,color:"0A9092",align:"center",after:5}),1700,{fill:"EAF6F6",border:"CFE7E7"}),
    sruTc(sruP("SIGNATAIRE",{size:14,bold:true,color:"0A9092",align:"center",after:5}),4700,{fill:"EAF6F6",border:"CFE7E7"}),
    sruTc(sruP("DATE",{size:14,bold:true,color:"0A9092",align:"center",after:5}),2500,{fill:"EAF6F6",border:"CFE7E7"})
  ]);
  const body=rows.map(([role,name,date])=>sruTableRow([
    sruTc(sruP(role,{size:16,bold:true,color:"5D6874",after:5}),1700),
    sruTc(sruP(name||"—",{size:17,color:"344252",after:5}),4700),
    sruTc(sruP(sruDateFr(date)||"—",{size:17,bold:true,color:"344252",align:"center",after:5}),2500)
  ]));
  return sruTable([head,...body],"8900");
}
function sruStamp(){
  const inside=
    sruP("ACTION IMMOBILIÈRE",{size:18,bold:true,color:"344252",align:"center",after:10})+
    sruP("VENTE • LOCATION • GESTION",{size:12,bold:true,color:"344252",align:"center",after:12})+
    sruP("6 rue La Bruyère • 29200 BREST",{size:11,align:"center",after:5})+
    sruP("02 98 46 41 41",{size:11,bold:true,align:"center",after:6})+
    sruP("RCS BREST 851 997 437 • CPI 2901 2019 000 042 411",{size:10,align:"center",after:4})+
    sruP("Adhérent SNPI n°21695 • Garantie financière QBE Europe SA/NV",{size:10,align:"center",after:3});
  return '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders><w:top w:val="double" w:sz="12" w:color="344252"/><w:left w:val="double" w:sz="12" w:color="344252"/><w:bottom w:val="double" w:sz="12" w:color="344252"/><w:right w:val="double" w:sz="12" w:color="344252"/></w:tblBorders></w:tblPr><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/><w:vAlign w:val="center"/><w:tcMar><w:top w:w="90" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar></w:tcPr>'+inside+'</w:tc></w:tr></w:tbl>';
}
async function buildSru(data,sru,buyer,buyerIndex=0){
  const sellers=(data.vendeurs||[]).map(sruName).filter(Boolean).join(" / ");
  const civ=buyer&&buyer.type==="morale"?"":((sru.civilites||{})[String(buyerIndex)]||"");
  const buyerName=[civ,sruName(buyer)].filter(Boolean).join(" ");
  const buyerAddress=buyer&&buyer.adresse||"";
  const propertyAddress=sru.adresseBien||data.adresseBien||"";
  const bien=sru.designationCourte||data.designation||propertyAddress;
  const dateNotif=sruDateFr(sru.dateNotification);
  const mode=sruModeLabelDoc(sru.modeEnvoi);
  const retour=sru.adresseRetour||"ACTION IMMOBILIÈRE - 6 rue La Bruyère - 29200 BREST";
  const salutation=civ==="Mme"?"Madame,":civ==="M."?"Monsieur,":"Madame, Monsieur,";
  const soussigne=civ==="Mme"?"Je soussignée":civ==="M."?"Je soussigné":"Je soussigné(e)";
  const signedDates=Object.values(sru.signatureDates||{}).filter(Boolean).sort();
  const firstDate=signedDates[0]||"";
  const accent="0A9C9B",dark="344252",muted="6B7785",light="EFF8F8";

  const page1=
    sruLogoDrawing()+
    sruP("NOTIFICATION DU DÉLAI DE RÉTRACTATION",{size:29,bold:true,color:dark,align:"center",after:28})+
    sruP("Article L.271-1 du Code de la construction et de l’habitation",{size:16,color:muted,align:"center",after:160})+
    sruTable([sruTableRow([sruInfoCell("DESTINATAIRE",buyerName,4300,light),sruInfoCell("ADRESSE DU DESTINATAIRE",buyerAddress,4600,light)])],"8900")+
    sruP("AVANT-CONTRAT",{size:18,bold:true,color:accent,before:120,after:50})+
    sruTable([sruTableRow([sruInfoCell("ADRESSE DU BIEN",propertyAddress,5200),sruInfoCell("MODE DE NOTIFICATION",mode,3700)])],"8900")+
    sruP("Dates de signature des parties",{size:15,bold:true,color:muted,before:80,after:40})+
    sruSignatureTable(data,sru)+
    sruP(salutation,{size:20,bold:true,before:110,after:70})+
    sruP("Veuillez trouver sous ce pli un original du compromis de vente, accompagné le cas échéant de ses annexes, signé par les parties aux dates indiquées ci-dessus.",{size:18,after:80,line:285})+
    sruP("Vendeur(s) : "+(sellers||"—"),{size:17,bold:true,color:dark,after:65})+
    sruP("Bien concerné : "+(bien||"—"),{size:17,shade:"F7F9FA",after:95,line:280})+
    sruP("DÉLAI DE RÉTRACTATION",{size:18,bold:true,color:accent,before:50,after:45})+
    sruP("Nous vous informons que le délai de rétractation de 10 jours dont vous bénéficiez commencera à courir le lendemain de la première présentation de la présente.",{size:18,bold:true,color:dark,after:70,line:285})+
    sruP("Si vous décidez d’utiliser votre faculté de rétractation, vous devez l’exercer par lettre recommandée avec demande d’avis de réception ; vous pouvez à cette fin utiliser le coupon de rétractation figurant en page suivante.",{size:17,after:90,line:280})+
    sruP("Fait à "+(sru.lieu||"")+" le "+(dateNotif||"________________"),{size:17,after:35})+
    sruP("Pour ACTION IMMOBILIÈRE",{size:15,bold:true,color:muted,after:5});

  const page2=
    sruLogoDrawing()+
    sruP("COUPON DE RÉTRACTATION",{size:27,bold:true,color:dark,align:"center",after:24})+
    sruP("À renvoyer par lettre recommandée avec accusé de réception à l’adresse indiquée ci-dessous.",{size:16,bold:true,color:accent,align:"center",after:140})+
    sruP(soussigne+" :",{size:16,bold:true,color:muted,after:25})+
    sruP(buyerName||"—",{size:21,bold:true,color:dark,shade:light,after:100})+
    sruP("déclare exercer ma faculté de rétractation concernant le compromis de vente portant sur le bien suivant :",{size:18,after:70,line:280})+
    sruP((propertyAddress?propertyAddress+" - ":"")+(bien||"—"),{size:18,bold:true,color:dark,shade:"F7F9FA",after:100,line:280})+
    sruTable([sruTableRow([sruInfoCell("NOTIFICATION ADRESSÉE À",buyerAddress,5400),sruInfoCell("1re DATE DE SIGNATURE",sruDateFr(firstDate)||"—",3500)])],"8900")+
    sruP("Date de rétractation : ____________________",{size:17,before:95,after:45})+
    sruP("Signature de l’acquéreur :",{size:16,bold:true,color:muted,after:110})+
    sruP(" ",{size:20,after:140})+
    sruP("ADRESSE DE RETOUR",{size:16,bold:true,color:accent,after:30})+
    sruP(retour,{size:17,bold:true,shade:light,after:100,line:280})+
    sruP("TAMPON DE L’AGENCE",{size:16,bold:true,color:accent,after:35})+
    sruStamp();

  const xml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>'
    +page1+'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'+page2
    +'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="650" w:right="850" w:bottom="650" w:left="850" w:header="350" w:footer="350" w:gutter="0"/></w:sectPr></w:body></w:document>';
  const types='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  const rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const docRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://www.action-immo.com/office12/action_immo_brest/catalog/images/H1a.png" TargetMode="External"/></Relationships>';
  const files=[
    {name:"[Content_Types].xml",data:te.encode(types)},
    {name:"_rels/.rels",data:te.encode(rels)},
    {name:"word/document.xml",data:te.encode(xml)},
    {name:"word/_rels/document.xml.rels",data:te.encode(docRels)}
  ];
  return new Blob([zip(files)],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
}

const SRU_EXACT_LOGO="data:image/png;base64,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[... ELLIPSIZATION ...]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";
function sruPdfModeLabel(mode){return mode==="LRE"?"Lettre recommandee electronique (LRE)":"Lettre recommandee avec accuse de reception (LRAR)"}
function sruPdfName(p){return p&&p.type==="morale"?(p.societe||""):[p&&p.prenoms,p&&p.nom].filter(Boolean).join(" ")}
function sruPdfDate(v){const m=String(v||"").match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?m[3]+"/"+m[2]+"/"+m[1]:String(v||"")}
async function buildSruPdf(data,sru,buyer,buyerIndex=0){
  if(!window.jspdf||!window.jspdf.jsPDF)throw new Error("Le moteur PDF n'est pas charge.");
  const {jsPDF}=window.jspdf,doc=new jsPDF({unit:"mm",format:"a4",orientation:"portrait",compress:true});
  const sruLogoData="data:image/png;base64,"+btoa(String.fromCharCode(...await sruLogoPng()));
  const W=210,M=17,CW=W-M*2,dark=[57,69,83],teal=[7,154,152],muted=[105,117,128],line=[220,226,231],pale=[245,249,249];
  const civ=buyer&&buyer.type==="morale"?"":((sru.civilites||{})[String(buyerIndex)]||"");
  const buyerName=[civ,sruPdfName(buyer)].filter(Boolean).join(" "),buyerAddress=buyer&&buyer.adresse||"";
  const sellers=(data.vendeurs||[]).map(sruPdfName).filter(Boolean).join(" / ");
  const propertyAddress=sru.adresseBien||data.adresseBien||"",bien=sru.designationCourte||data.designation||propertyAddress;
  const mode=sruPdfModeLabel(sru.modeEnvoi),retour=sru.adresseRetour||"ACTION IMMOBILIERE - 6 rue La Bruyere - 29200 BREST";
  const salutation=civ==="Mme"?"Madame,":civ==="M."?"Monsieur,":"Madame, Monsieur,";
  const soussigne=civ==="Mme"?"Je soussignee":civ==="M."?"Je soussigne":"Je soussigne(e)";
  const signRows=[];
  (data.vendeurs||[]).forEach((p,i)=>signRows.push(["Vendeur",sruPdfName(p),(sru.signatureDates||{})["vendeur-"+i]||""]));
  (data.acquereurs||[]).forEach((p,i)=>signRows.push(["Acquereur",sruPdfName(p),(sru.signatureDates||{})["acquereur-"+i]||""]));
  const signedDates=signRows.map(r=>r[2]).filter(Boolean).sort(),firstDate=signedDates[0]||"";

  const setText=(size=10,color=dark,style="normal")=>{doc.setFont("helvetica",style);doc.setFontSize(size);doc.setTextColor(...color)};
  const logo=()=>{const w=46,h=w*(596/842);doc.addImage(sruLogoData,"PNG",(W-w)/2,8,w,h,undefined,"FAST");return 8+h};
  const rule=y=>{doc.setDrawColor(...teal);doc.setLineWidth(.7);doc.line(M,y,W-M,y)};
  const title=(text,y)=>{setText(16,dark,"bold");doc.text(text,W/2,y,{align:"center"});return y+7};
  const subtitle=(text,y)=>{setText(8.6,muted,"normal");doc.text(text,W/2,y,{align:"center"});return y+6};
  const section=(text,y)=>{setText(9,teal,"bold");doc.text(text.toUpperCase(),M,y);doc.setDrawColor(...line);doc.setLineWidth(.25);doc.line(M,y+2,W-M,y+2);return y+7};
  const box=(label,value,x,y,w,h=20)=>{
    doc.setFillColor(...pale);doc.setDrawColor(...line);doc.roundedRect(x,y,w,h,2,2,"FD");
    setText(7.2,muted,"bold");doc.text(label.toUpperCase(),x+4,y+5);
    setText(9.2,dark,"bold");const lines=doc.splitTextToSize(value||"-",w-8);doc.text(lines,x+4,y+11,{maxWidth:w-8,lineHeightFactor:1.15});
  };
  const para=(text,y,opt={})=>{
    setText(opt.size||9.3,opt.color||dark,opt.bold?"bold":"normal");
    const lines=doc.splitTextToSize(text,CW-(opt.indent||0));doc.text(lines,M+(opt.indent||0),y,{lineHeightFactor:1.35});
    return y+lines.length*(opt.size||9.3)*0.47+(opt.after||3);
  };
  const pageFooter=()=>{
    setText(7,muted,"normal");doc.text("ACTION IMMOBILIERE - 6 rue La Bruyere - 29200 BREST",M,290);
    doc.text("Notification SRU",W-M,290,{align:"right"});
  };

  let y=logo(); y=title("NOTIFICATION DU DELAI DE RETRACTATION",y+2); y=subtitle("Article L.271-1 du Code de la construction et de l'habitation",y); rule(y); y+=6;
  box("Destinataire",buyerName,M,y,86,20);box("Adresse du destinataire",buyerAddress,M+90,y,CW-90,20);y+=27;
  y=section("Avant-contrat",y);
  box("Adresse du bien",propertyAddress,M,y,104,20);box("Mode de notification",mode,M+108,y,CW-108,20);y+=27;
  setText(7.5,muted,"bold");doc.text("DATES DE SIGNATURE DES PARTIES",M,y);y+=4;
  const col=[28,92,42],x0=M;
  doc.setFillColor(233,247,247);doc.setDrawColor(...line);doc.rect(x0,y,CW,8,"FD");
  setText(7,teal,"bold");doc.text("QUALITE",x0+3,y+5);doc.text("SIGNATAIRE",x0+col[0]+3,y+5);doc.text("DATE",x0+col[0]+col[1]+3,y+5);
  y+=8;
  for(const r of signRows){
    const rh=9;doc.setDrawColor(...line);doc.rect(x0,y,CW,rh);
    doc.line(x0+col[0],y,x0+col[0],y+rh);doc.line(x0+col[0]+col[1],y,x0+col[0]+col[1],y+rh);
    setText(7.7,dark,"normal");doc.text(r[0],x0+3,y+5.7);doc.text(doc.splitTextToSize(r[1]||"-",col[1]-5),x0+col[0]+3,y+5.7);
    setText(7.7,dark,"bold");doc.text(sruPdfDate(r[2])||"-",x0+col[0]+col[1]+3,y+5.7);y+=rh;
  }
  y+=7;setText(10,dark,"bold");doc.text(salutation,M,y);y+=7;
  y=para("Veuillez trouver sous ce pli un original du compromis de vente, accompagne le cas echeant de ses annexes, signe par les parties aux dates indiquees ci-dessus.",y);
  y=para("Vendeur(s) : "+(sellers||"-"),y,{bold:true,after:2});
  y=para("Bien concerne : "+(bien||"-"),y,{after:5});
  y=section("Delai de retractation",y);
  y=para("Nous vous informons que le delai de retractation de 10 jours dont vous beneficiez commencera a courir le lendemain de la premiere presentation de la presente.",y,{bold:true});
  y=para("Si vous decidez d'utiliser votre faculte de retractation, vous devez l'exercer par lettre recommandee avec demande d'avis de reception ; vous pouvez a cette fin utiliser le coupon de retractation figurant en page suivante.",y);
  y=para("Fait a "+(sru.lieu||"")+" le "+(sruPdfDate(sru.dateNotification)||"________________"),y,{after:1});
  setText(8.5,muted,"bold");doc.text("Pour ACTION IMMOBILIERE",M,y+2);pageFooter();

  doc.addPage();
  y=logo(); y=title("COUPON DE RETRACTATION",y+1); y=subtitle("A renvoyer par lettre recommandee avec accuse de reception a l'adresse indiquee ci-dessous.",y);rule(y);y+=7;
  y=section("Acquereur",y);
  box("Identite",buyerName,M,y,86,20);box("Adresse de notification",buyerAddress,M+90,y,CW-90,20);y+=28;
  y=para(soussigne+" "+(buyerName||"-")+", declare exercer ma faculte de retractation concernant le compromis de vente portant sur le bien suivant :",y);
  box("Bien concerne",(propertyAddress?propertyAddress+" - ":"")+(bien||"-"),M,y,CW,24);y+=31;
  box("Premiere date de signature",sruPdfDate(firstDate)||"-",M,y,70,18);box("Date de retractation","________________",M+74,y,58,18);box("Signature","",M+136,y,CW-136,18);y+=27;
  y=section("Adresse de retour",y);box("Destinataire du retour",retour,M,y,CW,20);y+=28;
  y=section("Tampon agence",y);
  const sw=98,sh=38,sx=(W-sw)/2,sy=y+2;
  doc.setDrawColor(...dark);doc.setLineWidth(.7);doc.roundedRect(sx,sy,sw,sh,3,3);
  doc.setLineWidth(.25);doc.roundedRect(sx+1.8,sy+1.8,sw-3.6,sh-3.6,2.2,2.2);
  setText(10,dark,"bold");doc.text("ACTION IMMOBILIERE",W/2,sy+8,{align:"center"});
  setText(7,teal,"bold");doc.text("VENTE - LOCATION - GESTION",W/2,sy+13,{align:"center"});
  setText(7,dark,"normal");doc.text("6 rue La Bruyere - 29200 BREST",W/2,sy+19,{align:"center"});
  doc.text("02 98 46 41 41",W/2,sy+23,{align:"center"});
  setText(6.2,dark,"normal");doc.text("RCS BREST 851 997 437 - CPI 2901 2019 000 042 411",W/2,sy+29,{align:"center"});
  doc.text("Adherent SNPI n°21695 - Garantie financiere QBE Europe SA/NV",W/2,sy+33,{align:"center"});
  pageFooter();
  return doc;
}
