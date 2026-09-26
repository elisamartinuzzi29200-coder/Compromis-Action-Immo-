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

function sruPdfModeLabel(mode){return mode==="LRE"?"Lettre recommandee electronique (LRE)":"Lettre recommandee avec accuse de reception (LRAR)"}
function sruPdfName(p){return p&&p.type==="morale"?(p.societe||""):[p&&p.prenoms,p&&p.nom].filter(Boolean).join(" ")}
function sruPdfDate(v){const m=String(v||"").match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?m[3]+"/"+m[2]+"/"+m[1]:String(v||"")}
async function buildSruPdf(data,sru,buyer,buyerIndex=0){
  if(!window.jspdf||!window.jspdf.jsPDF)throw new Error("Le moteur PDF n'est pas charge.");
  const {jsPDF}=window.jspdf,doc=new jsPDF({unit:"mm",format:"a4",orientation:"portrait",compress:true});
  const logoResponse=await fetch("./assets/action-immo-logo-sru.b64?v=2",{cache:"no-store"});
  if(!logoResponse.ok)throw new Error("Impossible de charger le logo Action Immobilière.");
  const logoBase64=(await logoResponse.text()).trim();
  const sruLogoData="data:image/png;base64,"+logoBase64;
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
  const logo=()=>{const w=46,h=w*(141/200);doc.addImage(sruLogoData,"PNG",(W-w)/2,8,w,h,undefined,"FAST");return 8+h};
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
