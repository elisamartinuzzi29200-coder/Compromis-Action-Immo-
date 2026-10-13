(function(){
  const TARGETED_IMPORT_VERSION="2026-10-13-v2";
  const originalRender=render;
  const originalHandleDocumentImport=handleDocumentImport;
  const originalExtractPerson=extractPersonFromId;
  const originalExtractTitle=extractTitleProperty;

  function setImportStatus(message){
    const node=$("status");
    if(node)node.textContent=message||"";
  }
  function escapeAttr(value){return esc(String(value==null?"":value))}
  function canvasCopy(source,crop,targetWidth){
    crop=crop||{x:0,y:0,w:source.width,h:source.height};
    const scale=targetWidth?Math.min(4,targetWidth/crop.w):1;
    const out=document.createElement("canvas");
    out.width=Math.max(1,Math.round(crop.w*scale));
    out.height=Math.max(1,Math.round(crop.h*scale));
    const ctx=out.getContext("2d",{willReadFrequently:true});
    ctx.imageSmoothingEnabled=true;ctx.imageSmoothingQuality="high";
    ctx.drawImage(source,crop.x,crop.y,crop.w,crop.h,0,0,out.width,out.height);
    return out;
  }
  function improveForOcr(canvas,binary){
    const ctx=canvas.getContext("2d",{willReadFrequently:true});
    const image=ctx.getImageData(0,0,canvas.width,canvas.height),d=image.data;
    for(let i=0;i<d.length;i+=4){
      const gray=.299*d[i]+.587*d[i+1]+.114*d[i+2];
      let value;
      if(binary)value=gray<185?0:gray>225?255:Math.round((gray-185)*6.375);
      else value=Math.max(0,Math.min(255,Math.round((gray-128)*1.28+128)));
      d[i]=d[i+1]=d[i+2]=value;
    }
    ctx.putImageData(image,0,0);return canvas;
  }
  function detectCardCrop(source){
    const targetWidth=Math.min(760,source.width),scale=targetWidth/source.width;
    const sample=document.createElement("canvas");
    sample.width=Math.max(1,Math.round(source.width*scale));
    sample.height=Math.max(1,Math.round(source.height*scale));
    const ctx=sample.getContext("2d",{willReadFrequently:true});
    ctx.drawImage(source,0,0,sample.width,sample.height);
    const pixels=ctx.getImageData(0,0,sample.width,sample.height).data;
    const tile=14,cols=Math.ceil(sample.width/tile),rows=Math.ceil(sample.height/tile),active=new Uint8Array(cols*rows);
    for(let row=1;row<rows-1;row++)for(let col=1;col<cols-1;col++){
      const x0=col*tile,y0=row*tile,x1=Math.min(sample.width,x0+tile),y1=Math.min(sample.height,y0+tile);
      let dark=0,total=0;
      for(let y=y0;y<y1;y+=2)for(let x=x0;x<x1;x+=2){const p=(y*sample.width+x)*4,g=.299*pixels[p]+.587*pixels[p+1]+.114*pixels[p+2];if(g<222)dark++;total++}
      if(total&&dark/total>.055)active[row*cols+col]=1;
    }
    const seen=new Uint8Array(active.length),components=[];
    for(let row=1;row<rows-1;row++)for(let col=1;col<cols-1;col++){
      const start=row*cols+col;if(!active[start]||seen[start])continue;
      const queue=[start];seen[start]=1;let count=0,minC=col,maxC=col,minR=row,maxR=row;
      while(queue.length){
        const pos=queue.pop(),r=Math.floor(pos/cols),c=pos-r*cols;count++;minC=Math.min(minC,c);maxC=Math.max(maxC,c);minR=Math.min(minR,r);maxR=Math.max(maxR,r);
        for(let dr=-1;dr<=1;dr++)for(let dc=-1;dc<=1;dc++){if(!dr&&!dc)continue;const nr=r+dr,nc=c+dc;if(nr<1||nc<1||nr>=rows-1||nc>=cols-1)continue;const np=nr*cols+nc;if(active[np]&&!seen[np]){seen[np]=1;queue.push(np)}}
      }
      const w=(maxC-minC+1)*tile,h=(maxR-minR+1)*tile,aspect=w/h,area=w*h,density=count/((maxC-minC+1)*(maxR-minR+1));
      if(w>sample.width*.24&&h>sample.height*.09&&aspect>1.12&&aspect<2.15&&area<sample.width*sample.height*.42){
        const aspectBonus=Math.max(.25,1-Math.abs(aspect-1.58)/1.58);components.push({minC,maxC,minR,maxR,score:area*density*(.6+.4*aspectBonus)})
      }
    }
    components.sort((a,b)=>b.score-a.score);
    if(!components.length)return {x:source.width*.22,y:source.height*.27,w:source.width*.72,h:source.height*.48};
    const best=components[0],pad=tile*2;
    const x=Math.max(0,best.minC*tile-pad),y=Math.max(0,best.minR*tile-pad),w=Math.min(sample.width-x,(best.maxC-best.minC+1)*tile+2*pad),h=Math.min(sample.height-y,(best.maxR-best.minR+1)*tile+2*pad);
    return {x:x/scale,y:y/scale,w:w/scale,h:h/scale};
  }
  async function createOcrWorker(language,label){
    if(!window.Tesseract)throw Error("Le module de lecture d'image n'est pas chargé. Recharge la page.");
    return Tesseract.createWorker(language||"fra+eng",1,{logger:m=>{if(m.status==="recognizing text")setImportStatus((label||"Lecture du document")+"… "+Math.round((m.progress||0)*100)+"%")}});
  }
  async function renderPdfPage(page,scale){
    const viewport=page.getViewport({scale:scale||2}),canvas=document.createElement("canvas"),ctx=canvas.getContext("2d",{willReadFrequently:true});
    canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);await page.render({canvasContext:ctx,viewport}).promise;return canvas;
  }
  async function recognizeIdentityCanvas(source,fullWorker,mrzWorker,pageLabel){
    const crop=detectCardCrop(source),card=improveForOcr(canvasCopy(source,crop,2200),false);
    await fullWorker.setParameters({tessedit_pageseg_mode:"11",preserve_interword_spaces:"1"});
    setImportStatus("Lecture "+pageLabel+"…");
    const full=(await fullWorker.recognize(card)).data.text||"";
    const mrzCrop={x:0,y:Math.floor(card.height*.46),w:card.width,h:Math.ceil(card.height*.54)},mrz=improveForOcr(canvasCopy(card,mrzCrop,2600),true);
    await mrzWorker.setParameters({tessedit_pageseg_mode:"6",tessedit_char_whitelist:"ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<"});
    const machine=(await mrzWorker.recognize(mrz)).data.text||"";
    return full+"\n"+machine;
  }
  async function readIdentityDocuments(files){
    const list=Array.from(files||[]),parts=[],fullWorker=await createOcrWorker("fra+eng","Lecture de la pièce d'identité"),mrzWorker=await createOcrWorker("eng","Lecture de la zone sécurisée");
    try{
      for(let fi=0;fi<list.length;fi++){
        const file=list[fi];
        if(/\.pdf$/i.test(file.name)||file.type==="application/pdf"){
          const bytes=new Uint8Array(await file.arrayBuffer()),pdf=await pdfjsLib.getDocument({data:bytes}).promise,max=Math.min(2,pdf.numPages);
          for(let n=1;n<=max;n++){const page=await pdf.getPage(n),canvas=await renderPdfPage(page,2.35);parts.push(await recognizeIdentityCanvas(canvas,fullWorker,mrzWorker,"identité page "+n+" / "+max))}
        }else{
          const image=await loadImageFile(file),canvas=canvasFromImage(image,null,1);parts.push(await recognizeIdentityCanvas(canvas,fullWorker,mrzWorker,"identité "+(fi+1)+" / "+list.length));
        }
      }
    }finally{await Promise.allSettled([fullWorker.terminate(),mrzWorker.terminate()])}
    return parts.join("\n");
  }
  function titleCandidateScore(text){
    const t=String(text||"").toUpperCase();let score=0;
    if(/D[ÉE]SIGNATION/.test(t))score+=12;
    if(/IDENTIFICATION (?:DU BIEN|DES IMMEUBLES?)/.test(t))score+=10;
    if(/DROITS IMMOBILIERS/.test(t))score+=5;
    if(/LOT NUM[ÉE]RO/.test(t))score+=5;
    if(/CADASTRE|CADASTRAL/.test(t))score+=3;
    if(/\b\d{5}\b/.test(t))score+=1;
    return score;
  }
  async function recognizeTitlePage(page,worker,scale,highQuality){
    const canvas=await renderPdfPage(page,scale),texts=[];
    if(highQuality&&canvas.width>canvas.height*1.12){
      const left=canvasCopy(canvas,{x:0,y:0,w:Math.round(canvas.width*.53),h:canvas.height}),right=canvasCopy(canvas,{x:Math.round(canvas.width*.47),y:0,w:Math.round(canvas.width*.53),h:canvas.height});
      await worker.setParameters({tessedit_pageseg_mode:"6",preserve_interword_spaces:"1"});
      texts.push((await worker.recognize(left)).data.text||"");texts.push((await worker.recognize(right)).data.text||"");
    }else{
      await worker.setParameters({tessedit_pageseg_mode:highQuality?"6":"11",preserve_interword_spaces:"1"});texts.push((await worker.recognize(canvas)).data.text||"");
    }
    return texts.join("\n");
  }
  async function readTitleDocuments(files){
    const list=Array.from(files||[]),all=[];
    for(let fi=0;fi<list.length;fi++){
      const file=list[fi],worker=await createOcrWorker("fra+eng","Repérage de la désignation");
      try{
        if(/\.pdf$/i.test(file.name)||file.type==="application/pdf"){
          const bytes=new Uint8Array(await file.arrayBuffer()),pdf=await pdfjsLib.getDocument({data:bytes}).promise,max=Math.min(8,pdf.numPages),quick=[],candidates=[];
          for(let n=1;n<=max;n++){
            setImportStatus("Recherche de la désignation - page "+n+" / "+max+"…");
            const page=await pdf.getPage(n),text=await recognizeTitlePage(page,worker,1.65,false),score=titleCandidateScore(text);quick.push({n,text,score});
            if(score>=12){candidates.push(n);break}
          }
          if(!candidates.length){const best=quick.slice().sort((a,b)=>b.score-a.score)[0];if(best&&best.score>0)candidates.push(best.n)}
          if(candidates.length){
            const pages=[...new Set(candidates.concat(candidates.map(n=>n+1)))].filter(n=>n>=1&&n<=pdf.numPages).slice(0,3);
            for(const n of pages){setImportStatus("Lecture précise du titre - page "+n+"…");all.push(await recognizeTitlePage(await pdf.getPage(n),worker,2.65,true))}
          }else all.push(quick.map(x=>x.text).join("\n"));
        }else{
          const image=await loadImageFile(file),canvas=canvasFromImage(image,null,Math.max(1,Math.min(2.4,2600/Math.max(image.width,image.height))));
          await worker.setParameters({tessedit_pageseg_mode:"11",preserve_interword_spaces:"1"});all.push((await worker.recognize(canvas)).data.text||"");
        }
      }finally{await worker.terminate()}
    }
    return all.join("\n");
  }
  function normalizeMrzLine(value){return String(value||"").toUpperCase().replace(/[ «‹›]/g,"<").replace(/[^A-Z0-9<]/g,"").replace(/O(?=\d)/g,"0")}
  function mrzBirthDate(raw){
    if(!/^\d{6}$/.test(raw||""))return "";
    const yy=Number(raw.slice(0,2)),mm=Number(raw.slice(2,4)),dd=Number(raw.slice(4,6)),now=new Date().getFullYear();
    let year=2000+yy;if(year>now)year-=100;
    if(mm<1||mm>12||dd<1||dd>31||year<now-120)return "";
    return String(year).padStart(4,"0")+"-"+String(mm).padStart(2,"0")+"-"+String(dd).padStart(2,"0");
  }
  function nextLinesAfterLabel(lines,re,max){
    for(let i=0;i<lines.length;i++){
      if(!re.test(lines[i]))continue;
      re.lastIndex=0;
      const same=lines[i].replace(re,"").replace(/^[\s:/;()\-]+/,"").trim();if(same)return same;
      const found=[];
      for(let j=i+1;j<lines.length&&found.length<(max||2);j++){
        const line=lines[j].trim();if(!line)continue;
        if(/^(?:DATE|SEXE|SEX|NATIONAL|TAILLE|HEIGHT|LIEU|PLACE|N[°O]|DOCUMENT|R[ÉE]PUBLIQUE|SIGNATURE)/i.test(line)&&found.length)break;
        found.push(line);
      }
      return found.join(" ");
    }
    return "";
  }
  function extractIdentityV2(text){
    const base=originalExtractPerson(text)||{},clean=cleanDocText(text),lines=clean.split(/\n+/).map(x=>x.trim()).filter(Boolean),mrz=lines.map(normalizeMrzLine).filter(x=>x.length>=25&&x.length<=38);
    let l1="",l2="",l3="";
    for(let i=0;i<mrz.length;i++)if(/^I[DCA]FRA/.test(mrz[i])||/^IDFRA/.test(mrz[i])){l1=mrz[i];l2=mrz[i+1]||"";l3=mrz[i+2]||"";break}
    let nom=base.nom||"",prenoms=base.prenoms||"",naissance=base.naissance||"",civilite=base.civilite||"";
    if(l3&&l3.includes("<<")){const parts=l3.split("<<");nom=parts[0].replace(/<+/g," ").trim()||nom;prenoms=(parts.slice(1).join(" ")).replace(/<+/g," ").trim()||prenoms}
    if(l2){const digits=l2.replace(/O/g,"0");const m=digits.match(/^(\d{6})\d([MF<])/);if(m){naissance=mrzBirthDate(m[1])||naissance;civilite=m[2]==="F"?"Madame":m[2]==="M"?"Monsieur":civilite}}
    if(!nom)nom=nextLinesAfterLabel(lines,/^(?:NOM|NOM\s*\/\s*SURNAME|SURNAME)\s*/i,1);
    if(!prenoms)prenoms=nextLinesAfterLabel(lines,/^(?:PR[ÉE]NOMS?|PR[ÉE]NOMS?\s*\/\s*GIVEN NAMES?|GIVEN NAMES?)\s*/i,1);
    if(!naissance){const m=clean.match(/(?:DATE DE NAISS(?:ANCE)?|DATE OF BIRTH)[^0-9]{0,30}(\d{1,2}[\s./-]+\d{1,2}[\s./-]+\d{2,4})/i);if(m)naissance=normalizeDateFr(m[1])}
    let lieu=base.lieuNaissance||nextLinesAfterLabel(lines,/^(?:LIEU DE NAISSANCE|LIEU DE NAISS|PLACE OF BIRTH)(?:\s*\/[^:]*)?\s*/i,1);
    let adresse=base.adresse||nextLinesAfterLabel(lines,/^(?:ADRESSE|ADDRESS|DOMICILE)(?:\s*\/[^:]*)?\s*/i,3);
    if(!adresse){const m=clean.match(/\b(\d{1,4}\s+(?:RUE|AVENUE|BOULEVARD|PLACE|IMPASSE|ALL[ÉE]E|ROUTE|CHEMIN|QUAI|COURS|SQUARE)\s+[^\n]{2,70})\s*\n?\s*(\d{5}\s+[A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý '\-]{2,60})/i);if(m)adresse=(m[1]+" "+m[2]).replace(/\s+/g," ").trim()}
    const tidy=v=>String(v||"").replace(/[<>|]/g," ").replace(/\s{2,}/g," ").replace(/\b(?:SEXE|SEX|NATIONALIT[ÉE]|DATE DE NAISS).*$/i,"").trim();
    nom=tidy(nom).toUpperCase();prenoms=tidy(prenoms).toLowerCase().replace(/(^|[\s-])([a-zà-öø-ÿ])/g,(m,a,b)=>a+b.toUpperCase());
    lieu=tidy(lieu).replace(/\b(?:N[°O] DU DOCUMENT|DATE D['’]?EXPIR).*$/i,"").trim();adresse=tidy(adresse).replace(/\b(?:DATE DE D[ÉE]LIVRANCE|DATE D['’]?EXPIR|FRANCE\b).*$/i,"").trim();
    return {civilite:civilite||"Monsieur",nom,prenoms,naissance,lieuNaissance:lieu,adresse};
  }
  function cleanDesignationBlock(value){
    return cleanDocText(value).split(/\n+/).map(x=>x.replace(/^\s*[|Il1]+\s*$/g,"").replace(/\s+/g," ").trim()).filter(x=>x&&!/^\d{1,3}$/.test(x)).join("\n").replace(/\n{3,}/g,"\n\n").trim();
  }
  function extractTitleV2(text){
    const base=originalExtractTitle(text)||{},t=cleanDocText(text);
    let designation=sectionValue(t,/(?:^|\n)\s*(?:(?:IDENTIFICATION DU BIEN|IDENTIFICATION DES IMMEUBLES?)\s*)?(?:D[ÉE]SIGNATION(?:\s+DES\s+BIENS(?:\s+IMMOBILIERS)?)?)\s*:?\s*/im,/(?:^|\n)\s*(?:[ÉE]TAT DESCRIPTIF DE DIVISION|IDENTIFICATION DES MEUBLES|USAGE DU BIEN|EFFET RELATIF|PROPRI[ÉE]T[ÉE][ -]JOUISSANCE|CHARGES ET CONDITIONS|PRIX|PUBLICIT[ÉE] FONCI[ÈE]RE)\b/im,18000);
    if(!designation){designation=sectionValue(t,/(?:^|\n)\s*(A\s+[A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý '\-]+\s*\([^)]*\)\s*\(?\d{5}\)?\s+\d{1,4}\s+(?:RUE|AVENUE|BOULEVARD|PLACE|IMPASSE|ALL[ÉE]E|ROUTE|CHEMIN|QUAI|COURS|SQUARE)[^\n]*)/im,/(?:^|\n)\s*(?:[ÉE]TAT DESCRIPTIF DE DIVISION|IDENTIFICATION DES MEUBLES|USAGE DU BIEN|EFFET RELATIF)\b/im,12000)}
    designation=cleanDesignationBlock(designation||base.designation||"");
    let adresse=base.adresseBien||"",m=t.match(/(?:^|\n)\s*A\s+([A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý '\-]+)\s*\([^)]*\)\s*\(?(\d{5})\)?\s+(\d{1,4}\s+(?:RUE|AVENUE|BOULEVARD|PLACE|IMPASSE|ALL[ÉE]E|ROUTE|CHEMIN|QUAI|COURS|SQUARE)\s+[^\n.;]{2,100})/im);
    if(m)adresse=(m[3]+" "+m[2]+" "+m[1]).replace(/\s+/g," ").trim();
    const copro=/LOTS? DE COPROPRI[ÉE]T[ÉE]|LOT NUM[ÉE]RO|MILLI[ÈE]MES? DES PARTIES COMMUNES/i.test(designation||t)?"oui":"";
    return {adresseBien:adresse,designation,copropriete:copro,origineVendeur:base.origineVendeur||"",origineActe:base.origineActe||""};
  }
  function reviewFields(meta,result){
    const identity=/^(vendeur|acquereur):cni/.test(meta||"");
    if(identity)return [
      ["civilite","Civilité","select"],["nom","Nom","text"],["prenoms","Prénom(s)","text"],["naissance","Date de naissance","date"],["lieuNaissance","Lieu de naissance","text"],["adresse","Adresse","textarea"]
    ];
    return [["adresseBien","Adresse du bien","textarea"],["designation","Désignation complète","textarea-long"],["copropriete","Copropriété","select-copro"],["origineVendeur","Acquis de","textarea"],["origineActe","Date / notaire","textarea"]];
  }
  function showReview(meta,result,rawText){
    return new Promise(resolve=>{
      const fields=reviewFields(meta,result).filter(([key])=>result[key]!==undefined),overlay=document.createElement("div");overlay.className="docReviewOverlay";
      overlay.innerHTML='<div class="docReviewModal" role="dialog" aria-modal="true"><div class="docReviewHead"><div><span>LECTURE DOCUMENTAIRE</span><h3>Vérifier les informations détectées</h3></div><button type="button" class="docReviewClose" aria-label="Fermer">×</button></div><div class="docReviewNotice">Le document a été analysé uniquement dans ce navigateur. Vérifie chaque donnée avant de l’appliquer au dossier ; le fichier et le texte brut ne sont pas enregistrés.</div><form class="docReviewForm">'+fields.map(([key,label,type])=>{
        const value=result[key]||"",checked=value!==""&&!(meta==="titre"&&(key==="origineVendeur"||key==="origineActe"));let control;
        if(type==="select")control='<select data-review-value="'+key+'"><option value="Monsieur" '+(value==="Monsieur"?"selected":"")+'>Monsieur</option><option value="Madame" '+(value==="Madame"?"selected":"")+'>Madame</option></select>';
        else if(type==="select-copro")control='<select data-review-value="'+key+'"><option value=""></option><option value="oui" '+(value==="oui"?"selected":"")+'>Oui</option><option value="non" '+(value==="non"?"selected":"")+'>Non</option></select>';
        else if(type==="textarea"||type==="textarea-long")control='<textarea class="'+(type==="textarea-long"?"long":"")+'" data-review-value="'+key+'">'+escapeAttr(value)+'</textarea>';
        else control='<input type="'+type+'" data-review-value="'+key+'" value="'+escapeAttr(value)+'">';
        return '<div class="docReviewField"><label class="docReviewCheck"><input type="checkbox" data-review-check="'+key+'" '+(checked?"checked":"")+'><span>'+label+'</span></label>'+control+'</div>';
      }).join("")+'<details class="docReviewRaw"><summary>Afficher le texte reconnu (contrôle technique)</summary><pre>'+escapeAttr(String(rawText||"").slice(0,12000))+'</pre></details><div class="docReviewActions"><button type="button" class="docReviewCancel">Annuler</button><button type="submit" class="primary">Appliquer les champs cochés</button></div></form></div>';
      document.body.appendChild(overlay);const close=value=>{overlay.remove();resolve(value)};
      overlay.querySelector(".docReviewClose").onclick=()=>close(null);overlay.querySelector(".docReviewCancel").onclick=()=>close(null);overlay.addEventListener("click",e=>{if(e.target===overlay)close(null)});
      overlay.querySelector("form").onsubmit=e=>{e.preventDefault();const selected={};fields.forEach(([key])=>{const check=overlay.querySelector('[data-review-check="'+key+'"]'),input=overlay.querySelector('[data-review-value="'+key+'"]');if(check&&check.checked&&input)selected[key]=input.value.trim()});close(selected)};
    });
  }
  async function targetedHandle(files,meta){
    const list=Array.from(files||[]);if(!list.length)return;
    try{
      setImportStatus(meta==="titre"?"Recherche de la désignation…":"Lecture de la pièce d'identité…");
      const raw=meta==="titre"?await readTitleDocuments(list):await readIdentityDocuments(list),result=meta==="titre"?extractTitleV2(raw):extractIdentityV2(raw);
      const useful=Object.entries(result).filter(([key,value])=>value!==""&&value!==false&&value!=null);
      if(meta==="titre"&&!result.designation){alert("Aucune désignation suffisamment fiable n'a été repérée. Vérifie que le fichier contient les pages intitulées « IDENTIFICATION DU BIEN » ou « DÉSIGNATION ». Le titre de propriété peut être incomplet.");return}
      if(!useful.length){alert("Aucune information suffisamment fiable n'a été détectée. Essaie un scan plus net ou les deux faces de la pièce.");return}
      const reviewed=await showReview(meta,result,raw);if(reviewed&&Object.keys(reviewed).length){applyImportResult(meta,reviewed);setImportStatus("Informations vérifiées et appliquées")}
      else setImportStatus("Import annulé - aucune donnée modifiée");
    }catch(error){setImportStatus("");alert("Impossible de lire ce document : "+(error&&error.message?error.message:error))}
    finally{const input=$("documentImportInput");if(input)input.value=""}
  }
  handleDocumentImport=async function(files,meta){
    if(meta==="titre"||/^(vendeur|acquereur):cni:/.test(meta||""))return targetedHandle(files,meta);
    return originalHandleDocumentImport(files,meta);
  };
  function launchTargetedImport(meta,multiple){
    pendingImport=meta;const input=$("documentImportInput");input.multiple=!!multiple;input.accept=".pdf,.jpg,.jpeg,.png,.webp,application/pdf,image/*";input.click();
  }
  function addTargetedButtons(){
    if(viewMode!=="editor")return;
    if(step===0){
      [["vendeurs","vendeur"],["acquereurs","acquereur"]].forEach(([collection,singular])=>{
        (data[collection]||[]).forEach((person,index)=>{
          if((person.type||"physique")==="morale")return;
          const anchor=document.querySelector('[data-person="'+collection+':'+index+':nom"]');if(!anchor)return;const card=anchor.closest(".box"),head=card&&Array.from(card.children).find(x=>x.classList&&x.classList.contains("boxhead"));if(!head||head.querySelector(".identityImportButton"))return;
          const button=document.createElement("button");button.type="button";button.className="identityImportButton";button.textContent="Lire une pièce d'identité";button.onclick=()=>launchTargetedImport(singular+":cni:"+index,true);head.appendChild(button);
        });
      });
    }
    if(step===1){
      const content=$("content");if(!content||content.querySelector(".titleImportCard"))return;const firstBox=content.querySelector(".box"),card=document.createElement("div");card.className="box titleImportCard";card.innerHTML='<div><strong>Extraire la désignation du titre de propriété</strong><p class="hint">PDF ou image. L’application recherche les rubriques « Identification du bien », « Désignation », références cadastrales et lots de copropriété. Selon la longueur du titre, la lecture peut prendre quelques minutes.</p><p class="docPrivacy">Analyse locale : le document n’est pas envoyé ni conservé.</p></div><button type="button" class="primary">Lire le titre de propriété</button>';
      card.querySelector("button").onclick=()=>launchTargetedImport("titre",false);if(firstBox)content.insertBefore(card,firstBox);else content.appendChild(card);
    }
  }
  render=function(){const value=originalRender.apply(this,arguments);addTargetedButtons();return value};
  window.__documentImportV2={version:TARGETED_IMPORT_VERSION,extractIdentity:extractIdentityV2,extractTitle:extractTitleV2};
})();