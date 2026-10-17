(()=>{
"use strict";
const SUPABASE_URL="https://nqzriajxhkiujlnldayx.supabase.co";
const SUPABASE_KEY="sb_publishable_y2unaxdazpqayrIq6kGSjw_trmoUiwJ";
const LOCAL_KEY="ai-compromis-dossiers";
const client=window.supabase.createClient(SUPABASE_URL,SUPABASE_KEY,{auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:true}});
const state={session:null,cache:[],ready:false,saveChain:Promise.resolve(),resolved:false};
let resolveReady;
window.COMPROMIS_CLOUD_READY=new Promise(resolve=>{resolveReady=resolve});

function clone(v){try{return structuredClone(v)}catch{return JSON.parse(JSON.stringify(v))}}
function localItems(){try{const v=JSON.parse(localStorage.getItem(LOCAL_KEY)||"[]");return Array.isArray(v)?v:[]}catch{return []}}
function setCloudStatus(text,bad=false){const el=document.getElementById("cloudStatus");if(el){el.textContent=text||"";el.classList.toggle("cloud-error",!!bad)}}
function showGate(message=""){
  document.body.classList.add("cloud-locked");
  const gate=document.getElementById("cloudGate");
  if(gate)gate.classList.remove("hidden");
  const msg=document.getElementById("cloudLoginMessage");
  if(msg)msg.textContent=message;
}
function hideGate(){
  const gate=document.getElementById("cloudGate");
  if(gate)gate.classList.add("hidden");
  document.body.classList.remove("cloud-locked");
}
function rowToItem(row){
  return {id:row.id,title:row.title||"Dossier sans nom",updatedAt:row.updated_at||row.created_at||new Date().toISOString(),data:row.dossier&&typeof row.dossier==="object"?row.dossier:{}};
}
async function fetchRows(){
  const {data,error}=await client.from("compromis").select("id,title,dossier,created_at,updated_at").order("updated_at",{ascending:false});
  if(error)throw error;
  state.cache=(data||[]).map(rowToItem);
  return state.cache;
}
async function upsertRaw(item){
  if(!state.session)throw new Error("Connexion requise");
  const updatedAt=item.updatedAt||new Date().toISOString();
  const payload={id:item.id,owner_id:state.session.user.id,title:item.title||"Dossier sans nom",dossier:clone(item.data||{}),updated_at:updatedAt};
  const {data,error}=await client.from("compromis").upsert(payload,{onConflict:"id"}).select("id,title,dossier,created_at,updated_at").single();
  if(error)throw error;
  const saved=rowToItem(data);
  const i=state.cache.findIndex(x=>x.id===saved.id);
  if(i>=0)state.cache[i]=saved;else state.cache.unshift(saved);
  state.cache.sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt)));
  try{localStorage.setItem(LOCAL_KEY,JSON.stringify(state.cache))}catch{}
  return saved;
}
async function mergeLocal(){
  const local=localItems();
  if(!local.length)return;
  const serverById=new Map(state.cache.map(x=>[x.id,x]));
  for(const item of local){
    if(!item||!item.id)continue;
    const server=serverById.get(item.id);
    const localDate=Date.parse(item.updatedAt||0)||0;
    const serverDate=Date.parse(server?.updatedAt||0)||0;
    if(!server||localDate>serverDate){
      await upsertRaw(item);
      serverById.set(item.id,item);
    }
  }
}
async function enter(session){
  state.session=session;
  hideGate();
  setCloudStatus("Connexion à vos dossiers…");
  try{
    await fetchRows();
    await mergeLocal();
    await fetchRows();
    state.ready=true;
    setCloudStatus("✓ Dossiers synchronisés");
  }catch(error){
    state.cache=localItems();
    state.ready=true;
    setCloudStatus("Mode local — synchronisation indisponible",true);
  }
  if(!state.resolved){state.resolved=true;resolveReady()}
}
async function initialize(){
  const form=document.getElementById("cloudLoginForm");
  if(form)form.addEventListener("submit",async ev=>{
    ev.preventDefault();
    const email=document.getElementById("cloudEmail")?.value.trim();
    const password=document.getElementById("cloudPassword")?.value||"";
    const button=form.querySelector("button[type=submit]");
    if(button)button.disabled=true;
    const msg=document.getElementById("cloudLoginMessage");
    if(msg)msg.textContent="Connexion…";
    const {data,error}=await client.auth.signInWithPassword({email,password});
    if(button)button.disabled=false;
    if(error){if(msg)msg.textContent="Connexion impossible : "+error.message;return}
    await enter(data.session);
  });
  const {data:{session}}=await client.auth.getSession();
  if(session)await enter(session);else showGate("Connectez-vous avec votre compte Action Immobilière pour retrouver vos compromis sur tous vos ordinateurs.");
  client.auth.onAuthStateChange((event,session)=>{
    if(event==="SIGNED_OUT"){state.session=null;state.ready=false;showGate("Vous êtes déconnecté.");return}
    if(session&&!state.session)enter(session);
  });
}
window.CompromisCloud={
  client,
  get ready(){return state.ready},
  get session(){return state.session},
  getDossiers(){return clone(state.cache)},
  setCache(list){state.cache=clone(Array.isArray(list)?list:[])},
  async refresh(){if(!state.session)return state.cache;await fetchRows();return clone(state.cache)},
  upsert(item){
    state.saveChain=state.saveChain.then(()=>upsertRaw(item));
    return state.saveChain;
  },
  async delete(id){
    if(!state.session)throw new Error("Connexion requise");
    const {error}=await client.from("compromis").delete().eq("id",id);
    if(error)throw error;
    state.cache=state.cache.filter(x=>x.id!==id);
    try{localStorage.setItem(LOCAL_KEY,JSON.stringify(state.cache))}catch{}
  }
};
initialize().catch(error=>{showGate("Impossible d’initialiser la connexion : "+error.message)});
})();