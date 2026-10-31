/* ---------- APP ---------- */
const app=document.getElementById("app");
let S={step:"lang",lang:"en",cat:null,qs:[],i:0,ans:[],fallback:false};
const esc=t=>String(t).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const max=()=>S.qs.reduce((s,q)=>s+Math.max(...q[1].map(o=>o[1])),0);

const STEPS={en:["Language","Category","Questions","Result"],si:["භාෂාව","කාණ්ඩය","ප්‍රශ්න","ප්‍රතිඵලය"],ta:["மொழி","பிரிவு","கேள்விகள்","முடிவு"]};
const ICONS={u18:"🎒",nw:"🎓",w:"💼"},GLYPH={en:"A",si:"අ",ta:"த"},SIDX={lang:0,cat:1,quiz:2,result:3};
const stepper=()=>{const L=STEPS[S.lang]||STEPS.en,a=SIDX[S.step];return `<ol class="steps" style="--p:${a/3}">${L.map((t,i)=>`<li class="${i<a?"done":i===a?"cur":""}"><i>${i<a?"✓":i+1}</i><span>${t}</span></li>`).join("")}</ol>`};
function pick(k){S.ans[S.i]=k;document.querySelectorAll(".opt[data-k]").forEach(b=>b.classList.toggle("on",+b.dataset.k===k));const n=document.getElementById("next");if(n)n.disabled=false}
function nextQ(){if(S.ans[S.i]==null)return;if(S.i===S.qs.length-1)S.step="result";else S.i++;render()}
function prevQ(){if(S.i>0)S.i--;else S.step="cat";render()}
addEventListener("keydown",e=>{if(S.step!=="quiz"||e.metaKey||e.ctrlKey||e.altKey)return;const n=S.qs[S.i][1].length,k="abcde".indexOf(e.key.toLowerCase()),d="12345".indexOf(e.key),x=k>=0?k:d;
 if(x>=0&&x<n)pick(x);else if(e.key==="ArrowRight")nextQ();else if(e.key==="ArrowLeft")prevQ()});
function confetti(){if(matchMedia("(prefers-reduced-motion: reduce)").matches)return;const c=["#f7941d","#fcc42a","#3a97d4","#243b73"],box=document.createElement("div");box.className="confetti";
 for(let i=0;i<48;i++){const s=document.createElement("i");s.style.cssText=`left:${Math.random()*100}%;background:${c[i%4]};animation-delay:${Math.random()*.7}s;animation-duration:${1.9+Math.random()*1.5}s`;box.appendChild(s)}
 document.body.appendChild(box);setTimeout(()=>box.remove(),4500)}
function render(){
  document.querySelector("header.top").classList.toggle("compact",S.step==="quiz");
  if(S.step==="lang"){
    app.innerHTML=stepper()+`<h1>Financial Health Questionnaire</h1><p class="sub">Choose your language<br>භාෂාව තෝරන්න<br>மொழியைத் தேர்ந்தெடுக்கவும்</p>`+
      LANGS.map(l=>`<button class="opt" data-lang="${l[0]}"><b>${GLYPH[l[0]]}</b><span>${l[1]}</span></button>`).join("");
    app.querySelectorAll("[data-lang]").forEach(b=>b.onclick=()=>{S.lang=b.dataset.lang;document.documentElement.lang=S.lang;S.step="cat";render()});
  }else if(S.step==="cat"){
    app.innerHTML=stepper()+`<h1>${T("cat")}</h1><p class="sub">${T("catSub")}</p>`+
      CATS.map(c=>`<button class="opt" data-cat="${c[0]}"><b>${ICONS[c[0]]}</b><span>${catName(c)}</span></button>`).join("")+
      `<div class="row"><button class="btn ghost" id="back">${T("back")}</button></div>`;
    app.querySelectorAll("[data-cat]").forEach(b=>b.onclick=()=>start(b.dataset.cat));
    document.getElementById("back").onclick=()=>{S.step="lang";render()};
  }else if(S.step==="quiz"){
    const q=S.qs[S.i],pk=S.ans[S.i],last=S.i===S.qs.length-1;
    app.innerHTML=`<div class="seg" role="progressbar" aria-valuemin="1" aria-valuemax="${S.qs.length}" aria-valuenow="${S.i+1}">${S.qs.map((_,i)=>`<i class="${i<S.i?"d":i===S.i?"c":""}"></i>`).join("")}</div>
      <div class="meta"><span>${T("q")(S.i+1,S.qs.length)}</span><span>${catName(CATS.find(c=>c[0]===S.cat))}</span></div>`+
      (S.fallback&&S.i===0?`<div class="note">Questions in this language are not available yet, so they are shown in English.</div>`:"")+
      `<div class="qcard"><h2>${esc(q[0])}</h2></div>`+
      q[1].map((o,k)=>`<button class="opt${pk===k?" on":""}" data-k="${k}"><b>${String.fromCharCode(65+k)}</b><span>${esc(o[0])}</span></button>`).join("")+
      `<div class="actions"><button class="btn ghost" id="prev">${T("back")}</button><button class="btn" id="next"${pk==null?" disabled":""}>${last?T("fin"):T("next")}</button></div>`;
    app.querySelectorAll("[data-k]").forEach(b=>b.onclick=()=>pick(+b.dataset.k));
    document.getElementById("prev").onclick=prevQ;document.getElementById("next").onclick=nextQ;
  }else{
    const total=S.qs.reduce((s,q,i)=>s+q[1][S.ans[i]][1],0),st=STICKERS.find(x=>total>=x[0]);
    if(!S.saved){saveFill(S.lang,S.cat,S.ans,total);S.saved=true}
    app.innerHTML=stepper()+`<div class="res"><h1>${T("thanks")}</h1>
      <div class="sticker" style="--sc:${st[3]}"><span class="m">${st[2]}</span><strong>${T("s")[STICKERS.indexOf(st)]}</strong><em>${total} / ${max()}</em></div>
      <button class="btn" id="again">${T("again")}</button></div><div class="dl-wrap"><button class="dl" id="share">Share by email</button><a href="#" class="dl" id="dl">Download fill-report.xlsx</a></div>`;
    document.getElementById("share").onclick=shareReport;
    document.getElementById("dl").onclick=e=>{e.preventDefault();downloadReport()};
    document.getElementById("again").onclick=()=>{S={step:"lang",lang:"en",cat:null,qs:[],i:0,ans:[],fallback:false};document.documentElement.lang="en";render()};
    if(total>=10)confetti();
  }
  window.scrollTo(0,0);
}

function start(cat){
  const set=QUESTIONS[cat];
  S.cat=cat;S.saved=false;S.fallback=!set[S.lang];S.qs=set[S.lang]||set.en;S.i=0;S.ans=[];S.step="quiz";render();
}
render();
