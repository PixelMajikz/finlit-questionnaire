/* Records every completed questionnaire in this browser and exports them as fill-report.xlsx (no libraries). */
const REPORT_KEY="finlit_fill_report";
const LANG_NAMES={en:"English",si:"Sinhala",ta:"Tamil"};
const CAT_NAMES={u18:"Under 18",nw:"Above 18 (non working)",w:"Above 18 (working)"};

const awardFor=s=>(STICKERS.find(x=>s>=x[0])||STICKERS[STICKERS.length-1])[1];
function loadFills(){try{return JSON.parse(localStorage.getItem(REPORT_KEY))||[]}catch(e){return[]}}
function saveFill(lang,cat,answers,score){
  const rows=loadFills();
  rows.push({n:rows.length+1,lang:LANG_NAMES[lang],cat:CAT_NAMES[cat],ans:answers.map(a=>String.fromCharCode(65+a)),score,award:awardFor(score)});
  try{localStorage.setItem(REPORT_KEY,JSON.stringify(rows))}catch(e){}
}

/* ---- minimal .xlsx writer: a zip with stored (uncompressed) entries ---- */
const crcTable=(()=>{const t=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
const crc32=b=>{let c=-1;for(const x of b)c=crcTable[(c^x)&255]^(c>>>8);return(c^-1)>>>0};
function zipFiles(files){
  const enc=new TextEncoder(),parts=[],cd=[];let off=0;
  const u16=n=>[n&255,(n>>8)&255],u32=n=>[n&255,(n>>8)&255,(n>>16)&255,(n>>24)&255];
  for(const [name,text] of files){
    const nm=enc.encode(name),d=enc.encode(text),crc=crc32(d);
    const lh=[...u32(0x04034b50),...u16(20),...u16(0x0800),...u16(0),...u16(0),...u16(0x21),...u32(crc),...u32(d.length),...u32(d.length),...u16(nm.length),...u16(0)];
    parts.push(new Uint8Array(lh),nm,d);
    cd.push(new Uint8Array([...u32(0x02014b50),...u16(20),...u16(20),...u16(0x0800),...u16(0),...u16(0),...u16(0x21),...u32(crc),...u32(d.length),...u32(d.length),...u16(nm.length),...u16(0),...u16(0),...u16(0),...u16(0),...u32(0),...u32(off)]),nm);
    off+=lh.length+nm.length+d.length;
  }
  const cdSize=cd.reduce((s,p)=>s+p.length,0);
  const end=new Uint8Array([...u32(0x06054b50),...u16(0),...u16(0),...u16(files.length),...u16(files.length),...u32(cdSize),...u32(off),...u16(0)]);
  return new Blob([...parts,...cd,end],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}
const xmlEsc=s=>String(s).replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c]));
const colName=i=>{let s="";for(i++;i>0;i=Math.floor((i-1)/26))s=String.fromCharCode(65+(i-1)%26)+s;return s};

function buildReport(rows){
  const head=["Number","Selected Language","Selected Age Category",...Array.from({length:10},(_,i)=>"Q"+(i+1)),"Final Score","Award"];
  const data=[head,...rows.map(r=>[r.n,r.lang,r.cat,...r.ans,r.score,r.award||awardFor(r.score)])];
  const sheetRows=data.map((row,ri)=>`<row r="${ri+1}">`+row.map((v,ci)=>{
    const ref=colName(ci)+(ri+1),st=ri===0?' s="1"':"";
    return typeof v==="number"?`<c r="${ref}"${st}><v>${v}</v></c>`:`<c r="${ref}"${st} t="inlineStr"><is><t>${xmlEsc(v)}</t></is></c>`;
  }).join("")+"</row>").join("");
  const ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main",rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships",ct="application/vnd.openxmlformats-officedocument.spreadsheetml.";
  const head0='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  return zipFiles([
    ["[Content_Types].xml",head0+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${ct}sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${ct}worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="${ct}styles+xml"/></Types>`],
    ["_rels/.rels",head0+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ["xl/workbook.xml",head0+`<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Fill Report" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels",head0+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`],
    ["xl/styles.xml",head0+`<styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`],
    ["xl/worksheets/sheet1.xml",head0+`<worksheet xmlns="${ns}"><cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="3" width="26" customWidth="1"/><col min="14" max="15" width="13" customWidth="1"/></cols><sheetData>${sheetRows}</sheetData></worksheet>`]
  ]);
}

function downloadReport(){
  const a=document.createElement("a");
  a.href=URL.createObjectURL(buildReport(loadFills()));
  a.download="fill-report.xlsx";
  document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(a.href),1500);
}

/* Password gate for the report download. Note: this runs in the browser, so it only keeps casual users out. */
const pwHash=s=>{let h1=0xdeadbeef,h2=0x41c6ce57;for(let i=0;i<s.length;i++){const c=s.charCodeAt(i);h1=Math.imul(h1^c,2654435761);h2=Math.imul(h2^c,1597334677)}h1=Math.imul(h1^(h1>>>16),2246822507)^Math.imul(h2^(h2>>>13),3266489909);h2=Math.imul(h2^(h2>>>16),2246822507)^Math.imul(h1^(h1>>>13),3266489909);return(4294967296*(2097151&h2)+(h1>>>0)).toString(36)};
const REPORT_PW_HASH="2c118swgd60";
function askPasswordThenDownload(){
  const o=document.createElement("div");o.className="modal";
  o.innerHTML=`<div class="modal-card" role="dialog" aria-modal="true" aria-labelledby="pwt"><h2 id="pwt">Enter password</h2><p>The fill report is password protected.</p><input id="pw" class="pw" type="password" autocomplete="off" aria-label="Password"><p class="pw-err" id="pwe" role="alert"></p><div class="row"><button class="btn ghost" id="pwc">Cancel</button><button class="btn" id="pwo">Download</button></div></div>`;
  document.body.appendChild(o);
  const inp=o.querySelector("#pw"),err=o.querySelector("#pwe");
  const go=()=>{if(pwHash(inp.value)===REPORT_PW_HASH){o.remove();downloadReport()}else{err.textContent="Incorrect password. Please try again.";inp.select()}};
  o.querySelector("#pwo").onclick=go;o.querySelector("#pwc").onclick=()=>o.remove();
  inp.addEventListener("keydown",e=>{if(e.key==="Enter")go();if(e.key==="Escape")o.remove()});
  inp.focus();
}
