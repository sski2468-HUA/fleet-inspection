/* ════════ 照片改存 Supabase Storage（共用：內部檢查缺失登錄 intaudit.html／船上改善回報 intaudit-reply.html）════════
   原本照片是 base64 直接塞在每筆資料裡（{name,dataUrl}），照片一多同步就慢、還會逾時。
   現在改成：本機編輯／離線草稿期間照片照舊是 base64（離線才能預覽、存草稿）；真正「寫進雲端」那一刻，
   才把每張照片上傳到 Storage 的 defect-photos 儲存桶，資料裡改存 {name,url,path}：
   - path：儲存桶裡的相對路徑（p/前兩碼/內容雜湊.jpg）。之後若搬到本地資料庫，只要換網址前綴，
     資料本身不用動。
   - url：完整網址，畫面直接拿來顯示（跟主系統、船上線上填單已經在用的格式一樣）。
   檔名用照片內容算出來（SHA-256），同一張照片不管上傳幾次、出現在登錄單或缺失紀錄裡，都只存一份。
   單張上傳失敗不擋存檔：那張保留原本的 base64，之後存檔會再試一次，資料不會不見。
   用法：PhotoCloud.init(CLOUD)；顯示用 PhotoCloud.src(p)；寫雲端前 await PhotoCloud.externalize(資料)。 */
const PhotoCloud=(()=>{
  const BUCKET="defect-photos";
  let cfg=null;
  const cache=new Map(); // dataUrl → {url,path}：同一張照片這次開頁面已經傳過，就不再重傳
  function init(c){cfg=c;}
  const isEmbedded=p=>p&&typeof p==="object"&&!Array.isArray(p)&&typeof p.dataUrl==="string"&&p.dataUrl.startsWith("data:image");
  /* PDF 附件（2026-10-27）：選檔當下直接上傳（不轉 base64），資料存 {name,url,path,type:"application/pdf",size}。
     縮圖位置顯示 PDF 圖示；圖示網址後面帶「#pdf=真正網址」，openLightbox 看到就改成另開 PDF。 */
  const PDF_MAX=5*1024*1024;   // 2026-11-05 PDF 統一每檔 5MB（節省雲端空間）
  const PDF_ICON="data:image/svg+xml;utf8,"+encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="8" fill="#fdecea"/><path d="M18 8h20l10 10v38H18z" fill="#fff" stroke="#c0392b" stroke-width="3"/><path d="M38 8v10h10" fill="none" stroke="#c0392b" stroke-width="3"/><text x="33" y="46" font-family="Arial" font-size="13" font-weight="700" fill="#c0392b" text-anchor="middle">PDF</text></svg>');
  const isPdf=p=>!!p&&typeof p==="object"&&(p.type==="application/pdf"||/\.pdf$/i.test(p.name||"")||/\.pdf(\?|#|$)/i.test(p.url||""));
  const isPdfFile=f=>!!f&&(f.type==="application/pdf"||/\.pdf$/i.test(f.name||""));
  /* 照片「待補傳」：上傳失敗時雲端只先存 {name,pending:true}（照片本體留在建檔那台電腦），別人看到的是這張提示圖 */
  const PENDING_ICON="data:image/svg+xml;utf8,"+encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 90"><rect width="120" height="90" rx="8" fill="#fff4e0"/><rect x="3" y="3" width="114" height="84" rx="6" fill="none" stroke="#e0a83a" stroke-width="2" stroke-dasharray="6 4"/><text x="60" y="42" font-family="sans-serif" font-size="13" font-weight="700" fill="#a05a10" text-anchor="middle">照片待補傳</text><text x="60" y="62" font-family="sans-serif" font-size="10" fill="#a05a10" text-anchor="middle">建檔人電腦存檔後補上</text></svg>');
  function src(p){if(!p)return "";if(isPdf(p))return PDF_ICON+"#pdf="+encodeURIComponent(p.url||"");return p.url||p.dataUrl||(p.pending?PENDING_ICON:"");}
  /* 還沒傳上 Storage 的照片數 */
  function countEmbedded(obj){const l=[];collect(obj,l);return l.length;}
  /* 複製一份要寫進雲端的資料：還沒傳上去的照片換成 {name,pending:true}，避免整張單因為照片太大存不進去 */
  function stripEmbedded(obj){
    return JSON.parse(JSON.stringify(obj,(k,v)=>isEmbedded(v)?{name:v.name||"照片",pending:true}:v));
  }
  function pdfUrlFromSrc(s){const m=/#pdf=(.*)$/.exec(String(s||""));return m?decodeURIComponent(m[1]):"";}
  async function uploadPdf(file){
    if(!cfg)throw new Error("PhotoCloud 尚未設定雲端連線");
    if(file.size>PDF_MAX)throw new Error(`超過 5MB（${(file.size/1024/1024).toFixed(1)}MB），請壓縮或拆檔`);
    if(navigator.onLine===false)throw new Error("PDF 需連網上傳，請連上網路後再選一次");
    const hash=await sha256Hex(await file.arrayBuffer());
    const path=`p/${hash.slice(0,2)}/${hash}.pdf`;
    const r=await fetch(`${cfg.url}/storage/v1/object/${BUCKET}/${path}`,{
      method:"POST",
      headers:{"apikey":cfg.key,"Authorization":"Bearer "+cfg.key,"Content-Type":"application/pdf","x-upsert":"false"},
      body:file
    });
    if(!r.ok){
      const t=await r.text().catch(()=>"");
      if(!(r.status===409||/exists|duplicate/i.test(t)))throw new Error(`HTTP ${r.status} ${t.slice(0,120)}`);
    }
    return {name:file.name||"附件.pdf",url:`${cfg.url}/storage/v1/object/public/${BUCKET}/${path}`,path,type:"application/pdf",size:file.size};
  }
  async function sha256Hex(buf){
    const h=await crypto.subtle.digest("SHA-256",buf);
    return [...new Uint8Array(h)].map(b=>b.toString(16).padStart(2,"0")).join("");
  }
  /* 2026-11-11：dataUrl 轉檔案不再用 fetch(dataUrl)——Mac／iPhone 的 Safari 對很長的 data: 網址會直接回
     「Load failed」，照片全部傳不上去、整張單也跟著存不進雲端。改成自己解 base64。 */
  function dataUrlToBlob(dataUrl){
    const i=dataUrl.indexOf(",");const head=dataUrl.slice(0,i),body=dataUrl.slice(i+1);
    const type=(/^data:([^;,]+)/.exec(head)||[])[1]||"image/jpeg";
    if(!/;base64/i.test(head))return new Blob([decodeURIComponent(body)],{type});
    const bin=atob(body),n=bin.length,u8=new Uint8Array(n);
    for(let k=0;k<n;k++)u8[k]=bin.charCodeAt(k);
    return new Blob([u8],{type});
  }
  const sleep=ms=>new Promise(r=>setTimeout(r,ms));
  async function upload(dataUrl){
    if(cache.has(dataUrl))return cache.get(dataUrl);
    if(!cfg)throw new Error("PhotoCloud 尚未設定雲端連線");
    const blob=dataUrlToBlob(dataUrl);
    const hash=await sha256Hex(await blob.arrayBuffer());
    const ext=blob.type==="image/png"?"png":"jpg";
    const path=`p/${hash.slice(0,2)}/${hash}.${ext}`;
    // 網路一時不穩（Safari 顯示 Load failed、Chrome 顯示 Failed to fetch）或伺服器忙碌（5xx）時自動重試，最多 3 次
    let lastErr=null;
    for(let tries=0;tries<3;tries++){
      if(tries)await sleep(tries===1?1200:3500);
      let r;
      try{
        r=await fetch(`${cfg.url}/storage/v1/object/${BUCKET}/${path}`,{
          method:"POST",
          headers:{"apikey":cfg.key,"Authorization":"Bearer "+cfg.key,"Content-Type":blob.type||"image/jpeg","x-upsert":"false"},
          body:blob
        });
      }catch(e){lastErr=new Error(`網路連線失敗（${e&&e.message||e}）`);continue;}
      if(r.ok){lastErr=null;break;}
      const t=await r.text().catch(()=>"");
      // 檔名＝內容雜湊，已經存在就代表同一張照片之前傳過了，直接沿用
      if(r.status===409||/exists|duplicate/i.test(t)){lastErr=null;break;}
      lastErr=new Error(`HTTP ${r.status} ${t.slice(0,120)}`);
      if(r.status<500)break; // 4xx 重試也沒用
    }
    if(lastErr)throw lastErr;
    const res={url:`${cfg.url}/storage/v1/object/public/${BUCKET}/${path}`,path};
    cache.set(dataUrl,res);
    return res;
  }
  function collect(o,out){
    if(!o||typeof o!=="object")return;
    if(Array.isArray(o)){o.forEach(x=>collect(x,out));return;}
    if(isEmbedded(o)){out.push(o);return;}
    for(const k in o){const v=o[k];if(v&&typeof v==="object")collect(v,out);}
  }
  /* 把 obj 裡所有內嵌照片上傳並換成 {name,url,path}（直接改 obj 本身）。回傳 {total,fail,first} */
  async function externalize(obj){
    const list=[];collect(obj,list);
    let fail=0,first="",netFail=0;
    for(const p of list){
      // 連續 2 張都是「連不上網路」就不再一張張重試（每張重試要等約 5 秒，10 張就要乾等快 1 分鐘），剩下的直接留到下次補傳
      if(netFail>=2&&!cache.has(p.dataUrl)){fail++;continue;}
      try{const r=await upload(p.dataUrl);p.url=r.url;p.path=r.path;delete p.dataUrl;netFail=0;}
      catch(e){fail++;if(!first)first=e.message||String(e);if(/^網路連線失敗/.test(e.message||""))netFail++;console.warn("照片上傳 Storage 失敗，這張先保留在資料裡",e);}
    }
    return {total:list.length,fail,first};
  }
  /* 把畫面上正在編輯的那份資料裡、已經傳過的照片也換成網址（不再上傳），避免下次存檔又整批重傳 */
  function applyCache(obj){
    const list=[];collect(obj,list);
    list.forEach(p=>{const r=cache.get(p.dataUrl);if(r){p.url=r.url;p.path=r.path;delete p.dataUrl;}});
  }
  /* Word／Excel 匯出需要圖片本體：網址照片下載回來轉成 dataUrl；讀不到回傳 "" */
  async function toDataUrl(p){
    if(!p||isPdf(p))return ""; // PDF 不能當圖片放進 Word／Excel
    if(p.dataUrl)return p.dataUrl;
    if(!p.url)return "";
    try{
      const b=await (await fetch(p.url)).blob();
      return await new Promise((res,rej)=>{const fr=new FileReader();fr.onload=()=>res(fr.result);fr.onerror=rej;fr.readAsDataURL(b);});
    }catch(e){console.warn("照片下載失敗",p.url,e);return "";}
  }
  return {init,src,externalize,applyCache,toDataUrl,BUCKET,isPdf,isPdfFile,uploadPdf,pdfUrlFromSrc,PDF_ICON,countEmbedded,stripEmbedded,PENDING_ICON};
})();
