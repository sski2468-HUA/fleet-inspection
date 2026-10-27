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
  const PDF_MAX=10*1024*1024;
  const PDF_ICON="data:image/svg+xml;utf8,"+encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="8" fill="#fdecea"/><path d="M18 8h20l10 10v38H18z" fill="#fff" stroke="#c0392b" stroke-width="3"/><path d="M38 8v10h10" fill="none" stroke="#c0392b" stroke-width="3"/><text x="33" y="46" font-family="Arial" font-size="13" font-weight="700" fill="#c0392b" text-anchor="middle">PDF</text></svg>');
  const isPdf=p=>!!p&&typeof p==="object"&&(p.type==="application/pdf"||/\.pdf$/i.test(p.name||"")||/\.pdf(\?|#|$)/i.test(p.url||""));
  const isPdfFile=f=>!!f&&(f.type==="application/pdf"||/\.pdf$/i.test(f.name||""));
  function src(p){if(!p)return "";if(isPdf(p))return PDF_ICON+"#pdf="+encodeURIComponent(p.url||"");return p.url||p.dataUrl||"";}
  function pdfUrlFromSrc(s){const m=/#pdf=(.*)$/.exec(String(s||""));return m?decodeURIComponent(m[1]):"";}
  async function uploadPdf(file){
    if(!cfg)throw new Error("PhotoCloud 尚未設定雲端連線");
    if(file.size>PDF_MAX)throw new Error(`超過 10MB（${(file.size/1024/1024).toFixed(1)}MB），請壓縮或拆檔`);
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
  async function upload(dataUrl){
    if(cache.has(dataUrl))return cache.get(dataUrl);
    if(!cfg)throw new Error("PhotoCloud 尚未設定雲端連線");
    const blob=await (await fetch(dataUrl)).blob();
    const hash=await sha256Hex(await blob.arrayBuffer());
    const ext=blob.type==="image/png"?"png":"jpg";
    const path=`p/${hash.slice(0,2)}/${hash}.${ext}`;
    const r=await fetch(`${cfg.url}/storage/v1/object/${BUCKET}/${path}`,{
      method:"POST",
      headers:{"apikey":cfg.key,"Authorization":"Bearer "+cfg.key,"Content-Type":blob.type||"image/jpeg","x-upsert":"false"},
      body:blob
    });
    if(!r.ok){
      const t=await r.text().catch(()=>"");
      // 檔名＝內容雜湊，已經存在就代表同一張照片之前傳過了，直接沿用
      if(!(r.status===409||/exists|duplicate/i.test(t)))throw new Error(`HTTP ${r.status} ${t.slice(0,120)}`);
    }
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
    let fail=0,first="";
    for(const p of list){
      try{const r=await upload(p.dataUrl);p.url=r.url;p.path=r.path;delete p.dataUrl;}
      catch(e){fail++;if(!first)first=e.message||String(e);console.warn("照片上傳 Storage 失敗，這張先保留在資料裡",e);}
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
  return {init,src,externalize,applyCache,toDataUrl,BUCKET,isPdf,isPdfFile,uploadPdf,pdfUrlFromSrc,PDF_ICON};
})();
