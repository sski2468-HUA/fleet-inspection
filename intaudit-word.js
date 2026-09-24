/* 內部檢查缺失登錄：共用資料定義與 Word（.docx）產生器。intaudit.html（督導端）與 intaudit-reply.html（船上端）共用。
   依賴頁面全域：JSZip、SHIPS（船舶清單，可為空）、toast；需與 intaudit-template.docx 放在同一資料夾。 */
const TYPE_DEFS=[
  {k:"onboard visit",label:"訪船",en:"Superintendent Visit",section:"訪船缺失",topicLabel:"訪船主題評估表",
   topics:["SAR-FM20 靜態航行檢查評估表","SAR-FM22 機艙管理檢查評估表","SAR-FM23 貨物操作檢查評估表","SAR-FM24 燃油加裝專項檢查表","SAR-FM25 主甲板、艏樓和繫泊檢查評估表","SAR-FM26 貨物機器間、貨泵艙、壓載泵艙和/或燃油泵艙檢查表","SAR-FM27 救生消防及住艙內外檢查表","其他"]},
  {k:"Internal audit",label:"內稽",en:"Internal Audit",section:"內稽缺失",topicLabel:"內稽區域／類別",topics:["甲板","機艙","ISM","MLC","ISPS","其他"]},
  {k:"ACCOMPANYING SHIP",label:"隨船",en:"On-Voyage Audit",section:"隨船缺失",topicLabel:"隨船主題評估表",
   topics:["SAR-FM21 動態航行檢查評估表","SAR-FM22 機艙管理檢查評估表","SAR-FM23 貨物操作檢查評估表","SAR-FM24 燃油加裝專項檢查表","SAR-FM25 主甲板、艏樓和繫泊檢查評估表","SAR-FM26 貨物機器間、貨泵艙、壓載泵艙和/或燃油泵艙檢查表","SAR-FM27 救生消防及住艙內外檢查表","其他"]},
  {k:"Cross ship visit",label:"交叉訪船",en:"Cross Visit",section:"交叉訪船缺失",topicLabel:"",topics:null},
  {k:"主題檢查",label:"主題檢查",en:"Topic Audit",section:"主題檢查缺失",topicLabel:"",topics:null}
];
/* 主題顯示文字：選「其他」時附上使用者輸入的說明 */
const topicText=d=>d.topic==="其他"&&d.topicOther?"其他："+d.topicOther:(d.topic||"");
/* 匯出檔名（Word／JSON／內控檔共用）：船名-檢查日期-Audit N_C _ Deficiency Items Details Records（檔名不能含斜線，改用底線） */
const exportBaseName=r=>`${r.ship||"船名"}-${(dateText(r)||"檢查日期").replace(/\//g,"-").replace(/～/g,"~")}-Audit N_C _ Deficiency Items Details Records`;
const typeDef=k=>TYPE_DEFS.find(t=>t.k===k);
const MODES=[{k:"FLOW",label:"FLOW 系統（DMP-FM01）",cls:"flow"},{k:"內控",label:"內控",cls:"ic"},{k:"系統內結案",label:"系統內結案",cls:"sys"}];
const modeOf=k=>MODES.find(m=>m.k===k);
const slash=s=>String(s||"").replace(/-/g,"/");
const dateText=r=>r.dateFrom?(slash(r.dateFrom)+(r.dateTo&&r.dateTo!==r.dateFrom?"～"+slash(r.dateTo):"")):"";
/* ══════ Word 匯出：以原「檢查缺失具體項目記錄表」的頁首／頁尾／版面設定為範本（intaudit-template.docx，
   保留公司標題列、SAR-FM30 頁尾、A4 橫向），內文依登錄單內容重新排版並附照片，輸出 .docx ══════ */
const W_NS_A="http://schemas.openxmlformats.org/drawingml/2006/main",W_NS_PIC="http://schemas.openxmlformats.org/drawingml/2006/picture";
const xe=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const wRun=(t,o)=>{o=o||{};const rp=`${o.b?"<w:b/>":""}${o.i?"<w:i/>":""}${o.color?`<w:color w:val="${o.color}"/>`:""}<w:sz w:val="${o.sz||20}"/><w:szCs w:val="${o.sz||20}"/>`;
  const lines=String(t??"").split(/\r?\n/);
  return `<w:r><w:rPr>${rp}</w:rPr>${lines.map((l,i)=>(i?"<w:br/>":"")+`<w:t xml:space="preserve">${xe(l)}</w:t>`).join("")}</w:r>`;};
const wPara=(inner,o)=>{o=o||{};return `<w:p><w:pPr>${o.keepNext?"<w:keepNext/>":""}${o.pb?"<w:pageBreakBefore/>":""}${o.shd?`<w:shd w:val="clear" w:color="auto" w:fill="${o.shd}"/>`:""}<w:spacing w:before="${o.before||0}" w:after="${o.after==null?40:o.after}"/>${o.jc?`<w:jc w:val="${o.jc}"/>`:""}</w:pPr>${inner}</w:p>`;};
const wCell=(w,inner,o)=>{o=o||{};return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${o.span?`<w:gridSpan w:val="${o.span}"/>`:""}${o.shd?`<w:shd w:val="clear" w:color="auto" w:fill="${o.shd}"/>`:""}<w:vAlign w:val="${o.va||"top"}"/></w:tcPr>${inner||wPara("")}</w:tc>`;};
const wTbl=(cols,rows)=>`<w:tbl><w:tblPr><w:tblW w:w="${cols.reduce((a,b)=>a+b,0)}" w:type="dxa"/><w:tblBorders>${["top","left","bottom","right","insideH","insideV"].map(s=>`<w:${s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`).join("")}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${cols.map(c=>`<w:gridCol w:w="${c}"/>`).join("")}</w:tblGrid>${rows.map(r=>`<w:tr><w:trPr><w:cantSplit/></w:trPr>${r}</w:tr>`).join("")}</w:tbl>`;
const LBL="E4F1F2";
const lblCell=(w,t,o)=>wCell(w,wPara(wRun(t,{b:true,sz:18}),{after:0}),Object.assign({shd:LBL},o||{}));
const txtCell=(w,t,o)=>wCell(w,wPara(wRun(t,{sz:19}),{after:0}),o);

async function imgSize(dataUrl){return new Promise(res=>{const im=new Image();im.onload=()=>res({w:im.naturalWidth,h:im.naturalHeight});im.onerror=()=>res({w:4,h:3});im.src=dataUrl;});}
function buildDocx_photos(ctx){
  /* 回傳 async 函式：把 photos 陣列轉成一段含多張圖片的 <w:p>，同時把圖檔登記到 ctx（之後寫進 zip） */
  return async function(photos,maxCm){
    if(!photos||!photos.length)return wPara("",{after:0});
    let runs="";
    for(const p of photos){
      const id=++ctx.n,sz=await imgSize(p.dataUrl),box=maxCm*360000,sc=Math.min(box/sz.w,box*0.8/sz.h);
      const cx=Math.round(sz.w*sc),cy=Math.round(sz.h*sc),rid="rIdIa"+id;
      ctx.files.push({name:`word/media/ia${id}.jpeg`,b64:p.dataUrl.split(",")[1],rid,target:`media/ia${id}.jpeg`});
      runs+=`<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="114300"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${1000+id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${W_NS_A}" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="${W_NS_A}"><a:graphicData uri="${W_NS_PIC}"><pic:pic xmlns:pic="${W_NS_PIC}"><pic:nvPicPr><pic:cNvPr id="${1000+id}" name="ia${id}.jpeg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>`;
    }
    return `<w:p><w:pPr><w:spacing w:before="20" w:after="20"/></w:pPr>${runs}</w:p>`;
  };
}
const MODE_SUFFIX={"FLOW":"需要登記到DMP-FM01中","內控":"需要登記內控中（不用輸入DMP-FM01）","系統內結案":"無需輸入FLOW系統追蹤，本表改善即可"};
const SECTION_HEAD={"onboard visit":"訪船缺失 Ship Visit","Internal audit":"內稽缺失 Internal Audit","ACCOMPANYING SHIP":"隨船缺失 On-Voyage Audit","Cross ship visit":"交叉訪船缺失 Cross Visit","主題檢查":"主題檢查缺失 Topic Inspection"};
async function buildDocBody(r,ctx,replyOf){
  const P=buildDocx_photos(ctx),W=14678;
  const ship=SHIPS.find(s=>s.code===r.ship),shipTxt=ship?`${ship.code} ${ship.name}`:r.ship;
  let b="";
  b+=wPara(wRun("檢查缺失具體項目記錄表 Audit N/C / Deficiency Items Details Records",{sz:28}),{jc:"center",after:120});
  const typeChecks=TYPE_DEFS.map(t=>`${(r.types||[]).includes(t.k)?"☑":"☐"} ${t.label} ${t.en}`).join("　　");
  b+=wTbl([2400,4939,2400,4939],[
    lblCell(2400,"訪問船舶 Ship Visited：")+txtCell(4939,shipTxt)+lblCell(2400,"日期 Date：")+txtCell(4939,dateText(r)),
    lblCell(2400,"檢查人 Auditor：")+txtCell(4939,r.inspector)+lblCell(2400,"地點 Ship's PSN：")+txtCell(4939,r.psn),
    lblCell(2400,"檢查種類 Audit Type：")+txtCell(12278,typeChecks,{span:3}),
    lblCell(2400,"船長 Captain：")+txtCell(4939,r.master)+lblCell(2400,"大副 C/O：")+txtCell(4939,r.co),
    lblCell(2400,"輪機長 C/E：")+txtCell(4939,r.ce)+lblCell(2400,"大管輪 2/E：")+txtCell(4939,r.e2)
  ]);
  b+=wPara(wRun("*這份缺失記錄僅為具體缺失項目記錄表格，因搭配不同的檢查種類要求的表格使用。This deficiency record is intended solely for documenting specific deficiencies and shall be used together with the relevant forms required for the applicable type of inspection.",{sz:16,color:"666666"}),{before:60,after:120});
  const crewTxt=a=>(a||[]).filter(x=>x.name||x.brief).map(x=>`${x.name||""}${x.rank?`（${x.rank}）`:""}：${x.brief||""}`).join("\n")||"N/A";
  b+=wTbl([2900,11778],[
    lblCell(2900,"總體評價：船上總體氣氛、TOP4配合程度、管理效果、船副管輪表現、乙級船員表現、船舶總體保養情況、機器狀況等\nOverall Assessment")+txtCell(11778,r.overall||""),
    lblCell(2900,"TOP 4 合照")+wCell(11778,await P(r.top4&&r.top4.photos,5.5)+(r.top4&&r.top4.cap?wPara(wRun(r.top4.cap,{sz:18}),{after:0}):"")),
    lblCell(2900,"表現亮眼船員姓名和簡述\nGood Performance Crew & Brief")+txtCell(11778,crewTxt(r.goodCrew)),
    lblCell(2900,"表現較差船員姓名和簡述\nPoor Performance Crew & Brief")+txtCell(11778,crewTxt(r.poorCrew))
  ]);
  const heading=(t,pb)=>wPara(wRun(t,{b:true,sz:22}),{shd:"D9E2E8",before:200,after:80,keepNext:true,pb});
  const tinyGap=()=>wPara(wRun("",{sz:2}),{after:0});
  if((r.best||[]).length){
    b+=heading("● 優良案例 Best Practice");
    for(let i=0;i<r.best.length;i++){
      const x=r.best[i];
      b+=wTbl([2400,12278],[
        lblCell(2400,"No.")+txtCell(12278,String(i+1)),
        lblCell(2400,"分類 Category")+txtCell(12278,x.category),
        lblCell(2400,"涉及人員姓名")+txtCell(12278,(x.persons&&x.persons.length)?x.persons.map(q=>q.name||"").join("\n"):x.people),
        lblCell(2400,"職務")+txtCell(12278,(x.persons&&x.persons.length)?x.persons.map(q=>q.rank||"").join("\n"):x.rank),
        lblCell(2400,"描述 Brief")+txtCell(12278,x.brief),
        lblCell(2400,"案例照片 Photos")+wCell(12278,await P(x.photos,5))
      ])+wPara("",{after:60});
    }
  }
  // 缺失：依「檢查種類（＋主題）＋結案方式」分組，各組一個標題，跟原記錄表的分區一致
  const order=TYPE_DEFS.map(t=>t.k),modeOrder=MODES.map(m=>m.k),groups=new Map();
  (r.defs||[]).forEach(d=>{
    const t=typeDef(d.type),topic=t&&t.topics&&["Internal audit"].includes(d.type)&&["MLC","ISPS","ISM"].includes(d.topic)?d.topic:"";
    const key=[order.indexOf(d.type),topic,modeOrder.indexOf(d.closeMode)].join("|");
    if(!groups.has(key))groups.set(key,{type:d.type,topic,mode:d.closeMode,items:[]});
    groups.get(key).items.push(d);
  });
  const cols=[900,2200,4400,3600,1800,1778];
  let defN=0; // 每一條缺失獨立一頁：第二條起（含換到下一個分區標題）都從新的一頁開始
  for(const key of [...groups.keys()].sort((a,b)=>{const A=a.split("|").map(Number),B=b.split("|").map(Number);return A[0]-B[0]||String(a.split("|")[1]).localeCompare(b.split("|")[1])||A[2]-B[2];})){
    const g=groups.get(key);
    b+=heading(`● ${SECTION_HEAD[g.type]||g.type}${g.topic?" - "+g.topic:""} - ${MODE_SUFFIX[g.mode]||""}`,defN>0);
    for(let i=0;i<g.items.length;i++){
      const d=g.items[i],mm=modeOf(d.closeMode),rp=(replyOf&&replyOf(d))||null;
      defN++;
      b+=wTbl(cols,[
        wCell(900,wPara(wRun("No.",{b:true,sz:18}),{after:0,pb:i>0}),{shd:LBL})+lblCell(2200,"分類 Category")+lblCell(4400,"待改善項目 Items to be corrected")+lblCell(3600,"改善完成證據/備註 Corrected Evidence / Remark")+lblCell(1800,"Reported @")+lblCell(1778,"Corrected @"),
        txtCell(900,String(i+1))+wCell(2200,wPara(wRun(d.category||"",{b:true,sz:19}),{after:0})+(topicText(d)?wPara(wRun(topicText(d),{sz:16,color:"666666"}),{after:0}):""))+wCell(4400,await P(d.photos,2.4))+wCell(3600,wPara(wRun(rp?rp.evidence||"":"",{sz:19}),{after:0})+(rp&&(rp.photos||[]).length?await P(rp.photos,2.4):""))+txtCell(1800,slash(d.reported))+txtCell(1778,rp?slash(rp.corrected):""),
        lblCell(3100,"Risk Level",{span:2})+txtCell(8000,d.risk||"",{span:2})+lblCell(1800,"結案方式")+txtCell(1778,mm?mm.label:""),
        lblCell(3100,"待改善描述 Finding Brief",{span:2})+txtCell(11578,d.finding||"",{span:4}),
        lblCell(3100,"改善措施簡述 Action Brief",{span:2})+txtCell(11578,rp?rp.action||"":"",{span:4})
      ].concat(rp&&rp.by?[lblCell(3100,"船上回報人 Replied by",{span:2})+txtCell(11578,`${rp.by}${rp.title?"（"+rp.title+"）":""}　${rp.at?new Date(rp.at).toLocaleString("zh-TW",{hour12:false}):""}${rp.status==="reviewed"?"　✓ 海技已審核結案":rp.status==="submitted"?"　（待海技審核）":""}`,{span:4})]:[]))+tinyGap();
    }
  }
  if((r.sugg||[]).length){
    b+=heading("● 船舶訴求/建議 Ship's Demand / Suggestion");
    const sRows=[lblCell(900,"No.")+lblCell(5200,"現狀 Situation")+lblCell(5200,"建議 Suggestion")+lblCell(3378,"照片 Photos")];
    for(let i=0;i<r.sugg.length;i++){const s=r.sugg[i];sRows.push(txtCell(900,String(i+1))+txtCell(5200,s.situation)+txtCell(5200,s.suggestion)+wCell(3378,await P(s.photos,3.2)));}
    b+=wTbl([900,5200,5200,3378],sRows);
  }
  return b;
}
/* 產生 .docx：replyOf(d) 回傳該缺失的船上回報內容（{action,corrected,evidence,photos,by,title,at,status}）或 null。
   回傳 {blob, photoCount}；下載由呼叫端負責（督導端與船上端的檔名相同：船名_檢查日期_檢查缺失具體項目記錄表.docx）。 */
async function iaMakeDocx(r,replyOf){
  if(typeof JSZip==="undefined")throw new Error("Word 元件載入失敗，請確認網路後重新整理頁面");
  const resp=await fetch("intaudit-template.docx");
  if(!resp.ok)throw new Error("找不到範本檔 intaudit-template.docx（HTTP "+resp.status+"），請確認已一併上傳");
  const zip=await JSZip.loadAsync(await resp.arrayBuffer());
  const ctx={n:0,files:[]};
  const body=await buildDocBody(r,ctx,replyOf);
  let doc=await zip.file("word/document.xml").async("string");
  if(!doc.includes("<!--BODY-->"))throw new Error("範本檔內容不符");
  doc=doc.replace("<!--BODY-->",()=>body);
  zip.file("word/document.xml",doc);
  let rels=await zip.file("word/_rels/document.xml.rels").async("string");
  ctx.files.forEach(f=>{zip.file(f.name,f.b64,{base64:true});rels=rels.replace("</Relationships>",`<Relationship Id="${f.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${f.target}"/></Relationships>`);});
  zip.file("word/_rels/document.xml.rels",rels);
  const blob=await zip.generateAsync({type:"blob",mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document",compression:"DEFLATE"});
  return {blob,photoCount:ctx.n};
}