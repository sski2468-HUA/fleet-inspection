/* ════════ Excel 匯出美化（共用）════════
   用法：原本 XLSX.writeFile(wb, 檔名) 改成 XlsxPretty.writeFile(wb, 檔名)。
   wb 照舊用 SheetJS（XLSX.utils.json_to_sheet / book_append_sheet）組好，這裡把每個分頁的內容與欄寬
   讀出來，改用 ExcelJS 重新產生有格式的檔案：
   - 每頁上方深藍標題列（檔名＋分頁名）與匯出時間
   - 表頭藍綠底白字、凍結、可篩選
   - 框線、隔列底色、長文字自動換行、數字靠右
   - 常見狀態字（已結案／未結案／待審核…）與風險（高／中／低）自動上色
   - A4 橫向、列印寬度縮成一頁
   需要先載入 ExcelJS；載入失敗時退回原本 SheetJS 的無格式匯出，不會匯不出來。 */
const XlsxPretty=(()=>{
  const C={navy:"FF1C3A5E",teal:"FF1E7F8C",line:"FFDDE4E8",zebra:"FFF5F8FA",ink:"FF15262F",ink3:"FF8195A0"};
  const FONT="Microsoft JhengHei";
  // 狀態字上色（只套在短字串，避免誤判長文字）
  const STATUS=[
    [/^(✓\s*)?(已結案|全部結案|結案|已審核|已處理|已匯入|已套用|Closed)$/i,"FFE2F3E8","FF20703F"],
    [/^(未結案|待承辦|已退回|已退回待重填|退回|待船上改善|逾期|未匯入|⚑.*|扣船|Open)$/i,"FFFDE8E6","FFC0392B"],
    [/^(處理中|待審核|待督導審核|船上已回報|審核中|進行中|已發改善單|已發單)$/,"FFFFF3CD","FF7A5A10"],
    [/^(草稿|FLOW.*|不適用|—|-)$/,"FFEEF1F3","FF52656F"]
  ];
  const RISK={"高":["FFFDE3E0","FFC0392B"],"中":["FFFDEED9","FFB5610F"],"低":["FFE2F3E8","FF2E8B57"]};
  const textLen=v=>{let n=0;for(const ch of String(v??""))n+=/[\u0000-ÿ]/.test(ch)?1:2;return n;};
  const lastCol=n=>{let s="";n++;while(n>0){const m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;};

  function addSheet(ewb,name,ws,title,stamp,noTitle){
    // noTitle：不加上方標題列，表頭留在第 1 列——給「匯出後可能再匯回系統」的檔案用，欄位位置跟原本一樣
    const R0=noTitle?1:3;
    const aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
    const nCol=Math.max(1,...aoa.map(r=>r.length));
    const sh=ewb.addWorksheet(name,{views:[{state:"frozen",ySplit:R0}],
      pageSetup:{orientation:"landscape",paperSize:9,fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:0.4,right:0.4,top:0.5,bottom:0.5,header:0.2,footer:0.2}}});
    // 欄寬：沿用原本設定的 wch；沒有設定的依內容估算
    const wch=(ws["!cols"]||[]).map(c=>c&&(c.wch||c.width));
    const widths=[];
    for(let c=0;c<nCol;c++){
      let w=wch[c];
      if(!w){w=Math.max(...aoa.slice(0,300).map(r=>textLen(r[c])),6)+2;}
      widths.push(Math.min(Math.max(w,6),60));
    }
    sh.columns=widths.map(w=>({width:w}));
    const L=lastCol(nCol-1);
    if(!noTitle){
    sh.mergeCells(`A1:${L}1`);
    const t=sh.getCell("A1");t.value=`${title}　｜　${name}`;t.font={name:FONT,size:14,bold:true,color:{argb:"FFFFFFFF"}};
    t.fill={type:"pattern",pattern:"solid",fgColor:{argb:C.navy}};t.alignment={vertical:"middle",indent:1};sh.getRow(1).height=28;
    sh.mergeCells(`A2:${L}2`);
    const s2=sh.getCell("A2");s2.value=`匯出時間：${stamp}　｜　共 ${Math.max(0,aoa.length-1)} 筆`;s2.font={name:FONT,size:9.5,color:{argb:C.ink3}};s2.alignment={vertical:"middle",indent:1};sh.getRow(2).height=18;
    }
    const thin={style:"thin",color:{argb:C.line}},box={top:thin,left:thin,bottom:thin,right:thin};
    // 判斷哪幾欄是數字欄、哪幾欄是長文字欄
    const body=aoa.slice(1);
    const isNum=[],isLong=[];
    for(let c=0;c<nCol;c++){
      const vals=body.map(r=>r[c]).filter(v=>v!==""&&v!=null);
      isNum[c]=vals.length>0&&vals.every(v=>typeof v==="number");
      isLong[c]=widths[c]>=24||vals.some(v=>textLen(v)>widths[c]*1.2);
    }
    aoa.forEach((row,ri)=>{
      const er=sh.getRow(R0+ri);
      for(let c=0;c<nCol;c++){
        const cell=er.getCell(c+1),v=row[c];
        cell.value=(v===undefined||v===null)?"":v;
        cell.border=box;
        if(ri===0){
          cell.font={name:FONT,size:10.5,bold:true,color:{argb:"FFFFFFFF"}};
          cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:C.teal}};
          cell.alignment={vertical:"middle",horizontal:"center",wrapText:true};
          continue;
        }
        cell.font={name:FONT,size:10.5,color:{argb:C.ink}};
        cell.alignment={vertical:"top",horizontal:isNum[c]?"right":"left",wrapText:!!isLong[c]};
        if(ri%2===0)cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:C.zebra}};
        const sv=typeof v==="string"?v.trim():"";
        if(sv&&textLen(sv)<=14){
          if(RISK[sv]){cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:RISK[sv][0]}};cell.font={name:FONT,size:10.5,bold:true,color:{argb:RISK[sv][1]}};cell.alignment={vertical:"top",horizontal:"center"};}
          else for(const [re,bg,fg] of STATUS){if(re.test(sv)){cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:bg}};cell.font={name:FONT,size:10.5,bold:true,color:{argb:fg}};cell.alignment={vertical:"top",horizontal:"center"};break;}}
        }
      }
      // 表頭字太長（例如「系統編號（請勿修改…）」）時自動加高，避免被切掉
      if(ri===0)er.height=Math.max(30,...row.map((v,c)=>Math.ceil(textLen(v)/Math.max(widths[c]-1,4))*15+6));
    });
    if(aoa.length>1)sh.autoFilter={from:{row:R0,column:1},to:{row:R0-1+aoa.length,column:nCol}};
    sh.headerFooter={oddFooter:`&L${title}&R&P / &N`};
    sh._R0=R0;
    return sh;
  }

  /* ── 照片嵌入：opts.photos={分頁名:[每一筆資料列 → {欄位名:[照片 dataUrl…]}]}（陣列順序＝資料列順序）。
     照片先縮小成 JPEG（最長邊 480px）再放進儲存格，檔案才不會大到開不了；儲存格裡排成一列小縮圖，
     在 Excel 點圖片可以拉大看。單一儲存格最多放 PHOTO_MAX 張。 ── */
  const PHOTO_W=96,PHOTO_H=72,PHOTO_GAP=4,PHOTO_MAX=6;
  function shrink(src){
    return new Promise(res=>{
      if(!src){res(null);return;}
      const img=new Image();
      img.onload=()=>{
        try{
          const k=Math.min(1,480/Math.max(img.naturalWidth,img.naturalHeight));
          const w=Math.max(1,Math.round(img.naturalWidth*k)),h=Math.max(1,Math.round(img.naturalHeight*k));
          const cv=document.createElement("canvas");cv.width=w;cv.height=h;
          const g=cv.getContext("2d");g.fillStyle="#fff";g.fillRect(0,0,w,h);g.drawImage(img,0,0,w,h);
          res({data:cv.toDataURL("image/jpeg",0.82),w,h});
        }catch(e){res(null);}
      };
      img.onerror=()=>res(null);
      img.src=src;
    });
  }
  async function placePhotos(ewb,sh,headers,rowsPhotos){
    const R0=sh._R0;
    const colIdx={};headers.forEach((h,i)=>{colIdx[h]=i;});
    const maxPer={};
    rowsPhotos.forEach(m=>Object.entries(m||{}).forEach(([k,l])=>{maxPer[k]=Math.max(maxPer[k]||0,Math.min((l||[]).length,PHOTO_MAX));}));
    // 照片欄加寬到放得下最多張的那一格（欄寬單位約 7px）
    Object.entries(maxPer).forEach(([k,n])=>{const c=colIdx[k];if(c==null||!n)return;const col=sh.getColumn(c+1);col.width=Math.max(col.width||10,Math.ceil((PHOTO_GAP+n*(PHOTO_W+PHOTO_GAP))/7)+1);});
    for(let i=0;i<rowsPhotos.length;i++){
      const m=rowsPhotos[i];if(!m)continue;
      let any=false;
      for(const [k,list] of Object.entries(m)){
        const c=colIdx[k];if(c==null)continue;
        const col=sh.getColumn(c+1),colPx=(col.width||10)*7;
        const srcs=(list||[]).slice(0,PHOTO_MAX);
        for(let j=0;j<srcs.length;j++){
          const s=await shrink(srcs[j]);if(!s)continue;
          const r=Math.min(PHOTO_W/s.w,PHOTO_H/s.h);
          const id=ewb.addImage({base64:s.data,extension:"jpeg"});
          sh.addImage(id,{tl:{col:c+(PHOTO_GAP+j*(PHOTO_W+PHOTO_GAP))/colPx,row:R0+i+0.08},ext:{width:Math.round(s.w*r),height:Math.round(s.h*r)},editAs:"oneCell"});
          any=true;
        }
        if((list||[]).length>PHOTO_MAX){const cell=sh.getRow(R0+1+i).getCell(c+1);cell.value=`（共 ${list.length} 張，只放前 ${PHOTO_MAX} 張）`;}
      }
      if(any){const row=sh.getRow(R0+1+i);row.height=Math.max(row.height||0,Math.ceil((PHOTO_H+10)*0.75));}
    }
  }

  async function writeFile(wb,filename,opts){
    opts=opts||{};
    if(typeof ExcelJS==="undefined"){XLSX.writeFile(wb,filename);return;}
    try{
      const ewb=new ExcelJS.Workbook();ewb.creator="FPMC";ewb.created=new Date();
      const title=opts.title||String(filename).replace(/\.xlsx?$/i,"").replace(/_\d{4}-\d{2}-\d{2}.*$/,"");
      const stamp=new Date().toLocaleString("zh-TW",{hour12:false});
      for(const n of wb.SheetNames){
        const sh=addSheet(ewb,n,wb.Sheets[n],title,stamp,!!opts.noTitle);
        const ph=opts.photos&&opts.photos[n];
        if(ph&&ph.length){
          const headers=(XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:""})[0]||[]).map(String);
          await placePhotos(ewb,sh,headers,ph);
        }
      }
      const buf=await ewb.xlsx.writeBuffer();
      const a=document.createElement("a");
      a.href=URL.createObjectURL(new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
      a.download=filename;document.body.appendChild(a);a.click();
      setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},1000);
    }catch(e){
      console.warn("XlsxPretty failed, fallback to plain export",e);
      XLSX.writeFile(wb,filename);
    }
  }
  return {writeFile};
})();
