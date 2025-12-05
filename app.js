// UPS Bill Summarizer - offline-friendly
(function(){
  const $ = (id) => document.getElementById(id);
  const statusRow = () => $("statusRow");

  function addChip(text, ok){
    const row = statusRow();
    if(!row) return;
    // clear placeholder on first real chip
    if(!addChip._cleared){
      row.innerHTML = "";
      addChip._cleared = true;
    }
    const el = document.createElement("div");
    el.className = "chip " + (ok ? "good" : "bad");
    el.textContent = text;
    row.appendChild(el);
  }

  function log(msg){
    const el = $("log");
    if(!el) return;
    el.textContent += ((el.textContent === "尚未開始") ? "" : "\n") + msg;
  }

  // visible startup ping
  addChip("app.js: loaded", true);
  addChip("protocol: " + location.protocol, true);

  window.addEventListener("error", (e) => {
    addChip("JS error: " + (e?.message || "unknown"), false);
    log("❌ JS error: " + (e?.message || "unknown"));
  });
  window.addEventListener("unhandledrejection", (e) => {
    const m = e?.reason?.message || String(e?.reason || "unknown");
    addChip("Promise reject: " + m, false);
    log("❌ Promise reject: " + m);
  });

  // Expect libraries from same-origin /vendor to satisfy strict CSP
  function loadScript(url){
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = url;
      s.async = true;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error("load fail: " + url));
      document.head.appendChild(s);
    });
  }

  async function boot(){
    try{
      await loadScript("./vendor/pdf.min.js");
      addChip("pdf.js: OK (local)", true);
    }catch(e){
      addChip("pdf.js: missing ./vendor/pdf.min.js", false);
      throw e;
    }

    try{
      // Set worker to local; if CSP blocks workers, we also disableWorker at document load time later.
      if(window.pdfjsLib?.GlobalWorkerOptions){
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.js";
        addChip("pdf worker: ./vendor/pdf.worker.min.js", true);
      }
    }catch(e){
      /* ignore */
    }

    try{
      await loadScript("./vendor/xlsx.full.min.js");
      addChip("xlsx: OK (local)", true);
    }catch(e){
      addChip("xlsx: missing ./vendor/xlsx.full.min.js", false);
      throw e;
    }

    addChip("runtime: ready", true);
    initApp();
  }

  // ===== App logic (same as earlier, shortened a bit for clarity) =====
  function initApp(){
    const runBtn = $("run");
    const exportBtn = $("export");

    function roundInt(n){ return Math.round((Number(n) + Number.EPSILON)); }
    function cleanNum(s){
      const t = String(s ?? "").replace(/[^\d\.\-]/g,"");
      const n = Number(t);
      return Number.isFinite(n) ? n : 0;
    }
    function dateToExcelSerial(dateUTC){
      const epoch = Date.UTC(1899,11,30);
      return (dateUTC.getTime() - epoch) / 86400000;
    }
    function isLikelyTracking(s){ return /^1Z[0-9A-Z]{16}$/i.test(s); }

    function groupIntoLines(items){
      const sorted = items
        .filter(it => it.str && it.str.trim() !== "")
        .map(it => ({...it, x: it.transform[4], y: it.transform[5]}))
        .sort((a,b)=> (b.y - a.y) || (a.x - b.x));

      const lines = [];
      const tol = 2.0;
      for(const it of sorted){
        let line = lines.find(l => Math.abs(l.y - it.y) <= tol);
        if(!line){ line = { y: it.y, items: [] }; lines.push(line); }
        line.items.push(it);
      }
      for(const l of lines){
        l.items.sort((a,b)=>a.x-b.x);
        l.text = l.items.map(i=>i.str).join(" ").replace(/\s+/g," ").trim();
      }
      return lines.sort((a,b)=>b.y-a.y);
    }

    const monthMap = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
    function parseMdFromLineText(t){
      const m1 = t.match(/\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b/i);
      if(m1){
        const dd = parseInt(m1[1],10);
        const mon = m1[2][0].toUpperCase()+m1[2].slice(1,3).toLowerCase();
        const mm = monthMap[mon];
        return mm ? `${mm}/${dd}` : null;
      }
      const m2 = t.match(/\b(\d{1,2})\/(\d{1,2})\b/);
      if(m2){
        const mm = parseInt(m2[1],10), dd = parseInt(m2[2],10);
        return `${mm}/${dd}`;
      }
      return null;
    }

    function shouldSkipLine(t){
      const s = t.toLowerCase();
      return s.includes("shipment total") ||
            (s.includes("total") && (s.includes("adjustment") || s.includes("non-taxable") || s.includes("taxable"))) ||
            s.includes("page:") ||
            (s.includes("invoice") && s.includes("date"));
    }

    function extractCharge(lineText){
      const nums = lineText.match(/-?\d{1,3}(?:,\d{3})*(?:\.\d+)?|-?\d+(?:\.\d+)?/g);
      if(!nums || nums.length === 0) return null;
      const amount = cleanNum(nums[nums.length-1]);
      let label = lineText.replace(nums[nums.length-1], "").trim();
      label = label.replace(/\b(TWD|NTD|USD|Charges?|Discount)\b/gi,"").trim();
      label = label.replace(/\s+/g," ").trim();
      if(label.length < 3) return null;
      return { label, amount };
    }

    function normalizeFeeLabel(label){
      const s = label.trim();
      const map = [
        [/^transportation/i, "Transportation"],
        [/fuel surcharge/i, "Fuel Surcharge"],
        [/disbursement fee/i, "Disbursement Fee"],
        [/entry prep fee/i, "Entry Prep Fee"],
        [/paper commercial invoice/i, "Paper Commercial Invoice"],
        [/^vat$/i, "VAT"],
      ];
      for(const [re, out] of map){ if(re.test(s)) return out; }
      return s;
    }

    function findInvoiceYear(allText){
      const m = allText.match(/\bInvoice\s*Date\b[^0-9]*(\d{4})[-\/]\d{1,2}[-\/]\d{1,2}/i)
            || allText.match(/\b(\d{4})[-\/]\d{1,2}[-\/]\d{1,2}\b/);
      return m ? parseInt(m[1],10) : null;
    }
    function mdToDateUTC(md, year){
      const [m,d] = String(md).split("/").map(x=>parseInt(x,10));
      const y = Number.isFinite(year) ? year : (new Date()).getFullYear();
      return new Date(Date.UTC(y, m-1, d));
    }

    async function parsePdf(file){
      const arr = await file.arrayBuffer();
      const pdf = await window.pdfjsLib.getDocument({ data: arr, disableWorker: true }).promise;

      let allText = "";
      for(let p=1; p<=pdf.numPages; p++){
        const page = await pdf.getPage(p);
        const content = await page.getTextContent();
        const lines = groupIntoLines(content.items);
        allText += "\n" + lines.map(l=>l.text).join("\n");
      }
      const year = findInvoiceYear(allText);
      log(`📄 解析：${file.name}（${pdf.numPages} 頁）→ 年份：${year ?? "未找到(用今年)"}`);

      const summary = [];
      const adjustments = [];

      for(let p=1; p<=pdf.numPages; p++){
        const page = await pdf.getPage(p);
        const content = await page.getTextContent();
        const lines = groupIntoLines(content.items);
        const pageText = lines.map(l=>l.text).join("\n").toLowerCase();

        const maybeWaybill = pageText.includes("tracking") || pageText.includes("waybill");
        if(!maybeWaybill) continue;

        const isAdjustmentPage = pageText.includes("adjustment");

        for(let i=0;i<lines.length;i++){
          const t = lines[i].text;
          if(shouldSkipLine(t)) continue;

          const md = parseMdFromLineText(t);
          if(!md) continue;

          const tokens = t.split(/\s+/);
          const trk = tokens.find(tok => isLikelyTracking(tok));
          if(!trk) continue;

          for(let j=i+1; j<Math.min(i+30, lines.length); j++){
            const lt = lines[j].text;
            if(shouldSkipLine(lt)) continue;

            const md2 = parseMdFromLineText(lt);
            const trk2 = lt.split(/\s+/).find(tok => isLikelyTracking(tok));
            if(md2 && trk2) break;

            const ch = extractCharge(lt);
            if(!ch) continue;

            const fee = normalizeFeeLabel(ch.label);
            const row = {file: file.name, md, year, tracking: trk, fee, net: ch.amount};
            (isAdjustmentPage ? adjustments : summary).push(row);
          }
        }
      }
      return {summary, adjustments};
    }

    function aggregateToWide(rows){
      const exemptRe = new RegExp(`^(?:${$("exempt").value})$`, "i");
      const taxRate = Number($("taxRate").value || "0.05");

      const agg = new Map();
      for(const r of rows){
        const key = `${r.file}||${r.year}||${r.md}||${r.tracking}||${r.fee}`;
        agg.set(key,(agg.get(key)||0)+Number(r.net||0));
      }

      const feeSet = new Set();
      const wideMap = new Map();
      for(const [k, amount] of agg.entries()){
        const [file, yearStr, md, tracking, fee] = k.split("||");
        const year = parseInt(yearStr,10);
        feeSet.add(fee);
        const k2 = `${file}||${year}||${md}||${tracking}`;
        if(!wideMap.has(k2)) wideMap.set(k2, {"檔名":file,"_year":year,"_md":md,"追蹤碼":tracking});
        wideMap.get(k2)[fee]=(wideMap.get(k2)[fee]||0)+amount;
      }

      const feeCols = Array.from(feeSet).sort((a,b)=>a.localeCompare(b));
      const out = Array.from(wideMap.values()).map(obj=>{
        for(const c of feeCols) obj[c]=obj[c]||0;
        const total = feeCols.reduce((s,c)=>s+Number(obj[c]||0),0);
        const taxBase = feeCols.reduce((s,c)=>s+(exemptRe.test(c)?0:Number(obj[c]||0)),0);
        const tax = taxBase*taxRate;

        for(const c of feeCols) obj[c]=roundInt(obj[c]);
        obj["折扣後運費合計"]=roundInt(total);
        obj["營業稅"]=roundInt(tax);
        obj["全部加總"]=roundInt(roundInt(total)+roundInt(tax));
        return obj;
      });

      const totalRow={"檔名":"全部加總","_year":"","_md":"","追蹤碼":""};
      for(const c of feeCols.concat(["折扣後運費合計","營業稅","全部加總"])){
        totalRow[c]=out.reduce((s,r)=>s+(Number(r[c])||0),0);
      }
      out.push(totalRow);

      const headers=["檔名","日期","追蹤碼",...feeCols,"折扣後運費合計","營業稅","全部加總"];
      const finalRows=out.map(r=>{
        const md=r._md||"";
        const year=r._year;
        const excelDate=md?dateToExcelSerial(mdToDateUTC(md,year)):"";
        const obj={"檔名":r["檔名"],"日期":excelDate,"追蹤碼":r["追蹤碼"]};
        for(const c of feeCols) obj[c]=r[c]??0;
        obj["折扣後運費合計"]=r["折扣後運費合計"]??0;
        obj["營業稅"]=r["營業稅"]??0;
        obj["全部加總"]=r["全部加總"]??0;
        return obj;
      });
      return {headers, rows: finalRows};
    }

    function renderPreview(headers, rows){
      const thead = $("tbl").querySelector("thead");
      const tbody = $("tbl").querySelector("tbody");
      thead.innerHTML=""; tbody.innerHTML="";
      const trh=document.createElement("tr");
      for(const h of headers){ const th=document.createElement("th"); th.textContent=h; trh.appendChild(th); }
      thead.appendChild(trh);

      const max=Math.min(200, rows.length);
      for(let i=0;i<max;i++){
        const tr=document.createElement("tr");
        for(const h of headers){
          const td=document.createElement("td");
          const v=rows[i][h];
          if(h==="日期" && typeof v==="number"){
            const js=new Date(Date.UTC(1899,11,30)+v*86400000);
            td.textContent=`${js.getUTCFullYear()}-${String(js.getUTCMonth()+1).padStart(2,"0")}-${String(js.getUTCDate()).padStart(2,"0")}`;
          }else td.textContent=(v??"").toString();
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
      }
    }

    function exportExcel(summaryPack, adjPack){
      const wb = window.XLSX.utils.book_new();
      const addSheet=(pack,name)=>{
        const ws = window.XLSX.utils.json_to_sheet(pack.rows, {header: pack.headers});
        window.XLSX.utils.book_append_sheet(wb, ws, name);
      };
      addSheet(summaryPack,"Summary");
      addSheet(adjPack,"Adjustments");
      window.XLSX.writeFile(wb, `UPS_Bill_Summary_${new Date().toISOString().slice(0,10)}.xlsx`);
    }

    let pickedFiles=[], lastSummary=null, lastAdj=null;

    function setFiles(files){
      pickedFiles=Array.from(files||[]).filter(f=>f.type==="application/pdf"||f.name.toLowerCase().endsWith(".pdf"));
      runBtn.disabled=pickedFiles.length===0;
      exportBtn.disabled=true;
      $("log").textContent=pickedFiles.length?`已選 ${pickedFiles.length} 份 PDF：\n- ${pickedFiles.map(f=>f.name).join("\n- ")}`:"尚未開始";
    }

    $("file").addEventListener("change", e=>setFiles(e.target.files));
    $("drop").addEventListener("dragover", e=>{e.preventDefault();});
    $("drop").addEventListener("drop", e=>{e.preventDefault(); setFiles(e.dataTransfer.files);});

    runBtn.addEventListener("click", async ()=>{
      runBtn.disabled=true; exportBtn.disabled=true;
      $("log").textContent="開始解析…";
      const allS=[], allA=[];
      for(const f of pickedFiles){
        const {summary, adjustments} = await parsePdf(f);
        allS.push(...summary); allA.push(...adjustments);
      }
      lastSummary=aggregateToWide(allS);
      lastAdj=aggregateToWide(allA);
      renderPreview(lastSummary.headers, lastSummary.rows);
      exportBtn.disabled=false; runBtn.disabled=false;
      log("✅ 完成");
    });

    exportBtn.addEventListener("click", ()=>{
      if(!lastSummary||!lastAdj) return;
      exportExcel(lastSummary, lastAdj);
    });

    addChip("UI: wired", true);
  }

  boot().catch(err=>{
    console.error(err);
    addChip(err?.message || String(err), false);
    log("❌ 啟動失敗：" + (err?.message || err));
  });
})();