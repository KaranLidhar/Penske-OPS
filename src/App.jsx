import { useState, useEffect } from "react";

// ── CONSTANTS ─────────────────────────────────────────────────────────────
const TRUCK_TYPES = ["High Roof Cargo Van","15 Cube","16 Cabover","26 FT CDL","26 FT G Class","18 Reefer","26 CDL Reefer","26 G Class Reefer","Day Cab","Sleeper"];

const LINE = {
  RL:  { bg:"#84cc16", text:"#1a2e05", label:"Ready Line" },
  WL:  { bg:"#7dd3fc", text:"#0c2a3e", label:"Wash Line" },
  SRL: { bg:"#f1f5f9", text:"#0f172a", label:"Service Ready" },
  SL:  { bg:"#f87171", text:"#3b0a0a", label:"Service Line" },
  SHOP:{ bg:"#e8b4bc", text:"#f9fafb", label:"Shop / Deadline" },
  PUR: { bg:"#a855f7", text:"#f5f3ff", label:"Purolator" },
};

// ── HELPERS ───────────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2,9);
const todayStr = () => new Date().toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric",year:"numeric"});
const todayKey = () => new Date().toISOString().slice(0,10);
const fmtDate  = d => { if(!d) return ""; return new Date(d+"T00:00:00").toLocaleDateString("en-US",{month:"short",day:"numeric"}); };
const fmtKey   = k => { if(!k) return ""; return new Date(k+"T00:00:00").toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"}); };
const twoWeeks = () => { const d=new Date(); d.setDate(d.getDate()+14); return d.toISOString().slice(0,10); };
const daysUntil = d => {
  if(!d) return null;
  const t=new Date(); t.setHours(0,0,0,0);
  return Math.round((new Date(d+"T00:00:00")-t)/(864e5));
};

const mkBoard = () => { const b={}; TRUCK_TYPES.forEach(t=>{b[t]=[];}); return b; };
const BLANK = () => ({ yard:mkBoard(), reso:mkBoard(), tomorrow:mkBoard(), pm:[], tasks:[], hikes:[], sent:[], checkins:[], pmScheduled:[], pmRows:[], pmInitialized:false, groundSwaps:[], groundNeeds:[], puroFleetRows:[], puroFleetOriginal:null, puroOrigBytes:null, puroRemovedUnits:[], contactLog:[], interBranch:[], handoffNotes:[], utilLog:[] });
// pmRows: PM checklist — starts empty, filled via upload or yard PM button
// pmInitialized: true once user has uploaded or added PM rows (prevents re-seeding)
// pmScheduled: [{unit,scheduledDate,swapRequired,swapUnit,customer}] for task generation

// ── PM DATA FROM EXCEL ────────────────────────────────────────────────────
// Parsed from uploaded PM schedule table (Belfield location)

// PM day logic:
// +8 to +1  = approaching PM window (yellow — get customer on notice)
// 0          = entering window TODAY
// -1 to -13  = inside PM window (orange — must book in soon)  
// -14        = hard deadline — DUE TODAY, must come in
// < -14      = past deadline — critical overdue
// Days until next PM:
// positive = days until window opens (approaching)
// 0        = window opens today
// -1 to -6 = inside window, book it in
// -7       = must act NOW (key threshold)
// -8 to -13= critical
// -14      = deadline — should NEVER be reached
// < -14    = past deadline, overdue
function urgencyColor(days){
  if(days < -14) return { bg:"#7f1d1d", text:"#fca5a5", label:"OVERDUE — PAST DEADLINE" };
  if(days===-14) return { bg:"#fee2e2", text:"#dc2626", label:"DEADLINE — DO NOT REACH" };
  if(days <  -7) return { bg:"#fff7ed", text:"#c2410c", label:"CRITICAL — GET IN NOW" };
  if(days === -7)return { bg:"#fef9c3", text:"#92400e", label:"-7 GET ON THIS NOW" };
  if(days <   0) return { bg:"#fefce8", text:"#ca8a04", label:"IN WINDOW — BOOK SOON" };
  if(days === 0) return { bg:"#fff7ed", text:"#c2410c", label:"WINDOW OPENS TODAY" };
  if(days <=  8) return { bg:"#f0fdf4", text:"#15803d", label:"APPROACHING — PLAN AHEAD" };
  return { bg:"#f3f4f6", text:"#6b7280", label:"UPCOMING" };
}

// ── PM TAB COMPONENT ──────────────────────────────────────────────────────
// PM rows live in main state (S.pmRows) — starts EMPTY, populated via Excel upload or yard PM button
function PMTab({ S, setS, notify, openModal }) {

  const [filterStatus, setFilterStatus] = useState("all");
  const [sortBy, setSortBy] = useState("days");
  const [expandedId, setExpandedId] = useState(null);
  const [uploadStatus, setUploadStatus] = useState(null);

  const rows = S.pmRows || [];
  const allYardUnits = Object.values(S.yard||{}).flat();
  function isUnitOnYard(unit){ return allYardUnits.some(u=>String(u.unit).trim()===String(unit).trim()); }

  // ── PASTE FROM EXCEL PARSER ──
  // Handles tab-separated text pasted directly from Excel
  // Columns: Owning Location | Renting Location | Product Line | Unit | PM Type | Customer | Next PM Due | Days until next PM | Due PM Defe | Comments
  const [pasteText, setPasteText] = useState("");
  const [pasteMode, setPasteMode] = useState(false);

  function parsePastedTable(text){
    const lines = text.split(/\r?\n/).filter(l=>l.trim());
    if(lines.length < 2) throw new Error("Need at least a header row and one data row");

    // Detect delimiter: tab (Excel paste) or comma (CSV)
    const delim = lines[0].includes('\t') ? '\t' : ',';

    function splitLine(line){
      if(delim === '\t') return line.split('\t').map(v=>v.trim());
      // CSV with quotes
      const result=[]; let cur=""; let inQ=false;
      for(let i=0;i<line.length;i++){
        if(line[i]==='"'){ inQ=!inQ; }
        else if(line[i]===','&&!inQ){ result.push(cur.trim()); cur=""; }
        else { cur+=line[i]; }
      }
      result.push(cur.trim());
      return result.map(v=>v.replace(/^"|"$/g,'').trim());
    }

    const headers = splitLine(lines[0]).map(h=>h.toLowerCase().replace(/[^a-z0-9 ]/g,' ').trim());

    function colIdx(...keys){
      for(const k of keys){
        const idx = headers.findIndex(h=>h.includes(k));
        if(idx >= 0) return idx;
      }
      return -1;
    }

    function getCol(row, ...keys){
      const idx = colIdx(...keys);
      return (idx >= 0 && row[idx] != null) ? row[idx].trim() : "";
    }

    // Map column positions from this specific Excel format
    const iOwning   = colIdx('owning');
    const iRenting  = colIdx('renting');
    const iUnit     = colIdx('unit');
    const iPMType   = colIdx('pm type','type');
    const iCustomer = colIdx('customer','cust');
    const iNextPM   = colIdx('next pm','pm due','due');
    const iDays     = colIdx('days until','days left','days');
    const iDefe     = colIdx('defe');
    const iComments = colIdx('comment');

    const parsed = [];

    for(let i=1; i<lines.length; i++){
      const row = splitLine(lines[i]);
      if(row.every(c=>!c)) continue;

      const unit = iUnit >= 0 ? row[iUnit]?.trim() : getCol(row,'unit');
      if(!unit || !/^\d{5,7}$/.test(unit.replace(/\s/g,''))) continue;

      const owning  = iOwning  >= 0 ? (row[iOwning]||"").trim()  : "";
      const renting = iRenting >= 0 ? (row[iRenting]||"").trim() : "";
      const pmType  = iPMType  >= 0 ? (row[iPMType]||"").trim()  : "";
      const customer= iCustomer>= 0 ? (row[iCustomer]||"").trim(): "";
      const nextPM  = iNextPM  >= 0 ? (row[iNextPM]||"").trim()  : "";
      const daysLeft= iDays    >= 0 ? parseInt(row[iDays])||0     : 0;
      const defeDays= iDefe    >= 0 ? (row[iDefe]||"").trim()     : "";
      const comments= iComments>= 0 ? (row[iComments]||"").trim() : "";

      // Only import units within actionable window
      if(daysLeft > 8) continue;

      // Determine if this is OUR truck (owning = Belfield)
      const isOurTruck = !owning || owning.toUpperCase().includes('BELFIELD') || owning.includes('262910');

      // Determine if truck is at another location
      const isAtOtherLocation = renting && !renting.toUpperCase().includes('BELFIELD') && !renting.includes('262910');
      const otherLocation = isAtOtherLocation ? renting.replace(/^\d+\s*-?\s*/,'').trim() : "";

      parsed.push({
        unit: unit.replace(/\s/g,''),
        pmType,
        customer,
        nextPM,
        daysLeft,
        defeDays,
        comments,
        owningLocation: owning,
        rentingLocation: renting,
        isOurTruck,
        isAtOtherLocation,
        otherLocation,
        swapRequired: true, // swap always needed for PM
      });
    }

    if(!parsed.length) throw new Error("No valid units found. Make sure you copied the full table including the header row.");
    return parsed;
  }

  // Parse CSV text into PM row objects (file upload fallback)
  function parseCSVtoPMRows(csv){
    return parsePastedTable(csv);
  }

  // Merge parsed rows — skip dupes, auto-create tasks for other-location trucks
  function mergePMRows(parsed){
    setS(s=>{
      const existing     = s.pmRows || [];
      const existingUnits= new Set(existing.map(r=>r.unit));
      let added=0, dupes=0, otherLoc=0;
      const newRows  = [...existing];
      const newTasks = [...(s.tasks||[])];

      for(const r of parsed){
        if(existingUnits.has(r.unit)){ dupes++; continue; }

        // Add PM row
        newRows.push({
          ...r,
          id: uid(),
          status: "pending",
          scheduledDate: "",
          swapUnit: "",
          notes: r.comments||"",
          location: r.rentingLocation||"",
          locationNotified: false,
          _prev: null,
        });
        existingUnits.add(r.unit);
        added++;

        // If truck is at another location -> auto-create recurring task
        if(r.isAtOtherLocation && r.otherLocation){
          otherLoc++;
          const taskText = `Message ${r.otherLocation} for PM update on unit #${r.unit} (PM ${r.daysLeft > 0 ? 'in '+r.daysLeft+'d' : Math.abs(r.daysLeft)+'d overdue'}) — follow up every 3 days`;
          const taskExists = newTasks.some(t=>t.unit===r.unit&&t.type==="pm-followup");
          if(!taskExists){
            newTasks.push({
              id: uid(),
              done: false,
              type: "pm-followup",
              unit: r.unit,
              location: r.otherLocation,
              text: taskText,
              repeatDays: 3,
              createdAt: new Date().toISOString(),
            });
          }
        }
      }

      // Deduplicate
      const seen = new Set();
      const deduped = newRows.filter(r=>{ if(seen.has(r.unit)) return false; seen.add(r.unit); return true; });

      const otherMsg = otherLoc > 0 ? ` · ${otherLoc} task${otherLoc!==1?'s':''} created for other-location units` : '';
      setUploadStatus({ok:true, msg:`✓ ${added} unit${added!==1?"s":""} imported · ${dupes} duplicate${dupes!==1?"s":""} skipped${otherMsg}`});
      return {...s, pmRows:deduped, tasks:newTasks, pmInitialized:true};
    });
  }

  function handlePaste(){
    if(!pasteText.trim()){ setUploadStatus({ok:false,msg:"Paste your Excel table first"}); return; }
    try{
      const parsed = parsePastedTable(pasteText);
      mergePMRows(parsed);
      setPasteText("");
      setPasteMode(false);
    } catch(err){
      setUploadStatus({ok:false, msg:"Parse error: "+err.message});
    }
  }


  function updateRow(id, patch){
    // If swapRequired is being turned ON, auto-increment ground needs for this unit's truck type
    if(patch.swapRequired===true){
      const row = (S.pmRows||[]).find(r=>r.id===id);
      if(row){
        setS(s=>{
          const tt = row.pmType||""; // use pmType as truck type hint, or skip if blank
          const updated = s.pmRows.map(r=>r.id===id?{...r,...patch}:r);
          // Find truck type from yard if available
          const yardUnit = Object.entries(s.yard||{}).find(([tt2,cards])=>cards.some(c=>String(c.unit)===String(row.unit)));
          const truckType = yardUnit ? yardUnit[0] : "";
          if(!truckType) return {...s, pmRows:updated};
          const existingNeeds = s.groundNeeds||[];
          const exists = existingNeeds.find(n=>n.tt===truckType);
          const newNeeds = exists
            ? existingNeeds.map(n=>n.tt===truckType?{...n,count:n.count+1}:n)
            : [...existingNeeds,{tt:truckType,count:1}];
          return {...s, pmRows:updated, groundNeeds:newNeeds};
        });
        return;
      }
    }
    // If swapRequired turned OFF, decrement
    if(patch.swapRequired===false){
      const row = (S.pmRows||[]).find(r=>r.id===id);
      if(row && row.swapRequired){
        setS(s=>{
          const updated = s.pmRows.map(r=>r.id===id?{...r,...patch}:r);
          const yardUnit = Object.entries(s.yard||{}).find(([tt2,cards])=>cards.some(c=>String(c.unit)===String(row.unit)));
          const truckType = yardUnit ? yardUnit[0] : "";
          if(!truckType) return {...s, pmRows:updated};
          const newNeeds = (s.groundNeeds||[])
            .map(n=>n.tt===truckType?{...n,count:Math.max(0,n.count-1)}:n)
            .filter(n=>n.count>0);
          return {...s, pmRows:updated, groundNeeds:newNeeds};
        });
        return;
      }
    }
    setS(s=>({...s, pmRows:s.pmRows.map(r=>r.id===id?{...r,...patch}:r)}));
  }

  function updateWithUndo(id, patch){
    setS(s=>({...s, pmRows:s.pmRows.map(r=>r.id===id?{...r,_prev:{status:r.status,scheduledDate:r.scheduledDate},...patch}:r)}));
  }

  function undoRow(id){
    setS(s=>({...s, pmRows:s.pmRows.map(r=>{
      if(r.id!==id||!r._prev) return r;
      return {...r,...r._prev,_prev:null};
    })}));
    notify("Action undone ↩");
  }

  function markScheduled(id, scheduledDate){
    if(!scheduledDate){ notify("Please pick a scheduled date first"); return; }
    const row = rows.find(r=>r.id===id);
    if(!row) return;
    updateWithUndo(id, { status:"scheduled", scheduledDate });
    // Sync to pmScheduled for newDay task generation
    setS(s=>{
      const existing=(s.pmScheduled||[]).filter(p=>p.unit!==row.unit);
      return {...s, pmScheduled:[...existing,{unit:row.unit,scheduledDate,swapRequired:row.swapRequired,swapUnit:row.swapUnit,customer:row.customer,pmType:row.pmType}]};
    });
    notify(`Unit ${row.unit} scheduled for ${fmtDate(scheduledDate)} ✓`);
  }

  function markDone(id){
    const row = rows.find(r=>r.id===id);
    updateWithUndo(id, { status:"done" });
    setS(s=>{
      // Remove from pmScheduled
      const pmScheduled=(s.pmScheduled||[]).filter(p=>p.unit!==row?.unit);
      // If unit is on yard and on SL/SHOP, auto-move to SRL (service done)
      let yard = s.yard;
      if(row){
        const updated = {};
        TRUCK_TYPES.forEach(tt=>{
          updated[tt]=(yard[tt]||[]).map(c=>{
            if(String(c.unit).trim()===String(row.unit).trim()&&(c.line==="SL"||c.line==="SHOP"))
              return {...c,line:"SRL"};
            return c;
          });
        });
        yard = updated;
      }
      return {...s, pmScheduled, yard};
    });
    if(row) notify("Unit "+row.unit+" PM done ✓ — moved to SRL on yard");
  }


  const filtered = rows
    .filter(r => filterStatus==="all" || r.status===filterStatus)
    .sort((a,b) => sortBy==="days" ? a.daysLeft-b.daysLeft : a.unit.localeCompare(b.unit));

  const critical    = rows.filter(r=>r.daysLeft<=-14&&r.status!=="done").length;
  const inWindow    = rows.filter(r=>r.daysLeft>-14&&r.daysLeft<0&&r.status!=="done").length;
  const approaching = rows.filter(r=>r.daysLeft>=0&&r.daysLeft<=8&&r.status!=="done").length;
  const overdue     = rows.filter(r=>r.daysLeft<0&&r.status!=="done").length;
  const urgent      = critical;
  const scheduled   = rows.filter(r=>r.status==="scheduled").length;
  const done        = rows.filter(r=>r.status==="done").length;

  return (
    <div>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:12,marginBottom:16}}>
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#fb923c",letterSpacing:"0.08em"}}>PM SCHEDULE</div>
          <div style={{fontSize:10,color:"#9c6b75",marginTop:1}}>Imported from Belfield PM table · set scheduled date · confirm done · undo any action</div>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {[["DUE TODAY / PAST",critical,"#ef4444","#fee2e2"],["IN WINDOW",inWindow,"#f97316","#fff7ed"],["APPROACHING",approaching,"#ca8a04","#fefce8"],["SCHEDULED",scheduled,"#16a34a","#dcfce7"],["DONE",done,"#9c6b75","#f3c0c8"]].map(([l,v,c,bg])=>(
            <div key={l} style={{background:bg,border:"1px solid "+c+"55",borderRadius:6,padding:"4px 10px",textAlign:"center"}}>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:c}}>{v}</div>
              <div style={{fontSize:9,color:c,opacity:0.9,letterSpacing:"0.06em"}}>{l}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Filter + Sort */}
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <span style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em"}}>Filter:</span>
        {[["all","All"],["pending","Pending"],["scheduled","Scheduled"],["done","Done"]].map(([v,l])=>(
          <button key={v} onClick={()=>setFilterStatus(v)} style={{border:"none",borderRadius:5,padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",background:filterStatus===v?"#f59e0b":"#f3c0c8",color:filterStatus===v?"#fdf2f4":"#a07880"}}>{l}</button>
        ))}
        <div style={{marginLeft:"auto",display:"flex",gap:6,alignItems:"center"}}>
          <span style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em"}}>Sort:</span>
          {[["days","By Date"],["unit","By Unit"]].map(([v,l])=>(
            <button key={v} onClick={()=>setSortBy(v)} style={{border:"none",borderRadius:5,padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",background:sortBy===v?"#e8b4bc":"#f3c0c8",color:sortBy===v?"#1a1a2e":"#a07880"}}>{l}</button>
          ))}
        </div>
      </div>

      {/* Checklist */}
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {filtered.map(row=>{
          const urg=urgencyColor(row.daysLeft);
          const isDone=row.status==="done";
          const isScheduled=row.status==="scheduled";
          const expanded=expandedId===row.id;
          const canUndo=!!row._prev;
          return (
            <div key={row.id} style={{background:"#ffffff",border:`1px solid ${isDone?"#f3c0c8":isScheduled?"#16a34a44":urg.bg}`,borderRadius:9,overflow:"hidden",opacity:isDone?0.6:1,transition:"all 0.2s"}}>

              {/* ── MAIN ROW ── */}
              <div style={{display:"flex",alignItems:"center",gap:10,padding:"11px 14px",cursor:"pointer",flexWrap:"wrap"}} onClick={()=>setExpandedId(expanded?null:row.id)}>

                {/* Status dot */}
                <div style={{width:10,height:10,borderRadius:"50%",background:isDone?"#4ade80":isScheduled?"#34d399":urg.bg,flexShrink:0,boxShadow:isScheduled?"0 0 6px #16a34a":undefined}}/>

                {/* Unit */}
                <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:isDone?"#e8b4bc":isScheduled?"#34d399":urg.text,minWidth:65,textDecoration:isDone?"line-through":undefined}}>{row.unit}</div>

                {/* PM type / service type */}
                {row.pmType&&(
                  <div style={{background:row.pmType==="WET"?"#eff6ff":row.pmType==="DRY"?"#fef9c3":row.pmType==="G1"?"#f0fdf4":"#f3c0c8",
                    color:row.pmType==="WET"?"#1d4ed8":row.pmType==="DRY"?"#92400e":row.pmType==="G1"?"#15803d":"#7a5560",
                    border:"1px solid "+(row.pmType==="WET"?"#bfdbfe":row.pmType==="DRY"?"#fde68a":row.pmType==="G1"?"#bbf7d0":"#f3c0c8"),
                    borderRadius:4,padding:"1px 7px",fontSize:10,fontWeight:700,flexShrink:0}}>
                    {row.pmType==="WET"?"💧 WET":row.pmType==="DRY"?"🌬 DRY":row.pmType==="G1"?"🪙 G1":row.pmType} SERVICE
                  </div>
                )}
                {/* Other location badge */}
                {row.isAtOtherLocation&&row.otherLocation&&(
                  <div style={{background:"#fdf4ff",color:"#7c3aed",border:"1px solid #d8b4fe",borderRadius:4,padding:"1px 7px",fontSize:9,fontWeight:700,flexShrink:0}}>
                    📍 {row.otherLocation}
                  </div>
                )}

                {/* ON YARD badge */}
                {isUnitOnYard(row.unit)&&!isDone&&(
                  <div style={{background:"#dcfce7",border:"1px solid #16a34a",borderRadius:4,padding:"1px 7px",fontSize:9,fontWeight:700,color:"#15803d",flexShrink:0}}>🟢 ON YARD</div>
                )}

                {/* Swap required badge */}
                {row.swapRequired&&!isDone&&(
                  <div style={{background:"#fff7ed",border:"1px solid #f97316",borderRadius:4,padding:"1px 7px",fontSize:9,fontWeight:700,color:"#c2410c",flexShrink:0}}>🔄 SWAP{row.swapUnit?" #"+row.swapUnit:""}</div>
                )}

                {/* Non-Belfield location badge */}
                {row.location&&row.location.toLowerCase().indexOf("belfield")===-1&&!isDone&&(
                  <div style={{background:"#fefce8",border:"1px solid #ca8a04",borderRadius:4,padding:"1px 7px",fontSize:9,fontWeight:700,color:"#92400e",flexShrink:0}}>
                    📍 {row.location}{row.locationNotified?" · Notified":""}
                  </div>
                )}

                {/* Customer */}
                <div style={{flex:1,fontSize:11,color:isDone?"#e8b4bc":"#a07880",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",minWidth:80}}>{row.customer}</div>

                {/* Days / status pill */}
                <div style={{background:isDone?"#f3c0c8":isScheduled?"#064e3b":urg.bg,color:isDone?"#4ade80":isScheduled?"#34d399":urg.text,borderRadius:5,padding:"2px 10px",fontSize:11,fontWeight:700,flexShrink:0,minWidth:88,textAlign:"center"}}>
                  {isDone?"✓ DONE":isScheduled?"📅 "+fmtDate(row.scheduledDate):row.daysLeft<=-14?"🔴 DUE NOW ("+Math.abs(row.daysLeft+14)+"d past)":row.daysLeft<0?"🟠 IN WINDOW ("+Math.abs(row.daysLeft)+"d / 14)":row.daysLeft===0?"🟡 WINDOW TODAY":"🟡 "+row.daysLeft+"d to window"}
                </div>

                {/* Next PM date */}
                <div style={{fontSize:10,color:"#e8b4bc",flexShrink:0,minWidth:55}}>{fmtDate(row.nextPM)}</div>

                {/* Action buttons */}
                <div style={{display:"flex",gap:5,flexShrink:0,alignItems:"center"}} onClick={e=>e.stopPropagation()}>
                  {canUndo&&(
                    <button onClick={()=>undoRow(row.id)} style={{background:"#f3c0c8",border:"1px solid #374151",color:"#6b4c52",borderRadius:5,padding:"4px 8px",fontSize:10,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}} title="Undo last action">
                      ↩ Undo
                    </button>
                  )}
                  {!isDone&&!isScheduled&&(
                    <button onClick={()=>markDone(row.id)} style={{background:"#dcfce7",border:"1px solid #16a34a",color:"#4ade80",borderRadius:5,padding:"4px 9px",fontSize:10,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                      ✓ Done
                    </button>
                  )}
                  {isDone&&(
                    <span style={{fontSize:10,color:"#4ade80",fontWeight:700}}>✅ Complete</span>
                  )}
                  {/* Delete button */}
                  <button
                    onClick={()=>setS(s=>({...s,pmRows:s.pmRows.filter(r=>r.id!==row.id),pmScheduled:(s.pmScheduled||[]).filter(p=>p.unit!==row.unit),pmInitialized:true}))}
                    style={{background:"transparent",border:"1px solid #374151",color:"#9c6b75",borderRadius:5,padding:"4px 7px",fontSize:11,fontWeight:700,cursor:"pointer",lineHeight:1}}
                    title="Delete this PM row">
                    ✕
                  </button>
                </div>

                <div style={{fontSize:10,color:"#e8b4bc",flexShrink:0}}>{expanded?"▲":"▼"}</div>
              </div>

              {/* ── EXPANDED PANEL ── */}
              {expanded&&(
                <div style={{borderTop:"1px solid #1f2937",padding:"14px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>

                  {/* Left: details */}
                  <div>
                    <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8}}>Details</div>
                    <div style={{fontSize:12,color:"#7a5560",lineHeight:2}}>
                      <div><span style={{color:"#9c6b75"}}>Customer: </span><span style={{color:"#6b4c52"}}>{row.customer}</span></div>
                      <div><span style={{color:"#9c6b75"}}>PM Type: </span><span style={{color:"#6b4c52"}}>{row.pmType}</span></div>
                      <div><span style={{color:"#9c6b75"}}>Due Date: </span><span style={{color:"#f59e0b"}}>{fmtDate(row.nextPM)}</span></div>
                      <div><span style={{color:"#9c6b75"}}>Deferral: </span><span style={{color:"#6b4c52"}}>{row.defeDays}</span></div>
                      {row.comment&&<div><span style={{color:"#9c6b75"}}>PM #: </span><span style={{color:"#6b4c52"}}>{row.comment}</span></div>}
                      {row.scheduledDate&&<div><span style={{color:"#9c6b75"}}>Scheduled: </span><span style={{color:"#34d399",fontWeight:700}}>{fmtDate(row.scheduledDate)}</span></div>}
                    </div>

                    {/* ON YARD banner */}
                    {isUnitOnYard(row.unit)&&!isDone&&(
                      <div style={{marginTop:8,background:"#dcfce7",border:"2px solid #16a34a",borderRadius:7,padding:"8px 12px"}}>
                        <div style={{fontSize:12,fontWeight:700,color:"#15803d"}}>🟢 This unit is currently ON YOUR YARD</div>
                        <div style={{fontSize:10,color:"#166534",marginTop:2}}>You can schedule the PM while it is here</div>
                      </div>
                    )}

                    {/* Other location banner */}
                    {row.isAtOtherLocation&&row.otherLocation&&!isDone&&(
                      <div style={{marginTop:8,background:"#fdf4ff",border:"2px solid #a855f7",borderRadius:7,padding:"10px 12px"}}>
                        <div style={{fontSize:12,fontWeight:700,color:"#7c3aed",marginBottom:4}}>📍 Unit at {row.otherLocation}</div>
                        <div style={{fontSize:10,color:"#6b21a8",marginBottom:6}}>A follow-up task has been auto-created to message this location every 3 days until the PM is scheduled.</div>
                        <div style={{fontSize:9,color:"#9c6b75"}}>Check the Tasks tab · type: PM Follow-up · repeat every 3 days</div>
                      </div>
                    )}

                    {/* Non-Belfield location panel */}
                    {row.location&&row.location.toLowerCase().indexOf("belfield")===-1&&!isDone&&(
                      <div style={{marginTop:8,background:"#fefce8",border:"2px solid #ca8a04",borderRadius:7,padding:"10px"}}>
                        <div style={{fontSize:11,fontWeight:700,color:"#92400e",marginBottom:6}}>📍 Unit at {row.location} — not at Belfield</div>
                        {!row.locationNotified?(
                          <button onClick={()=>updateWithUndo(row.id,{locationNotified:true,status:"scheduled",scheduledDate:row.scheduledDate||""})}
                            style={{width:"100%",background:"#fef08a",border:"1px solid #ca8a04",color:"#713f12",borderRadius:5,padding:"6px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                            📞 Mark Location Notified -> Move to Scheduled
                          </button>
                        ):(
                          <div>
                            <div style={{fontSize:10,color:"#15803d",fontWeight:700,marginBottom:6}}>✅ Location notified — waiting for drop-off date</div>
                            <input type="date" value={row.scheduledDate||""} onChange={e=>updateRow(row.id,{scheduledDate:e.target.value})}
                              style={{width:"100%",background:"#fff",border:"1px solid #ca8a04",color:"#1a1a2e",borderRadius:5,padding:"6px 10px",fontFamily:"inherit",fontSize:12,outline:"none",marginBottom:6}}/>
                            {row.scheduledDate&&<div style={{fontSize:10,color:"#92400e"}}>PM date set: {fmtDate(row.scheduledDate)}</div>}
                            <button onClick={()=>updateRow(row.id,{locationNotified:false})}
                              style={{marginTop:6,background:"transparent",border:"1px solid #f3c0c8",color:"#9c6b75",borderRadius:5,padding:"4px 8px",fontSize:10,cursor:"pointer",fontFamily:"inherit"}}>
                              ↩ Undo Notified
                            </button>
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  {/* Right: actions */}
                  <div style={{display:"flex",flexDirection:"column",gap:8}}>
                    <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:2}}>Actions</div>

                    {/* Schedule date picker + confirm */}
                    {!isDone&&(
                      <div style={{background:"#f3c0c8",borderRadius:7,padding:"10px"}}>
                        <div style={{fontSize:10,color:"#7a5560",marginBottom:5}}>📅 Schedule PM for date:</div>
                        <input
                          type="date"
                          value={row.scheduledDate||""}
                          onChange={e=>updateRow(row.id,{scheduledDate:e.target.value})}
                          style={{background:"#ffffff",border:"1px solid #374151",color:"#1a1a2e",borderRadius:5,padding:"6px 10px",fontFamily:"inherit",fontSize:12,width:"100%",outline:"none",marginBottom:6}}
                        />
                        <button
                          onClick={()=>markScheduled(row.id, row.scheduledDate)}
                          style={{width:"100%",background:isScheduled?"#064e3b":"#1e3a5f",border:`1px solid ${isScheduled?"#16a34a":"#2d5080"}`,color:isScheduled?"#4ade80":"#93c5fd",borderRadius:5,padding:"6px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                          {isScheduled?"✓ Scheduled — update date":"📧 Mark Scheduled + Email Sent"}
                        </button>
                      </div>
                    )}

                    {/* Swap required */}
                    <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",background:row.swapRequired?"#fff7ed":"#f9fafb",border:"1px solid "+(row.swapRequired?"#f97316":"#f3c0c8"),borderRadius:6,padding:"8px 10px"}}>
                      <input type="checkbox" checked={!!row.swapRequired} onChange={e=>updateRow(row.id,{swapRequired:e.target.checked})} style={{width:14,height:14,accentColor:"#f97316",cursor:"pointer"}}/>
                      <div>
                        <div style={{fontSize:12,color:row.swapRequired?"#c2410c":"#a07880",fontWeight:row.swapRequired?"700":"400"}}>🔄 Swap Required</div>
                        {row.swapRequired&&<div style={{fontSize:9,color:"#92400e",marginTop:1}}>A swap unit must be on yard before this PM can be done</div>}
                      </div>
                    </label>
                    {row.swapRequired&&(
                      <div>
                        <input placeholder="Assign swap unit # (links to Ground Units tab)..." value={row.swapUnit||""} onChange={e=>updateRow(row.id,{swapUnit:e.target.value})}
                          style={{background:"#fff",border:"1px solid #f97316",borderRadius:5,padding:"6px 10px",fontFamily:"inherit",fontSize:12,color:"#1a1a2e",outline:"none",width:"100%"}}/>
                        {row.swapUnit&&(
                          <div style={{fontSize:10,color:"#c2410c",marginTop:4,padding:"4px 8px",background:"#fff7ed",borderRadius:4}}>
                            Swap unit #{row.swapUnit} assigned — check Ground Units tab to track availability
                          </div>
                        )}
                      </div>
                    )}



                    {/* Mark done from expanded */}
                    {!isDone&&(
                      <button onClick={()=>markDone(row.id)} style={{background:"#dcfce7",border:"1px solid #16a34a",color:"#4ade80",borderRadius:5,padding:"6px 10px",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
                        ✅ Confirm PM Done
                      </button>
                    )}

                    {/* Undo from expanded */}
                    {canUndo&&(
                      <button onClick={()=>undoRow(row.id)} style={{background:"#f3c0c8",border:"1px solid #374151",color:"#6b4c52",borderRadius:5,padding:"6px 10px",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>
                        ↩ Undo Last Action
                      </button>
                    )}

                    {/* Notes */}
                    <textarea placeholder="Notes..." value={row.notes} onChange={e=>updateRow(row.id,{notes:e.target.value})}
                      style={{background:"#f3c0c8",border:"1px solid #374151",borderRadius:5,padding:"7px 10px",fontFamily:"inherit",fontSize:11,color:"#6b4c52",outline:"none",width:"100%",resize:"vertical",minHeight:56}}/>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* ── UPLOAD SECTION ── */}
      <div style={{marginTop:20,background:"#ffffff",border:"1px solid #f3c0c8",borderRadius:10,padding:"16px"}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:15,color:"#fb923c",letterSpacing:"0.06em",marginBottom:4}}>UPLOAD NEW PM TABLE</div>
        <div style={{fontSize:10,color:"#9c6b75",marginBottom:12}}>Upload an Excel (.xlsx), CSV (.csv), or screenshot image — new units are merged in, duplicates are ignored, existing statuses are preserved</div>

        {/* PASTE TABLE — primary input method */}
        {!pasteMode?(
          <button onClick={()=>setPasteMode(true)}
            style={{width:"100%",background:"#f59e0b",border:"none",borderRadius:9,color:"#fff",fontSize:13,fontWeight:700,padding:"14px",cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:8}}>
            📋 Paste PM Table from Excel
          </button>
        ):(
          <div>
            <div style={{fontSize:11,color:"#9c6b75",marginBottom:6}}>
              Copy the full table in Excel (Ctrl+A or select all rows including header) -> Ctrl+C -> paste below
            </div>
            <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:7,padding:"8px 12px",marginBottom:8,fontSize:10,color:"#92400e"}}>
              <strong>Expected columns:</strong> Owning Location · Renting Location · Unit · PM Type · Customer · Next PM Due · Days until next PM · Comments
            </div>
            <textarea
              value={pasteText}
              onChange={e=>setPasteText(e.target.value)}
              placeholder="Paste your Excel PM table here... Tip: Select all rows in Excel (including the header row) then Ctrl+C, click here, Ctrl+V"
              style={{width:"100%",minHeight:160,background:"#fff",border:"1.5px solid #f59e0b",borderRadius:8,padding:"10px",fontFamily:"monospace",fontSize:11,color:"#1a1a2e",outline:"none",resize:"vertical",lineHeight:1.5}}
            />
            <div style={{display:"flex",gap:8,marginTop:8}}>
              <button onClick={handlePaste}
                style={{flex:1,background:"#f59e0b",border:"none",borderRadius:7,color:"#fff",fontSize:13,fontWeight:700,padding:"10px",cursor:"pointer",fontFamily:"inherit"}}>
                ✓ Import Table
              </button>
              <button onClick={()=>{setPasteMode(false);setPasteText("");}}
                style={{background:"transparent",border:"1px solid #f3c0c8",borderRadius:7,color:"#9c6b75",fontSize:12,padding:"10px 16px",cursor:"pointer",fontFamily:"inherit"}}>
                Cancel
              </button>
            </div>
          </div>
        )}

        {/* File upload — fallback */}
        {!pasteMode&&(
          <div style={{marginTop:8}}>
            <div style={{fontSize:10,color:"#9c6b75",textAlign:"center",marginBottom:6}}>or upload a file</div>
            <label style={{display:"block",border:"2px dashed #f3c0c8",borderRadius:8,padding:"12px",textAlign:"center",cursor:"pointer",background:"#fdf2f4"}}
              onDragOver={e=>{e.preventDefault();e.currentTarget.style.borderColor="#f59e0b";}}
              onDragLeave={e=>{e.currentTarget.style.borderColor="#f3c0c8";}}
              onDrop={e=>{e.preventDefault();e.currentTarget.style.borderColor="#f3c0c8";handlePMUpload(e.dataTransfer.files[0]);}}>
              <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>handlePMUpload(e.target.files[0])}/>
              <div style={{fontSize:11,color:"#9c6b75"}}>📎 Drop Excel / CSV file here</div>
            </label>
          </div>
        )}

        {uploadStatus&&(
          <div style={{marginTop:10,padding:"9px 12px",background:uploadStatus.ok?"#f0fdf4":"#fff5f5",border:`1px solid ${uploadStatus.ok?"#16a34a":"#ef4444"}`,borderRadius:7,fontSize:11,color:uploadStatus.ok?"#15803d":"#dc2626",fontWeight:500}}>
            {uploadStatus.msg}
          </div>
        )}
      </div>
    </div>
  );

  // ── UPLOAD HANDLER ── defined inside PMTab scope so it can access state
  function handlePMUpload(file){
    if(!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    setUploadStatus({ok:true,msg:"Reading file..."});

    if(ext==="csv"){
      const reader = new FileReader();
      reader.onload = e => {
        try { mergePMRows(parseCSVtoPMRows(e.target.result)); }
        catch(err){ setUploadStatus({ok:false,msg:"CSV error: "+err.message}); }
      };
      reader.onerror = ()=>setUploadStatus({ok:false,msg:"Could not read file"});
      reader.readAsText(file);

    } else if(ext==="xlsx"||ext==="xls"){
      // Load SheetJS dynamically then parse
      const script = document.createElement('script');
      script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      script.onload = () => {
        const reader = new FileReader();
        reader.onload = e => {
          try {
            const wb = window.XLSX.read(new Uint8Array(e.target.result), {type:"array"});
            const ws = wb.Sheets[wb.SheetNames[0]];
            const csv = window.XLSX.utils.sheet_to_csv(ws);
            mergePMRows(parseCSVtoPMRows(csv));
          } catch(err){ setUploadStatus({ok:false,msg:"Excel error: "+err.message}); }
        };
        reader.readAsArrayBuffer(file);
      };
      script.onerror = ()=>setUploadStatus({ok:false,msg:"Could not load Excel parser. Please save your file as CSV from Excel (File -> Save As -> CSV) and upload that instead."});
      document.head.appendChild(script);

    } else if(["png","jpg","jpeg"].includes(ext)){
      setUploadStatus({ok:false,msg:"📸 Image detected — to import from a screenshot: open the image, manually note the unit numbers, then add them via the yard PM button (🔧 PM on each card). For auto-import, export your Excel as CSV (File -> Save As -> CSV) and upload that."});

    } else {
      setUploadStatus({ok:false,msg:"Please upload an Excel (.xlsx) or CSV (.csv) file"});
    }
  }
}

// ── GROUND UNITS TAB ─────────────────────────────────────────────────────
function GroundUnitsTab({ S, setS, notify, TRUCK_TYPES }) {
  const [swaps, setSwaps_] = useState(S.groundSwaps||[]);
  const [addForm, setAddForm] = useState(null);
  const [showHistory, setShowHistory] = useState(false);

  function setSwaps(fn){
    const n = typeof fn==="function" ? fn(swaps) : fn;
    setSwaps_(n);
    setS(s=>({...s, groundSwaps:n}));
  }

  // Auto-compute units needed per truck type from open swaps (no replacement assigned yet)
  const allYardFlat = TRUCK_TYPES.flatMap(tt=>(S.yard[tt]||[]).map(c=>({...c,tt})));
  function isOnYard(unit){ return allYardFlat.some(c=>String(c.unit).trim()===String(unit).trim()&&["RL","WL","SRL"].includes(c.line)); }

  const openSwaps   = swaps.filter(s=>s.swapStatus!=="done");
  const doneSwaps   = swaps.filter(s=>s.swapStatus==="done");

  // Units needed = open swaps with no replacement assigned, grouped by truck type
  const needsByType = {};
  openSwaps.forEach(s=>{
    if(!s.replacementUnit){
      needsByType[s.tt] = (needsByType[s.tt]||0) + 1;
    }
  });

  const REASONS = ["PM","E-Test","Plate Renewal","Damage","Customer Request","Unit Deadline","Other"];

  // E-Test flow is different: swap goes OUT first (customer gets it before their unit is tested)
  // PM / Plate Renewal / Damage: customer brings unit IN first, gets swap while work is done
  const FLOWS = {
    "E-Test": [
      {key:"needs_unit",      label:"Needs Swap Unit",   color:"#f59e0b", bg:"#fef9c3"},
      {key:"unit_assigned",   label:"Swap Assigned",     color:"#3b82f6", bg:"#eff6ff"},
      {key:"swap_done",       label:"Swap Given Out",    color:"#a855f7", bg:"#fdf4ff"},
      {key:"etest_progress",  label:"E-Test In Progress",color:"#0891b2", bg:"#ecfeff"},
      {key:"etest_done",      label:"E-Test Passed",     color:"#16a34a", bg:"#dcfce7"},
      {key:"done",            label:"Swap Returned",     color:"#9c6b75", bg:"#f3c0c8"},
    ],
    "default": [
      {key:"needs_unit",    label:"Needs Unit",      color:"#f59e0b", bg:"#fef9c3"},
      {key:"unit_assigned", label:"Unit Assigned",   color:"#3b82f6", bg:"#eff6ff"},
      {key:"swap_ready",    label:"Swap Ready",      color:"#a855f7", bg:"#fdf4ff"},
      {key:"swap_done",     label:"Swap Executed",   color:"#16a34a", bg:"#dcfce7"},
      {key:"done",          label:"Complete",        color:"#9c6b75", bg:"#f3c0c8"},
    ],
  };

  function getFlow(reason){ return FLOWS[reason] || FLOWS["default"]; }
  function getStatusInfo(key, reason){
    const flow = getFlow(reason);
    return flow.find(s=>s.key===key) || flow[0];
  }

  const DONE_LABELS = {
    "PM":            "✓ PM Done",
    "E-Test":        "✓ Swap Returned",
    "Plate Renewal": "✓ Plate Done",
    "Damage":        "✓ Repair Done",
    "Customer Request": "✓ Swap Done",
    "Unit Deadline": "✓ Unit Returned",
    "Other":         "✓ Done",
  };

  // Helper: human-readable description of what the next step means
  function getNextLabel(sw, nextStatus){
    if(!nextStatus) return "";
    if(sw.reason==="E-Test"){
      if(nextStatus.key==="unit_assigned")   return "-> Assign Swap Unit";
      if(nextStatus.key==="swap_done")       return "-> Swap Given to Customer";
      if(nextStatus.key==="etest_progress")  return "-> Unit In for E-Test";
      if(nextStatus.key==="etest_done")      return "-> E-Test Passed";
      if(nextStatus.key==="done")            return "✓ Swap Returned";
    }
    if(nextStatus.key==="done") return DONE_LABELS[sw.reason]||"✓ Done";
    return "-> "+nextStatus.label;
  }

  function daysOpen(sw){
    if(!sw.createdAt) return 0;
    return Math.round((Date.now()-new Date(sw.createdAt))/(864e5));
  }

  function addSwap(){
    if(!addForm?.currentUnit?.trim()||!addForm?.tt||!addForm?.reason){
      notify("Fill in unit #, truck type, and reason"); return;
    }
    const entry = {
      id:uid(),
      currentUnit:  addForm.currentUnit.trim(),
      tt:           addForm.tt,
      reason:       addForm.reason,
      customer:     addForm.customer?.trim()||"",
      replacementUnit: addForm.replacementUnit?.trim()||"",
      swapStatus:   addForm.replacementUnit?.trim() ? "unit_assigned" : "needs_unit",
      notes:        addForm.notes?.trim()||"",
      createdAt:    new Date().toISOString(),
      updatedAt:    new Date().toISOString(),
    };
    setSwaps(prev=>[entry,...prev]);
    setAddForm(null);
    notify("Swap added ✓");
  }

  function updateSwap(id, patch){
    setSwaps(prev=>prev.map(s=>s.id===id ? Object.assign({},s,patch,{updatedAt:new Date().toISOString()}) : s));
  }

  function advanceStatus(sw){
    const flow = getFlow(sw.reason);
    const idx = flow.findIndex(s=>s.key===sw.swapStatus);
    if(idx < flow.length-1){
      const next = flow[idx+1].key;
      updateSwap(sw.id, {swapStatus:next});
      if(next==="done") notify("Complete ✓ — moved to history");
      else notify("-> "+flow[idx+1].label);
    }
  }

  function deleteSwap(id){ setSwaps(prev=>prev.filter(s=>s.id!==id)); }

  return (
    <div>
      <style>{`@keyframes blink{0%,100%{opacity:1}50%{opacity:0.2}}`}</style>

      {/* ── HEADER ── */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:12,marginBottom:16}}>
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#f59e0b",letterSpacing:"0.08em"}}>🔄 GROUND UNITS — SWAP TRACKER</div>
          <div style={{fontSize:10,color:"#9c6b75",marginTop:1}}>Track unit swaps · PM · E-Test · Plate Renewal · status flows automatically</div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center"}}>
          {doneSwaps.length>0&&(
            <button onClick={()=>setShowHistory(h=>!h)} style={{background:"transparent",border:"1px solid #f3c0c8",borderRadius:6,color:"#9c6b75",fontSize:11,fontWeight:600,padding:"6px 12px",cursor:"pointer",fontFamily:"inherit"}}>
              {showHistory?"Hide":"📋"} History ({doneSwaps.length})
            </button>
          )}
          <button onClick={()=>setAddForm({currentUnit:"",tt:"",reason:"",customer:"",replacementUnit:"",notes:""})}
            style={{background:"#f59e0b",border:"none",borderRadius:6,color:"#fff",fontSize:11,fontWeight:700,padding:"7px 16px",cursor:"pointer",fontFamily:"inherit"}}>
            + Add Swap
          </button>
        </div>
      </div>

      {/* ── UNITS NEEDED — auto-computed ── */}
      {Object.keys(needsByType).length>0&&(
        <div style={{marginBottom:20}}>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:14,color:"#f59e0b",marginBottom:10,letterSpacing:"0.05em"}}>
            UNITS NEEDED (auto-calculated from open swaps without replacement)
          </div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            {Object.entries(needsByType).map(([tt,count])=>(
              <div key={tt} style={{background:"#fff",border:"2px solid #f59e0b",borderRadius:10,padding:"12px 20px",textAlign:"center",minWidth:120}}>
                <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:48,color:"#f59e0b",lineHeight:1}}>{count}</div>
                <div style={{fontSize:11,color:"#1a1a2e",fontWeight:600,marginTop:4}}>{tt}</div>
                <div style={{fontSize:9,color:"#f59e0b",marginTop:2}}>NEEDED</div>
              </div>
            ))}
          </div>
        </div>
      )}
      {Object.keys(needsByType).length===0&&openSwaps.length>0&&(
        <div style={{background:"#dcfce7",border:"1px solid #16a34a",borderRadius:8,padding:"10px 14px",marginBottom:16,fontSize:12,color:"#15803d",fontWeight:600}}>
          ✅ All open swaps have a replacement unit assigned
        </div>
      )}

      {/* ── ADD SWAP FORM ── */}
      {addForm&&(
        <div style={{background:"#fff",border:"2px solid #f59e0b",borderRadius:10,padding:"16px",marginBottom:16}}>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:14,color:"#f59e0b",marginBottom:12,letterSpacing:"0.06em"}}>NEW SWAP</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Current Unit #</div>
              <input placeholder="e.g. 515857" value={addForm.currentUnit} onChange={e=>setAddForm(f=>({...f,currentUnit:e.target.value}))}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Truck Type</div>
              <select value={addForm.tt} onChange={e=>setAddForm(f=>({...f,tt:e.target.value}))}>
                <option value="">Select...</option>
                {TRUCK_TYPES.map(t=><option key={t} value={t}>{t}</option>)}
              </select>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Reason</div>
              <select value={addForm.reason} onChange={e=>setAddForm(f=>({...f,reason:e.target.value}))}>
                <option value="">Select reason...</option>
                {REASONS.map(r=><option key={r} value={r}>{r}</option>)}
              </select>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Customer Name</div>
              <input placeholder="Who has this unit?" value={addForm.customer} onChange={e=>setAddForm(f=>({...f,customer:e.target.value}))}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Replacement Unit # (optional)</div>
              <input placeholder="Leave blank if not assigned yet" value={addForm.replacementUnit} onChange={e=>setAddForm(f=>({...f,replacementUnit:e.target.value}))}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Notes</div>
              <input placeholder="Any additional info..." value={addForm.notes} onChange={e=>setAddForm(f=>({...f,notes:e.target.value}))}/>
            </div>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={addSwap} style={{background:"#f59e0b",border:"none",borderRadius:6,color:"#fff",fontSize:12,fontWeight:700,padding:"8px 20px",cursor:"pointer",fontFamily:"inherit"}}>Add Swap</button>
            <button onClick={()=>setAddForm(null)} style={{background:"transparent",border:"1px solid #f3c0c8",borderRadius:6,color:"#9c6b75",fontSize:12,padding:"8px 14px",cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
          </div>
        </div>
      )}

      {/* ── OPEN SWAPS ── */}
      {openSwaps.length===0&&!addForm&&(
        <div style={{textAlign:"center",padding:"32px 0",color:"#e8b4bc",fontSize:12,background:"#fff",border:"1px solid #f3c0c8",borderRadius:10}}>
          No active swaps — hit + Add Swap to track one
        </div>
      )}

      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {openSwaps.map(sw=>{
          const status   = getStatusInfo(sw.swapStatus||"needs_unit", sw.reason);
          const days     = daysOpen(sw);
          const replOnYard = sw.replacementUnit && isOnYard(sw.replacementUnit);
          const currOnYard = isOnYard(sw.currentUnit);
          const isUrgent = days >= 5;
          const doneLabel = DONE_LABELS[sw.reason] || "Done";
          const flow = getFlow(sw.reason); const nextStatus = flow[flow.findIndex(s=>s.key===status.key)+1];

          return (
            <div key={sw.id} style={{background:"#fff",border:`2px solid ${isUrgent&&status.key==="needs_unit"?"#ef4444":status.color}44`,borderRadius:10,overflow:"hidden"}}>

              {/* Status bar */}
              <div style={{background:status.bg,borderBottom:`1px solid ${status.color}33`,padding:"7px 14px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <div style={{width:8,height:8,borderRadius:"50%",background:status.color,flexShrink:0}}/>
                  <span style={{fontSize:11,fontWeight:700,color:status.color}}>{status.label}</span>
                </div>
                {/* Status progress pills */}
                <div style={{display:"flex",gap:4,flex:1,flexWrap:"wrap"}}>
                  {getFlow(sw.reason).filter(s=>s.key!=="done").map((s,i)=>{
                    const idx     = getFlow(sw.reason).findIndex(x=>x.key===status.key);
                    const sIdx    = getFlow(sw.reason).findIndex(x=>x.key===s.key);
                    const isPast  = sIdx < idx;
                    const isCurr  = sIdx === idx;
                    return (
                      <div key={s.key} style={{display:"flex",alignItems:"center",gap:3}}>
                        <div style={{width:6,height:6,borderRadius:"50%",background:isPast||isCurr?s.color:"#e8b4bc"}}/>
                        <span style={{fontSize:9,color:isCurr?s.color:isPast?"#9c6b75":"#e8b4bc",fontWeight:isCurr?700:400}}>{s.label}</span>
                        {i<3&&<span style={{fontSize:9,color:"#e8b4bc",marginLeft:2}}>-></span>}
                      </div>
                    );
                  })}
                </div>
                {/* Days open badge */}
                <div style={{background:isUrgent?"#fee2e2":"#f3c0c8",border:`1px solid ${isUrgent?"#fca5a5":"#f3c0c8"}`,borderRadius:5,padding:"2px 8px",fontSize:9,fontWeight:700,color:isUrgent?"#dc2626":"#9c6b75",flexShrink:0}}>
                  {days===0?"Today":days+"d open"}{isUrgent?" ⚠":""}
                </div>
                {sw.reason==="E-Test"&&(
                  <div style={{fontSize:9,color:"#0891b2",fontWeight:600,background:"#ecfeff",border:"1px solid #a5f3fc",borderRadius:4,padding:"2px 8px",flexShrink:0}}>
                    🔬 Swap goes OUT first — E-Test happens before customer takes unit
                  </div>
                )}
              </div>

              {/* Main content */}
              <div style={{padding:"12px 14px",display:"flex",gap:16,flexWrap:"wrap",alignItems:"flex-start"}}>

                {/* Current unit */}
                <div style={{textAlign:"center",minWidth:80}}>
                  <div style={{fontSize:9,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Current</div>
                  <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,color:"#1a1a2e"}}>{sw.currentUnit}</div>
                  <div style={{fontSize:9,color:"#9c6b75"}}>{sw.tt}</div>
                  {currOnYard&&<div style={{fontSize:8,background:"#dcfce7",color:"#16a34a",borderRadius:3,padding:"1px 5px",fontWeight:700,marginTop:2}}>ON YARD</div>}
                </div>

                {/* Reason badge */}
                <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minWidth:90}}>
                  <div style={{background:
                    sw.reason==="PM"?"#fff7ed":
                    sw.reason==="E-Test"?"#eff6ff":
                    sw.reason==="Plate Renewal"?"#fdf4ff":
                    sw.reason==="Damage"?"#fff5f5":"#f3c0c8",
                    border:`1px solid ${
                    sw.reason==="PM"?"#f97316":
                    sw.reason==="E-Test"?"#3b82f6":
                    sw.reason==="Plate Renewal"?"#a855f7":
                    sw.reason==="Damage"?"#ef4444":"#f3c0c8"}`,
                    borderRadius:6,padding:"4px 10px",fontSize:11,fontWeight:700,
                    color:
                    sw.reason==="PM"?"#c2410c":
                    sw.reason==="E-Test"?"#2563eb":
                    sw.reason==="Plate Renewal"?"#7c3aed":
                    sw.reason==="Damage"?"#dc2626":"#9c6b75",
                    textAlign:"center"
                  }}>
                    {sw.reason==="PM"?"🔧":sw.reason==="E-Test"?"🔬":sw.reason==="Plate Renewal"?"🪪":sw.reason==="Damage"?"💥":"🔄"} {sw.reason}
                  </div>
                  {sw.customer&&<div style={{fontSize:9,color:"#6b4c52",marginTop:4,textAlign:"center"}}>{sw.customer}</div>}
                </div>

                <div style={{fontSize:18,color:"#e8b4bc",alignSelf:"center"}}>-></div>

                {/* Replacement unit */}
                <div style={{textAlign:"center",minWidth:80}}>
                  <div style={{fontSize:9,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Replacement</div>
                  {sw.replacementUnit?(
                    <div>
                      <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:24,color:"#16a34a"}}>{sw.replacementUnit}</div>
                      {replOnYard?(
                        <div style={{fontSize:8,background:"#dcfce7",color:"#16a34a",border:"1px solid #86efac",borderRadius:4,padding:"2px 6px",fontWeight:700,marginTop:2}}>✅ ON YARD</div>
                      ):(
                        <div style={{fontSize:8,background:"#fef9c3",color:"#ca8a04",border:"1px solid #fde68a",borderRadius:4,padding:"2px 6px",fontWeight:700,marginTop:2}}>NOT ON YARD</div>
                      )}
                    </div>
                  ):(
                    <div>
                      <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#f59e0b",animation:"blink 1.2s step-start infinite"}}>PENDING</div>
                      <input
                        placeholder="Assign #"
                        onKeyDown={e=>{if(e.key==="Enter"&&e.target.value.trim()){updateSwap(sw.id,{replacementUnit:e.target.value.trim(),swapStatus:"unit_assigned"});e.target.value="";notify("Unit assigned ✓");}}}
                        onBlur={e=>{if(e.target.value.trim()){updateSwap(sw.id,{replacementUnit:e.target.value.trim(),swapStatus:"unit_assigned"});e.target.value="";}}}
                        style={{marginTop:4,width:100,background:"#fef9c3",border:"1px solid #f59e0b",borderRadius:5,padding:"4px 8px",fontFamily:"inherit",fontSize:11,color:"#1a1a2e",outline:"none",textAlign:"center"}}
                      />
                    </div>
                  )}
                </div>

                {/* Notes */}
                {sw.notes&&(
                  <div style={{flex:1,fontSize:11,color:"#6b4c52",background:"#fdf2f4",borderRadius:6,padding:"6px 10px",alignSelf:"center",minWidth:120}}>
                    {sw.notes}
                  </div>
                )}

                {/* Actions */}
                <div style={{display:"flex",flexDirection:"column",gap:6,marginLeft:"auto",alignItems:"flex-end"}}>
                  {nextStatus&&(
                    <button onClick={()=>advanceStatus(sw)} style={{background:nextStatus.key==="done"?"#16a34a":"#f59e0b",border:"none",borderRadius:6,color:"#fff",fontSize:11,fontWeight:700,padding:"6px 14px",cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>
                      {getNextLabel(sw, nextStatus)}
                    </button>
                  )}
                  <div style={{display:"flex",gap:6}}>
                    {!sw._editingNote?(
                      <button onClick={()=>updateSwap(sw.id,{_editingNote:true})} style={{background:"transparent",border:"1px solid #f3c0c8",borderRadius:5,color:"#9c6b75",fontSize:10,padding:"4px 8px",cursor:"pointer",fontFamily:"inherit"}}>{sw.notes?"Edit Note":"+ Note"}</button>
                    ):(
                      <div style={{display:"flex",gap:4,alignItems:"center"}}>
                        <input autoFocus defaultValue={sw.notes||""} onKeyDown={e=>{if(e.key==="Enter"){updateSwap(sw.id,{notes:e.target.value,_editingNote:false});}if(e.key==="Escape")updateSwap(sw.id,{_editingNote:false});}}
                          style={{width:140,background:"#fff",border:"1px solid #f59e0b",borderRadius:5,padding:"3px 7px",fontFamily:"inherit",fontSize:11,color:"#1a1a2e",outline:"none"}}
                          placeholder="Type note, Enter to save"/>
                        <button onClick={e=>{const v=e.target.previousSibling.value;updateSwap(sw.id,{notes:v,_editingNote:false});}} style={{background:"#f59e0b",border:"none",borderRadius:4,color:"#fff",fontSize:10,fontWeight:700,padding:"3px 7px",cursor:"pointer",fontFamily:"inherit"}}>✓</button>
                      </div>
                    )}
                    <button onClick={()=>deleteSwap(sw.id)} style={{background:"transparent",border:"1px solid #fca5a5",borderRadius:5,color:"#dc2626",fontSize:10,padding:"4px 8px",cursor:"pointer",fontFamily:"inherit"}}>✕</button>
                  </div>
                </div>

              </div>
            </div>
          );
        })}
      </div>

      {/* ── COMPLETED HISTORY ── */}
      {showHistory&&doneSwaps.length>0&&(
        <div style={{marginTop:20}}>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:14,color:"#9c6b75",marginBottom:10,letterSpacing:"0.06em"}}>COMPLETED SWAPS</div>
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {doneSwaps.map(sw=>(
              <div key={sw.id} style={{background:"#f9fafb",border:"1px solid #f3c0c8",borderRadius:8,padding:"10px 14px",display:"flex",alignItems:"center",gap:14,opacity:0.8,flexWrap:"wrap"}}>
                <div>
                  <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:16,color:"#9c6b75"}}>{sw.currentUnit}</span>
                  <span style={{fontSize:12,color:"#e8b4bc",margin:"0 6px"}}>-></span>
                  <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:16,color:"#16a34a"}}>{sw.replacementUnit||"—"}</span>
                </div>
                <div style={{fontSize:10,color:"#9c6b75"}}>{sw.tt}</div>
                <div style={{fontSize:10,background:"#f3c0c8",color:"#9c6b75",borderRadius:4,padding:"1px 7px"}}>{sw.reason}</div>
                {sw.customer&&<div style={{fontSize:10,color:"#6b4c52"}}>{sw.customer}</div>}
                <div style={{fontSize:9,color:"#9c6b75",marginLeft:"auto"}}>{DONE_LABELS[sw.reason]||"Done"} · {daysOpen(sw)}d total</div>
                <button onClick={()=>deleteSwap(sw.id)} style={{background:"none",border:"none",color:"#e8b4bc",cursor:"pointer",fontSize:11}}>✕</button>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ── PUROLATOR FLEET SHEET TAB
function PuroFleetTab({ S, setS, notify }) {
  const [allRows, setAllRows_]   = useState(S.puroFleetRows||[]);  // ALL units from file
  const [editId, setEditId]      = useState(null);
  const [editVals, setEditVals]  = useState({});
  const [msg, setMsg]            = useState(null);
  const [fname, setFname]        = useState("");
  const [sec, setSec]            = useState("all");
  const [showAdd, setShowAdd]    = useState(false);
  const [addForm, setAddForm]    = useState({unit:"",location:"GTT",nextPM:"",status:"Out - Local",comments:"",section:"26 Van (NON-CDL)"});
  const [origBytes, setOrigBytes] = useState(S.puroOrigBytes||null);
  const [copied, setCopied]      = useState(false);

  function setAllRows(fn){
    const n = typeof fn==="function" ? fn(allRows) : fn;
    setAllRows_(n);
    setS(s=>({...s, puroFleetRows:n}));
  }

  const yardPuro = Object.values(S.yard||{}).flat().filter(c=>c.isPuro);
  function onYard(u){ return yardPuro.some(p=>String(p.unit).trim()===String(u).trim()); }

  // Only show units with a comment — the ones that need attention
  const rows = allRows.filter(function(r){ return r.comments && r.comments.trim(); });

  // PM units from PM tab that are Purolator (match any puro unit or isPuro on yard)
  const puroUnitsInPM = (S.pmRows||[]).filter(function(r){
    return r.status !== "done" && allRows.some(function(p){ return String(p.unit)===String(r.unit); });
  });

  // PM suggestions: puro units NOT already in our rows (no comment) but coming up for PM
  const pmSuggestions = puroUnitsInPM.filter(function(r){
    var alreadyHasComment = rows.some(function(p){ return String(p.unit)===String(r.unit); });
    return !alreadyHasComment;
  });

  // Puro units with "due service" already in comment
  const dueServiceKeywords = ["due service","due for service","service due","scheduled for service","in for service"];
  function hasDueService(comment){
    var c = (comment||"").toLowerCase();
    return dueServiceKeywords.some(function(k){ return c.indexOf(k)>-1; });
  }

  // ── UPLOAD ──
  function upload(file){
    if(!file) return;
    setFname(file.name);
    setMsg({ok:true, msg:"Reading file..."});
    var ext = file.name.split(".").pop().toLowerCase();

    function processBuffer(buf){
      try{
        var X = window.XLSX;
        var wb = X.read(new Uint8Array(buf), {type:"array"});
        var ws = wb.Sheets[wb.SheetNames[0]];
        var range = X.utils.decode_range(ws["!ref"]||"A1:O41");
        var parsed = [], section = "";

        function cv(r,c){
          var a = X.utils.encode_cell({r:r,c:c});
          var cell = ws[a];
          if(!cell||cell.v===null||cell.v===undefined) return "";
          return String(cell.v).trim();
        }

        for(var r=range.s.r; r<=range.e.r; r++){
          var e = cv(r,4);
          if(!e) continue;
          if(e.indexOf("Van")>-1)     { section="26 Van (NON-CDL)"; continue; }
          if(e.indexOf("Tractor")>-1) { section="T/A Tractor (CDL)"; continue; }
          if(e==="Purolator Location"||e.indexOf("PUROLATOR")>-1||e.indexOf("BELFIELD")>-1) continue;
          var clean = e.replace(/\s/g,"");
          if(!/^\d{5,7}$/.test(clean)) continue;
          parsed.push({
            id:uid(), unit:clean,
            location:cv(r,5), nextPM:cv(r,6), status:cv(r,7), comments:cv(r,8),
            section:section, _modified:false, _orig:cv(r,8), _isNew:false,
          });
        }

        if(!parsed.length){ setMsg({ok:false,msg:"No units found — check this is the Purolator fleet sheet"}); return; }
        setAllRows(parsed);
        var withComment = parsed.filter(function(r){ return r.comments&&r.comments.trim(); }).length;
        var m = parsed.filter(function(r){ return onYard(r.unit); }).length;
        setMsg({ok:true, msg:"Loaded "+parsed.length+" units total · "+withComment+" have comments (shown below) · "+m+" on yard"});
      }catch(err){ setMsg({ok:false,msg:"Excel error: "+err.message}); }
    }

    function load(){
      var reader = new FileReader();
      reader.onload = function(e){ processBuffer(e.target.result); };
      reader.readAsArrayBuffer(file);
    }

    if(ext==="xlsx"||ext==="xls"){
      if(window.XLSX){ load(); }
      else{
        var s=document.createElement("script");
        s.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
        s.onload=load;
        s.onerror=function(){ setMsg({ok:false,msg:"Could not load Excel library"}); };
        document.head.appendChild(s);
      }
    } else {
      setMsg({ok:false,msg:"Please upload the .xlsx master sheet"});
    }
  }

  // ── EDIT ──
  function startEdit(row){ setEditId(row.id); setEditVals({location:row.location,nextPM:row.nextPM,status:row.status,comments:row.comments}); }
  function saveEdit(id){
    setAllRows(function(prev){
      return prev.map(function(r){
        if(r.id!==id) return r;
        var changed = editVals.comments!==r._orig||editVals.status!==r.status||editVals.nextPM!==r.nextPM||editVals.location!==r.location;
        return Object.assign({},r,editVals,{_modified:changed});
      });
    });
    setEditId(null);
    notify("Row updated");
  }

  // ── ADD UNIT FROM PM SUGGESTION ──
  function addFromPM(pmRow){
    if(allRows.find(function(r){ return r.unit===pmRow.unit; })){
      // Unit exists but has no comment — add "Due for Service" comment
      setAllRows(function(prev){
        return prev.map(function(r){
          if(r.unit!==pmRow.unit) return r;
          return Object.assign({},r,{comments:"Due for Service - PM "+urgencyColor(pmRow.daysLeft).label,_modified:true});
        });
      });
      notify("Unit "+pmRow.unit+" flagged — Due for Service added");
    } else {
      // New unit entirely
      var newRow = {id:uid(),unit:pmRow.unit,location:"GTT",nextPM:pmRow.nextPM||"",status:"Out - Local",comments:"Due for Service - PM "+urgencyColor(pmRow.daysLeft).label,section:"26 Van (NON-CDL)",_modified:true,_orig:"",_isNew:true};
      setAllRows(function(prev){ return [...prev,newRow]; });
      notify("Unit "+pmRow.unit+" added with PM comment");
    }
  }

  // ── ADD MANUAL UNIT ──
  function addUnit(){
    if(!addForm.unit.trim()){ notify("Enter a unit number"); return; }
    if(allRows.find(function(r){ return r.unit===addForm.unit.trim(); })){ notify("Unit already in list"); return; }
    setAllRows(function(prev){
      return [...prev,{id:uid(),unit:addForm.unit.trim(),location:addForm.location,nextPM:addForm.nextPM,status:addForm.status||"Out - Local",comments:addForm.comments,section:addForm.section,_modified:true,_orig:"",_isNew:true}];
    });
    setAddForm({unit:"",location:"GTT",nextPM:"",status:"Out - Local",comments:"",section:"26 Van (NON-CDL)"});
    setShowAdd(false);
    notify("Unit added");
  }

  // ── REMOVE ──
  function removeUnit(id){
    setAllRows(function(prev){ return prev.filter(function(r){ return r.id!==id; }); });
    notify("Unit removed");
  }

  // ── COPY-PASTABLE OUTPUT ──
  // Only outputs rows that have comments (the ones that matter)
  function copyOutput(){
    var outputRows = rows.filter(function(r){ return r.comments&&r.comments.trim(); });
    if(!outputRows.length){ notify("No units with comments to copy"); return; }
    var header = "Unit #\tLocation\tNext PM Date\tStatus\tComments\tSection";
    var body = outputRows.map(function(r){
      return [r.unit,r.location,r.nextPM,r.status,r.comments,r.section].join("\t");
    }).join("\n");
    var text = header+"\n"+body;
    // Reliable cross-browser copy: create a textarea, select it, execCommand
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    ta.style.top = "0";
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    try{
      document.execCommand("copy");
      setCopied(true);
      setTimeout(function(){ setCopied(false); },2500);
      notify("Copied "+outputRows.length+" rows — paste into Excel");
    }catch(e){
      notify("Copy failed — please select the table text manually");
    }
    document.body.removeChild(ta);
  }

  var sections = [...new Set(rows.map(function(r){ return r.section; }).filter(Boolean))];
  var displayed = sec==="all" ? rows : rows.filter(function(r){ return r.section===sec; });
  var matched   = rows.filter(function(r){ return onYard(r.unit); });

  return (
    <div>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:12,marginBottom:16}}>
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#a855f7",letterSpacing:"0.08em"}}>PUROLATOR FLEET UPDATE</div>
          <div style={{fontSize:10,color:"#9c6b75",marginTop:1}}>
            Upload master sheet · only units with comments shown · copy output to paste back into master sheet
          </div>
        </div>
        {rows.length>0&&(
          <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
            <div style={{textAlign:"center",background:"#fff",border:"1px solid #a855f7",borderRadius:7,padding:"5px 12px"}}>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#a855f7"}}>{rows.length}</div>
              <div style={{fontSize:9,color:"#9c6b75",textTransform:"uppercase"}}>With Comments</div>
            </div>
            <div style={{textAlign:"center",background:"#fff",border:"1px solid #f59e0b",borderRadius:7,padding:"5px 12px"}}>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#f59e0b"}}>{rows.filter(function(r){return r._modified;}).length}</div>
              <div style={{fontSize:9,color:"#9c6b75",textTransform:"uppercase"}}>Modified</div>
            </div>
            <div style={{textAlign:"center",background:"#fff",border:"1px solid #16a34a",borderRadius:7,padding:"5px 12px"}}>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#16a34a"}}>{matched.length}</div>
              <div style={{fontSize:9,color:"#9c6b75",textTransform:"uppercase"}}>On Yard</div>
            </div>
            <button onClick={function(){setShowAdd(function(v){return !v;});}} style={{background:"#16a34a",border:"none",borderRadius:6,color:"#fff",fontSize:11,fontWeight:700,padding:"8px 14px",cursor:"pointer",fontFamily:"inherit"}}>+ Add Unit</button>
            <button onClick={copyOutput} style={{background:copied?"#16a34a":"#7c3aed",border:"none",borderRadius:6,color:"#fff",fontSize:11,fontWeight:700,padding:"8px 16px",cursor:"pointer",fontFamily:"inherit"}}>
              {copied?"✓ Copied!":"📋 Copy Output"}
            </button>
          </div>
        )}
      </div>

      {/* Upload */}
      <label style={{display:"block",border:"2px dashed #d8b4fe",borderRadius:10,padding:"18px",textAlign:"center",cursor:"pointer",background:"#fdf4ff",marginBottom:12}}>
        <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={function(e){upload(e.target.files[0]);}}/>
        <div style={{fontSize:20,marginBottom:3}}>🟣</div>
        <div style={{fontSize:12,color:"#a855f7",fontWeight:600}}>{allRows.length>0?"Re-upload Master Sheet":"Upload Purolator Master Fleet Sheet (.xlsx)"}</div>
        <div style={{fontSize:10,color:"#7c3aed",marginTop:2}}>Only units with comments will be shown · all others = no change needed</div>
      </label>

      {msg&&(
        <div style={{marginBottom:12,padding:"9px 13px",background:msg.ok?"#f0fdf4":"#fff5f5",border:"1px solid "+(msg.ok?"#16a34a":"#ef4444"),borderRadius:7,fontSize:11,color:msg.ok?"#15803d":"#dc2626"}}>
          {msg.msg}
        </div>
      )}

      {/* ── PM SUGGESTIONS PANEL ── */}
      {pmSuggestions.length>0&&(
        <div style={{background:"#fff7ed",border:"2px solid #f97316",borderRadius:10,padding:"14px 16px",marginBottom:16}}>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:14,color:"#c2410c",marginBottom:8,letterSpacing:"0.06em"}}>
            🔧 PUROLATOR UNITS WITH PM DUE — NOT YET ON THIS SHEET
          </div>
          <div style={{fontSize:10,color:"#92400e",marginBottom:10}}>
            These Purolator units are tracked in your PM tab but have no comment in the fleet sheet yet. Add them so Purolator knows they need service.
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {pmSuggestions.map(function(r){
              var urg = urgencyColor(r.daysLeft);
              var existsInAll = allRows.some(function(p){ return p.unit===r.unit; });
              return (
                <div key={r.id} style={{background:"#fff",border:"1.5px solid #fed7aa",borderRadius:8,padding:"10px 12px",minWidth:160}}>
                  <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#c2410c"}}>#{r.unit}</div>
                  <div style={{fontSize:9,color:"#92400e",marginBottom:4}}>{r.customer}</div>
                  <div style={{display:"inline-block",background:urg.bg,color:urg.text,borderRadius:4,padding:"1px 7px",fontSize:9,fontWeight:700,marginBottom:8}}>{urg.label} · {r.daysLeft}d</div>
                  <div>
                    <button onClick={function(){addFromPM(r);}} style={{width:"100%",background:"#f97316",border:"none",borderRadius:5,color:"#fff",fontSize:10,fontWeight:700,padding:"5px",cursor:"pointer",fontFamily:"inherit"}}>
                      {existsInAll?"+ Add Comment":"+ Add to Sheet"}
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ── PM UNITS ALREADY WITH DUE SERVICE COMMENT ── */}
      {puroUnitsInPM.filter(function(r){ return rows.some(function(p){ return p.unit===r.unit&&hasDueService(p.comments); }); }).length>0&&(
        <div style={{background:"#f0fdf4",border:"1.5px solid #86efac",borderRadius:9,padding:"10px 14px",marginBottom:12}}>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:12,color:"#15803d",marginBottom:6,letterSpacing:"0.05em"}}>✅ ALREADY FLAGGED IN SHEET</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {puroUnitsInPM.filter(function(r){ return rows.some(function(p){ return p.unit===r.unit&&hasDueService(p.comments); }); }).map(function(r){
              return (
                <div key={r.id} style={{background:"#fff",border:"1px solid #86efac",borderRadius:6,padding:"4px 10px",fontSize:11}}>
                  <span style={{fontFamily:"'Bebas Neue',sans-serif",color:"#15803d"}}>#{r.unit}</span>
                  <span style={{color:"#9c6b75",marginLeft:5,fontSize:9}}>Due Service already in comments ✓</span>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Add unit form */}
      {showAdd&&(
        <div style={{background:"#f0fdf4",border:"2px solid #16a34a",borderRadius:10,padding:"14px",marginBottom:14}}>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:13,color:"#15803d",marginBottom:10,letterSpacing:"0.06em"}}>ADD UNIT</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Unit #</div>
              <input placeholder="e.g. 515857" value={addForm.unit} onChange={function(e){var v=e.target.value;setAddForm(function(f){return Object.assign({},f,{unit:v});});}}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Location</div>
              <select value={addForm.location} onChange={function(e){var v=e.target.value;setAddForm(function(f){return Object.assign({},f,{location:v});});}}>
                <option value="GTT">GTT</option>
                <option value="Vulcan">Vulcan</option>
              </select>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Section</div>
              <select value={addForm.section} onChange={function(e){var v=e.target.value;setAddForm(function(f){return Object.assign({},f,{section:v});});}}>
                <option value="26 Van (NON-CDL)">26 Van (NON-CDL)</option>
                <option value="T/A Tractor (CDL)">T/A Tractor (CDL)</option>
              </select>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Status</div>
              <input value={addForm.status} onChange={function(e){var v=e.target.value;setAddForm(function(f){return Object.assign({},f,{status:v});});}}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Next PM Date</div>
              <input placeholder="e.g. Apr/29/2026" value={addForm.nextPM} onChange={function(e){var v=e.target.value;setAddForm(function(f){return Object.assign({},f,{nextPM:v});});}}/>
            </div>
            <div>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Comments</div>
              <input placeholder="What's the update?" value={addForm.comments} onChange={function(e){var v=e.target.value;setAddForm(function(f){return Object.assign({},f,{comments:v});});}}/>
            </div>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={addUnit} style={{background:"#16a34a",border:"none",borderRadius:6,color:"#fff",fontSize:12,fontWeight:700,padding:"8px 20px",cursor:"pointer",fontFamily:"inherit"}}>Add Unit</button>
            <button onClick={function(){setShowAdd(false);}} style={{background:"transparent",border:"1px solid #f3c0c8",borderRadius:6,color:"#9c6b75",fontSize:12,padding:"8px 14px",cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
          </div>
        </div>
      )}

      {/* Section filter */}
      {sections.length>0&&(
        <div style={{display:"flex",gap:6,marginBottom:10,flexWrap:"wrap"}}>
          {["all",...sections].map(function(s){
            return (
              <button key={s} onClick={function(){setSec(s);}}
                style={{border:"none",borderRadius:5,padding:"4px 12px",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",background:sec===s?"#a855f7":"#fff",color:sec===s?"#fff":"#9c6b75"}}>
                {s==="all"?"All Units with Comments":s}
              </button>
            );
          })}
        </div>
      )}

      {/* ── EXCEL-STYLE TABLE ── */}
      {rows.length>0&&(
        <div style={{background:"#fff",border:"2px solid #d8b4fe",borderRadius:10,overflow:"hidden",marginBottom:16}}>

          {/* Table toolbar */}
          <div style={{background:"#7c3aed",padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:15,color:"#fff",letterSpacing:"0.06em"}}>
              FLEET UPDATE TABLE — {rows.length} UNITS WITH COMMENTS
            </div>
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <div style={{fontSize:10,color:"#e9d5ff"}}>Select &amp; copy any cell · or use button -></div>
              <button onClick={copyOutput} style={{background:copied?"#16a34a":"#fff",border:"none",borderRadius:6,color:copied?"#fff":"#7c3aed",fontSize:11,fontWeight:700,padding:"6px 14px",cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
                {copied?"✓ Copied!":"📋 Copy All Rows"}
              </button>
            </div>
          </div>

          {/* Sticky header row — matches Excel column order */}
          <div style={{display:"grid",gridTemplateColumns:"90px 90px 100px 110px 1fr 120px 60px",background:"#ede9fe",borderBottom:"2px solid #a855f7",padding:"0"}}>
            {["Unit #","Location","Next PM","Status","Comments","Section",""].map(function(h,i){
              return (
                <div key={i} style={{padding:"7px 10px",fontSize:10,fontWeight:700,color:"#5b21b6",textTransform:"uppercase",letterSpacing:"0.06em",borderRight:i<6?"1px solid #c4b5fd":"none"}}>
                  {h}
                </div>
              );
            })}
          </div>

          {/* Data rows */}
          {displayed.map(function(row,i){
            var yard       = onYard(row.unit);
            var hasPM      = puroUnitsInPM.some(function(r){ return r.unit===row.unit; });
            var isDueSvc   = hasDueService(row.comments);
            var editing    = editId===row.id;
            var rowBg      = row._isNew?"#f0fdf4":row._modified?"#fefce8":yard?"#faf5ff":"#fff";
            var altBg      = row._isNew?"#f0fdf4":row._modified?"#fef9c3":yard?"#f5f3ff":i%2===1?"#fdf4ff":"#fff";

            if(editing){
              return (
                <div key={row.id} style={{background:"#fdf4ff",borderBottom:"1px solid #d8b4fe",padding:"10px 12px"}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                    <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:15,color:"#7c3aed"}}>Editing #{row.unit}</span>
                    <button onClick={function(){saveEdit(row.id);}} style={{background:"#7c3aed",border:"none",borderRadius:5,color:"#fff",fontSize:11,fontWeight:700,padding:"4px 14px",cursor:"pointer",fontFamily:"inherit",marginLeft:"auto"}}>Save</button>
                    <button onClick={function(){setEditId(null);}} style={{background:"transparent",border:"1px solid #d8b4fe",borderRadius:5,color:"#9c6b75",fontSize:11,padding:"4px 10px",cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 2fr",gap:10}}>
                    {[["location","Location"],["nextPM","Next PM Date"],["status","Status"],["comments","Comments"]].map(function(kl){
                      return (
                        <div key={kl[0]}>
                          <div style={{fontSize:9,color:"#7c3aed",textTransform:"uppercase",marginBottom:3,fontWeight:600}}>{kl[1]}</div>
                          <input value={editVals[kl[0]]||""} onChange={function(e){var v=e.target.value;setEditVals(function(f){return Object.assign({},f,{[kl[0]]:v});});}}
                            style={{width:"100%",background:"#fff",border:"1.5px solid #a855f7",color:"#1a1a2e",borderRadius:5,padding:"6px 8px",fontFamily:"inherit",fontSize:12,outline:"none"}}/>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            }

            return (
              <div key={row.id+"v"} style={{display:"grid",gridTemplateColumns:"90px 90px 100px 110px 1fr 120px 60px",borderBottom:i<displayed.length-1?"1px solid #ede9fe":"none",background:i%2===0?rowBg:altBg,alignItems:"stretch"}}>

                {/* Unit # cell */}
                <div style={{padding:"8px 10px",borderRight:"1px solid #ede9fe",display:"flex",flexDirection:"column",justifyContent:"center"}}>
                  <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:15,color:yard?"#7c3aed":row._isNew?"#16a34a":"#1a1a2e",userSelect:"text"}}>{row.unit}</div>
                  {yard&&<div style={{fontSize:7,background:"#ede9fe",color:"#7c3aed",borderRadius:3,padding:"1px 4px",fontWeight:700,marginTop:2,display:"inline-block"}}>ON YARD</div>}
                  {row._isNew&&<div style={{fontSize:7,background:"#dcfce7",color:"#16a34a",borderRadius:3,padding:"1px 4px",fontWeight:700,marginTop:2,display:"inline-block"}}>NEW</div>}
                  {row._modified&&!row._isNew&&<div style={{fontSize:7,background:"#fef9c3",color:"#ca8a04",borderRadius:3,padding:"1px 4px",fontWeight:700,marginTop:2,display:"inline-block"}}>EDITED</div>}
                </div>

                {/* Location */}
                <div style={{padding:"8px 10px",borderRight:"1px solid #ede9fe",fontSize:12,color:"#374151",display:"flex",alignItems:"center",userSelect:"text"}}>{row.location}</div>

                {/* Next PM */}
                <div style={{padding:"8px 10px",borderRight:"1px solid #ede9fe",fontSize:11,color:"#d97706",display:"flex",alignItems:"center",userSelect:"text"}}>{row.nextPM||"—"}</div>

                {/* Status */}
                <div style={{padding:"8px 10px",borderRight:"1px solid #ede9fe",display:"flex",alignItems:"center"}}>
                  <span style={{fontSize:11,background:"#f0f4ff",color:"#3730a3",border:"1px solid #c7d2fe",borderRadius:4,padding:"2px 7px",userSelect:"text"}}>{row.status}</span>
                </div>

                {/* Comments — the main column */}
                <div style={{padding:"8px 10px",borderRight:"1px solid #ede9fe",display:"flex",flexDirection:"column",justifyContent:"center",gap:3}}>
                  <div style={{fontSize:12,color:isDueSvc?"#c2410c":hasPM?"#7c3aed":row._modified?"#b45309":"#1a1a2e",fontWeight:isDueSvc||hasPM?600:400,userSelect:"text",lineHeight:1.4}}>
                    {row.comments||<span style={{color:"#d1d5db",fontStyle:"italic"}}>no comment</span>}
                  </div>
                  {isDueSvc&&<div style={{fontSize:9,color:"#16a34a",fontWeight:700}}>✓ Due for Service flagged</div>}
                  {hasPM&&!isDueSvc&&<div style={{fontSize:9,color:"#f97316",fontWeight:700}}>⚠ PM due — no service comment yet</div>}
                </div>

                {/* Section */}
                <div style={{padding:"8px 10px",borderRight:"1px solid #ede9fe",fontSize:10,color:"#6b7280",display:"flex",alignItems:"center",userSelect:"text"}}>{row.section}</div>

                {/* Actions */}
                <div style={{padding:"6px 8px",display:"flex",flexDirection:"column",gap:4,alignItems:"center",justifyContent:"center"}}>
                  <button onClick={function(){startEdit(row);}} style={{background:"#ede9fe",border:"none",borderRadius:4,color:"#7c3aed",fontSize:9,fontWeight:700,padding:"3px 8px",cursor:"pointer",fontFamily:"inherit",width:"100%"}}>Edit</button>
                  <button onClick={function(){removeUnit(row.id);}} style={{background:"#fee2e2",border:"none",borderRadius:4,color:"#dc2626",fontSize:9,fontWeight:700,padding:"3px 8px",cursor:"pointer",fontFamily:"inherit",width:"100%"}}>✕</button>
                </div>

              </div>
            );
          })}

          {/* Footer summary */}
          <div style={{background:"#ede9fe",padding:"7px 14px",display:"flex",gap:16,fontSize:10,color:"#5b21b6",fontWeight:600,borderTop:"2px solid #a855f7"}}>
            <span>Total: {rows.length}</span>
            <span style={{color:"#16a34a"}}>New: {rows.filter(function(r){return r._isNew;}).length}</span>
            <span style={{color:"#ca8a04"}}>Edited: {rows.filter(function(r){return r._modified&&!r._isNew;}).length}</span>
            <span style={{color:"#c2410c"}}>Due Service: {rows.filter(function(r){return hasDueService(r.comments);}).length}</span>
            <span style={{color:"#7c3aed"}}>On Yard: {matched.length}</span>
          </div>
        </div>
      )}

      {allRows.length>0&&rows.length===0&&(
        <div style={{textAlign:"center",padding:"32px 0",background:"#fff",border:"1.5px solid #f3c0c8",borderRadius:10}}>
          <div style={{fontSize:24,marginBottom:8}}>✅</div>
          <div style={{fontSize:13,color:"#15803d",fontWeight:600}}>All {allRows.length} units loaded — none have comments</div>
          <div style={{fontSize:11,color:"#9c6b75",marginTop:4}}>No updates needed for Purolator. Units with comments will appear here when you add them.</div>
        </div>
      )}

      {allRows.length===0&&!msg&&(
        <div style={{textAlign:"center",padding:"40px 0",color:"#d8b4fe",fontSize:12}}>Upload the Purolator master fleet sheet to get started</div>
      )}
    </div>
  );
}

// ── CONTACTS & INTER-BRANCH TAB ────────────────────────────────────────────
function ContactsTab({ S, setS, notify, dayLabel }) {
  const [view, setView]         = useState("log");   // "log" | "branch" | "handoff" | "digest"
  const [logForm, setLogForm]   = useState({ unit:"", customer:"", type:"call", outcome:"", note:"", branch:"" });
  const [branchForm, setBF]     = useState({ unit:"", tt:"", branch:"", since:"", expectedBack:"", note:"" });
  const [handoff, setHandoff]   = useState("");
  const [digestCopied, setDigestCopied] = useState(false);

  const log       = S.contactLog   || [];
  const branches  = S.interBranch  || [];
  const handoffs  = S.handoffNotes || [];

  function addLog() {
    if(!logForm.unit.trim() && !logForm.customer.trim()) { notify("Enter a unit or customer name"); return; }
    const entry = { id:uid(), ts:new Date().toISOString(), unit:logForm.unit.trim(), customer:logForm.customer.trim(), type:logForm.type, outcome:logForm.outcome.trim(), note:logForm.note.trim(), branch:logForm.branch.trim(), day:dayLabel };
    setS(s=>({...s, contactLog:[entry,...(s.contactLog||[])]}));
    setLogForm({ unit:"", customer:"", type:"call", outcome:"", note:"", branch:"" });
    notify("Contact logged ✓");
  }

  function delLog(id) { setS(s=>({...s, contactLog:(s.contactLog||[]).filter(x=>x.id!==id)})); }

  function addBranch() {
    if(!branchForm.unit.trim()||!branchForm.branch.trim()) { notify("Enter unit # and branch name"); return; }
    const entry = { id:uid(), unit:branchForm.unit.trim(), tt:branchForm.tt, branch:branchForm.branch.trim(), since:branchForm.since||new Date().toISOString().slice(0,10), expectedBack:branchForm.expectedBack, note:branchForm.note.trim(), active:true };
    setS(s=>({...s, interBranch:[entry,...(s.interBranch||[])]}));
    setBF({ unit:"", tt:"", branch:"", since:"", expectedBack:"", note:"" });
    notify("Inter-branch unit tracked ✓");
  }

  function returnBranch(id) {
    setS(s=>({...s, interBranch:(s.interBranch||[]).map(x=>x.id===id?{...x,active:false,returnedOn:new Date().toISOString().slice(0,10)}:x)}));
    notify("Marked as returned ✓");
  }

  function saveHandoff() {
    if(!handoff.trim()) return;
    const entry = { id:uid(), ts:new Date().toISOString(), day:dayLabel, note:handoff.trim() };
    setS(s=>({...s, handoffNotes:[entry,...(s.handoffNotes||[]).slice(0,13)]}));
    setHandoff("");
    notify("Handoff note saved ✓");
  }

  function buildDigest() {
    const allYard = TRUCK_TYPES.flatMap(tt=>(S.yard[tt]||[]).map(c=>({...c,tt})));
    const pmDue   = (S.pmRows||[]).filter(r=>r.status!=="done").slice(0,10);
    const overReso = TRUCK_TYPES.flatMap(tt=>(S.reso[tt]||[]).map(c=>({...c,tt}))).filter(c=>daysUntil(c.returnDate)<=0);
    const goingOut = allYard.filter(c=>c.goingOut);
    const activeBr = (S.interBranch||[]).filter(x=>x.active);

    let d = "=== BRANCH OPS DAILY DIGEST — "+dayLabel+" ===\n\n";
    d += "YARD: "+allYard.length+" units | Available: "+allYard.filter(c=>["RL","WL","SRL"].includes(c.line)&&!c.goingOut).length+"\n";
    if(goingOut.length) d += "Going Out: "+goingOut.map(c=>"#"+c.unit+" ("+c.tt+")").join(", ")+"\n";
    d += "\n";

    if(pmDue.length) {
      d += "PM UNITS NEEDING ATTENTION:\n";
      pmDue.forEach(r=>{ d += "  #"+r.unit+" — "+r.daysLeft+"d ("+urgencyColor(r.daysLeft).label+") | "+r.customer+"\n"; });
      d += "\n";
    }

    if(overReso.length) {
      d += "OVERDUE RETURNS ("+overReso.length+"):\n";
      overReso.forEach(c=>{ d += "  #"+c.unit+" ("+c.tt+") — overdue "+(Math.abs(daysUntil(c.returnDate)))+"d"+( c.customer?" | "+c.customer:"")+"\n"; });
      d += "\n";
    }

    const todaysLog = log.filter(x=>x.day===dayLabel);
    if(todaysLog.length) {
      d += "CUSTOMER CONTACTS TODAY ("+todaysLog.length+"):\n";
      todaysLog.forEach(x=>{ d += "  "+x.type.toUpperCase()+" — "+(x.unit?"#"+x.unit+" ":"")+x.customer+" | "+x.outcome+"\n"; });
      d += "\n";
    }

    if(activeBr.length) {
      d += "UNITS AT OTHER BRANCHES ("+activeBr.length+"):\n";
      activeBr.forEach(x=>{ d += "  #"+x.unit+" -> "+x.branch+(x.expectedBack?" | Back: "+x.expectedBack:"")+"\n"; });
      d += "\n";
    }

    const lastHandoff = handoffs[0];
    if(lastHandoff) {
      d += "LAST HANDOFF NOTE ("+lastHandoff.day+"):\n"+lastHandoff.note+"\n\n";
    }

    d += "Tasks: "+(S.tasks||[]).filter(t=>t.done).length+"/"+(S.tasks||[]).length+" done\n";
    return d;
  }

  function copyDigest() {
    const text = buildDigest();
    navigator.clipboard.writeText(text).then(()=>{ setDigestCopied(true); setTimeout(()=>setDigestCopied(false),2000); notify("Digest copied to clipboard ✓"); });
  }

  function emailDigest() {
    const text = buildDigest();
    const subject = encodeURIComponent("Branch Ops Digest — "+dayLabel);
    const body    = encodeURIComponent(text);
    window.open("mailto:?subject="+subject+"&body="+body);
  }

  const VIEWS = [["log","📞 Contact Log"],["branch","🏢 Inter-Branch"],["handoff","📝 Handoff Notes"],["digest","📧 Email Digest"]];

  return (
    <div>
      <div style={{display:"flex",gap:6,marginBottom:16,flexWrap:"wrap"}}>
        {VIEWS.map(([v,l])=>(
          <button key={v} onClick={()=>setView(v)} style={{border:"none",borderRadius:7,padding:"6px 14px",fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit",background:view===v?"#e11d48":"#fff",color:view===v?"#fff":"#9c6b75"}}>
            {l}
          </button>
        ))}
      </div>

      {/* ── CONTACT LOG ── */}
      {view==="log"&&(
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#e11d48",marginBottom:12,letterSpacing:"0.06em"}}>CUSTOMER CONTACT LOG</div>

          {/* Add form */}
          <div style={{background:"#fff",border:"1.5px solid #f3c0c8",borderRadius:10,padding:"14px",marginBottom:16}}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:13,color:"#9c6b75",marginBottom:10,letterSpacing:"0.05em"}}>LOG A CONTACT</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Unit # (optional)</div>
                <input placeholder="e.g. 515857" value={logForm.unit} onChange={e=>setLogForm(f=>({...f,unit:e.target.value}))}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Customer / Contact</div>
                <input placeholder="Name or company" value={logForm.customer} onChange={e=>setLogForm(f=>({...f,customer:e.target.value}))}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Type</div>
                <select value={logForm.type} onChange={e=>setLogForm(f=>({...f,type:e.target.value}))}>
                  <option value="call">📞 Call</option>
                  <option value="email">📧 Email</option>
                  <option value="in-person">🤝 In Person</option>
                  <option value="voicemail">📬 Voicemail</option>
                  <option value="text">💬 Text</option>
                </select>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Outcome</div>
                <input placeholder="e.g. Confirmed drop-off Friday" value={logForm.outcome} onChange={e=>setLogForm(f=>({...f,outcome:e.target.value}))}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Notes</div>
                <input placeholder="Additional details..." value={logForm.note} onChange={e=>setLogForm(f=>({...f,note:e.target.value}))}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Branch (if inter-branch)</div>
                <input placeholder="e.g. Mississauga" value={logForm.branch} onChange={e=>setLogForm(f=>({...f,branch:e.target.value}))}/>
              </div>
            </div>
            <button onClick={addLog} style={{background:"#e11d48",border:"none",borderRadius:7,color:"#fff",fontSize:12,fontWeight:700,padding:"8px 20px",cursor:"pointer",fontFamily:"inherit"}}>Log Contact</button>
          </div>

          {/* Log list */}
          {log.length===0&&<div style={{textAlign:"center",padding:"32px 0",color:"#e8b4bc",fontSize:12}}>No contacts logged yet</div>}
          {log.map(x=>(
            <div key={x.id} style={{background:"#fff",border:"1.5px solid #f3c0c8",borderRadius:9,padding:"10px 14px",marginBottom:8,display:"flex",gap:12,alignItems:"flex-start"}}>
              <div style={{flexShrink:0,width:36,height:36,borderRadius:"50%",background:x.type==="call"?"#fee2e2":x.type==="email"?"#eff6ff":x.type==="in-person"?"#dcfce7":"#fef9c3",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>
                {x.type==="call"?"📞":x.type==="email"?"📧":x.type==="in-person"?"🤝":x.type==="voicemail"?"📬":"💬"}
              </div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                  {x.unit&&<span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:16,color:"#e11d48"}}>#{x.unit}</span>}
                  <span style={{fontSize:13,fontWeight:600,color:"#1a1a2e"}}>{x.customer}</span>
                  {x.branch&&<span style={{fontSize:10,background:"#eff6ff",color:"#2563eb",borderRadius:4,padding:"1px 7px",fontWeight:600}}>🏢 {x.branch}</span>}
                  <span style={{fontSize:10,color:"#9c6b75",marginLeft:"auto"}}>{new Date(x.ts).toLocaleString("en-US",{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"})}</span>
                </div>
                {x.outcome&&<div style={{fontSize:12,color:"#15803d",marginTop:3,fontWeight:500}}>-> {x.outcome}</div>}
                {x.note&&<div style={{fontSize:11,color:"#9c6b75",marginTop:2}}>{x.note}</div>}
              </div>
              <button onClick={()=>delLog(x.id)} style={{background:"none",border:"none",color:"#f3c0c8",cursor:"pointer",fontSize:12,flexShrink:0}}>✕</button>
            </div>
          ))}
        </div>
      )}

      {/* ── INTER-BRANCH ── */}
      {view==="branch"&&(
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#2563eb",marginBottom:4,letterSpacing:"0.06em"}}>INTER-BRANCH UNIT TRACKER</div>
          <div style={{fontSize:10,color:"#9c6b75",marginBottom:14}}>Track units running at other branches — or units they sent you</div>

          <div style={{background:"#fff",border:"1.5px solid #bfdbfe",borderRadius:10,padding:"14px",marginBottom:16}}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:13,color:"#2563eb",marginBottom:10,letterSpacing:"0.05em"}}>ADD INTER-BRANCH UNIT</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Unit #</div>
                <input placeholder="e.g. 515857" value={branchForm.unit} onChange={e=>setBF(f=>({...f,unit:e.target.value}))}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Branch Name</div>
                <input placeholder="e.g. Mississauga, GTT" value={branchForm.branch} onChange={e=>setBF(f=>({...f,branch:e.target.value}))}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Truck Type</div>
                <select value={branchForm.tt} onChange={e=>setBF(f=>({...f,tt:e.target.value}))}>
                  <option value="">Select...</option>
                  {TRUCK_TYPES.map(t=><option key={t} value={t}>{t}</option>)}
                </select>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Since Date</div>
                <input type="date" value={branchForm.since} onChange={e=>setBF(f=>({...f,since:e.target.value}))}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Expected Back</div>
                <input type="date" value={branchForm.expectedBack} onChange={e=>setBF(f=>({...f,expectedBack:e.target.value}))}/>
              </div>
              <div>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:3}}>Notes</div>
                <input placeholder="Reason / contact at branch..." value={branchForm.note} onChange={e=>setBF(f=>({...f,note:e.target.value}))}/>
              </div>
            </div>
            <button onClick={addBranch} style={{background:"#2563eb",border:"none",borderRadius:7,color:"#fff",fontSize:12,fontWeight:700,padding:"8px 20px",cursor:"pointer",fontFamily:"inherit"}}>Track Unit</button>
          </div>

          {/* Active */}
          {branches.filter(x=>x.active).length===0&&<div style={{textAlign:"center",padding:"32px 0",color:"#e8b4bc",fontSize:12}}>No inter-branch units tracked</div>}
          {branches.filter(x=>x.active).map(x=>{
            const daysOut = x.since ? Math.round((new Date()-new Date(x.since+"T00:00:00"))/(864e5)) : 0;
            const daysBack = x.expectedBack ? daysUntil(x.expectedBack) : null;
            return (
              <div key={x.id} style={{background:"#fff",border:"1.5px solid #bfdbfe",borderRadius:9,padding:"12px 14px",marginBottom:8}}>
                <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                  <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#2563eb"}}>#{x.unit}</div>
                  <div style={{flex:1}}>
                    <div style={{fontSize:13,fontWeight:600,color:"#1a1a2e"}}>🏢 {x.branch}</div>
                    <div style={{fontSize:10,color:"#9c6b75"}}>{x.tt} · Out {daysOut} day{daysOut!==1?"s":""}{x.note?" · "+x.note:""}</div>
                  </div>
                  {daysBack!==null&&(
                    <div style={{textAlign:"center",background:daysBack<0?"#fee2e2":daysBack===0?"#fef9c3":"#f0fdf4",border:"1px solid "+(daysBack<0?"#fca5a5":daysBack===0?"#fde68a":"#bbf7d0"),borderRadius:6,padding:"3px 10px"}}>
                      <div style={{fontSize:12,fontWeight:700,color:daysBack<0?"#dc2626":daysBack===0?"#ca8a04":"#16a34a"}}>{daysBack<0?Math.abs(daysBack)+"d overdue":daysBack===0?"Due today":daysBack+"d left"}</div>
                      <div style={{fontSize:9,color:"#9c6b75"}}>return</div>
                    </div>
                  )}
                  <button onClick={()=>returnBranch(x.id)} style={{background:"#dcfce7",border:"1px solid #16a34a",borderRadius:6,color:"#15803d",fontSize:11,fontWeight:700,padding:"5px 12px",cursor:"pointer",fontFamily:"inherit"}}>✓ Returned</button>
                </div>
              </div>
            );
          })}

          {/* History */}
          {branches.filter(x=>!x.active).length>0&&(
            <div style={{marginTop:16}}>
              <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8}}>Returned Units (History)</div>
              {branches.filter(x=>!x.active).map(x=>(
                <div key={x.id} style={{background:"#f9fafb",border:"1px solid #f3c0c8",borderRadius:7,padding:"8px 12px",marginBottom:6,display:"flex",gap:10,alignItems:"center",opacity:0.7}}>
                  <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:15,color:"#9c6b75"}}>#{x.unit}</span>
                  <span style={{fontSize:11,color:"#9c6b75"}}>{x.branch} · returned {x.returnedOn||""}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* ── HANDOFF NOTES ── */}
      {view==="handoff"&&(
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#7c3aed",marginBottom:4,letterSpacing:"0.06em"}}>SHIFT HANDOFF NOTES</div>
          <div style={{fontSize:10,color:"#9c6b75",marginBottom:14}}>Write your end-of-day note — saves with today's date and carries into daily history</div>

          <div style={{background:"#fff",border:"1.5px solid #d8b4fe",borderRadius:10,padding:"14px",marginBottom:16}}>
            <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:6}}>Today — {dayLabel}</div>
            <textarea
              value={handoff}
              onChange={e=>setHandoff(e.target.value)}
              placeholder="What happened today? Tip: units to watch, customers to follow up, any issues, things next shift needs to know"
              style={{width:"100%",minHeight:140,background:"#fdf4ff",border:"1.5px solid #d8b4fe",borderRadius:8,padding:"10px",fontFamily:"inherit",fontSize:13,color:"#1a1a2e",outline:"none",resize:"vertical",lineHeight:1.6}}
            />
            <div style={{display:"flex",gap:8,marginTop:10}}>
              <button onClick={saveHandoff} style={{background:"#7c3aed",border:"none",borderRadius:7,color:"#fff",fontSize:12,fontWeight:700,padding:"8px 20px",cursor:"pointer",fontFamily:"inherit"}}>Save Handoff Note</button>
            </div>
          </div>

          {handoffs.length===0&&<div style={{textAlign:"center",padding:"24px 0",color:"#e8b4bc",fontSize:12}}>No handoff notes saved yet</div>}
          {handoffs.map((h,i)=>(
            <div key={h.id} style={{background:"#fff",border:"1.5px solid "+(i===0?"#d8b4fe":"#f3c0c8"),borderRadius:9,padding:"12px 14px",marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                <div style={{fontSize:11,fontWeight:600,color:i===0?"#7c3aed":"#9c6b75"}}>{h.day}{i===0?" (latest)":""}</div>
                <div style={{fontSize:10,color:"#9c6b75"}}>{new Date(h.ts).toLocaleTimeString("en-US",{hour:"2-digit",minute:"2-digit"})}</div>
              </div>
              <div style={{fontSize:13,color:"#1a1a2e",whiteSpace:"pre-wrap",lineHeight:1.6}}>{h.note}</div>
            </div>
          ))}
        </div>
      )}

      {/* ── EMAIL DIGEST ── */}
      {view==="digest"&&(
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#16a34a",marginBottom:4,letterSpacing:"0.06em"}}>DAILY EMAIL DIGEST</div>
          <div style={{fontSize:10,color:"#9c6b75",marginBottom:14}}>One-click summary of everything — PM units, overdue returns, contacts, inter-branch, handoff notes</div>

          <div style={{background:"#f0fdf4",border:"1.5px solid #86efac",borderRadius:10,padding:"14px",marginBottom:16}}>
            <pre style={{fontFamily:"monospace",fontSize:11,color:"#1a1a2e",whiteSpace:"pre-wrap",lineHeight:1.7,maxHeight:360,overflowY:"auto"}}>{buildDigest()}</pre>
          </div>

          <div style={{display:"flex",gap:10}}>
            <button onClick={copyDigest} style={{background:digestCopied?"#16a34a":"#fff",border:"1.5px solid "+(digestCopied?"#16a34a":"#f3c0c8"),borderRadius:7,color:digestCopied?"#fff":"#6b4c52",fontSize:12,fontWeight:700,padding:"9px 20px",cursor:"pointer",fontFamily:"inherit"}}>
              {digestCopied?"✓ Copied!":"📋 Copy to Clipboard"}
            </button>
            <button onClick={emailDigest} style={{background:"#e11d48",border:"none",borderRadius:7,color:"#fff",fontSize:12,fontWeight:700,padding:"9px 20px",cursor:"pointer",fontFamily:"inherit"}}>
              📧 Open in Email
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ── UTILIZATION TAB ────────────────────────────────────────────────────────
function UtilizationTab({ S, history, dayLabel, notify }) {
  const allYard = TRUCK_TYPES.flatMap(tt=>(S.yard[tt]||[]).map(c=>({...c,tt})));
  const allReso = TRUCK_TYPES.flatMap(tt=>(S.reso[tt]||[]).map(c=>({...c,tt})));

  // Compute utilization per truck type
  const typeStats = TRUCK_TYPES.map(tt=>{
    const yard = (S.yard[tt]||[]);
    const reso  = (S.reso[tt]||[]);
    const total = yard.length + reso.length;
    const out   = reso.length;
    const avail = yard.filter(c=>["RL","WL","SRL"].includes(c.line)&&!c.goingOut&&!c.isPuro).length;
    const shop  = yard.filter(c=>c.line==="SL"||c.line==="SHOP").length;
    const pct   = total>0 ? Math.round((out/total)*100) : 0;
    return { tt, total, out, avail, shop, pct };
  }).filter(x=>x.total>0);

  // History trend — units out per day
  const trend = [...history].reverse().slice(-14).map(h=>{
    const outCount = TRUCK_TYPES.reduce((n,tt)=>(n+(h.snap.reso[tt]||[]).length),0);
    const yardCount = TRUCK_TYPES.reduce((n,tt)=>(n+(h.snap.yard[tt]||[]).length),0);
    return { day:"Day "+h.dayNum, label:h.label.split(",")[0], out:outCount, yard:yardCount };
  });

  // Idle units (on yard >0 days, not going out)
  const idleUnits = allYard.filter(c=>!c.goingOut&&!c.awaitingArrival&&(c.line==="RL"||c.line==="WL"));

  // PM pressure: units on yard with PM due
  const pmOnYard = (S.pmRows||[]).filter(r=>r.status!=="done"&&allYard.some(u=>String(u.unit)===String(r.unit)));

  const totalFleet = allYard.length + allReso.length;
  const utilRate   = totalFleet>0 ? Math.round((allReso.length/totalFleet)*100) : 0;

  return (
    <div>
      <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#1a1a2e",marginBottom:4,letterSpacing:"0.08em"}}>📊 UTILIZATION &amp; FLEET HEALTH</div>
      <div style={{fontSize:10,color:"#9c6b75",marginBottom:18}}>Live snapshot of fleet usage — updated as you move units</div>

      {/* Top stats */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(130px,1fr))",gap:10,marginBottom:20}}>
        {[
          ["Fleet Util. Rate",utilRate+"%","#7c3aed","overall utilization"],
          ["Total On Reso",allReso.length,"#f59e0b","units currently out"],
          ["Available",allYard.filter(c=>["RL","WL","SRL"].includes(c.line)&&!c.isPuro).length,"#16a34a","RL + WL + SRL"],
          ["In Shop/SL",allYard.filter(c=>c.line==="SL"||c.line==="SHOP").length,"#ef4444","needing service"],
          ["PM Pressure",pmOnYard.length,"#fb923c","PM due on yard now"],
        ].map(([l,v,c,sub])=>(
          <div key={l} style={{background:"#fff",border:"1.5px solid "+c+"44",borderRadius:10,padding:"12px 14px",textAlign:"center"}}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:28,color:c,lineHeight:1}}>{v}</div>
            <div style={{fontSize:11,color:"#1a1a2e",fontWeight:600,marginTop:3}}>{l}</div>
            <div style={{fontSize:9,color:"#9c6b75",marginTop:1}}>{sub}</div>
          </div>
        ))}
      </div>

      {/* Per truck type breakdown */}
      <div style={{background:"#fff",border:"1.5px solid #f3c0c8",borderRadius:10,overflow:"hidden",marginBottom:20}}>
        <div style={{background:"#fff0f3",padding:"10px 16px",fontFamily:"'Bebas Neue',sans-serif",fontSize:14,color:"#e11d48",letterSpacing:"0.06em"}}>BY TRUCK TYPE</div>
        <div style={{display:"grid",gridTemplateColumns:"2fr 60px 60px 60px 60px 80px",padding:"7px 14px",borderBottom:"1px solid #f3c0c8",fontSize:9,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em"}}>
          <div>Type</div><div style={{textAlign:"center"}}>Total</div><div style={{textAlign:"center"}}>Out</div><div style={{textAlign:"center"}}>Avail</div><div style={{textAlign:"center"}}>Shop</div><div style={{textAlign:"center"}}>Util %</div>
        </div>
        {typeStats.length===0&&<div style={{padding:"16px",textAlign:"center",color:"#e8b4bc",fontSize:12}}>No units tracked yet</div>}
        {typeStats.map(x=>(
          <div key={x.tt} style={{display:"grid",gridTemplateColumns:"2fr 60px 60px 60px 60px 80px",padding:"9px 14px",borderBottom:"1px solid #fdf2f4",alignItems:"center"}}>
            <div style={{fontSize:12,color:"#1a1a2e",fontWeight:500}}>{x.tt}</div>
            <div style={{textAlign:"center",fontSize:12,color:"#9c6b75"}}>{x.total}</div>
            <div style={{textAlign:"center",fontSize:12,color:"#f59e0b",fontWeight:600}}>{x.out}</div>
            <div style={{textAlign:"center",fontSize:12,color:"#16a34a",fontWeight:600}}>{x.avail}</div>
            <div style={{textAlign:"center",fontSize:12,color:x.shop>0?"#ef4444":"#9c6b75"}}>{x.shop}</div>
            <div style={{textAlign:"center"}}>
              <div style={{display:"inline-block",background:x.pct>=80?"#dcfce7":x.pct>=50?"#fef9c3":"#fee2e2",border:"1px solid "+(x.pct>=80?"#86efac":x.pct>=50?"#fde68a":"#fca5a5"),borderRadius:5,padding:"2px 8px",fontSize:11,fontWeight:700,color:x.pct>=80?"#16a34a":x.pct>=50?"#ca8a04":"#dc2626"}}>{x.pct}%</div>
            </div>
          </div>
        ))}
      </div>

      {/* Historical trend */}
      {trend.length>0&&(
        <div style={{background:"#fff",border:"1.5px solid #f3c0c8",borderRadius:10,overflow:"hidden",marginBottom:20}}>
          <div style={{background:"#fff0f3",padding:"10px 16px",fontFamily:"'Bebas Neue',sans-serif",fontSize:14,color:"#e11d48",letterSpacing:"0.06em"}}>UNITS OUT — LAST {trend.length} DAYS</div>
          <div style={{padding:"14px 16px"}}>
            {trend.map((d,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}>
                <div style={{fontSize:10,color:"#9c6b75",width:80,flexShrink:0}}>{d.label}</div>
                <div style={{flex:1,background:"#fdf2f4",borderRadius:4,height:18,position:"relative"}}>
                  <div style={{position:"absolute",left:0,top:0,height:"100%",background:"#e11d48",borderRadius:4,width:Math.min(100,d.yard>0?Math.round((d.out/(d.out+d.yard))*100):0)+"%",opacity:0.7}}/>
                </div>
                <div style={{fontSize:11,fontWeight:600,color:"#e11d48",width:30,textAlign:"right",flexShrink:0}}>{d.out}</div>
              </div>
            ))}
            <div style={{fontSize:9,color:"#9c6b75",marginTop:8}}>Bar = % of fleet on reso · Number = units out</div>
          </div>
        </div>
      )}

      {/* PM pressure on yard */}
      {pmOnYard.length>0&&(
        <div style={{background:"#fff7ed",border:"1.5px solid #fed7aa",borderRadius:10,padding:"12px 16px",marginBottom:16}}>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:13,color:"#c2410c",marginBottom:8,letterSpacing:"0.06em"}}>🔧 PM UNITS CURRENTLY ON YARD ({pmOnYard.length})</div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {pmOnYard.map(r=>(
              <div key={r.id} style={{background:"#fff",border:"1px solid #fed7aa",borderRadius:7,padding:"7px 12px"}}>
                <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:16,color:"#c2410c"}}>#{r.unit}</div>
                <div style={{fontSize:9,color:"#92400e"}}>{urgencyColor(r.daysLeft).label} · {r.daysLeft}d</div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ── APP ───────────────────────────────────────────────────────────────────
export default function App() {
  // ── LOCALSTORAGE PERSISTENCE ──────────────────────────────────────────
  function loadLS(key, fallback){
    try{ const v=localStorage.getItem(key); return v?JSON.parse(v):fallback; }
    catch(e){ return fallback; }
  }
  function saveLS(key, val){
    try{ localStorage.setItem(key, JSON.stringify(val)); }catch(e){}
  }

  const [S, setS_]       = useState(()=>{ const saved=loadLS("branchops-state",null); return saved?{...BLANK(),...saved}:BLANK(); });
  const [history, setHistory] = useState(()=>loadLS("branchops-history",[]));
  const [dayNum, setDayNum]   = useState(()=>loadLS("branchops-daynum",1));
  const [dayLabel, setDayLabel] = useState(()=>loadLS("branchops-daylabel",todayStr()));
  const [tab, setTab]  = useState("dash");
  const [modal, setModal] = useState(null); // { type, tt, card }
  const [form, setForm]   = useState({});
  const [search, setSearch] = useState("");
  const [histOpen, setHistOpen] = useState(false);
  const [historyViewDay, setHistoryViewDay] = useState(null); // snapshot being viewed
  const [goModal, setGoModal] = useState(null); // { card, tt }
  const [goForm, setGoForm]   = useState({ customer:"", returnDate:"" });
  const [removeQ, setRemoveQ] = useState("");
  const [taskInput, setTaskInput] = useState("");
  const [notification, setNotification] = useState("");

  // Auto-save to localStorage whenever state changes
  useEffect(()=>{ saveLS("branchops-state", S); }, [S]);
  useEffect(()=>{ saveLS("branchops-history", history); }, [history]);
  useEffect(()=>{ saveLS("branchops-daynum", dayNum); }, [dayNum]);
  useEffect(()=>{ saveLS("branchops-daylabel", dayLabel); }, [dayLabel]);

  const setS = fn => setS_(s => fn(s));

  function notify(msg){ setNotification(msg); setTimeout(()=>setNotification(""),2500); }

  // ── DAY MANAGEMENT ──────────────────────────────────────────────────────
  function newDay(){
    // Save current operational day to history (keyed by day number, never overwrites)
    const snap = JSON.parse(JSON.stringify(S));
    const currentLabel = dayLabel;
    const currentNum = dayNum;
    setHistory(h => [...h, { dayNum: currentNum, label: currentLabel, snap }]);

    // Next operational day
    const nextNum = currentNum + 1;
    const nextLabel = todayStr(); // use real date for the label
    setDayNum(nextNum);
    setDayLabel(nextLabel);

    // Build new day state
    const ns = BLANK();

    // Yard: keep units physically present (not wentOut), reset daily flags
    TRUCK_TYPES.forEach(tt=>{
      ns.yard[tt] = (S.yard[tt]||[])
        .filter(c => !c.wentOut)
        .map(c=>({...c, goingOut:false, wentOut:false}));
    });

    // Reso: all cards carry forward until checked in
    TRUCK_TYPES.forEach(tt=>{
      ns.reso[tt] = (S.reso[tt]||[]).map(c=>({...c, checkInPending:false}));
    });

    // PM board + hikes carry forward
    ns.pm    = (S.pm||[]).map(c=>({...c}));
    ns.hikes = (S.hikes||[]).map(c=>({...c}));
    ns.pmScheduled = (S.pmScheduled||[]).map(c=>({...c}));
    // PM checklist rows: only carry pending + scheduled — done rows stay in history only
    ns.pmRows = (S.pmRows||[]).filter(r=>r.status!=="done").map(r=>({...r,_prev:null}));

    // Auto return reminders
    TRUCK_TYPES.forEach(tt=>{
      (ns.reso[tt]||[]).forEach(card=>{
        const d = daysUntil(card.returnDate);
        if(d===1) ns.tasks.push({ id:uid(), done:false, type:"return", unit:card.unit, tt,
          text:`Remind customer to drop off unit ${card.unit} — due TOMORROW` });
        if(d===0) ns.tasks.push({ id:uid(), done:false, type:"return", unit:card.unit, tt,
          text:`Unit ${card.unit} is due back TODAY — confirm drop-off` });
        if(d<0) ns.tasks.push({ id:uid(), done:false, type:"overdue", unit:card.unit, tt,
          text:`⚠️ Unit ${card.unit} is ${Math.abs(d)} day(s) OVERDUE — follow up with customer` });
      });
    });

    // Auto PM tasks from scheduled PM rows
    (S.pmScheduled||[]).forEach(pmRow=>{
      if(!pmRow.scheduledDate) return;
      const d = daysUntil(pmRow.scheduledDate);
      if(d===1){
        ns.tasks.push({ id:uid(), done:false, type:"pm", unit:pmRow.unit,
          text:`Remind customer to drop off unit ${pmRow.unit} for PM — scheduled TOMORROW` });
        if(pmRow.swapRequired){
          if(pmRow.swapUnit){
            ns.tasks.push({ id:uid(), done:false, type:"pm-swap", unit:pmRow.unit,
              text:"Make sure swap unit "+pmRow.swapUnit+" is available for unit "+pmRow.unit+"s PM tomorrow"+(pmRow.pmType?" ("+pmRow.pmType+" type)":"") });
          } else {
            ns.tasks.push({ id:uid(), done:false, type:"pm-swap", unit:pmRow.unit,
              text:`🔄 Unit ${pmRow.unit} PM is tomorrow — make sure a${pmRow.pmType?" "+pmRow.pmType:""} swap unit is available` });
          }
        }
      }
      if(d===0){
        ns.tasks.push({ id:uid(), done:false, type:"pm", unit:pmRow.unit,
          text:`Unit ${pmRow.unit} PM is scheduled TODAY — confirm drop-off` });
        if(pmRow.swapRequired&&pmRow.swapUnit)
          ns.tasks.push({ id:uid(), done:false, type:"pm-swap", unit:pmRow.unit,
            text:`✅ Is swap unit ${pmRow.swapUnit} here for unit ${pmRow.unit}?` });
      }
    });

    setS_(ns);
    setTab("dash");
    notify(`Day ${nextNum} started ✓`);
  }

  // View a past day's snapshot (read-only peek — does not replace live state)
  // History is just for reference, we never go "back"
  function viewDay(entry){ setHistOpen(false); setTab("dash"); /* future: show snapshot modal */ }

  // ── YARD ────────────────────────────────────────────────────────────────
  function saveYard(){
    if(!form.unit?.trim()) return;
    const tt=modal.tt;
    const hikeId = uid();
    const isHikeIn  = !!form.hikeIn  && !modal.card;
    const isHikeOut = !!form.hikeOut && !modal.card;
    const card={
      id:form.id||uid(), unit:form.unit.trim(), line:form.line||"RL",
      isPuro:!!form.isPuro, note:form.note||"", shopDate:form.shopDate||"",
      goingOut:!!form.goingOut, wentOut:!!form.wentOut,
      awaitingArrival: isHikeIn,
      hikeId: isHikeIn ? hikeId : undefined,
    };
    setS(s=>{
      let ns;
      if(isHikeOut){
        // Hike out: don't add to yard, add to hikes outbound + sent
        const hikeCard={id:hikeId,unit:card.unit,tt,dir:"out",location:"",arrival:"",placed:false,ready:false,pmDue:false,note:form.note||""};
        const sentExists=s.sent.find(c=>c.unit===card.unit);
        ns={...s,
          hikes:[...s.hikes,hikeCard],
          sent:sentExists?s.sent:[...s.sent,{id:uid(),unit:card.unit,tt,location:"",note:"Hike out"}],
        };
      } else {
        // Normal add (or hike in — card goes to yard as awaiting arrival)
        const arr=modal.card?s.yard[tt].map(c=>c.id===card.id?card:c):[...s.yard[tt],card];
        ns={...s,yard:{...s.yard,[tt]:arr}};
        // Hike in: also add to hikes inbound
        if(isHikeIn && !ns.hikes.find(h=>h.unit===card.unit&&h.dir==="in")){
          const hikeCard={id:hikeId,unit:card.unit,tt,dir:"in",location:"",arrival:"",placed:false,ready:false,pmDue:false,note:form.note||""};
          ns={...ns,hikes:[...ns.hikes,hikeCard]};
        }
      }
      // Quick action side effects (skip for hike out since unit isn't on yard)
      if(!isHikeOut){
        if(form.addPM && !ns.pm.find(p=>p.unit===card.unit))
          ns={...ns,pm:[...ns.pm,{id:uid(),unit:card.unit,tt,pmDate:"",note:""}]};
        // Auto-add to PM checklist if PM Due toggle was on
        if(form.addPM && !ns.pmRows.find(r=>r.unit===card.unit))
          ns={...ns,pmRows:[...ns.pmRows,{id:uid(),unit:card.unit,tt,pmType:"",customer:"",nextPM:"",daysLeft:0,defeDays:"",comment:"",location:"Belfield",status:"pending",scheduledDate:"",swapRequired:false,swapUnit:"",notes:"From yard — PM due",locationNotified:false,_prev:null}],pmInitialized:true};
        if(form.addTomorrow && !(ns.tomorrow[tt]||[]).find(c=>c.unit===card.unit))
          ns={...ns,tomorrow:{...ns.tomorrow,[tt]:[...(ns.tomorrow[tt]||[]),{id:uid(),unit:card.unit,note:"From yard",hold:true}]}};
        if(form.addCheckin){
          if(!ns.tasks.find(t=>t.unit===card.unit&&t.type==="checkin"))
            ns={...ns,tasks:[...ns.tasks,{id:uid(),done:false,type:"checkin",unit:card.unit,text:`Check in unit ${card.unit} (${tt})`}]};
          if(!ns.checkins.find(c=>c.unit===card.unit))
            ns={...ns,checkins:[...ns.checkins,{id:uid(),unit:card.unit,tt,hikedFrom:"",note:""}]};
        }
      }
      return ns;
    });
    closeModal();
    notify(isHikeOut?`Unit ${form.unit.trim()} hiked out -> Hikes ↑ ✓`:isHikeIn?`Unit ${form.unit.trim()} added as Awaiting Arrival -> Hikes ↓ ✓`:"Unit saved ✓");
  }

  function markGoingOut(tt,card){
    setS(s=>({...s,yard:{...s.yard,[tt]:s.yard[tt].map(c=>c.id===card.id?{...c,goingOut:!c.goingOut,wentOut:c.goingOut?false:c.wentOut}:c)}}));
  }

  function openWentOut(tt,card){
    setGoModal({card,tt});
    setGoForm({customer:card.customer||"",returnDate:twoWeeks()});
  }

  function confirmWentOut(){
    const {card,tt}=goModal;
    const {customer,returnDate}=goForm;
    if(!returnDate) return;
    setS(s=>({
      ...s,
      yard:{...s.yard,[tt]:s.yard[tt].filter(c=>c.id!==card.id)},
      reso:{...s.reso,[tt]:[...(s.reso[tt]||[]),{id:uid(),unit:card.unit,returnDate,customer,note:"Going out today",tt}]},
    }));
    setGoModal(null);
    notify(`Unit ${card.unit} moved to Short Term Reso ✓`);
  }

  function quickPM(tt,card){
    setS(s=>{
      // Add to yard pm list
      const ns = s.pm.find(p=>p.unit===card.unit)?s:{...s,pm:[...s.pm,{id:uid(),unit:card.unit,tt,pmDate:"",note:""}]};
      // Also add to PM checklist rows if not already there
      const pmRowExists=(ns.pmRows||[]).find(r=>r.unit===card.unit);
      if(!pmRowExists){
        const newRow={id:uid(),unit:card.unit,pmType:"",customer:"",nextPM:"",daysLeft:0,defeDays:"",comment:"",status:"pending",scheduledDate:"",swapRequired:false,swapUnit:"",notes:"From yard",_prev:null};
        return {...ns,pmRows:[...(ns.pmRows||[]),newRow],pmInitialized:true};
      }
      return {...ns,pmInitialized:true};
    });
    notify(`Unit ${card.unit} added to PM schedule ✓`);
  }
  function quickTomorrow(tt,card){
    setS(s=>(s.tomorrow[tt]||[]).find(c=>c.unit===card.unit)?s:{...s,tomorrow:{...s.tomorrow,[tt]:[...(s.tomorrow[tt]||[]),{id:uid(),unit:card.unit,note:"From yard",hold:true}]}});
    notify(`Unit ${card.unit} added to Tomorrow ✓`);
  }

  // Quick hike out — opens a destination modal before acting
  const [hikeOutModal, setHikeOutModal] = useState(null); // { card, tt }
  const [hikeOutDest, setHikeOutDest]   = useState("");

  function quickHikeOut(tt, card){
    setHikeOutModal({card, tt});
    setHikeOutDest("");
  }

  function confirmHikeOut(){
    const {card, tt} = hikeOutModal;
    const location = hikeOutDest.trim();
    setS(s=>{
      const hikeExists = s.hikes.find(h=>h.unit===card.unit&&h.dir==="out");
      if(hikeExists) return s;
      const hikeCard = {id:uid(),unit:card.unit,tt,dir:"out",location,arrival:"",placed:false,ready:false,pmDue:false,note:"Hike out from yard"};
      const sentExists = s.sent.find(c=>c.unit===card.unit);
      const newSent = sentExists ? s.sent : [...s.sent,{id:uid(),unit:card.unit,tt,location,note:"Hike out"}];
      const newYard = {...s.yard,[tt]:s.yard[tt].filter(c=>c.id!==card.id)};
      return {...s, hikes:[...s.hikes,hikeCard], sent:newSent, yard:newYard};
    });
    notify(`Unit ${card.unit} hiked out to ${location||"unknown"} ✓`);
    setHikeOutModal(null);
    setHikeOutDest("");
  }

  // Quick hike in — marks unit as awaiting arrival, adds to hikes section (inbound)
  const [hikeInModal, setHikeInModal]   = useState(null); // { card, tt }
  const [hikeInFrom,  setHikeInFrom]    = useState("");

  function quickHikeIn(tt, card){
    setHikeInModal({card, tt});
    setHikeInFrom("");
  }

  function confirmHikeIn(){
    const {card, tt} = hikeInModal;
    const location = hikeInFrom.trim();
    setS(s=>{
      const hikeExists = s.hikes.find(h=>h.unit===card.unit&&h.dir==="in");
      if(hikeExists) return s;
      const hikeCard = {id:uid(),unit:card.unit,tt,dir:"in",location,arrival:"",placed:false,ready:false,pmDue:false,note:"Hike in to yard"};
      const newYard = {...s.yard,[tt]:s.yard[tt].map(c=>c.id===card.id?{...c,awaitingArrival:true,hikeId:hikeCard.id,note:`Hiked from ${location||"other location"}`}:c)};
      return {...s, hikes:[...s.hikes,hikeCard], yard:newYard};
    });
    notify(`Unit ${card.unit} awaiting arrival from ${location||"other location"} ✓`);
    setHikeInModal(null);
    setHikeInFrom("");
  }

  // ── RESO ────────────────────────────────────────────────────────────────
  function saveReso(){
    if(!form.unit?.trim()) return;
    const tt=modal.tt;
    const card={id:form.id||uid(),unit:form.unit.trim(),returnDate:form.returnDate||"",customer:form.customer||"",note:form.note||"",tt};
    setS(s=>{ const arr=modal.card?s.reso[tt].map(c=>c.id===card.id?card:c):[...s.reso[tt],card]; return {...s,reso:{...s.reso,[tt]:arr}}; });
    closeModal();
  }

  function checkInFromReso(tt,card){
    setS(s=>{
      const taskExists=s.tasks.find(t=>t.unit===card.unit&&t.type==="checkin"&&!t.done);
      if(taskExists) return s;
      const taskId=uid();
      const newTask={
        id:taskId, done:false, type:"checkin", unit:card.unit,
        text:`Check in unit ${card.unit} (${tt}) — returning from reso`,
        resoCardId:card.id, resoTT:tt, customer:card.customer||"",
      };
      // Mark reso card pending
      const newReso={...s.reso,[tt]:s.reso[tt].map(c=>c.id===card.id?{...c,checkInPending:true}:c)};
      // Also add to checkins panel so it shows as "Awaiting Check In"
      const ciExists=s.checkins.find(c=>c.unit===card.unit);
      const newCI=ciExists?s.checkins:[...s.checkins,{id:uid(),unit:card.unit,tt,hikedFrom:"",note:"Awaiting check in from reso",awaitingCheckin:true}];
      return {...s, tasks:[...s.tasks,newTask], reso:newReso, checkins:newCI};
    });
    notify(`Unit ${card.unit} added to tasks & check-in list — tick off when unit arrives ✓`);
  }

  // ── TOMORROW ────────────────────────────────────────────────────────────
  function saveTomorrow(){
    if(!form.unit?.trim()) return;
    const tt=modal.tt;
    const card={id:form.id||uid(),unit:form.unit.trim(),note:form.note||"",hold:!!form.hold};
    setS(s=>{ const arr=modal.card?s.tomorrow[tt].map(c=>c.id===card.id?card:c):[...s.tomorrow[tt],card]; return {...s,tomorrow:{...s.tomorrow,[tt]:arr}}; });
    closeModal();
  }
  function toggleHold(tt,id){ setS(s=>({...s,tomorrow:{...s.tomorrow,[tt]:s.tomorrow[tt].map(c=>c.id===id?{...c,hold:!c.hold}:c)}})); }

  // ── PM ──────────────────────────────────────────────────────────────────
  function savePM(){
    if(!form.unit?.trim()) return;
    const card={id:form.id||uid(),unit:form.unit.trim(),tt:form.tt||"",pmDate:form.pmDate||"",note:form.note||""};
    setS(s=>({...s,pm:modal.card?s.pm.map(c=>c.id===card.id?card:c):[...s.pm,card]}));
    closeModal();
  }

  // ── HIKES ───────────────────────────────────────────────────────────────
  function saveHike(){
    if(!form.unit?.trim()) return;
    const card={id:form.id||uid(),unit:form.unit.trim(),tt:form.tt||"",dir:form.dir||"in",location:form.location||"",arrival:form.arrival||"",placed:!!form.placed,ready:!!form.ready,pmDue:!!form.pmDue,note:form.note||"",awaitingArrival:form.dir==="in"};
    setS(s=>{
      const newHikes=modal.card?s.hikes.map(c=>c.id===card.id?card:c):[...s.hikes,card];
      let ns={...s,hikes:newHikes};
      // PM side effect
      if(card.pmDue&&!ns.pm.find(p=>p.unit===card.unit))
        ns={...ns,pm:[...ns.pm,{id:uid(),unit:card.unit,tt:card.tt,pmDate:"",note:"Via hike"}]};
      // Outbound hike: also add to Sent panel so it's tracked there
      if(card.dir==="out"&&!ns.sent.find(c=>c.unit===card.unit))
        ns={...ns,sent:[...ns.sent,{id:uid(),unit:card.unit,tt:card.tt,location:card.location,note:`Hike out · arrival ${card.arrival||"TBD"}`}]};
      // Inbound hike: add to yard as "Awaiting Arrival" so it shows on the board
      if(card.dir==="in"&&!modal.card){
        const ttKey=card.tt||TRUCK_TYPES[0];
        if(!(ns.yard[ttKey]||[]).find(c=>c.unit===card.unit))
          ns={...ns,yard:{...ns.yard,[ttKey]:[...(ns.yard[ttKey]||[]),{id:uid(),unit:card.unit,line:"RL",isPuro:false,note:"Awaiting arrival",shopDate:"",goingOut:false,wentOut:false,awaitingArrival:true,hikeId:card.id}]}};
      }
      return ns;
    });
    closeModal();
    notify(card.dir==="out"?`Outbound hike for ${card.unit} placed — added to Sent ✓`:`Inbound hike for ${card.unit} — added to yard as Awaiting Arrival ✓`);
  }
  function toggleHikeField(id,f){ setS(s=>({...s,hikes:s.hikes.map(h=>h.id===id?{...h,[f]:!h[f]}:h)})); }

  // Confirm inbound hike arrived — removes awaiting flag, becomes normal yard unit
  function confirmHikeArrival(tt, card){
    setS(s=>{
      const yard = {...s.yard,[tt]:s.yard[tt].map(c=>c.id===card.id?{...c,awaitingArrival:false,note:"",hikeId:undefined}:c)};
      const hikes = s.hikes.map(h=>h.id===card.hikeId?{...h,placed:true,ready:true}:h);
      // Auto-add arrival task if not already there
      const taskExists = s.tasks.find(t=>t.unit===card.unit&&t.type==="hike-arrive");
      const tasks = taskExists ? s.tasks : [...s.tasks,{id:uid(),done:false,type:"hike-arrive",unit:card.unit,tt,text:"Unit "+card.unit+" arrived from hike — inspect and assign line"}];
      return {...s, yard, hikes, tasks};
    });
    notify("Unit "+card.unit+" arrived ✓ — task created");
  }

  // ── SENT / CHECKINS ─────────────────────────────────────────────────────
  function saveSent(){
    if(!form.unit?.trim()) return;
    const card={id:form.id||uid(),unit:form.unit.trim(),tt:form.tt||"",location:form.location||"",note:form.note||""};
    setS(s=>({...s,sent:modal.card?s.sent.map(c=>c.id===card.id?card:c):[...s.sent,card]}));
    closeModal();
  }
  function saveCheckin(){
    if(!form.unit?.trim()) return;
    const card={id:form.id||uid(),unit:form.unit.trim(),tt:form.tt||"",hikedFrom:form.customer||"",note:form.note||""};
    setS(s=>{
      const newCI = modal.card ? s.checkins.map(c=>c.id===card.id?card:c) : [...s.checkins,card];
      // If new check-in (hike placed), add unit to yard as awaiting arrival
      if(!modal.card && card.tt){
        const yardExists = (s.yard[card.tt]||[]).find(c=>c.unit===card.unit);
        if(!yardExists){
          const newYardCard = {id:uid(),unit:card.unit,line:"RL",isPuro:false,note:`Hiked from ${card.hikedFrom||"other location"}`,shopDate:"",goingOut:false,wentOut:false,awaitingArrival:true};
          return {...s, checkins:newCI, yard:{...s.yard,[card.tt]:[...(s.yard[card.tt]||[]),newYardCard]}};
        }
      }
      return {...s, checkins:newCI};
    });
    closeModal();
    notify(`Unit ${form.unit.trim()} added to yard as Awaiting Arrival ✓`);
  }

  // ── TASKS ───────────────────────────────────────────────────────────────
  function toggleTask(id){
    setS(s=>{
      const task = s.tasks.find(t=>t.id===id);
      if(!task) return s;
      const nowDone = !task.done;
      // Base: toggle the task
      let ns={...s, tasks:s.tasks.map(t=>t.id===id?{...t,done:nowDone}:t)};
      // If completing a reso check-in task -> move unit to yard + remove from reso
      if(nowDone && task.type==="checkin" && task.resoCardId && task.resoTT){
        const tt=task.resoTT;
        // Add to yard as WL (just returned from rental)
        const yardExists=(ns.yard[tt]||[]).find(c=>c.unit===task.unit);
        if(!yardExists){
          ns={...ns,yard:{...ns.yard,[tt]:[...(ns.yard[tt]||[]),
            {id:uid(),unit:task.unit,line:"WL",isPuro:false,note:"Returned from reso",shopDate:"",goingOut:false,wentOut:false}
          ]}};
        }
        // Remove from reso
        ns={...ns,reso:{...ns.reso,[tt]:ns.reso[tt].filter(c=>c.id!==task.resoCardId)}};
        // Remove from checkins panel (it's now on yard)
        ns={...ns,checkins:ns.checkins.filter(c=>c.unit!==task.unit)};
      }
      // If un-completing a reso check-in task -> restore reso card pending state
      if(!nowDone && task.type==="checkin" && task.resoCardId && task.resoTT){
        const tt=task.resoTT;
        // Remove from yard if it was added
        ns={...ns,yard:{...ns.yard,[tt]:ns.yard[tt].filter(c=>c.unit!==task.unit||c.note!=="Returned from reso")}};
        // Restore reso card (it may have been removed — can't restore if gone, but mark un-pending if still there)
        ns={...ns,reso:{...ns.reso,[tt]:ns.reso[tt].map(c=>c.id===task.resoCardId?{...c,checkInPending:false}:c)}};
      }
      return ns;
    });
  }
  function addTask(text){ if(!text.trim()) return; setS(s=>({...s,tasks:[...s.tasks,{id:uid(),done:false,type:"general",unit:"",text:text.trim()}]})); }
  function delTask(id){ setS(s=>({...s,tasks:s.tasks.filter(t=>t.id!==id)})); }

  // ── REMOVE UNIT EVERYWHERE ──────────────────────────────────────────────
  function removeUnit(u){
    if(!u.trim()) return;
    setS(s=>{
      const y={},r={},t={};
      TRUCK_TYPES.forEach(tt=>{ y[tt]=(s.yard[tt]||[]).filter(c=>c.unit!==u); r[tt]=(s.reso[tt]||[]).filter(c=>c.unit!==u); t[tt]=(s.tomorrow[tt]||[]).filter(c=>c.unit!==u); });
      return {...s,yard:y,reso:r,tomorrow:t,pm:s.pm.filter(c=>c.unit!==u),tasks:s.tasks.filter(c=>c.unit!==u),hikes:s.hikes.filter(c=>c.unit!==u),sent:s.sent.filter(c=>c.unit!==u),checkins:s.checkins.filter(c=>c.unit!==u)};
    });
    notify(`Unit ${u} removed from all sections ✓`);
    setRemoveQ("");
  }

  // ── MODAL HELPERS ────────────────────────────────────────────────────────
  function openModal(type,tt=null,card=null){
    setModal({type,tt,card});
    if(card && type==="yard"){
      // Check both the yard PM list AND the imported PM checklist (pmRows)
    const hasPM=S.pm.find(p=>p.unit===card.unit) || (S.pmRows||[]).find(p=>p.unit===card.unit&&p.status!=="done");
      const hasTom=Object.values(S.tomorrow).flat().find(c=>c.unit===card.unit);
      const hasCI=S.checkins.find(c=>c.unit===card.unit);
      setForm({...card,addPM:!!hasPM,addTomorrow:!!hasTom,addCheckin:!!hasCI,goingOut:!!card.goingOut});
    } else {
      setForm(card?{...card}:{unit:"",line:"RL",isPuro:false,note:"",shopDate:"",returnDate:"",customer:"",pmDate:"",dir:"in",location:"",arrival:"",placed:false,ready:false,pmDue:false,hold:false,addPM:false,addTomorrow:false,addCheckin:false,tt:tt||""});
    }
  }
  function closeModal(){ setModal(null); setForm({}); }
  const sf = k => e => setForm(f=>({...f,[k]:e.target.type==="checkbox"?e.target.checked:e.target.value}));

  // ── STATS ────────────────────────────────────────────────────────────────
  const totalYard  = TRUCK_TYPES.reduce((a,t)=>a+(S.yard[t]||[]).length,0);
  const totalReso  = TRUCK_TYPES.reduce((a,t)=>a+(S.reso[t]||[]).length,0);
  const totalTom   = TRUCK_TYPES.reduce((a,t)=>a+(S.tomorrow[t]||[]).length,0);
  const avail      = TRUCK_TYPES.reduce((a,t)=>a+(S.yard[t]||[]).filter(c=>["RL","WL","SRL"].includes(c.line)&&!c.isPuro&&!c.goingOut).length,0);
  const goingOut   = TRUCK_TYPES.reduce((a,t)=>a+(S.yard[t]||[]).filter(c=>c.goingOut).length,0);
  const tasksDone  = S.tasks.filter(t=>t.done).length;
  // Cross-module: PM units currently on yard
  const allYardFlat = TRUCK_TYPES.flatMap(t=>S.yard[t]||[]);
  const pmDueOnYard = (S.pmRows||[]).filter(r=>r.status!=="done"&&allYardFlat.some(u=>String(u.unit).trim()===String(r.unit).trim())).length;
  // Cross-module: active PM count for header badge
  const activePMs = (S.pmRows||[]).filter(r=>r.status!=="done").length;
  // Units on yard that need to be held for PM swaps
  const holdForSwapUnits = new Set(
    (S.pmRows||[]).filter(r=>r.status!=="done"&&r.swapUnit&&r.swapUnit.trim())
      .map(r=>String(r.swapUnit).trim())
  );
  const returnAlerts = [];
  TRUCK_TYPES.forEach(tt=>{ (S.reso[tt]||[]).forEach(c=>{ const d=daysUntil(c.returnDate); if(d===0||d===-1||d<0) returnAlerts.push({...c,tt,days:d}); else if(d===1) returnAlerts.push({...c,tt,days:d}); }); });

  // ── SEARCH ───────────────────────────────────────────────────────────────
  const searchResults = !search.trim() ? null : (() => {
    const q=search.trim().toLowerCase(), res=[];
    TRUCK_TYPES.forEach(tt=>{
      (S.yard[tt]||[]).forEach(c=>{ if(c.unit.toLowerCase().includes(q)) res.push({where:"Yard",tt,unit:c.unit,detail:c.line}); });
      (S.reso[tt]||[]).forEach(c=>{ if(c.unit.toLowerCase().includes(q)) res.push({where:"Reso",tt,unit:c.unit,detail:c.returnDate?`Back ${fmtDate(c.returnDate)}`:""}); });
      (S.tomorrow[tt]||[]).forEach(c=>{ if(c.unit.toLowerCase().includes(q)) res.push({where:"Tomorrow",tt,unit:c.unit,detail:c.hold?"HOLD":""}); });
    });
    S.pm.forEach(c=>{ if(c.unit.toLowerCase().includes(q)) res.push({where:"PM",tt:c.tt,unit:c.unit,detail:c.pmDate?fmtDate(c.pmDate):""}); });
    S.hikes.forEach(c=>{ if(c.unit.toLowerCase().includes(q)) res.push({where:`Hike ${c.dir==="in"?"↓":"↑"}`,tt:c.tt,unit:c.unit,detail:c.location}); });
    S.sent.forEach(c=>{ if(c.unit.toLowerCase().includes(q)) res.push({where:"Sent",tt:c.tt,unit:c.unit,detail:c.location}); });
    return res;
  })();

  const TABS = [["dash","📋 Dashboard"],["pm","🔧 PM"],["ground","🔄 Ground Units"],["puro","🟣 Purolator Fleet"],["hikes","✈️ Hikes"],["other","📤 Sent & CI"],["tasks","✅ Tasks"],["contacts","📞 Contacts"],["stats","📊 Utilization"]];

  // ── YARD CARD (reused in both dashboard and yard tab) ────────────────────
  const YardCard = ({card,tt}) => {
    const ls=card.isPuro?LINE.PUR:(LINE[card.line]||LINE.RL);
    // Check both the yard PM list AND the imported PM checklist (pmRows)
    const hasPM=S.pm.find(p=>p.unit===card.unit) || (S.pmRows||[]).find(p=>p.unit===card.unit&&p.status!=="done");
    const hasTom=Object.values(S.tomorrow||{}).flat().find(c=>c.unit===card.unit);
    const hasCheckinPending=S.tasks.find(t=>t.unit===card.unit&&t.type==="checkin"&&!t.done);

    // Awaiting arrival (inbound hike) — special state
    if(card.awaitingArrival){
      return (
        <div className="ucard" style={{background:"#f0fff4",border:"2px dashed #16a34a",color:"#4ade80",position:"relative"}}>
          <div className="unum">{card.unit}</div>
          <div className="usub" style={{color:"#4ade80",opacity:0.7}}>✈️ Awaiting arrival</div>
          {card.note&&<div className="usub">{card.note}</div>}
          <div className="qa-row" onClick={e=>e.stopPropagation()}>
            <button style={{background:"#16a34a",border:"none",borderRadius:4,color:"#fdf2f4",fontSize:9,fontWeight:700,padding:"3px 8px",cursor:"pointer",fontFamily:"inherit"}}
              onClick={()=>confirmHikeArrival(tt,card)}>✅ Arrived</button>
          </div>
          <button className="xcbtn" onClick={e=>{e.stopPropagation();setS(s=>({...s,yard:{...s.yard,[tt]:s.yard[tt].filter(c=>c.id!==card.id)}}));}}>✕</button>
        </div>
      );
    }

    const tomEntry = hasTom; // the tomorrow entry for this unit
    return (
      <div className={`ucard ${card.goingOut?"ucard-go":""}`} style={{background:ls.bg,color:ls.text,outline:hasTom&&!card.goingOut?"2px solid #f59e0b":undefined,outlineOffset:hasTom&&!card.goingOut?"1px":undefined}} onClick={()=>openModal("yard",tt,card)}>
        <div className={`unum ${hasPM?"pm-b":""}`}>{card.unit}</div>
        <div className="usub">{card.isPuro?"PURO":card.line}{card.note?" · "+card.note:""}</div>
        {card.shopDate&&<div className="usub">Out: {fmtDate(card.shopDate)}</div>}
        {hasPM&&<div className="usub">🔧 PM sch.</div>}
        {/* Tomorrow strip — shown on card body */}
        {hasTom&&!card.goingOut&&(
          <div style={{marginTop:4,background:"#fff7ed",border:"1px solid #f59e0b",borderRadius:3,padding:"2px 5px",fontSize:8,color:"#fcd34d",fontWeight:700}}>
            📅 NEEDED TOMORROW{tomEntry?.hold?" · 🔴 HOLD":""}
          </div>
        )}
        {card.goingOut&&(
          <div className="go-strip" style={{background:card.wentOut?"#14532d":undefined,borderColor:card.wentOut?"#16a34a":undefined}}>
            {card.wentOut?"✅ WENT OUT":"🚀 GOING OUT"}{card.returnDate?" · back "+fmtDate(card.returnDate):""}
          </div>
        )}
        {holdForSwapUnits.has(String(card.unit))&&(
          <div style={{marginTop:3,background:"#eff6ff",border:"1px solid #93c5fd",borderRadius:3,padding:"2px 5px",fontSize:7,color:"#1d4ed8",fontWeight:700,letterSpacing:"0.04em"}}>🔒 HOLD — PM SWAP</div>
        )}
        {hasCheckinPending&&(
          <div style={{marginTop:4,background:"#dcfce7",border:"1px solid #4ade80",borderRadius:3,padding:"2px 5px",fontSize:8,color:"#4ade80",fontWeight:700}}>
            ✅ AWAITING CHECK IN
          </div>
        )}
        <div className="qa-row" onClick={e=>e.stopPropagation()}>
          <button className={`qa-go ${card.goingOut?"on":""}`} onClick={()=>markGoingOut(tt,card)}>{card.goingOut?"✓ Out":"🚀 Out"}</button>
          {card.goingOut&&!card.wentOut&&<button className="qa-btn" style={{background:"#eff6ff",color:"#93c5fd"}} onClick={()=>openWentOut(tt,card)}>📋 Went Out</button>}
          {card.wentOut&&<span className="qa-badge green">✓ In Reso</span>}
          {/* Hike Out — removes from yard, adds to hikes outbound */}
          {!card.goingOut&&!card.wentOut&&!card.awaitingArrival&&(
            <button className="qa-btn" style={{background:"#4c1d95",color:"#c4b5fd"}} onClick={()=>quickHikeOut(tt,card)}>↑ Hike Out</button>
          )}
          {/* Hike In — marks awaiting arrival, adds to hikes inbound */}
          {!card.awaitingArrival&&(
            <button className="qa-btn" style={{background:"#14532d",color:"#86efac"}} onClick={()=>quickHikeIn(tt,card)}>↓ Hike In</button>
          )}
          {hasPM&&<span className="qa-badge orange">🔧 PM</span>}
        </div>
        <button className="xcbtn" onClick={e=>{e.stopPropagation();setS(s=>({...s,yard:{...s.yard,[tt]:s.yard[tt].filter(c=>c.id!==card.id)}}));}}>✕</button>
      </div>
    );
  };

  // ── RESO CARD (reused) ───────────────────────────────────────────────────
  const ResoCard = ({card,tt}) => {
    const days=daysUntil(card.returnDate);
    const overdue=days!==null&&days<0, urgent=days===0, soon=days===1;
    const cdLabel=days===null?"":overdue?`${Math.abs(days)}d overdue`:urgent?"due TODAY":soon?"due TOMORROW":`${days}d left`;
    const cdColor=overdue||urgent?"#ef4444":soon?"#f59e0b":"#475569";
    return (
      <div className={`reso-card ${urgent||overdue?"r-urgent":soon?"r-soon":""}`} style={card.checkInPending?{borderColor:"#4ade80",boxShadow:"0 0 8px #4ade8033"}:{}} onClick={()=>openModal("reso",tt,card)}>
        <div style={{fontSize:13,fontWeight:700,color:"#93c5fd"}}>{card.unit}</div>
        {card.customer&&<div style={{fontSize:9,color:"#7dd3fc",marginTop:1}}>{card.customer}</div>}
        {card.returnDate&&<div style={{fontSize:9,color:"#64748b",marginTop:2}}>Back {urgent?"TODAY":soon?"TOMORROW":`${fmtDate(card.returnDate)}`}</div>}
        {cdLabel&&<div style={{fontSize:10,fontWeight:700,color:cdColor}}>{cdLabel}</div>}
        {!card.checkInPending&&(
          <button onClick={e=>{e.stopPropagation();checkInFromReso(tt,card);}} className="ci-btn">✅ Add to Daily Tasks</button>
        )}
        {card.checkInPending&&(
          <div style={{marginTop:6,background:"#1c3a1c",border:"2px solid #4ade80",borderRadius:6,padding:"6px 8px",textAlign:"center"}}>
            <div style={{fontSize:10,fontWeight:700,color:"#4ade80"}}>⏳ AWAITING CHECK-IN</div>
            <div style={{fontSize:9,color:"#86efac",marginTop:2}}>Task added -> tick off in Daily Tasks when unit arrives</div>
          </div>
        )}
        <button className="xcbtn" onClick={e=>{e.stopPropagation();setS(s=>({...s,reso:{...s.reso,[tt]:s.reso[tt].filter(c=>c.id!==card.id)}}));}}>✕</button>
      </div>
    );
  };

  // ── BOARD GRID (reused) ──────────────────────────────────────────────────
  const BoardGrid = ({data,renderCard,addCard,style={}}) => (
    <div className="grid" style={style}>
      {TRUCK_TYPES.map(tt=>(
        <div key={tt}>
          <div className="col-hdr" title={tt}>{tt}</div>
          <div className="bcol">
            {(data[tt]||[]).map(c=>renderCard(c,tt))}
            {addCard&&<div className="add-btn" onClick={()=>addCard(tt)}>+</div>}
          </div>
        </div>
      ))}
    </div>
  );

  // ── RENDER ───────────────────────────────────────────────────────────────
  return (
    <div style={{fontFamily:"'Inter',sans-serif",minHeight:"100vh",background:"#fdf2f4",color:"#1a1a2e"}}>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Bebas+Neue&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:5px;height:5px;}
        ::-webkit-scrollbar-thumb{background:#f3c0c8;border-radius:4px;}
        ::-webkit-scrollbar-track{background:#fdf2f4;}
        body,#root{font-family:'Inter',sans-serif;}
        input,select,textarea{background:#fff;border:1.5px solid #f3c0c8;color:#1a1a2e;border-radius:8px;padding:8px 12px;font-family:inherit;font-size:13px;width:100%;outline:none;transition:border 0.15s;box-shadow:0 1px 3px rgba(225,29,72,0.06);}
        input:focus,select:focus,textarea:focus{border-color:#e11d48;box-shadow:0 0 0 3px rgba(225,29,72,0.1);}
        select option{background:#fff;color:#1a1a2e;}
        textarea{resize:vertical;min-height:56px;}
        .btn{cursor:pointer;border:none;border-radius:8px;font-family:inherit;font-size:12px;font-weight:600;padding:8px 16px;transition:all 0.15s;}
        .btn-amber{background:#f59e0b;color:#fff;box-shadow:0 2px 6px rgba(245,158,11,0.3);}.btn-amber:hover{background:#d97706;transform:translateY(-1px);}
        .btn-ghost{background:#fff;color:#6b4c52;border:1.5px solid #f3c0c8;}.btn-ghost:hover{border-color:#e11d48;color:#e11d48;}
        .btn-green{background:#16a34a;color:#fff;box-shadow:0 2px 6px rgba(22,163,74,0.3);}.btn-green:hover{background:#15803d;transform:translateY(-1px);}
        .btn-red{background:#fee2e2;color:#dc2626;border:1.5px solid #fca5a5;}.btn-red:hover{background:#fecaca;}
        .btn-sm{padding:5px 12px;font-size:11px;}
        .overlay{position:fixed;inset:0;background:rgba(100,20,40,0.45);backdrop-filter:blur(4px);z-index:100;display:flex;align-items:center;justify-content:center;padding:16px;}
        .modal{background:#fff;border:1.5px solid #f3c0c8;border-radius:16px;padding:24px;width:100%;max-width:440px;max-height:90vh;overflow-y:auto;box-shadow:0 20px 60px rgba(225,29,72,0.15);}
        .field{display:flex;flex-direction:column;gap:5px;margin-bottom:14px;}
        .field label{font-size:10px;color:#9c6b75;letter-spacing:0.1em;text-transform:uppercase;font-weight:600;}
        .row2{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
        .section-title{font-family:'Bebas Neue',sans-serif;letter-spacing:0.08em;font-size:22px;margin-bottom:4px;}
        .section-sub{font-size:11px;color:#9c6b75;margin-bottom:10px;}
        .grid{display:grid;grid-template-columns:repeat(10,minmax(110px,1fr));gap:2px;background:#f9d5dc;border:1.5px solid #f3c0c8;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(225,29,72,0.08);}
        .col-hdr{font-size:9px;color:#9c6b75;text-transform:uppercase;letter-spacing:0.07em;text-align:center;padding:6px 4px;border-bottom:1.5px solid #f3c0c8;background:#fff0f3;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:600;}
        .bcol{background:#fdf2f4;padding:7px;min-height:110px;display:flex;flex-direction:column;gap:6px;}
        .ucard{border-radius:10px;padding:8px 9px 6px;cursor:pointer;position:relative;transition:transform 0.15s,box-shadow 0.15s;box-shadow:0 2px 6px rgba(0,0,0,0.08);}
        .ucard:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,0.12);}
        .ucard-go{outline:2px solid #f97316;outline-offset:2px;box-shadow:0 0 12px rgba(249,115,22,0.35)!important;}
        .unum{font-size:13px;font-weight:700;line-height:1.2;}
        .pm-b{text-decoration:underline dotted;text-underline-offset:2px;}
        .usub{font-size:9px;opacity:0.75;margin-top:2px;line-height:1.3;}
        .go-strip{margin-top:5px;background:#fff7ed;border:1px solid #f97316;border-radius:4px;padding:2px 6px;font-size:8px;color:#c2410c;font-weight:700;}
        .qa-row{display:flex;gap:3px;margin-top:6px;flex-wrap:wrap;}
        .qa-go{border:1px solid #fed7aa;border-radius:5px;cursor:pointer;font-size:8px;padding:3px 7px;font-family:inherit;font-weight:700;background:#fff7ed;color:#ea580c;transition:all 0.1s;}
        .qa-go.on{background:#f97316;color:#fff;border-color:#f97316;}
        .qa-btn{border:1px solid #bfdbfe;border-radius:5px;cursor:pointer;font-size:8px;padding:3px 7px;font-family:inherit;font-weight:700;background:#eff6ff;color:#2563eb;}
        .qa-badge{border-radius:4px;font-size:8px;padding:2px 6px;font-weight:700;display:inline-flex;align-items:center;}
        .qa-badge.green{background:#dcfce7;color:#16a34a;border:1px solid #bbf7d0;}
        .qa-badge.amber{background:#fef9c3;color:#ca8a04;border:1px solid #fde68a;}
        .qa-badge.orange{background:#fff7ed;color:#ea580c;border:1px solid #fed7aa;}
        .add-btn{background:#fff;border:2px dashed #f3c0c8;border-radius:8px;color:#f3c0c8;font-size:20px;text-align:center;cursor:pointer;padding:7px;user-select:none;transition:all 0.15s;}
        .add-btn:hover{border-color:#e11d48;color:#e11d48;background:#fff5f7;}
        .xcbtn{position:absolute;top:4px;right:4px;background:rgba(255,255,255,0.85);border:none;border-radius:4px;cursor:pointer;font-size:9px;padding:2px 5px;color:#9c6b75;font-weight:700;line-height:1;}
        .xcbtn:hover{background:#fff;color:#e11d48;}
        .tab{cursor:pointer;padding:9px 16px;font-size:12px;font-weight:600;border:none;background:transparent;color:#9c6b75;font-family:inherit;border-bottom:2.5px solid transparent;transition:all 0.15s;white-space:nowrap;}
        .tab.on{color:#e11d48;border-bottom:2.5px solid #e11d48;}.tab:hover:not(.on){color:#6b4c52;}
        .reso-card{background:#fff;border:1.5px solid #bfdbfe;border-radius:10px;padding:10px;cursor:pointer;position:relative;transition:transform 0.15s,box-shadow 0.15s;box-shadow:0 2px 8px rgba(59,130,246,0.08);}
        .reso-card:hover{transform:translateY(-1px);box-shadow:0 4px 14px rgba(59,130,246,0.12);}
        .r-urgent{border-color:#fca5a5!important;box-shadow:0 0 10px rgba(239,68,68,0.2)!important;}
        .r-soon{border-color:#fde68a!important;}
        .ci-btn{margin-top:7px;width:100%;background:#dcfce7;border:1.5px solid #16a34a;border-radius:6px;color:#15803d;font-size:10px;font-weight:700;padding:5px;cursor:pointer;font-family:inherit;transition:all 0.15s;}
        .ci-btn:hover{background:#bbf7d0;}
        .ci-pending{margin-top:6px;background:#f0fdf4;border:1.5px solid #4ade80;border-radius:6px;color:#15803d;font-size:9px;padding:4px 7px;text-align:center;font-weight:600;}
        .tom-card{background:#fff;border:1.5px solid #fde68a;border-radius:10px;padding:8px 10px;cursor:pointer;position:relative;box-shadow:0 2px 8px rgba(234,179,8,0.1);}
        .hold-badge{background:#fee2e2;color:#dc2626;border-radius:4px;font-size:8px;padding:2px 6px;font-weight:700;display:inline-block;margin-top:3px;}
        .pm-card{background:#fff;border:1.5px solid #fed7aa;border-radius:10px;padding:10px 12px;cursor:pointer;position:relative;transition:border-color 0.15s;box-shadow:0 2px 6px rgba(249,115,22,0.08);}
        .pm-card:hover{border-color:#f97316;}
        .hike-card{border-radius:10px;padding:12px;position:relative;box-shadow:0 2px 8px rgba(0,0,0,0.05);}
        .hike-in{background:#f0fdf4;border:1.5px solid #86efac;}.hike-out{background:#fdf4ff;border:1.5px solid #d8b4fe;}
        .side-card{background:#fff;border:1.5px solid #f3c0c8;border-radius:10px;padding:10px 12px;position:relative;transition:all 0.15s;cursor:pointer;box-shadow:0 2px 6px rgba(225,29,72,0.05);}
        .side-card:hover{border-color:#e11d48;box-shadow:0 4px 14px rgba(225,29,72,0.1);}
        .chk-box{width:17px;height:17px;border-radius:4px;border:2px solid #f3c0c8;background:#fff;cursor:pointer;appearance:none;flex-shrink:0;margin-top:2px;transition:all 0.15s;}
        .chk-box:checked{background:#e11d48;border-color:#e11d48;}
        .tog{display:flex;align-items:center;gap:8px;cursor:pointer;}
        .tog input[type=checkbox]{width:15px;height:15px;cursor:pointer;accent-color:#e11d48;}
        .stat-box{text-align:center;}
        .stat-num{font-family:'Bebas Neue',sans-serif;font-size:24px;line-height:1;}
        .stat-lbl{font-size:9px;color:#9c6b75;text-transform:uppercase;letter-spacing:0.08em;margin-top:2px;font-weight:600;}
        .avail{background:#dcfce7;border:1.5px solid #16a34a;border-radius:8px;padding:5px 14px;display:inline-flex;flex-direction:column;align-items:center;box-shadow:0 2px 8px rgba(22,163,74,0.15);}
        .notif{position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#1a1a2e;border:none;border-radius:10px;padding:11px 22px;font-size:13px;color:#fff;z-index:200;pointer-events:none;animation:fadein 0.2s;white-space:nowrap;box-shadow:0 4px 20px rgba(0,0,0,0.2);}
        @keyframes fadein{from{opacity:0;transform:translateX(-50%) translateY(10px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}
        .alert-bar{background:#fff5f5;border-bottom:2px solid #fca5a5;padding:7px 18px;display:flex;gap:10px;flex-wrap:wrap;align-items:center;}
        .alert-chip{background:#fee2e2;color:#dc2626;border:1px solid #fca5a5;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:700;}
        .alert-chip.soon{background:#fef9c3;color:#ca8a04;border-color:#fde68a;}
        .search-res{display:flex;flex-wrap:wrap;gap:8px;padding:12px 18px;border-bottom:1.5px solid #f3c0c8;background:#fff;}
        .search-chip{background:#fff;border:1.5px solid #f3c0c8;border-radius:8px;padding:7px 12px;min-width:115px;box-shadow:0 1px 4px rgba(225,29,72,0.06);}
      `}</style>

      {/* ── HEADER ── */}
      <div style={{borderBottom:"1px solid #1f2937",padding:"12px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
        <div>
          <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:26,color:"#e11d48",letterSpacing:"0.1em"}}>BRANCH OPS</div>
          <div style={{fontSize:11,color:"#9c6b75",fontWeight:500}}>Day {dayNum} · {dayLabel}</div>
        </div>

        {/* Stats */}
        <div style={{display:"flex",gap:16,alignItems:"center",flexWrap:"wrap"}}>
          {[["Yard",totalYard,"#7dd3fc"],["Reso",totalReso,"#f59e0b"],["Tmrw",totalTom,"#fcd34d"],["PM",activePMs,"#fb923c"],["Hikes",S.hikes.length,"#67e8f9"]].map(([l,v,c])=>(
            <div key={l} className="stat-box"><div className="stat-num" style={{color:c}}>{v}</div><div className="stat-lbl">{l}</div></div>
          ))}
          <div className="avail">
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#4ade80"}}>{avail}</div>
            <div style={{fontSize:9,color:"#166534",textTransform:"uppercase",letterSpacing:"0.06em"}}>Available</div>
          </div>
          {goingOut>0&&<div className="stat-box"><div className="stat-num" style={{color:"#f97316"}}>{goingOut}</div><div className="stat-lbl">Going Out</div></div>}
          {S.tasks.length>0&&<div className="stat-box"><div className="stat-num" style={{color:tasksDone===S.tasks.length?"#4ade80":"#a07880"}}>{tasksDone}/{S.tasks.length}</div><div className="stat-lbl">Tasks</div></div>}
        </div>

        {/* Controls */}
        <div style={{display:"flex",gap:7,alignItems:"center",flexWrap:"wrap"}}>
          {/* Search */}
          <div style={{position:"relative"}}>
            <input style={{background:"#ffffff",border:"1px solid #f3c0c8",borderRadius:6,padding:"6px 12px",fontFamily:"inherit",fontSize:12,color:"#1a1a2e",outline:"none",width:170}} placeholder="🔍 Search unit #" value={search} onChange={e=>setSearch(e.target.value)}/>
            {search&&<button style={{position:"absolute",right:8,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"#9c6b75",cursor:"pointer"}} onClick={()=>setSearch("")}>✕</button>}
          </div>
          {/* Remove */}
          <div style={{display:"flex",gap:5}}>
            <input style={{background:"#1c0a0a",border:"1px solid #ef444466",borderRadius:6,padding:"6px 10px",fontFamily:"inherit",fontSize:12,color:"#fca5a5",outline:"none",width:130}} placeholder="Unit # remove" value={removeQ} onChange={e=>setRemoveQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&removeQ&&removeUnit(removeQ)}/>
            <button className="btn btn-red btn-sm" onClick={()=>removeQ&&removeUnit(removeQ)}>🗑</button>
          </div>
          {/* History */}
          <button className="btn btn-ghost btn-sm" onClick={()=>setHistOpen(true)}>📅 History {history.length>0?`(${history.length})`:""}</button>
          <button className="btn btn-green btn-sm" onClick={newDay}>🌅 Start Day {dayNum+1}</button>
        </div>
      </div>

      {/* Return Alerts */}
      {returnAlerts.length>0&&(
        <div className="alert-bar">
          <span style={{fontSize:10,color:"#ef4444",fontWeight:700}}>⚠️ RETURNS:</span>
          {returnAlerts.map((a,i)=>(
            <span key={i} className={`alert-chip ${a.days===1?"soon":""}`}>
              #{a.unit} — {a.days<0?`${Math.abs(a.days)}d OVERDUE`:a.days===0?"TODAY":"TOMORROW"} ({fmtDate(a.returnDate)})
            </span>
          ))}
        </div>
      )}

      {/* Search results */}
      {searchResults&&(
        <div className="search-res">
          {searchResults.length===0
            ?<div style={{fontSize:11,color:"#9c6b75"}}>No results for "{search}"</div>
            :searchResults.map((r,i)=>(
              <div key={i} className="search-chip">
                <div style={{fontSize:12,fontWeight:700,color:"#1a1a2e"}}>{r.unit}</div>
                <div style={{fontSize:9,color:"#7a5560",marginTop:1}}>{r.where}{r.tt?" · "+r.tt:""}</div>
                {r.detail&&<div style={{fontSize:9,color:"#f59e0b",marginTop:1}}>{r.detail}</div>}
              </div>
            ))
          }
        </div>
      )}

      {/* Legend */}
      <div style={{padding:"5px 18px",borderBottom:"1px solid #1f2937",display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
        {Object.entries(LINE).map(([k,v])=>(
          <div key={k} style={{display:"flex",alignItems:"center",gap:4}}>
            <div style={{width:8,height:8,borderRadius:2,background:v.bg,flexShrink:0}}/>
            <span style={{fontSize:9,color:"#7a5560"}}>{k} – {v.label}</span>
          </div>
        ))}
        <span style={{fontSize:9,color:"#7a5560",marginLeft:4}}>· <strong style={{color:"#1a1a2e",textDecoration:"underline dotted"}}>underline</strong> = PM due</span>
      </div>

      {/* Tabs */}
      <div style={{display:"flex",padding:"0 18px",borderBottom:"1px solid #1f2937",overflowX:"auto"}}>
        {TABS.map(([id,label])=>(
          <button key={id} className={`tab ${tab===id?"on":""}`} onClick={()=>setTab(id)}>{label}</button>
        ))}
      </div>

      <div style={{padding:"20px 20px",overflowX:"auto"}}>

        {/* ══ DASHBOARD ══ */}
        {tab==="dash"&&(
          <div style={{display:"flex",flexDirection:"column",gap:28}}>

            {/* MY YARD TODAY */}
            <div>
              <div className="section-title" style={{color:"#1a1a2e"}}>MY YARD TODAY</div>
              <div className="section-sub">Tap card to edit · 🚀 Out = going out today -> Went Out moves to Reso</div>
              <BoardGrid
                data={S.yard}
                renderCard={(card,tt)=><YardCard key={card.id} card={card} tt={tt}/>}
                addCard={tt=>openModal("yard",tt)}
              />
            </div>

            {/* SHORT TERM RESO */}
            <div>
              <div className="section-title" style={{color:"#93c5fd"}}>SHORT TERM RESO</div>
              <div className="section-sub">Carries forward daily · Check In returns unit to yard as WL</div>
              <BoardGrid
                data={S.reso}
                style={{background:"#f0f4ff",borderColor:"#1e3a5f"}}
                renderCard={(card,tt)=><ResoCard key={card.id} card={card} tt={tt}/>}
                addCard={tt=>openModal("reso",tt)}
              />
            </div>

            {/* NEED FOR TOMORROW */}
            <div>
              <div className="section-title" style={{color:"#fcd34d"}}>NEED FOR TOMORROW</div>
              <div className="section-sub">🔴 HOLD = reserved for reso — do not give out</div>
              <div className="grid" style={{background:"#fff7e6",borderColor:"#78350f"}}>
                {TRUCK_TYPES.map(tt=>(
                  <div key={tt}>
                    <div className="col-hdr" style={{background:"#fff7e6",borderBottom:"1px solid #78350f",color:"#92400e"}} title={tt}>{tt}</div>
                    <div className="bcol" style={{background:"#fff9f0"}}>
                      {(S.tomorrow[tt]||[]).map(card=>(
                        <div key={card.id} className="tom-card" onClick={()=>openModal("tomorrow",tt,card)}>
                          <div style={{fontSize:13,fontWeight:700,color:"#fcd34d"}}>{card.unit}</div>
                          {card.note&&<div style={{fontSize:9,color:"#92400e",marginTop:1}}>{card.note}</div>}
                          {card.hold?<span className="hold-badge">🔴 HOLD</span>:<span style={{fontSize:8,color:"#78350f",display:"inline-block",marginTop:2}}>available</span>}
                          <div style={{marginTop:4}} onClick={e=>e.stopPropagation()}>
                            <label className="tog">
                              <input type="checkbox" checked={!!card.hold} onChange={()=>toggleHold(tt,card.id)}/>
                              <span style={{fontSize:9,color:"#92400e"}}>Hold for reso</span>
                            </label>
                          </div>
                          <button className="xcbtn" onClick={e=>{e.stopPropagation();setS(s=>({...s,tomorrow:{...s.tomorrow,[tt]:s.tomorrow[tt].filter(c=>c.id!==card.id)}}));}}>✕</button>
                        </div>
                      ))}
                      <div className="add-btn" style={{borderColor:"#78350f",color:"#78350f"}} onClick={()=>openModal("tomorrow",tt)}>+</div>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* ── SERVICE / SHOP ── */}
            {(() => {
              const serviceUnits = TRUCK_TYPES.flatMap(tt =>
                (S.yard[tt]||[]).filter(c=>c.line==="SL"||c.line==="SHOP").map(c=>({...c,tt})) // SRL excluded — it's ready
              );
              if(serviceUnits.length===0) return null;
              return (
                <div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:8,marginBottom:8}}>
                    <div>
                      <div className="section-title" style={{color:"#f87171"}}>🔧 SUB &amp; DEAD — SERVICE PROGRESS</div>
                      <div className="section-sub">SL = Service Line · SHOP = In shop/deadline · set ready date · mark done when fixed</div>
                    </div>
                    <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#f87171"}}>{serviceUnits.length} unit{serviceUnits.length!==1?"s":""}</div>
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:8}}>
                    {serviceUnits.map(card=>{
                      const ls=card.line==="SHOP"?LINE.SHOP:LINE.SL;
                      const daysLeft=card.shopDate?daysUntil(card.shopDate):null;
                      const overdue=daysLeft!==null&&daysLeft<0;
                      const today=daysLeft===0;
                      return (
                        <div key={card.id} style={{background:"#ffffff",border:`1px solid ${overdue?"#ef4444":today?"#f59e0b":card.line==="SHOP"?"#e8b4bc":"#7f1d1d"}`,borderRadius:9,padding:"12px 14px",display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
                          {/* Line badge */}
                          <div style={{background:ls.bg,color:ls.text,borderRadius:5,padding:"3px 10px",fontSize:11,fontWeight:700,flexShrink:0}}>{card.line}</div>
                          {/* Unit + truck type */}
                          <div>
                            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:card.line==="SHOP"?"#6b4c52":"#f87171"}}>{card.unit}</div>
                            <div style={{fontSize:9,color:"#9c6b75"}}>{card.tt}</div>
                          </div>
                          {/* Note */}
                          {card.note&&<div style={{fontSize:11,color:"#7a5560",flex:1}}>{card.note}</div>}
                          {/* Ready date picker */}
                          <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
                            <span style={{fontSize:10,color:"#9c6b75"}}>Ready date:</span>
                            <input
                              type="date"
                              value={card.shopDate||""}
                              onChange={e=>{
                                const val=e.target.value;
                                setS(s=>({...s,yard:{...s.yard,[card.tt]:s.yard[card.tt].map(c=>c.id===card.id?{...c,shopDate:val}:c)}}));
                              }}
                              onClick={e=>e.stopPropagation()}
                              style={{background:"#f3c0c8",border:"1px solid #374151",color:"#1a1a2e",borderRadius:5,padding:"4px 8px",fontFamily:"inherit",fontSize:11,outline:"none"}}
                            />
                          </div>
                          {/* Countdown */}
                          {daysLeft!==null&&(
                            <div style={{background:overdue?"#7f1d1d":today?"#78350f":"#f3c0c8",color:overdue?"#fca5a5":today?"#f59e0b":"#a07880",borderRadius:5,padding:"3px 10px",fontSize:11,fontWeight:700,flexShrink:0}}>
                              {overdue?`${Math.abs(daysLeft)}d overdue`:today?"Ready TODAY":`${daysLeft}d left`}
                            </div>
                          )}
                          {/* Mark fixed — moves to RL */}
                          <button
                            onClick={()=>{
                              setS(s=>({...s,yard:{...s.yard,[card.tt]:s.yard[card.tt].map(c=>c.id===card.id?{...c,line:"SRL",note:"Fixed — Service Ready",shopDate:""}:c)}}));
                              notify(`Unit ${card.unit} marked fixed — moved to SRL ✓`);
                            }}
                            style={{background:"#1e293b",border:"1px solid #94a3b8",color:"#f1f5f9",borderRadius:5,padding:"4px 10px",fontSize:10,fontWeight:700,cursor:"pointer",fontFamily:"inherit",flexShrink:0}}>
                            ✓ Fixed -> SRL
                          </button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })()}

            {/* ── PUROLATOR ── */}
            {(() => {
              const puroUnits = TRUCK_TYPES.flatMap(tt =>
                (S.yard[tt]||[]).filter(c=>c.isPuro).map(c=>({...c,tt}))
              );
              const puroReso = TRUCK_TYPES.flatMap(tt =>
                (S.reso[tt]||[]).filter(c=>c.customer&&c.customer.toUpperCase().includes("PURO")).map(c=>({...c,tt}))
              );
              if(puroUnits.length===0&&puroReso.length===0) return null;
              return (
                <div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",flexWrap:"wrap",gap:8,marginBottom:8}}>
                    <div>
                      <div className="section-title" style={{color:"#a855f7"}}>🟣 PUROLATOR UNITS</div>
                      <div className="section-sub">All Purolator units on yard + in reso at a glance</div>
                    </div>
                    <div style={{display:"flex",gap:12}}>
                      <div style={{textAlign:"center"}}>
                        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#a855f7"}}>{puroUnits.length}</div>
                        <div style={{fontSize:9,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.06em"}}>On Yard</div>
                      </div>
                      <div style={{textAlign:"center"}}>
                        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#7dd3fc"}}>{puroReso.length}</div>
                        <div style={{fontSize:9,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.06em"}}>In Reso</div>
                      </div>
                    </div>
                  </div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                    {puroUnits.map(card=>(
                      <div key={card.id} style={{background:"#fce7ef",border:"1px solid #7c3aed",borderRadius:8,padding:"10px 14px",minWidth:140}}>
                        <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}>
                          <div style={{width:8,height:8,borderRadius:"50%",background:"#a855f7"}}/>
                          <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#c4b5fd"}}>{card.unit}</span>
                        </div>
                        <div style={{fontSize:9,color:"#6b21a8"}}>{card.tt}</div>
                        <div style={{fontSize:9,color:"#7c3aed",marginTop:2}}>{card.line} · On Yard</div>
                        {card.note&&<div style={{fontSize:9,color:"#9c6b75",marginTop:2}}>{card.note}</div>}
                      </div>
                    ))}
                    {puroReso.map(card=>(
                      <div key={card.id} style={{background:"#0f0a1e",border:"1px solid #4c1d95",borderRadius:8,padding:"10px 14px",minWidth:140,opacity:0.85}}>
                        <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4}}>
                          <div style={{width:8,height:8,borderRadius:"50%",background:"#7c3aed"}}/>
                          <span style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#a78bfa"}}>{card.unit}</span>
                        </div>
                        <div style={{fontSize:9,color:"#6b21a8"}}>{card.tt}</div>
                        <div style={{fontSize:9,color:"#7c3aed",marginTop:2}}>In Reso · back {fmtDate(card.returnDate)}</div>
                        {card.customer&&<div style={{fontSize:9,color:"#9c6b75",marginTop:2}}>{card.customer}</div>}
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}

          </div>
        )}

        {/* ══ PM ══ */}
        {tab==="pm"&&(
          <PMTab S={S} setS={setS} notify={notify} openModal={openModal}/>
        )}

                {/* ══ GROUND UNITS ══ */}
        {tab==="ground"&&(
          <GroundUnitsTab S={S} setS={setS} notify={notify} TRUCK_TYPES={TRUCK_TYPES}/>
        )}

        {/* ══ PUROLATOR FLEET SHEET ══ */}
        {tab==="puro"&&(
          <PuroFleetTab S={S} setS={setS} notify={notify}/>
        )}

        {/* ══ HIKES ══ */}
        {tab==="hikes"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <div><div className="section-title" style={{color:"#67e8f9"}}>HIKE TRACKER</div><div className="section-sub">↓ Inbound = coming to you · ↑ Outbound = sent out</div></div>
              <button className="btn btn-amber btn-sm" onClick={()=>openModal("hike")}>+ Add Hike</button>
            </div>
            {S.hikes.length===0&&<div style={{color:"#e8b4bc",fontSize:12,padding:"24px 0",textAlign:"center"}}>No hikes tracked</div>}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:8}}>
              {S.hikes.map(h=>(
                <div key={h.id} className={`hike-card ${h.dir==="in"?"hike-in":"hike-out"}`}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                    <div>
                      <span style={{fontSize:15,fontWeight:700,color:h.dir==="in"?"#4ade80":"#c084fc"}}>{h.unit}</span>
                      <span style={{fontSize:9,marginLeft:6,color:h.dir==="in"?"#166534":"#6b21a8",background:h.dir==="in"?"#d1fae511":"#f3e8ff11",padding:"1px 5px",borderRadius:3}}>{h.dir==="in"?"↓ IN":"↑ OUT"}</span>
                    </div>
                    <button className="xcbtn" style={{position:"static"}} onClick={()=>setS(s=>({...s,hikes:s.hikes.filter(x=>x.id!==h.id)}))}>✕</button>
                  </div>
                  <div style={{fontSize:9,color:"#7a5560",marginTop:3}}>{h.tt}</div>
                  {h.location&&<div style={{fontSize:10,color:"#6b4c52",marginTop:2}}>{h.dir==="in"?"From":"To"}: {h.location}</div>}
                  {h.arrival&&<div style={{fontSize:10,color:"#f59e0b",marginTop:2}}>📅 {fmtDate(h.arrival)}</div>}
                  {h.note&&<div style={{fontSize:9,color:"#9c6b75",marginTop:3}}>{h.note}</div>}
                  <div style={{marginTop:8,display:"flex",flexDirection:"column",gap:4}}>
                    {[["placed","Hike Placed","#4ade80"],["ready","Unit Ready","#7dd3fc"],["pmDue","PM Due","#fb923c"]].map(([f,l,c])=>(
                      <label key={f} className="tog">
                        <input type="checkbox" checked={!!h[f]} onChange={()=>toggleHikeField(h.id,f)}/>
                        <span style={{fontSize:10,color:h[f]?c:"#a07880",fontWeight:h[f]?"600":"400"}}>{l}</span>
                      </label>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ══ SENT & CI ══ */}
        {tab==="other"&&(
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,maxWidth:800}}>
            <div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <div><div className="section-title" style={{fontSize:17,color:"#a78bfa"}}>NON-REV'D UNITS</div><div className="section-sub">Sent to other locations</div></div>
                <button className="btn btn-amber btn-sm" onClick={()=>openModal("sent")}>+ Add</button>
              </div>
              {S.sent.length===0&&<div style={{color:"#e8b4bc",fontSize:12,padding:"20px 0",textAlign:"center"}}>None sent out</div>}
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {S.sent.map(card=>(
                  <div key={card.id} className="side-card" onClick={()=>openModal("sent",null,card)}>
                    <div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontSize:13,fontWeight:700,color:"#a78bfa"}}>{card.unit}</span><span style={{fontSize:9,color:"#7a5560"}}>{card.tt}</span></div>
                    {card.location&&<div style={{fontSize:10,color:"#7c3aed",marginTop:2}}>-> {card.location}</div>}
                    {card.note&&<div style={{fontSize:9,color:"#9c6b75",marginTop:2}}>{card.note}</div>}
                    <button className="xcbtn" onClick={e=>{e.stopPropagation();setS(s=>({...s,sent:s.sent.filter(c=>c.id!==card.id)}));}}>✕</button>
                  </div>
                ))}
              </div>
            </div>
            <div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <div><div className="section-title" style={{fontSize:17,color:"#34d399"}}>CHECK IN'S</div><div className="section-sub">Off contract — auto-added via CI action</div></div>
                <button className="btn btn-amber btn-sm" onClick={()=>openModal("checkin")}>+ Add</button>
              </div>
              {S.checkins.length===0&&<div style={{color:"#e8b4bc",fontSize:12,padding:"20px 0",textAlign:"center"}}>No check-ins</div>}
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {S.checkins.map(card=>(
                  <div key={card.id} className="side-card" style={{borderColor:"#064e3b"}} onClick={()=>openModal("checkin",null,card)}>
                    <div style={{display:"flex",justifyContent:"space-between"}}><span style={{fontSize:13,fontWeight:700,color:"#34d399"}}>{card.unit}</span><span style={{fontSize:9,color:"#7a5560"}}>{card.tt}</span></div>
                    {(card.hikedFrom||card.customer)&&<div style={{fontSize:10,color:"#059669",marginTop:2}}>✈️ From: {card.hikedFrom||card.customer}</div>}
                    {card.note&&<div style={{fontSize:9,color:"#9c6b75",marginTop:2}}>{card.note}</div>}
                    <button className="xcbtn" onClick={e=>{e.stopPropagation();setS(s=>({...s,checkins:s.checkins.filter(c=>c.id!==card.id)}));}}>✕</button>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ══ TASKS ══ */}
        {tab==="tasks"&&(
          <div style={{maxWidth:600}}>
            <div className="section-title" style={{color:"#a3e635",marginBottom:4}}>DAILY TASKS</div>
            <div className="section-sub">Return reminders auto-appear · check-ins auto-added via CI action</div>
            <div style={{display:"flex",gap:8,marginBottom:16}}>
              <input placeholder="Add a task..." value={taskInput} onChange={e=>setTaskInput(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"){addTask(taskInput);setTaskInput("");}}}/>
              <button className="btn btn-amber" style={{flexShrink:0,padding:"8px 14px"}} onClick={()=>{addTask(taskInput);setTaskInput("");}}>Add</button>
            </div>
            {S.tasks.length===0&&<div style={{color:"#e8b4bc",fontSize:12,padding:"20px 0",textAlign:"center"}}>No tasks yet</div>}
            {[["overdue","🚨 Overdue Units","#ef4444"],["return","⚠️ Return Reminders","#f59e0b"],["pm","🔧 PM Reminders","#fb923c"],["pm-swap","🔄 Swap Checks","#f59e0b"],["checkin","✅ Check In Tasks","#34d399"],["general","General","#a07880"]].map(([type,label,color])=>{
              const group=S.tasks.filter(t=>t.type===type);
              if(group.length===0) return null;
              return (
                <div key={type} style={{marginBottom:16}}>
                  <div style={{fontSize:10,color,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:6}}>{label}</div>
                  {group.map(t=>(
                    <div key={t.id} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"8px 0",borderBottom:"1px solid #1f2937"}}>
                      <input type="checkbox" className="chk-box" checked={t.done} onChange={()=>toggleTask(t.id)}/>
                      <div style={{flex:1}}>
                        <div style={{fontSize:13,color:t.done?"#e8b4bc":"#1a1a2e",textDecoration:t.done?"line-through":"none"}}>
                          {t.unit&&<span style={{color:"#f59e0b",marginRight:5,fontWeight:700}}>#{t.unit}</span>}
                          {t.text}
                        </div>
                      </div>
                      <button style={{background:"none",border:"none",color:"#e8b4bc",cursor:"pointer",fontSize:11}} onClick={()=>delTask(t.id)}>✕</button>
                    </div>
                  ))}
                </div>
              );
            })}
          </div>
        )}

        {/* ══ CONTACTS & INTER-BRANCH TAB ══ */}
        {tab==="contacts"&&(
          <ContactsTab S={S} setS={setS} notify={notify} dayLabel={dayLabel}/>
        )}

        {/* ══ UTILIZATION TAB ══ */}
        {tab==="stats"&&(
          <UtilizationTab S={S} history={history} dayLabel={dayLabel} notify={notify}/>
        )}

      </div>

      {/* ══ WENT OUT MODAL ══ */}
      {goModal&&(
        <div className="overlay" onClick={()=>setGoModal(null)}>
          <div className="modal" style={{background:"#fff7ed",border:"1px solid #f97316",maxWidth:360}} onClick={e=>e.stopPropagation()}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#f97316",marginBottom:4}}>WENT OUT — #{goModal.card.unit}</div>
            <div style={{fontSize:10,color:"#92400e",marginBottom:16}}>{goModal.tt} · unit left the yard · will move to Short Term Reso</div>
            <div className="field"><label>Customer</label><input placeholder="Customer name" value={goForm.customer} onChange={e=>setGoForm(f=>({...f,customer:e.target.value}))}/></div>
            <div className="field">
              <label>Return Date <span style={{color:"#78350f"}}>(default 2 weeks)</span></label>
              <input type="date" value={goForm.returnDate} onChange={e=>setGoForm(f=>({...f,returnDate:e.target.value}))}/>
            </div>
            <div style={{background:"#fdf2f4",borderRadius:6,padding:"8px 10px",fontSize:10,color:"#9c6b75",marginBottom:14}}>
              Unit removed from yard · added to Short Term Reso · return reminder auto-added on due date
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button className="btn btn-ghost" onClick={()=>setGoModal(null)}>Cancel</button>
              <button className="btn" style={{background:"#ea580c",color:"#fff7ed"}} onClick={confirmWentOut}>Confirm Went Out</button>
            </div>
          </div>
        </div>
      )}

      {/* ══ MAIN MODAL ══ */}
      {modal&&(
        <div className="overlay" onClick={closeModal}>
          <div className="modal" onClick={e=>e.stopPropagation()}>

            {modal.type==="yard"&&<>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:17,color:"#f59e0b",marginBottom:14}}>{modal.card?"EDIT":"ADD"} UNIT — {modal.tt}</div>
              <div className="row2">
                <div className="field"><label>Unit #</label><input placeholder="e.g. 529835" value={form.unit||""} onChange={sf("unit")}/></div>
                <div className="field"><label>Line</label>
                  <select value={form.line||"RL"} onChange={sf("line")}>
                    <option value="RL">RL – Ready Line</option>
                    <option value="WL">WL – Wash Line</option>
                    <option value="SRL">SRL – Service Ready Line</option>
                    <option value="SL">SL – Service Line</option>
                    <option value="SHOP">SHOP – Shop/Deadline</option>
                  </select>
                </div>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12}}>
                {[["isPuro","Purolator unit","#a855f7"],["addTomorrow","📅 Need Tomorrow","#fcd34d"],["addPM","🔧 PM Due","#fb923c"],["addCheckin","✅ Check In","#34d399"]].map(([k,l,c])=>(
                  <label key={k} className="tog" style={{background:"#ffffff",border:"1px solid #f3c0c8",borderRadius:6,padding:"6px 10px",cursor:"pointer"}}>
                    <input type="checkbox" checked={!!form[k]} onChange={sf(k)}/>
                    <span style={{fontSize:11,color:form[k]?c:"#a07880"}}>{l}</span>
                  </label>
                ))}
              </div>
              {/* Hike actions — mutually exclusive */}
              <div style={{marginBottom:12}}>
                <div style={{fontSize:10,color:"#9c6b75",textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:6}}>Hike</div>
                <div style={{display:"flex",gap:8}}>
                  <label className="tog" style={{flex:1,background:form.hikeOut?"#f5d0fe":"#ffffff",border:`1px solid ${form.hikeOut?"#7c3aed":"#f3c0c8"}`,borderRadius:6,padding:"8px 10px",cursor:"pointer"}}
                    onClick={()=>setForm(f=>({...f,hikeOut:!f.hikeOut,hikeIn:false}))}>
                    <input type="checkbox" checked={!!form.hikeOut} onChange={()=>{}} style={{accentColor:"#a855f7"}}/>
                    <span style={{fontSize:11,color:form.hikeOut?"#c4b5fd":"#a07880"}}>↑ Hike Out</span>
                  </label>
                  <label className="tog" style={{flex:1,background:form.hikeIn?"#052e16":"#ffffff",border:`1px solid ${form.hikeIn?"#16a34a":"#f3c0c8"}`,borderRadius:6,padding:"8px 10px",cursor:"pointer"}}
                    onClick={()=>setForm(f=>({...f,hikeIn:!f.hikeIn,hikeOut:false}))}>
                    <input type="checkbox" checked={!!form.hikeIn} onChange={()=>{}} style={{accentColor:"#4ade80"}}/>
                    <span style={{fontSize:11,color:form.hikeIn?"#86efac":"#a07880"}}>↓ Hike In</span>
                  </label>
                </div>
                {form.hikeOut&&<div style={{fontSize:10,color:"#a855f7",marginTop:5,padding:"5px 8px",background:"#fce7ef",borderRadius:5}}>Unit will be removed from yard and added to Hikes ↑ outbound</div>}
                {form.hikeIn&&<div style={{fontSize:10,color:"#4ade80",marginTop:5,padding:"5px 8px",background:"#f0fff4",borderRadius:5}}>Unit will stay on yard as Awaiting Arrival and added to Hikes ↓ inbound</div>}
              </div>
              {form.line==="SHOP"&&<div className="field"><label>Expected Out</label><input type="date" value={form.shopDate||""} onChange={sf("shopDate")}/></div>}
              <div className="field"><label>Note</label><textarea placeholder="Any notes..." value={form.note||""} onChange={sf("note")}/></div>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button className="btn btn-ghost" onClick={closeModal}>Cancel</button>
                <button className="btn btn-amber" onClick={saveYard}>{modal.card?"Save":"Add Unit"}</button>
              </div>
            </>}

            {modal.type==="reso"&&<>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:17,color:"#93c5fd",marginBottom:14}}>{modal.card?"EDIT":"ADD"} RESO — {modal.tt}</div>
              <div className="field"><label>Unit #</label><input value={form.unit||""} onChange={sf("unit")}/></div>
              <div className="field"><label>Customer</label><input value={form.customer||""} onChange={sf("customer")}/></div>
              <div className="field"><label>Return Date</label><input type="date" value={form.returnDate||""} onChange={sf("returnDate")}/></div>
              <div className="field"><label>Note</label><textarea value={form.note||""} onChange={sf("note")}/></div>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button className="btn btn-ghost" onClick={closeModal}>Cancel</button>
                <button className="btn btn-amber" onClick={saveReso}>{modal.card?"Save":"Add"}</button>
              </div>
            </>}

            {modal.type==="tomorrow"&&<>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:17,color:"#fcd34d",marginBottom:14}}>{modal.card?"EDIT":"ADD"} TOMORROW — {modal.tt}</div>
              <div className="field"><label>Unit #</label><input value={form.unit||""} onChange={sf("unit")}/></div>
              <label className="tog" style={{marginBottom:12,display:"flex"}}>
                <input type="checkbox" checked={!!form.hold} onChange={sf("hold")}/>
                <span style={{fontSize:12,color:"#fca5a5"}}>🔴 Hold — do not give out</span>
              </label>
              <div className="field"><label>Note</label><textarea value={form.note||""} onChange={sf("note")}/></div>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button className="btn btn-ghost" onClick={closeModal}>Cancel</button>
                <button className="btn btn-amber" onClick={saveTomorrow}>{modal.card?"Save":"Add"}</button>
              </div>
            </>}

            {modal.type==="pm"&&<>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:17,color:"#fb923c",marginBottom:14}}>{modal.card?"EDIT":"SCHEDULE"} PM</div>
              <div className="row2">
                <div className="field"><label>Unit #</label><input value={form.unit||""} onChange={sf("unit")}/></div>
                <div className="field"><label>Truck Type</label>
                  <select value={form.tt||""} onChange={sf("tt")}>
                    <option value="">Select...</option>
                    {TRUCK_TYPES.map(t=><option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
              </div>
              <div className="field"><label>PM Date</label><input type="date" value={form.pmDate||""} onChange={sf("pmDate")}/></div>
              <div className="field"><label>Note</label><textarea value={form.note||""} onChange={sf("note")}/></div>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button className="btn btn-ghost" onClick={closeModal}>Cancel</button>
                <button className="btn btn-amber" onClick={savePM}>{modal.card?"Save":"Schedule"}</button>
              </div>
            </>}

            {modal.type==="hike"&&<>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:17,color:"#67e8f9",marginBottom:14}}>{modal.card?"EDIT":"ADD"} HIKE</div>
              <div className="row2">
                <div className="field"><label>Unit #</label><input value={form.unit||""} onChange={sf("unit")}/></div>
                <div className="field"><label>Direction</label>
                  <select value={form.dir||"in"} onChange={sf("dir")}>
                    <option value="in">↓ Inbound</option>
                    <option value="out">↑ Outbound</option>
                  </select>
                </div>
              </div>
              <div className="row2">
                <div className="field"><label>Truck Type</label>
                  <select value={form.tt||""} onChange={sf("tt")}>
                    <option value="">Select...</option>
                    {TRUCK_TYPES.map(t=><option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div className="field"><label>{form.dir==="out"?"To":"From"} Location</label><input value={form.location||""} onChange={sf("location")}/></div>
              </div>
              <div className="field"><label>Expected Date</label><input type="date" value={form.arrival||""} onChange={sf("arrival")}/></div>
              <div style={{display:"flex",gap:12,marginBottom:10,flexWrap:"wrap"}}>
                {[["placed","Hike Placed"],["ready","Unit Ready"],["pmDue","PM Due"]].map(([k,l])=>(
                  <label key={k} className="tog">
                    <input type="checkbox" checked={!!form[k]} onChange={sf(k)}/>
                    <span style={{fontSize:12,color:"#6b4c52"}}>{l}</span>
                  </label>
                ))}
              </div>
              {form.pmDue&&<div style={{fontSize:10,color:"#fb923c",marginBottom:10,background:"#fffbeb",borderRadius:5,padding:"6px 10px"}}>⚠️ PM Due auto-adds to PM Schedule</div>}
              <div className="field"><label>Note</label><textarea value={form.note||""} onChange={sf("note")}/></div>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button className="btn btn-ghost" onClick={closeModal}>Cancel</button>
                <button className="btn btn-amber" onClick={saveHike}>{modal.card?"Save":"Add Hike"}</button>
              </div>
            </>}

            {modal.type==="sent"&&<>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:17,color:"#a78bfa",marginBottom:14}}>{modal.card?"EDIT":"ADD"} NON-REV'D</div>
              <div className="row2">
                <div className="field"><label>Unit #</label><input value={form.unit||""} onChange={sf("unit")}/></div>
                <div className="field"><label>Truck Type</label>
                  <select value={form.tt||""} onChange={sf("tt")}>
                    <option value="">Select...</option>
                    {TRUCK_TYPES.map(t=><option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
              </div>
              <div className="field"><label>Sent To</label><input value={form.location||""} onChange={sf("location")}/></div>
              <div className="field"><label>Note</label><textarea value={form.note||""} onChange={sf("note")}/></div>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button className="btn btn-ghost" onClick={closeModal}>Cancel</button>
                <button className="btn btn-amber" onClick={saveSent}>{modal.card?"Save":"Add"}</button>
              </div>
            </>}

            {modal.type==="checkin"&&<>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:17,color:"#34d399",marginBottom:14}}>{modal.card?"EDIT":"ADD"} CHECK IN</div>
              <div className="row2">
                <div className="field"><label>Unit #</label><input value={form.unit||""} onChange={sf("unit")}/></div>
                <div className="field"><label>Truck Type</label>
                  <select value={form.tt||""} onChange={sf("tt")}>
                    <option value="">Select...</option>
                    {TRUCK_TYPES.map(t=><option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
              </div>
              <div className="field"><label>Hiked From</label><input placeholder="e.g. Concord" value={form.customer||""} onChange={sf("customer")}/></div>
              <div className="field"><label>Note</label><textarea value={form.note||""} onChange={sf("note")}/></div>
              <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
                <button className="btn btn-ghost" onClick={closeModal}>Cancel</button>
                <button className="btn btn-amber" onClick={saveCheckin}>{modal.card?"Save":"Add"}</button>
              </div>
            </>}

          </div>
        </div>
      )}

      {/* ══ HIKE IN SOURCE MODAL ══ */}
      {hikeInModal&&(
        <div className="overlay" onClick={()=>setHikeInModal(null)}>
          <div className="modal" style={{background:"#f0fff4",border:"1px solid #16a34a",maxWidth:340}} onClick={e=>e.stopPropagation()}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#4ade80",marginBottom:4}}>HIKE IN — #{hikeInModal.card.unit}</div>
            <div style={{fontSize:10,color:"#166534",marginBottom:14}}>{hikeInModal.tt} · where is this unit coming from?</div>
            <div className="field">
              <label>Coming From</label>
              <input
                placeholder="e.g. Concord, Belfield..."
                value={hikeInFrom}
                onChange={e=>setHikeInFrom(e.target.value)}
                onKeyDown={e=>e.key==="Enter"&&confirmHikeIn()}
                autoFocus
              />
            </div>
            <div style={{fontSize:10,color:"#9c6b75",marginBottom:14,padding:"6px 8px",background:"#fdf2f4",borderRadius:5}}>
              Unit stays on yard as Awaiting Arrival · added to Hikes ↓ inbound
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button className="btn btn-ghost" onClick={()=>setHikeInModal(null)}>Cancel</button>
              <button className="btn" style={{background:"#16a34a",color:"#f0fdf4"}} onClick={confirmHikeIn}>Confirm Hike In</button>
            </div>
          </div>
        </div>
      )}

      {/* ══ HIKE OUT DESTINATION MODAL ══ */}
      {hikeOutModal&&(
        <div className="overlay" onClick={()=>setHikeOutModal(null)}>
          <div className="modal" style={{background:"#fff0f6",border:"1px solid #7c3aed",maxWidth:340}} onClick={e=>e.stopPropagation()}>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#c4b5fd",marginBottom:4}}>HIKE OUT — #{hikeOutModal.card.unit}</div>
            <div style={{fontSize:10,color:"#6b21a8",marginBottom:14}}>{hikeOutModal.tt} · where is this unit going?</div>
            <div className="field">
              <label>Destination Location</label>
              <input
                placeholder="e.g. Concord, Belfield..."
                value={hikeOutDest}
                onChange={e=>setHikeOutDest(e.target.value)}
                onKeyDown={e=>e.key==="Enter"&&confirmHikeOut()}
                autoFocus
              />
            </div>
            <div style={{fontSize:10,color:"#9c6b75",marginBottom:14,padding:"6px 8px",background:"#fdf2f4",borderRadius:5}}>
              Unit will be removed from yard · added to Hikes ↑ and Sent Out
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button className="btn btn-ghost" onClick={()=>setHikeOutModal(null)}>Cancel</button>
              <button className="btn" style={{background:"#7c3aed",color:"#f5f3ff"}} onClick={confirmHikeOut}>Confirm Hike Out</button>
            </div>
          </div>
        </div>
      )}

      {/* Toast notification */}
      {notification&&<div className="notif">{notification}</div>}

      {/* ══ HISTORY LIST MODAL ══ */}
      {histOpen&&!historyViewDay&&(
        <div className="overlay" onClick={()=>setHistOpen(false)}>
          <div className="modal" style={{maxWidth:520,maxHeight:"85vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,color:"#f59e0b",letterSpacing:"0.08em"}}>OPERATIONS HISTORY</div>
              <button onClick={()=>setHistOpen(false)} style={{background:"none",border:"none",color:"#9c6b75",cursor:"pointer",fontSize:18}}>✕</button>
            </div>
            {history.length===0&&(
              <div style={{textAlign:"center",padding:"32px 0",color:"#e8b4bc",fontSize:12}}>
                No history yet — hit 🌅 New Day to save today's operations
              </div>
            )}
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {[...history].reverse().map(h=>{
                const yardTotal = TRUCK_TYPES.reduce((a,t)=>a+(h.snap.yard[t]||[]).length,0);
                const resoTotal = TRUCK_TYPES.reduce((a,t)=>a+(h.snap.reso[t]||[]).length,0);
                const wentOut   = TRUCK_TYPES.reduce((a,t)=>a+(h.snap.yard[t]||[]).filter(c=>c.wentOut).length,0);
                const tasksDone = (h.snap.tasks||[]).filter(t=>t.done).length;
                const tasksTotal= (h.snap.tasks||[]).length;
                const pmDone    = (h.snap.pmRows||[]).filter(r=>r.status==="done").length;
                return (
                  <div key={h.dayNum} style={{background:"#f3c0c8",border:"1px solid #374151",borderRadius:9,padding:"12px 14px",cursor:"pointer",transition:"border-color 0.15s"}}
                    onMouseEnter={e=>e.currentTarget.style.borderColor="#f59e0b"}
                    onMouseLeave={e=>e.currentTarget.style.borderColor="#e8b4bc"}
                    onClick={()=>{setHistoryViewDay(h);setHistOpen(false);}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                      <div>
                        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:18,color:"#f59e0b"}}>Day {h.dayNum}</div>
                        <div style={{fontSize:11,color:"#7a5560",marginTop:1}}>{h.label}</div>
                      </div>
                      <span style={{fontSize:10,color:"#f59e0b",marginTop:4}}>View -></span>
                    </div>
                    <div style={{display:"flex",gap:16,marginTop:10,flexWrap:"wrap"}}>
                      {[
                        ["🚛 On Yard", yardTotal, "#7dd3fc"],
                        ["📋 In Reso", resoTotal, "#f59e0b"],
                        ["🚀 Went Out", wentOut, "#f97316"],
                        ["✅ Tasks", `${tasksDone}/${tasksTotal}`, tasksDone===tasksTotal&&tasksTotal>0?"#4ade80":"#a07880"],
                        ["🔧 PM Done", pmDone, "#fb923c"],
                      ].map(([l,v,c])=>(
                        <div key={l} style={{textAlign:"center"}}>
                          <div style={{fontSize:14,fontWeight:700,color:c}}>{v}</div>
                          <div style={{fontSize:9,color:"#9c6b75"}}>{l}</div>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* ══ HISTORY DAY DETAIL MODAL ══ */}
      {historyViewDay&&(
        <div className="overlay" onClick={()=>setHistoryViewDay(null)}>
          <div style={{background:"#fdf2f4",border:"1px solid #f3c0c8",borderRadius:12,width:"100%",maxWidth:980,maxHeight:"92vh",overflowY:"auto",padding:20}} onClick={e=>e.stopPropagation()}>
            {/* Header */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18,paddingBottom:12,borderBottom:"1px solid #1f2937"}}>
              <div>
                <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:22,color:"#f59e0b",letterSpacing:"0.08em"}}>DAY {historyViewDay.dayNum} — SNAPSHOT</div>
                <div style={{fontSize:11,color:"#9c6b75",marginTop:2}}>{historyViewDay.label} · read-only</div>
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <button onClick={()=>{setHistoryViewDay(null);setHistOpen(true);}} style={{background:"#f3c0c8",border:"1px solid #374151",color:"#6b4c52",borderRadius:6,padding:"5px 12px",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}> Back</button>
                <button onClick={()=>setHistoryViewDay(null)} style={{background:"none",border:"none",color:"#9c6b75",cursor:"pointer",fontSize:18}}>✕</button>
              </div>
            </div>

            {(() => {
              const h = historyViewDay.snap;
              const LINE_H = { RL:{bg:"#84cc16",text:"#1a2e05"}, WL:{bg:"#7dd3fc",text:"#0c2a3e"}, SRL:{bg:"#f1f5f9",text:"#0f172a"}, SL:{bg:"#f87171",text:"#3b0a0a"}, SHOP:{bg:"#e8b4bc",text:"#f9fafb"}, PUR:{bg:"#a855f7",text:"#f5f3ff"} };

              const Section = ({title,color,children}) => (
                <div style={{marginBottom:20}}>
                  <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:16,color:color||"#1a1a2e",letterSpacing:"0.06em",marginBottom:8,paddingBottom:4,borderBottom:"1px solid #1f2937"}}>{title}</div>
                  {children}
                </div>
              );

              const yardUnits = TRUCK_TYPES.flatMap(tt=>(h.yard[tt]||[]).map(c=>({...c,tt})));
              const resoUnits = TRUCK_TYPES.flatMap(tt=>(h.reso[tt]||[]).map(c=>({...c,tt})));
              const tomUnits  = TRUCK_TYPES.flatMap(tt=>(h.tomorrow[tt]||[]).map(c=>({...c,tt})));
              const tasks     = h.tasks||[];
              const pmRows    = h.pmRows||[];
              const hikes     = h.hikes||[];
              const sent      = h.sent||[];

              return (
                <div>
                  {/* Yard */}
                  <Section title={`🚛 My Yard (${yardUnits.length} units)`} color="#7dd3fc">
                    {yardUnits.length===0?<div style={{fontSize:11,color:"#e8b4bc"}}>No units on yard</div>:(
                      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                        {yardUnits.map(c=>{
                          const ls=c.isPuro?LINE_H.PUR:(LINE_H[c.line]||LINE_H.RL);
                          return (
                            <div key={c.id} style={{background:ls.bg,color:ls.text,borderRadius:6,padding:"5px 10px",fontSize:11}}>
                              <div style={{fontWeight:700}}>{c.unit}</div>
                              <div style={{fontSize:9,opacity:0.75}}>{c.isPuro?"PURO":c.line} · {c.tt}</div>
                              {c.wentOut&&<div style={{fontSize:8,fontWeight:700}}>✅ WENT OUT</div>}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </Section>

                  {/* Reso */}
                  <Section title={`📋 Short Term Reso (${resoUnits.length} units)`} color="#93c5fd">
                    {resoUnits.length===0?<div style={{fontSize:11,color:"#e8b4bc"}}>No units in reso</div>:(
                      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                        {resoUnits.map(c=>(
                          <div key={c.id} style={{background:"#f0f4ff",border:"1px solid #1e3a5f",borderRadius:6,padding:"6px 10px",fontSize:11}}>
                            <div style={{fontWeight:700,color:"#93c5fd"}}>{c.unit}</div>
                            {c.customer&&<div style={{fontSize:9,color:"#7dd3fc"}}>{c.customer}</div>}
                            {c.returnDate&&<div style={{fontSize:9,color:"#f59e0b"}}>Back {fmtDate(c.returnDate)}</div>}
                            <div style={{fontSize:9,color:"#9c6b75"}}>{c.tt}</div>
                          </div>
                        ))}
                      </div>
                    )}
                  </Section>

                  {/* Tomorrow */}
                  {tomUnits.length>0&&(
                    <Section title={`📅 Need for Tomorrow (${tomUnits.length})`} color="#fcd34d">
                      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                        {tomUnits.map(c=>(
                          <div key={c.id} style={{background:"#fff7e6",border:"1px solid #78350f",borderRadius:6,padding:"5px 10px",fontSize:11}}>
                            <div style={{fontWeight:700,color:"#fcd34d"}}>{c.unit}</div>
                            <div style={{fontSize:9,color:"#92400e"}}>{c.tt}{c.hold?" · 🔴 HOLD":""}</div>
                          </div>
                        ))}
                      </div>
                    </Section>
                  )}

                  {/* Tasks */}
                  <Section title={`✅ Daily Tasks (${tasks.filter(t=>t.done).length}/${tasks.length} done)`} color="#a3e635">
                    {tasks.length===0?<div style={{fontSize:11,color:"#e8b4bc"}}>No tasks</div>:(
                      <div style={{display:"flex",flexDirection:"column",gap:5}}>
                        {tasks.map(t=>(
                          <div key={t.id} style={{display:"flex",alignItems:"center",gap:8,fontSize:12,color:t.done?"#e8b4bc":"#6b4c52",textDecoration:t.done?"line-through":"none"}}>
                            <span style={{fontSize:14}}>{t.done?"✅":"⬜"}</span>
                            {t.unit&&<span style={{color:"#f59e0b",fontWeight:700}}>#{t.unit}</span>}
                            {t.text}
                          </div>
                        ))}
                      </div>
                    )}
                  </Section>

                  {/* PM */}
                  {pmRows.length>0&&(
                    <Section title={`🔧 PM Checklist (${pmRows.filter(r=>r.status==="done").length} done)`} color="#fb923c">
                      <div style={{display:"flex",flexDirection:"column",gap:5}}>
                        {pmRows.map(r=>(
                          <div key={r.id} style={{display:"flex",alignItems:"center",gap:10,fontSize:12,background:"#f3c0c8",borderRadius:5,padding:"6px 10px",opacity:r.status==="done"?0.6:1}}>
                            <span style={{fontSize:13}}>{r.status==="done"?"✅":r.status==="scheduled"?"📅":"⬜"}</span>
                            <span style={{fontWeight:700,color:r.status==="done"?"#4ade80":r.status==="scheduled"?"#34d399":"#fb923c"}}>{r.unit}</span>
                            <span style={{color:"#7a5560",fontSize:10}}>{r.pmType} · {r.customer}</span>
                            {r.scheduledDate&&<span style={{color:"#f59e0b",fontSize:10,marginLeft:"auto"}}>📅 {fmtDate(r.scheduledDate)}</span>}
                          </div>
                        ))}
                      </div>
                    </Section>
                  )}

                  {/* Hikes */}
                  {hikes.length>0&&(
                    <Section title={`✈️ Hikes (${hikes.length})`} color="#67e8f9">
                      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                        {hikes.map(h=>(
                          <div key={h.id} style={{background:h.dir==="in"?"#0a1f12":"#fff0f6",border:`1px solid ${h.dir==="in"?"#166534":"#6b21a8"}`,borderRadius:6,padding:"5px 10px",fontSize:11}}>
                            <div style={{fontWeight:700,color:h.dir==="in"?"#4ade80":"#c084fc"}}>{h.unit} {h.dir==="in"?"↓":"↑"}</div>
                            <div style={{fontSize:9,color:"#9c6b75"}}>{h.location||"—"}</div>
                          </div>
                        ))}
                      </div>
                    </Section>
                  )}

                  {/* Sent */}
                  {sent.length>0&&(
                    <Section title={`📤 Sent Out (${sent.length})`} color="#a78bfa">
                      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
                        {sent.map(c=>(
                          <div key={c.id} style={{background:"#f3c0c8",border:"1px solid #374151",borderRadius:6,padding:"5px 10px",fontSize:11}}>
                            <div style={{fontWeight:700,color:"#a78bfa"}}>{c.unit}</div>
                            <div style={{fontSize:9,color:"#7a5560"}}>{c.location||"—"}</div>
                          </div>
                        ))}
                      </div>
                    </Section>
                  )}
                </div>
              );
            })()}
          </div>
        </div>
      )}

    </div>
  );
}
