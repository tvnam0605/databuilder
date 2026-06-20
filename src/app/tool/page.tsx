"use client";
import { useSession, signOut } from "next-auth/react";
import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Types ──────────────────────────────────────────────────────────────
type Row = Record<string, unknown>;

type MappedRow = {
  "TO Number": string;
  "SPX Tracking Number": string;
  "TO Status": string;
  "Receive Status": string;
  "Sender Name": string;
  "Receiver Name": string;
  "Complete Time": string;
  "LH Trip Number": string;
  "LH Trip Name": string;
  "Vehicle Plate Number": string;
  Driver: string;
  "Station Name": string;
  "Actual Arrival Time": string;
  "Actual Departure Time": string;
  _source: string;
  _mapped: boolean;
};

type DS = {
  all: Row[];
  view: Row[];
  cols: string[];
  vis: Set<string>;
  page: number;
  sortCol: string | null;
  sortAsc: boolean;
};

// ── Constants ──────────────────────────────────────────────────────────
const PAGE_SIZE = 15;
const TO_PRI = [
  "TO Number","SPX Tracking Number","TO Status","Receive Status",
  "Sender Name","Receiver Name","Current Station","Complete Time",
  "Line Hual Trip Number","TO Order Quantity","Weight","Remark",
];
const TRIP_PRI = [
  "LH Trip Number","LH Trip Name","Vehicle Plate Number","Driver",
  "Station Number","Station Name","Schedule Arrival Time","Actual Arrival Time",
  "Actual Departure Time","Inbound(TO)","Outbound(TO)","Inbound(order)","Outbound(order)",
  "Occupancy Rate","Cost Type","Late Arrival Status","Remark",
];
const MAP_COLS = [
  "TO Number","SPX Tracking Number","TO Status","Receive Status",
  "Sender Name","Receiver Name","Complete Time",
  "LH Trip Number","LH Trip Name","Vehicle Plate Number","Driver",
  "Station Name","Actual Arrival Time","Actual Departure Time","_source",
];
const BADGE_CLS: Record<string, string> = {
  Received: "bg", Abnormal: "br", "In Transit": "bb",
  Created: "bgr", Completed: "bg", Forward: "bo", Transporting: "bo",
};

// ── Helpers ────────────────────────────────────────────────────────────
const str = (x: unknown) => (x === null || x === undefined ? "" : String(x));
const esc = (t: unknown) =>
  str(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
const cleanTrip = (s: string) => s.replace(/^'+/, "").split("/")[0].trim();

// Tách paste input thành tokens (space / newline / comma / tab)
const splitTokens = (s: string): string[] =>
  s.split(/[\s,\t\n]+/).map(t => t.trim()).filter(Boolean);

// 1 token → contains; nhiều token → exact match any (case-insensitive)
const matchMulti = (val: string, filter: string): boolean => {
  if (!filter.trim()) return true;
  const tokens = splitTokens(filter);
  if (tokens.length === 0) return true;
  const v = val.toLowerCase();
  if (tokens.length === 1) return v.includes(tokens[0].toLowerCase());
  return tokens.some(t => v === t.toLowerCase());
};
const fmtDate = (v: unknown): string => {
  if (!v) return "";
  const d = v instanceof Date ? v : new Date(str(v));
  return isNaN(d.getTime()) ? str(v) : d.toLocaleString("en-GB");
};
const deduped = <T,>(arr: T[]) => [...new Set(arr)];
const orderedVis = (cols: string[], pri: string[]) =>
  new Set([...pri.filter(c => cols.includes(c)), ...cols.filter(c => !pri.includes(c))]);
const dateStr = () => {
  const d = new Date();
  const date = d.toISOString().slice(0,10).replace(/-/g,"");
  const time = d.toTimeString().slice(0,8).replace(/:/g,"");
  return `${date}_${time}`;
};

function parseBuffer(buf: ArrayBuffer): Row[] {
  try {
    const wb = XLSX.read(new Uint8Array(buf), { type:"array", cellDates:true, raw:false });
    return XLSX.utils.sheet_to_json<Row>(wb.Sheets[wb.SheetNames[0]], { defval:"" });
  } catch { return []; }
}

function parseCsvBuf(buf: ArrayBuffer): Row[] {
  try {
    const text = new TextDecoder("utf-8").decode(new Uint8Array(buf));
    const wb = XLSX.read(text, { type:"string", cellDates:true });
    return XLSX.utils.sheet_to_json<Row>(wb.Sheets[wb.SheetNames[0]], { defval:"" });
  } catch { return []; }
}

async function readFile(f: File): Promise<Row[]> {
  const name = f.name.toLowerCase();
  const buf = await f.arrayBuffer();
  if (name.endsWith(".xlsx")) return parseBuffer(buf);
  if (name.endsWith(".csv"))  return parseCsvBuf(buf);
  if (name.endsWith(".zip")) {
    // dynamic import jszip at runtime
    const JSZip = (await import("jszip")).default;
    const zip = await JSZip.loadAsync(buf);
    const rows: Row[] = [];
    for (const [path, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue;
      const lp = path.toLowerCase();
      const ab = await entry.async("arraybuffer");
      if (lp.endsWith(".xlsx")) rows.push(...parseBuffer(ab));
      else if (lp.endsWith(".csv")) rows.push(...parseCsvBuf(ab));
    }
    return rows;
  }
  return [];
}

// pageRange
function pageRange(cur: number, tot: number): (number|"...")[] {
  if (tot <= 7) return Array.from({length:tot},(_,i)=>i+1);
  if (cur<=4)     return [1,2,3,4,5,"...",tot];
  if (cur>=tot-3) return [1,"...",tot-4,tot-3,tot-2,tot-1,tot];
  return [1,"...",cur-1,cur,cur+1,"...",tot];
}

function makeDS(): DS {
  return { all:[], view:[], cols:[], vis:new Set(), page:1, sortCol:null, sortAsc:true };
}

// ── Component ──────────────────────────────────────────────────────────
export default function ToolPage() {
  const { data: session } = useSession();

  // Raw data
  const [rawTO, setRawTO]   = useState<Row[]>([]);
  const [rawHO, setRawHO]   = useState<Row[]>([]);
  const [rawEN, setRawEN]   = useState<Row[]>([]);

  // Upload state
  const [uState, setUState] = useState<Record<"to"|"ho"|"en", "idle"|"ok"|"err">>({ to:"idle", ho:"idle", en:"idle" });
  const [uInfo, setUInfo]   = useState<Record<"to"|"ho"|"en", string>>({ to:"", ho:"", en:"" });
  const [prog, setProg]     = useState<{pct:number; lbl:string}|null>(null);

  // Datasets
  const [dsTO, setDsTO]     = useState<DS>(makeDS());
  const [dsMap, setDsMap]   = useState<DS>(makeDS());
  const [dsTrip, setDsTrip] = useState<DS>(makeDS());

  // Stats
  const [stats, setStats]   = useState({ tot:0, rec:0, abn:0, rvr:0, snd:0, map:0 });
  const [mapRatio, setMapRatio] = useState("");

  // Filter fields – TO
  const [fTO, setFTO]   = useState({ to:"", spx:"", rec:"", snd:"", tos:"", rs:"", dt:"" });
  // Filter fields – Map
  const [fMap, setFMap] = useState({ to:"", trp:"", plt:"", drv:"", rcv:"", src:"", ms:"" });
  // Filter fields – Trip
  const [fTrip, setFTrip] = useState({ trp:"", plt:"", drv:"", stn:"", src:"" });

  // Dropdown options
  const [toStatuses, setToStatuses]  = useState<string[]>([]);
  const [rcvStatuses, setRcvStatuses] = useState<string[]>([]);

  // Tab + modal
  const [tab, setTab]             = useState<"to"|"map"|"trip"|"manual">("to");
  const [sumOpen, setSumOpen]     = useState(false);
  const [expOpen, setExpOpen]     = useState(false);
  const [expType, setExpType]     = useState<string>("all");

  // Toast
  const [toast, setToast]   = useState<{msg:string;cls:string}|null>(null);
  const toastRef             = useRef<ReturnType<typeof setTimeout>|null>(null);

  const showToast = (msg:string, cls="t-inf") => {
    if (toastRef.current) clearTimeout(toastRef.current);
    setToast({msg,cls});
    toastRef.current = setTimeout(()=>setToast(null), 2700);
  };

  // pill toggle ref cache
  const [toVis, setToVis]   = useState<Set<string>>(new Set());
  const [mapVis, setMapVis] = useState<Set<string>>(new Set());
  const [tripVis, setTripVis] = useState<Set<string>>(new Set());

  // ── Upload handler ──────────────────────────────────────────────────
  const handleUpload = useCallback(async (files: FileList, key:"to"|"ho"|"en") => {
    setProg({pct:0, lbl:"Đang đọc file..."});
    try {
      let rows: Row[] = [];
      for (let i=0; i<files.length; i++) {
        rows = rows.concat(await readFile(files[i]));
        setProg({pct:Math.round((i+1)/files.length*100), lbl:`${i+1}/${files.length} file...`});
      }
      if (key==="to") setRawTO(rows);
      if (key==="ho") setRawHO(rows);
      if (key==="en") setRawEN(rows);
      setUState(s=>({...s,[key]: rows.length?"ok":"err"}));
      setUInfo(s=>({...s,[key]: rows.length ? `${rows.length.toLocaleString()} dòng` : "Không đọc được"}));
    } catch {
      setUState(s=>({...s,[key]:"err"}));
      setUInfo(s=>({...s,[key]:"Lỗi đọc file"}));
    }
    setProg(null);
  }, []);

  const clearDS = (key:"to"|"ho"|"en") => {
    if (key==="to") setRawTO([]);
    if (key==="ho") setRawHO([]);
    if (key==="en") setRawEN([]);
    setUState(s=>({...s,[key]:"idle"}));
    setUInfo(s=>({...s,[key]:""}));
  };

  // ── Rebuild when raw data changes ──────────────────────────────────
  useEffect(()=>{
    // TO dataset
    if (rawTO.length===0) { setDsTO(makeDS()); return; }
    const cols = Object.keys(rawTO[0]??{});
    const vis  = orderedVis(cols, TO_PRI);
    setToVis(vis);
    setToStatuses([...new Set(rawTO.map(r=>str(r["TO Status"])).filter(Boolean))].sort());
    setRcvStatuses([...new Set(rawTO.map(r=>str(r["Receive Status"])).filter(Boolean))].sort());
    const ds: DS = {all:rawTO, view:rawTO, cols, vis, page:1, sortCol:null, sortAsc:true};
    setDsTO(ds);
  }, [rawTO]);

  useEffect(()=>{
    if (rawHO.length===0 && rawEN.length===0) { setDsTrip(makeDS()); return; }
    const allTrips: Row[] = [
      ...rawHO.map(r=>({...r, _source:"handover"})),
      ...rawEN.map(r=>({...r, _source:"ended"})),
    ];
    const cols = deduped([...TRIP_PRI,"_source",...Object.keys(allTrips[0]??{})].filter(c=>allTrips.some(r=>c in r)));
    const vis  = orderedVis(cols, TRIP_PRI);
    setTripVis(vis);
    setDsTrip({all:allTrips, view:allTrips, cols, vis, page:1, sortCol:null, sortAsc:true});
  }, [rawHO, rawEN]);

  useEffect(()=>{
    if (rawTO.length===0 || (rawHO.length===0 && rawEN.length===0)) { setDsMap(makeDS()); return; }
    const allTrips: Row[] = [
      ...rawHO.map(r=>({...r, _source:"handover"})),
      ...rawEN.map(r=>({...r, _source:"ended"})),
    ];
    const lookup: Record<string,Row> = {};
    allTrips.forEach(tr=>{
      const tn = cleanTrip(str(tr["LH Trip Number"]));
      if (!tn) return;
      const sn = parseInt(str(tr["Station Number"]))||999;
      if (!lookup[tn]||(parseInt(str(lookup[tn]["Station Number"]))||999)>sn) lookup[tn]=tr;
    });
    let mapped=0;
    const mappedRows: Row[] = rawTO.map(to=>{
      const tn   = cleanTrip(str(to["Line Hual Trip Number"]));
      const trip = tn ? lookup[tn] : null;
      if (trip) mapped++;
      return {
        "TO Number":             str(to["TO Number"]),
        "SPX Tracking Number":   str(to["SPX Tracking Number"]),
        "TO Status":             str(to["TO Status"]),
        "Receive Status":        str(to["Receive Status"]),
        "Sender Name":           str(to["Sender Name"]),
        "Receiver Name":         str(to["Receiver Name"]),
        "Complete Time":         fmtDate(to["Complete Time"]),
        "LH Trip Number":        trip ? str(trip["LH Trip Number"]) : tn,
        "LH Trip Name":          trip ? str(trip["LH Trip Name"])   : "",
        "Vehicle Plate Number":  trip ? str(trip["Vehicle Plate Number"]) : "",
        "Driver":                trip ? str(trip["Driver"])          : "",
        "Station Name":          trip ? str(trip["Station Name"])    : "",
        "Actual Arrival Time":   trip ? fmtDate(trip["Actual Arrival Time"])   : "",
        "Actual Departure Time": trip ? fmtDate(trip["Actual Departure Time"]) : "",
        "_source":               trip ? str(trip["_source"]) : "",
        "_mapped":               !!trip,
      } as Row;
    });
    const pct = ((mapped/(rawTO.length||1))*100).toFixed(1);
    setMapRatio(`${mapped.toLocaleString()} / ${rawTO.length.toLocaleString()} TO mapped (${pct}%)`);
    const vis = new Set(MAP_COLS);
    setMapVis(vis);
    setDsMap({all:mappedRows, view:mappedRows, cols:MAP_COLS, vis, page:1, sortCol:null, sortAsc:true});
  }, [rawTO, rawHO, rawEN]);

  // ── Apply filters ──────────────────────────────────────────────────
  const applyFilter = useCallback((key:"to"|"map"|"trip", ds:DS, setDs: React.Dispatch<React.SetStateAction<DS>>) => {
    let rows = ds.all.slice();

    if (key==="to") {
      const {to,spx,rec,snd,tos,rs,dt} = fTO;
      rows = rows.filter(r=>{
        if (!matchMulti(str(r["TO Number"]), to)) return false;
        if (!matchMulti(str(r["SPX Tracking Number"]), spx)) return false;
        if (!matchMulti(str(r["Receiver Name"]), rec)) return false;
        if (!matchMulti(str(r["Sender Name"]), snd)) return false;
        if (tos && str(r["TO Status"])!==tos) return false;
        if (rs  && str(r["Receive Status"])!==rs) return false;
        if (dt) {
          const v = r["Complete Time"];
          const d2 = v instanceof Date ? v : new Date(str(v));
          if (isNaN(d2.getTime())||d2.toISOString().slice(0,10)!==dt) return false;
        }
        return true;
      });
      const rec2=rows.filter(r=>r["Receive Status"]==="Received").length;
      const abn=rows.filter(r=>r["Receive Status"]==="Abnormal").length;
      setStats({
        tot:rows.length, rec:rec2, abn, map:0,
        rvr:new Set(rows.map(r=>r["Receiver Name"])).size,
        snd:new Set(rows.map(r=>r["Sender Name"])).size,
      });
    }
    if (key==="map") {
      const {to,trp,plt,drv,rcv,src,ms} = fMap;
      rows = rows.filter(r=>{
        if (!matchMulti(str(r["TO Number"]), to)) return false;
        if (!matchMulti(str(r["LH Trip Number"]), trp)) return false;
        if (!matchMulti(str(r["Vehicle Plate Number"]), plt)) return false;
        if (!matchMulti(str(r["Driver"]), drv)) return false;
        if (!matchMulti(str(r["Receiver Name"]), rcv)) return false;
        if (src && str(r["_source"])!==src) return false;
        if (ms==="1" && !r["_mapped"]) return false;
        if (ms==="0" && r["_mapped"]) return false;
        return true;
      });
    }
    if (key==="trip") {
      const {trp,plt,drv,stn,src} = fTrip;
      rows = rows.filter(r=>{
        if (!matchMulti(str(r["LH Trip Number"]), trp)) return false;
        if (!matchMulti(str(r["Vehicle Plate Number"]), plt)) return false;
        if (!matchMulti(str(r["Driver"]), drv)) return false;
        if (!matchMulti(str(r["Station Name"]), stn)) return false;
        if (src && str(r["_source"])!==src) return false;
        return true;
      });
    }

    if (ds.sortCol) {
      rows.sort((a,b)=>{
        let av=a[ds.sortCol!], bv=b[ds.sortCol!];
        if (av instanceof Date && bv instanceof Date) return ds.sortAsc ? (av as Date).getTime()-(bv as Date).getTime() : (bv as Date).getTime()-(av as Date).getTime();
        const as2=str(av), bs2=str(bv);
        return ds.sortAsc ? as2.localeCompare(bs2,"vi") : bs2.localeCompare(as2,"vi");
      });
    }
    setDs(d=>({...d, view:rows, page:1}));
  }, [fTO, fMap, fTrip]);

  useEffect(()=>{ if (dsTO.all.length) applyFilter("to",dsTO,setDsTO); }, [fTO, dsTO.all, dsTO.sortCol, dsTO.sortAsc]); // eslint-disable-line
  useEffect(()=>{ if (dsMap.all.length) applyFilter("map",dsMap,setDsMap); }, [fMap, dsMap.all, dsMap.sortCol, dsMap.sortAsc]); // eslint-disable-line
  useEffect(()=>{ if (dsTrip.all.length) applyFilter("trip",dsTrip,setDsTrip); }, [fTrip, dsTrip.all, dsTrip.sortCol, dsTrip.sortAsc]); // eslint-disable-line

  useEffect(()=>{
    if (rawTO.length) {
      const rec=rawTO.filter(r=>r["Receive Status"]==="Received").length;
      const abn=rawTO.filter(r=>r["Receive Status"]==="Abnormal").length;
      setStats({
        tot:rawTO.length, rec, abn, map:0,
        rvr:new Set(rawTO.map(r=>r["Receiver Name"])).size,
        snd:new Set(rawTO.map(r=>r["Sender Name"])).size,
      });
    }
  }, [rawTO]);

  useEffect(()=>{
    if (dsMap.all.length) {
      const mapped = dsMap.all.filter(r=>r["_mapped"]).length;
      setStats(s=>({...s, map:mapped}));
    }
  }, [dsMap.all]);

  // ── Sort ───────────────────────────────────────────────────────────
  const doSort = (key:"to"|"map"|"trip", col:string) => {
    const setter = key==="to" ? setDsTO : key==="map" ? setDsMap : setDsTrip;
    setter(d=>{
      const asc = d.sortCol===col ? !d.sortAsc : true;
      return {...d, sortCol:col, sortAsc:asc};
    });
  };

  // ── Export ────────────────────────────────────────────────────────
  const doExport = () => {
    // Chọn đúng ds/vis theo tab đang xem
    let ds: DS, vis: Set<string>, prefix: string;
    if (expType === "mapping") {
      ds = dsMap; vis = mapVis; prefix = "Mapping_";
    } else if (tab === "map") {
      ds = dsMap; vis = mapVis; prefix = "Mapping_";
    } else if (tab === "trip") {
      ds = dsTrip; vis = tripVis; prefix = "Trip_";
    } else {
      ds = dsTO; vis = toVis; prefix = "";
    }

    // Chỉ xuất cột đang hiển thị, bỏ internal cols
    const cols = ds.cols.filter(c => vis.has(c) && c !== "_mapped" && c !== "_source" || (vis.has(c) && c === "_source"));
    const toRow = (r: Row) => {
      const obj: Record<string, unknown> = {};
      ds.cols.filter(c => vis.has(c) && c !== "_mapped").forEach(c => {
        let v = r[c];
        if (v instanceof Date) v = fmtDate(v);
        obj[c] = v;
      });
      return obj;
    };
    const data = ds.view;
    const wb = XLSX.utils.book_new();

    if (expType === "all" || expType === "mapping") {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.map(toRow)), "Data");
    } else if (expType === "by-receiver") {
      const grp: Record<string, Row[]> = {};
      data.forEach(r => { const k = str(r["Receiver Name"]) || "Unknown"; (grp[k] = grp[k] || []).push(r); });
      const ks = Object.keys(grp);
      if (ks.length === 1) {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(grp[ks[0]].map(toRow)), "Data");
      } else {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ks.map(k => ({ "Receiver Name": k, Count: grp[k].length }))), "Overview");
        ks.forEach(k => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(grp[k].map(toRow)), k.slice(0, 31)));
      }
    } else if (expType === "by-status") {
      const grp: Record<string, Row[]> = {};
      data.forEach(r => { const k = str(r["Receive Status"]) || "Unknown"; (grp[k] = grp[k] || []).push(r); });
      Object.entries(grp).forEach(([k, rows]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.map(toRow)), (k || "Unknown").slice(0, 31)));
    } else if (expType === "pivot") {
      const recvrs = [...new Set(data.map(r => str(r["Receiver Name"]) || "?"))].sort();
      const stats2 = [...new Set(data.map(r => str(r["Receive Status"]) || "?"))].sort();
      const piv: Row[] = [];
      recvrs.forEach(rc => stats2.forEach(st => {
        const cnt = data.filter(r => r["Receiver Name"] === rc && r["Receive Status"] === st).length;
        if (cnt) piv.push({ "Receiver Name": rc, "Receive Status": st, Count: cnt });
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(piv), "Pivot");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data.map(toRow)), "Detail");
    }
    void cols;
    XLSX.writeFile(wb, `LH_${prefix}${dateStr()}.xlsx`);
    setExpOpen(false);
    showToast("Đã xuất file!", "t-ok");
  };

  // ── MultiInput: textarea expand khi paste nhiều giá trị ──────────
  const MultiInput = ({
    ph, v, sv, w=130
  }: {ph:string; v:string; sv:(v:string)=>void; w?:number}) => {
    const tokens = v.trim() ? v.split(/[\s,\t\n]+/).filter(Boolean) : [];
    const isMulti = tokens.length > 1;
    return (
      <div style={{position:"relative", display:"inline-flex", flexDirection:"column", gap:2}}>
        <textarea
          placeholder={ph}
          value={v}
          rows={1}
          onChange={e=>sv(e.target.value)}
          style={{
            width:w, minHeight:32, maxHeight:120,
            padding:"6px 10px",
            border:`1px solid ${isMulti?"var(--or)":"var(--bd)"}`,
            borderRadius:"var(--r)",
            background: isMulti?"var(--orb)":"var(--srf)",
            color:"var(--tx)", fontSize:12,
            resize:"vertical", outline:"none",
            fontFamily:"inherit",
            lineHeight:1.4,
            transition:"border-color .12s, background .12s",
          }}
          onFocus={e=>{ if (!isMulti) e.currentTarget.style.borderColor="var(--or)"; }}
          onBlur={e=>{ if (!isMulti) e.currentTarget.style.borderColor="var(--bd)"; }}
        />
        {isMulti && (
          <span style={{
            position:"absolute", bottom:4, right:6,
            fontSize:9, fontWeight:600,
            color:"var(--or)", background:"var(--orb)",
            padding:"0 4px", borderRadius:3, pointerEvents:"none",
          }}>
            {tokens.length}
          </span>
        )}
      </div>
    );
  };

  // ── Upload Box Component ──────────────────────────────────────────
  const UpBox = ({ k, label, sub }: { k:"to"|"ho"|"en"; label:string; sub:string }) => {
    const ref = useRef<HTMLInputElement>(null);
    const st  = uState[k];
    return (
      <div
        className={`upbox${st==="ok"?" ok":st==="err"?" err":""}`}
        onClick={()=>ref.current?.click()}
        onDragOver={e=>{e.preventDefault(); e.currentTarget.classList.add("drag");}}
        onDragLeave={e=>e.currentTarget.classList.remove("drag")}
        onDrop={e=>{
          e.preventDefault(); e.currentTarget.classList.remove("drag");
          handleUpload(e.dataTransfer.files, k);
        }}
      >
        <input ref={ref} type="file" multiple accept=".xlsx,.csv,.zip"
          onChange={e=>{ if (e.target.files?.length) handleUpload(e.target.files, k); }} />
        {st==="ok" && (
          <button className="clr-btn" onClick={e=>{e.stopPropagation();clearDS(k);}}>×</button>
        )}
        <div style={{fontWeight:600, fontSize:12, marginBottom:3, color:st==="ok"?"var(--gn)":st==="err"?"var(--rd)":"var(--tx)"}}>
          {label}
        </div>
        <div style={{fontSize:11, color:"var(--tx3)"}}>{sub}</div>
        {uInfo[k] && (
          <div style={{fontSize:11, marginTop:5, fontWeight:500, color:st==="ok"?"var(--gn)":"var(--rd)"}}>
            {st==="ok"?"✓ ":""}{uInfo[k]}
          </div>
        )}
      </div>
    );
  };

  // ── Table renderer ─────────────────────────────────────────────────
  const renderTable = (
    ds: DS,
    key: "to"|"map"|"trip",
    vis: Set<string>,
    setVis: React.Dispatch<React.SetStateAction<Set<string>>>
  ) => {
    const cols = ds.cols.filter(c=>vis.has(c)&&c!=="_mapped");
    const total = ds.view.length;
    const pages = Math.ceil(total/PAGE_SIZE)||1;
    const page  = ds.page;
    const slice = ds.view.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

    const setDs = key==="to"?setDsTO:key==="map"?setDsMap:setDsTrip;

    return (
      <>
        {/* Pills */}
        <div className="pills-bar">
          <span className="lbl">Cột:</span>
          {ds.cols.filter(c=>c!=="_mapped").map(col=>(
            <span key={col} className={`pill${vis.has(col)?" on":""}`}
              onClick={()=>{
                const next=new Set(vis);
                next.has(col)?next.delete(col):next.add(col);
                setVis(next);
              }}>
              {col}
            </span>
          ))}
        </div>

        {/* Table */}
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                {cols.map(c=>{
                  const cls=ds.sortCol===c?(ds.sortAsc?"asc":"desc"):"";
                  return (
                    <th key={c} className={cls} onClick={()=>doSort(key,c)}>{c}</th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {slice.length===0
                ? <tr className="empty-state"><td colSpan={Math.max(cols.length,1)}>Không có dữ liệu khớp</td></tr>
                : slice.map((row,i)=>{
                    const mapped = row["_mapped"]===true;
                    return (
                      <tr key={i} className={mapped?"mapped":""}>
                        {cols.map(col=>{
                          let raw=row[col];
                          if (raw instanceof Date) raw=fmtDate(raw);
                          const v=raw===null||raw===undefined?"":String(raw);
                          const clean=v.startsWith("'")?v.slice(1):v;

                          if (col==="Receive Status"||col==="TO Status") {
                            const bc=BADGE_CLS[clean]||"bgr";
                            return <td key={col}><span className={`badge ${bc}`}>{esc(clean||"—")}</span></td>;
                          }
                          if (col==="_source"&&clean) {
                            return <td key={col}><span className={`badge ${clean==="handover"?"bo":"bb"}`}>{clean}</span></td>;
                          }
                          if (col==="Vehicle Plate Number"&&clean) {
                            return <td key={col}><span className="badge bp">{esc(clean)}</span></td>;
                          }
                          if (!clean) return <td key={col} className="muted">—</td>;
                          return <td key={col} title={clean}>{esc(clean)}</td>;
                        })}
                      </tr>
                    );
                  })
              }
            </tbody>
          </table>
        </div>

        {/* Pager */}
        {total>0 && (
          <div className="pager">
            <span>{((page-1)*PAGE_SIZE+1)}–{Math.min(page*PAGE_SIZE,total)} / {total.toLocaleString()} dòng</span>
            <div className="pager-btns">
              <button className="pg" disabled={page===1} onClick={()=>setDs(d=>({...d,page:d.page-1}))}>‹</button>
              {pageRange(page,pages).map((p,i)=>
                p==="..."
                  ? <span key={i} style={{padding:"0 3px",color:"var(--tx3)"}}>…</span>
                  : <button key={i} className={`pg${p===page?" on":""}`} onClick={()=>setDs(d=>({...d,page:p as number}))}>{p}</button>
              )}
              <button className="pg" disabled={page===pages} onClick={()=>setDs(d=>({...d,page:d.page+1}))}>›</button>
            </div>
            <div className="pg-jump">
              Trang <input type="number" className="inp" style={{width:40,height:26,textAlign:"center",padding:"0 4px"}}
                defaultValue={page} min={1} max={pages}
                onKeyDown={e=>{ if (e.key==="Enter") { const v=Math.max(1,Math.min(pages,+(e.currentTarget.value))); setDs(d=>({...d,page:v})); }}}
              /> <button className="btn sm" onClick={()=>{}}>Go</button>
            </div>
          </div>
        )}
      </>
    );
  };

  // ── Summary ────────────────────────────────────────────────────────
  const SummaryModal = () => {
    const data = dsTO.view;
    const total = data.length;
    const grp: Record<string,number> = {};
    const byRec: Record<string,number> = {};
    data.forEach(r=>{
      const rec=str(r["Receiver Name"])||"Unknown";
      const st=str(r["Receive Status"])||"Unknown";
      const k=`${rec}|||${st}`;
      grp[k]=(grp[k]||0)+1;
      byRec[rec]=(byRec[rec]||0)+1;
    });

    return (
      <div className="overlay" onClick={e=>{if (e.target===e.currentTarget) setSumOpen(false);}}>
        <div className="modal">
          <div className="modal-head">
            <h2>Summary — Receiver × Receive Status</h2>
            <button className="mclose" onClick={()=>setSumOpen(false)}>×</button>
          </div>
          <div className="modal-body">
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:14}}>
              {Object.entries(byRec).sort((a,b)=>b[1]-a[1]).map(([n,c])=>(
                <div key={n} style={{background:"var(--srf2)",border:"1px solid var(--bd)",borderRadius:"var(--r)",padding:"8px 12px",minWidth:90}}>
                  <div style={{fontSize:18,fontWeight:700}}>{c.toLocaleString()}</div>
                  <div style={{fontSize:11,color:"var(--tx3)",marginTop:1}}>{n}</div>
                </div>
              ))}
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr style={{background:"var(--srf3)"}}>
                  {["Receiver Name","Receive Status","#TO","%",""].map(h=>(
                    <th key={h} style={{padding:"7px 10px",textAlign:h==="#TO"||h==="%"?"right":"left",fontSize:11,fontWeight:600,color:"var(--tx2)",textTransform:"uppercase",letterSpacing:".4px",borderBottom:"1px solid var(--bd)"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Object.entries(grp).sort((a,b)=>b[1]-a[1]).map(([k,cnt])=>{
                  const [rec,st]=k.split("|||");
                  const pct=(cnt/total*100).toFixed(1);
                  return (
                    <tr key={k} style={{borderBottom:"1px solid var(--bd)"}}>
                      <td style={{padding:"7px 10px"}}>{rec}</td>
                      <td style={{padding:"7px 10px"}}><span className={`badge ${BADGE_CLS[st]||"bgr"}`}>{st}</span></td>
                      <td style={{padding:"7px 10px",textAlign:"right",fontWeight:500}}>{cnt.toLocaleString()}</td>
                      <td style={{padding:"7px 10px",textAlign:"right"}}>{pct}%</td>
                      <td style={{padding:"7px 10px"}}>
                        <div className="mini-bar"><div className="mini-fill" style={{width:`${pct}%`}}/></div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{background:"var(--orb)",borderTop:"1.5px solid var(--orbd)"}}>
                  <td colSpan={2} style={{padding:"7px 10px",fontWeight:600}}>Total</td>
                  <td style={{padding:"7px 10px",textAlign:"right",fontWeight:600}}>{total.toLocaleString()}</td>
                  <td style={{padding:"7px 10px",textAlign:"right",fontWeight:600}}>100%</td>
                  <td/>
                </tr>
              </tfoot>
            </table>
          </div>
          <div className="modal-foot">
            <button className="btn b-gn" onClick={()=>{
              const wb=XLSX.utils.book_new();
              const rows=Object.entries(grp).sort((a,b)=>b[1]-a[1]).map(([k,c])=>{
                const [rec,st]=k.split("|||");
                return {"Receiver Name":rec,"Receive Status":st,Count:c,"%":+(c/total*100).toFixed(1)};
              });
              XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Summary");
              XLSX.writeFile(wb,`LH_Summary_${dateStr()}.xlsx`);
              showToast("Đã xuất Summary!","t-ok");
            }}>Export Summary</button>
            <button className="btn b-or" onClick={()=>setSumOpen(false)}>Đóng</button>
          </div>
        </div>
      </div>
    );
  };

  const hasTO   = rawTO.length>0;
  const hasTrip = rawHO.length>0||rawEN.length>0;
  const hasMap  = hasTO&&hasTrip;

  const tabCount = {
    to:   dsTO.view.length,
    map:  dsMap.view.length,
    trip: dsTrip.view.length,
  };

  // Avatar initials
  const name  = session?.user?.name ?? "";
  const email = session?.user?.email ?? "";
  const initials = name
    ? (name.split(" ").shift()?.[0]??""+(name.split(" ").pop()?.[0]??"")).toUpperCase().slice(0,2)
    : email[0]?.toUpperCase() ?? "?";

  return (
    <div style={{minHeight:"100vh", background:"var(--bg)"}}>
      {/* ── Topbar ── */}
      <div className="topbar">
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{width:28,height:28,borderRadius:7,background:"var(--or)",display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 1px 4px rgba(217,82,4,.3)"}}>
            <svg width="15" height="15" fill="none" viewBox="0 0 24 24">
              <rect x="2" y="3" width="20" height="14" rx="2" stroke="white" strokeWidth="1.8"/>
              <path d="M2 8h20" stroke="white" strokeWidth="1.5" strokeOpacity=".6"/>
            </svg>
          </div>
          <span style={{fontWeight:700,fontSize:13,letterSpacing:"-.2px"}}>Linehaul Data Builder</span>
          <span style={{fontSize:10,fontWeight:600,background:"var(--orb)",color:"var(--or)",border:"1px solid var(--orbd)",borderRadius:99,padding:"1px 7px",lineHeight:"16px"}}>Live</span>
        </div>
        <div style={{flex:1}}/>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          <button className="btn sm" disabled={!hasTO} onClick={()=>setSumOpen(true)}>Summary</button>
          <button className="btn sm b-gn" disabled={!hasTO} onClick={()=>setExpOpen(true)}>Export</button>
          <div style={{width:1,height:18,background:"var(--bd)",margin:"0 4px"}}/>
          {/* User avatar */}
          <div style={{display:"flex",alignItems:"center",gap:7}}>
            {session?.user?.image
              ? <img src={session.user.image} alt="" style={{width:26,height:26,borderRadius:"50%",border:"1px solid var(--bd)"}}/>
              : <div style={{width:26,height:26,borderRadius:"50%",background:"#CECBF6",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:600,color:"#3C3489"}}>{initials}</div>
            }
            <span style={{fontSize:12,color:"var(--tx2)",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{email}</span>
          </div>
          <button className="btn sm b-ghost" onClick={()=>signOut({callbackUrl:"/login"})}>Đăng xuất</button>
        </div>
      </div>

      {/* ── Tabs ── */}
      <div className="tabs">
        {(["to","map","trip","manual"] as const).map(t=>(
          <div key={t} className={`tab${tab===t?" on":""}`} onClick={()=>setTab(t)}>
            {t==="to"?"TO Orders":t==="map"?"LH Mapping":t==="trip"?"Trip Detail":"Hướng dẫn"}
            {t!=="manual" && <span className="tcnt">{tabCount[t as "to"|"map"|"trip"].toLocaleString()}</span>}
          </div>
        ))}
      </div>

      <div className="main">

        {/* ══ TO Orders ══ */}
        {tab==="to" && (
          <div className="panel on">
            {/* Import */}
            <div className="card">
              <div className="card-head"><span className="card-title">Import file</span></div>
              <div style={{padding:14}}>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
                  <UpBox k="to" label="TO Order" sub="xlsx / csv / zip"/>
                  <UpBox k="ho" label="LH Trip — Handover" sub="xlsx / csv / zip"/>
                  <UpBox k="en" label="LH Trip — Ended" sub="xlsx / csv / zip"/>
                </div>
                {prog && (
                  <div className="prog-row" style={{display:"block",marginTop:10}}>
                    <div className="prog-track"><div className="prog-fill" style={{width:`${prog.pct}%`}}/></div>
                    <div style={{fontSize:11,color:"var(--tx3)",marginTop:3}}>{prog.lbl}</div>
                  </div>
                )}
              </div>
            </div>

            {/* Stats */}
            {hasTO && (
              <div className="card">
                <div className="stats">
                  {[
                    {v:stats.tot,l:"Total TO",    cls:"c-or"},
                    {v:stats.rec,l:"Received",    cls:"c-gn"},
                    {v:stats.abn,l:"Abnormal",    cls:"c-rd"},
                    {v:stats.rvr,l:"Receivers",   cls:""},
                    {v:stats.snd,l:"Senders",     cls:""},
                    {v:stats.map,l:"Mapped trip", cls:"c-pu"},
                  ].map((s,i)=>(
                    <div key={i} className={`stat ${s.cls}`}>
                      <div className="v">{s.v.toLocaleString()}</div>
                      <div className="l">{s.l}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Data table */}
            {hasTO && (
              <div className="card">
                <div className="filters">
                  <MultiInput ph="TO Number"    v={fTO.to}  sv={v=>setFTO(f=>({...f,to:v}))}/>
                  <MultiInput ph="SPX Tracking" v={fTO.spx} sv={v=>setFTO(f=>({...f,spx:v}))}/>
                  <MultiInput ph="Receiver"     v={fTO.rec} sv={v=>setFTO(f=>({...f,rec:v}))}/>
                  <MultiInput ph="Sender"       v={fTO.snd} sv={v=>setFTO(f=>({...f,snd:v}))}/>
                  <select className="inp" style={{width:140}} value={fTO.tos} onChange={e=>setFTO(f=>({...f,tos:e.target.value}))}>
                    <option value="">TO Status</option>
                    {toStatuses.map(s=><option key={s}>{s}</option>)}
                  </select>
                  <select className="inp" style={{width:150}} value={fTO.rs} onChange={e=>setFTO(f=>({...f,rs:e.target.value}))}>
                    <option value="">Receive Status</option>
                    {rcvStatuses.map(s=><option key={s}>{s}</option>)}
                  </select>
                  <input type="date" className="inp" style={{width:140}} value={fTO.dt} onChange={e=>setFTO(f=>({...f,dt:e.target.value}))}/>
                  <span className="sp"/>
                  <button className="btn b-ghost sm" onClick={()=>setFTO({to:"",spx:"",rec:"",snd:"",tos:"",rs:"",dt:""})}>Xóa filter</button>
                </div>
                {renderTable(dsTO,"to",toVis,setToVis)}
              </div>
            )}
          </div>
        )}

        {/* ══ LH Mapping ══ */}
        {tab==="map" && (
          <div className="panel on">
            {!hasMap && (
              <div style={{background:"var(--orb)",border:"1px solid var(--orbd)",borderRadius:8,padding:"10px 14px",fontSize:12,color:"var(--or)"}}>
                Upload TO data + ít nhất 1 file LH Trip (Handover hoặc Ended) để xem mapping.
              </div>
            )}
            {hasMap && (
              <div className="card">
                <div className="card-head">
                  <span className="card-title">TO ↔ LH Trip Mapping</span>
                  <span style={{fontSize:12,color:"var(--tx3)"}}>{mapRatio}</span>
                </div>
                <div className="filters">
                  <MultiInput ph="TO Number"      v={fMap.to}  sv={v=>setFMap(f=>({...f,to:v}))}/>
                  <MultiInput ph="LH Trip Number" v={fMap.trp} sv={v=>setFMap(f=>({...f,trp:v}))}/>
                  <MultiInput ph="Plate Number"   v={fMap.plt} sv={v=>setFMap(f=>({...f,plt:v}))}/>
                  <MultiInput ph="Driver"         v={fMap.drv} sv={v=>setFMap(f=>({...f,drv:v}))}/>
                  <MultiInput ph="Receiver"       v={fMap.rcv} sv={v=>setFMap(f=>({...f,rcv:v}))}/>
                  <select className="inp" style={{width:130}} value={fMap.src} onChange={e=>setFMap(f=>({...f,src:e.target.value}))}>
                    <option value="">Source (all)</option>
                    <option value="handover">Handover</option>
                    <option value="ended">Ended</option>
                  </select>
                  <select className="inp" style={{width:130}} value={fMap.ms} onChange={e=>setFMap(f=>({...f,ms:e.target.value}))}>
                    <option value="">Mapped (all)</option>
                    <option value="1">Mapped</option>
                    <option value="0">Not mapped</option>
                  </select>
                  <span className="sp"/>
                  <button className="btn b-ghost sm" onClick={()=>setFMap({to:"",trp:"",plt:"",drv:"",rcv:"",src:"",ms:""})}>Xóa filter</button>
                </div>
                {renderTable(dsMap,"map",mapVis,setMapVis)}
              </div>
            )}
          </div>
        )}

        {/* ══ Trip Detail ══ */}
        {tab==="trip" && (
          <div className="panel on">
            {!hasTrip && (
              <div style={{background:"var(--blb)",border:"1px solid var(--blbd)",borderRadius:8,padding:"10px 14px",fontSize:12,color:"var(--bl)"}}>
                Upload file LH Trip (Handover hoặc Ended) để xem trip detail.
              </div>
            )}
            {hasTrip && (
              <div className="card">
                <div className="card-head"><span className="card-title">LH Trip Detail</span></div>
                <div className="filters">
                  <MultiInput ph="LH Trip Number" v={fTrip.trp} sv={v=>setFTrip(f=>({...f,trp:v}))} w={140}/>
                  <MultiInput ph="Plate Number"   v={fTrip.plt} sv={v=>setFTrip(f=>({...f,plt:v}))} w={140}/>
                  <MultiInput ph="Driver"         v={fTrip.drv} sv={v=>setFTrip(f=>({...f,drv:v}))} w={140}/>
                  <MultiInput ph="Station"        v={fTrip.stn} sv={v=>setFTrip(f=>({...f,stn:v}))} w={140}/>
                  <select className="inp" style={{width:140}} value={fTrip.src} onChange={e=>setFTrip(f=>({...f,src:e.target.value}))}>
                    <option value="">Source (all)</option>
                    <option value="handover">Handover</option>
                    <option value="ended">Ended</option>
                  </select>
                  <span className="sp"/>
                  <button className="btn b-ghost sm" onClick={()=>setFTrip({trp:"",plt:"",drv:"",stn:"",src:""})}>Xóa filter</button>
                </div>
                {renderTable(dsTrip,"trip",tripVis,setTripVis)}
              </div>
            )}
          </div>
        )}

        {/* ══ Manual ══ */}
        {tab==="manual" && (
          <div className="panel on">
            <div className="card" style={{padding:"24px 28px",maxWidth:720}}>
              <div style={{marginBottom:20}}>
                <div style={{fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:".8px",color:"var(--or)",marginBottom:4}}>Hướng dẫn</div>
                <h2 style={{fontSize:18,fontWeight:700,letterSpacing:"-.3px",marginBottom:8}}>Linehaul Data Builder</h2>
                <p style={{fontSize:13,color:"var(--tx2)",lineHeight:1.7}}>Công cụ phân tích TO Order và ghép thông tin chuyến xe Linehaul.</p>
              </div>
              {[
                {n:"1",t:"Import file",d:"Upload 3 loại file vào ô tương ứng: TO Order, LH Trip Handover, LH Trip Ended. Hỗ trợ xlsx, csv, và zip (nhiều file gộp lại). Ô chuyển xanh = đọc thành công."},
                {n:"2",t:"Xem & filter",d:"Tab TO Orders hiển thị toàn bộ dữ liệu. Filter real-time theo TO Number, SPX Tracking, Receiver, Sender, Status, Date. Toggle cột bằng pill dưới filter. Nhấn header để sort."},
                {n:"3",t:"Mapping",d:'Tab LH Mapping ghép TO với thông tin chuyến xe (biển số, tài xế, giờ thực tế) qua "Line Hual Trip Number". Hàng tím = đã ghép được. Filter "Not mapped" để audit nhanh.'},
                {n:"4",t:"Export",d:"Nhấn Export → chọn chế độ: All in one sheet, Split by Receiver, Split by Status, Pivot + Detail, hoặc Mapping (TO + Trip info)."},
              ].map(s=>(
                <div key={s.n} style={{display:"flex",gap:14,padding:"14px 0",borderBottom:"1px solid var(--bd)"}}>
                  <div style={{width:28,height:28,background:"var(--or)",color:"#fff",borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,flexShrink:0}}>
                    {s.n}
                  </div>
                  <div>
                    <div style={{fontWeight:600,fontSize:13,marginBottom:4}}>{s.t}</div>
                    <div style={{fontSize:12,color:"var(--tx2)",lineHeight:1.65}}>{s.d}</div>
                  </div>
                </div>
              ))}
              <div style={{marginTop:16,padding:"12px 14px",background:"var(--srf2)",borderLeft:"3px solid var(--or)",borderRadius:"0 6px 6px 0",fontSize:12,color:"var(--tx2)"}}>
                <strong style={{color:"var(--or)"}}>Mẹo: </strong>
                Dữ liệu xử lý hoàn toàn trên trình duyệt — không gửi ra server, không lo lộ dữ liệu.
              </div>
            </div>
          </div>
        )}
      </div>

      {/* ── Summary Modal ── */}
      {sumOpen && <SummaryModal/>}

      {/* ── Export Modal ── */}
      {expOpen && (
        <div className="overlay" onClick={e=>{if (e.target===e.currentTarget) setExpOpen(false);}}>
          <div className="modal" style={{maxWidth:460}}>
            <div className="modal-head">
              <h2>Export Data</h2>
              <button className="mclose" onClick={()=>setExpOpen(false)}>×</button>
            </div>
            <div className="modal-body">
              {/* Info: đang ở tab nào, bao nhiêu dòng */}
              {(()=>{
                const curDs = tab==="map" ? dsMap : tab==="trip" ? dsTrip : dsTO;
                const tabLabel = tab==="map" ? "LH Mapping" : tab==="trip" ? "Trip Detail" : "TO Orders";
                return (
                  <div style={{background:"var(--srf2)",border:"1px solid var(--bd)",borderRadius:6,padding:"8px 12px",marginBottom:12,fontSize:12}}>
                    <span style={{color:"var(--tx3)"}}>Tab hiện tại: </span>
                    <strong>{tabLabel}</strong>
                    <span style={{color:"var(--tx3)"}}> · </span>
                    <strong>{curDs.view.length.toLocaleString()}</strong>
                    <span style={{color:"var(--tx3)"}}> dòng sau filter</span>
                  </div>
                );
              })()}
              {/* Options theo tab */}
              {(()=>{
                // All tabs: always show "all" + "mapping"
                // TO tab: thêm by-receiver, by-status, pivot
                const opts: {v:string;t:string;d:string}[] = [
                  {v:"all", t:"All data in one sheet", d:`Xuất toàn bộ data tab ${tab==="map"?"Mapping":tab==="trip"?"Trip":"TO"} đang filter, 1 sheet`},
                ];
                if (tab==="to") {
                  opts.push(
                    {v:"by-receiver", t:"Split by Receiver",       d:"Mỗi receiver 1 sheet"},
                    {v:"by-status",   t:"Split by Receive Status", d:"Received / Abnormal — riêng từng sheet"},
                    {v:"pivot",       t:"Pivot + Detail",          d:"Sheet pivot (Receiver × Status) + sheet chi tiết"},
                  );
                }
                opts.push({v:"mapping", t:"Mapping (TO + Trip info)", d:"Xuất tab LH Mapping — TO ghép với biển số, tài xế, giờ thực tế"});
                return opts.map(o=>(
                  <label key={o.v} className="eopt">
                    <input type="radio" name="et" value={o.v} checked={expType===o.v} onChange={()=>setExpType(o.v)} style={{accentColor:"var(--or)"}}/>
                    <div>
                      <div style={{fontWeight:500,fontSize:13}}>{o.t}</div>
                      <div style={{fontSize:11,color:"var(--tx3)"}}>{o.d}</div>
                    </div>
                  </label>
                ));
              })()}
            </div>
            <div className="modal-foot">
              <button className="btn b-ghost" onClick={()=>setExpOpen(false)}>Hủy</button>
              <button className="btn b-gn" onClick={doExport}>Tải xuống</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Toast ── */}
      <div className={`toast${toast?" show":""} ${toast?.cls??""}`}>{toast?.msg}</div>
    </div>
  );
}