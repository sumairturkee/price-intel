'use client';
import { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../lib/supabase';
import { findMyMatch } from '../lib/matching';

const DEFAULT_PRICE_LIST = [
  {model:"Poco C81X (Fresh)",variant:"3/64",price:10100},{model:"Poco C81 (Fresh)",variant:"4/64",price:11100},
  {model:"Poco C85X 5G (Activated)",variant:"4/64",price:12400},{model:"Poco C85X 5G (Activated)",variant:"4/128",price:13400},
  {model:"Poco C85X 5G (Fresh)",variant:"4/64",price:12600},{model:"Poco C85X 5G (Fresh)",variant:"4/128",price:13600},
  {model:"Poco C85 5G (Activated)",variant:"4/128",price:14200},{model:"Poco C85 5G (Activated)",variant:"6/128",price:15100},{model:"Poco C85 5G (Activated)",variant:"8/128",price:16600},
  {model:"Poco C85 5G (Fresh)",variant:"4/128",price:14700},{model:"Poco C85 5G (Fresh)",variant:"6/128",price:15400},{model:"Poco C85 5G (Fresh)",variant:"8/128",price:16900},
  {model:"Poco M7 Plus 5G (Activated)",variant:"4/128",price:14000},{model:"Poco M7 Plus 5G (Activated)",variant:"6/128",price:15300},{model:"Poco M7 Plus 5G (Activated)",variant:"8/128",price:16300},
  {model:"Poco M7 Plus 5G (Fresh)",variant:"4/128",price:14500},{model:"Poco M7 Plus 5G (Fresh)",variant:"6/128",price:15800},{model:"Poco M7 Plus 5G (Fresh)",variant:"8/128",price:16800},
  {model:"Poco M8 5G (Fresh)",variant:"6/128",price:19100},{model:"Poco M8 5G (Fresh)",variant:"8/128",price:21100},
  {model:"RealMe P4 Lite 5G",variant:"4/64",price:14200},{model:"RealMe P4 Lite 5G",variant:"4/128",price:15700},{model:"RealMe P4 Lite 5G",variant:"6/128",price:17700},
  {model:"RealMe P4x 5G (New MRP)",variant:"6/128",price:18700},{model:"RealMe P4x 5G (New MRP)",variant:"8/128",price:19900},
  {model:"RealMe P4 AI 5G",variant:"6/128",price:20100},{model:"RealMe P4 Power 5G",variant:"8/128",price:25900},
  {model:"Narzo 100 Lite 5G",variant:"4/64",price:14200},{model:"Narzo 100 Lite 5G",variant:"4/128",price:15700},{model:"Narzo 100 Lite 5G",variant:"6/128",price:17700},
  {model:"Narzo 90x 5G",variant:"6/128",price:16900},{model:"Narzo 90x 5G",variant:"8/128",price:17900},
  {model:"Moto Edge 70 Fusion 5G",variant:"8/128",price:26700},{model:"Moto Edge 70 Fusion 5G",variant:"8/256",price:29700},{model:"Moto Edge 70 Fusion 5G",variant:"12/256",price:32700},
  {model:"Moto Edge 60 Pro 5G",variant:"8/256",price:27700},{model:"Moto Edge 60 Pro 5G",variant:"12/256",price:31700},
  {model:"Samsung M06 5G",variant:"4/128",price:12100},{model:"Samsung M06 5G",variant:"6/128",price:14100},
  {model:"Samsung M17 E 5G",variant:"4/128",price:12700},{model:"Samsung M17 E 5G",variant:"6/128",price:14500},
  {model:"Samsung M17 5G",variant:"4/128",price:14300},{model:"Samsung M17 5G",variant:"6/128",price:16300},
  {model:"Samsung M36 5G",variant:"6/128",price:17100},{model:"Samsung M36 5G",variant:"8/128",price:18800},
  {model:"Samsung F70 E 5G",variant:"4/128",price:12700},{model:"Samsung F70 E 5G",variant:"6/128",price:14500},
  {model:"Samsung F36 5G",variant:"6/128",price:16100},{model:"Samsung F36 5G",variant:"8/128",price:18500},
  {model:"Oppo K14x 5G",variant:"4/64",price:14400},{model:"Oppo K14x 5G",variant:"4/128",price:15500},{model:"Oppo K14x 5G",variant:"6/128",price:17500},
  {model:"Redmi A7 Pro 5G",variant:"4/64",price:13000},{model:"Redmi A7 Pro 5G",variant:"4/128",price:14000},
  {model:"Lava Bold N1 5G",variant:"6/128",price:12300},
  {model:"IQoo Z11X 5G",variant:"6/128",price:21500},{model:"IQoo Z11X 5G",variant:"8/128",price:23000},
];

const C = {
  bg:'#0D1117',surface:'#161B22',surface2:'#1C2128',border:'#21262D',
  accent:'#00D4AA',accentDim:'#00D4AA18',amber:'#F0A500',amberDim:'#F0A50018',
  red:'#FF4C6A',redDim:'#FF4C6A18',blue:'#58A6FF',blueDim:'#58A6FF18',
  muted:'#8B949E',text:'#E6EDF3',textDim:'#C9D1D9',
};

const fmt = n => n>0 ? `₹${Number(n).toLocaleString('en-IN')}` : '—';

function parseExcelFile(file) {
  return new Promise((res,rej) => {
    const r = new FileReader();
    r.onload = e => {
      try {
        const wb = XLSX.read(e.target.result,{type:'array'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws,{defval:''});
        const items = [];
        for (const row of rows) {
          const keys = Object.keys(row).map(k=>k.trim());
          const mk = keys.find(k=>/model/i.test(k));
          const vk = keys.find(k=>/variant|ram|storage/i.test(k));
          const pk = keys.find(k=>/price|mrp|rate/i.test(k));
          if (!mk||!pk) continue;
          const model = String(row[mk]).trim();
          const variant = vk ? String(row[vk]).trim() : '';
          const raw = row[pk];
          const price = typeof raw==='number' ? raw : parseInt(String(raw).replace(/[^0-9]/g,''),10)||0;
          if (model && price>0) items.push({model,variant,price});
        }
        res(items);
      } catch(err){rej(err);}
    };
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}

function downloadExcel(rows, compNames, dateStr) {
  const headers = ['Model','Variant','My Price',...compNames,'Cheapest Comp','Diff %','Status'];
  const data = rows.map(row => {
    const prices = compNames.map(cn=>row.compPrices[cn]).filter(p=>p>0);
    const cheap = prices.length ? Math.min(...prices) : '';
    let diff='',status='';
    if (row.myPrice&&cheap){const pct=(row.myPrice-cheap)/cheap*100;diff=parseFloat(pct.toFixed(1));status=pct>3?'EXPENSIVE':pct<-3?'CHEAPER':'MATCHED';}
    else if(!row.myPrice){status='NEW MODEL';}
    return [row.model,row.variant,row.myPrice||0,...compNames.map(cn=>row.compPrices[cn]||''),cheap,diff,status];
  });
  const ws = XLSX.utils.aoa_to_sheet([headers,...data]);
  ws['!cols']=[{wch:32},{wch:8},{wch:12},...compNames.map(()=>({wch:12})),{wch:14},{wch:8},{wch:12}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Price Comparison');
  XLSX.writeFile(wb,`Price_Comparison_${dateStr}.xlsx`);
}

function Pill({children,color}){return <span style={{fontSize:10,fontWeight:700,letterSpacing:'0.07em',padding:'2px 7px',borderRadius:3,background:color+'22',color,textTransform:'uppercase'}}>{children}</span>;}

function MultiSelect({label,options,selected,onChange,color=C.accent}){
  const [open,setOpen]=useState(false);
  const ref=useRef(null);
  useEffect(()=>{
    function h(e){if(ref.current&&!ref.current.contains(e.target))setOpen(false);}
    document.addEventListener('mousedown',h);
    return ()=>document.removeEventListener('mousedown',h);
  },[]);
  const isAll=selected.size===0;
  const toggle=(opt)=>{
    const next=new Set(selected);
    if(isAll){options.forEach(o=>{if(o!==opt)next.add(o);});}
    else if(next.has(opt)){next.delete(opt);if(next.size===0)return onChange(new Set());}
    else{next.add(opt);if(next.size===options.length)return onChange(new Set());}
    onChange(next);
  };
  const btnLabel=isAll?label:selected.size===1?[...selected][0]:`${selected.size} selected`;
  return(
    <div ref={ref} style={{position:'relative'}}>
      <button onClick={()=>setOpen(o=>!o)} style={{display:'flex',alignItems:'center',gap:6,background:!isAll?color+'18':C.surface,color:!isAll?color:C.textDim,border:`1px solid ${!isAll?color+'55':C.border}`,borderRadius:6,padding:'7px 11px',fontSize:13,cursor:'pointer',whiteSpace:'nowrap',fontWeight:!isAll?700:400}}>
        <span>{btnLabel}</span><span style={{fontSize:9,color:C.muted}}>▼</span>
        {!isAll&&<span onClick={e=>{e.stopPropagation();onChange(new Set());}} style={{marginLeft:2,color:C.muted,fontSize:12}}>✕</span>}
      </button>
      {open&&(
        <div style={{position:'absolute',top:'calc(100% + 4px)',left:0,zIndex:100,background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,boxShadow:'0 8px 24px #0008',minWidth:180,maxHeight:320,overflowY:'auto',padding:'6px 0'}}>
          <div style={{display:'flex',borderBottom:`1px solid ${C.border}`,padding:'4px 8px 8px',marginBottom:2}}>
            <button onClick={()=>onChange(new Set())} style={{flex:1,background:'transparent',border:'none',color,fontSize:11,fontWeight:700,cursor:'pointer',textAlign:'left',padding:'2px 4px'}}>All</button>
            <button onClick={()=>onChange(new Set(options))} style={{flex:1,background:'transparent',border:'none',color:C.muted,fontSize:11,cursor:'pointer',textAlign:'right',padding:'2px 4px'}}>Select all</button>
          </div>
          {options.map(opt=>{
            const checked=isAll||selected.has(opt);
            return(
              <div key={opt} onClick={()=>toggle(opt)} style={{display:'flex',alignItems:'center',gap:9,padding:'7px 14px',cursor:'pointer',fontSize:13,color:(!isAll&&selected.has(opt))?color:C.textDim,background:'transparent'}}
                onMouseEnter={e=>e.currentTarget.style.background=C.border} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                <div style={{width:15,height:15,borderRadius:3,flexShrink:0,border:`2px solid ${checked?color:C.border}`,background:checked?color:'transparent',display:'flex',alignItems:'center',justifyContent:'center'}}>
                  {checked&&<span style={{color:'#000',fontSize:9,fontWeight:900,lineHeight:1}}>✓</span>}
                </div>
                <span style={{fontFamily:'monospace'}}>{opt}</span>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function PastePanel({name,existingItems,onConfirm,myModels=[]}){
  const [step,setStep]=useState('paste');
  const [raw,setRaw]=useState('');
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState('');
  const [draft,setDraft]=useState([]);
  const [ow,setOw]=useState(true);
  async function parse(){
    if(!raw.trim())return;
    setLoading(true);setError('');
    try{
      const res=await fetch('/api/parse',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({rawText:raw,myModels})});
      const data=await res.json();
      if(data.error)throw new Error(data.error);
      if(!data.items?.length)throw new Error('No items found');
      setDraft(data.items.map(it=>({...it,_sel:true})));
      setStep('verify');
    }catch(e){setError(e.message||'Parse failed');}
    setLoading(false);
  }
  const sel=draft.filter(r=>r._sel).length;
  const overlaps=draft.filter(r=>r._sel&&existingItems.some(e=>e.model?.toLowerCase()===r.model?.toLowerCase()&&e.variant===r.variant)).length;
  if(step==='paste')return(
    <div style={{padding:24,maxWidth:700}}>
      <p style={{color:C.muted,fontSize:13,marginBottom:12}}>Paste <strong style={{color:C.text}}>{name}</strong>'s WhatsApp price message. Claude will extract all models, variants & prices for you to verify.</p>
      <textarea value={raw} onChange={e=>setRaw(e.target.value)} placeholder={`Paste ${name}'s WhatsApp message here…`}
        style={{width:'100%',minHeight:180,background:C.bg,color:C.text,border:`1px solid ${C.border}`,borderRadius:8,padding:14,fontSize:13,fontFamily:'monospace',resize:'vertical',outline:'none',boxSizing:'border-box'}}/>
      {error&&<p style={{color:C.red,fontSize:12,marginTop:6}}>{error}</p>}
      <div style={{display:'flex',gap:10,marginTop:12,alignItems:'center'}}>
        <button onClick={parse} disabled={loading||!raw.trim()} style={{padding:'10px 24px',borderRadius:6,border:'none',background:loading?C.border:C.accent,color:'#000',fontWeight:700,fontSize:13,cursor:loading?'default':'pointer',opacity:!raw.trim()?0.4:1}}>
          {loading?'Parsing…':'Parse & Preview →'}
        </button>
        {existingItems.length>0&&<span style={{color:C.muted,fontSize:12}}>{existingItems.length} items loaded</span>}
      </div>
    </div>
  );
  return(
    <div style={{padding:24,maxWidth:820}}>
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:16,flexWrap:'wrap',gap:8}}>
        <div style={{display:'flex',alignItems:'center',gap:8}}><h3 style={{margin:0,fontWeight:700}}>Verify Parsed Data</h3><Pill color={C.blue}>STEP 2 OF 2</Pill></div>
        <button onClick={()=>setStep('paste')} style={{background:'transparent',border:`1px solid ${C.border}`,color:C.muted,borderRadius:6,padding:'6px 14px',cursor:'pointer',fontSize:12}}>← Back</button>
      </div>
      {existingItems.length>0&&(
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:'14px 16px',marginBottom:14}}>
          {[
            [true, C.accent, 'Replace list', 'Clears '+existingItems.length+' items, saves '+sel+' selected'],
            [false, C.amber, 'Add/Update', overlaps ? 'Merges · '+overlaps+' overlap'+(overlaps>1?'s':'')+' updated' : 'Merges into existing list'],
          ].map(([val,color,title,desc])=>(
            <label key={String(val)} onClick={()=>setOw(val)} style={{display:'flex',alignItems:'flex-start',gap:10,cursor:'pointer',marginBottom:val?10:0}}>
              <div style={{width:17,height:17,borderRadius:'50%',border:`2px solid ${ow===val?color:C.border}`,background:ow===val?color:'transparent',flexShrink:0,marginTop:1,display:'flex',alignItems:'center',justifyContent:'center'}}>
                {ow===val&&<div style={{width:7,height:7,borderRadius:'50%',background:'#000'}}/>}
              </div>
              <div><div style={{fontWeight:600,fontSize:13,color:ow===val?color:C.textDim}}>{title}</div><div style={{fontSize:11,color:C.muted,marginTop:1}}>{desc}</div></div>
            </label>
          ))}
        </div>
      )}
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:6,fontSize:12}}>
        <div style={{display:'flex',alignItems:'center',gap:8}}>
          <div onClick={()=>setDraft(d=>d.map(r=>({...r,_sel:!d.every(r=>r._sel)})))} style={{width:15,height:15,borderRadius:3,border:`2px solid ${draft.every(r=>r._sel)?C.accent:C.border}`,background:draft.every(r=>r._sel)?C.accent:'transparent',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center'}}>
            {draft.every(r=>r._sel)&&<span style={{color:'#000',fontSize:9,fontWeight:900}}>✓</span>}
          </div>
          <span style={{color:C.textDim}}><strong style={{color:C.accent}}>{sel}</strong> of {draft.length} selected</span>
        </div>
        <div style={{display:'flex',gap:6}}>
          <button onClick={()=>setDraft(d=>d.map(r=>({...r,_sel:true})))} style={{background:'transparent',border:`1px solid ${C.border}`,color:C.muted,borderRadius:5,padding:'3px 9px',cursor:'pointer',fontSize:11}}>All</button>
          <button onClick={()=>setDraft(d=>d.map(r=>({...r,_sel:false})))} style={{background:'transparent',border:`1px solid ${C.border}`,color:C.muted,borderRadius:5,padding:'3px 9px',cursor:'pointer',fontSize:11}}>None</button>
        </div>
      </div>
      <div style={{border:`1px solid ${C.border}`,borderRadius:8,overflow:'hidden',marginBottom:14}}>
        <div style={{display:'grid',gridTemplateColumns:'32px 1fr 90px 110px 32px',background:C.surface2,borderBottom:`1px solid ${C.border}`,padding:'7px 12px',gap:8}}>
          {['','MODEL','VARIANT','PRICE',''].map((h,i)=><div key={i} style={{fontSize:10,fontWeight:700,color:C.muted,letterSpacing:'0.07em',textAlign:i>=3?'right':'left'}}>{h}</div>)}
        </div>
        <div style={{maxHeight:360,overflowY:'auto'}}>
          {draft.map((row,i)=>(
            <div key={i} style={{display:'grid',gridTemplateColumns:'32px 1fr 90px 110px 32px',padding:'6px 12px',gap:8,alignItems:'center',background:row._sel?(i%2===0?'transparent':C.surface+'66'):C.surface+'33',borderBottom:`1px solid ${C.border}`,opacity:row._sel?1:0.45}}>
              <div onClick={()=>setDraft(d=>d.map((r,idx)=>idx===i?{...r,_sel:!r._sel}:r))} style={{width:15,height:15,borderRadius:3,border:`2px solid ${row._sel?C.accent:C.border}`,background:row._sel?C.accent:'transparent',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center'}}>
                {row._sel&&<span style={{color:'#000',fontSize:9,fontWeight:900}}>✓</span>}
              </div>
              <input value={row.model} onChange={e=>setDraft(d=>d.map((r,idx)=>idx===i?{...r,model:e.target.value}:r))}
                style={{background:'transparent',color:C.text,border:'1px solid transparent',borderRadius:4,padding:'3px 6px',fontSize:13,outline:'none',width:'100%',boxSizing:'border-box'}}
                onFocus={e=>e.target.style.borderColor=C.accent} onBlur={e=>e.target.style.borderColor='transparent'}/>
              <input value={row.variant} onChange={e=>setDraft(d=>d.map((r,idx)=>idx===i?{...r,variant:e.target.value}:r))} placeholder="—"
                style={{background:'transparent',color:C.muted,border:'1px solid transparent',borderRadius:4,padding:'3px 6px',fontSize:13,outline:'none',fontFamily:'monospace',width:'100%',boxSizing:'border-box'}}
                onFocus={e=>e.target.style.borderColor=C.accent} onBlur={e=>e.target.style.borderColor='transparent'}/>
              <input type="number" value={row.price} onChange={e=>setDraft(d=>d.map((r,idx)=>idx===i?{...r,price:Number(e.target.value)||0}:r))}
                style={{background:'transparent',color:C.accent,border:'1px solid transparent',borderRadius:4,padding:'3px 6px',fontSize:13,outline:'none',fontFamily:'monospace',textAlign:'right',width:'100%',boxSizing:'border-box'}}
                onFocus={e=>e.target.style.borderColor=C.accent} onBlur={e=>e.target.style.borderColor='transparent'}/>
              <div onClick={()=>setDraft(d=>d.filter((_,idx)=>idx!==i))} style={{color:C.muted,cursor:'pointer',fontSize:13,textAlign:'center'}}
                onMouseEnter={e=>e.currentTarget.style.color=C.red} onMouseLeave={e=>e.currentTarget.style.color=C.muted}>✕</div>
            </div>
          ))}
        </div>
      </div>
      <button onClick={()=>{
        const items=draft.filter(r=>r._sel).map(({model,variant,price})=>({model,variant,price}));
        if(items.length)onConfirm(items,ow);
        setStep('paste');setRaw('');setDraft([]);
      }} disabled={sel===0} style={{padding:'10px 28px',borderRadius:6,border:'none',background:sel>0?C.accent:C.border,color:'#000',fontWeight:700,fontSize:13,cursor:sel>0?'pointer':'default',opacity:sel>0?1:0.5}}>
        {ow||existingItems.length===0?`Save ${sel} Items`:`Add ${sel} Items`}
      </button>
    </div>
  );
}

export default function PriceIntel() {
  const [myList,setMyList]           = useState(DEFAULT_PRICE_LIST);
  const [competitors,setCompetitors] = useState([]);
  const [overrides,setOverrides]     = useState([]);
  const [loading,setLoading]         = useState(true);
  const [syncStatus,setSyncStatus]   = useState('');
  const [activeTab,setActiveTab]     = useState('table');
  const [myPasteTab,setMyPasteTab]   = useState('excel');
  const [addingComp,setAddingComp]   = useState(false);
  const [newCompName,setNewCompName] = useState('');
  const [search,setSearch]           = useState('');
  const [filter,setFilter]           = useState('all');
  const [brandFilter,setBrandFilter] = useState('All');
  const [modelFilter,setModelFilter] = useState('All');
  const [variantFilter,setVariantFilter] = useState(new Set());
  const [compFilter,setCompFilter]   = useState(new Set());
  const [sortCol,setSortCol]         = useState('model');
  const [sortDir,setSortDir]         = useState(1);
  const [editingCell,setEditingCell] = useState(null);
  const [editVal,setEditVal]         = useState('');
  const [addRowOpen,setAddRowOpen]   = useState(false);
  const [addRowModel,setAddRowModel] = useState('');
  const [addRowVariant,setAddRowVariant] = useState('');
  const [addRowMyPrice,setAddRowMyPrice] = useState('');
  const [savedDates,setSavedDates]   = useState([]);
  const [recallDate,setRecallDate]   = useState('');
  const [sessionDropOpen,setSessionDropOpen] = useState(false);
  const [saveStatus,setSaveStatus]   = useState('');
  const fileRef = useRef(null);

  useEffect(()=>{ loadAll(); loadSavedDates(); const u=setupRealtime(); return ()=>u(); },[]);

  async function loadSavedDates(){
    const {data} = await supabase.from('daily_snapshots').select('date_key').order('date_key',{ascending:false});
    if(data) setSavedDates(data.map(d=>d.date_key));
  }

  async function saveSnapshot(){
    setSaveStatus('saving');
    const dateKey = new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'short',year:'numeric'}).replace(/ /g,'-');
    const snapshot = { myList, competitors, overrides };
    const {error} = await supabase.from('daily_snapshots').upsert(
      {date_key:dateKey, snapshot, saved_at:new Date().toISOString()},
      {onConflict:'date_key'}
    );
    if(error){ setSaveStatus('error'); console.error(error); setTimeout(()=>setSaveStatus(''),3000); return; }
    setSavedDates(prev=>[dateKey,...prev.filter(d=>d!==dateKey)].sort((a,b)=>b.localeCompare(a)));
    setSaveStatus('saved'); setRecallDate(dateKey);
    setTimeout(()=>setSaveStatus(''),2500);
  }

  async function recallSnapshot(dateKey){
    setSessionDropOpen(false);
    const {data} = await supabase.from('daily_snapshots').select('snapshot').eq('date_key',dateKey).single();
    if(!data) return;
    const snap = data.snapshot;
    setMyList(snap.myList || DEFAULT_PRICE_LIST);
    setCompetitors(snap.competitors || []);
    setOverrides(snap.overrides || []);
    setRecallDate(dateKey);
    setActiveTab('table');
  }

  async function deleteSnapshot(dateKey, e){
    e.stopPropagation();
    await supabase.from('daily_snapshots').delete().eq('date_key',dateKey);
    setSavedDates(prev=>prev.filter(d=>d!==dateKey));
    if(recallDate===dateKey) setRecallDate('');
  }


  async function loadAll(){
    setLoading(true);
    try {
      const [pr,cr,cpr,or] = await Promise.all([
        supabase.from('my_price_list').select('*').order('id'),
        supabase.from('competitors').select('*').order('created_at'),
        supabase.from('competitor_prices').select('*').order('competitor_id,id'),
        supabase.from('overrides').select('*'),
      ]);
      if(pr.data?.length) setMyList(pr.data.map(r=>({model:r.model,variant:r.variant,price:r.price})));
      if(cr.data){
        setCompetitors(cr.data.map(c=>({
          id:c.id, name:c.name,
          items:(cpr.data||[]).filter(p=>p.competitor_id===c.id).map(p=>({model:p.model,variant:p.variant,price:p.price})),
        })));
      }
      if(or.data) setOverrides(or.data);
    } finally { setLoading(false); }
  }

  function setupRealtime(){
    const ch = supabase.channel('realtime-all')
      .on('postgres_changes',{event:'*',schema:'public',table:'my_price_list'},loadAll)
      .on('postgres_changes',{event:'*',schema:'public',table:'competitors'},loadAll)
      .on('postgres_changes',{event:'*',schema:'public',table:'competitor_prices'},loadAll)
      .on('postgres_changes',{event:'*',schema:'public',table:'overrides'},loadAll)
      .subscribe();
    return ()=>supabase.removeChannel(ch);
  }

  function getOv(key,type,comp=''){
    const o=overrides.find(o=>o.row_key===key&&o.type===type&&o.comp_name===(comp||''));
    return o?o.value:null;
  }

  async function setOv(key,type,value,comp=''){
    setSyncStatus('saving');
    const val = String(value);
    const compName = comp||'';
    // Update local state immediately so UI responds without waiting for realtime
    setOverrides(prev => {
      const idx = prev.findIndex(o=>o.row_key===key&&o.type===type&&o.comp_name===compName);
      const newOv = {row_key:key,type,comp_name:compName,value:val,id:idx>=0?prev[idx].id:Date.now()};
      if(idx>=0){ const next=[...prev]; next[idx]=newOv; return next; }
      return [...prev, newOv];
    });
    const {error} = await supabase.from('overrides').upsert(
      {row_key:key,type,comp_name:compName,value:val},
      {onConflict:'row_key,type,comp_name'}
    );
    if(error) console.error('setOv error:', error);
    setSyncStatus('saved'); setTimeout(()=>setSyncStatus(''),2000);
  }

  const deletedKeys = useMemo(()=>new Set(overrides.filter(o=>o.type==='deleted').map(o=>o.row_key)),[overrides]);
  const manualRows  = useMemo(()=>overrides.filter(o=>o.type==='manualRow').map(o=>{try{return JSON.parse(o.value);}catch{return null;}}).filter(Boolean),[overrides]);

  const tableRows = useMemo(()=>{
    const rows={};
    for(const m of myList){
      const key=m.model+'|||'+m.variant;
      if(deletedKeys.has(key))continue;
      const pov=getOv(key,'myPrice');
      rows[key]={model:m.model,variant:m.variant,myPrice:pov!==null?parseInt(pov)||0:m.price,compPrices:{},isNew:false,key};
    }
    for(const comp of competitors){
      for(const item of comp.items){
        const match=findMyMatch(item.model,item.variant,myList);
        let key;
        if(match){key=match.model+'|||'+match.variant;}
        else{
          key=item.model+'|||'+(item.variant||'');
          if(!rows[key]){if(deletedKeys.has(key))continue;rows[key]={model:item.model,variant:item.variant||'',myPrice:0,compPrices:{},isNew:true,key};}
        }
        if(!rows[key]||deletedKeys.has(key))continue;
        const ov=getOv(key,'compPrice',comp.name);
        const price=ov!==null?parseInt(ov)||0:item.price;
        if(!rows[key].compPrices[comp.name]||price<rows[key].compPrices[comp.name]) rows[key].compPrices[comp.name]=price;
      }
    }
    for(const mr of manualRows){
      if(deletedKeys.has(mr.key)||rows[mr.key])continue;
      const pov=getOv(mr.key,'myPrice');
      rows[mr.key]={model:mr.model,variant:mr.variant,myPrice:pov!==null?parseInt(pov)||0:mr.myPrice,compPrices:{...(mr.compPrices||{})},isNew:false,isManual:true,key:mr.key};
    }
    // Apply compPrice overrides AFTER manual rows exist, so manually-added rows
    // also pick up prices typed directly into their competitor cells
    overrides.filter(o=>o.type==='compPrice'&&o.comp_name&&parseInt(o.value)>0).forEach(o=>{
      if(rows[o.row_key]&&!deletedKeys.has(o.row_key)) rows[o.row_key].compPrices[o.comp_name]=parseInt(o.value);
    });
    return Object.values(rows);
  },[myList,competitors,overrides,deletedKeys,manualRows]);

  const compNames = competitors.map(c=>c.name);
  const visibleCompNames = compNames.filter(n=>compFilter.size===0||compFilter.has(n));
  const brands = useMemo(()=>['All',...new Set(tableRows.map(r=>r.model.split(' ')[0]))].sort(),(tableRows));
  const allVariants = useMemo(()=>[...new Set(tableRows.map(r=>r.variant||'—'))].sort((a,b)=>{const p=s=>{const x=s.split('/');return[parseInt(x[0])||0,parseInt(x[1])||0];};const[ar,as]=p(a),[br,bs]=p(b);return ar!==br?ar-br:as-bs;}),[tableRows]);
  const allModels = useMemo(()=>[...new Set(tableRows.map(r=>r.model))].sort(),[tableRows]);

  let visible=tableRows;
  if(filter==='mine')  visible=visible.filter(r=>!r.isNew&&!r.isManual);
  if(filter==='new')   visible=visible.filter(r=>r.isNew);
  if(brandFilter!=='All') visible=visible.filter(r=>r.model.startsWith(brandFilter));
  if(modelFilter!=='All') visible=visible.filter(r=>r.model===modelFilter);
  if(variantFilter.size>0) visible=visible.filter(r=>variantFilter.has(r.variant||'—'));
  if(search.trim()){const q=search.toLowerCase();visible=visible.filter(r=>(r.model+' '+r.variant).toLowerCase().includes(q));}
  visible=[...visible].sort((a,b)=>{
    if(sortCol==='model')return sortDir*a.model.localeCompare(b.model);
    if(sortCol==='variant')return sortDir*a.variant.localeCompare(b.variant);
    if(sortCol==='myPrice')return sortDir*(a.myPrice-b.myPrice);
    return sortDir*((a.compPrices[sortCol]||0)-(b.compPrices[sortCol]||0));
  });

  function cheapest(row){const p=Object.values(row.compPrices).filter(p=>p>0);return p.length?Math.min(...p):null;}
  function diffColor(myP,cheap){if(!cheap||!myP)return C.muted;const pct=(myP-cheap)/cheap*100;return pct>3?C.red:pct<-3?C.accent:C.amber;}

  const stats={
    mine:tableRows.filter(r=>!r.isNew).length,
    newM:tableRows.filter(r=>r.isNew).length,
    higher:tableRows.filter(r=>{const c=cheapest(r);return c&&r.myPrice&&(r.myPrice-c)/c*100>3}).length,
    lower:tableRows.filter(r=>{const c=cheapest(r);return c&&r.myPrice&&(r.myPrice-c)/c*100<-3}).length,
  };

  async function saveMyList(items){
    setSyncStatus('saving');
    await supabase.from('my_price_list').delete().neq('id',0);
    if(items.length) await supabase.from('my_price_list').insert(items.map(({model,variant,price})=>({model,variant,price})));
    setSyncStatus('saved');setTimeout(()=>setSyncStatus(''),2000);
  }

  async function addCompetitor(){
    const name=newCompName.trim();
    if(!name||competitors.find(c=>c.name===name))return;
    const{data}=await supabase.from('competitors').insert({name}).select().single();
    if(data){setCompetitors(p=>[...p,{id:data.id,name:data.name,items:[]}]);setActiveTab(data.name);}
    setNewCompName('');setAddingComp(false);
  }

  async function removeCompetitor(id,name){
    await supabase.from('competitors').delete().eq('id',id);
    if(activeTab===name)setActiveTab('table');
  }

  async function saveCompItems(compId,name,items,ow){
    setSyncStatus('saving');
    if(ow){await supabase.from('competitor_prices').delete().eq('competitor_id',compId);}
    else{
      const{data:ex}=await supabase.from('competitor_prices').select('id,model,variant').eq('competitor_id',compId);
      const del=(ex||[]).filter(e=>items.some(i=>i.model.toLowerCase()===e.model.toLowerCase()&&i.variant===e.variant));
      if(del.length) await supabase.from('competitor_prices').delete().in('id',del.map(r=>r.id));
    }
    if(items.length) await supabase.from('competitor_prices').insert(items.map(({model,variant,price})=>({competitor_id:compId,model,variant,price})));
    const toClean=overrides.filter(o=>o.type==='compPrice'&&o.comp_name===name);
    for(const o of toClean) await supabase.from('overrides').delete().eq('id',o.id);
    setSyncStatus('saved');setTimeout(()=>setSyncStatus(''),2000);
    setActiveTab('table');
  }

  function startEdit(key,col,val){setEditingCell({key,col});setEditVal(String(val||''));}

  async function commitEdit(){
    if(!editingCell)return;
    const{key,col}=editingCell;
    if(col==='__model__'){if(editVal.trim())await setOv(key,'modelName',editVal.trim());}
    else{const num=parseInt(editVal.replace(/[^0-9]/g,''),10);if(!isNaN(num)&&num>=0)col==='myPrice'?await setOv(key,'myPrice',num):await setOv(key,'compPrice',num,col);}
    setEditingCell(null);setEditVal('');
  }

  async function deleteRow(key){await setOv(key,'deleted','true');}

  async function resetRowEdits(key){
    const ids=overrides.filter(o=>o.row_key===key&&['myPrice','compPrice','modelName'].includes(o.type)).map(o=>o.id);
    for(const id of ids) await supabase.from('overrides').delete().eq('id',id);
  }

  async function addManualRow(){
    const model=addRowModel.trim();
    const variant=addRowVariant.trim();
    const myPrice=parseInt(addRowMyPrice.replace(/[^0-9]/g,''),10)||0;
    if(!model)return;
    // Use timestamp to guarantee unique key
    const key='manual_'+Date.now()+'_'+Math.random().toString(36).slice(2,7);
    await setOv(key,'manualRow',JSON.stringify({key,model,variant,myPrice,compPrices:{}}));
    setAddRowModel('');setAddRowVariant('');setAddRowMyPrice('');setAddRowOpen(false);
  }

  async function handleMyExcel(e){
    const file=e.target.files[0];if(!file)return;
    try{const items=await parseExcelFile(file);if(!items.length)throw new Error('No valid rows');await saveMyList(items);}
    catch(err){alert(err.message);}
    e.target.value='';
  }

  const today=new Date().toLocaleDateString('en-IN',{day:'2-digit',month:'short',year:'numeric'}).replace(/ /g,'-');

  if(loading)return(
    <div style={{minHeight:'100vh',background:C.bg,display:'flex',alignItems:'center',justifyContent:'center',color:C.muted,fontFamily:'Inter,sans-serif',fontSize:15}}>
      <div style={{textAlign:'center'}}><div style={{fontSize:36,marginBottom:12}}>📱</div><div>Loading Price Intel…</div></div>
    </div>
  );

  const SideBtn=({icon,label,badge,active,color=C.accent,onClick})=>(
    <div onClick={onClick} style={{padding:'8px 12px',borderRadius:6,cursor:'pointer',background:active?color+'18':'transparent',color:active?color:C.textDim,border:`1px solid ${active?color+'44':'transparent'}`,display:'flex',justifyContent:'space-between',alignItems:'center',fontSize:13}}>
      <span>{icon} {label}</span>
      {badge>0&&<span style={{background:color,color:'#000',borderRadius:10,padding:'1px 7px',fontSize:10,fontWeight:700}}>{badge}</span>}
    </div>
  );

  const renderPriceCell=(col,value,color,rowKey)=>{
    const isActive=editingCell?.key===rowKey&&editingCell?.col===col;
    const wasEdited=overrides.some(o=>o.row_key===rowKey&&(col==='myPrice'?o.type==='myPrice':o.type==='compPrice'&&o.comp_name===col));
    if(isActive)return(
      <td key={col} style={{padding:'4px 8px',textAlign:'right'}}>
        <input autoFocus value={editVal} onChange={e=>setEditVal(e.target.value)}
          onBlur={commitEdit} onKeyDown={e=>{if(e.key==='Enter')commitEdit();if(e.key==='Escape'){setEditingCell(null);setEditVal('');}}}
          placeholder="Enter price"
          style={{width:90,background:C.surface2,color:C.text,border:`2px solid ${C.accent}`,borderRadius:4,padding:'4px 8px',fontSize:13,textAlign:'right',fontFamily:'monospace',outline:'none'}}/>
      </td>
    );
    return(
      <td key={col} onClick={()=>startEdit(rowKey,col,value||0)} title="Click to edit"
        style={{padding:'8px 12px',textAlign:'right',fontFamily:'monospace',color:value>0?color:C.muted+'77',fontWeight:value>0?600:400,cursor:'pointer',position:'relative',userSelect:'none'}}>
        <span style={{borderBottom:wasEdited?`2px dashed ${C.blue}`:'none',paddingBottom:wasEdited?1:0}}>{value>0?fmt(value):'+ add'}</span>
        {wasEdited&&<span style={{position:'absolute',top:3,right:3,fontSize:8,color:C.blue,fontWeight:700}}>✎</span>}
      </td>
    );
  };

  return(
    <div style={{minHeight:'100vh',background:C.bg,color:C.text,fontFamily:"'Inter','Segoe UI',sans-serif",fontSize:14,display:'flex',flexDirection:'column'}}>
      {/* Header */}
      <div style={{borderBottom:`1px solid ${C.border}`,padding:'13px 20px',display:'flex',alignItems:'center',justifyContent:'space-between',flexShrink:0,flexWrap:'wrap',gap:10}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
          <img src="/logo.jpeg" alt="Shri Maa" style={{width:30,height:30,borderRadius:'50%',objectFit:'cover',flexShrink:0}}/>
          <span style={{fontSize:22}}>📱</span>
          <span style={{fontWeight:800,fontSize:17,letterSpacing:'-0.02em'}}>Price Intel</span>
          <span style={{color:C.muted,fontWeight:400,fontSize:13}}> — Daily Competitor Tracker</span>
          {recallDate&&<span style={{fontSize:11,fontWeight:700,padding:'2px 8px',borderRadius:3,background:C.blueDim,color:C.blue,marginLeft:4}}>📅 {recallDate}</span>}
          {syncStatus&&<span style={{fontSize:11,fontWeight:700,padding:'2px 8px',borderRadius:3,background:syncStatus==='saving'?C.amberDim:C.accentDim,color:syncStatus==='saving'?C.amber:C.accent}}>{syncStatus==='saving'?'⟳ Syncing…':'✓ Synced'}</span>}
        </div>
        <div style={{display:'flex',alignItems:'center',gap:14,flexWrap:'wrap'}}>
          {/* Save snapshot button */}
          <button onClick={saveSnapshot} disabled={saveStatus==='saving'} style={{
            display:'flex',alignItems:'center',gap:6,padding:'7px 14px',borderRadius:6,border:'none',
            background:saveStatus==='saved'?C.accent:saveStatus==='error'?C.red:C.surface2,
            color:saveStatus==='saved'?'#000':saveStatus==='error'?'#fff':C.textDim,
            fontWeight:700,fontSize:12,cursor:'pointer',whiteSpace:'nowrap',
            border:`1px solid ${saveStatus==='saved'?C.accent:saveStatus==='error'?C.red:C.border}`,
          }}>
            {saveStatus==='saving'?'Saving…':saveStatus==='saved'?'✓ Saved!':saveStatus==='error'?'✗ Error':'💾 Save'}
          </button>

          {/* Recall dropdown */}
          {savedDates.length>0&&(
            <div style={{position:'relative'}}>
              <button onClick={()=>setSessionDropOpen(o=>!o)} style={{
                display:'flex',alignItems:'center',gap:6,padding:'7px 14px',borderRadius:6,
                background:recallDate?C.blueDim:C.surface2,color:recallDate?C.blue:C.textDim,
                border:`1px solid ${recallDate?C.blue+'55':C.border}`,fontWeight:recallDate?700:400,fontSize:12,
                cursor:'pointer',whiteSpace:'nowrap',
              }}>
                {recallDate?`📅 ${recallDate}`:'📅 Recall'}
                <span style={{fontSize:9,color:C.muted}}>▼</span>
                {recallDate&&<span onClick={e=>{e.stopPropagation();setRecallDate('');loadAll();}} style={{color:C.muted,fontSize:12}}>✕</span>}
              </button>
              {sessionDropOpen&&(
                <>
                  <div onClick={()=>setSessionDropOpen(false)} style={{position:'fixed',inset:0,zIndex:99}}/>
                  <div style={{position:'absolute',top:'calc(100% + 4px)',right:0,zIndex:100,background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,boxShadow:'0 8px 24px #0008',minWidth:200,maxHeight:320,overflowY:'auto',padding:'6px 0'}}>
                    <div style={{padding:'6px 14px 8px',borderBottom:`1px solid ${C.border}`,fontSize:11,color:C.muted,fontWeight:700,letterSpacing:'0.06em',textTransform:'uppercase'}}>Saved Dates</div>
                    {savedDates.map(d=>(
                      <div key={d} onClick={()=>recallSnapshot(d)} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'9px 14px',cursor:'pointer',fontSize:13,color:d===recallDate?C.blue:C.textDim,background:d===recallDate?C.blueDim:'transparent'}}
                        onMouseEnter={e=>{if(d!==recallDate)e.currentTarget.style.background=C.border;}}
                        onMouseLeave={e=>{if(d!==recallDate)e.currentTarget.style.background='transparent';}}>
                        <span>📅 {d}</span>
                        <span onClick={e=>deleteSnapshot(d,e)} style={{color:C.muted,fontSize:13,padding:'2px 5px'}}
                          onMouseEnter={e=>e.currentTarget.style.color=C.red} onMouseLeave={e=>e.currentTarget.style.color=C.muted}>🗑</span>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>
          )}

          {competitors.some(c=>c.items.length>0)&&(
            <div style={{display:'flex',gap:16,fontSize:12}}>
              <span style={{color:C.muted}}>My models: <strong style={{color:C.text}}>{stats.mine}</strong></span>
              <span style={{color:C.red}}>Expensive: <strong>{stats.higher}</strong></span>
              <span style={{color:C.accent}}>Cheaper: <strong>{stats.lower}</strong></span>
              {stats.newM>0&&<span style={{color:C.amber}}>New: <strong>{stats.newM}</strong></span>}
            </div>
          )}
        </div>
      </div>

      <div style={{display:'flex',flex:1,overflow:'hidden'}}>
        {/* Sidebar */}
        <div style={{width:212,borderRight:`1px solid ${C.border}`,padding:12,flexShrink:0,display:'flex',flexDirection:'column',gap:4,overflowY:'auto'}}>
          <div style={{fontSize:10,fontWeight:700,color:C.muted,letterSpacing:'0.1em',textTransform:'uppercase',padding:'4px 4px 6px'}}>Views</div>
          <SideBtn icon="📊" label="Comparison" active={activeTab==='table'} badge={competitors.some(c=>c.items.length>0)?visible.length:0} onClick={()=>setActiveTab('table')}/>
          <SideBtn icon="📋" label="My Price List" active={activeTab==='mylist'} badge={myList.length} onClick={()=>setActiveTab('mylist')}/>
          <div style={{fontSize:10,fontWeight:700,color:C.muted,letterSpacing:'0.1em',textTransform:'uppercase',padding:'12px 4px 6px'}}>Competitors</div>
          {competitors.map(c=>(
            <div key={c.id} onClick={()=>setActiveTab(c.name)} style={{padding:'8px 12px',borderRadius:6,cursor:'pointer',background:activeTab===c.name?C.amberDim:'transparent',color:activeTab===c.name?C.amber:C.textDim,border:`1px solid ${activeTab===c.name?C.amber+'44':'transparent'}`,display:'flex',justifyContent:'space-between',alignItems:'center',fontSize:13}}>
              <span>🏪 {c.name}</span>
              <div style={{display:'flex',gap:5,alignItems:'center'}}>
                {c.items.length>0&&<span style={{background:C.amber,color:'#000',borderRadius:10,padding:'1px 6px',fontSize:10,fontWeight:700}}>{c.items.length}</span>}
                <span onClick={e=>{e.stopPropagation();removeCompetitor(c.id,c.name);}} style={{color:C.muted,fontSize:11,padding:'1px 4px',cursor:'pointer'}}>✕</span>
              </div>
            </div>
          ))}
          {addingComp?(
            <div style={{display:'flex',gap:4,marginTop:4}}>
              <input autoFocus value={newCompName} onChange={e=>setNewCompName(e.target.value)} onKeyDown={e=>{if(e.key==='Enter')addCompetitor();if(e.key==='Escape')setAddingComp(false);}} placeholder="Name…"
                style={{flex:1,background:C.bg,color:C.text,border:`1px solid ${C.amber}`,borderRadius:5,padding:'6px 8px',fontSize:12,outline:'none'}}/>
              <button onClick={addCompetitor} style={{background:C.amber,color:'#000',border:'none',borderRadius:5,padding:'6px 10px',cursor:'pointer',fontWeight:700,fontSize:13}}>+</button>
            </div>
          ):(
            <button onClick={()=>setAddingComp(true)} style={{background:'transparent',border:`1px dashed ${C.border}`,color:C.muted,borderRadius:6,padding:'7px 12px',cursor:'pointer',fontSize:12,textAlign:'left',marginTop:4}}>+ Add Competitor</button>
          )}
        </div>

        {/* Main */}
        <div style={{flex:1,overflowY:'auto',overflowX:'auto'}}>

          {/* My Price List */}
          {activeTab==='mylist'&&(
            <div style={{padding:24,maxWidth:820}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:20,flexWrap:'wrap',gap:12}}>
                <div><h2 style={{fontWeight:700,margin:0}}>My Price List</h2><p style={{color:C.muted,fontSize:13,margin:'4px 0 0'}}>{myList.length} variants — synced across all devices in real time</p></div>
                <Pill color={C.accent}>SHARED</Pill>
              </div>
              <div style={{display:'flex',gap:0,marginBottom:20,borderBottom:`1px solid ${C.border}`}}>
                {[['excel','📂 Upload Excel'],['whatsapp','💬 Paste WhatsApp']].map(([t,label])=>(
                  <button key={t} onClick={()=>setMyPasteTab(t)} style={{padding:'9px 20px',border:'none',cursor:'pointer',fontSize:13,fontWeight:600,background:'transparent',color:myPasteTab===t?C.accent:C.muted,borderBottom:`2px solid ${myPasteTab===t?C.accent:'transparent'}`,marginBottom:-1}}>{label}</button>
                ))}
              </div>
              {myPasteTab==='excel'&&(
                <div style={{border:`2px dashed ${C.border}`,borderRadius:10,padding:24,marginBottom:24,background:C.surface,textAlign:'center'}}>
                  <div style={{fontSize:28,marginBottom:8}}>📂</div>
                  <div style={{fontWeight:600,marginBottom:4}}>Upload Today's Excel</div>
                  <div style={{color:C.muted,fontSize:12,marginBottom:16}}>Columns: <code style={{background:C.bg,padding:'2px 6px',borderRadius:3}}>Model</code> <code style={{background:C.bg,padding:'2px 6px',borderRadius:3}}>Variant</code> <code style={{background:C.bg,padding:'2px 6px',borderRadius:3}}>Price</code></div>
                  <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleMyExcel} style={{display:'none'}}/>
                  <button onClick={()=>fileRef.current?.click()} style={{padding:'10px 28px',borderRadius:6,border:'none',background:C.accent,color:'#000',fontWeight:700,fontSize:13,cursor:'pointer'}}>Choose Excel File</button>
                </div>
              )}
              {myPasteTab==='whatsapp'&&(
                <PastePanel name="My List" existingItems={myList} myModels={myList}
                  onConfirm={(items,ow)=>{
                    if(ow){saveMyList(items);}
                    else{const merged=[...myList];for(const ni of items){const idx=merged.findIndex(e=>e.model.toLowerCase()===ni.model.toLowerCase()&&e.variant===ni.variant);if(idx>=0)merged[idx]=ni;else merged.push(ni);}saveMyList(merged);}
                  }}/>
              )}
              <div style={{display:'flex',flexDirection:'column',gap:2}}>
                <div style={{display:'flex',justifyContent:'space-between',padding:'6px 12px',background:C.surface2,borderRadius:'6px 6px 0 0',fontSize:11,fontWeight:700,color:C.muted,letterSpacing:'0.06em',textTransform:'uppercase'}}>
                  <span>Model</span><div style={{display:'flex',gap:56}}><span>Variant</span><span>Price</span></div>
                </div>
                {myList.map((m,i)=>(
                  <div key={i} style={{display:'flex',justifyContent:'space-between',padding:'7px 12px',background:i%2===0?C.surface:'transparent',fontSize:13}}>
                    <span style={{color:C.textDim}}>{m.model}</span>
                    <div style={{display:'flex',gap:20}}>
                      <span style={{color:C.muted,fontFamily:'monospace',width:50,textAlign:'right'}}>{m.variant||'—'}</span>
                      <span style={{color:C.accent,fontFamily:'monospace',width:82,textAlign:'right',fontWeight:600}}>{fmt(m.price)}</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Competitor panel */}
          {competitors.find(c=>c.name===activeTab)&&(()=>{
            const comp=competitors.find(c=>c.name===activeTab);
            return(
              <div style={{maxWidth:840}}>
                <div style={{padding:'24px 24px 0',display:'flex',alignItems:'center',gap:10}}>
                  <h2 style={{fontWeight:700,margin:0}}>{activeTab}</h2>
                  <Pill color={C.amber}>COMPETITOR</Pill>
                  {comp.items.length>0&&<Pill color={C.muted}>{comp.items.length} models</Pill>}
                </div>
                <PastePanel name={activeTab} existingItems={comp.items} myModels={myList}
                  onConfirm={(items,ow)=>saveCompItems(comp.id,comp.name,items,ow)}/>
              </div>
            );
          })()}

          {/* Comparison Table */}
          {activeTab==='table'&&(
            <div style={{padding:20}}>
              {!competitors.some(c=>c.items.length>0)&&competitors.length===0?(
                <div style={{textAlign:'center',paddingTop:80,color:C.muted}}>
                  <div style={{fontSize:48,marginBottom:16}}>📊</div>
                  <div style={{fontSize:16,fontWeight:600,color:C.textDim,marginBottom:8}}>Ready to compare</div>
                  <div style={{fontSize:13,maxWidth:400,margin:'0 auto'}}>
                    Your {myList.length} models are synced to the database. Add a competitor from the sidebar — everyone on the team sees updates instantly.
                  </div>
                </div>
              ):(
                <>
                  {/* Controls */}
                  <div style={{display:'flex',gap:8,marginBottom:12,flexWrap:'wrap',alignItems:'center'}}>
                    <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search model…"
                      style={{background:C.surface,color:C.text,border:`1px solid ${C.border}`,borderRadius:6,padding:'7px 12px',fontSize:13,outline:'none',width:185}}/>
                    <select value={brandFilter} onChange={e=>setBrandFilter(e.target.value)} style={{background:C.surface,color:C.text,border:`1px solid ${C.border}`,borderRadius:6,padding:'7px 10px',fontSize:13,outline:'none',cursor:'pointer'}}>
                      {['All',...new Set(tableRows.map(r=>r.model.split(' ')[0]))].sort().map(b=><option key={b}>{b}</option>)}
                    </select>
                    <select value={modelFilter} onChange={e=>setModelFilter(e.target.value)} style={{background:C.surface,color:C.text,border:`1px solid ${C.border}`,borderRadius:6,padding:'7px 10px',fontSize:13,outline:'none',cursor:'pointer',maxWidth:180}}>
                      <option value="All">All Models</option>
                      {allModels.map(m=><option key={m} value={m}>{m}</option>)}
                    </select>
                    <MultiSelect label="All Variants" options={allVariants} selected={variantFilter} onChange={setVariantFilter} color={C.accent}/>
                    {compNames.length>0&&<MultiSelect label="All Competitors" options={compNames} selected={compFilter} onChange={setCompFilter} color={C.amber}/>}
                    {['all','mine','new'].map(f=>(
                      <button key={f} onClick={()=>setFilter(f)} style={{padding:'7px 13px',borderRadius:6,border:'none',cursor:'pointer',fontSize:12,fontWeight:600,background:filter===f?C.accent:C.surface,color:filter===f?'#000':C.muted}}>
                        {f==='all'?'All':f==='mine'?'My Models':'🆕 New'}
                      </button>
                    ))}
                    <span style={{flex:1}}/>

                    <button onClick={()=>downloadExcel(visible,visibleCompNames,today)} disabled={visible.length===0}
                      style={{display:'flex',alignItems:'center',gap:7,padding:'8px 16px',borderRadius:6,border:`1px solid ${C.accent}44`,background:C.accentDim,color:C.accent,fontWeight:700,fontSize:12,cursor:'pointer',whiteSpace:'nowrap',opacity:visible.length===0?0.4:1}}>
                      ⬇ Download Excel
                    </button>
                  </div>

                  {/* Legend */}
                  <div style={{display:'flex',gap:14,marginBottom:10,fontSize:11,color:C.muted,flexWrap:'wrap',alignItems:'center'}}>
                    <span><span style={{color:C.red}}>■</span> My price &gt;3% higher</span>
                    <span><span style={{color:C.amber}}>■</span> Within ±3%</span>
                    <span><span style={{color:C.accent}}>■</span> My price lower</span>
                    <span><span style={{color:C.amber}}>■</span> NEW = not in catalog</span>
                    <span>· Click price to edit · 🗑 delete</span>
                    {deletedKeys.size>0&&(
                      <button onClick={async()=>{const ids=overrides.filter(o=>o.type==='deleted').map(o=>o.id);for(const id of ids)await supabase.from('overrides').delete().eq('id',id);}} style={{background:C.redDim,border:`1px solid ${C.red}44`,color:C.red,borderRadius:5,padding:'3px 10px',fontSize:11,cursor:'pointer',fontWeight:600}}>
                        ↺ Restore {deletedKeys.size} deleted
                      </button>
                    )}
                    <span style={{marginLeft:'auto',color:C.muted}}>{visible.length} rows</span>
                  </div>

                  {/* Table */}
                  <div style={{borderRadius:8,border:`1px solid ${C.border}`,overflowX:'auto'}}>
                    <table style={{width:'100%',borderCollapse:'collapse',fontSize:13,minWidth:500}}>
                      <thead>
                        <tr style={{background:C.surface2,borderBottom:`1px solid ${C.border}`}}>
                          {['model','variant','myPrice',...visibleCompNames].map(col=>(
                            <th key={col} onClick={()=>{if(sortCol===col)setSortDir(d=>-d);else{setSortCol(col);setSortDir(1);}}}
                              style={{padding:'10px 12px',textAlign:col==='model'||col==='variant'?'left':'right',cursor:'pointer',color:sortCol===col?C.accent:C.muted,fontWeight:700,fontSize:10,letterSpacing:'0.06em',textTransform:'uppercase',whiteSpace:'nowrap',userSelect:'none'}}>
                              {col==='myPrice'?'MY PRICE':col.toUpperCase()}{sortCol===col&&(sortDir===1?' ↑':' ↓')}
                            </th>
                          ))}
                          <th style={{padding:'10px 12px',textAlign:'right',color:C.muted,fontWeight:700,fontSize:10,textTransform:'uppercase',whiteSpace:'nowrap'}}>DIFF</th>
                          <th style={{padding:'10px 12px',width:56}}/>
                        </tr>
                      </thead>
                      <tbody>
                        {visible.map((row,i)=>{
                          const cheap=cheapest(row);
                          const dc=diffColor(row.myPrice,cheap);
                          const diffPct=row.myPrice&&cheap?((row.myPrice-cheap)/cheap*100).toFixed(1):null;
                          const rowKey=row.key||(row.model+'|||'+row.variant);
                          const modelOv=getOv(rowKey,'modelName');
                          const hasEdits=overrides.some(o=>o.row_key===rowKey&&['myPrice','compPrice','modelName'].includes(o.type));
                          const isModelEditing=editingCell?.key===rowKey&&editingCell?.col==='__model__';
                          const displayModel=modelOv||row.model;
                          return(
                            <tr key={rowKey} style={{borderBottom:`1px solid ${C.border}`,background:row.isNew?C.amberDim:hasEdits?C.blueDim:i%2===0?'transparent':C.surface+'55'}}>
                              <td style={{padding:'8px 12px'}}>
                                {isModelEditing?(
                                  <input autoFocus value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit}
                                    onKeyDown={e=>{if(e.key==='Enter')commitEdit();if(e.key==='Escape'){setEditingCell(null);setEditVal('');}}}
                                    style={{width:'100%',background:C.surface2,color:C.text,border:`2px solid ${C.accent}`,borderRadius:4,padding:'4px 8px',fontSize:13,outline:'none',boxSizing:'border-box'}}/>
                                ):(
                                  <span onClick={()=>{setEditingCell({key:rowKey,col:'__model__'});setEditVal(displayModel);}} title="Click to edit" style={{cursor:'pointer',borderBottom:modelOv?`2px dashed ${C.blue}`:'none'}}>{displayModel}</span>
                                )}
                                {row.isNew&&<span style={{marginLeft:7}}><Pill color={C.amber}>NEW</Pill></span>}
                                {row.isManual&&<span style={{marginLeft:7}}><Pill color={C.blue}>MANUAL</Pill></span>}
                                {hasEdits&&<span style={{marginLeft:6}}><Pill color={C.blue}>EDITED</Pill></span>}
                              </td>
                              <td style={{padding:'8px 12px',color:C.muted,fontFamily:'monospace'}}>{row.variant||'—'}</td>
                              {renderPriceCell('myPrice',row.myPrice,C.accent,rowKey)}
                              {visibleCompNames.map(cn=>renderPriceCell(cn,row.compPrices[cn]||0,C.amber,rowKey))}
                              <td style={{padding:'8px 12px',textAlign:'right',fontFamily:'monospace',color:dc,fontWeight:700,fontSize:12}}>
                                {diffPct!==null?(Number(diffPct)>0?`+${diffPct}%`:`${diffPct}%`):'—'}
                              </td>
                              <td style={{padding:'4px 6px',textAlign:'center',width:56}}>
                                <div style={{display:'flex',gap:3,justifyContent:'center',alignItems:'center'}}>
                                  {hasEdits&&(
                                    <span onClick={()=>resetRowEdits(rowKey)} title="Reset edits"
                                      style={{display:'inline-flex',alignItems:'center',justifyContent:'center',width:22,height:22,borderRadius:4,cursor:'pointer',fontSize:11,color:C.blue,border:'1px solid transparent'}}
                                      onMouseEnter={e=>{e.currentTarget.style.borderColor=C.blue+'55';e.currentTarget.style.background=C.blueDim;}}
                                      onMouseLeave={e=>{e.currentTarget.style.borderColor='transparent';e.currentTarget.style.background='transparent';}}>↺</span>
                                  )}
                                  <span onClick={()=>deleteRow(rowKey)} title="Delete row"
                                    style={{display:'inline-flex',alignItems:'center',justifyContent:'center',width:22,height:22,borderRadius:4,cursor:'pointer',fontSize:12,color:C.muted,border:'1px solid transparent'}}
                                    onMouseEnter={e=>{e.currentTarget.style.color=C.red;e.currentTarget.style.borderColor=C.red+'55';e.currentTarget.style.background=C.redDim;}}
                                    onMouseLeave={e=>{e.currentTarget.style.color=C.muted;e.currentTarget.style.borderColor='transparent';e.currentTarget.style.background='transparent';}}>🗑</span>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>

                    {visible.length===0&&<div style={{padding:40,textAlign:'center',color:C.muted}}>No models match your filter.</div>}

                    {/* Add Row */}
                    {!addRowOpen?(
                      <div onClick={()=>setAddRowOpen(true)} style={{display:'flex',alignItems:'center',gap:8,padding:'10px 14px',cursor:'pointer',color:C.muted,fontSize:13,borderTop:`1px solid ${C.border}`}}
                        onMouseEnter={e=>e.currentTarget.style.background=C.surface} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
                        <span style={{color:C.accent,fontSize:16,fontWeight:700,lineHeight:1}}>+</span>
                        <span>Add row manually</span>
                      </div>
                    ):(
                      <div style={{borderTop:`1px solid ${C.border}`,padding:'12px 14px',background:C.surface}}>
                        <div style={{fontSize:11,fontWeight:700,color:C.muted,letterSpacing:'0.08em',textTransform:'uppercase',marginBottom:10}}>New Row</div>
                        <div style={{display:'flex',gap:8,alignItems:'flex-end',flexWrap:'wrap'}}>
                          {[
                            {label:'Model Name *',val:addRowModel,set:setAddRowModel,ph:'e.g. Samsung S25 FE (Fresh)',flex:'2 1 200px',mono:false},
                            {label:'Variant',val:addRowVariant,set:setAddRowVariant,ph:'e.g. 8/128',flex:'0 1 100px',mono:true},
                            {label:'My Price',val:addRowMyPrice,set:setAddRowMyPrice,ph:'e.g. 44300',flex:'0 1 110px',mono:true},
                          ].map(({label,val,set,ph,flex,mono})=>(
                            <div key={label} style={{display:'flex',flexDirection:'column',gap:4,flex}}>
                              <label style={{fontSize:11,color:C.muted}}>{label}</label>
                              <input value={val} onChange={e=>set(e.target.value)}
                                onKeyDown={e=>{if(e.key==='Enter')addManualRow();if(e.key==='Escape')setAddRowOpen(false);}}
                                placeholder={ph}
                                style={{background:C.bg,color:C.text,border:`1px solid ${C.border}`,borderRadius:6,padding:'8px 10px',fontSize:13,outline:'none',fontFamily:mono?'monospace':'inherit'}}
                                onFocus={e=>e.target.style.borderColor=C.accent} onBlur={e=>e.target.style.borderColor=C.border}/>
                            </div>
                          ))}
                          <div style={{display:'flex',gap:6}}>
                            <button onClick={addManualRow} disabled={!addRowModel.trim()} style={{padding:'8px 18px',borderRadius:6,border:'none',background:addRowModel.trim()?C.accent:C.border,color:'#000',fontWeight:700,fontSize:13,cursor:addRowModel.trim()?'pointer':'default'}}>Add</button>
                            <button onClick={()=>{setAddRowOpen(false);setAddRowModel('');setAddRowVariant('');setAddRowMyPrice('');}} style={{padding:'8px 14px',borderRadius:6,border:`1px solid ${C.border}`,background:'transparent',color:C.muted,fontSize:13,cursor:'pointer'}}>Cancel</button>
                          </div>
                        </div>
                        <div style={{fontSize:11,color:C.muted,marginTop:8}}>After adding, click any price cell to fill competitor prices.</div>
                      </div>
                    )}
                  </div>
                </>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
