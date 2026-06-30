import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ê³µí†µ ë°ì´í„°
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const MONTHLY_TREND = [
  {month:"1ì›”",income:4200000,expense:2850000},
  {month:"2ì›”",income:4200000,expense:3120000},
  {month:"3ì›”",income:4500000,expense:2680000},
  {month:"4ì›”",income:4200000,expense:3450000},
  {month:"5ì›”",income:4700000,expense:2990000},
  {month:"6ì›”",income:4200000,expense:1103000},
];
const MONTHLY_CATS = [
  {name:"ì‹ë¹„",  amount:387000,icon:"ðŸš",color:"#F97316",avg:280000},
  {name:"ì‡¼í•‘",  amount:298000,icon:"ðŸ›ï¸",color:"#EC4899",avg:150000},
  {name:"ì¹´íŽ˜",  amount:142500,icon:"â˜•", color:"#A78BFA",avg:95000 },
  {name:"êµí†µ",  amount:84200, icon:"ðŸšŒ",color:"#3B82F6",avg:80000 },
  {name:"íŽ¸ì˜ì ",amount:68700, icon:"ðŸª",color:"#10B981",avg:55000 },
  {name:"ê¸°íƒ€",  amount:52300, icon:"ðŸ“¦",color:"#94A3B8",avg:60000 },
];
const CAT_TREND = {
  ì‹ë¹„:  [280000,310000,265000,320000,295000,387000],
  ì‡¼í•‘:  [150000,210000,130000,280000,160000,298000],
  ì¹´íŽ˜:  [95000, 88000, 102000,115000,91000, 142500],
  êµí†µ:  [80000, 75000, 83000, 79000, 82000, 84200 ],
  íŽ¸ì˜ì :[55000, 62000, 48000, 71000, 58000, 68700 ],
  ê¸°íƒ€:  [60000, 55000, 72000, 48000, 63000, 52300 ],
};
const WEEKLY_CARDS = [
  {name:"ì‹ í•œì¹´ë“œ",amount:87500, color:"#3B82F6",emoji:"ðŸ’³"},
  {name:"í˜„ëŒ€ì¹´ë“œ",amount:124000,color:"#8B5CF6",emoji:"ðŸ’œ"},
  {name:"ì‚¼ì„±ì¹´ë“œ",amount:43200, color:"#10B981",emoji:"ðŸ’š"},
];
const MONTHLY_CARDS = [
  {name:"ì‹ í•œì¹´ë“œ",amount:342000,limit:500000,color:"#3B82F6"},
  {name:"í˜„ëŒ€ì¹´ë“œ",amount:487000,limit:600000,color:"#8B5CF6"},
  {name:"ì‚¼ì„±ì¹´ë“œ",amount:198000,limit:400000,color:"#10B981"},
];
const FAMILY = {
  code:"MONEY-7X3K",
  members:[
    {id:1,name:"ë‚˜",   role:"ê´€ë¦¬ìž",emoji:"ðŸ‘¨",color:"#6366F1",joined:"2026.01.01"},
    {id:2,name:"ë°°ìš°ìž",role:"ë©¤ë²„",  emoji:"ðŸ‘©",color:"#EC4899",joined:"2026.01.03"},
    {id:3,name:"ìžë…€1",role:"ë·°ì–´",  emoji:"ðŸ‘¦",color:"#10B981",joined:"2026.03.15"},
  ],
};
const INIT_TX = [
  {id:1,name:"ìŠ¤íƒ€ë²…ìŠ¤",  category:"ì¹´íŽ˜",  amount:-6500, date:"2026.06.19",time:"10:23",card:"ì‹ í•œì¹´ë“œ",icon:"â˜•", author:1,memo:""},
  {id:2,name:"GS25",      category:"íŽ¸ì˜ì ",amount:-4200, date:"2026.06.19",time:"08:11",card:"í˜„ëŒ€ì¹´ë“œ",icon:"ðŸª", author:2,memo:""},
  {id:3,name:"ë‹¹ê·¼ë§ˆì¼“",  category:"ìˆ˜ìž…",  amount:+35000,date:"2026.06.18",time:"19:45",card:"-",      icon:"ðŸ¥•", author:1,memo:"ì¤‘ê³  íŒë§¤"},
  {id:4,name:"ì¿ íŒ¡",      category:"ì‡¼í•‘",  amount:-29800,date:"2026.06.18",time:"14:22",card:"ì‚¼ì„±ì¹´ë“œ",icon:"ðŸ“¦", author:2,memo:""},
  {id:5,name:"ì˜¬ë¦¬ë¸Œì˜",  category:"ë·°í‹°",  amount:-18500,date:"2026.06.16",time:"17:45",card:"ì‹ í•œì¹´ë“œ",icon:"ðŸŒ¿", author:1,memo:""},
  {id:6,name:"ì´ë§ˆíŠ¸24",  category:"íŽ¸ì˜ì ",amount:-3200, date:"2026.06.15",time:"09:02",card:"í˜„ëŒ€ì¹´ë“œ",icon:"ðŸª", author:3,memo:"í•™êµ ì•ž"},
  {id:7,name:"ë²„ê±°í‚¹",    category:"ì‹ë¹„",  amount:-12500,date:"2026.06.14",time:"13:30",card:"ì‚¼ì„±ì¹´ë“œ",icon:"ðŸ”", author:1,memo:""},
  {id:8,name:"ì¹´ì¹´ì˜¤íŽ˜ì´",category:"ìˆ˜ìž…",  amount:+20000,date:"2026.06.13",time:"11:00",card:"-",      icon:"ðŸ’›", author:2,memo:"ìš©ëˆ"},
];
const SMS_RAW = [
  {id:1,raw:"[ì‹ í•œì¹´ë“œ] 06/19 10:23 ìŠ¤íƒ€ë²…ìŠ¤ê°•ë‚¨ì  6,500ì› ìŠ¹ì¸", card:"ì‹ í•œì¹´ë“œ",amount:6500, date:"06/19 10:23"},
  {id:2,raw:"[í˜„ëŒ€ì¹´ë“œ] 06/19 08:11 GS25ìƒì¼ë™ì  4,200ì› ìŠ¹ì¸",   card:"í˜„ëŒ€ì¹´ë“œ",amount:4200, date:"06/19 08:11"},
  {id:3,raw:"[ì‚¼ì„±ì¹´ë“œ] 06/18 14:22 ì¿ íŒ¡ 29,800ì› ìŠ¹ì¸",          card:"ì‚¼ì„±ì¹´ë“œ",amount:29800,date:"06/18 14:22"},
  {id:4,raw:"[ì‹ í•œì¹´ë“œ] 06/18 17:45 ì˜¬ë¦¬ë¸Œì˜ê³ ë•ì  18,500ì› ìŠ¹ì¸",card:"ì‹ í•œì¹´ë“œ",amount:18500,date:"06/18 17:45"},
  {id:5,raw:"[í˜„ëŒ€ì¹´ë“œ] 06/17 09:02 ì´ë§ˆíŠ¸24 3,200ì› ìŠ¹ì¸",       card:"í˜„ëŒ€ì¹´ë“œ",amount:3200, date:"06/17 09:02"},
];
const CARD_PROVIDERS = [
  {id:"shinhan",name:"ì‹ í•œì¹´ë“œ",prefix:"[ì‹ í•œì¹´ë“œ]",sender:"15440092",emoji:"ðŸ’™"},
  {id:"hyundai",name:"í˜„ëŒ€ì¹´ë“œ",prefix:"[í˜„ëŒ€ì¹´ë“œ]",sender:"15778000",emoji:"ðŸ–¤"},
  {id:"samsung",name:"ì‚¼ì„±ì¹´ë“œ",prefix:"[ì‚¼ì„±ì¹´ë“œ]",sender:"15881700",emoji:"ðŸ’™"},
  {id:"kb",     name:"KBêµ­ë¯¼ì¹´ë“œ",prefix:"[KBêµ­ë¯¼ì¹´ë“œ]",sender:"15881688",emoji:"ðŸ’›"},
  {id:"lotte",  name:"ë¡¯ë°ì¹´ë“œ",prefix:"[ë¡¯ë°ì¹´ë“œ]",sender:"15771700",emoji:"â¤ï¸"},
  {id:"hana",   name:"í•˜ë‚˜ì¹´ë“œ",prefix:"[í•˜ë‚˜ì¹´ë“œ]",sender:"15991111",emoji:"ðŸ’š"},
  {id:"kakao",  name:"ì¹´ì¹´ì˜¤íŽ˜ì´",prefix:"[ì¹´ì¹´ì˜¤íŽ˜ì´]",sender:"15994490",emoji:"ðŸ’›"},
];
const CATEGORIES = [
  {name:"ì‹ë¹„",icon:"ðŸš",color:"#F97316"},{name:"ì¹´íŽ˜",icon:"â˜•",color:"#A78BFA"},
  {name:"ì‡¼í•‘",icon:"ðŸ›ï¸",color:"#EC4899"},{name:"êµí†µ",icon:"ðŸšŒ",color:"#3B82F6"},
  {name:"íŽ¸ì˜ì ",icon:"ðŸª",color:"#10B981"},{name:"ì˜ë£Œ",icon:"ðŸ¥",color:"#EF4444"},
  {name:"ë¬¸í™”",icon:"ðŸŽ¬",color:"#F59E0B"},{name:"ë·°í‹°",icon:"ðŸŒ¿",color:"#06B6D4"},
  {name:"ìš´ë™",icon:"ðŸ’ª",color:"#84CC16"},{name:"ìˆ˜ìž…",icon:"ðŸ’°",color:"#10B981"},
  {name:"ê¸°íƒ€",icon:"ðŸ“¦",color:"#94A3B8"},
];
const WARN_THR=120,ALERT_THR=150;
const getLevel=(c,a)=>{const p=(c/a)*100;return p>=ALERT_THR?"alert":p>=WARN_THR?"warn":"ok";};
const LVL={
  alert:{badge:"#EF4444",badgeBg:"#FEE2E2",border:"#FECACA",bg:"#FEE2E2",text:"#991B1B",icon:"ðŸš¨",label:"ê²½ê³ "},
  warn: {badge:"#F97316",badgeBg:"#FFF7ED",border:"#FFEDD5",bg:"#FFF7ED",text:"#92400E",icon:"âš ï¸",label:"ì£¼ì˜"},
  ok:   {badge:"#10B981",badgeBg:"#DCFCE7",border:"#DCFCE7",bg:"#F0FDF4",text:"#14532D",icon:"âœ…",label:"ì •ìƒ"},
};
const fmt=(n)=>{const a=Math.abs(n).toLocaleString();return n<0?`-${a}ì›`:`+${a}ì›`;};
const parseMerchant=(raw)=>{const m=raw.match(/\d{2}:\d{2}\s(.+?)\s[\d,]+ì›/);return m?m[1]:"ì•Œ ìˆ˜ ì—†ìŒ";};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ê³µí†µ UI
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const Toggle=({value,onChange})=>(
  <div onClick={()=>onChange(!value)} style={{width:"44px",height:"24px",borderRadius:"99px",background:value?"#6366F1":"#CBD5E1",position:"relative",cursor:"pointer",transition:"background 0.2s",flexShrink:0}}>
    <div style={{width:"20px",height:"20px",borderRadius:"50%",background:"white",position:"absolute",top:"2px",left:value?"22px":"2px",transition:"left 0.2s",boxShadow:"0 1px 3px rgba(0,0,0,0.2)"}}/>
  </div>
);

const NavBar=({screen,setScreen,alertCount})=>(
  <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:"430px",background:"white",borderTop:"1px solid #F1F5F9",display:"flex",justifyContent:"space-around",padding:"8px 0 16px",boxShadow:"0 -4px 20px rgba(0,0,0,0.06)",zIndex:100}}>
    {[{id:"home",icon:"ðŸ ",label:"í™ˆ"},{id:"excel",icon:"ðŸ“¥",label:"ì—‘ì…€"},{id:"stats",icon:"ðŸ“Š",label:"í†µê³„"},{id:"alert",icon:"ðŸš¨",label:"ì•Œë¦¼",badge:alertCount},{id:"family",icon:"ðŸ‘¨â€ðŸ‘©â€ðŸ‘¦",label:"ê°€ì¡±"},{id:"settings",icon:"âš™ï¸",label:"ì„¤ì •"}].map(({id,icon,label,badge})=>(
      <button key={id} onClick={()=>setScreen(id)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"2px",padding:"6px 8px",background:"none",border:"none",cursor:"pointer",color:screen===id?"#6366F1":"#94A3B8",position:"relative"}}>
        <span style={{fontSize:"18px"}}>{icon}</span>
        <span style={{fontSize:"9px",fontWeight:screen===id?"700":"400"}}>{label}</span>
        {badge>0&&<div style={{position:"absolute",top:"2px",right:"2px",width:"14px",height:"14px",borderRadius:"50%",background:"#EF4444",color:"white",fontSize:"8px",fontWeight:"700",display:"flex",alignItems:"center",justifyContent:"center"}}>{badge}</div>}
      </button>
    ))}
  </div>
);

// ì°¨íŠ¸ë“¤
const BarChartH=({data})=>{
  const max=Math.max(...data.map(d=>d.amount));
  return <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>{data.map(d=>{const pct=Math.round((d.amount/max)*100);return(<div key={d.name}><div style={{display:"flex",justifyContent:"space-between",marginBottom:"5px"}}><div style={{display:"flex",alignItems:"center",gap:"6px"}}><span style={{fontSize:"14px"}}>{d.icon}</span><span style={{fontSize:"13px",color:"#334155"}}>{d.name}</span></div><span style={{fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>{d.amount.toLocaleString()}ì›</span></div><div style={{background:"#F1F5F9",borderRadius:"99px",height:"8px"}}><div style={{width:`${pct}%`,height:"100%",background:d.color,borderRadius:"99px"}}/></div></div>);})}</div>;
};

const TrendChart=({data,selectedIdx,onSelect})=>{
  const w=340,h=160,pad={t:10,b:30,l:10,r:10};
  const iW=w-pad.l-pad.r,iH=h-pad.t-pad.b;
  const max=Math.max(...data.map(d=>Math.max(d.income,d.expense)));
  const bW=iW/data.length*0.4;
  const xP=(i)=>pad.l+(i/(data.length-1))*iW;
  const yP=(v)=>pad.t+iH-(v/max)*iH;
  const iPts=data.map((d,i)=>`${xP(i)},${yP(d.income)}`).join(" ");
  const ePts=data.map((d,i)=>`${xP(i)},${yP(d.expense)}`).join(" ");
  return(
    <svg width="100%" viewBox={`0 0 ${w} ${h}`} style={{overflow:"visible"}}>
      {data.map((d,i)=>{const x=xP(i),bH=(d.expense/max)*iH,iS=selectedIdx===i;return(<g key={i} onClick={()=>onSelect(i)} style={{cursor:"pointer"}}><rect x={x-bW/2} y={pad.t+iH-bH} width={bW} height={bH} rx="4" fill={iS?"#6366F1":"#E0E7FF"} opacity={iS?1:0.7}/><text x={x} y={h-6} textAnchor="middle" fontSize="10" fill={iS?"#6366F1":"#94A3B8"} fontWeight={iS?"700":"400"}>{d.month}</text></g>);})}
      <polyline points={iPts} fill="none" stroke="#10B981" strokeWidth="2" strokeLinecap="round" strokeDasharray="4 2"/>
      {data.map((d,i)=><circle key={i} cx={xP(i)} cy={yP(d.income)} r="3" fill="#10B981"/>)}
      <polyline points={ePts} fill="none" stroke="#6366F1" strokeWidth="2.5" strokeLinecap="round"/>
      {data.map((d,i)=><circle key={i} cx={xP(i)} cy={yP(d.expense)} r={selectedIdx===i?5:3} fill={selectedIdx===i?"#6366F1":"white"} stroke="#6366F1" strokeWidth="2"/>)}
    </svg>
  );
};

const CatChart=({data,color})=>{
  const w=300,h=80,p={t:8,b:8,l:8,r:8};
  const iW=w-p.l-p.r,iH=h-p.t-p.b;
  const max=Math.max(...data),min=Math.min(...data),range=max-min||1;
  const xP=(i)=>p.l+(i/(data.length-1))*iW;
  const yP=(v)=>p.t+iH-((v-min)/range)*iH;
  const pts=data.map((v,i)=>`${xP(i)},${yP(v)}`).join(" ");
  const area=[`${p.l},${p.t+iH}`,...data.map((v,i)=>`${xP(i)},${yP(v)}`),`${p.l+iW},${p.t+iH}`].join(" ");
  return(
    <svg width="100%" viewBox={`0 0 ${w} ${h}`}>
      <defs><linearGradient id="cg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor={color} stopOpacity="0.3"/><stop offset="100%" stopColor={color} stopOpacity="0.02"/></linearGradient></defs>
      <polygon points={area} fill="url(#cg)"/>
      <polyline points={pts} fill="none" stroke={color} strokeWidth="2" strokeLinecap="round"/>
      {data.map((v,i)=><circle key={i} cx={xP(i)} cy={yP(v)} r={i===data.length-1?4:2.5} fill={i===data.length-1?color:"white"} stroke={color} strokeWidth="1.5"/>)}
    </svg>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// íŽ¸ì§‘ ëª¨ë‹¬
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const EditModal=({tx,onSave,onDelete,onClose})=>{
  const [name,setName]=useState(tx.name);
  const [amount,setAmount]=useState(String(Math.abs(tx.amount)));
  const [type,setType]=useState(tx.amount<0?"expense":"income");
  const [cat,setCat]=useState(tx.category);
  const [card,setCard]=useState(tx.card);
  const [memo,setMemo]=useState(tx.memo||"");
  const [date,setDate]=useState(tx.date.replaceAll(".","-"));
  const [confirmDel,setConfirmDel]=useState(false);

  const handleSave=()=>{
    const finalAmt=type==="expense"?-parseInt(amount||0):parseInt(amount||0);
    const catMeta=CATEGORIES.find(c=>c.name===cat)||CATEGORIES.at(-1);
    onSave({...tx,name,amount:finalAmt,category:cat,icon:catMeta.icon,card,memo,date:date.replaceAll("-",".")});
  };

  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{background:"white",borderRadius:"24px 24px 0 0",width:"100%",maxWidth:"430px",padding:"20px 20px 40px",maxHeight:"90vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"16px"}}>
          <h2 style={{margin:0,fontSize:"17px",fontWeight:"700",color:"#1E293B"}}>ë‚´ì—­ íŽ¸ì§‘</h2>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:"24px",color:"#94A3B8",lineHeight:1}}>Ã—</button>
        </div>
        <div style={{display:"flex",background:"#F1F5F9",borderRadius:"12px",padding:"4px",marginBottom:"14px"}}>
          {[["expense","ì§€ì¶œ"],["income","ìˆ˜ìž…"]].map(([k,l])=>(
            <button key={k} onClick={()=>setType(k)} style={{flex:1,padding:"8px",border:"none",cursor:"pointer",borderRadius:"10px",background:type===k?"white":"transparent",color:type===k?(k==="expense"?"#6366F1":"#10B981"):"#64748B",fontWeight:type===k?"700":"400",fontSize:"14px",transition:"all 0.2s"}}>{l}</button>
          ))}
        </div>
        {[
          {label:"ìƒí˜¸ëª…",content:<input value={name} onChange={e=>setName(e.target.value)} style={{width:"100%",padding:"11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"14px",color:"#1E293B",outline:"none",boxSizing:"border-box"}}/>},
          {label:"ê¸ˆì•¡",content:<div style={{position:"relative"}}><input type="number" value={amount} onChange={e=>setAmount(e.target.value)} style={{width:"100%",padding:"11px 44px 11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"16px",fontWeight:"700",color:type==="expense"?"#6366F1":"#10B981",outline:"none",boxSizing:"border-box"}}/><span style={{position:"absolute",right:"14px",top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}>ì›</span></div>},
          {label:"ë‚ ì§œ",content:<input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{width:"100%",padding:"11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"14px",color:"#6366F1",fontWeight:"600",outline:"none",boxSizing:"border-box"}}/>},
        ].map(({label,content})=>(
          <div key={label} style={{marginBottom:"12px"}}>
            <p style={{margin:"0 0 6px",fontSize:"12px",fontWeight:"600",color:"#64748B"}}>{label}</p>
            {content}
          </div>
        ))}
        <div style={{marginBottom:"12px"}}>
          <p style={{margin:"0 0 8px",fontSize:"12px",fontWeight:"600",color:"#64748B"}}>ì¹´í…Œê³ ë¦¬</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"6px"}}>
            {CATEGORIES.map(c=>{const iS=cat===c.name;return(<button key={c.name} onClick={()=>setCat(c.name)} style={{background:iS?c.color+"18":"#F8FAFC",border:`2px solid ${iS?c.color:"transparent"}`,borderRadius:"10px",padding:"8px 4px",cursor:"pointer",textAlign:"center"}}><div style={{fontSize:"18px",marginBottom:"2px"}}>{c.icon}</div><p style={{margin:0,fontSize:"10px",color:iS?c.color:"#64748B",fontWeight:iS?"700":"400"}}>{c.name}</p></button>);})}</div>
        </div>
        {[
          {label:"ê²°ì œìˆ˜ë‹¨",content:<input value={card} onChange={e=>setCard(e.target.value)} placeholder="ì‹ í•œì¹´ë“œ, í˜„ê¸ˆ ë“±" style={{width:"100%",padding:"11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"14px",outline:"none",boxSizing:"border-box"}}/>},
          {label:"ë©”ëª¨",content:<textarea value={memo} onChange={e=>setMemo(e.target.value)} rows={2} placeholder="ë©”ëª¨ (ì„ íƒ)" style={{width:"100%",padding:"11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"14px",outline:"none",resize:"none",boxSizing:"border-box",fontFamily:"inherit"}}/>},
        ].map(({label,content})=>(
          <div key={label} style={{marginBottom:"12px"}}>
            <p style={{margin:"0 0 6px",fontSize:"12px",fontWeight:"600",color:"#64748B"}}>{label}</p>
            {content}
          </div>
        ))}
        <div style={{display:"flex",gap:"10px",marginTop:"8px"}}>
          {!confirmDel?(
            <>
              <button onClick={()=>setConfirmDel(true)} style={{padding:"14px 18px",border:"2px solid #FEE2E2",borderRadius:"14px",background:"white",color:"#EF4444",fontSize:"14px",fontWeight:"700",cursor:"pointer"}}>ðŸ—‘ï¸</button>
              <button onClick={handleSave} style={{flex:1,padding:"14px",border:"none",borderRadius:"14px",background:"linear-gradient(135deg,#6366F1,#8B5CF6)",color:"white",fontSize:"15px",fontWeight:"700",cursor:"pointer"}}>ì €ìž¥í•˜ê¸°</button>
            </>
          ):(
            <>
              <button onClick={()=>setConfirmDel(false)} style={{flex:1,padding:"14px",border:"2px solid #E2E8F0",borderRadius:"14px",background:"white",color:"#64748B",fontSize:"14px",fontWeight:"700",cursor:"pointer"}}>ì·¨ì†Œ</button>
              <button onClick={()=>onDelete(tx.id)} style={{flex:1,padding:"14px",border:"none",borderRadius:"14px",background:"#EF4444",color:"white",fontSize:"14px",fontWeight:"700",cursor:"pointer"}}>ì •ë§ ì‚­ì œ</button>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

// ì´ˆëŒ€ ëª¨ë‹¬
const InviteModal=({code,onClose})=>{
  const [copied,setCopied]=useState(false);
  const [joinCode,setJoinCode]=useState("");
  const [joined,setJoined]=useState(false);
  const [tab,setTab]=useState("invite");
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:"20px"}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{background:"white",borderRadius:"20px",width:"100%",maxWidth:"390px",padding:"24px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"16px"}}>
          <h2 style={{margin:0,fontSize:"17px",fontWeight:"700"}}>ê°€ì¡± ì´ˆëŒ€</h2>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:"24px",color:"#94A3B8",lineHeight:1}}>Ã—</button>
        </div>
        <div style={{display:"flex",background:"#F1F5F9",borderRadius:"10px",padding:"3px",marginBottom:"16px"}}>
          {[["invite","ì´ˆëŒ€í•˜ê¸°"],["join","ì½”ë“œ ìž…ë ¥"]].map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"7px",border:"none",cursor:"pointer",borderRadius:"8px",background:tab===k?"white":"transparent",color:tab===k?"#6366F1":"#64748B",fontWeight:tab===k?"700":"400",fontSize:"13px"}}>{l}</button>
          ))}
        </div>
        {tab==="invite"&&(
          <>
            <div style={{background:"linear-gradient(135deg,#EEF2FF,#F5F3FF)",border:"2px dashed #A5B4FC",borderRadius:"16px",padding:"20px",textAlign:"center",marginBottom:"14px"}}>
              <p style={{margin:"0 0 6px",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>ì´ˆëŒ€ ì½”ë“œ</p>
              <p style={{margin:0,fontSize:"28px",fontWeight:"700",color:"#4338CA",letterSpacing:"4px"}}>{code}</p>
            </div>
            <div style={{display:"flex",gap:"8px",marginBottom:"12px"}}>
              <button onClick={()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);}} style={{flex:1,padding:"12px",border:"none",borderRadius:"12px",background:copied?"#10B981":"#6366F1",color:"white",fontSize:"13px",fontWeight:"700",cursor:"pointer",transition:"background 0.2s"}}>{copied?"âœ… ë³µì‚¬ëì–´ìš”":"ðŸ“‹ ì½”ë“œ ë³µì‚¬"}</button>
              <button style={{flex:1,padding:"12px",border:"none",borderRadius:"12px",background:"#F9E000",color:"#1A1A1A",fontSize:"13px",fontWeight:"700",cursor:"pointer"}}>ðŸ’¬ ì¹´í†¡ ê³µìœ </button>
            </div>
            <p style={{margin:0,fontSize:"12px",color:"#94A3B8",textAlign:"center"}}>ì½”ë“œ ìœ íš¨ê¸°ê°„ 7ì¼ Â· ê°€ì¡±ì´ ìž…ë ¥í•˜ë©´ ì¦‰ì‹œ ì—°ê²°</p>
          </>
        )}
        {tab==="join"&&(
          !joined?(
            <>
              <input value={joinCode} onChange={e=>setJoinCode(e.target.value.toUpperCase())} placeholder="MONEY-XXXX" maxLength={10} style={{width:"100%",padding:"14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"18px",fontWeight:"700",color:"#4338CA",textAlign:"center",letterSpacing:"3px",outline:"none",boxSizing:"border-box",marginBottom:"12px"}}/>
              <button onClick={()=>{if(joinCode.length>=5)setJoined(true);}} style={{width:"100%",padding:"14px",border:"none",borderRadius:"12px",background:joinCode.length>=5?"linear-gradient(135deg,#6366F1,#8B5CF6)":"#CBD5E1",color:"white",fontSize:"15px",fontWeight:"700",cursor:joinCode.length>=5?"pointer":"not-allowed"}}>ì°¸ì—¬í•˜ê¸°</button>
            </>
          ):(
            <div style={{textAlign:"center",padding:"20px 0"}}>
              <div style={{fontSize:"48px",marginBottom:"12px"}}>ðŸŽ‰</div>
              <p style={{fontSize:"16px",fontWeight:"700",color:"#1E293B",margin:"0 0 6px"}}>ì—°ê²°ëì–´ìš”!</p>
              <p style={{fontSize:"13px",color:"#64748B",margin:"0 0 16px"}}>ì´ì œ ê°€ì¡±ê³¼ í•¨ê»˜ ì‚¬ìš©í•  ìˆ˜ ìžˆì–´ìš”</p>
              <button onClick={onClose} style={{padding:"12px 24px",border:"none",borderRadius:"12px",background:"linear-gradient(135deg,#6366F1,#8B5CF6)",color:"white",fontSize:"14px",fontWeight:"700",cursor:"pointer"}}>í™•ì¸</button>
            </div>
          )
        )}
      </div>
    </div>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// í™”ë©´ 1: í™ˆ
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const HomeScreen=({setScreen,txList})=>{
  const [tab,setTab]=useState("week");
  const [aiAnalysis,setAiAnalysis]=useState(null);
  const [aiLoading,setAiLoading]=useState(false);
  const totalMonthly=MONTHLY_CARDS.reduce((s,c)=>s+c.amount,0);
  const totalWeekly=WEEKLY_CARDS.reduce((s,c)=>s+c.amount,0);
  const totalCat=MONTHLY_CATS.reduce((s,c)=>s+c.amount,0);

  const handleAI=async()=>{
    setAiLoading(true);setAiAnalysis(null);
    try{
      const summary=MONTHLY_CATS.map(c=>`${c.name}:${c.amount.toLocaleString()}ì›`).join(", ");
      const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1000,messages:[{role:"user",content:`ì¹œê·¼í•œ ê°€ê³„ë¶€ AIì•¼. ì´ë²ˆë‹¬ ì§€ì¶œ: ${summary}. ì´ëª¨ì§€ ì¨ì„œ 3ì¤„ë¡œ ë¶„ì„í•´ì¤˜.\n1. í•œì¤„ ìš”ì•½\n2. ì£¼ëª©í•  ì \n3. ì ˆì•½ íŒ`}]})});
      const data=await res.json();
      setAiAnalysis(data.content?.[0]?.text||"ë¶„ì„ ì‹¤íŒ¨");
    }catch{setAiAnalysis("ë¶„ì„ ì¤‘ ì˜¤ë¥˜ê°€ ë°œìƒí–ˆì–´ìš”.");}
    setAiLoading(false);
  };

  const recent=txList.slice(0,5);

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 28px",color:"white"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"20px"}}>
          <div><p style={{fontSize:"13px",opacity:0.8,margin:0}}>2026ë…„ 6ì›”</p><h1 style={{fontSize:"22px",fontWeight:"700",margin:"4px 0 0"}}>ì•ˆë…•í•˜ì„¸ìš” ðŸ‘‹</h1></div>
          <button onClick={()=>setScreen("excel")} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:"12px",padding:"8px 12px",color:"white",fontSize:"12px",fontWeight:"600",cursor:"pointer"}}>ðŸ“¥ ì—‘ì…€ ë¶ˆëŸ¬ì˜¤ê¸°</button>
        </div>
        <div style={{background:"rgba(255,255,255,0.15)",borderRadius:"16px",padding:"16px 20px"}}>
          <p style={{margin:0,fontSize:"12px",opacity:0.8}}>ì´ë²ˆ ë‹¬ ì´ ì§€ì¶œ</p>
          <p style={{margin:"4px 0 0",fontSize:"28px",fontWeight:"700"}}>{totalMonthly.toLocaleString()}ì›</p>
        </div>
      </div>
      <div style={{padding:"20px 16px 100px"}}>
        {/* AI ë¶„ì„ */}
        <div style={{background:aiAnalysis?"linear-gradient(135deg,#EEF2FF,#F5F3FF)":"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",border:aiAnalysis?"1px solid #C7D2FE":"1px solid #F1F5F9",marginBottom:"20px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:aiAnalysis?"12px":"0"}}>
            <div style={{display:"flex",alignItems:"center",gap:"8px"}}><span style={{fontSize:"18px"}}>ðŸ¤–</span><span style={{fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>AI ì§€ì¶œ ë¶„ì„</span></div>
            <button onClick={handleAI} disabled={aiLoading} style={{background:aiLoading?"#E2E8F0":"linear-gradient(135deg,#6366F1,#8B5CF6)",border:"none",borderRadius:"10px",padding:"6px 12px",color:aiLoading?"#94A3B8":"white",fontSize:"12px",fontWeight:"600",cursor:aiLoading?"not-allowed":"pointer"}}>{aiLoading?"ë¶„ì„ ì¤‘...":aiAnalysis?"ë‹¤ì‹œ ë¶„ì„":"ë¶„ì„í•˜ê¸°"}</button>
          </div>
          {aiLoading&&<p style={{margin:"8px 0 0",fontSize:"12px",color:"#94A3B8"}}>Claudeê°€ ë¶„ì„í•˜ê³  ìžˆì–´ìš”...</p>}
          {aiAnalysis&&!aiLoading&&<div style={{background:"white",borderRadius:"12px",padding:"12px 14px",fontSize:"13px",color:"#334155",lineHeight:"1.8",whiteSpace:"pre-line"}}>{aiAnalysis}</div>}
          {!aiAnalysis&&!aiLoading&&<p style={{margin:"8px 0 0",fontSize:"12px",color:"#94A3B8"}}>ì´ë²ˆ ë‹¬ ì§€ì¶œ íŒ¨í„´ì„ AIê°€ ë¶„ì„í•´ë“œë ¤ìš”</p>}
        </div>

        {/* íƒ­ */}
        <div style={{display:"flex",background:"#E2E8F0",borderRadius:"12px",padding:"4px",marginBottom:"20px"}}>
          {[["week","ì´ë²ˆ ì£¼"],["month","ì´ë²ˆ ë‹¬"]].map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"8px",border:"none",cursor:"pointer",borderRadius:"10px",background:tab===k?"white":"transparent",color:tab===k?"#6366F1":"#64748B",fontWeight:tab===k?"700":"400",fontSize:"14px",boxShadow:tab===k?"0 1px 4px rgba(0,0,0,0.1)":"none",transition:"all 0.2s"}}>{l}</button>
          ))}
        </div>

        {/* ì¹´ë“œë³„ ì§€ì¶œ */}
        <div style={{marginBottom:"24px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
            <h2 style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>ì¹´ë“œë³„ ì§€ì¶œ</h2>
            <span style={{fontSize:"12px",color:"#94A3B8"}}>í•©ê³„ {(tab==="week"?totalWeekly:totalMonthly).toLocaleString()}ì›</span>
          </div>
          {tab==="week"?(
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {WEEKLY_CARDS.map(c=>(
                <div key={c.name} style={{background:"white",borderRadius:"14px",padding:"14px 16px",display:"flex",alignItems:"center",gap:"12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <div style={{width:"40px",height:"40px",borderRadius:"10px",background:c.color+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"18px"}}>{c.emoji}</div>
                  <div style={{flex:1}}><p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{c.name}</p><p style={{margin:"2px 0 0",fontSize:"12px",color:"#94A3B8"}}>ì´ë²ˆ ì£¼</p></div>
                  <p style={{margin:0,fontSize:"15px",fontWeight:"700",color:c.color}}>{c.amount.toLocaleString()}ì›</p>
                </div>
              ))}
            </div>
          ):(
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {MONTHLY_CARDS.map(c=>{const pct=Math.round((c.amount/c.limit)*100);return(
                <div key={c.name} style={{background:"white",borderRadius:"14px",padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:"8px"}}><p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{c.name}</p><p style={{margin:0,fontSize:"14px",fontWeight:"700",color:c.color}}>{c.amount.toLocaleString()}ì›</p></div>
                  <div style={{background:"#F1F5F9",borderRadius:"99px",height:"6px"}}><div style={{width:`${pct}%`,height:"100%",background:pct>80?"#EF4444":c.color,borderRadius:"99px"}}/></div>
                  <div style={{display:"flex",justifyContent:"space-between",marginTop:"4px"}}><p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>í•œë„ {c.limit.toLocaleString()}ì›</p><p style={{margin:0,fontSize:"11px",color:pct>80?"#EF4444":"#94A3B8"}}>{pct}%</p></div>
                </div>
              );})}
            </div>
          )}
        </div>

        {/* ì¹´í…Œê³ ë¦¬ ë°” ì°¨íŠ¸ */}
        <div style={{marginBottom:"24px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
            <h2 style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>ì›”ë³„ ì¹´í…Œê³ ë¦¬ ì§€ì¶œ</h2>
            <span style={{fontSize:"12px",color:"#94A3B8"}}>ì´ {totalCat.toLocaleString()}ì›</span>
          </div>
          <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"10px"}}><BarChartH data={MONTHLY_CATS}/></div>
          <div style={{background:"#FFF7ED",borderRadius:"12px",padding:"10px 14px",display:"flex",alignItems:"center",gap:"8px",border:"1px solid #FFEDD5"}}>
            <span style={{fontSize:"16px"}}>ðŸ”¥</span>
            <p style={{margin:0,fontSize:"12px",color:"#92400E"}}>ì´ë²ˆ ë‹¬ <strong>{MONTHLY_CATS[0].name}</strong>ì— ê°€ìž¥ ë§Žì´ ì¼ì–´ìš” â€” <strong>{MONTHLY_CATS[0].amount.toLocaleString()}ì›</strong></p>
          </div>
        </div>

        {/* ìµœê·¼ ë‚´ì—­ */}
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
            <h2 style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>ìµœê·¼ ë‚´ì—­</h2>
            <button onClick={()=>setScreen("txlist")} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>ì „ì²´ë³´ê¸° â†’</button>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
            {recent.map(tx=>(
              <div key={tx.id} style={{background:"white",borderRadius:"14px",padding:"12px 16px",display:"flex",alignItems:"center",gap:"12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                <div style={{width:"38px",height:"38px",borderRadius:"10px",background:tx.amount>0?"#DCFCE7":"#F1F5F9",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"18px",flexShrink:0}}>{tx.icon}</div>
                <div style={{flex:1,minWidth:0}}>
                  <p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{tx.name}</p>
                  <p style={{margin:"2px 0 0",fontSize:"11px",color:"#94A3B8"}}>{tx.category} Â· {tx.date}</p>
                </div>
                <div style={{textAlign:"right"}}>
                  <p style={{margin:0,fontSize:"14px",fontWeight:"700",color:tx.amount>0?"#10B981":"#1E293B"}}>{fmt(tx.amount)}</p>
                  {tx.card!=="-"&&<p style={{margin:"2px 0 0",fontSize:"10px",color:"#CBD5E1"}}>{tx.card}</p>}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// í™”ë©´ 2: ì „ì²´ ë‚´ì—­ (ìˆ˜ì • í¬í•¨)
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const TxListScreen=({txList,setTxList,onBack})=>{
  const [editTx,setEditTx]=useState(null);
  const [filterCat,setFilterCat]=useState(null);
  const [search,setSearch]=useState("");

  const handleSave=(updated)=>{setTxList(p=>p.map(t=>t.id===updated.id?updated:t));setEditTx(null);};
  const handleDelete=(id)=>{setTxList(p=>p.filter(t=>t.id!==id));setEditTx(null);};

  const filtered=txList
    .filter(t=>!filterCat||t.category===filterCat)
    .filter(t=>!search||t.name.includes(search)||t.category.includes(search)||t.memo.includes(search));

  const totalExp=filtered.filter(t=>t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);
  const totalInc=filtered.filter(t=>t.amount>0).reduce((s,t)=>s+t.amount,0);

  // ë‚ ì§œë³„ ê·¸ë£¹
  const grouped=filtered.reduce((acc,tx)=>{
    if(!acc[tx.date])acc[tx.date]=[];
    acc[tx.date].push(tx);
    return acc;
  },{});

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 20px",color:"white"}}>
        <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"16px"}}>
          <button onClick={onBack} style={{background:"none",border:"none",color:"white",fontSize:"22px",cursor:"pointer",padding:0}}>â†</button>
          <h1 style={{margin:0,fontSize:"18px",fontWeight:"700"}}>ì „ì²´ ë‚´ì—­</h1>
        </div>
        {/* ê²€ìƒ‰ */}
        <div style={{position:"relative"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="ìƒí˜¸ëª…, ì¹´í…Œê³ ë¦¬, ë©”ëª¨ ê²€ìƒ‰" style={{width:"100%",padding:"10px 14px 10px 36px",border:"none",borderRadius:"12px",fontSize:"13px",background:"rgba(255,255,255,0.2)",color:"white",outline:"none",boxSizing:"border-box"}}/>
          <span style={{position:"absolute",left:"12px",top:"50%",transform:"translateY(-50%)",fontSize:"14px",opacity:0.7}}>ðŸ”</span>
        </div>
      </div>

      <div style={{padding:"0 0 100px"}}>
        {/* ìš”ì•½ */}
        <div style={{display:"flex",gap:"0",background:"white",borderBottom:"1px solid #F1F5F9"}}>
          <div style={{flex:1,padding:"14px 16px",textAlign:"center",borderRight:"1px solid #F1F5F9"}}>
            <p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>ì´ ì§€ì¶œ</p>
            <p style={{margin:"4px 0 0",fontSize:"16px",fontWeight:"700",color:"#6366F1"}}>{totalExp.toLocaleString()}ì›</p>
          </div>
          <div style={{flex:1,padding:"14px 16px",textAlign:"center"}}>
            <p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>ì´ ìˆ˜ìž…</p>
            <p style={{margin:"4px 0 0",fontSize:"16px",fontWeight:"700",color:"#10B981"}}>{totalInc.toLocaleString()}ì›</p>
          </div>
        </div>

        {/* ì¹´í…Œê³ ë¦¬ í•„í„° */}
        <div style={{display:"flex",gap:"6px",overflowX:"auto",padding:"12px 16px",background:"white",borderBottom:"1px solid #F1F5F9"}}>
          <button onClick={()=>setFilterCat(null)} style={{padding:"5px 12px",border:"none",borderRadius:"99px",background:!filterCat?"#6366F1":"#F1F5F9",color:!filterCat?"white":"#64748B",fontWeight:!filterCat?"700":"400",fontSize:"12px",flexShrink:0,cursor:"pointer"}}>ì „ì²´</button>
          {CATEGORIES.map(c=>(
            <button key={c.name} onClick={()=>setFilterCat(filterCat===c.name?null:c.name)} style={{display:"flex",alignItems:"center",gap:"4px",padding:"5px 10px",border:"none",borderRadius:"99px",background:filterCat===c.name?c.color:"#F1F5F9",color:filterCat===c.name?"white":"#64748B",fontWeight:filterCat===c.name?"700":"400",fontSize:"12px",flexShrink:0,cursor:"pointer"}}>
              <span>{c.icon}</span><span>{c.name}</span>
            </button>
          ))}
        </div>

        {/* ë‚ ì§œë³„ ë‚´ì—­ */}
        <div style={{padding:"0 16px"}}>
          {Object.keys(grouped).sort((a,b)=>b.localeCompare(a)).map(date=>{
            const dayTxs=grouped[date];
            const dayTotal=dayTxs.reduce((s,t)=>s+t.amount,0);
            return(
              <div key={date} style={{marginTop:"16px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}>
                  <span style={{fontSize:"12px",fontWeight:"700",color:"#64748B"}}>{date}</span>
                  <span style={{fontSize:"12px",fontWeight:"700",color:dayTotal<0?"#EF4444":"#10B981"}}>{fmt(dayTotal)}</span>
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:"6px"}}>
                  {dayTxs.map(tx=>(
                    <button key={tx.id} onClick={()=>setEditTx(tx)} style={{background:"white",border:"none",borderRadius:"14px",padding:"12px 14px",display:"flex",alignItems:"center",gap:"12px",boxShadow:"0 1px 4px rgba(0,0,0,0.06)",cursor:"pointer",textAlign:"left",width:"100%",transition:"transform 0.1s"}}
                      onMouseDown={e=>e.currentTarget.style.transform="scale(0.98)"}
                      onMouseUp={e=>e.currentTarget.style.transform="scale(1)"}>
                      <div style={{width:"38px",height:"38px",borderRadius:"10px",background:tx.amount>0?"#DCFCE7":"#F1F5F9",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"18px",flexShrink:0}}>{tx.icon}</div>
                      <div style={{flex:1,minWidth:0}}>
                        <p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{tx.name}</p>
                        <div style={{display:"flex",gap:"6px",marginTop:"2px"}}>
                          <span style={{fontSize:"10px",color:"#94A3B8"}}>{tx.category}</span>
                          {tx.time&&<span style={{fontSize:"10px",color:"#CBD5E1"}}>{tx.time}</span>}
                          {tx.memo&&<span style={{fontSize:"10px",color:"#94A3B8"}}>Â· {tx.memo}</span>}
                        </div>
                      </div>
                      <div style={{textAlign:"right"}}>
                        <p style={{margin:0,fontSize:"14px",fontWeight:"700",color:tx.amount>0?"#10B981":"#1E293B"}}>{fmt(tx.amount)}</p>
                        {tx.card!=="-"&&<p style={{margin:"2px 0 0",fontSize:"10px",color:"#CBD5E1"}}>{tx.card}</p>}
                      </div>
                      <span style={{fontSize:"12px",color:"#E2E8F0"}}>âœï¸</span>
                    </button>
                  ))}
                </div>
              </div>
            );
          })}
          {filtered.length===0&&(
            <div style={{textAlign:"center",padding:"60px 0",color:"#94A3B8"}}>
              <div style={{fontSize:"40px",marginBottom:"8px"}}>ðŸ“­</div>
              <p style={{margin:0}}>ë‚´ì—­ì´ ì—†ì–´ìš”</p>
            </div>
          )}
        </div>
      </div>

      {editTx&&<EditModal tx={editTx} onSave={handleSave} onDelete={handleDelete} onClose={()=>setEditTx(null)}/>}
    </div>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// í™”ë©´ 3: SMS ë¶ˆëŸ¬ì˜¤ê¸°
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const SmsScreen=({setTxList})=>{
  const [step,setStep]=useState("idle");
  const [items,setItems]=useState([]);
  const [selected,setSelected]=useState([]);
  const [done,setDone]=useState(false);

  const handleImport=async()=>{
    setStep("parsing");setItems([]);setSelected([]);
    await new Promise(r=>setTimeout(r,1200));
    const parsed=SMS_RAW.map(s=>({...s,merchant:parseMerchant(s),category:null,icon:"ðŸ“¦"}));
    setStep("classifying");
    try{
      const merchants=parsed.map(p=>p.merchant);
      const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1000,messages:[{role:"user",content:`ìƒí˜¸ëª…ìœ¼ë¡œ ì¹´í…Œê³ ë¦¬ ë¶„ë¥˜. ì¹´í…Œê³ ë¦¬: ì‹ë¹„,ì¹´íŽ˜,ì‡¼í•‘,êµí†µ,íŽ¸ì˜ì ,ì˜ë£Œ,ë¬¸í™”,ë·°í‹°,ìš´ë™,ê¸°íƒ€\nJSON ë°°ì—´ë¡œë§Œ. ì˜ˆ:["ì¹´íŽ˜"]\nìƒí˜¸ëª…:\n${merchants.map((m,i)=>`${i+1}. ${m}`).join("\n")}`}]})});
      const data=await res.json();
      const cats=JSON.parse((data.content?.[0]?.text||"[]").replace(/```json|```/g,"").trim());
      const enriched=parsed.map((p,i)=>{
        const cat=cats[i]||"ê¸°íƒ€";
        const meta=CATEGORIES.find(c=>c.name===cat)||CATEGORIES.at(-1);
        return{...p,category:cat,icon:meta.icon,color:meta.color};
      });
      setItems(enriched);setSelected(enriched.map(e=>e.id));
    }catch{
      const fb=parsed.map(p=>({...p,category:"ê¸°íƒ€",icon:"ðŸ“¦",color:"#94A3B8"}));
      setItems(fb);setSelected(fb.map(e=>e.id));
    }
    setStep("done");
  };

  const toggle=(id)=>setSelected(p=>p.includes(id)?p.filter(i=>i!==id):[...p,id]);

  const handleConfirm=()=>{
    const toAdd=items.filter(it=>selected.includes(it.id)).map(it=>({
      id:Date.now()+it.id,name:it.merchant,category:it.category,
      amount:-it.amount,date:it.date.replace(/(\d{2})\/(\d{2})/,"2026.$1.$2"),
      time:it.date.split(" ")[1]||"",card:it.card,icon:it.icon||"ðŸ“¦",author:1,memo:"",
    }));
    setTxList(p=>[...toAdd,...p]);
    setDone(true);
    setTimeout(()=>{setDone(false);setStep("idle");setItems([]);setSelected([]);},1500);
  };

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 24px",color:"white"}}>
        <h1 style={{margin:"0 0 4px",fontSize:"20px",fontWeight:"700"}}>ë¬¸ìžì—ì„œ ë¶ˆëŸ¬ì˜¤ê¸° ðŸ“©</h1>
        <p style={{margin:0,fontSize:"13px",opacity:0.75}}>AIê°€ ì¹´í…Œê³ ë¦¬ë¥¼ ìžë™ìœ¼ë¡œ ë¶„ë¥˜í•´ìš”</p>
      </div>
      <div style={{padding:"20px 16px 100px"}}>
        {step==="idle"&&(
          <div style={{textAlign:"center",padding:"40px 0"}}>
            <div style={{fontSize:"56px",marginBottom:"16px"}}>ðŸ“±</div>
            <p style={{color:"#1E293B",fontSize:"16px",fontWeight:"700",marginBottom:"8px"}}>ì¹´ë“œ ê²°ì œ ë¬¸ìž ë¶ˆëŸ¬ì˜¤ê¸°</p>
            <p style={{color:"#94A3B8",fontSize:"13px",marginBottom:"32px",lineHeight:"1.6"}}>ìµœê·¼ ìˆ˜ì‹ ëœ ì¹´ë“œì‚¬ ê²°ì œ ë¬¸ìžë¥¼ ì½ì–´ì™€<br/>AIê°€ ì¹´í…Œê³ ë¦¬ë¥¼ ìžë™ìœ¼ë¡œ ë¶„ë¥˜í•´ë“œë ¤ìš”</p>
            <button onClick={handleImport} style={{background:"linear-gradient(135deg,#6366F1,#8B5CF6)",border:"none",borderRadius:"14px",padding:"14px 32px",color:"white",fontSize:"15px",fontWeight:"700",cursor:"pointer"}}>ðŸ“© ë¬¸ìž ë¶ˆëŸ¬ì˜¤ê¸° ì‹œìž‘</button>
          </div>
        )}
        {(step==="parsing"||step==="classifying")&&(
          <div style={{textAlign:"center",padding:"60px 0"}}>
            <div style={{fontSize:"40px",marginBottom:"20px"}}>{step==="parsing"?"ðŸ“±":"ðŸ¤–"}</div>
            <p style={{color:"#1E293B",fontSize:"15px",fontWeight:"700",marginBottom:"8px"}}>{step==="parsing"?"ë¬¸ìž ë‚´ì—­ ì½ëŠ” ì¤‘...":"AIê°€ ì¹´í…Œê³ ë¦¬ ë¶„ë¥˜ ì¤‘..."}</p>
            <p style={{color:"#94A3B8",fontSize:"12px"}}>{step==="classifying"?"Claudeê°€ ìƒí˜¸ëª…ì„ ë¶„ì„í•˜ê³  ìžˆì–´ìš”":"ìž ì‹œë§Œ ê¸°ë‹¤ë ¤ì£¼ì„¸ìš”"}</p>
          </div>
        )}
        {step==="done"&&!done&&(
          <>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
              <p style={{margin:0,fontSize:"13px",color:"#64748B"}}>{items.length}ê±´ ì¸ì‹ Â· AI ì¹´í…Œê³ ë¦¬ ë¶„ë¥˜ ì™„ë£Œ âœ¨</p>
              <button onClick={()=>setSelected(selected.length===items.length?[]:items.map(s=>s.id))} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>{selected.length===items.length?"ì „ì²´í•´ì œ":"ì „ì²´ì„ íƒ"}</button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {items.map(item=>{const iS=selected.includes(item.id);return(
                <button key={item.id} onClick={()=>toggle(item.id)} style={{background:iS?"#EEF2FF":"white",border:`2px solid ${iS?"#6366F1":"transparent"}`,borderRadius:"14px",padding:"14px 16px",display:"flex",alignItems:"center",gap:"12px",cursor:"pointer",textAlign:"left",width:"100%",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <div style={{width:"22px",height:"22px",borderRadius:"6px",background:iS?"#6366F1":"white",border:`2px solid ${iS?"#6366F1":"#CBD5E1"}`,display:"flex",alignItems:"center",justifyContent:"center",color:"white",fontSize:"13px",fontWeight:"700",flexShrink:0}}>{iS?"âœ“":""}</div>
                  <div style={{width:"40px",height:"40px",borderRadius:"10px",background:(item.color||"#94A3B8")+"20",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"20px",flexShrink:0}}>{item.icon}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",justifyContent:"space-between"}}>
                      <div>
                        <p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{item.merchant}</p>
                        <div style={{display:"flex",alignItems:"center",gap:"6px",marginTop:"3px"}}>
                          <span style={{fontSize:"10px",padding:"2px 6px",borderRadius:"99px",background:(item.color||"#94A3B8")+"20",color:item.color||"#94A3B8",fontWeight:"700"}}>ðŸ¤– {item.category}</span>
                          <span style={{fontSize:"11px",color:"#94A3B8"}}>{item.card} Â· {item.date}</span>
                        </div>
                      </div>
                      <p style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>{item.amount.toLocaleString()}ì›</p>
                    </div>
                    <div style={{marginTop:"8px",background:"#F8FAFC",borderRadius:"8px",padding:"6px 8px"}}><p style={{margin:0,fontSize:"10px",color:"#94A3B8",fontFamily:"monospace"}}>{item.raw}</p></div>
                  </div>
                </button>
              );})}
            </div>
          </>
        )}
        {done&&<div style={{textAlign:"center",padding:"60px 0"}}><div style={{fontSize:"50px",marginBottom:"16px"}}>âœ…</div><p style={{color:"#10B981",fontSize:"16px",fontWeight:"700"}}>{selected.length}ê±´ ê°€ê³„ë¶€ì— ì¶”ê°€ëì–´ìš”!</p></div>}
      </div>
      {step==="done"&&!done&&(
        <div style={{position:"fixed",bottom:"72px",left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:"430px",background:"white",borderTop:"1px solid #F1F5F9",padding:"12px 16px 8px",boxShadow:"0 -4px 20px rgba(0,0,0,0.08)"}}>
          <button onClick={handleConfirm} disabled={selected.length===0} style={{width:"100%",padding:"15px",border:"none",borderRadius:"14px",background:selected.length>0?"linear-gradient(135deg,#6366F1,#8B5CF6)":"#CBD5E1",color:"white",fontSize:"16px",fontWeight:"700",cursor:selected.length>0?"pointer":"not-allowed"}}>{selected.length>0?`${selected.length}ê±´ ê°€ê³„ë¶€ì— ì¶”ê°€`:"í•­ëª©ì„ ì„ íƒí•´ì£¼ì„¸ìš”"}</button>
        </div>
      )}
    </div>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// í™”ë©´ 4: í†µê³„
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const StatsScreen=()=>{
  const [tab,setTab]=useState("trend");
  const [selM,setSelM]=useState(5);
  const [selC,setSelC]=useState("ì‹ë¹„");
  const months=["1ì›”","2ì›”","3ì›”","4ì›”","5ì›”","6ì›”"];
  const sel=MONTHLY_TREND[selM];
  const saving=sel.income-sel.expense;
  const sr=Math.round((saving/sel.income)*100);
  const avgExp=Math.round(MONTHLY_TREND.reduce((s,m)=>s+m.expense,0)/MONTHLY_TREND.length);
  const maxM=[...MONTHLY_TREND].sort((a,b)=>b.expense-a.expense)[0];

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 24px",color:"white"}}>
        <p style={{margin:0,fontSize:"13px",opacity:0.8}}>2026ë…„ ìƒë°˜ê¸°</p>
        <h1 style={{fontSize:"20px",fontWeight:"700",margin:"4px 0 16px"}}>ì§€ì¶œ í†µê³„ ðŸ“Š</h1>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px"}}>
          {[{l:"ì›”í‰ê·  ì§€ì¶œ",v:`${Math.round(avgExp/10000)}ë§Œì›`},{l:"ìµœë‹¤ ì§€ì¶œì›”",v:maxM.month},{l:"ì´ë²ˆë‹¬ ì €ì¶•ë¥ ",v:`${sr}%`}].map(({l,v})=>(
            <div key={l} style={{background:"rgba(255,255,255,0.15)",borderRadius:"12px",padding:"10px",textAlign:"center"}}><p style={{margin:0,fontSize:"9px",opacity:0.8}}>{l}</p><p style={{margin:"4px 0 0",fontSize:"16px",fontWeight:"700"}}>{v}</p></div>
          ))}
        </div>
      </div>
      <div style={{padding:"20px 16px 100px"}}>
        <div style={{display:"flex",background:"#E2E8F0",borderRadius:"12px",padding:"4px",marginBottom:"20px"}}>
          {[["trend","ì›”ë³„ ì¶”ì´"],["category","ì¹´í…Œê³ ë¦¬"]].map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"8px",border:"none",cursor:"pointer",borderRadius:"10px",background:tab===k?"white":"transparent",color:tab===k?"#6366F1":"#64748B",fontWeight:tab===k?"700":"400",fontSize:"14px",boxShadow:tab===k?"0 1px 4px rgba(0,0,0,0.1)":"none"}}>{l}</button>
          ))}
        </div>
        {tab==="trend"&&(
          <>
            <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px"}}>
              <div style={{display:"flex",gap:"16px",marginBottom:"12px"}}>
                {[["#10B981","ìˆ˜ìž…"],["#6366F1","ì§€ì¶œ"]].map(([c,l])=>(
                  <div key={l} style={{display:"flex",alignItems:"center",gap:"5px"}}><div style={{width:"12px",height:"3px",background:c,borderRadius:"2px"}}/><span style={{fontSize:"11px",color:"#64748B"}}>{l}</span></div>
                ))}
              </div>
              <TrendChart data={MONTHLY_TREND} selectedIdx={selM} onSelect={setSelM}/>
            </div>
            <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"16px"}}>
              <h2 style={{margin:"0 0 14px",fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>{sel.month} ìƒì„¸{selM===5&&<span style={{fontSize:"11px",color:"#6366F1",marginLeft:"6px"}}>ì§„í–‰ì¤‘</span>}</h2>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"10px"}}>
                {[{l:"ìˆ˜ìž…",v:sel.income,c:"#10B981"},{l:"ì§€ì¶œ",v:sel.expense,c:"#6366F1"},{l:saving>=0?"ì €ì¶•":"ì´ˆê³¼",v:Math.abs(saving),c:saving>=0?"#3B82F6":"#EF4444"}].map(({l,v,c})=>(
                  <div key={l} style={{background:c+"10",borderRadius:"12px",padding:"12px 10px",textAlign:"center",border:`1px solid ${c}20`}}><p style={{margin:0,fontSize:"11px",color:"#64748B"}}>{l}</p><p style={{margin:"4px 0 0",fontSize:"14px",fontWeight:"700",color:c}}>{Math.round(v/10000)}ë§Œ</p></div>
                ))}
              </div>
              <div style={{marginTop:"14px"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"6px"}}><span style={{fontSize:"12px",color:"#64748B"}}>ì €ì¶•ë¥ </span><span style={{fontSize:"12px",fontWeight:"700",color:sr>=20?"#10B981":sr>=10?"#F97316":"#EF4444"}}>{sr}%</span></div>
                <div style={{background:"#F1F5F9",borderRadius:"99px",height:"8px"}}><div style={{width:`${Math.max(0,sr)}%`,height:"100%",background:sr>=20?"#10B981":sr>=10?"#F97316":"#EF4444",borderRadius:"99px"}}/></div>
              </div>
            </div>
            <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
              <h2 style={{margin:"0 0 14px",fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>ì›”ë³„ ì§€ì¶œ ë¹„êµ</h2>
              {MONTHLY_TREND.map((m,i)=>{const pct=Math.round((m.expense/Math.max(...MONTHLY_TREND.map(d=>d.expense)))*100);const iS=i===selM;return(
                <div key={m.month} onClick={()=>setSelM(i)} style={{marginBottom:i<5?"10px":"0",cursor:"pointer"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:"4px"}}><span style={{fontSize:"12px",fontWeight:iS?"700":"400",color:iS?"#6366F1":"#64748B"}}>{m.month}</span><span style={{fontSize:"12px",fontWeight:"700",color:iS?"#6366F1":"#1E293B"}}>{m.expense.toLocaleString()}ì›</span></div>
                  <div style={{background:"#F1F5F9",borderRadius:"99px",height:"6px"}}><div style={{width:`${pct}%`,height:"100%",background:iS?"#6366F1":"#C7D2FE",borderRadius:"99px"}}/></div>
                </div>
              );})}
            </div>
          </>
        )}
        {tab==="category"&&(
          <>
            <div style={{display:"flex",gap:"8px",overflowX:"auto",paddingBottom:"4px",marginBottom:"16px"}}>
              {Object.keys(CAT_TREND).map(cat=>{const m=CATEGORIES.find(c=>c.name===cat)||CATEGORIES.at(-1);const iS=selC===cat;return(
                <button key={cat} onClick={()=>setSelC(cat)} style={{display:"flex",alignItems:"center",gap:"5px",padding:"7px 12px",border:"none",borderRadius:"99px",cursor:"pointer",background:iS?m.color:"#F1F5F9",color:iS?"white":"#64748B",fontWeight:iS?"700":"400",fontSize:"12px",flexShrink:0}}><span>{m.icon}</span><span>{cat}</span></button>
              );})}
            </div>
            {(()=>{const m=CATEGORIES.find(c=>c.name===selC)||CATEGORIES.at(-1);return(
              <>
                <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"12px"}}>
                    <div><h2 style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>{m.icon} {selC} ì›”ë³„ ì¶”ì´</h2><p style={{margin:"4px 0 0",fontSize:"11px",color:"#94A3B8"}}>ìµœê·¼ 6ê°œì›”</p></div>
                    <div style={{textAlign:"right"}}><p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>ì´ë²ˆë‹¬</p><p style={{margin:"2px 0 0",fontSize:"16px",fontWeight:"700",color:m.color}}>{CAT_TREND[selC].at(-1).toLocaleString()}ì›</p></div>
                  </div>
                  <CatChart data={CAT_TREND[selC]} color={m.color}/>
                  <div style={{display:"flex",justifyContent:"space-between",marginTop:"4px"}}>
                    {months.map((mo,i)=><span key={mo} style={{fontSize:"10px",color:i===months.length-1?m.color:"#94A3B8",fontWeight:i===months.length-1?"700":"400"}}>{mo}</span>)}
                  </div>
                </div>
                <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <h2 style={{margin:"0 0 12px",fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>ì›”ë³„ ìƒì„¸</h2>
                  {CAT_TREND[selC].map((v,i)=>{const prev=i>0?CAT_TREND[selC][i-1]:null;const diff=prev?v-prev:0;const pct=Math.round((v/Math.max(...CAT_TREND[selC]))*100);const iL=i===CAT_TREND[selC].length-1;return(
                    <div key={i} style={{marginBottom:i<5?"10px":"0"}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"4px"}}>
                        <div style={{display:"flex",alignItems:"center",gap:"6px"}}><span style={{fontSize:"12px",fontWeight:iL?"700":"400",color:iL?m.color:"#64748B"}}>{months[i]}</span>{prev&&<span style={{fontSize:"10px",color:diff>0?"#EF4444":"#10B981",fontWeight:"600"}}>{diff>0?"â–²":"â–¼"} {Math.abs(diff).toLocaleString()}ì›</span>}</div>
                        <span style={{fontSize:"12px",fontWeight:iL?"700":"400",color:iL?m.color:"#1E293B"}}>{v.toLocaleString()}ì›</span>
                      </div>
                      <div style={{background:"#F1F5F9",borderRadius:"99px",height:"5px"}}><div style={{width:`${pct}%`,height:"100%",background:iL?m.color:m.color+"50",borderRadius:"99px"}}/></div>
                    </div>
                  );})}
                </div>
              </>
            );})()}
          </>
        )}
      </div>
    </div>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// í™”ë©´ 5: ì•Œë¦¼
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const AlertScreen=()=>{
  const [expanded,setExpanded]=useState(null);
  const [showAll,setShowAll]=useState(false);
  const analyzed=MONTHLY_CATS.map(cat=>{
    const level=getLevel(cat.amount,cat.avg);
    const pct=Math.round((cat.amount/cat.avg)*100);
    const diff=cat.amount-cat.avg;
    const spark=[Math.round(cat.avg*0.9),Math.round(cat.avg*1.05),cat.avg,cat.amount];
    return{...cat,level,pct,diff,spark};
  });
  const alerts=analyzed.filter(c=>c.level==="alert");
  const warns=analyzed.filter(c=>c.level==="warn");
  const oks=analyzed.filter(c=>c.level==="ok");
  const issues=[...alerts,...warns];

  return(
    <div>
      <div style={{background:alerts.length>0?"linear-gradient(135deg,#EF4444,#DC2626)":warns.length>0?"linear-gradient(135deg,#F97316,#EA580C)":"linear-gradient(135deg,#10B981,#059669)",padding:"48px 24px 24px",color:"white"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div><p style={{margin:0,fontSize:"13px",opacity:0.8}}>2026ë…„ 6ì›” ì§€ì¶œ ì•Œë¦¼</p><h1 style={{fontSize:"20px",fontWeight:"700",margin:"4px 0 0"}}>{alerts.length>0?`${alerts.length}ê°œ ì¹´í…Œê³ ë¦¬ ê²½ê³  ðŸš¨`:warns.length>0?`${warns.length}ê°œ ì¹´í…Œê³ ë¦¬ ì£¼ì˜ âš ï¸`:"ì´ë²ˆ ë‹¬ ì§€ì¶œ ì–‘í˜¸í•´ìš” âœ…"}</h1></div>
          <div style={{background:"rgba(255,255,255,0.2)",borderRadius:"12px",padding:"8px 12px",textAlign:"center"}}><p style={{margin:0,fontSize:"22px",fontWeight:"700"}}>{issues.length}</p><p style={{margin:0,fontSize:"10px",opacity:0.8}}>ì£¼ì˜ í•­ëª©</p></div>
        </div>
        <div style={{display:"flex",gap:"8px",marginTop:"16px"}}>
          {[{level:"alert",count:alerts.length,label:"ê²½ê³ "},{level:"warn",count:warns.length,label:"ì£¼ì˜"},{level:"ok",count:oks.length,label:"ì •ìƒ"}].map(({level,count,label})=>(
            <div key={level} style={{background:"rgba(255,255,255,0.2)",borderRadius:"10px",padding:"6px 12px",display:"flex",alignItems:"center",gap:"6px"}}><span style={{fontSize:"14px"}}>{LVL[level].icon}</span><span style={{fontSize:"13px",fontWeight:"700"}}>{count}</span><span style={{fontSize:"11px",opacity:0.8}}>{label}</span></div>
          ))}
        </div>
      </div>
      <div style={{padding:"20px 16px 100px"}}>
        {issues.length>0&&(
          <div style={{marginBottom:"24px"}}>
            <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>ì£¼ì˜ê°€ í•„ìš”í•œ ì¹´í…Œê³ ë¦¬</p>
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {issues.map(cat=>{const s=LVL[cat.level];const iO=expanded===cat.name;return(
                <div key={cat.name} style={{background:"white",borderRadius:"16px",boxShadow:"0 2px 10px rgba(0,0,0,0.07)",border:`1px solid ${s.border}`,overflow:"hidden"}}>
                  <div onClick={()=>setExpanded(iO?null:cat.name)} style={{padding:"14px 16px",cursor:"pointer"}}>
                    <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
                      <div style={{width:"44px",height:"44px",borderRadius:"12px",background:cat.color+"20",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"22px",flexShrink:0,position:"relative"}}>
                        {cat.icon}
                        <div style={{position:"absolute",top:"-4px",right:"-4px",width:"16px",height:"16px",borderRadius:"50%",background:s.badge,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"9px"}}>{cat.level==="alert"?"ðŸš¨":"âš ï¸"}</div>
                      </div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:"6px",marginBottom:"4px"}}>
                          <span style={{fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>{cat.name}</span>
                          <span style={{fontSize:"10px",padding:"2px 7px",borderRadius:"99px",background:s.badgeBg,color:s.badge,fontWeight:"700",border:`1px solid ${s.border}`}}>{s.label} {cat.pct}%</span>
                        </div>
                        <div style={{background:"#F1F5F9",borderRadius:"99px",height:"6px",overflow:"hidden"}}><div style={{width:`${Math.min(cat.pct,100)}%`,height:"100%",background:s.badge,borderRadius:"99px"}}/></div>
                        <div style={{display:"flex",justifyContent:"space-between",marginTop:"4px"}}><span style={{fontSize:"11px",color:"#94A3B8"}}>í‰ê·  {cat.avg.toLocaleString()}ì›</span><span style={{fontSize:"11px",fontWeight:"700",color:s.badge}}>{cat.amount.toLocaleString()}ì›</span></div>
                      </div>
                      <span style={{fontSize:"16px",color:"#94A3B8"}}>{iO?"â–²":"â–¼"}</span>
                    </div>
                  </div>
                  {iO&&(
                    <div style={{padding:"0 16px 14px",borderTop:`1px solid ${s.border}`}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:"12px",marginBottom:"10px"}}>
                        <div>
                          <p style={{margin:"0 0 4px",fontSize:"12px",color:"#64748B"}}>ìµœê·¼ 4ê°œì›” ì¶”ì´</p>
                          <div style={{display:"flex",gap:"12px"}}>
                            {["3ë‹¬ì „","2ë‹¬ì „","ì§€ë‚œë‹¬","ì´ë²ˆë‹¬"].map((l,i)=>(
                              <div key={l} style={{textAlign:"center"}}><p style={{margin:0,fontSize:"9px",color:"#94A3B8"}}>{l}</p><p style={{margin:"2px 0 0",fontSize:"10px",fontWeight:"600",color:i===3?s.badge:"#64748B"}}>{Math.round(cat.spark[i]/1000)}ë§Œ</p></div>
                            ))}
                          </div>
                        </div>
                        <div style={{background:s.bg,borderRadius:"12px",padding:"10px 14px",border:`1px solid ${s.border}`,textAlign:"right"}}>
                          <p style={{margin:0,fontSize:"11px",color:s.text}}>í‰ê·  ëŒ€ë¹„</p>
                          <p style={{margin:"2px 0 0",fontSize:"18px",fontWeight:"700",color:s.badge}}>+{cat.diff.toLocaleString()}ì›</p>
                          <p style={{margin:"2px 0 0",fontSize:"11px",color:s.text}}>{cat.pct-100}% ì´ˆê³¼</p>
                        </div>
                      </div>
                      <div style={{background:"#F8FAFC",borderRadius:"10px",padding:"10px 12px",display:"flex",gap:"8px"}}>
                        <span style={{fontSize:"14px",flexShrink:0}}>ðŸ¤–</span>
                        <p style={{margin:0,fontSize:"12px",color:"#334155",lineHeight:"1.6"}}>
                          {cat.name==="ì‡¼í•‘"&&"ì´ë²ˆ ë‹¬ ì‡¼í•‘ì´ í‰ì†Œë³´ë‹¤ ë§Žì´ ëŠ˜ì—ˆì–´ìš”. ìž¥ë°”êµ¬ë‹ˆì— ë‹´ì•„ë‘ê³  3ì¼ í›„ ë‹¤ì‹œ í™•ì¸í•˜ëŠ” ìŠµê´€ì´ ì¶©ë™êµ¬ë§¤ë¥¼ ì¤„ì´ëŠ” ë° ë„ì›€ì´ ë¼ìš”."}
                          {cat.name==="ì¹´íŽ˜"&&"ì¹´íŽ˜ ì§€ì¶œì´ í‰ì†Œë³´ë‹¤ ë†’ì•„ìš”. í™ˆì¹´íŽ˜ë¡œ ì „í™˜í•˜ê±°ë‚˜ í…€ë¸”ëŸ¬ ì§€ì°¸ ì‹œ í• ì¸ë˜ëŠ” ë§¤ìž¥ì„ í™œìš©í•´ë³´ì„¸ìš”."}
                          {!["ì‡¼í•‘","ì¹´íŽ˜"].includes(cat.name)&&`${cat.name} ì§€ì¶œì´ í‰ì†Œë³´ë‹¤ ${cat.pct-100}% ë†’ì•„ìš”. ë¶ˆí•„ìš”í•œ í•­ëª©ì´ ì—†ëŠ”ì§€ ì ê²€í•´ë³´ì„¸ìš”.`}
                        </p>
                      </div>
                    </div>
                  )}
                </div>
              );})}
            </div>
          </div>
        )}
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"10px"}}>
            <p style={{margin:0,fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>ì •ìƒ ë²”ìœ„ ì¹´í…Œê³ ë¦¬</p>
            <button onClick={()=>setShowAll(p=>!p)} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>{showAll?"ì ‘ê¸°":"íŽ¼ì¹˜ê¸°"}</button>
          </div>
          {showAll?(
            <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
              {oks.map(cat=>(
                <div key={cat.name} style={{background:"white",borderRadius:"14px",padding:"12px 16px",display:"flex",alignItems:"center",gap:"12px",boxShadow:"0 1px 4px rgba(0,0,0,0.05)",border:"1px solid #DCFCE7"}}>
                  <div style={{width:"38px",height:"38px",borderRadius:"10px",background:cat.color+"15",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"20px",flexShrink:0}}>{cat.icon}</div>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:"6px"}}><span style={{fontSize:"13px",fontWeight:"600",color:"#1E293B"}}>{cat.name}</span><span style={{fontSize:"10px",padding:"1px 6px",borderRadius:"99px",background:"#DCFCE7",color:"#15803D",fontWeight:"700"}}>âœ… {cat.pct}%</span></div>
                    <div style={{background:"#F1F5F9",borderRadius:"99px",height:"4px",marginTop:"6px"}}><div style={{width:`${Math.min(cat.pct,100)}%`,height:"100%",background:"#10B981",borderRadius:"99px"}}/></div>
                  </div>
                  <div style={{textAlign:"right"}}><p style={{margin:0,fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>{cat.amount.toLocaleString()}ì›</p><p style={{margin:"2px 0 0",fontSize:"10px",color:"#94A3B8"}}>í‰ê·  {cat.avg.toLocaleString()}ì›</p></div>
                </div>
              ))}
            </div>
          ):(
            <div style={{background:"#F0FDF4",borderRadius:"12px",padding:"12px 16px",border:"1px solid #DCFCE7",display:"flex",alignItems:"center",gap:"10px"}}>
              <span style={{fontSize:"18px"}}>âœ…</span>
              <p style={{margin:0,fontSize:"13px",color:"#15803D"}}><strong>{oks.length}ê°œ</strong> ì¹´í…Œê³ ë¦¬ëŠ” í‰ê·  ë²”ìœ„ ë‚´ ì§€ì¶œ ì¤‘ì´ì—ìš”</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// í™”ë©´ 6: ê°€ì¡± ê³µìœ 
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const FamilyScreen=({txList,setTxList})=>{
  const [tab,setTab]=useState("txlist");
  const [filterAuthor,setFilterAuthor]=useState(null);
  const [showInvite,setShowInvite]=useState(false);
  const [editTx,setEditTx]=useState(null);
  const members=FAMILY.members;
  const getMember=(id)=>members.find(m=>m.id===id)||members[0];
  const handleSave=(updated)=>{setTxList(p=>p.map(t=>t.id===updated.id?updated:t));setEditTx(null);};
  const handleDelete=(id)=>{setTxList(p=>p.filter(t=>t.id!==id));setEditTx(null);};
  const filtered=filterAuthor?txList.filter(t=>t.author===filterAuthor):txList;
  const totalExp=txList.filter(t=>t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 20px",color:"white"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"16px"}}>
          <div><p style={{margin:0,fontSize:"13px",opacity:0.8}}>ìš°ë¦¬ì§‘ ê°€ê³„ë¶€ ðŸ‘¨â€ðŸ‘©â€ðŸ‘¦</p><h1 style={{fontSize:"20px",fontWeight:"700",margin:"4px 0 0"}}>ê°€ì¡± ê³µìœ </h1></div>
          <button onClick={()=>setShowInvite(true)} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:"12px",padding:"8px 12px",color:"white",fontSize:"12px",fontWeight:"600",cursor:"pointer"}}>ðŸ‘¥ ì´ˆëŒ€í•˜ê¸°</button>
        </div>
        <div style={{display:"flex",gap:"8px",alignItems:"center"}}>
          {members.map(m=>(
            <button key={m.id} onClick={()=>setFilterAuthor(filterAuthor===m.id?null:m.id)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"4px",background:"none",border:"none",cursor:"pointer"}}>
              <div style={{width:"40px",height:"40px",borderRadius:"50%",background:filterAuthor===m.id?"white":m.color+"40",border:`2px solid ${filterAuthor===m.id?"white":"transparent"}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"20px"}}>{m.emoji}</div>
              <span style={{fontSize:"10px",opacity:filterAuthor===m.id?1:0.7,fontWeight:filterAuthor===m.id?"700":"400"}}>{m.name}</span>
            </button>
          ))}
          {filterAuthor&&<button onClick={()=>setFilterAuthor(null)} style={{marginLeft:"4px",background:"rgba(255,255,255,0.2)",border:"none",borderRadius:"99px",padding:"4px 10px",color:"white",fontSize:"11px",cursor:"pointer",fontWeight:"600"}}>ì „ì²´ë³´ê¸°</button>}
        </div>
      </div>
      <div style={{display:"flex",background:"#E2E8F0",padding:"4px"}}>
        {[["txlist","ë‚´ì—­ ëª©ë¡"],["family","ë©¤ë²„ ê´€ë¦¬"]].map(([k,l])=>(
          <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"10px",border:"none",cursor:"pointer",background:tab===k?"white":"transparent",color:tab===k?"#6366F1":"#64748B",fontWeight:tab===k?"700":"400",fontSize:"14px"}}>{l}</button>
        ))}
      </div>
      <div style={{padding:"16px 16px 100px"}}>
        {tab==="txlist"&&(
          <>
            <div style={{background:"white",borderRadius:"16px",padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><p style={{margin:0,fontSize:"12px",color:"#94A3B8"}}>ì´ë²ˆë‹¬ ê°€ì¡± ì´ ì§€ì¶œ</p><p style={{margin:"4px 0 0",fontSize:"20px",fontWeight:"700",color:"#6366F1"}}>{totalExp.toLocaleString()}ì›</p></div>
              <div style={{display:"flex",gap:"6px"}}>
                {members.map(m=>{const mt=txList.filter(t=>t.author===m.id&&t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);const pct=totalExp>0?Math.round((mt/totalExp)*100):0;return(<div key={m.id} style={{textAlign:"center"}}><div style={{fontSize:"18px"}}>{m.emoji}</div><p style={{margin:"2px 0 0",fontSize:"10px",color:m.color,fontWeight:"700"}}>{pct}%</p></div>);})}
              </div>
            </div>
            <div style={{background:"white",borderRadius:"14px",padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"16px"}}>
              <p style={{margin:"0 0 10px",fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>ë©¤ë²„ë³„ ì§€ì¶œ ë¹„ì¤‘</p>
              <div style={{display:"flex",height:"8px",borderRadius:"99px",overflow:"hidden",marginBottom:"10px"}}>
                {members.map(m=>{const mt=txList.filter(t=>t.author===m.id&&t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);const pct=totalExp>0?(mt/totalExp)*100:0;return <div key={m.id} style={{width:`${pct}%`,background:m.color}}/>;}) }
              </div>
              <div style={{display:"flex",gap:"12px"}}>
                {members.map(m=>{const mt=txList.filter(t=>t.author===m.id&&t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);return(<div key={m.id} style={{display:"flex",alignItems:"center",gap:"5px"}}><div style={{width:"8px",height:"8px",borderRadius:"2px",background:m.color,flexShrink:0}}/><span style={{fontSize:"11px",color:"#64748B"}}>{m.name} {mt.toLocaleString()}ì›</span></div>);})}
              </div>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
              {filtered.map(tx=>{const author=getMember(tx.author);return(
                <button key={tx.id} onClick={()=>setEditTx(tx)} style={{background:"white",border:"none",borderRadius:"14px",padding:"14px 16px",display:"flex",alignItems:"center",gap:"12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",cursor:"pointer",textAlign:"left",width:"100%"}}>
                  <div style={{width:"42px",height:"42px",borderRadius:"12px",background:tx.amount>0?"#DCFCE7":"#F1F5F9",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"20px",flexShrink:0,position:"relative"}}>
                    {tx.icon}
                    <div style={{position:"absolute",bottom:"-4px",right:"-4px",width:"16px",height:"16px",borderRadius:"50%",background:author.color,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"9px",border:"2px solid white"}}>{author.emoji}</div>
                  </div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",justifyContent:"space-between"}}>
                      <div><p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{tx.name}</p><div style={{display:"flex",alignItems:"center",gap:"5px",marginTop:"2px"}}><span style={{fontSize:"10px",color:"#94A3B8"}}>{tx.category}</span><span style={{fontSize:"10px",color:"#CBD5E1"}}>Â·</span><span style={{fontSize:"10px",color:author.color,fontWeight:"600"}}>{author.name}</span>{tx.memo&&<><span style={{fontSize:"10px",color:"#CBD5E1"}}>Â·</span><span style={{fontSize:"10px",color:"#94A3B8"}}>{tx.memo}</span></>}</div></div>
                      <div style={{textAlign:"right"}}><p style={{margin:0,fontSize:"14px",fontWeight:"700",color:tx.amount>0?"#10B981":"#1E293B"}}>{fmt(tx.amount)}</p><p style={{margin:"2px 0 0",fontSize:"10px",color:"#CBD5E1"}}>{tx.date}</p></div>
                    </div>
                  </div>
                  <span style={{fontSize:"12px",color:"#CBD5E1"}}>âœï¸</span>
                </button>
              );})}
            </div>
          </>
        )}
        {tab==="family"&&(
          <>
            <div style={{background:"white",borderRadius:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",overflow:"hidden",marginBottom:"16px"}}>
              {members.map((m,i)=>(
                <div key={m.id} style={{display:"flex",alignItems:"center",gap:"12px",padding:"14px 16px",borderBottom:i<members.length-1?"1px solid #F8FAFC":"none"}}>
                  <div style={{width:"44px",height:"44px",borderRadius:"50%",background:m.color+"20",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"22px"}}>{m.emoji}</div>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:"6px",marginBottom:"2px"}}>
                      <p style={{margin:0,fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>{m.name}</p>
                      <span style={{fontSize:"10px",padding:"2px 7px",borderRadius:"99px",fontWeight:"700",background:m.role==="ê´€ë¦¬ìž"?"#EEF2FF":m.role==="ë©¤ë²„"?"#F0FDF4":"#F8FAFC",color:m.role==="ê´€ë¦¬ìž"?"#6366F1":m.role==="ë©¤ë²„"?"#10B981":"#94A3B8"}}>{m.role}</span>
                    </div>
                    <p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>{m.joined} ì°¸ì—¬</p>
                  </div>
                  <div style={{textAlign:"right"}}><p style={{margin:0,fontSize:"13px",fontWeight:"700",color:m.color}}>{txList.filter(t=>t.author===m.id&&t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0).toLocaleString()}ì›</p><p style={{margin:"2px 0 0",fontSize:"10px",color:"#94A3B8"}}>ì´ë²ˆë‹¬</p></div>
                </div>
              ))}
            </div>
            <button onClick={()=>setShowInvite(true)} style={{width:"100%",padding:"14px",border:"2px dashed #C7D2FE",borderRadius:"16px",background:"#EEF2FF",color:"#6366F1",fontSize:"14px",fontWeight:"700",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:"8px",marginBottom:"16px"}}>
              <span style={{fontSize:"18px"}}>ðŸ‘¥</span> ê°€ì¡± ì´ˆëŒ€í•˜ê¸°
            </button>
            <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
              <p style={{margin:"0 0 12px",fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>ìš°ë¦¬ ê°€ì¡± ì½”ë“œ</p>
              <div style={{background:"linear-gradient(135deg,#EEF2FF,#F5F3FF)",border:"2px dashed #A5B4FC",borderRadius:"12px",padding:"16px",textAlign:"center"}}>
                <p style={{margin:"0 0 6px",fontSize:"11px",color:"#6366F1",fontWeight:"600"}}>ì´ˆëŒ€ ì½”ë“œ</p>
                <p style={{margin:0,fontSize:"24px",fontWeight:"700",color:"#4338CA",letterSpacing:"3px"}}>{FAMILY.code}</p>
              </div>
            </div>
          </>
        )}
      </div>
      {editTx&&<EditModal tx={editTx} onSave={handleSave} onDelete={handleDelete} onClose={()=>setEditTx(null)}/>}
      {showInvite&&<InviteModal code={FAMILY.code} onClose={()=>setShowInvite(false)}/>}
    </div>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// í™”ë©´ 7: ì„¤ì •
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
const SettingsScreen=()=>{
  const [selCards,setSelCards]=useState(["shinhan","hyundai","samsung"]);
  const [period,setPeriod]=useState("30");
  const [minAmt,setMinAmt]=useState("1000");
  const [exCancel,setExCancel]=useState(true);
  const [exOverseas,setExOverseas]=useState(false);
  const [keywords,setKeywords]=useState(["ìŠ¹ì¸ì·¨ì†Œ","í•´ì™¸ê²°ì œ","ì„ ë¶ˆì¶©ì „","í¬ì¸íŠ¸ì ë¦½"]);
  const [newKw,setNewKw]=useState("");
  const [autoCat,setAutoCat]=useState(true);
  const [autoAI,setAutoAI]=useState(true);
  const [saved,setSaved]=useState(false);
  const toggleCard=(id)=>setSelCards(p=>p.includes(id)?p.filter(c=>c!==id):[...p,id]);
  const addKw=()=>{if(!newKw.trim())return;setKeywords(p=>[...p,newKw.trim()]);setNewKw("");};

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 24px",color:"white"}}>
        <h1 style={{margin:"0 0 4px",fontSize:"20px",fontWeight:"700"}}>ë¬¸ìž ë¶ˆëŸ¬ì˜¤ê¸° ì„¤ì • âš™ï¸</h1>
        <p style={{margin:0,fontSize:"13px",opacity:0.75}}>ì–´ë–¤ ë¬¸ìžë¥¼ ì½ì–´ì˜¬ì§€ ì„¤ì •í•´ìš”</p>
      </div>
      <div style={{padding:"20px 16px 140px"}}>
        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>ì¹´ë“œì‚¬ ì„ íƒ</p>
        <div style={{background:"white",borderRadius:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px",overflow:"hidden"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 16px",borderBottom:"1px solid #F1F5F9",background:"#FAFBFF"}}>
            <span style={{fontSize:"13px",color:"#64748B",fontWeight:"600"}}>{selCards.length}ê°œ ì„ íƒë¨</span>
            <button onClick={()=>setSelCards(selCards.length===CARD_PROVIDERS.length?[]:CARD_PROVIDERS.map(c=>c.id))} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>{selCards.length===CARD_PROVIDERS.length?"ì „ì²´í•´ì œ":"ì „ì²´ì„ íƒ"}</button>
          </div>
          {CARD_PROVIDERS.map((card,i)=>{const iS=selCards.includes(card.id);return(
            <div key={card.id} onClick={()=>toggleCard(card.id)} style={{display:"flex",alignItems:"center",gap:"12px",padding:"12px 16px",borderBottom:i<CARD_PROVIDERS.length-1?"1px solid #F8FAFC":"none",cursor:"pointer",background:iS?"#FAFBFF":"white"}}>
              <div style={{width:"22px",height:"22px",borderRadius:"6px",background:iS?"#6366F1":"white",border:`2px solid ${iS?"#6366F1":"#CBD5E1"}`,display:"flex",alignItems:"center",justifyContent:"center",color:"white",fontSize:"13px",fontWeight:"700",flexShrink:0}}>{iS?"âœ“":""}</div>
              <span style={{fontSize:"16px"}}>{card.emoji}</span>
              <div style={{flex:1}}><p style={{margin:0,fontSize:"14px",fontWeight:iS?"600":"400",color:"#1E293B"}}>{card.name}</p><p style={{margin:"1px 0 0",fontSize:"10px",color:"#94A3B8",fontFamily:"monospace"}}>ë°œì‹ ë²ˆí˜¸ {card.sender}</p></div>
              <span style={{fontSize:"10px",padding:"2px 6px",borderRadius:"99px",background:"#F1F5F9",color:"#64748B",fontFamily:"monospace"}}>{card.prefix}</span>
            </div>
          );})}
        </div>

        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>ì½ê¸° ê¸°ê°„</p>
        <div style={{background:"white",borderRadius:"16px",padding:"12px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px"}}>
            {[{v:"7",l:"ìµœê·¼ 1ì£¼"},{v:"30",l:"ìµœê·¼ 1ê°œì›”"},{v:"90",l:"ìµœê·¼ 3ê°œì›”"},{v:"custom",l:"ì§ì ‘ ì„¤ì •"}].map(o=>(
              <button key={o.v} onClick={()=>setPeriod(o.v)} style={{padding:"10px",border:"none",borderRadius:"10px",cursor:"pointer",background:period===o.v?"#6366F1":"#F1F5F9",color:period===o.v?"white":"#64748B",fontSize:"13px",fontWeight:period===o.v?"700":"400"}}>{o.l}</button>
            ))}
          </div>
        </div>

        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>ê¸ˆì•¡ í•„í„°</p>
        <div style={{background:"white",borderRadius:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px",overflow:"hidden"}}>
          {[
            {icon:"ðŸ’°",label:"ìµœì†Œ ê¸ˆì•¡ ì´í•˜ ì œì™¸",sub:`${parseInt(minAmt||0).toLocaleString()}ì› ë¯¸ë§Œ ì œì™¸`,right:<div style={{display:"flex",alignItems:"center",gap:"4px"}}><input type="number" value={minAmt} onChange={e=>setMinAmt(e.target.value)} onClick={e=>e.stopPropagation()} style={{width:"70px",padding:"6px 8px",border:"2px solid #E2E8F0",borderRadius:"8px",fontSize:"13px",color:"#6366F1",fontWeight:"600",outline:"none",textAlign:"right"}}/><span style={{fontSize:"12px",color:"#94A3B8"}}>ì›</span></div>},
            {icon:"ðŸ”„",label:"ìŠ¹ì¸ì·¨ì†Œ ì œì™¸",sub:"ì·¨ì†ŒÂ·í™˜ë¶ˆ ë¬¸ìžëŠ” ë¶ˆëŸ¬ì˜¤ì§€ ì•Šì•„ìš”",right:<Toggle value={exCancel} onChange={setExCancel}/>},
            {icon:"âœˆï¸",label:"í•´ì™¸ê²°ì œ ì œì™¸",sub:"í•´ì™¸ ê°€ë§¹ì  ê²°ì œ ì œì™¸",right:<Toggle value={exOverseas} onChange={setExOverseas}/>},
          ].map((row,i,arr)=>(
            <div key={row.label} style={{display:"flex",alignItems:"center",gap:"12px",padding:"13px 16px",borderBottom:i<arr.length-1?"1px solid #F8FAFC":"none"}}>
              <span style={{fontSize:"18px"}}>{row.icon}</span>
              <div style={{flex:1}}><p style={{margin:0,fontSize:"14px",fontWeight:"500",color:"#1E293B"}}>{row.label}</p><p style={{margin:"2px 0 0",fontSize:"11px",color:"#94A3B8"}}>{row.sub}</p></div>
              {row.right}
            </div>
          ))}
        </div>

        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>ì œì™¸ í‚¤ì›Œë“œ</p>
        <div style={{background:"white",borderRadius:"16px",padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px"}}>
          <div style={{display:"flex",flexWrap:"wrap",gap:"6px",marginBottom:"10px"}}>
            {keywords.map(kw=>(
              <div key={kw} style={{display:"flex",alignItems:"center",gap:"4px",background:"#FEE2E2",borderRadius:"99px",padding:"4px 10px"}}>
                <span style={{fontSize:"12px",color:"#EF4444",fontWeight:"600"}}>{kw}</span>
                <button onClick={()=>setKeywords(p=>p.filter(k=>k!==kw))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",fontSize:"14px",padding:"0",lineHeight:1,fontWeight:"700"}}>Ã—</button>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"8px"}}>
            <input value={newKw} onChange={e=>setNewKw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addKw()} placeholder="í‚¤ì›Œë“œ ì¶”ê°€" style={{flex:1,padding:"8px 12px",border:"2px solid #E2E8F0",borderRadius:"10px",fontSize:"13px",outline:"none"}}/>
            <button onClick={addKw} style={{background:"#6366F1",border:"none",borderRadius:"10px",padding:"8px 14px",color:"white",fontSize:"13px",fontWeight:"700",cursor:"pointer"}}>ì¶”ê°€</button>
          </div>
        </div>

        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>AI ìžë™í™” ì„¤ì •</p>
        <div style={{background:"white",borderRadius:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px",overflow:"hidden"}}>
          {[{icon:"ðŸ¤–",label:"AI ì¹´í…Œê³ ë¦¬ ìžë™ ë¶„ë¥˜",sub:"ìƒí˜¸ëª…ìœ¼ë¡œ ì¹´í…Œê³ ë¦¬ë¥¼ ìžë™ìœ¼ë¡œ ë¶„ë¥˜í•´ìš”",value:autoCat,onChange:setAutoCat},{icon:"ðŸ“Š",label:"AI ì§€ì¶œ ë¶„ì„ ìžë™ ì‹¤í–‰",sub:"í™ˆ í™”ë©´ ì§„ìž… ì‹œ ìžë™ìœ¼ë¡œ ë¶„ì„í•´ìš”",value:autoAI,onChange:setAutoAI}].map((row,i,arr)=>(
            <div key={row.label} style={{display:"flex",alignItems:"center",gap:"12px",padding:"13px 16px",borderBottom:i<arr.length-1?"1px solid #F8FAFC":"none"}}>
              <span style={{fontSize:"18px"}}>{row.icon}</span>
              <div style={{flex:1}}><p style={{margin:0,fontSize:"14px",fontWeight:"500",color:"#1E293B"}}>{row.label}</p><p style={{margin:"2px 0 0",fontSize:"11px",color:"#94A3B8"}}>{row.sub}</p></div>
              <Toggle value={row.value} onChange={row.onChange}/>
            </div>
          ))}
        </div>
      </div>
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:"430px",background:"white",borderTop:"1px solid #F1F5F9",padding:"12px 16px 72px",boxShadow:"0 -4px 20px rgba(0,0,0,0.08)"}}>
        <button onClick={()=>{setSaved(true);setTimeout(()=>setSaved(false),2000);}} style={{width:"100%",padding:"14px",border:"none",borderRadius:"14px",background:saved?"#10B981":"linear-gradient(135deg,#6366F1,#8B5CF6)",color:"white",fontSize:"15px",fontWeight:"700",cursor:"pointer",transition:"background 0.3s"}}>{saved?"âœ… ì„¤ì •ì´ ì €ìž¥ëì–´ìš”!":"ì„¤ì • ì €ìž¥í•˜ê¸°"}</button>
      </div>
    </div>
  );
};

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ë©”ì¸ ì•± ë¼ìš°í„°
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

// â”€â”€ ì¹´í…Œê³ ë¦¬ ë§¤í•‘ â”€â”€
const SUBCAT_MAP = {
  "ì»¤í”¼/ìŒë£Œ": { name:"ì¹´íŽ˜",   icon:"â˜•",  color:"#A78BFA" },
  "ì‹ì‚¬/ê°„ì‹": { name:"ì‹ë¹„",   icon:"ðŸš", color:"#F97316" },
  "ì‹ìž¬ë£Œ":    { name:"ì‹ë¹„",   icon:"ðŸ›’", color:"#F97316" },
  "ì˜ë¥˜/ìž¡í™”": { name:"ì‡¼í•‘",   icon:"ðŸ›ï¸", color:"#EC4899" },
  "ì „ìžì œí’ˆ":  { name:"ì‡¼í•‘",   icon:"ðŸ“±", color:"#EC4899" },
  "í™”ìž¥í’ˆ":    { name:"ë·°í‹°",   icon:"ðŸ’„", color:"#06B6D4" },
  "ë³‘ì›ë¹„/ì•½ê°’":{ name:"ì˜ë£Œ",  icon:"ðŸ¥", color:"#EF4444" },
  "ì˜í™”/ê³µì—°": { name:"ë¬¸í™”",   icon:"ðŸŽ¬", color:"#F59E0B" },
  "ì—¬í–‰":      { name:"ë¬¸í™”",   icon:"âœˆï¸", color:"#F59E0B" },
  "ë„ì„œ":      { name:"ë¬¸í™”",   icon:"ðŸ“š", color:"#F59E0B" },
  "íƒì‹œë¹„":    { name:"êµí†µ",   icon:"ðŸš•", color:"#3B82F6" },
  "ëŒ€ì¤‘êµí†µ":  { name:"êµí†µ",   icon:"ðŸšŒ", color:"#3B82F6" },
  "ìœ ë¥˜ë¹„":    { name:"êµí†µ",   icon:"â›½", color:"#3B82F6" },
  "ì£¼ì°¨/í†µí–‰": { name:"êµí†µ",   icon:"ðŸš—", color:"#3B82F6" },
  "êµìž¬ë¹„":    { name:"ê¸°íƒ€",   icon:"ðŸ“š", color:"#6366F1" },
  "í†µì‹ ë¹„":    { name:"ê¸°íƒ€",   icon:"ðŸ“¡", color:"#94A3B8" },
  "ìƒí™œì„¸ê¸ˆ":  { name:"ê¸°íƒ€",   icon:"ðŸ›ï¸", color:"#94A3B8" },
  "ìžë™ì°¨ë³´í—˜":{ name:"ê¸°íƒ€",   icon:"ðŸš—", color:"#94A3B8" },
  "ë³´ìž¥ë³´í—˜":  { name:"ê¸°íƒ€",   icon:"ðŸ›¡ï¸", color:"#94A3B8" },
};
const CAT_MAP = {
  "ì‹ìžìž¬":    { name:"ì‹ë¹„",   icon:"ðŸš", color:"#F97316" },
  "ì‡¼í•‘ë¹„":    { name:"ì‡¼í•‘",   icon:"ðŸ›ï¸", color:"#EC4899" },
  "êµí†µë¹„":    { name:"êµí†µ",   icon:"ðŸšŒ", color:"#3B82F6" },
  "ì°¨ëŸ‰ìœ ì§€ë¹„":{ name:"êµí†µ",   icon:"ðŸšŒ", color:"#3B82F6" },
  "ê±´ê°•ê´€ë¦¬ë¹„":{ name:"ì˜ë£Œ",   icon:"ðŸ¥", color:"#EF4444" },
  "ë¬¸í™”ìƒí™œë¹„":{ name:"ë¬¸í™”",   icon:"ðŸŽ¬", color:"#F59E0B" },
  "êµìœ¡ë¹„":    { name:"ê¸°íƒ€",   icon:"ðŸ“š", color:"#6366F1" },
  "ë¯¸ìš©ë¹„":    { name:"ë·°í‹°",   icon:"ðŸŒ¿", color:"#06B6D4" },
  "ìƒí™œë¹„":    { name:"ê¸°íƒ€",   icon:"ðŸ ", color:"#94A3B8" },
  "ê¸ˆìœµë³´í—˜ë¹„":{ name:"ê¸°íƒ€",   icon:"ðŸ’°", color:"#94A3B8" },
  "ì—†ìŒ":      { name:"ê¸°íƒ€",   icon:"ðŸ“¦", color:"#94A3B8" },
};

const getMeta = (subcat, cat) =>
  SUBCAT_MAP[subcat] || CAT_MAP[cat] || { name:"ê¸°íƒ€", icon:"ðŸ“¦", color:"#94A3B8" };

// XLS íŒŒì‹±
const parseXls = (buffer) => {
  const wb = XLSX.read(buffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const headerIdx = raw.findIndex(row => row.includes("ì§€ì¶œì¼"));
  if (headerIdx === -1) throw new Error("ë˜‘ë˜‘ê°€ê³„ë¶€ í˜•ì‹ì´ ì•„ë‹ˆì—ìš”");
  const headers = raw[headerIdx];
  const dataRows = raw.slice(headerIdx + 1).filter(row =>
    row[0] && String(row[0]).includes(".")
  );
  return dataRows.map((row, i) => {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j]; });
    const meta = getMeta(String(obj["ì„¸ë¶€ì¹´í…Œê³ ë¦¬"]||""), String(obj["ì¹´í…Œê³ ë¦¬"]||""));
    return {
      id: Date.now() + i,
      name: String(obj["ì§€ì¶œë‚´ì—­"] || ""),
      amount: -Math.abs(parseInt(obj["ì§€ì¶œê¸ˆì•¡"]) || 0),
      date: String(obj["ì§€ì¶œì¼"]||"").split(" ")[0],
      time: String(obj["ì§€ì¶œì‹œê°„"]||""),
      card: String(obj["ê²°ì œ"]||"-"),
      category: meta.name,
      icon: meta.icon,
      color: meta.color,
      memo: String(obj["ë©”ëª¨"]||""),
      author: 1,
      subcat: String(obj["ì„¸ë¶€ì¹´í…Œê³ ë¦¬"]||""),
    };
  });
};

// ë‚ ì§œ í¬ë§·
const fmtDate = (d) => {
  if (!d) return "";
  const dt = new Date(d);
  return isNaN(dt) ? String(d) : dt.toLocaleDateString("ko-KR");
};

// â”€â”€ ë©”ì¸ â”€â”€
const ExcelImportScreen = ({ onImport }) => {
  // í´ë” í•¸ë“¤ (File System Access API)
  const [dirHandle, setDirHandle] = useState(null);
  const [latestFile, setLatestFile] = useState(null); // { name, handle, lastModified }
  const [scanning, setScanning] = useState(false);
  const [step, setStep] = useState("idle"); // idle | preview | done
  const [parsed, setParsed] = useState([]);
  const [selected, setSelected] = useState([]);
  const [filterCat, setFilterCat] = useState(null);
  const [error, setError] = useState(null);
  const fileRef = useRef();

  // File System Access API ì§€ì› ì—¬ë¶€
  const fsa = typeof window !== "undefined" && "showDirectoryPicker" in window;

  // í´ë”ì—ì„œ ìµœì‹  xls íŒŒì¼ ì°¾ê¸°
  const scanDir = async (handle) => {
    setScanning(true);
    setError(null);
    try {
      let latest = null;
      for await (const [name, fh] of handle.entries()) {
        if (fh.kind !== "file") continue;
        if (!name.match(/\.(xls|xlsx)$/i)) continue;
        const file = await fh.getFile();
        if (!latest || file.lastModified > latest.lastModified) {
          latest = { name, handle: fh, lastModified: file.lastModified, file };
        }
      }
      if (!latest) throw new Error("í´ë”ì— ì—‘ì…€ íŒŒì¼ì´ ì—†ì–´ìš”");
      setLatestFile(latest);
    } catch (e) {
      if (e.name !== "AbortError") setError(e.message || "í´ë”ë¥¼ ì½ì„ ìˆ˜ ì—†ì–´ìš”");
    }
    setScanning(false);
  };

  // í´ë” ì„ íƒ
  const handleSelectDir = async () => {
    try {
      const handle = await window.showDirectoryPicker({ mode: "read" });
      setDirHandle(handle);
      setLatestFile(null);
      setParsed([]);
      setStep("idle");
      await scanDir(handle);
    } catch (e) {
      if (e.name !== "AbortError") setError("í´ë” ì„ íƒ ì·¨ì†Œë¨");
    }
  };

  // ìƒˆë¡œê³ ì¹¨ (ìµœì‹  íŒŒì¼ ë‹¤ì‹œ ìŠ¤ìº”)
  const handleRefresh = async () => {
    if (!dirHandle) return;
    setLatestFile(null);
    setParsed([]);
    setStep("idle");
    await scanDir(dirHandle);
  };

  // íŒŒì¼ ë¶ˆëŸ¬ì˜¤ê¸° (FSA)
  const handleLoadLatest = async () => {
    if (!latestFile) return;
    setError(null);
    try {
      const file = latestFile.file || await latestFile.handle.getFile();
      const buffer = await file.arrayBuffer();
      const rows = parseXls(new Uint8Array(buffer));
      if (rows.length === 0) throw new Error("ë°ì´í„°ê°€ ì—†ì–´ìš”");
      setParsed(rows);
      setSelected(rows.map(r => r.id));
      setStep("preview");
    } catch (e) {
      setError(e.message || "íŒŒì¼ì„ ì½ì„ ìˆ˜ ì—†ì–´ìš”");
    }
  };

  // ìˆ˜ë™ íŒŒì¼ ì„ íƒ (FSA ë¯¸ì§€ì› fallback)
  const handleManualFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const rows = parseXls(new Uint8Array(ev.target.result));
        if (rows.length === 0) throw new Error("ë°ì´í„°ê°€ ì—†ì–´ìš”");
        setLatestFile({ name: file.name, lastModified: file.lastModified });
        setParsed(rows);
        setSelected(rows.map(r => r.id));
        setStep("preview");
      } catch (e) {
        setError(e.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const toggle = (id) => setSelected(p => p.includes(id) ? p.filter(i=>i!==id) : [...p,id]);

  const handleConfirm = () => {
    onImport?.(parsed.filter(r => selected.includes(r.id)));
    setStep("done");
    setTimeout(() => { setStep("idle"); setParsed([]); setSelected([]); }, 2000);
  };

  const categories = [...new Set(parsed.map(r => r.category))];
  const filtered = filterCat ? parsed.filter(r=>r.category===filterCat) : parsed;
  const totalAmt = parsed.filter(r=>selected.includes(r.id)).reduce((s,r)=>s+Math.abs(r.amount),0);
  const catStats = parsed.reduce((acc,r) => {
    if (!acc[r.category]) acc[r.category] = {count:0,amount:0,icon:r.icon,color:r.color};
    acc[r.category].count++;
    acc[r.category].amount += Math.abs(r.amount);
    return acc;
  }, {});

  return (
    <div style={{
      minHeight:"100vh", background:"#F8FAFC",
      fontFamily:"'Apple SD Gothic Neo','Noto Sans KR',sans-serif",
      maxWidth:"430px", margin:"0 auto",
    }}>
      {/* í—¤ë” */}
      <div style={{
        background:"linear-gradient(135deg,#10B981 0%,#059669 100%)",
        padding:"48px 24px 24px", color:"white",
      }}>
        <h1 style={{margin:"0 0 4px",fontSize:"20px",fontWeight:"700"}}>
          ðŸ“Š ë˜‘ë˜‘ê°€ê³„ë¶€ ë¶ˆëŸ¬ì˜¤ê¸°
        </h1>
        <p style={{margin:0,fontSize:"13px",opacity:0.85}}>
          í´ë”ë¥¼ í•œ ë²ˆ ì§€ì •í•˜ë©´ ìµœì‹  íŒŒì¼ì„ ìžë™ìœ¼ë¡œ ì°¾ì•„ìš”
        </p>
      </div>

      <div style={{padding:"20px 16px 100px"}}>

        {/* â”€â”€ í´ë” ì§€ì • ì˜ì—­ â”€â”€ */}
        <div style={{
          background:"white", borderRadius:"16px", padding:"16px",
          boxShadow:"0 1px 6px rgba(0,0,0,0.06)", marginBottom:"16px",
        }}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
            <p style={{margin:0,fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>
              ðŸ“ ë‹¤ìš´ë¡œë“œ í´ë” ì—°ê²°
            </p>
            {dirHandle && (
              <button onClick={handleRefresh} disabled={scanning} style={{
                background:"none",border:"none",cursor:"pointer",
                fontSize:"12px",color:"#10B981",fontWeight:"600",
              }}>{scanning?"ìŠ¤ìº” ì¤‘...":"ðŸ”„ ìƒˆë¡œê³ ì¹¨"}</button>
            )}
          </div>

          {fsa ? (
            <>
              {/* í´ë” ì„ íƒ ë²„íŠ¼ */}
              <button onClick={handleSelectDir} style={{
                width:"100%", padding:"14px",
                border:`2px dashed ${dirHandle?"#A7F3D0":"#CBD5E1"}`,
                borderRadius:"14px",
                background:dirHandle?"#F0FDF4":"#F8FAFC",
                cursor:"pointer", display:"flex", alignItems:"center", gap:"12px",
                marginBottom: latestFile ? "12px" : "0",
              }}>
                <span style={{fontSize:"28px"}}>{dirHandle?"ðŸ“‚":"ðŸ“"}</span>
                <div style={{textAlign:"left"}}>
                  <p style={{margin:0,fontSize:"13px",fontWeight:"700",color:dirHandle?"#065F46":"#64748B"}}>
                    {dirHandle ? `ðŸ“‚ ${dirHandle.name}` : "í´ë” ì„ íƒí•˜ê¸°"}
                  </p>
                  <p style={{margin:"2px 0 0",fontSize:"11px",color:dirHandle?"#10B981":"#94A3B8"}}>
                    {dirHandle ? "í´ë¦­í•´ì„œ í´ë” ë³€ê²½" : "ë˜‘ë˜‘ê°€ê³„ë¶€ ì—‘ì…€ì´ ì €ìž¥ë˜ëŠ” í´ë”"}
                  </p>
                </div>
              </button>

              {/* ìŠ¤ìº” ì¤‘ */}
              {scanning && (
                <div style={{
                  background:"#F0FDF4",borderRadius:"12px",padding:"12px 14px",
                  display:"flex",alignItems:"center",gap:"10px",marginTop:"10px",
                }}>
                  <span style={{fontSize:"16px"}}>ðŸ”</span>
                  <p style={{margin:0,fontSize:"13px",color:"#065F46"}}>ìµœì‹  íŒŒì¼ ì°¾ëŠ” ì¤‘...</p>
                </div>
              )}

              {/* ìµœì‹  íŒŒì¼ ê°ì§€ë¨ */}
              {latestFile && !scanning && step === "idle" && (
                <div style={{
                  background:"linear-gradient(135deg,#F0FDF4,#ECFDF5)",
                  border:"2px solid #A7F3D0",
                  borderRadius:"14px", padding:"14px 16px",
                  marginTop:"10px",
                }}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"12px"}}>
                    <div>
                      <p style={{margin:0,fontSize:"11px",color:"#059669",fontWeight:"600"}}>
                        âœ¨ ìµœì‹  íŒŒì¼ ë°œê²¬
                      </p>
                      <p style={{margin:"4px 0 0",fontSize:"14px",fontWeight:"700",color:"#065F46"}}>
                        {latestFile.name}
                      </p>
                      <p style={{margin:"2px 0 0",fontSize:"11px",color:"#6EE7B7"}}>
                        {new Date(latestFile.lastModified).toLocaleString("ko-KR")}
                      </p>
                    </div>
                    <span style={{fontSize:"32px"}}>ðŸ“„</span>
                  </div>
                  <button onClick={handleLoadLatest} style={{
                    width:"100%", padding:"12px",
                    border:"none", borderRadius:"12px",
                    background:"linear-gradient(135deg,#10B981,#059669)",
                    color:"white", fontSize:"14px", fontWeight:"700",
                    cursor:"pointer", display:"flex", alignItems:"center",
                    justifyContent:"center", gap:"8px",
                  }}>
                    <span style={{fontSize:"16px"}}>â¬‡ï¸</span>
                    ì´ íŒŒì¼ ë¶ˆëŸ¬ì˜¤ê¸°
                  </button>
                </div>
              )}
            </>
          ) : (
            /* FSA ë¯¸ì§€ì› - ìˆ˜ë™ íŒŒì¼ ì„ íƒ */
            <>
              <div style={{
                background:"#FFF7ED",borderRadius:"10px",padding:"10px 12px",
                border:"1px solid #FFEDD5",marginBottom:"10px",
              }}>
                <p style={{margin:0,fontSize:"11px",color:"#92400E"}}>
                  ðŸ’¡ ì´ ë¸Œë¼ìš°ì €ëŠ” í´ë” ìžë™ ê°ì§€ë¥¼ ì§€ì›í•˜ì§€ ì•Šì•„ìš”. íŒŒì¼ì„ ì§ì ‘ ì„ íƒí•´ì£¼ì„¸ìš”.
                </p>
              </div>
              <input ref={fileRef} type="file" accept=".xls,.xlsx" onChange={handleManualFile} style={{display:"none"}}/>
              <button onClick={()=>fileRef.current.click()} style={{
                width:"100%",padding:"14px",border:"2px dashed #CBD5E1",
                borderRadius:"14px",background:"#F8FAFC",cursor:"pointer",
                display:"flex",alignItems:"center",justifyContent:"center",gap:"10px",
              }}>
                <span style={{fontSize:"20px"}}>ðŸ“‚</span>
                <span style={{fontSize:"14px",fontWeight:"700",color:"#64748B"}}>ì—‘ì…€ íŒŒì¼ ì§ì ‘ ì„ íƒ</span>
              </button>
            </>
          )}
        </div>

        {/* ì—ëŸ¬ */}
        {error && (
          <div style={{
            background:"#FEE2E2",borderRadius:"12px",padding:"12px 14px",
            border:"1px solid #FECACA",display:"flex",gap:"8px",marginBottom:"12px",
          }}>
            <span>âš ï¸</span>
            <p style={{margin:0,fontSize:"13px",color:"#991B1B"}}>{error}</p>
          </div>
        )}

        {/* ì‚¬ìš© ì•ˆë‚´ (í´ë” ë¯¸ì„ íƒ ì‹œ) */}
        {!dirHandle && !latestFile && step === "idle" && (
          <div style={{
            background:"white", borderRadius:"16px", padding:"16px",
            boxShadow:"0 1px 6px rgba(0,0,0,0.06)",
          }}>
            <p style={{margin:"0 0 12px",fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>ðŸ“‹ ì‚¬ìš© ë°©ë²•</p>
            {[
              ["1ï¸âƒ£","ë˜‘ë˜‘ê°€ê³„ë¶€ ì•± â†’ ë‚´ë³´ë‚´ê¸° â†’ ì—‘ì…€ ì €ìž¥"],
              ["2ï¸âƒ£","ì €ìž¥ í´ë”(ë³´í†µ ë‹¤ìš´ë¡œë“œ)ë¥¼ ìœ„ì—ì„œ ì„ íƒ"],
              ["3ï¸âƒ£","ì•±ì´ ìµœì‹  íŒŒì¼ì„ ìžë™ìœ¼ë¡œ ì°¾ì•„ì¤˜ìš”"],
              ["4ï¸âƒ£","ë¶ˆëŸ¬ì˜¤ê¸° ë²„íŠ¼ í•œ ë²ˆë§Œ ëˆ„ë¥´ë©´ ì™„ë£Œ!"],
            ].map(([n,t])=>(
              <div key={n} style={{display:"flex",gap:"10px",marginBottom:"8px",alignItems:"flex-start"}}>
                <span style={{fontSize:"16px",flexShrink:0}}>{n}</span>
                <span style={{fontSize:"13px",color:"#334155",lineHeight:"1.5"}}>{t}</span>
              </div>
            ))}
            <div style={{
              marginTop:"12px",background:"#EEF2FF",borderRadius:"10px",
              padding:"10px 12px",border:"1px solid #C7D2FE",
            }}>
              <p style={{margin:0,fontSize:"11px",color:"#4338CA",lineHeight:"1.6"}}>
                ðŸ’¡ í´ë”ëŠ” í•œ ë²ˆë§Œ ì§€ì •í•˜ë©´ ë¼ìš”. ë‹¤ìŒì— ì•±ì„ ì—´ì–´ë„ ê°™ì€ í´ë”ì—ì„œ ìµœì‹  íŒŒì¼ì„ ìžë™ìœ¼ë¡œ ì°¾ì•„ìš”.
              </p>
            </div>
          </div>
        )}

        {/* â”€â”€ ë¯¸ë¦¬ë³´ê¸° â”€â”€ */}
        {step === "preview" && (
          <>
            {/* ìš”ì•½ ì¹´ë“œ */}
            <div style={{
              background:"white",borderRadius:"16px",padding:"14px 16px",
              boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px",
            }}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
                <p style={{margin:0,fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>
                  {latestFile?.name || "ë¶ˆëŸ¬ì˜¨ íŒŒì¼"}
                </p>
                <button onClick={()=>{setStep("idle");setParsed([]);setSelected([]);}} style={{
                  background:"none",border:"none",cursor:"pointer",
                  fontSize:"12px",color:"#10B981",fontWeight:"600",
                }}>ë‹¤ì‹œ ì„ íƒ</button>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px"}}>
                {[
                  {l:"ì´ ê±´ìˆ˜",v:`${parsed.length}ê±´`},
                  {l:"ì„ íƒ",v:`${selected.length}ê±´`},
                  {l:"í•©ê³„",v:`${Math.round(totalAmt/10000)}ë§Œì›`},
                ].map(({l,v})=>(
                  <div key={l} style={{background:"#F0FDF4",borderRadius:"10px",padding:"10px",textAlign:"center",border:"1px solid #D1FAE5"}}>
                    <p style={{margin:0,fontSize:"10px",color:"#059669"}}>{l}</p>
                    <p style={{margin:"3px 0 0",fontSize:"15px",fontWeight:"700",color:"#065F46"}}>{v}</p>
                  </div>
                ))}
              </div>
            </div>

            {/* ì¹´í…Œê³ ë¦¬ í†µê³„ */}
            <div style={{
              background:"white",borderRadius:"16px",padding:"14px 16px",
              boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px",
            }}>
              <p style={{margin:"0 0 10px",fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>ì¹´í…Œê³ ë¦¬ë³„ ìš”ì•½</p>
              {Object.entries(catStats).sort((a,b)=>b[1].amount-a[1].amount).map(([cat,s])=>{
                const max = Math.max(...Object.values(catStats).map(v=>v.amount));
                const pct = Math.round((s.amount/max)*100);
                return (
                  <div key={cat} style={{marginBottom:"8px"}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:"3px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:"5px"}}>
                        <span style={{fontSize:"13px"}}>{s.icon}</span>
                        <span style={{fontSize:"12px",color:"#334155"}}>{cat}</span>
                        <span style={{fontSize:"10px",color:"#94A3B8"}}>{s.count}ê±´</span>
                      </div>
                      <span style={{fontSize:"12px",fontWeight:"700",color:"#1E293B"}}>{s.amount.toLocaleString()}ì›</span>
                    </div>
                    <div style={{background:"#F1F5F9",borderRadius:"99px",height:"5px"}}>
                      <div style={{width:`${pct}%`,height:"100%",background:s.color,borderRadius:"99px"}}/>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* ì¹´í…Œê³ ë¦¬ í•„í„° */}
            <div style={{display:"flex",gap:"6px",overflowX:"auto",paddingBottom:"4px",marginBottom:"10px"}}>
              <button onClick={()=>setFilterCat(null)} style={{
                padding:"5px 12px",border:"none",borderRadius:"99px",
                background:!filterCat?"#10B981":"#F1F5F9",
                color:!filterCat?"white":"#64748B",
                fontWeight:!filterCat?"700":"400",fontSize:"12px",flexShrink:0,cursor:"pointer",
              }}>ì „ì²´ {parsed.length}</button>
              {categories.map(cat=>{
                const s=catStats[cat];
                return (
                  <button key={cat} onClick={()=>setFilterCat(filterCat===cat?null:cat)} style={{
                    display:"flex",alignItems:"center",gap:"4px",padding:"5px 10px",
                    border:"none",borderRadius:"99px",
                    background:filterCat===cat?s.color:"#F1F5F9",
                    color:filterCat===cat?"white":"#64748B",
                    fontWeight:filterCat===cat?"700":"400",fontSize:"12px",flexShrink:0,cursor:"pointer",
                  }}>
                    <span>{s.icon}</span><span>{cat}</span>
                  </button>
                );
              })}
            </div>

            {/* ì „ì²´ì„ íƒ */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}>
              <p style={{margin:0,fontSize:"12px",color:"#64748B"}}>{filtered.length}ê±´</p>
              <button onClick={()=>setSelected(
                selected.length===parsed.length?[]:parsed.map(r=>r.id)
              )} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#10B981",fontWeight:"600"}}>
                {selected.length===parsed.length?"ì „ì²´í•´ì œ":"ì „ì²´ì„ íƒ"}
              </button>
            </div>

            {/* ë‚´ì—­ ë¦¬ìŠ¤íŠ¸ */}
            <div style={{display:"flex",flexDirection:"column",gap:"6px"}}>
              {filtered.map(row=>{
                const isSel=selected.includes(row.id);
                return (
                  <button key={row.id} onClick={()=>toggle(row.id)} style={{
                    background:isSel?"#F0FDF4":"white",
                    border:`2px solid ${isSel?"#10B981":"transparent"}`,
                    borderRadius:"12px",padding:"11px 13px",
                    display:"flex",alignItems:"center",gap:"10px",
                    cursor:"pointer",textAlign:"left",width:"100%",
                    boxShadow:"0 1px 4px rgba(0,0,0,0.05)",transition:"all 0.12s",
                  }}>
                    <div style={{
                      width:"20px",height:"20px",borderRadius:"5px",flexShrink:0,
                      background:isSel?"#10B981":"white",
                      border:`2px solid ${isSel?"#10B981":"#CBD5E1"}`,
                      display:"flex",alignItems:"center",justifyContent:"center",
                      color:"white",fontSize:"11px",fontWeight:"700",
                    }}>{isSel?"âœ“":""}</div>
                    <div style={{
                      width:"34px",height:"34px",borderRadius:"9px",flexShrink:0,
                      background:(row.color||"#94A3B8")+"20",
                      display:"flex",alignItems:"center",justifyContent:"center",fontSize:"17px",
                    }}>{row.icon}</div>
                    <div style={{flex:1,minWidth:0}}>
                      <p style={{margin:0,fontSize:"13px",fontWeight:"600",color:"#1E293B",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.name}</p>
                      <div style={{display:"flex",gap:"5px",marginTop:"2px",alignItems:"center"}}>
                        <span style={{fontSize:"10px",padding:"1px 5px",borderRadius:"99px",background:(row.color||"#94A3B8")+"20",color:row.color||"#94A3B8",fontWeight:"600",flexShrink:0}}>{row.category}</span>
                        <span style={{fontSize:"10px",color:"#94A3B8",flexShrink:0}}>{row.date}</span>
                        <span style={{fontSize:"10px",color:"#CBD5E1",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{row.card}</span>
                      </div>
                    </div>
                    <p style={{margin:0,fontSize:"13px",fontWeight:"700",color:"#1E293B",flexShrink:0}}>
                      {Math.abs(row.amount).toLocaleString()}ì›
                    </p>
                  </button>
                );
              })}
            </div>
          </>
        )}

        {/* â”€â”€ done â”€â”€ */}
        {step === "done" && (
          <div style={{textAlign:"center",padding:"80px 0"}}>
            <div style={{fontSize:"56px",marginBottom:"16px"}}>âœ…</div>
            <p style={{color:"#10B981",fontSize:"18px",fontWeight:"700",margin:"0 0 6px"}}>
              {selected.length}ê±´ ì¶”ê°€ëì–´ìš”!
            </p>
            <p style={{color:"#94A3B8",fontSize:"13px"}}>ê°€ê³„ë¶€ì— ë°˜ì˜ëì–´ìš”</p>
          </div>
        )}
      </div>

      {/* í•˜ë‹¨ í™•ì¸ ë²„íŠ¼ */}
      {step === "preview" && (
        <div style={{
          position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",
          width:"100%",maxWidth:"430px",background:"white",
          borderTop:"1px solid #F1F5F9",padding:"12px 16px 28px",
          boxShadow:"0 -4px 20px rgba(0,0,0,0.08)",
        }}>
          <button onClick={handleConfirm} disabled={selected.length===0} style={{
            width:"100%",padding:"15px",border:"none",borderRadius:"14px",
            background:selected.length>0
              ?"linear-gradient(135deg,#10B981,#059669)"
              :"#CBD5E1",
            color:"white",fontSize:"15px",fontWeight:"700",
            cursor:selected.length>0?"pointer":"not-allowed",
          }}>
            {selected.length>0
              ?`${selected.length}ê±´ ê°€ê³„ë¶€ì— ì¶”ê°€ (${Math.round(totalAmt/10000)}ë§Œì›)`
              :"í•­ëª©ì„ ì„ íƒí•´ì£¼ì„¸ìš”"}
          </button>
        </div>
      )}
    </div>
  );
}


export default function MoneyAppV2() {
  const [screen, setScreen] = useState("home");
  const [txList, setTxList] = useState(INIT_TX);
  const alertCount = MONTHLY_CATS.filter(c=>getLevel(c.amount,c.avg)!=="ok").length;

  const renderScreen = () => {
    switch(screen) {
      case "home":     return <HomeScreen setScreen={setScreen} txList={txList}/>;
      case "txlist":   return <TxListScreen txList={txList} setTxList={setTxList} onBack={()=>setScreen("home")}/>;
      case "excel":    return <ExcelImportScreen onImport={(rows)=>setTxList(p=>[...rows,...p])}/>;
      case "stats":    return <StatsScreen/>;
      case "alert":    return <AlertScreen/>;
      case "family":   return <FamilyScreen txList={txList} setTxList={setTxList}/>;
      case "settings": return <SettingsScreen/>;
      default:         return <HomeScreen setScreen={setScreen} txList={txList}/>;
    }
  };

  const showNav = screen !== "txlist"; // ì „ì²´ë‚´ì—­ í™”ë©´ì€ ìžì²´ ë’¤ë¡œê°€ê¸° ì‚¬ìš©

  return (
    <div style={{
      minHeight:"100vh", background:"#F8FAFC",
      fontFamily:"'Apple SD Gothic Neo','Noto Sans KR',sans-serif",
      maxWidth:"430px", margin:"0 auto",
    }}>
      <div style={{paddingBottom: showNav ? "72px" : "0"}}>
        {renderScreen()}
      </div>
      {showNav && <NavBar screen={screen} setScreen={setScreen} alertCount={alertCount}/>}
    </div>
  );
}
