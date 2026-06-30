import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ════════════════════════════════
// 공통 데이터
// ════════════════════════════════
const MONTHLY_TREND = [
  {month:"1월",income:4200000,expense:2850000},
  {month:"2월",income:4200000,expense:3120000},
  {month:"3월",income:4500000,expense:2680000},
  {month:"4월",income:4200000,expense:3450000},
  {month:"5월",income:4700000,expense:2990000},
  {month:"6월",income:4200000,expense:1103000},
];
const MONTHLY_CATS = [
  {name:"식비",  amount:387000,icon:"🍚",color:"#F97316",avg:280000},
  {name:"쇼핑",  amount:298000,icon:"🛍️",color:"#EC4899",avg:150000},
  {name:"카페",  amount:142500,icon:"☕", color:"#A78BFA",avg:95000 },
  {name:"교통",  amount:84200, icon:"🚌",color:"#3B82F6",avg:80000 },
  {name:"편의점",amount:68700, icon:"🏪",color:"#10B981",avg:55000 },
  {name:"기타",  amount:52300, icon:"📦",color:"#94A3B8",avg:60000 },
];
const CAT_TREND = {
  식비:  [280000,310000,265000,320000,295000,387000],
  쇼핑:  [150000,210000,130000,280000,160000,298000],
  카페:  [95000, 88000, 102000,115000,91000, 142500],
  교통:  [80000, 75000, 83000, 79000, 82000, 84200 ],
  편의점:[55000, 62000, 48000, 71000, 58000, 68700 ],
  기타:  [60000, 55000, 72000, 48000, 63000, 52300 ],
};
const WEEKLY_CARDS = [
  {name:"신한카드",amount:87500, color:"#3B82F6",emoji:"💳"},
  {name:"현대카드",amount:124000,color:"#8B5CF6",emoji:"💜"},
  {name:"삼성카드",amount:43200, color:"#10B981",emoji:"💚"},
];
const MONTHLY_CARDS = [
  {name:"신한카드",amount:342000,limit:500000,color:"#3B82F6"},
  {name:"현대카드",amount:487000,limit:600000,color:"#8B5CF6"},
  {name:"삼성카드",amount:198000,limit:400000,color:"#10B981"},
];
const FAMILY = {
  code:"MONEY-7X3K",
  members:[
    {id:1,name:"나",   role:"관리자",emoji:"👨",color:"#6366F1",joined:"2026.01.01"},
    {id:2,name:"배우자",role:"멤버",  emoji:"👩",color:"#EC4899",joined:"2026.01.03"},
    {id:3,name:"자녀1",role:"뷰어",  emoji:"👦",color:"#10B981",joined:"2026.03.15"},
  ],
};
const INIT_TX = [
  {id:1,name:"스타벅스",  category:"카페",  amount:-6500, date:"2026.06.19",time:"10:23",card:"신한카드",icon:"☕", author:1,memo:""},
  {id:2,name:"GS25",      category:"편의점",amount:-4200, date:"2026.06.19",time:"08:11",card:"현대카드",icon:"🏪", author:2,memo:""},
  {id:3,name:"당근마켓",  category:"수입",  amount:+35000,date:"2026.06.18",time:"19:45",card:"-",      icon:"🥕", author:1,memo:"중고 판매"},
  {id:4,name:"쿠팡",      category:"쇼핑",  amount:-29800,date:"2026.06.18",time:"14:22",card:"삼성카드",icon:"📦", author:2,memo:""},
  {id:5,name:"올리브영",  category:"뷰티",  amount:-18500,date:"2026.06.16",time:"17:45",card:"신한카드",icon:"🌿", author:1,memo:""},
  {id:6,name:"이마트24",  category:"편의점",amount:-3200, date:"2026.06.15",time:"09:02",card:"현대카드",icon:"🏪", author:3,memo:"학교 앞"},
  {id:7,name:"버거킹",    category:"식비",  amount:-12500,date:"2026.06.14",time:"13:30",card:"삼성카드",icon:"🍔", author:1,memo:""},
  {id:8,name:"카카오페이",category:"수입",  amount:+20000,date:"2026.06.13",time:"11:00",card:"-",      icon:"💛", author:2,memo:"용돈"},
];
const SMS_RAW = [
  {id:1,raw:"[신한카드] 06/19 10:23 스타벅스강남점 6,500원 승인", card:"신한카드",amount:6500, date:"06/19 10:23"},
  {id:2,raw:"[현대카드] 06/19 08:11 GS25상일동점 4,200원 승인",   card:"현대카드",amount:4200, date:"06/19 08:11"},
  {id:3,raw:"[삼성카드] 06/18 14:22 쿠팡 29,800원 승인",          card:"삼성카드",amount:29800,date:"06/18 14:22"},
  {id:4,raw:"[신한카드] 06/18 17:45 올리브영고덕점 18,500원 승인",card:"신한카드",amount:18500,date:"06/18 17:45"},
  {id:5,raw:"[현대카드] 06/17 09:02 이마트24 3,200원 승인",       card:"현대카드",amount:3200, date:"06/17 09:02"},
];
const CARD_PROVIDERS = [
  {id:"shinhan",name:"신한카드",prefix:"[신한카드]",sender:"15440092",emoji:"💙"},
  {id:"hyundai",name:"현대카드",prefix:"[현대카드]",sender:"15778000",emoji:"🖤"},
  {id:"samsung",name:"삼성카드",prefix:"[삼성카드]",sender:"15881700",emoji:"💙"},
  {id:"kb",     name:"KB국민카드",prefix:"[KB국민카드]",sender:"15881688",emoji:"💛"},
  {id:"lotte",  name:"롯데카드",prefix:"[롯데카드]",sender:"15771700",emoji:"❤️"},
  {id:"hana",   name:"하나카드",prefix:"[하나카드]",sender:"15991111",emoji:"💚"},
  {id:"kakao",  name:"카카오페이",prefix:"[카카오페이]",sender:"15994490",emoji:"💛"},
];
const CATEGORIES = [
  {name:"식비",icon:"🍚",color:"#F97316"},{name:"카페",icon:"☕",color:"#A78BFA"},
  {name:"쇼핑",icon:"🛍️",color:"#EC4899"},{name:"교통",icon:"🚌",color:"#3B82F6"},
  {name:"편의점",icon:"🏪",color:"#10B981"},{name:"의료",icon:"🏥",color:"#EF4444"},
  {name:"문화",icon:"🎬",color:"#F59E0B"},{name:"뷰티",icon:"🌿",color:"#06B6D4"},
  {name:"운동",icon:"💪",color:"#84CC16"},{name:"수입",icon:"💰",color:"#10B981"},
  {name:"기타",icon:"📦",color:"#94A3B8"},
];
const WARN_THR=120,ALERT_THR=150;
const getLevel=(c,a)=>{const p=(c/a)*100;return p>=ALERT_THR?"alert":p>=WARN_THR?"warn":"ok";};
const LVL={
  alert:{badge:"#EF4444",badgeBg:"#FEE2E2",border:"#FECACA",bg:"#FEE2E2",text:"#991B1B",icon:"🚨",label:"경고"},
  warn: {badge:"#F97316",badgeBg:"#FFF7ED",border:"#FFEDD5",bg:"#FFF7ED",text:"#92400E",icon:"⚠️",label:"주의"},
  ok:   {badge:"#10B981",badgeBg:"#DCFCE7",border:"#DCFCE7",bg:"#F0FDF4",text:"#14532D",icon:"✅",label:"정상"},
};
const fmt=(n)=>{const a=Math.abs(n).toLocaleString();return n<0?`-${a}원`:`+${a}원`;};
const parseMerchant=(raw)=>{const m=raw.match(/\d{2}:\d{2}\s(.+?)\s[\d,]+원/);return m?m[1]:"알 수 없음";};

// ════════════════════════════════
// 공통 UI
// ════════════════════════════════
const Toggle=({value,onChange})=>(
  <div onClick={()=>onChange(!value)} style={{width:"44px",height:"24px",borderRadius:"99px",background:value?"#6366F1":"#CBD5E1",position:"relative",cursor:"pointer",transition:"background 0.2s",flexShrink:0}}>
    <div style={{width:"20px",height:"20px",borderRadius:"50%",background:"white",position:"absolute",top:"2px",left:value?"22px":"2px",transition:"left 0.2s",boxShadow:"0 1px 3px rgba(0,0,0,0.2)"}}/>
  </div>
);

const NavBar=({screen,setScreen,alertCount})=>(
  <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:"430px",background:"white",borderTop:"1px solid #F1F5F9",display:"flex",justifyContent:"space-around",padding:"8px 0 16px",boxShadow:"0 -4px 20px rgba(0,0,0,0.06)",zIndex:100}}>
    {[{id:"home",icon:"🏠",label:"홈"},{id:"excel",icon:"📥",label:"엑셀"},{id:"stats",icon:"📊",label:"통계"},{id:"alert",icon:"🚨",label:"알림",badge:alertCount},{id:"family",icon:"👨‍👩‍👦",label:"가족"},{id:"settings",icon:"⚙️",label:"설정"}].map(({id,icon,label,badge})=>(
      <button key={id} onClick={()=>setScreen(id)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"2px",padding:"6px 8px",background:"none",border:"none",cursor:"pointer",color:screen===id?"#6366F1":"#94A3B8",position:"relative"}}>
        <span style={{fontSize:"18px"}}>{icon}</span>
        <span style={{fontSize:"9px",fontWeight:screen===id?"700":"400"}}>{label}</span>
        {badge>0&&<div style={{position:"absolute",top:"2px",right:"2px",width:"14px",height:"14px",borderRadius:"50%",background:"#EF4444",color:"white",fontSize:"8px",fontWeight:"700",display:"flex",alignItems:"center",justifyContent:"center"}}>{badge}</div>}
      </button>
    ))}
  </div>
);

// 차트들
const BarChartH=({data})=>{
  const max=Math.max(...data.map(d=>d.amount));
  return <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>{data.map(d=>{const pct=Math.round((d.amount/max)*100);return(<div key={d.name}><div style={{display:"flex",justifyContent:"space-between",marginBottom:"5px"}}><div style={{display:"flex",alignItems:"center",gap:"6px"}}><span style={{fontSize:"14px"}}>{d.icon}</span><span style={{fontSize:"13px",color:"#334155"}}>{d.name}</span></div><span style={{fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>{d.amount.toLocaleString()}원</span></div><div style={{background:"#F1F5F9",borderRadius:"99px",height:"8px"}}><div style={{width:`${pct}%`,height:"100%",background:d.color,borderRadius:"99px"}}/></div></div>);})}</div>;
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

// ════════════════════════════════
// 편집 모달
// ════════════════════════════════
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
          <h2 style={{margin:0,fontSize:"17px",fontWeight:"700",color:"#1E293B"}}>내역 편집</h2>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:"24px",color:"#94A3B8",lineHeight:1}}>×</button>
        </div>
        <div style={{display:"flex",background:"#F1F5F9",borderRadius:"12px",padding:"4px",marginBottom:"14px"}}>
          {[["expense","지출"],["income","수입"]].map(([k,l])=>(
            <button key={k} onClick={()=>setType(k)} style={{flex:1,padding:"8px",border:"none",cursor:"pointer",borderRadius:"10px",background:type===k?"white":"transparent",color:type===k?(k==="expense"?"#6366F1":"#10B981"):"#64748B",fontWeight:type===k?"700":"400",fontSize:"14px",transition:"all 0.2s"}}>{l}</button>
          ))}
        </div>
        {[
          {label:"상호명",content:<input value={name} onChange={e=>setName(e.target.value)} style={{width:"100%",padding:"11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"14px",color:"#1E293B",outline:"none",boxSizing:"border-box"}}/>},
          {label:"금액",content:<div style={{position:"relative"}}><input type="number" value={amount} onChange={e=>setAmount(e.target.value)} style={{width:"100%",padding:"11px 44px 11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"16px",fontWeight:"700",color:type==="expense"?"#6366F1":"#10B981",outline:"none",boxSizing:"border-box"}}/><span style={{position:"absolute",right:"14px",top:"50%",transform:"translateY(-50%)",color:"#94A3B8"}}>원</span></div>},
          {label:"날짜",content:<input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{width:"100%",padding:"11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"14px",color:"#6366F1",fontWeight:"600",outline:"none",boxSizing:"border-box"}}/>},
        ].map(({label,content})=>(
          <div key={label} style={{marginBottom:"12px"}}>
            <p style={{margin:"0 0 6px",fontSize:"12px",fontWeight:"600",color:"#64748B"}}>{label}</p>
            {content}
          </div>
        ))}
        <div style={{marginBottom:"12px"}}>
          <p style={{margin:"0 0 8px",fontSize:"12px",fontWeight:"600",color:"#64748B"}}>카테고리</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"6px"}}>
            {CATEGORIES.map(c=>{const iS=cat===c.name;return(<button key={c.name} onClick={()=>setCat(c.name)} style={{background:iS?c.color+"18":"#F8FAFC",border:`2px solid ${iS?c.color:"transparent"}`,borderRadius:"10px",padding:"8px 4px",cursor:"pointer",textAlign:"center"}}><div style={{fontSize:"18px",marginBottom:"2px"}}>{c.icon}</div><p style={{margin:0,fontSize:"10px",color:iS?c.color:"#64748B",fontWeight:iS?"700":"400"}}>{c.name}</p></button>);})}</div>
        </div>
        {[
          {label:"결제수단",content:<input value={card} onChange={e=>setCard(e.target.value)} placeholder="신한카드, 현금 등" style={{width:"100%",padding:"11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"14px",outline:"none",boxSizing:"border-box"}}/>},
          {label:"메모",content:<textarea value={memo} onChange={e=>setMemo(e.target.value)} rows={2} placeholder="메모 (선택)" style={{width:"100%",padding:"11px 14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"14px",outline:"none",resize:"none",boxSizing:"border-box",fontFamily:"inherit"}}/>},
        ].map(({label,content})=>(
          <div key={label} style={{marginBottom:"12px"}}>
            <p style={{margin:"0 0 6px",fontSize:"12px",fontWeight:"600",color:"#64748B"}}>{label}</p>
            {content}
          </div>
        ))}
        <div style={{display:"flex",gap:"10px",marginTop:"8px"}}>
          {!confirmDel?(
            <>
              <button onClick={()=>setConfirmDel(true)} style={{padding:"14px 18px",border:"2px solid #FEE2E2",borderRadius:"14px",background:"white",color:"#EF4444",fontSize:"14px",fontWeight:"700",cursor:"pointer"}}>🗑️</button>
              <button onClick={handleSave} style={{flex:1,padding:"14px",border:"none",borderRadius:"14px",background:"linear-gradient(135deg,#6366F1,#8B5CF6)",color:"white",fontSize:"15px",fontWeight:"700",cursor:"pointer"}}>저장하기</button>
            </>
          ):(
            <>
              <button onClick={()=>setConfirmDel(false)} style={{flex:1,padding:"14px",border:"2px solid #E2E8F0",borderRadius:"14px",background:"white",color:"#64748B",fontSize:"14px",fontWeight:"700",cursor:"pointer"}}>취소</button>
              <button onClick={()=>onDelete(tx.id)} style={{flex:1,padding:"14px",border:"none",borderRadius:"14px",background:"#EF4444",color:"white",fontSize:"14px",fontWeight:"700",cursor:"pointer"}}>정말 삭제</button>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

// 초대 모달
const InviteModal=({code,onClose})=>{
  const [copied,setCopied]=useState(false);
  const [joinCode,setJoinCode]=useState("");
  const [joined,setJoined]=useState(false);
  const [tab,setTab]=useState("invite");
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:"20px"}} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div style={{background:"white",borderRadius:"20px",width:"100%",maxWidth:"390px",padding:"24px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"16px"}}>
          <h2 style={{margin:0,fontSize:"17px",fontWeight:"700"}}>가족 초대</h2>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:"24px",color:"#94A3B8",lineHeight:1}}>×</button>
        </div>
        <div style={{display:"flex",background:"#F1F5F9",borderRadius:"10px",padding:"3px",marginBottom:"16px"}}>
          {[["invite","초대하기"],["join","코드 입력"]].map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"7px",border:"none",cursor:"pointer",borderRadius:"8px",background:tab===k?"white":"transparent",color:tab===k?"#6366F1":"#64748B",fontWeight:tab===k?"700":"400",fontSize:"13px"}}>{l}</button>
          ))}
        </div>
        {tab==="invite"&&(
          <>
            <div style={{background:"linear-gradient(135deg,#EEF2FF,#F5F3FF)",border:"2px dashed #A5B4FC",borderRadius:"16px",padding:"20px",textAlign:"center",marginBottom:"14px"}}>
              <p style={{margin:"0 0 6px",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>초대 코드</p>
              <p style={{margin:0,fontSize:"28px",fontWeight:"700",color:"#4338CA",letterSpacing:"4px"}}>{code}</p>
            </div>
            <div style={{display:"flex",gap:"8px",marginBottom:"12px"}}>
              <button onClick={()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);}} style={{flex:1,padding:"12px",border:"none",borderRadius:"12px",background:copied?"#10B981":"#6366F1",color:"white",fontSize:"13px",fontWeight:"700",cursor:"pointer",transition:"background 0.2s"}}>{copied?"✅ 복사됐어요":"📋 코드 복사"}</button>
              <button style={{flex:1,padding:"12px",border:"none",borderRadius:"12px",background:"#F9E000",color:"#1A1A1A",fontSize:"13px",fontWeight:"700",cursor:"pointer"}}>💬 카톡 공유</button>
            </div>
            <p style={{margin:0,fontSize:"12px",color:"#94A3B8",textAlign:"center"}}>코드 유효기간 7일 · 가족이 입력하면 즉시 연결</p>
          </>
        )}
        {tab==="join"&&(
          !joined?(
            <>
              <input value={joinCode} onChange={e=>setJoinCode(e.target.value.toUpperCase())} placeholder="MONEY-XXXX" maxLength={10} style={{width:"100%",padding:"14px",border:"2px solid #E2E8F0",borderRadius:"12px",fontSize:"18px",fontWeight:"700",color:"#4338CA",textAlign:"center",letterSpacing:"3px",outline:"none",boxSizing:"border-box",marginBottom:"12px"}}/>
              <button onClick={()=>{if(joinCode.length>=5)setJoined(true);}} style={{width:"100%",padding:"14px",border:"none",borderRadius:"12px",background:joinCode.length>=5?"linear-gradient(135deg,#6366F1,#8B5CF6)":"#CBD5E1",color:"white",fontSize:"15px",fontWeight:"700",cursor:joinCode.length>=5?"pointer":"not-allowed"}}>참여하기</button>
            </>
          ):(
            <div style={{textAlign:"center",padding:"20px 0"}}>
              <div style={{fontSize:"48px",marginBottom:"12px"}}>🎉</div>
              <p style={{fontSize:"16px",fontWeight:"700",color:"#1E293B",margin:"0 0 6px"}}>연결됐어요!</p>
              <p style={{fontSize:"13px",color:"#64748B",margin:"0 0 16px"}}>이제 가족과 함께 사용할 수 있어요</p>
              <button onClick={onClose} style={{padding:"12px 24px",border:"none",borderRadius:"12px",background:"linear-gradient(135deg,#6366F1,#8B5CF6)",color:"white",fontSize:"14px",fontWeight:"700",cursor:"pointer"}}>확인</button>
            </div>
          )
        )}
      </div>
    </div>
  );
};

// ════════════════════════════════
// 화면 1: 홈
// ════════════════════════════════
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
      const summary=MONTHLY_CATS.map(c=>`${c.name}:${c.amount.toLocaleString()}원`).join(", ");
      const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1000,messages:[{role:"user",content:`친근한 가계부 AI야. 이번달 지출: ${summary}. 이모지 써서 3줄로 분석해줘.\n1. 한줄 요약\n2. 주목할 점\n3. 절약 팁`}]})});
      const data=await res.json();
      setAiAnalysis(data.content?.[0]?.text||"분석 실패");
    }catch{setAiAnalysis("분석 중 오류가 발생했어요.");}
    setAiLoading(false);
  };

  const recent=txList.slice(0,5);

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 28px",color:"white"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"20px"}}>
          <div><p style={{fontSize:"13px",opacity:0.8,margin:0}}>2026년 6월</p><h1 style={{fontSize:"22px",fontWeight:"700",margin:"4px 0 0"}}>안녕하세요 👋</h1></div>
          <button onClick={()=>setScreen("excel")} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:"12px",padding:"8px 12px",color:"white",fontSize:"12px",fontWeight:"600",cursor:"pointer"}}>📥 엑셀 불러오기</button>
        </div>
        <div style={{background:"rgba(255,255,255,0.15)",borderRadius:"16px",padding:"16px 20px"}}>
          <p style={{margin:0,fontSize:"12px",opacity:0.8}}>이번 달 총 지출</p>
          <p style={{margin:"4px 0 0",fontSize:"28px",fontWeight:"700"}}>{totalMonthly.toLocaleString()}원</p>
        </div>
      </div>
      <div style={{padding:"20px 16px 100px"}}>
        {/* AI 분석 */}
        <div style={{background:aiAnalysis?"linear-gradient(135deg,#EEF2FF,#F5F3FF)":"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",border:aiAnalysis?"1px solid #C7D2FE":"1px solid #F1F5F9",marginBottom:"20px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:aiAnalysis?"12px":"0"}}>
            <div style={{display:"flex",alignItems:"center",gap:"8px"}}><span style={{fontSize:"18px"}}>🤖</span><span style={{fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>AI 지출 분석</span></div>
            <button onClick={handleAI} disabled={aiLoading} style={{background:aiLoading?"#E2E8F0":"linear-gradient(135deg,#6366F1,#8B5CF6)",border:"none",borderRadius:"10px",padding:"6px 12px",color:aiLoading?"#94A3B8":"white",fontSize:"12px",fontWeight:"600",cursor:aiLoading?"not-allowed":"pointer"}}>{aiLoading?"분석 중...":aiAnalysis?"다시 분석":"분석하기"}</button>
          </div>
          {aiLoading&&<p style={{margin:"8px 0 0",fontSize:"12px",color:"#94A3B8"}}>Claude가 분석하고 있어요...</p>}
          {aiAnalysis&&!aiLoading&&<div style={{background:"white",borderRadius:"12px",padding:"12px 14px",fontSize:"13px",color:"#334155",lineHeight:"1.8",whiteSpace:"pre-line"}}>{aiAnalysis}</div>}
          {!aiAnalysis&&!aiLoading&&<p style={{margin:"8px 0 0",fontSize:"12px",color:"#94A3B8"}}>이번 달 지출 패턴을 AI가 분석해드려요</p>}
        </div>

        {/* 탭 */}
        <div style={{display:"flex",background:"#E2E8F0",borderRadius:"12px",padding:"4px",marginBottom:"20px"}}>
          {[["week","이번 주"],["month","이번 달"]].map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"8px",border:"none",cursor:"pointer",borderRadius:"10px",background:tab===k?"white":"transparent",color:tab===k?"#6366F1":"#64748B",fontWeight:tab===k?"700":"400",fontSize:"14px",boxShadow:tab===k?"0 1px 4px rgba(0,0,0,0.1)":"none",transition:"all 0.2s"}}>{l}</button>
          ))}
        </div>

        {/* 카드별 지출 */}
        <div style={{marginBottom:"24px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
            <h2 style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>카드별 지출</h2>
            <span style={{fontSize:"12px",color:"#94A3B8"}}>합계 {(tab==="week"?totalWeekly:totalMonthly).toLocaleString()}원</span>
          </div>
          {tab==="week"?(
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {WEEKLY_CARDS.map(c=>(
                <div key={c.name} style={{background:"white",borderRadius:"14px",padding:"14px 16px",display:"flex",alignItems:"center",gap:"12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <div style={{width:"40px",height:"40px",borderRadius:"10px",background:c.color+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"18px"}}>{c.emoji}</div>
                  <div style={{flex:1}}><p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{c.name}</p><p style={{margin:"2px 0 0",fontSize:"12px",color:"#94A3B8"}}>이번 주</p></div>
                  <p style={{margin:0,fontSize:"15px",fontWeight:"700",color:c.color}}>{c.amount.toLocaleString()}원</p>
                </div>
              ))}
            </div>
          ):(
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {MONTHLY_CARDS.map(c=>{const pct=Math.round((c.amount/c.limit)*100);return(
                <div key={c.name} style={{background:"white",borderRadius:"14px",padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:"8px"}}><p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{c.name}</p><p style={{margin:0,fontSize:"14px",fontWeight:"700",color:c.color}}>{c.amount.toLocaleString()}원</p></div>
                  <div style={{background:"#F1F5F9",borderRadius:"99px",height:"6px"}}><div style={{width:`${pct}%`,height:"100%",background:pct>80?"#EF4444":c.color,borderRadius:"99px"}}/></div>
                  <div style={{display:"flex",justifyContent:"space-between",marginTop:"4px"}}><p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>한도 {c.limit.toLocaleString()}원</p><p style={{margin:0,fontSize:"11px",color:pct>80?"#EF4444":"#94A3B8"}}>{pct}%</p></div>
                </div>
              );})}
            </div>
          )}
        </div>

        {/* 카테고리 바 차트 */}
        <div style={{marginBottom:"24px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
            <h2 style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>월별 카테고리 지출</h2>
            <span style={{fontSize:"12px",color:"#94A3B8"}}>총 {totalCat.toLocaleString()}원</span>
          </div>
          <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"10px"}}><BarChartH data={MONTHLY_CATS}/></div>
          <div style={{background:"#FFF7ED",borderRadius:"12px",padding:"10px 14px",display:"flex",alignItems:"center",gap:"8px",border:"1px solid #FFEDD5"}}>
            <span style={{fontSize:"16px"}}>🔥</span>
            <p style={{margin:0,fontSize:"12px",color:"#92400E"}}>이번 달 <strong>{MONTHLY_CATS[0].name}</strong>에 가장 많이 썼어요 — <strong>{MONTHLY_CATS[0].amount.toLocaleString()}원</strong></p>
          </div>
        </div>

        {/* 최근 내역 */}
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
            <h2 style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>최근 내역</h2>
            <button onClick={()=>setScreen("txlist")} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>전체보기 →</button>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
            {recent.map(tx=>(
              <div key={tx.id} style={{background:"white",borderRadius:"14px",padding:"12px 16px",display:"flex",alignItems:"center",gap:"12px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                <div style={{width:"38px",height:"38px",borderRadius:"10px",background:tx.amount>0?"#DCFCE7":"#F1F5F9",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"18px",flexShrink:0}}>{tx.icon}</div>
                <div style={{flex:1,minWidth:0}}>
                  <p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{tx.name}</p>
                  <p style={{margin:"2px 0 0",fontSize:"11px",color:"#94A3B8"}}>{tx.category} · {tx.date}</p>
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

// ════════════════════════════════
// 화면 2: 전체 내역 (수정 포함)
// ════════════════════════════════
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

  // 날짜별 그룹
  const grouped=filtered.reduce((acc,tx)=>{
    if(!acc[tx.date])acc[tx.date]=[];
    acc[tx.date].push(tx);
    return acc;
  },{});

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 20px",color:"white"}}>
        <div style={{display:"flex",alignItems:"center",gap:"12px",marginBottom:"16px"}}>
          <button onClick={onBack} style={{background:"none",border:"none",color:"white",fontSize:"22px",cursor:"pointer",padding:0}}>←</button>
          <h1 style={{margin:0,fontSize:"18px",fontWeight:"700"}}>전체 내역</h1>
        </div>
        {/* 검색 */}
        <div style={{position:"relative"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="상호명, 카테고리, 메모 검색" style={{width:"100%",padding:"10px 14px 10px 36px",border:"none",borderRadius:"12px",fontSize:"13px",background:"rgba(255,255,255,0.2)",color:"white",outline:"none",boxSizing:"border-box"}}/>
          <span style={{position:"absolute",left:"12px",top:"50%",transform:"translateY(-50%)",fontSize:"14px",opacity:0.7}}>🔍</span>
        </div>
      </div>

      <div style={{padding:"0 0 100px"}}>
        {/* 요약 */}
        <div style={{display:"flex",gap:"0",background:"white",borderBottom:"1px solid #F1F5F9"}}>
          <div style={{flex:1,padding:"14px 16px",textAlign:"center",borderRight:"1px solid #F1F5F9"}}>
            <p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>총 지출</p>
            <p style={{margin:"4px 0 0",fontSize:"16px",fontWeight:"700",color:"#6366F1"}}>{totalExp.toLocaleString()}원</p>
          </div>
          <div style={{flex:1,padding:"14px 16px",textAlign:"center"}}>
            <p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>총 수입</p>
            <p style={{margin:"4px 0 0",fontSize:"16px",fontWeight:"700",color:"#10B981"}}>{totalInc.toLocaleString()}원</p>
          </div>
        </div>

        {/* 카테고리 필터 */}
        <div style={{display:"flex",gap:"6px",overflowX:"auto",padding:"12px 16px",background:"white",borderBottom:"1px solid #F1F5F9"}}>
          <button onClick={()=>setFilterCat(null)} style={{padding:"5px 12px",border:"none",borderRadius:"99px",background:!filterCat?"#6366F1":"#F1F5F9",color:!filterCat?"white":"#64748B",fontWeight:!filterCat?"700":"400",fontSize:"12px",flexShrink:0,cursor:"pointer"}}>전체</button>
          {CATEGORIES.map(c=>(
            <button key={c.name} onClick={()=>setFilterCat(filterCat===c.name?null:c.name)} style={{display:"flex",alignItems:"center",gap:"4px",padding:"5px 10px",border:"none",borderRadius:"99px",background:filterCat===c.name?c.color:"#F1F5F9",color:filterCat===c.name?"white":"#64748B",fontWeight:filterCat===c.name?"700":"400",fontSize:"12px",flexShrink:0,cursor:"pointer"}}>
              <span>{c.icon}</span><span>{c.name}</span>
            </button>
          ))}
        </div>

        {/* 날짜별 내역 */}
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
                          {tx.memo&&<span style={{fontSize:"10px",color:"#94A3B8"}}>· {tx.memo}</span>}
                        </div>
                      </div>
                      <div style={{textAlign:"right"}}>
                        <p style={{margin:0,fontSize:"14px",fontWeight:"700",color:tx.amount>0?"#10B981":"#1E293B"}}>{fmt(tx.amount)}</p>
                        {tx.card!=="-"&&<p style={{margin:"2px 0 0",fontSize:"10px",color:"#CBD5E1"}}>{tx.card}</p>}
                      </div>
                      <span style={{fontSize:"12px",color:"#E2E8F0"}}>✏️</span>
                    </button>
                  ))}
                </div>
              </div>
            );
          })}
          {filtered.length===0&&(
            <div style={{textAlign:"center",padding:"60px 0",color:"#94A3B8"}}>
              <div style={{fontSize:"40px",marginBottom:"8px"}}>📭</div>
              <p style={{margin:0}}>내역이 없어요</p>
            </div>
          )}
        </div>
      </div>

      {editTx&&<EditModal tx={editTx} onSave={handleSave} onDelete={handleDelete} onClose={()=>setEditTx(null)}/>}
    </div>
  );
};

// ════════════════════════════════
// 화면 3: SMS 불러오기
// ════════════════════════════════
const SmsScreen=({setTxList})=>{
  const [step,setStep]=useState("idle");
  const [items,setItems]=useState([]);
  const [selected,setSelected]=useState([]);
  const [done,setDone]=useState(false);

  const handleImport=async()=>{
    setStep("parsing");setItems([]);setSelected([]);
    await new Promise(r=>setTimeout(r,1200));
    const parsed=SMS_RAW.map(s=>({...s,merchant:parseMerchant(s),category:null,icon:"📦"}));
    setStep("classifying");
    try{
      const merchants=parsed.map(p=>p.merchant);
      const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-6",max_tokens:1000,messages:[{role:"user",content:`상호명으로 카테고리 분류. 카테고리: 식비,카페,쇼핑,교통,편의점,의료,문화,뷰티,운동,기타\nJSON 배열로만. 예:["카페"]\n상호명:\n${merchants.map((m,i)=>`${i+1}. ${m}`).join("\n")}`}]})});
      const data=await res.json();
      const cats=JSON.parse((data.content?.[0]?.text||"[]").replace(/```json|```/g,"").trim());
      const enriched=parsed.map((p,i)=>{
        const cat=cats[i]||"기타";
        const meta=CATEGORIES.find(c=>c.name===cat)||CATEGORIES.at(-1);
        return{...p,category:cat,icon:meta.icon,color:meta.color};
      });
      setItems(enriched);setSelected(enriched.map(e=>e.id));
    }catch{
      const fb=parsed.map(p=>({...p,category:"기타",icon:"📦",color:"#94A3B8"}));
      setItems(fb);setSelected(fb.map(e=>e.id));
    }
    setStep("done");
  };

  const toggle=(id)=>setSelected(p=>p.includes(id)?p.filter(i=>i!==id):[...p,id]);

  const handleConfirm=()=>{
    const toAdd=items.filter(it=>selected.includes(it.id)).map(it=>({
      id:Date.now()+it.id,name:it.merchant,category:it.category,
      amount:-it.amount,date:it.date.replace(/(\d{2})\/(\d{2})/,"2026.$1.$2"),
      time:it.date.split(" ")[1]||"",card:it.card,icon:it.icon||"📦",author:1,memo:"",
    }));
    setTxList(p=>[...toAdd,...p]);
    setDone(true);
    setTimeout(()=>{setDone(false);setStep("idle");setItems([]);setSelected([]);},1500);
  };

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 24px",color:"white"}}>
        <h1 style={{margin:"0 0 4px",fontSize:"20px",fontWeight:"700"}}>문자에서 불러오기 📩</h1>
        <p style={{margin:0,fontSize:"13px",opacity:0.75}}>AI가 카테고리를 자동으로 분류해요</p>
      </div>
      <div style={{padding:"20px 16px 100px"}}>
        {step==="idle"&&(
          <div style={{textAlign:"center",padding:"40px 0"}}>
            <div style={{fontSize:"56px",marginBottom:"16px"}}>📱</div>
            <p style={{color:"#1E293B",fontSize:"16px",fontWeight:"700",marginBottom:"8px"}}>카드 결제 문자 불러오기</p>
            <p style={{color:"#94A3B8",fontSize:"13px",marginBottom:"32px",lineHeight:"1.6"}}>최근 수신된 카드사 결제 문자를 읽어와<br/>AI가 카테고리를 자동으로 분류해드려요</p>
            <button onClick={handleImport} style={{background:"linear-gradient(135deg,#6366F1,#8B5CF6)",border:"none",borderRadius:"14px",padding:"14px 32px",color:"white",fontSize:"15px",fontWeight:"700",cursor:"pointer"}}>📩 문자 불러오기 시작</button>
          </div>
        )}
        {(step==="parsing"||step==="classifying")&&(
          <div style={{textAlign:"center",padding:"60px 0"}}>
            <div style={{fontSize:"40px",marginBottom:"20px"}}>{step==="parsing"?"📱":"🤖"}</div>
            <p style={{color:"#1E293B",fontSize:"15px",fontWeight:"700",marginBottom:"8px"}}>{step==="parsing"?"문자 내역 읽는 중...":"AI가 카테고리 분류 중..."}</p>
            <p style={{color:"#94A3B8",fontSize:"12px"}}>{step==="classifying"?"Claude가 상호명을 분석하고 있어요":"잠시만 기다려주세요"}</p>
          </div>
        )}
        {step==="done"&&!done&&(
          <>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
              <p style={{margin:0,fontSize:"13px",color:"#64748B"}}>{items.length}건 인식 · AI 카테고리 분류 완료 ✨</p>
              <button onClick={()=>setSelected(selected.length===items.length?[]:items.map(s=>s.id))} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>{selected.length===items.length?"전체해제":"전체선택"}</button>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {items.map(item=>{const iS=selected.includes(item.id);return(
                <button key={item.id} onClick={()=>toggle(item.id)} style={{background:iS?"#EEF2FF":"white",border:`2px solid ${iS?"#6366F1":"transparent"}`,borderRadius:"14px",padding:"14px 16px",display:"flex",alignItems:"center",gap:"12px",cursor:"pointer",textAlign:"left",width:"100%",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <div style={{width:"22px",height:"22px",borderRadius:"6px",background:iS?"#6366F1":"white",border:`2px solid ${iS?"#6366F1":"#CBD5E1"}`,display:"flex",alignItems:"center",justifyContent:"center",color:"white",fontSize:"13px",fontWeight:"700",flexShrink:0}}>{iS?"✓":""}</div>
                  <div style={{width:"40px",height:"40px",borderRadius:"10px",background:(item.color||"#94A3B8")+"20",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"20px",flexShrink:0}}>{item.icon}</div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",justifyContent:"space-between"}}>
                      <div>
                        <p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{item.merchant}</p>
                        <div style={{display:"flex",alignItems:"center",gap:"6px",marginTop:"3px"}}>
                          <span style={{fontSize:"10px",padding:"2px 6px",borderRadius:"99px",background:(item.color||"#94A3B8")+"20",color:item.color||"#94A3B8",fontWeight:"700"}}>🤖 {item.category}</span>
                          <span style={{fontSize:"11px",color:"#94A3B8"}}>{item.card} · {item.date}</span>
                        </div>
                      </div>
                      <p style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>{item.amount.toLocaleString()}원</p>
                    </div>
                    <div style={{marginTop:"8px",background:"#F8FAFC",borderRadius:"8px",padding:"6px 8px"}}><p style={{margin:0,fontSize:"10px",color:"#94A3B8",fontFamily:"monospace"}}>{item.raw}</p></div>
                  </div>
                </button>
              );})}
            </div>
          </>
        )}
        {done&&<div style={{textAlign:"center",padding:"60px 0"}}><div style={{fontSize:"50px",marginBottom:"16px"}}>✅</div><p style={{color:"#10B981",fontSize:"16px",fontWeight:"700"}}>{selected.length}건 가계부에 추가됐어요!</p></div>}
      </div>
      {step==="done"&&!done&&(
        <div style={{position:"fixed",bottom:"72px",left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:"430px",background:"white",borderTop:"1px solid #F1F5F9",padding:"12px 16px 8px",boxShadow:"0 -4px 20px rgba(0,0,0,0.08)"}}>
          <button onClick={handleConfirm} disabled={selected.length===0} style={{width:"100%",padding:"15px",border:"none",borderRadius:"14px",background:selected.length>0?"linear-gradient(135deg,#6366F1,#8B5CF6)":"#CBD5E1",color:"white",fontSize:"16px",fontWeight:"700",cursor:selected.length>0?"pointer":"not-allowed"}}>{selected.length>0?`${selected.length}건 가계부에 추가`:"항목을 선택해주세요"}</button>
        </div>
      )}
    </div>
  );
};

// ════════════════════════════════
// 화면 4: 통계
// ════════════════════════════════
const StatsScreen=()=>{
  const [tab,setTab]=useState("trend");
  const [selM,setSelM]=useState(5);
  const [selC,setSelC]=useState("식비");
  const months=["1월","2월","3월","4월","5월","6월"];
  const sel=MONTHLY_TREND[selM];
  const saving=sel.income-sel.expense;
  const sr=Math.round((saving/sel.income)*100);
  const avgExp=Math.round(MONTHLY_TREND.reduce((s,m)=>s+m.expense,0)/MONTHLY_TREND.length);
  const maxM=[...MONTHLY_TREND].sort((a,b)=>b.expense-a.expense)[0];

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 24px",color:"white"}}>
        <p style={{margin:0,fontSize:"13px",opacity:0.8}}>2026년 상반기</p>
        <h1 style={{fontSize:"20px",fontWeight:"700",margin:"4px 0 16px"}}>지출 통계 📊</h1>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px"}}>
          {[{l:"월평균 지출",v:`${Math.round(avgExp/10000)}만원`},{l:"최다 지출월",v:maxM.month},{l:"이번달 저축률",v:`${sr}%`}].map(({l,v})=>(
            <div key={l} style={{background:"rgba(255,255,255,0.15)",borderRadius:"12px",padding:"10px",textAlign:"center"}}><p style={{margin:0,fontSize:"9px",opacity:0.8}}>{l}</p><p style={{margin:"4px 0 0",fontSize:"16px",fontWeight:"700"}}>{v}</p></div>
          ))}
        </div>
      </div>
      <div style={{padding:"20px 16px 100px"}}>
        <div style={{display:"flex",background:"#E2E8F0",borderRadius:"12px",padding:"4px",marginBottom:"20px"}}>
          {[["trend","월별 추이"],["category","카테고리"]].map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"8px",border:"none",cursor:"pointer",borderRadius:"10px",background:tab===k?"white":"transparent",color:tab===k?"#6366F1":"#64748B",fontWeight:tab===k?"700":"400",fontSize:"14px",boxShadow:tab===k?"0 1px 4px rgba(0,0,0,0.1)":"none"}}>{l}</button>
          ))}
        </div>
        {tab==="trend"&&(
          <>
            <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px"}}>
              <div style={{display:"flex",gap:"16px",marginBottom:"12px"}}>
                {[["#10B981","수입"],["#6366F1","지출"]].map(([c,l])=>(
                  <div key={l} style={{display:"flex",alignItems:"center",gap:"5px"}}><div style={{width:"12px",height:"3px",background:c,borderRadius:"2px"}}/><span style={{fontSize:"11px",color:"#64748B"}}>{l}</span></div>
                ))}
              </div>
              <TrendChart data={MONTHLY_TREND} selectedIdx={selM} onSelect={setSelM}/>
            </div>
            <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"16px"}}>
              <h2 style={{margin:"0 0 14px",fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>{sel.month} 상세{selM===5&&<span style={{fontSize:"11px",color:"#6366F1",marginLeft:"6px"}}>진행중</span>}</h2>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"10px"}}>
                {[{l:"수입",v:sel.income,c:"#10B981"},{l:"지출",v:sel.expense,c:"#6366F1"},{l:saving>=0?"저축":"초과",v:Math.abs(saving),c:saving>=0?"#3B82F6":"#EF4444"}].map(({l,v,c})=>(
                  <div key={l} style={{background:c+"10",borderRadius:"12px",padding:"12px 10px",textAlign:"center",border:`1px solid ${c}20`}}><p style={{margin:0,fontSize:"11px",color:"#64748B"}}>{l}</p><p style={{margin:"4px 0 0",fontSize:"14px",fontWeight:"700",color:c}}>{Math.round(v/10000)}만</p></div>
                ))}
              </div>
              <div style={{marginTop:"14px"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"6px"}}><span style={{fontSize:"12px",color:"#64748B"}}>저축률</span><span style={{fontSize:"12px",fontWeight:"700",color:sr>=20?"#10B981":sr>=10?"#F97316":"#EF4444"}}>{sr}%</span></div>
                <div style={{background:"#F1F5F9",borderRadius:"99px",height:"8px"}}><div style={{width:`${Math.max(0,sr)}%`,height:"100%",background:sr>=20?"#10B981":sr>=10?"#F97316":"#EF4444",borderRadius:"99px"}}/></div>
              </div>
            </div>
            <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
              <h2 style={{margin:"0 0 14px",fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>월별 지출 비교</h2>
              {MONTHLY_TREND.map((m,i)=>{const pct=Math.round((m.expense/Math.max(...MONTHLY_TREND.map(d=>d.expense)))*100);const iS=i===selM;return(
                <div key={m.month} onClick={()=>setSelM(i)} style={{marginBottom:i<5?"10px":"0",cursor:"pointer"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:"4px"}}><span style={{fontSize:"12px",fontWeight:iS?"700":"400",color:iS?"#6366F1":"#64748B"}}>{m.month}</span><span style={{fontSize:"12px",fontWeight:"700",color:iS?"#6366F1":"#1E293B"}}>{m.expense.toLocaleString()}원</span></div>
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
                    <div><h2 style={{margin:0,fontSize:"15px",fontWeight:"700",color:"#1E293B"}}>{m.icon} {selC} 월별 추이</h2><p style={{margin:"4px 0 0",fontSize:"11px",color:"#94A3B8"}}>최근 6개월</p></div>
                    <div style={{textAlign:"right"}}><p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>이번달</p><p style={{margin:"2px 0 0",fontSize:"16px",fontWeight:"700",color:m.color}}>{CAT_TREND[selC].at(-1).toLocaleString()}원</p></div>
                  </div>
                  <CatChart data={CAT_TREND[selC]} color={m.color}/>
                  <div style={{display:"flex",justifyContent:"space-between",marginTop:"4px"}}>
                    {months.map((mo,i)=><span key={mo} style={{fontSize:"10px",color:i===months.length-1?m.color:"#94A3B8",fontWeight:i===months.length-1?"700":"400"}}>{mo}</span>)}
                  </div>
                </div>
                <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
                  <h2 style={{margin:"0 0 12px",fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>월별 상세</h2>
                  {CAT_TREND[selC].map((v,i)=>{const prev=i>0?CAT_TREND[selC][i-1]:null;const diff=prev?v-prev:0;const pct=Math.round((v/Math.max(...CAT_TREND[selC]))*100);const iL=i===CAT_TREND[selC].length-1;return(
                    <div key={i} style={{marginBottom:i<5?"10px":"0"}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"4px"}}>
                        <div style={{display:"flex",alignItems:"center",gap:"6px"}}><span style={{fontSize:"12px",fontWeight:iL?"700":"400",color:iL?m.color:"#64748B"}}>{months[i]}</span>{prev&&<span style={{fontSize:"10px",color:diff>0?"#EF4444":"#10B981",fontWeight:"600"}}>{diff>0?"▲":"▼"} {Math.abs(diff).toLocaleString()}원</span>}</div>
                        <span style={{fontSize:"12px",fontWeight:iL?"700":"400",color:iL?m.color:"#1E293B"}}>{v.toLocaleString()}원</span>
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

// ════════════════════════════════
// 화면 5: 알림
// ════════════════════════════════
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
          <div><p style={{margin:0,fontSize:"13px",opacity:0.8}}>2026년 6월 지출 알림</p><h1 style={{fontSize:"20px",fontWeight:"700",margin:"4px 0 0"}}>{alerts.length>0?`${alerts.length}개 카테고리 경고 🚨`:warns.length>0?`${warns.length}개 카테고리 주의 ⚠️`:"이번 달 지출 양호해요 ✅"}</h1></div>
          <div style={{background:"rgba(255,255,255,0.2)",borderRadius:"12px",padding:"8px 12px",textAlign:"center"}}><p style={{margin:0,fontSize:"22px",fontWeight:"700"}}>{issues.length}</p><p style={{margin:0,fontSize:"10px",opacity:0.8}}>주의 항목</p></div>
        </div>
        <div style={{display:"flex",gap:"8px",marginTop:"16px"}}>
          {[{level:"alert",count:alerts.length,label:"경고"},{level:"warn",count:warns.length,label:"주의"},{level:"ok",count:oks.length,label:"정상"}].map(({level,count,label})=>(
            <div key={level} style={{background:"rgba(255,255,255,0.2)",borderRadius:"10px",padding:"6px 12px",display:"flex",alignItems:"center",gap:"6px"}}><span style={{fontSize:"14px"}}>{LVL[level].icon}</span><span style={{fontSize:"13px",fontWeight:"700"}}>{count}</span><span style={{fontSize:"11px",opacity:0.8}}>{label}</span></div>
          ))}
        </div>
      </div>
      <div style={{padding:"20px 16px 100px"}}>
        {issues.length>0&&(
          <div style={{marginBottom:"24px"}}>
            <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>주의가 필요한 카테고리</p>
            <div style={{display:"flex",flexDirection:"column",gap:"10px"}}>
              {issues.map(cat=>{const s=LVL[cat.level];const iO=expanded===cat.name;return(
                <div key={cat.name} style={{background:"white",borderRadius:"16px",boxShadow:"0 2px 10px rgba(0,0,0,0.07)",border:`1px solid ${s.border}`,overflow:"hidden"}}>
                  <div onClick={()=>setExpanded(iO?null:cat.name)} style={{padding:"14px 16px",cursor:"pointer"}}>
                    <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
                      <div style={{width:"44px",height:"44px",borderRadius:"12px",background:cat.color+"20",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"22px",flexShrink:0,position:"relative"}}>
                        {cat.icon}
                        <div style={{position:"absolute",top:"-4px",right:"-4px",width:"16px",height:"16px",borderRadius:"50%",background:s.badge,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"9px"}}>{cat.level==="alert"?"🚨":"⚠️"}</div>
                      </div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:"6px",marginBottom:"4px"}}>
                          <span style={{fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>{cat.name}</span>
                          <span style={{fontSize:"10px",padding:"2px 7px",borderRadius:"99px",background:s.badgeBg,color:s.badge,fontWeight:"700",border:`1px solid ${s.border}`}}>{s.label} {cat.pct}%</span>
                        </div>
                        <div style={{background:"#F1F5F9",borderRadius:"99px",height:"6px",overflow:"hidden"}}><div style={{width:`${Math.min(cat.pct,100)}%`,height:"100%",background:s.badge,borderRadius:"99px"}}/></div>
                        <div style={{display:"flex",justifyContent:"space-between",marginTop:"4px"}}><span style={{fontSize:"11px",color:"#94A3B8"}}>평균 {cat.avg.toLocaleString()}원</span><span style={{fontSize:"11px",fontWeight:"700",color:s.badge}}>{cat.amount.toLocaleString()}원</span></div>
                      </div>
                      <span style={{fontSize:"16px",color:"#94A3B8"}}>{iO?"▲":"▼"}</span>
                    </div>
                  </div>
                  {iO&&(
                    <div style={{padding:"0 16px 14px",borderTop:`1px solid ${s.border}`}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingTop:"12px",marginBottom:"10px"}}>
                        <div>
                          <p style={{margin:"0 0 4px",fontSize:"12px",color:"#64748B"}}>최근 4개월 추이</p>
                          <div style={{display:"flex",gap:"12px"}}>
                            {["3달전","2달전","지난달","이번달"].map((l,i)=>(
                              <div key={l} style={{textAlign:"center"}}><p style={{margin:0,fontSize:"9px",color:"#94A3B8"}}>{l}</p><p style={{margin:"2px 0 0",fontSize:"10px",fontWeight:"600",color:i===3?s.badge:"#64748B"}}>{Math.round(cat.spark[i]/1000)}만</p></div>
                            ))}
                          </div>
                        </div>
                        <div style={{background:s.bg,borderRadius:"12px",padding:"10px 14px",border:`1px solid ${s.border}`,textAlign:"right"}}>
                          <p style={{margin:0,fontSize:"11px",color:s.text}}>평균 대비</p>
                          <p style={{margin:"2px 0 0",fontSize:"18px",fontWeight:"700",color:s.badge}}>+{cat.diff.toLocaleString()}원</p>
                          <p style={{margin:"2px 0 0",fontSize:"11px",color:s.text}}>{cat.pct-100}% 초과</p>
                        </div>
                      </div>
                      <div style={{background:"#F8FAFC",borderRadius:"10px",padding:"10px 12px",display:"flex",gap:"8px"}}>
                        <span style={{fontSize:"14px",flexShrink:0}}>🤖</span>
                        <p style={{margin:0,fontSize:"12px",color:"#334155",lineHeight:"1.6"}}>
                          {cat.name==="쇼핑"&&"이번 달 쇼핑이 평소보다 많이 늘었어요. 장바구니에 담아두고 3일 후 다시 확인하는 습관이 충동구매를 줄이는 데 도움이 돼요."}
                          {cat.name==="카페"&&"카페 지출이 평소보다 높아요. 홈카페로 전환하거나 텀블러 지참 시 할인되는 매장을 활용해보세요."}
                          {!["쇼핑","카페"].includes(cat.name)&&`${cat.name} 지출이 평소보다 ${cat.pct-100}% 높아요. 불필요한 항목이 없는지 점검해보세요.`}
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
            <p style={{margin:0,fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>정상 범위 카테고리</p>
            <button onClick={()=>setShowAll(p=>!p)} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>{showAll?"접기":"펼치기"}</button>
          </div>
          {showAll?(
            <div style={{display:"flex",flexDirection:"column",gap:"8px"}}>
              {oks.map(cat=>(
                <div key={cat.name} style={{background:"white",borderRadius:"14px",padding:"12px 16px",display:"flex",alignItems:"center",gap:"12px",boxShadow:"0 1px 4px rgba(0,0,0,0.05)",border:"1px solid #DCFCE7"}}>
                  <div style={{width:"38px",height:"38px",borderRadius:"10px",background:cat.color+"15",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"20px",flexShrink:0}}>{cat.icon}</div>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:"6px"}}><span style={{fontSize:"13px",fontWeight:"600",color:"#1E293B"}}>{cat.name}</span><span style={{fontSize:"10px",padding:"1px 6px",borderRadius:"99px",background:"#DCFCE7",color:"#15803D",fontWeight:"700"}}>✅ {cat.pct}%</span></div>
                    <div style={{background:"#F1F5F9",borderRadius:"99px",height:"4px",marginTop:"6px"}}><div style={{width:`${Math.min(cat.pct,100)}%`,height:"100%",background:"#10B981",borderRadius:"99px"}}/></div>
                  </div>
                  <div style={{textAlign:"right"}}><p style={{margin:0,fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>{cat.amount.toLocaleString()}원</p><p style={{margin:"2px 0 0",fontSize:"10px",color:"#94A3B8"}}>평균 {cat.avg.toLocaleString()}원</p></div>
                </div>
              ))}
            </div>
          ):(
            <div style={{background:"#F0FDF4",borderRadius:"12px",padding:"12px 16px",border:"1px solid #DCFCE7",display:"flex",alignItems:"center",gap:"10px"}}>
              <span style={{fontSize:"18px"}}>✅</span>
              <p style={{margin:0,fontSize:"13px",color:"#15803D"}}><strong>{oks.length}개</strong> 카테고리는 평균 범위 내 지출 중이에요</p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

// ════════════════════════════════
// 화면 6: 가족 공유
// ════════════════════════════════
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
          <div><p style={{margin:0,fontSize:"13px",opacity:0.8}}>우리집 가계부 👨‍👩‍👦</p><h1 style={{fontSize:"20px",fontWeight:"700",margin:"4px 0 0"}}>가족 공유</h1></div>
          <button onClick={()=>setShowInvite(true)} style={{background:"rgba(255,255,255,0.2)",border:"none",borderRadius:"12px",padding:"8px 12px",color:"white",fontSize:"12px",fontWeight:"600",cursor:"pointer"}}>👥 초대하기</button>
        </div>
        <div style={{display:"flex",gap:"8px",alignItems:"center"}}>
          {members.map(m=>(
            <button key={m.id} onClick={()=>setFilterAuthor(filterAuthor===m.id?null:m.id)} style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"4px",background:"none",border:"none",cursor:"pointer"}}>
              <div style={{width:"40px",height:"40px",borderRadius:"50%",background:filterAuthor===m.id?"white":m.color+"40",border:`2px solid ${filterAuthor===m.id?"white":"transparent"}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"20px"}}>{m.emoji}</div>
              <span style={{fontSize:"10px",opacity:filterAuthor===m.id?1:0.7,fontWeight:filterAuthor===m.id?"700":"400"}}>{m.name}</span>
            </button>
          ))}
          {filterAuthor&&<button onClick={()=>setFilterAuthor(null)} style={{marginLeft:"4px",background:"rgba(255,255,255,0.2)",border:"none",borderRadius:"99px",padding:"4px 10px",color:"white",fontSize:"11px",cursor:"pointer",fontWeight:"600"}}>전체보기</button>}
        </div>
      </div>
      <div style={{display:"flex",background:"#E2E8F0",padding:"4px"}}>
        {[["txlist","내역 목록"],["family","멤버 관리"]].map(([k,l])=>(
          <button key={k} onClick={()=>setTab(k)} style={{flex:1,padding:"10px",border:"none",cursor:"pointer",background:tab===k?"white":"transparent",color:tab===k?"#6366F1":"#64748B",fontWeight:tab===k?"700":"400",fontSize:"14px"}}>{l}</button>
        ))}
      </div>
      <div style={{padding:"16px 16px 100px"}}>
        {tab==="txlist"&&(
          <>
            <div style={{background:"white",borderRadius:"16px",padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><p style={{margin:0,fontSize:"12px",color:"#94A3B8"}}>이번달 가족 총 지출</p><p style={{margin:"4px 0 0",fontSize:"20px",fontWeight:"700",color:"#6366F1"}}>{totalExp.toLocaleString()}원</p></div>
              <div style={{display:"flex",gap:"6px"}}>
                {members.map(m=>{const mt=txList.filter(t=>t.author===m.id&&t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);const pct=totalExp>0?Math.round((mt/totalExp)*100):0;return(<div key={m.id} style={{textAlign:"center"}}><div style={{fontSize:"18px"}}>{m.emoji}</div><p style={{margin:"2px 0 0",fontSize:"10px",color:m.color,fontWeight:"700"}}>{pct}%</p></div>);})}
              </div>
            </div>
            <div style={{background:"white",borderRadius:"14px",padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"16px"}}>
              <p style={{margin:"0 0 10px",fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>멤버별 지출 비중</p>
              <div style={{display:"flex",height:"8px",borderRadius:"99px",overflow:"hidden",marginBottom:"10px"}}>
                {members.map(m=>{const mt=txList.filter(t=>t.author===m.id&&t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);const pct=totalExp>0?(mt/totalExp)*100:0;return <div key={m.id} style={{width:`${pct}%`,background:m.color}}/>;}) }
              </div>
              <div style={{display:"flex",gap:"12px"}}>
                {members.map(m=>{const mt=txList.filter(t=>t.author===m.id&&t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0);return(<div key={m.id} style={{display:"flex",alignItems:"center",gap:"5px"}}><div style={{width:"8px",height:"8px",borderRadius:"2px",background:m.color,flexShrink:0}}/><span style={{fontSize:"11px",color:"#64748B"}}>{m.name} {mt.toLocaleString()}원</span></div>);})}
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
                      <div><p style={{margin:0,fontSize:"14px",fontWeight:"600",color:"#1E293B"}}>{tx.name}</p><div style={{display:"flex",alignItems:"center",gap:"5px",marginTop:"2px"}}><span style={{fontSize:"10px",color:"#94A3B8"}}>{tx.category}</span><span style={{fontSize:"10px",color:"#CBD5E1"}}>·</span><span style={{fontSize:"10px",color:author.color,fontWeight:"600"}}>{author.name}</span>{tx.memo&&<><span style={{fontSize:"10px",color:"#CBD5E1"}}>·</span><span style={{fontSize:"10px",color:"#94A3B8"}}>{tx.memo}</span></>}</div></div>
                      <div style={{textAlign:"right"}}><p style={{margin:0,fontSize:"14px",fontWeight:"700",color:tx.amount>0?"#10B981":"#1E293B"}}>{fmt(tx.amount)}</p><p style={{margin:"2px 0 0",fontSize:"10px",color:"#CBD5E1"}}>{tx.date}</p></div>
                    </div>
                  </div>
                  <span style={{fontSize:"12px",color:"#CBD5E1"}}>✏️</span>
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
                      <span style={{fontSize:"10px",padding:"2px 7px",borderRadius:"99px",fontWeight:"700",background:m.role==="관리자"?"#EEF2FF":m.role==="멤버"?"#F0FDF4":"#F8FAFC",color:m.role==="관리자"?"#6366F1":m.role==="멤버"?"#10B981":"#94A3B8"}}>{m.role}</span>
                    </div>
                    <p style={{margin:0,fontSize:"11px",color:"#94A3B8"}}>{m.joined} 참여</p>
                  </div>
                  <div style={{textAlign:"right"}}><p style={{margin:0,fontSize:"13px",fontWeight:"700",color:m.color}}>{txList.filter(t=>t.author===m.id&&t.amount<0).reduce((s,t)=>s+Math.abs(t.amount),0).toLocaleString()}원</p><p style={{margin:"2px 0 0",fontSize:"10px",color:"#94A3B8"}}>이번달</p></div>
                </div>
              ))}
            </div>
            <button onClick={()=>setShowInvite(true)} style={{width:"100%",padding:"14px",border:"2px dashed #C7D2FE",borderRadius:"16px",background:"#EEF2FF",color:"#6366F1",fontSize:"14px",fontWeight:"700",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",gap:"8px",marginBottom:"16px"}}>
              <span style={{fontSize:"18px"}}>👥</span> 가족 초대하기
            </button>
            <div style={{background:"white",borderRadius:"16px",padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)"}}>
              <p style={{margin:"0 0 12px",fontSize:"14px",fontWeight:"700",color:"#1E293B"}}>우리 가족 코드</p>
              <div style={{background:"linear-gradient(135deg,#EEF2FF,#F5F3FF)",border:"2px dashed #A5B4FC",borderRadius:"12px",padding:"16px",textAlign:"center"}}>
                <p style={{margin:"0 0 6px",fontSize:"11px",color:"#6366F1",fontWeight:"600"}}>초대 코드</p>
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

// ════════════════════════════════
// 화면 7: 설정
// ════════════════════════════════
const SettingsScreen=()=>{
  const [selCards,setSelCards]=useState(["shinhan","hyundai","samsung"]);
  const [period,setPeriod]=useState("30");
  const [minAmt,setMinAmt]=useState("1000");
  const [exCancel,setExCancel]=useState(true);
  const [exOverseas,setExOverseas]=useState(false);
  const [keywords,setKeywords]=useState(["승인취소","해외결제","선불충전","포인트적립"]);
  const [newKw,setNewKw]=useState("");
  const [autoCat,setAutoCat]=useState(true);
  const [autoAI,setAutoAI]=useState(true);
  const [saved,setSaved]=useState(false);
  const toggleCard=(id)=>setSelCards(p=>p.includes(id)?p.filter(c=>c!==id):[...p,id]);
  const addKw=()=>{if(!newKw.trim())return;setKeywords(p=>[...p,newKw.trim()]);setNewKw("");};

  return(
    <div>
      <div style={{background:"linear-gradient(135deg,#6366F1 0%,#8B5CF6 100%)",padding:"48px 24px 24px",color:"white"}}>
        <h1 style={{margin:"0 0 4px",fontSize:"20px",fontWeight:"700"}}>문자 불러오기 설정 ⚙️</h1>
        <p style={{margin:0,fontSize:"13px",opacity:0.75}}>어떤 문자를 읽어올지 설정해요</p>
      </div>
      <div style={{padding:"20px 16px 140px"}}>
        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>카드사 선택</p>
        <div style={{background:"white",borderRadius:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px",overflow:"hidden"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 16px",borderBottom:"1px solid #F1F5F9",background:"#FAFBFF"}}>
            <span style={{fontSize:"13px",color:"#64748B",fontWeight:"600"}}>{selCards.length}개 선택됨</span>
            <button onClick={()=>setSelCards(selCards.length===CARD_PROVIDERS.length?[]:CARD_PROVIDERS.map(c=>c.id))} style={{background:"none",border:"none",cursor:"pointer",fontSize:"12px",color:"#6366F1",fontWeight:"600"}}>{selCards.length===CARD_PROVIDERS.length?"전체해제":"전체선택"}</button>
          </div>
          {CARD_PROVIDERS.map((card,i)=>{const iS=selCards.includes(card.id);return(
            <div key={card.id} onClick={()=>toggleCard(card.id)} style={{display:"flex",alignItems:"center",gap:"12px",padding:"12px 16px",borderBottom:i<CARD_PROVIDERS.length-1?"1px solid #F8FAFC":"none",cursor:"pointer",background:iS?"#FAFBFF":"white"}}>
              <div style={{width:"22px",height:"22px",borderRadius:"6px",background:iS?"#6366F1":"white",border:`2px solid ${iS?"#6366F1":"#CBD5E1"}`,display:"flex",alignItems:"center",justifyContent:"center",color:"white",fontSize:"13px",fontWeight:"700",flexShrink:0}}>{iS?"✓":""}</div>
              <span style={{fontSize:"16px"}}>{card.emoji}</span>
              <div style={{flex:1}}><p style={{margin:0,fontSize:"14px",fontWeight:iS?"600":"400",color:"#1E293B"}}>{card.name}</p><p style={{margin:"1px 0 0",fontSize:"10px",color:"#94A3B8",fontFamily:"monospace"}}>발신번호 {card.sender}</p></div>
              <span style={{fontSize:"10px",padding:"2px 6px",borderRadius:"99px",background:"#F1F5F9",color:"#64748B",fontFamily:"monospace"}}>{card.prefix}</span>
            </div>
          );})}
        </div>

        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>읽기 기간</p>
        <div style={{background:"white",borderRadius:"16px",padding:"12px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px"}}>
            {[{v:"7",l:"최근 1주"},{v:"30",l:"최근 1개월"},{v:"90",l:"최근 3개월"},{v:"custom",l:"직접 설정"}].map(o=>(
              <button key={o.v} onClick={()=>setPeriod(o.v)} style={{padding:"10px",border:"none",borderRadius:"10px",cursor:"pointer",background:period===o.v?"#6366F1":"#F1F5F9",color:period===o.v?"white":"#64748B",fontSize:"13px",fontWeight:period===o.v?"700":"400"}}>{o.l}</button>
            ))}
          </div>
        </div>

        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>금액 필터</p>
        <div style={{background:"white",borderRadius:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px",overflow:"hidden"}}>
          {[
            {icon:"💰",label:"최소 금액 이하 제외",sub:`${parseInt(minAmt||0).toLocaleString()}원 미만 제외`,right:<div style={{display:"flex",alignItems:"center",gap:"4px"}}><input type="number" value={minAmt} onChange={e=>setMinAmt(e.target.value)} onClick={e=>e.stopPropagation()} style={{width:"70px",padding:"6px 8px",border:"2px solid #E2E8F0",borderRadius:"8px",fontSize:"13px",color:"#6366F1",fontWeight:"600",outline:"none",textAlign:"right"}}/><span style={{fontSize:"12px",color:"#94A3B8"}}>원</span></div>},
            {icon:"🔄",label:"승인취소 제외",sub:"취소·환불 문자는 불러오지 않아요",right:<Toggle value={exCancel} onChange={setExCancel}/>},
            {icon:"✈️",label:"해외결제 제외",sub:"해외 가맹점 결제 제외",right:<Toggle value={exOverseas} onChange={setExOverseas}/>},
          ].map((row,i,arr)=>(
            <div key={row.label} style={{display:"flex",alignItems:"center",gap:"12px",padding:"13px 16px",borderBottom:i<arr.length-1?"1px solid #F8FAFC":"none"}}>
              <span style={{fontSize:"18px"}}>{row.icon}</span>
              <div style={{flex:1}}><p style={{margin:0,fontSize:"14px",fontWeight:"500",color:"#1E293B"}}>{row.label}</p><p style={{margin:"2px 0 0",fontSize:"11px",color:"#94A3B8"}}>{row.sub}</p></div>
              {row.right}
            </div>
          ))}
        </div>

        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>제외 키워드</p>
        <div style={{background:"white",borderRadius:"16px",padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px"}}>
          <div style={{display:"flex",flexWrap:"wrap",gap:"6px",marginBottom:"10px"}}>
            {keywords.map(kw=>(
              <div key={kw} style={{display:"flex",alignItems:"center",gap:"4px",background:"#FEE2E2",borderRadius:"99px",padding:"4px 10px"}}>
                <span style={{fontSize:"12px",color:"#EF4444",fontWeight:"600"}}>{kw}</span>
                <button onClick={()=>setKeywords(p=>p.filter(k=>k!==kw))} style={{background:"none",border:"none",cursor:"pointer",color:"#EF4444",fontSize:"14px",padding:"0",lineHeight:1,fontWeight:"700"}}>×</button>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:"8px"}}>
            <input value={newKw} onChange={e=>setNewKw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addKw()} placeholder="키워드 추가" style={{flex:1,padding:"8px 12px",border:"2px solid #E2E8F0",borderRadius:"10px",fontSize:"13px",outline:"none"}}/>
            <button onClick={addKw} style={{background:"#6366F1",border:"none",borderRadius:"10px",padding:"8px 14px",color:"white",fontSize:"13px",fontWeight:"700",cursor:"pointer"}}>추가</button>
          </div>
        </div>

        <p style={{margin:"0 0 10px",fontSize:"12px",fontWeight:"700",color:"#94A3B8",letterSpacing:"0.5px"}}>AI 자동화 설정</p>
        <div style={{background:"white",borderRadius:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"20px",overflow:"hidden"}}>
          {[{icon:"🤖",label:"AI 카테고리 자동 분류",sub:"상호명으로 카테고리를 자동으로 분류해요",value:autoCat,onChange:setAutoCat},{icon:"📊",label:"AI 지출 분석 자동 실행",sub:"홈 화면 진입 시 자동으로 분석해요",value:autoAI,onChange:setAutoAI}].map((row,i,arr)=>(
            <div key={row.label} style={{display:"flex",alignItems:"center",gap:"12px",padding:"13px 16px",borderBottom:i<arr.length-1?"1px solid #F8FAFC":"none"}}>
              <span style={{fontSize:"18px"}}>{row.icon}</span>
              <div style={{flex:1}}><p style={{margin:0,fontSize:"14px",fontWeight:"500",color:"#1E293B"}}>{row.label}</p><p style={{margin:"2px 0 0",fontSize:"11px",color:"#94A3B8"}}>{row.sub}</p></div>
              <Toggle value={row.value} onChange={row.onChange}/>
            </div>
          ))}
        </div>
      </div>
      <div style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",width:"100%",maxWidth:"430px",background:"white",borderTop:"1px solid #F1F5F9",padding:"12px 16px 72px",boxShadow:"0 -4px 20px rgba(0,0,0,0.08)"}}>
        <button onClick={()=>{setSaved(true);setTimeout(()=>setSaved(false),2000);}} style={{width:"100%",padding:"14px",border:"none",borderRadius:"14px",background:saved?"#10B981":"linear-gradient(135deg,#6366F1,#8B5CF6)",color:"white",fontSize:"15px",fontWeight:"700",cursor:"pointer",transition:"background 0.3s"}}>{saved?"✅ 설정이 저장됐어요!":"설정 저장하기"}</button>
      </div>
    </div>
  );
};

// ════════════════════════════════
// 메인 앱 라우터
// ════════════════════════════════

// ── 카테고리 매핑 ──
const SUBCAT_MAP = {
  "커피/음료": { name:"카페",   icon:"☕",  color:"#A78BFA" },
  "식사/간식": { name:"식비",   icon:"🍚", color:"#F97316" },
  "식재료":    { name:"식비",   icon:"🛒", color:"#F97316" },
  "의류/잡화": { name:"쇼핑",   icon:"🛍️", color:"#EC4899" },
  "전자제품":  { name:"쇼핑",   icon:"📱", color:"#EC4899" },
  "화장품":    { name:"뷰티",   icon:"💄", color:"#06B6D4" },
  "병원비/약값":{ name:"의료",  icon:"🏥", color:"#EF4444" },
  "영화/공연": { name:"문화",   icon:"🎬", color:"#F59E0B" },
  "여행":      { name:"문화",   icon:"✈️", color:"#F59E0B" },
  "도서":      { name:"문화",   icon:"📚", color:"#F59E0B" },
  "택시비":    { name:"교통",   icon:"🚕", color:"#3B82F6" },
  "대중교통":  { name:"교통",   icon:"🚌", color:"#3B82F6" },
  "유류비":    { name:"교통",   icon:"⛽", color:"#3B82F6" },
  "주차/통행": { name:"교통",   icon:"🚗", color:"#3B82F6" },
  "교재비":    { name:"기타",   icon:"📚", color:"#6366F1" },
  "통신비":    { name:"기타",   icon:"📡", color:"#94A3B8" },
  "생활세금":  { name:"기타",   icon:"🏛️", color:"#94A3B8" },
  "자동차보험":{ name:"기타",   icon:"🚗", color:"#94A3B8" },
  "보장보험":  { name:"기타",   icon:"🛡️", color:"#94A3B8" },
};
const CAT_MAP = {
  "식자재":    { name:"식비",   icon:"🍚", color:"#F97316" },
  "쇼핑비":    { name:"쇼핑",   icon:"🛍️", color:"#EC4899" },
  "교통비":    { name:"교통",   icon:"🚌", color:"#3B82F6" },
  "차량유지비":{ name:"교통",   icon:"🚌", color:"#3B82F6" },
  "건강관리비":{ name:"의료",   icon:"🏥", color:"#EF4444" },
  "문화생활비":{ name:"문화",   icon:"🎬", color:"#F59E0B" },
  "교육비":    { name:"기타",   icon:"📚", color:"#6366F1" },
  "미용비":    { name:"뷰티",   icon:"🌿", color:"#06B6D4" },
  "생활비":    { name:"기타",   icon:"🏠", color:"#94A3B8" },
  "금융보험비":{ name:"기타",   icon:"💰", color:"#94A3B8" },
  "없음":      { name:"기타",   icon:"📦", color:"#94A3B8" },
};

const getMeta = (subcat, cat) =>
  SUBCAT_MAP[subcat] || CAT_MAP[cat] || { name:"기타", icon:"📦", color:"#94A3B8" };

// XLS 파싱
const parseXls = (buffer) => {
  const wb = XLSX.read(buffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const headerIdx = raw.findIndex(row => row.includes("지출일"));
  if (headerIdx === -1) throw new Error("똑똑가계부 형식이 아니에요");
  const headers = raw[headerIdx];
  const dataRows = raw.slice(headerIdx + 1).filter(row =>
    row[0] && String(row[0]).includes(".")
  );
  return dataRows.map((row, i) => {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = row[j]; });
    const meta = getMeta(String(obj["세부카테고리"]||""), String(obj["카테고리"]||""));
    return {
      id: Date.now() + i,
      name: String(obj["지출내역"] || ""),
      amount: -Math.abs(parseInt(obj["지출금액"]) || 0),
      date: String(obj["지출일"]||"").split(" ")[0],
      time: String(obj["지출시간"]||""),
      card: String(obj["결제"]||"-"),
      category: meta.name,
      icon: meta.icon,
      color: meta.color,
      memo: String(obj["메모"]||""),
      author: 1,
      subcat: String(obj["세부카테고리"]||""),
    };
  });
};

// 날짜 포맷
const fmtDate = (d) => {
  if (!d) return "";
  const dt = new Date(d);
  return isNaN(dt) ? String(d) : dt.toLocaleDateString("ko-KR");
};

// ── 메인 ──
const ExcelImportScreen = ({ onImport }) => {
  // 폴더 핸들 (File System Access API)
  const [dirHandle, setDirHandle] = useState(null);
  const [latestFile, setLatestFile] = useState(null); // { name, handle, lastModified }
  const [scanning, setScanning] = useState(false);
  const [step, setStep] = useState("idle"); // idle | preview | done
  const [parsed, setParsed] = useState([]);
  const [selected, setSelected] = useState([]);
  const [filterCat, setFilterCat] = useState(null);
  const [error, setError] = useState(null);
  const fileRef = useRef();

  // File System Access API 지원 여부
  const fsa = typeof window !== "undefined" && "showDirectoryPicker" in window;

  // 폴더에서 최신 xls 파일 찾기
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
      if (!latest) throw new Error("폴더에 엑셀 파일이 없어요");
      setLatestFile(latest);
    } catch (e) {
      if (e.name !== "AbortError") setError(e.message || "폴더를 읽을 수 없어요");
    }
    setScanning(false);
  };

  // 폴더 선택
  const handleSelectDir = async () => {
    try {
      const handle = await window.showDirectoryPicker({ mode: "read" });
      setDirHandle(handle);
      setLatestFile(null);
      setParsed([]);
      setStep("idle");
      await scanDir(handle);
    } catch (e) {
      if (e.name !== "AbortError") setError("폴더 선택 취소됨");
    }
  };

  // 새로고침 (최신 파일 다시 스캔)
  const handleRefresh = async () => {
    if (!dirHandle) return;
    setLatestFile(null);
    setParsed([]);
    setStep("idle");
    await scanDir(dirHandle);
  };

  // 파일 불러오기 (FSA)
  const handleLoadLatest = async () => {
    if (!latestFile) return;
    setError(null);
    try {
      const file = latestFile.file || await latestFile.handle.getFile();
      const buffer = await file.arrayBuffer();
      const rows = parseXls(new Uint8Array(buffer));
      if (rows.length === 0) throw new Error("데이터가 없어요");
      setParsed(rows);
      setSelected(rows.map(r => r.id));
      setStep("preview");
    } catch (e) {
      setError(e.message || "파일을 읽을 수 없어요");
    }
  };

  // 수동 파일 선택 (FSA 미지원 fallback)
  const handleManualFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError(null);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const rows = parseXls(new Uint8Array(ev.target.result));
        if (rows.length === 0) throw new Error("데이터가 없어요");
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
      {/* 헤더 */}
      <div style={{
        background:"linear-gradient(135deg,#10B981 0%,#059669 100%)",
        padding:"48px 24px 24px", color:"white",
      }}>
        <h1 style={{margin:"0 0 4px",fontSize:"20px",fontWeight:"700"}}>
          📊 똑똑가계부 불러오기
        </h1>
        <p style={{margin:0,fontSize:"13px",opacity:0.85}}>
          폴더를 한 번 지정하면 최신 파일을 자동으로 찾아요
        </p>
      </div>

      <div style={{padding:"20px 16px 100px"}}>

        {/* ── 폴더 지정 영역 ── */}
        <div style={{
          background:"white", borderRadius:"16px", padding:"16px",
          boxShadow:"0 1px 6px rgba(0,0,0,0.06)", marginBottom:"16px",
        }}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
            <p style={{margin:0,fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>
              📁 다운로드 폴더 연결
            </p>
            {dirHandle && (
              <button onClick={handleRefresh} disabled={scanning} style={{
                background:"none",border:"none",cursor:"pointer",
                fontSize:"12px",color:"#10B981",fontWeight:"600",
              }}>{scanning?"스캔 중...":"🔄 새로고침"}</button>
            )}
          </div>

          {fsa ? (
            <>
              {/* 폴더 선택 버튼 */}
              <button onClick={handleSelectDir} style={{
                width:"100%", padding:"14px",
                border:`2px dashed ${dirHandle?"#A7F3D0":"#CBD5E1"}`,
                borderRadius:"14px",
                background:dirHandle?"#F0FDF4":"#F8FAFC",
                cursor:"pointer", display:"flex", alignItems:"center", gap:"12px",
                marginBottom: latestFile ? "12px" : "0",
              }}>
                <span style={{fontSize:"28px"}}>{dirHandle?"📂":"📁"}</span>
                <div style={{textAlign:"left"}}>
                  <p style={{margin:0,fontSize:"13px",fontWeight:"700",color:dirHandle?"#065F46":"#64748B"}}>
                    {dirHandle ? `📂 ${dirHandle.name}` : "폴더 선택하기"}
                  </p>
                  <p style={{margin:"2px 0 0",fontSize:"11px",color:dirHandle?"#10B981":"#94A3B8"}}>
                    {dirHandle ? "클릭해서 폴더 변경" : "똑똑가계부 엑셀이 저장되는 폴더"}
                  </p>
                </div>
              </button>

              {/* 스캔 중 */}
              {scanning && (
                <div style={{
                  background:"#F0FDF4",borderRadius:"12px",padding:"12px 14px",
                  display:"flex",alignItems:"center",gap:"10px",marginTop:"10px",
                }}>
                  <span style={{fontSize:"16px"}}>🔍</span>
                  <p style={{margin:0,fontSize:"13px",color:"#065F46"}}>최신 파일 찾는 중...</p>
                </div>
              )}

              {/* 최신 파일 감지됨 */}
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
                        ✨ 최신 파일 발견
                      </p>
                      <p style={{margin:"4px 0 0",fontSize:"14px",fontWeight:"700",color:"#065F46"}}>
                        {latestFile.name}
                      </p>
                      <p style={{margin:"2px 0 0",fontSize:"11px",color:"#6EE7B7"}}>
                        {new Date(latestFile.lastModified).toLocaleString("ko-KR")}
                      </p>
                    </div>
                    <span style={{fontSize:"32px"}}>📄</span>
                  </div>
                  <button onClick={handleLoadLatest} style={{
                    width:"100%", padding:"12px",
                    border:"none", borderRadius:"12px",
                    background:"linear-gradient(135deg,#10B981,#059669)",
                    color:"white", fontSize:"14px", fontWeight:"700",
                    cursor:"pointer", display:"flex", alignItems:"center",
                    justifyContent:"center", gap:"8px",
                  }}>
                    <span style={{fontSize:"16px"}}>⬇️</span>
                    이 파일 불러오기
                  </button>
                </div>
              )}
            </>
          ) : (
            /* FSA 미지원 - 수동 파일 선택 */
            <>
              <div style={{
                background:"#FFF7ED",borderRadius:"10px",padding:"10px 12px",
                border:"1px solid #FFEDD5",marginBottom:"10px",
              }}>
                <p style={{margin:0,fontSize:"11px",color:"#92400E"}}>
                  💡 이 브라우저는 폴더 자동 감지를 지원하지 않아요. 파일을 직접 선택해주세요.
                </p>
              </div>
              <input ref={fileRef} type="file" accept=".xls,.xlsx" onChange={handleManualFile} style={{display:"none"}}/>
              <button onClick={()=>fileRef.current.click()} style={{
                width:"100%",padding:"14px",border:"2px dashed #CBD5E1",
                borderRadius:"14px",background:"#F8FAFC",cursor:"pointer",
                display:"flex",alignItems:"center",justifyContent:"center",gap:"10px",
              }}>
                <span style={{fontSize:"20px"}}>📂</span>
                <span style={{fontSize:"14px",fontWeight:"700",color:"#64748B"}}>엑셀 파일 직접 선택</span>
              </button>
            </>
          )}
        </div>

        {/* 에러 */}
        {error && (
          <div style={{
            background:"#FEE2E2",borderRadius:"12px",padding:"12px 14px",
            border:"1px solid #FECACA",display:"flex",gap:"8px",marginBottom:"12px",
          }}>
            <span>⚠️</span>
            <p style={{margin:0,fontSize:"13px",color:"#991B1B"}}>{error}</p>
          </div>
        )}

        {/* 사용 안내 (폴더 미선택 시) */}
        {!dirHandle && !latestFile && step === "idle" && (
          <div style={{
            background:"white", borderRadius:"16px", padding:"16px",
            boxShadow:"0 1px 6px rgba(0,0,0,0.06)",
          }}>
            <p style={{margin:"0 0 12px",fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>📋 사용 방법</p>
            {[
              ["1️⃣","똑똑가계부 앱 → 내보내기 → 엑셀 저장"],
              ["2️⃣","저장 폴더(보통 다운로드)를 위에서 선택"],
              ["3️⃣","앱이 최신 파일을 자동으로 찾아줘요"],
              ["4️⃣","불러오기 버튼 한 번만 누르면 완료!"],
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
                💡 폴더는 한 번만 지정하면 돼요. 다음에 앱을 열어도 같은 폴더에서 최신 파일을 자동으로 찾아요.
              </p>
            </div>
          </div>
        )}

        {/* ── 미리보기 ── */}
        {step === "preview" && (
          <>
            {/* 요약 카드 */}
            <div style={{
              background:"white",borderRadius:"16px",padding:"14px 16px",
              boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px",
            }}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
                <p style={{margin:0,fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>
                  {latestFile?.name || "불러온 파일"}
                </p>
                <button onClick={()=>{setStep("idle");setParsed([]);setSelected([]);}} style={{
                  background:"none",border:"none",cursor:"pointer",
                  fontSize:"12px",color:"#10B981",fontWeight:"600",
                }}>다시 선택</button>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px"}}>
                {[
                  {l:"총 건수",v:`${parsed.length}건`},
                  {l:"선택",v:`${selected.length}건`},
                  {l:"합계",v:`${Math.round(totalAmt/10000)}만원`},
                ].map(({l,v})=>(
                  <div key={l} style={{background:"#F0FDF4",borderRadius:"10px",padding:"10px",textAlign:"center",border:"1px solid #D1FAE5"}}>
                    <p style={{margin:0,fontSize:"10px",color:"#059669"}}>{l}</p>
                    <p style={{margin:"3px 0 0",fontSize:"15px",fontWeight:"700",color:"#065F46"}}>{v}</p>
                  </div>
                ))}
              </div>
            </div>

            {/* 카테고리 통계 */}
            <div style={{
              background:"white",borderRadius:"16px",padding:"14px 16px",
              boxShadow:"0 1px 6px rgba(0,0,0,0.06)",marginBottom:"12px",
            }}>
              <p style={{margin:"0 0 10px",fontSize:"13px",fontWeight:"700",color:"#1E293B"}}>카테고리별 요약</p>
              {Object.entries(catStats).sort((a,b)=>b[1].amount-a[1].amount).map(([cat,s])=>{
                const max = Math.max(...Object.values(catStats).map(v=>v.amount));
                const pct = Math.round((s.amount/max)*100);
                return (
                  <div key={cat} style={{marginBottom:"8px"}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:"3px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:"5px"}}>
                        <span style={{fontSize:"13px"}}>{s.icon}</span>
                        <span style={{fontSize:"12px",color:"#334155"}}>{cat}</span>
                        <span style={{fontSize:"10px",color:"#94A3B8"}}>{s.count}건</span>
                      </div>
                      <span style={{fontSize:"12px",fontWeight:"700",color:"#1E293B"}}>{s.amount.toLocaleString()}원</span>
                    </div>
                    <div style={{background:"#F1F5F9",borderRadius:"99px",height:"5px"}}>
                      <div style={{width:`${pct}%`,height:"100%",background:s.color,borderRadius:"99px"}}/>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* 카테고리 필터 */}
            <div style={{display:"flex",gap:"6px",overflowX:"auto",paddingBottom:"4px",marginBottom:"10px"}}>
              <button onClick={()=>setFilterCat(null)} style={{
                padding:"5px 12px",border:"none",borderRadius:"99px",
                background:!filterCat?"#10B981":"#F1F5F9",
                color:!filterCat?"white":"#64748B",
                fontWeight:!filterCat?"700":"400",fontSize:"12px",flexShrink:0,cursor:"pointer",
              }}>전체 {parsed.length}</button>
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

            {/* 전체선택 */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"8px"}}>
              <p style={{margin:0,fontSize:"12px",color:"#64748B"}}>{filtered.length}건</p>
              <button onClick={()=>setSelected(
                selected.length===parsed.length
