/* ============================================================
   APEX — Executive Apartment Portfolio Dashboard · script.js
   No animations · static data sourced from Excel workbook
   ============================================================ */

/* ---------- 1. DATA OBJECTS ---------- */
const DATA = {
  months: ["Jul '23","Aug '23","Sep '23","Oct '23","Nov '23","Dec '23","Jan '24","Feb '24","Mar '24","Apr '24","May '24","Jun '24","Jul '24","Aug '24","Sep '24","Oct '24","Nov '24","Dec '24","Jan '25","Feb '25","Mar '25","Apr '25","May '25","Jun '25"],
  occTrend: [92.2,92.2,92.7,92.8,93.1,93.2,93.7,93.5,94.1,94.2,94.2,94.9,95.1,95.0,95.5,95.5,95.6,95.9,95.9,96.3,96.5,96.5,96.7,96.8],
  noiTrend: [3346,3127,3402,3119,3265,3132,3451,3205,3377,3238,3392,3410,3585,3490,3457,3420,3484,3281,3464,3369,3421,3569,3406,3491],
  revActual: [5837,5834,5890,5973,5882,5948,5798,5951,6095,6151,5896,6057,5990,5998,6164,6149,5918,6128,6214,6053,6166,5976,6258,6138],
  revBudget: [5852,5882,5929,6004,5941,5972,5846,6001,6135,6203,5922,6109,6005,6015,6199,6164,5946,6151,6225,6091,6182,6018,6274,6171],
  portfolio: {occ:94.7,noi:80898570,margin:61.7,cap:5.11,deln:3.3,dscr:1.66,value:562403085,equity:272123085,loan:290280000,eqGrowth:67.0,debtYield:9.9,coc:4.0},
  properties: [
    {name:"Parkside at Midtown",region:"Southeast",noi:5836013,margin:63.2,occ:91.9,value:49049914,loan:26292000,equity:22757914,eqGrowth:42.8,debtYield:8.85,dscr:1.54,coc:2.5,cap:4.74,deln:4.3,vacDays:17.1,aged:262,conv:22.0,budgetVar:0.0},
    {name:"The Hadley",region:"Southeast",noi:5102340,margin:60.9,occ:90.4,value:43215000,loan:23100000,equity:20115000,eqGrowth:38.4,debtYield:8.42,dscr:1.49,coc:2.1,cap:4.61,deln:4.9,vacDays:19.6,aged:248,conv:21.3,budgetVar:-0.6},
    {name:"Vista Ridge",region:"Southwest",noi:8211460,margin:62.8,occ:95.1,value:71840000,loan:37800000,equity:34040000,eqGrowth:71.2,debtYield:9.41,dscr:1.71,coc:4.6,cap:5.34,deln:2.7,vacDays:14.2,aged:201,conv:26.8,budgetVar:-1.5},
    {name:"Magnolia Park",region:"Southwest",noi:7044900,margin:59.4,occ:93.7,value:60120000,loan:33200000,equity:26920000,eqGrowth:55.0,debtYield:8.94,dscr:1.58,coc:3.4,cap:5.08,deln:3.8,vacDays:18.9,aged:233,conv:23.6,budgetVar:-0.9},
    {name:"Crestline Flats",region:"Southwest",noi:5398220,margin:61.7,occ:94.8,value:46500000,loan:24600000,equity:21900000,eqGrowth:58.7,debtYield:9.12,dscr:1.64,coc:3.9,cap:5.21,deln:3.1,vacDays:15.3,aged:188,conv:25.4,budgetVar:0.2},
    {name:"Lakeshore Commons",region:"Midwest",noi:6122870,margin:64.1,occ:96.2,value:54300000,loan:27500000,equity:26800000,eqGrowth:78.7,debtYield:9.73,dscr:1.78,coc:5.1,cap:5.46,deln:2.4,vacDays:12.8,aged:159,conv:28.1,budgetVar:0.8},
    {name:"Iron Gate",region:"Midwest",noi:4388510,margin:57.9,occ:92.6,value:38900000,loan:21800000,equity:17100000,eqGrowth:41.3,debtYield:8.05,dscr:1.45,coc:2.0,cap:4.83,deln:4.6,vacDays:20.4,aged:276,conv:22.7,budgetVar:-1.8},
    {name:"Summit Pointe",region:"Mountain",noi:7588340,margin:60.6,occ:94.3,value:65400000,loan:35900000,equity:29500000,eqGrowth:59.5,debtYield:8.86,dscr:1.57,coc:3.3,cap:5.02,deln:4.1,vacDays:18.2,aged:284,conv:23.9,budgetVar:-1.2},
    {name:"Harborview",region:"Southeast",noi:9044120,margin:65.3,occ:96.7,value:79200000,loan:39600000,equity:39600000,eqGrowth:84.1,debtYield:9.97,dscr:1.83,coc:5.6,cap:5.58,deln:2.1,vacDays:11.9,aged:142,conv:29.4,budgetVar:1.1},
    {name:"Camden Row",region:"Southeast",noi:8161796,margin:62.0,occ:95.4,value:53774171,loan:20488000,equity:33286171,eqGrowth:88.3,debtYield:10.4,dscr:1.69,coc:4.4,cap:5.18,deln:2.9,vacDays:16.0,aged:219,conv:25.0,budgetVar:-0.3}
  ]
};

/* ---------- 2. FORMAT HELPERS ---------- */
const fmtMoney = (n) => {
  if (Math.abs(n) >= 1e9) return '$' + (n/1e9).toFixed(2) + 'B';
  if (Math.abs(n) >= 1e6) return '$' + (n/1e6).toFixed(1) + 'M';
  if (Math.abs(n) >= 1e3) return '$' + (n/1e3).toFixed(0) + 'K';
  return '$' + n;
};
const fmtPct = (n) => n.toFixed(1) + '%';
const hexA = (hex,a)=>{const n=parseInt(hex.slice(1),16);return `rgba(${n>>16&255},${n>>8&255},${n&255},${a})`;};

/* ---------- 3. CHART THEME (no animation) ---------- */
const C = { teal:'#5eead4', blue:'#7aa2ff', gold:'#e6c068', green:'#4ade80',
  amber:'#fbbf24', red:'#fb6f6f', grid:'rgba(150,178,228,.07)', txt:'#a3b3d0', txt3:'#697a9a' };
Chart.defaults.font.family = "'Sora', sans-serif";
Chart.defaults.color = C.txt;
Chart.defaults.font.size = 12;
Chart.defaults.animation = false;

const baseOpts = (extra={}) => Object.assign({
  responsive:true, maintainAspectRatio:false, animation:false,
  plugins:{ legend:{display:false},
    tooltip:{ backgroundColor:'#0a0f1c', borderColor:'rgba(150,178,228,.25)', borderWidth:1,
      titleColor:'#eaf0fb', bodyColor:'#a3b3d0', padding:12, cornerRadius:10, displayColors:false } },
  scales:{ x:{ grid:{color:C.grid,drawBorder:false}, ticks:{color:C.txt3,maxRotation:0,autoSkip:true,maxTicksLimit:8} },
           y:{ grid:{color:C.grid,drawBorder:false}, ticks:{color:C.txt3} } }
}, extra);

/* ---------- 4. MINI SPARKLINE (inline SVG) ---------- */
function sparkline(data, color) {
  const w=58, h=24, min=Math.min(...data), max=Math.max(...data), rng=(max-min)||1;
  const pts = data.map((v,i)=>`${(i/(data.length-1))*w},${h-((v-min)/rng)*(h-4)-2}`).join(' ');
  const id = 'sg'+Math.random().toString(36).slice(2,7);
  return `<svg class="kpi-spark" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">
    <defs><linearGradient id="${id}" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="${color}" stop-opacity=".35"/><stop offset="100%" stop-color="${color}" stop-opacity="0"/>
    </linearGradient></defs>
    <polyline points="${pts}" fill="none" stroke="${color}" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
    <polygon points="0,${h} ${pts} ${w},${h}" fill="url(#${id})"/></svg>`;
}

/* ---------- 5. KPI CARD RENDERER (instant values) ---------- */
function kpiCard(k) {
  const arrow = k.trendDir==='up' ? '▲' : k.trendDir==='down' ? '▼' : '◆';
  const spark = k.spark ? sparkline(k.spark, k.sparkColor||C.teal) : '';
  return `
    <div class="kpi">
      <div class="kpi-top">
        <span class="kpi-name"><span class="dot"></span>${k.name}</span>
        ${spark}
      </div>
      <div class="kpi-val">${k.value}<span class="u">${k.unit||''}</span></div>
      <div class="kpi-foot">
        <span class="kpi-trend ${k.trendDir}">${arrow} ${k.trend}</span>
      </div>
      <p class="kpi-mean">${k.meaning}</p>
    </div>`;
}

function renderExecKpis() {
  const p = DATA.portfolio;
  const occS = DATA.occTrend.slice(-12);
  const noiS = DATA.noiTrend.slice(-12);
  const cards = [
    {name:'Physical Occupancy', value:fmtPct(p.occ), unit:'', trend:'+4.6 pts', trendDir:'up', spark:occS, sparkColor:C.teal, meaning:'Share of units occupied across the portfolio.'},
    {name:'Portfolio NOI', value:fmtMoney(p.noi), unit:'', trend:'24-month total', trendDir:'up', spark:noiS, sparkColor:C.blue, meaning:'Net operating income generated over the period.'},
    {name:'NOI Margin', value:fmtPct(p.margin), unit:'', trend:'Healthy band', trendDir:'up', spark:noiS, sparkColor:C.green, meaning:'NOI as a share of effective gross income.'},
    {name:'Portfolio Cap Rate', value:fmtPct(p.cap), unit:'', trend:'Income yield', trendDir:'flat', spark:occS, sparkColor:C.gold, meaning:'Unlevered yield on current asset value.'},
    {name:'Delinquency 30+', value:fmtPct(p.deln), unit:'', trend:'Below 5% target', trendDir:'down', spark:occS.map(x=>100-x), sparkColor:C.amber, meaning:'Receivables past due 30+ days.'},
    {name:'Average DSCR', value:p.dscr.toFixed(2), unit:'x', trend:'Comfortable cover', trendDir:'up', spark:noiS, sparkColor:C.teal, meaning:'Operations cover debt service comfortably.'},
    {name:'Total Current Value', value:fmtMoney(p.value), unit:'', trend:'+15% appreciation', trendDir:'up', spark:noiS, sparkColor:C.blue, meaning:'Estimated market value of the portfolio.'},
    {name:'Total Equity Value', value:fmtMoney(p.equity), unit:'', trend:'+'+p.eqGrowth+'%', trendDir:'up', spark:noiS, sparkColor:C.green, meaning:'Net equity after outstanding debt.'}
  ];
  document.getElementById('kpiGrid').innerHTML = cards.map(kpiCard).join('');
}

function renderInvestmentKpis() {
  const p = DATA.portfolio;
  const eqS = DATA.properties.map(x=>x.equity);
  const cards = [
    {name:'Estimated Value', value:fmtMoney(p.value), unit:'', trend:'Direct-cap', trendDir:'up', spark:eqS, sparkColor:C.teal, meaning:'Income-based valuation of all assets.'},
    {name:'Loan Balance', value:fmtMoney(p.loan), unit:'', trend:'~52% LTV', trendDir:'flat', spark:eqS, sparkColor:C.gold, meaning:'Outstanding debt against the portfolio.'},
    {name:'Equity Value', value:fmtMoney(p.equity), unit:'', trend:'+'+p.eqGrowth+'%', trendDir:'up', spark:eqS, sparkColor:C.blue, meaning:'Value remaining after debt is repaid.'},
    {name:'Equity Growth', value:fmtPct(p.eqGrowth), unit:'', trend:'Since acquisition', trendDir:'up', spark:eqS, sparkColor:C.green, meaning:'Appreciation of invested equity.'},
    {name:'Debt Yield', value:fmtPct(p.debtYield), unit:'', trend:'Lender-safe', trendDir:'up', spark:eqS, sparkColor:C.teal, meaning:'NOI return on the loan balance.'},
    {name:'DSCR', value:p.dscr.toFixed(2), unit:'x', trend:'>1.25 covenant', trendDir:'up', spark:eqS, sparkColor:C.blue, meaning:'Debt service coverage ratio.'},
    {name:'Cash-on-Cash', value:fmtPct(p.coc), unit:'', trend:'Annualized', trendDir:'flat', spark:eqS, sparkColor:C.gold, meaning:'Cash yield on invested equity.'},
    {name:'Cap Rate', value:fmtPct(p.cap), unit:'', trend:'Portfolio blend', trendDir:'flat', spark:eqS, sparkColor:C.green, meaning:'Unlevered income yield.'}
  ];
  document.getElementById('invKpiGrid').innerHTML = cards.map(kpiCard).join('');
}

/* ---------- 6. HERO BADGES + MARQUEE ---------- */
function renderHeroBadges() {
  const p = DATA.portfolio;
  const badges = [
    {v:'10', l:'Properties'},
    {v:'24', l:'Months modeled'},
    {v:fmtMoney(p.value), l:'Portfolio value'},
    {v:fmtPct(p.occ), l:'Occupancy'}
  ];
  document.getElementById('heroBadges').innerHTML = badges.map(b=>
    `<div class="hbadge"><span class="v">${b.v}</span><span class="l">${b.l}</span></div>`).join('');

  const items = ['Power Query','Power Pivot','DAX Measures','Star Schema','Power Automate','Executive Reporting','Chart.js','Data Modeling'];
  const row = items.map(i=>`<span>${i}</span>`).join('');
  document.getElementById('marquee').innerHTML = row + row;
}

/* ---------- 7. HERO DASHBOARD MOCKUP ---------- */
function renderMock() {
  const p = DATA.portfolio;
  document.getElementById('mockKpis').innerHTML = [
    {l:'Occupancy', v:fmtPct(p.occ)},
    {l:'NOI', v:fmtMoney(p.noi)},
    {l:'DSCR', v:p.dscr.toFixed(2)+'x'}
  ].map(k=>`<div class="mock-mini"><div class="l">${k.l}</div><div class="v">${k.v}</div></div>`).join('');

  const top = [...DATA.properties].sort((a,b)=>b.noi-a.noi).slice(0,4);
  const max = top[0].noi;
  document.getElementById('mockRows').innerHTML = top.map(t=>`
    <div class="mock-row"><span class="nm">${t.name}</span>
      <span class="bar" style="width:${42*(t.noi/max)}px"></span>
      <span class="rb" style="color:${t.margin>=60?C.green:C.amber};background:${t.margin>=60?'rgba(74,222,128,.12)':'rgba(251,191,36,.12)'}">${t.margin.toFixed(0)}%</span>
    </div>`).join('');

  document.getElementById('mockFloat1').innerHTML = `<div class="l">Cap Rate</div><div class="v">${fmtPct(p.cap)} <small>▲</small></div>`;
  document.getElementById('mockFloat2').innerHTML = `<div class="l">Equity</div><div class="v">${fmtMoney(p.equity)} <small>+${p.eqGrowth}%</small></div>`;

  const ctx = document.getElementById('mockChart').getContext('2d');
  new Chart(ctx, {
    type:'line',
    data:{ labels:DATA.months.slice(-12), datasets:[{ data:DATA.noiTrend.slice(-12),
      borderColor:C.teal, backgroundColor:hexA(C.teal,0.12), borderWidth:2, fill:true, tension:.45, pointRadius:0 }]},
    options:{ responsive:true, maintainAspectRatio:false, animation:false,
      plugins:{legend:{display:false},tooltip:{enabled:false}},
      scales:{x:{display:false},y:{display:false}} }
  });
}

/* ---------- 8. CHART RENDERERS ---------- */
function lineChart(id, data, color, fmt) {
  const ctx = document.getElementById(id).getContext('2d');
  new Chart(ctx, {
    type:'line',
    data:{ labels:DATA.months, datasets:[{ data, borderColor:color, backgroundColor:hexA(color,0.13),
      borderWidth:2.5, fill:true, tension:.4, pointRadius:0, pointHoverRadius:5, pointHoverBackgroundColor:color }]},
    options:baseOpts({ plugins:{ legend:{display:false},
      tooltip:{ backgroundColor:'#0a0f1c',borderColor:'rgba(150,178,228,.25)',borderWidth:1,padding:12,cornerRadius:10,displayColors:false,
        callbacks:{ label:(c)=> ' '+(fmt?fmt(c.parsed.y):c.parsed.y) } } } })
  });
}
function multiLineChart(id, datasets) {
  const ctx = document.getElementById(id).getContext('2d');
  new Chart(ctx, {
    type:'line',
    data:{ labels:DATA.months, datasets:datasets.map(d=>({ label:d.label, data:d.data, borderColor:d.color,
      backgroundColor:hexA(d.color,0.08), borderWidth:2.5, borderDash:d.dash||[], fill:false, tension:.4, pointRadius:0, pointHoverRadius:5 })) },
    options:baseOpts({ plugins:{
      legend:{display:true,position:'top',align:'end',labels:{usePointStyle:true,pointStyle:'line',boxWidth:24,color:C.txt}},
      tooltip:{ backgroundColor:'#0a0f1c',borderColor:'rgba(150,178,228,.25)',borderWidth:1,padding:12,cornerRadius:10,
        callbacks:{ label:(c)=> ' '+c.dataset.label+': $'+(c.parsed.y/1000).toFixed(2)+'M' } } },
      scales:{ x:{grid:{color:C.grid},ticks:{color:C.txt3,maxTicksLimit:8}},
               y:{grid:{color:C.grid},ticks:{color:C.txt3,callback:(v)=>'$'+(v/1000).toFixed(1)+'M'}} } })
  });
}
function barChart(id, labels, data, color, horizontal=false, fmt) {
  const ctx = document.getElementById(id).getContext('2d');
  new Chart(ctx, {
    type:'bar',
    data:{ labels, datasets:[{ data, backgroundColor:labels.map((_,i)=>hexA(color,0.5+0.5*(i/labels.length))),
      borderColor:color, borderWidth:1.5, borderRadius:7, maxBarThickness:34 }]},
    options:baseOpts({ indexAxis: horizontal?'y':'x',
      plugins:{ legend:{display:false},
        tooltip:{ backgroundColor:'#0a0f1c',borderColor:'rgba(150,178,228,.25)',borderWidth:1,padding:12,cornerRadius:10,displayColors:false,
          callbacks:{ label:(c)=> ' '+(fmt?fmt(horizontal?c.parsed.x:c.parsed.y):(horizontal?c.parsed.x:c.parsed.y)) } } },
      scales:{ x:{grid:{color:C.grid},ticks:{color:C.txt3,callback:horizontal?(v)=>fmt?fmt(v):v:undefined}},
               y:{grid:{color:C.grid},ticks:{color:C.txt3,callback:!horizontal?(v)=>fmt?fmt(v):v:undefined}} } })
  });
}
function renderCharts() {
  lineChart('occChart', DATA.occTrend, C.teal, (v)=>v+'%');
  lineChart('noiChart', DATA.noiTrend, C.blue, (v)=>'$'+(v/1000).toFixed(2)+'M');
  multiLineChart('revChart', [
    {label:'Actual EGI', data:DATA.revActual, color:C.teal},
    {label:'Budget', data:DATA.revBudget, color:C.gold, dash:[6,4]}
  ]);
  const byMargin = [...DATA.properties].sort((a,b)=>b.margin-a.margin).slice(0,5);
  barChart('topChart', byMargin.map(p=>p.name), byMargin.map(p=>p.margin), C.green, true, (v)=>v.toFixed(1)+'%');
  const byOcc = [...DATA.properties].sort((a,b)=>a.occ-b.occ).slice(0,5);
  barChart('botChart', byOcc.map(p=>p.name), byOcc.map(p=>p.occ), C.red, true, (v)=>v.toFixed(1)+'%');
  const byVal = [...DATA.properties].sort((a,b)=>b.value-a.value);
  barChart('valChart', byVal.map(p=>p.name), byVal.map(p=>p.value), C.blue, true, fmtMoney);
  const byEq = [...DATA.properties].sort((a,b)=>b.equity-a.equity);
  barChart('eqChart', byEq.map(p=>p.name), byEq.map(p=>p.equity), C.teal, true, fmtMoney);
  const byDscr = [...DATA.properties].sort((a,b)=>b.dscr-a.dscr);
  barChart('dscrChart', byDscr.map(p=>p.name), byDscr.map(p=>p.dscr), C.gold, true, (v)=>v.toFixed(2)+'x');
  const byDy = [...DATA.properties].sort((a,b)=>b.debtYield-a.debtYield);
  barChart('dyChart', byDy.map(p=>p.name), byDy.map(p=>p.debtYield), C.blue, true, (v)=>v.toFixed(1)+'%');
}

/* ---------- 9. EXCEPTIONS ---------- */
function buildExceptions() {
  const rows = [];
  DATA.properties.forEach(p => {
    if (p.occ < 92)        rows.push({p, type:'Low Occupancy', metric:fmtPct(p.occ), score:(92-p.occ)});
    if (p.deln > 4)        rows.push({p, type:'High Delinquency', metric:fmtPct(p.deln), score:(p.deln)*4});
    if (p.budgetVar < -1)  rows.push({p, type:'Negative NOI Variance', metric:fmtPct(p.budgetVar), score:-p.budgetVar*3});
    if (p.vacDays > 18)    rows.push({p, type:'High Vacancy Days', metric:p.vacDays+' days', score:(p.vacDays-18)});
    if (p.aged > 230)      rows.push({p, type:'Work Order Delay', metric:p.aged+' aged WOs', score:(p.aged-230)/10});
    if (p.conv < 24)       rows.push({p, type:'Weak Leasing Conversion', metric:fmtPct(p.conv), score:(24-p.conv)});
  });
  rows.forEach(r=>{ r.risk = r.score>=6 ? 'High' : r.score>=2.5 ? 'Medium' : 'Low'; });
  const order = {High:0,Medium:1,Low:2};
  rows.sort((a,b)=> order[a.risk]-order[b.risk] || b.score-a.score);
  return rows.slice(0,12);
}
const ACTIONS = {
  'Low Occupancy':'Launch concession & marketing push; review pricing.',
  'High Delinquency':'Escalate collections; enforce payment plans.',
  'Negative NOI Variance':'Audit opex line items vs budget.',
  'High Vacancy Days':'Accelerate make-ready turn workflow.',
  'Work Order Delay':'Add maintenance capacity; triage aged tickets.',
  'Weak Leasing Conversion':'Coach leasing team; review tour-to-lease funnel.'
};
function renderRiskSummary(rows) {
  const high = rows.filter(r=>r.risk==='High').length;
  const med  = rows.filter(r=>r.risk==='Medium').length;
  const low  = rows.filter(r=>r.risk==='Low').length;
  const tiles = [
    {cls:'t-high', l:'High Priority', v:high, s:'Immediate owner action'},
    {cls:'t-med',  l:'Medium Priority', v:med, s:'Monitor & plan'},
    {cls:'t-low',  l:'Low Priority', v:low, s:'Routine follow-up'},
    {cls:'t-tot',  l:'Total Flags', v:rows.length, s:'Across all properties'}
  ];
  document.getElementById('riskSummary').innerHTML = tiles.map(t=>`
    <div class="risk-tile ${t.cls}"><span class="rt-bar"></span>
      <div class="rt-label">${t.l}</div><div class="rt-val">${t.v}</div><div class="rt-sub">${t.s}</div></div>`).join('');
}
function renderExceptionTable() {
  const rows = buildExceptions();
  renderRiskSummary(rows);
  document.getElementById('exBody').innerHTML = rows.map(r => `
    <tr>
      <td class="prop-name"><span class="pin"></span>${r.p.name}</td>
      <td>${r.p.region}</td>
      <td>${r.type}</td>
      <td class="metric-val">${r.metric}</td>
      <td><span class="risk-badge risk-${r.risk}"><i></i>${r.risk}</span></td>
      <td class="action-txt">${ACTIONS[r.type]}</td>
    </tr>`).join('');
}

/* ---------- 10. CASE STUDY ---------- */
function renderCaseStudy() {
  document.getElementById('caseMetrics').innerHTML = [
    {v:'8→1', l:'Reports unified'},
    {v:'~9 wk', l:'Delivery time'},
    {v:'500', l:'Max properties'},
    {v:'1-click', l:'Refresh'}
  ].map(m=>`<div class="case-metric"><div class="v">${m.v}</div><div class="l">${m.l}</div></div>`).join('');

  const steps = [
    {k:'Discovery', t:'Business Problem', d:'Eight disconnected monthly exports — box scores, financials, delinquency, leasing, work orders, turnover — manually stitched before anyone could answer a basic question.'},
    {k:'Design', t:'Solution Built', d:'A five-layer Excel-native BI platform: raw landing zone, Power Query staging, Power Pivot star schema, DAX measures, and three executive dashboards.'},
    {k:'Model', t:'Data Model', d:'8 fact tables and 6 conformed dimensions, single-direction relationships, and a marked date table powering full time-intelligence.'},
    {k:'Automate', t:'Automation Workflow', d:'Reports arrive by email, land in shared storage, trigger a Power Query refresh, and flow into the model with zero manual steps.'},
    {k:'Deliver', t:'Dashboard Output', d:'KPI cards, trend charts, ranked performers, a color-coded risk register, and full investment-return analysis in one executive view.'},
    {k:'Impact', t:'Business Impact', d:'Manual consolidation replaced by a one-click refresh; same-day visibility into performance, risk and returns; scales from 10 to 500 properties.'}
  ];
  document.getElementById('caseTimeline').innerHTML = steps.map((s,i)=>`
    <div class="case-step"><span class="num">0${i+1}</span>
      <span class="ck">${s.k}</span><h4>${s.t}</h4><p>${s.d}</p></div>`).join('');
}

/* ---------- 11. WORKFLOW ---------- */
function renderWorkflow() {
  const nodes = [
    {i:'📊', t:'Exported Reports', s:'RealPage / PMS'},
    {i:'✉️', t:'Gmail Attachments', s:'Scheduled email'},
    {i:'⚙️', t:'Automation Tool', s:'Power Automate'},
    {i:'📁', t:'Shared Folder', s:'OneDrive / SharePoint'},
    {i:'🔄', t:'Power Query', s:'Clean · map · validate'},
    {i:'🧱', t:'Excel Data Model', s:'Power Pivot + DAX'},
    {i:'📈', t:'Executive Dashboard', s:'One-click refresh'}
  ];
  document.getElementById('flow').innerHTML = nodes.map((n,i)=>{
    const node = `<div class="flow-node"><span class="flow-step">0${i+1}</span>
      <span class="flow-ico">${n.i}</span><div class="t">${n.t}</div><div class="s">${n.s}</div></div>`;
    return i < nodes.length-1 ? node + `<div class="flow-arrow">→</div>` : node;
  }).join('');
}

/* ---------- 12. TECH STACK ---------- */
function renderStack() {
  const stack = [
    {i:'📗', n:'Excel', d:'Delivery & dashboard surface'},
    {i:'🔌', n:'Power Query', d:'M — ingest, clean, map'},
    {i:'🧮', n:'Power Pivot', d:'In-memory data model'},
    {i:'∑', n:'DAX', d:'Measures & time-intelligence'},
    {i:'🤖', n:'Power Automate', d:'Make.com / Zapier alt.'},
    {i:'📉', n:'Chart.js', d:'Interactive web charts'},
    {i:'🗂️', n:'Data Modeling', d:'Star schema design'},
    {i:'🎯', n:'Executive Reporting', d:'Decision support'}
  ];
  document.getElementById('stackGrid').innerHTML = stack.map(s=>`
    <div class="stack-item"><span class="stack-ico">${s.i}</span><span class="n">${s.n}</span><span class="d">${s.d}</span></div>`).join('');
}

/* ---------- 13. NAV SCROLL STATE (no reveal animation) ---------- */
function initInteractions() {
  const nav = document.getElementById('nav');
  window.addEventListener('scroll', ()=> nav.classList.toggle('scrolled', window.scrollY>30), {passive:true});
  document.getElementById('yr').textContent = new Date().getFullYear();
}

/* ---------- 14. INIT ---------- */
document.addEventListener('DOMContentLoaded', ()=>{
  renderHeroBadges();
  renderMock();
  renderExecKpis();
  renderInvestmentKpis();
  renderCharts();
  renderExceptionTable();
  renderCaseStudy();
  renderWorkflow();
  renderStack();
  initInteractions();
});
