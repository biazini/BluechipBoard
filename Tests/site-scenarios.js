/* Scenarios for Test-Site.ps1. This script runs before the page's own script: it can change the embedded data,
   prepare localStorage and register window.__BB_TEST__, which the page calls after building itself.
   Each scenario is chosen with ?s=<name> and runs in a fresh browser profile (except "reopen2"). */
(function(){
const S=new URLSearchParams(location.search).get('s')||'';
window.__R=[];window.__errs=window.__errs||[];
const ok=(n,c,d)=>window.__R.push([n,!!c,d===undefined?'':String(d)]);
const near=(a,b,eps)=>a!=null&&b!=null&&Math.abs(a-b)<=(eps==null?1e-6:eps)*Math.max(1,Math.abs(b));
window.confirm=()=>true;
const el=document.getElementById('dados');
const pair=p=>Array.isArray(p)?p:(p&&Array.isArray(p.value)?p.value:[]);
const arr=x=>Array.isArray(x)?x:(x==null?[]:[x]);
const seed=o=>{localStorage.clear();Object.keys(o).forEach(k=>localStorage.setItem('bb.'+k,JSON.stringify(o[k])));};
const get=k=>{try{return JSON.parse(localStorage.getItem('bb.'+k));}catch(e){return undefined;}};
const $=s=>document.querySelector(s),txt=s=>($(s)||{}).textContent||'';
const V1={app:'Bluechip Board',version:1,exported:'2026-10-03T15:32:10.177Z',portfolio:{SXR8:{q:0,c:0}},lots:[],etfLots:[{id:'musjvneia0ov',d:'2026-10-01',q:1,p:732.58}]};
const histPrice=(d,id,iso)=>{const p=(d.historico[id]||[]).map(pair).find(x=>x[0]===iso);return p?+p[1]:null;};
const lastIso=(d,id)=>{const h=(d.historico[id]||[]).map(pair);return h.length?h[h.length-1][0]:null;};
let G={};
/* dividendos de teste no formato do script (Get-Dividendo) */
const DV=(d,o)=>Object.assign({id:'AAPL',simbolo:'AAPL',moeda:'USD',estado:'ok',fonte:'Yahoo Finance',obtidoEm:d.geradoEm,anualPorAcao:1.08,ttmPorAcao:1.06,frequencia:4,rendimentoPct:0.32,ultimo:['2026-08-10',0.27],pagamentos:[['2026-05-11',0.27],['2026-08-10',0.27]],preco:333.69,precoData:d.geradoEm,nota:'',erro:''},o);
const DV3=d=>({AAPL:DV(d,{}),NVDA:DV(d,{id:'NVDA',simbolo:'NVDA',anualPorAcao:1,rendimentoPct:0.43,ultimo:['2026-09-10',0.25]}),GOOGL:DV(d,{id:'GOOGL',simbolo:'GOOGL',anualPorAcao:0.88,rendimentoPct:0.26,ultimo:['2026-09-04',0.22]})});
const fmtE=v=>'€'+Number(v).toLocaleString('en-GB',{minimumFractionDigits:2,maximumFractionDigits:2});
const divRow=id=>[...document.querySelectorAll('#tbl-div tbody tr')].find(r=>r.textContent.startsWith({AAPL:'Apple',NVDA:'NVIDIA',GOOGL:'Alphabet'}[id]));
/* downloads da página registados em vez de feitos (nome do ficheiro e conteúdo) */
const capDl=()=>{window.__dl=[];window.__blobs=[];const oc=URL.createObjectURL.bind(URL);URL.createObjectURL=b=>{window.__blobs.push(b);return oc(b);};HTMLAnchorElement.prototype.click=function(){window.__dl.push(this.download);};};
const lsSnap=()=>JSON.stringify(Object.keys(localStorage).sort().map(k=>[k,localStorage.getItem(k)]));
/* leitor de CSV (";", aspas duplas) para conferir o que a exportação escreve */
const csvParse=t=>{t=t.replace(/^﻿/,'');const R=[];let row=[],c='',q=false;for(let i=0;i<t.length;i++){const ch=t[i];if(q){if(ch==='"'){if(t[i+1]==='"'){c+='"';i++;}else q=false;}else c+=ch;}else if(ch==='"')q=true;else if(ch===';'){row.push(c);c='';}else if(ch==='\r'&&t[i+1]==='\n'){row.push(c);R.push(row);row=[];c='';i++;}else c+=ch;}if(c||row.length){row.push(c);R.push(row);}return R;};
const csvCol=(R,h)=>{const i=R[0].indexOf(h);return R.slice(1).map(r=>r[i]);};
/* rentabilidade: TIR independente (bisseção simples sobre o valor atual; base = dias de um período, 365 por ano) e dados fixos */
const xirrInd=(fl,base)=>{const f=r=>fl.reduce((s,x)=>s+x[1]/Math.pow(1+r,(x[0]-fl[0][0])/(base*864e5)),0);let lo=-0.999999,hi=1;while(f(hi)>0&&hi<1e15)hi*=2;for(let i=0;i<400;i++){const m=(lo+hi)/2;if(f(m)>0)lo=m;else hi=m;}return(lo+hi)/2;};
const SX={ret:{'2024-11-04':100,'2025-06-02':200,'2026-10-01':250},retsales:{'2025-01-06':100,'2025-06-02':200,'2026-01-05':225,'2026-10-01':250},
 retshort:{'2026-01-05':100,'2026-04-01':60,'2026-10-01':250},retmissing:{'2010-05-19':80,'2025-01-06':100,'2026-10-01':250}};
const retFix=d=>{const P=o=>Object.keys(o).sort().map(k=>[k,o[k]]);d.historico.SXR8=P(SX[S]);d.historico.EUNK=P({'2025-01-06':50,'2026-01-05':60,'2026-04-01':75,'2026-10-01':70});
 [['SXR8',250],['EUNK',70],['BTC',80000]].forEach(([id,v])=>{const a=d.ativos.find(x=>x.id===id);a.pontos=[['2026-10-01',v]];a.parcial=false;a.fonte='Yahoo Finance';});
 G.fim=new Intl.DateTimeFormat('en-CA',{timeZone:'Europe/Lisbon',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(d.geradoEm));};
/* alocação-alvo: preços fixos (SXR8 €250, EUNK €70, IS3N €50, BTC €80,000) e uma carteira de €3,500 (SXR8 €2,000, EUNK €700, BTC €800) */
const tgtFix=d=>[['SXR8',250],['EUNK',70],['IS3N',50],['BTC',80000]].forEach(([id,v])=>{const a=d.ativos.find(x=>x.id===id);a.pontos=[['2026-10-01',v]];a.parcial=false;a.fonte='Yahoo Finance';});
/* política: valores de teste; texto hostil numerado (window.__xss = n se correr); preços e histórico fixos para os alertas de queda */
const POLV=()=>({horizon:'',allocation:'',monthly:'',drop20:'',drop30:'',sell:'',at:'2026-10-01T00:00:00.000Z'});
const HOSTIL=n=>`<img src=x onerror="window.__xss=${n}"><\/textarea><script>window.__xss=${n}<\/script>`;  /* <\/ : o script de cenários vai dentro de uma tag script */
const txt2=e=>e?e.textContent:'';
const polFix=d=>{const a=d.ativos.find(x=>x.id==='AAPL');a.pontos=[['2025-10-06',180],['2026-01-05',200],['2026-10-01',150]];a.parcial=false;
 d.historico.AAPL=[['2010-01-04',100],['2011-03-01',70],['2011-12-01',100],['2015-01-02',120],['2016-02-01',90],['2017-01-03',125],['2020-02-19',150],['2020-03-23',100],['2020-08-03',160],['2026-01-05',200],['2026-10-01',150]];
 const b=d.ativos.find(x=>x.id==='BTC');b.pontos=[['2025-10-02',90000],['2025-12-01',100000],['2026-10-01',68000]];b.parcial=false;d.historico.BTC=[['2014-09-17',400],['2017-12-17',16000],['2018-12-15',3000],['2021-11-10',60000],['2025-12-01',100000],['2026-10-01',68000]];};
/* stress test: carteira fixa (SXR8 €1,000, AAPL $2,000 = €1,600 a 1.25, IS3N €500) e históricos fixos; IS3N só desde 2021 */
const STFX={'2018-09-20':1.17,'2018-12-24':1.14,'2019-04-01':1.12,'2020-02-19':1.08,'2020-03-23':1.07,'2020-06-01':1.11,'2022-01-03':1.13,'2022-10-12':0.97,'2023-06-01':1.07,'2026-09-25':1.25,'2026-09-28':1.25,'2026-09-29':1.25,'2026-09-30':1.25,'2026-10-01':1.25};
const STH={AAPL:{'2018-09-20':50,'2018-12-24':40,'2019-04-01':52,'2020-02-19':80,'2020-03-23':56,'2020-06-01':85,'2022-01-03':180,'2022-10-12':140,'2023-06-01':190,'2026-10-01':200},
 SXR8:{'2018-09-20':45,'2018-12-21':38,'2019-06-03':46,'2020-02-19':60,'2020-03-23':42,'2020-11-02':61,'2022-01-03':90,'2022-10-12':75,'2024-01-02':95,'2026-10-01':250},
 IS3N:{'2021-01-04':30,'2022-01-03':33,'2022-10-12':26,'2026-10-01':50}};
const stFix=d=>{const P=o=>Object.keys(o).sort().map(k=>[k,o[k]]);d.historico.FX=P(STFX);d.fx.yahoo=P(STFX).filter(p=>p[0]>='2026-09-25');Object.keys(STH).forEach(id=>{d.historico[id]=P(STH[id]);});
 [['SXR8',250],['IS3N',50],['AAPL',200]].forEach(([id,v])=>{const a=d.ativos.find(x=>x.id===id);a.pontos=[['2026-09-30',v],['2026-10-01',v]];a.parcial=false;a.fonte='Yahoo Finance';});};
/* simulador de estratégias: série semanal sintética da Bitcoin; cálculo independente das três estratégias (máximo das 52 semanas por força bruta) */
const stratFix=d=>{const P=[];for(let i=0,t0=Date.UTC(2019,0,7);i<=403;i++){const t=t0+i*7*864e5;P.push([new Date(t).toISOString().slice(0,10),Math.round(20000*(1+i/200)*(1+0.35*Math.sin(i/9)))]);}
 d.historico.BTC=P;const a=d.ativos.find(x=>x.id==='BTC');a.pontos=P.slice(-60);a.parcial=false;};
const stratInd=(pts,amt,from,X)=>{const f=Date.parse(from+'-01T00:00:00Z'),P=pts.filter(p=>p[0]>=f),last=pts[pts.length-1][1];let mes=-1,inv=0,uA=0,cash=0,uC=0,n=0,buys=0;
 for(const p of P){const d=new Date(p[0]),m=d.getUTCFullYear()*12+d.getUTCMonth();if(m!==mes){mes=m;inv+=amt;uA+=amt/p[1];cash+=amt;n++;}
  const prev=pts.filter(q=>q[0]<p[0]&&q[0]>p[0]-365*864e5),mx=prev.length?Math.max(...prev.map(q=>q[1])):null;if(cash>0&&mx!=null&&p[1]<=mx*(1-X/100)){uC+=cash/p[1];cash=0;buys++;}}
 return{inv,n,a:uA*last,b:inv/P[0][1]*last,c:uC*last+cash,cash,buys,uA};};
/* retornos rolantes: séries sintéticas (dias úteis para o ETF, todos os dias para a Bitcoin) e cálculo independente */
const serieDias=(ini,fim,uteis,f)=>{const o=[];for(let t=Date.parse(ini+'T00:00:00Z'),e=Date.parse(fim+'T00:00:00Z'),i=0;t<=e;t+=864e5){const w=new Date(t).getUTCDay();if(uteis&&(w===0||w===6))continue;o.push([new Date(t).toISOString().slice(0,10),f(i++)]);}return o;};
const rollFix=d=>{d.historico.SXR8=serieDias('2016-01-04','2026-10-01',true,i=>+(50*Math.pow(1.0003,i)*(1+0.18*Math.sin(i/90))).toFixed(4));
 d.historico.IS3N=serieDias('2023-04-03','2026-10-01',true,i=>+(30*(1+i/2000)*(1+0.1*Math.sin(i/40))).toFixed(4));
 d.historico.BTC=serieDias('2018-01-01','2026-10-01',false,i=>Math.round(8000*Math.pow(1.001,i)*(1+0.4*Math.sin(i/120))));};
const rollInd=(h,anos,k1)=>{const c=h.map(p=>p[1]),k=k1*anos;if(c.length<=k)return null;const r=[];for(let i=k;i<c.length;i++)r.push(c[i]/c[i-k]-1);const s=r.slice().sort((a,b)=>a-b),m=s.length%2?s[(s.length-1)/2]:(s[s.length/2-1]+s[s.length/2])/2,an=v=>Math.pow(1+v,1/anos)-1;
 return{n:r.length,min:s[0],med:m,max:s[s.length-1],neg:r.filter(v=>v<0).length/r.length,ann:anos>1?{min:an(s[0]),med:an(m),max:an(s[s.length-1])}:null};};
/* exposição agregada: preços fixos, agregados de teste (SXR8 com país do índice, EUNK do ficheiro, EUNN Unavailable) */
const AG={SXR8:{paises:[{n:'United States',w:99.8},{n:'Cash/Other',w:0.2}],setores:[{n:'Information Technology',w:40},{n:'Financials',w:30},{n:'Health Care',w:29.8},{n:'Cash/Other',w:0.2}],moedas:[{n:'USD',w:99.8},{n:'Cash/Other',w:0.2}],fontePaises:'index'},
 EUNK:{paises:[{n:'United Kingdom',w:40},{n:'France',w:35},{n:'Switzerland',w:24.5},{n:'Cash/Other',w:0.5}],setores:[{n:'Financials',w:50},{n:'Industrials',w:49.5},{n:'Cash/Other',w:0.5}],moedas:[{n:'GBP',w:40},{n:'EUR',w:35},{n:'CHF',w:24.5},{n:'Cash/Other',w:0.5}],fontePaises:'file'}};
const expoFix=d=>{const P=o=>Object.keys(o).sort().map(k=>[k,o[k]]);d.fx.yahoo=P({'2026-09-25':1.25,'2026-09-28':1.25,'2026-09-29':1.25,'2026-09-30':1.25,'2026-10-01':1.25});
 [['SXR8',250],['EUNK',70],['EUNN',60],['BTC',80000],['AAPL',200],['NVDA',100],['GOOGL',150]].forEach(([id,v])=>{const a=d.ativos.find(x=>x.id===id);a.pontos=[['2026-09-30',v],['2026-10-01',v]];a.parcial=false;a.fonte='Yahoo Finance';});
 d.versao='1.4';d.etfs.SXR8=Object.assign({},d.etfs.SXR8,{aoVivo:true,AAPL:7,NVDA:0,GOOGL:0,agregados:AG.SXR8,top10:[{t:'AAPL',n:'Apple',s:'Information Technology',w:7},{t:'MSFT',n:'Microsoft',s:'Information Technology',w:6},{t:'SHEL',n:'Shell',s:'Energy',w:1}]});
 d.etfs.EUNK=Object.assign({},d.etfs.EUNK,{aoVivo:true,AAPL:0,NVDA:0,GOOGL:0,agregados:AG.EUNK,top10:[{t:'ASML',n:'ASML',s:'Information Technology',w:44},{t:'SHEL',n:'Shell',s:'Energy',w:2}]});
 d.etfs.EUNN=Object.assign({},d.etfs.EUNN,{aoVivo:false,agregados:null,top10:[]});};
const EXSEED=()=>({deleted:{},lots:[{id:'b',d:'2025-01-06',q:0.005,c:300}],sales:[],savedAt:'2026-10-03T10:00:00.000Z',
 buys:[{id:'x',a:'SXR8',d:'2025-01-06',q:4,p:200,u:'2026-10-01'},{id:'k',a:'EUNK',d:'2025-01-06',q:10,p:60,u:'2026-10-01'},{id:'j',a:'EUNN',d:'2025-01-06',q:2,p:50,u:'2026-10-01'},{id:'a',a:'AAPL',d:'2025-01-06',q:5,p:150,u:'2026-10-01'}]});
/* venda simulada e comissões: hoje em Lisboa (como a página) e datas relativas */
const hojeIso=()=>new Intl.DateTimeFormat('en-CA',{timeZone:'Europe/Lisbon',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
const diasAntes=n=>new Date(Date.parse(hojeIso()+'T00:00:00Z')-n*864e5).toISOString().slice(0,10);
const semExport=b=>JSON.stringify(Object.assign({},b,{exported:null}));
const TGTSEED=()=>({deleted:{},sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'t1',a:'SXR8',d:'2025-01-06',q:8,p:200,u:'2026-10-01'},{id:'t2',a:'EUNK',d:'2025-01-06',q:10,p:60,u:'2026-10-01'}],lots:[{id:'tb',d:'2025-01-06',q:0.01,c:600}]});
/* repartição independente, como na especificação: alvos para (total + M); défice = max(0, alvo − atual); Σ défices ≥ M → proporcional; senão cobre e reparte o resto pelos pesos */
const reparteInd=(cur,w,M)=>{const T=Object.keys(cur).reduce((s,k)=>s+cur[k],0),ids=Object.keys(w),def={};ids.forEach(k=>def[k]=Math.max(0,w[k]*(T+M)-(cur[k]||0)));const S=ids.reduce((s,k)=>s+def[k],0),o={};ids.forEach(k=>o[k]=S>=M?(S>0?M*def[k]/S:0):def[k]+(M-S)*w[k]);return o;};
const mesesInd=(cur,w,M,band)=>{const c=Object.assign({},cur);for(let n=0;n<=600;n++){const T=Object.keys(c).reduce((s,k)=>s+c[k],0);if(T>0&&Object.keys(w).every(k=>Math.abs((c[k]||0)/T*100-w[k]*100)<=band+1e-9))return n;const a=reparteInd(c,w,M);Object.keys(a).forEach(k=>c[k]=(c[k]||0)+a[k]);}return null;};
const TAXSEED={deleted:{},savedAt:'2026-10-03T10:00:00.000Z',
 buys:[{id:'a0',a:'AAPL',d:'2025-01-10',q:1,p:200,u:'2026-10-02'},{id:'a1',a:'AAPL',d:'2026-01-05',q:100,p:100,u:'2026-10-02'},{id:'a2',a:'AAPL',d:'2026-02-02',q:50,p:120,u:'2026-10-02'},
  {id:'g1',a:'GOOGL',d:'2026-01-06',q:10,p:150,u:'2026-10-02'},{id:'n1',a:'NVDA',d:'2026-02-02',q:5,p:150,u:'2026-10-02'},{id:'x1',a:'SXR8',d:'2026-01-07',q:3,p:600,u:'2026-10-02'}],
 sales:[{id:'s0',a:'AAPL',d:'2025-03-03',q:1,p:210,u:'2026-10-02'},{id:'s1',a:'AAPL',d:'2026-06-01',q:120,p:150,u:'2026-10-02'},{id:'gs1',a:'GOOGL',d:'2026-03-02',q:3,p:170,u:'2026-10-02'},
  {id:'gs2',a:'GOOGL',d:'2026-05-04',q:7,p:160,u:'2026-10-02'},{id:'=x;"q"',a:'SXR8',d:'2026-04-01',q:1,p:650,u:'2026-10-02'},{id:'bs1',a:'BTC',d:'2026-04-01',q:0.6,p:70000}],
 lots:[{id:'L1',d:'2025-03-01',q:0.5,c:15000},{id:'L2',d:'2026-02-01',q:0.3,c:20000}]};

/* backup da versão 1 (sem ids) com duas compras iguais no mesmo dia */
const V1DUP=()=>({app:'Bluechip Board',version:1,exported:'2026-10-03T15:32:10.177Z',lots:[],etfLots:[{d:'2026-09-01',q:1,p:700},{d:'2026-09-01',q:1,p:700},{d:'2026-09-02',q:2,p:710}]});
const SC={
 newuser:{test(api){
  ok('price cards: the five original assets, the three new ETFs and EUR/USD last',[...document.querySelectorAll('.tk')].map(t=>t.dataset.tk).join(',')==='AAPL,NVDA,GOOGL,SXR8,EUNK,IS3N,EUNN,BTC,FX',[...document.querySelectorAll('.tk')].map(t=>t.dataset.tk).join(','));
  ok('every card shows the date of its price',[...document.querySelectorAll('.tk .asof')].length===9,[...document.querySelectorAll('.tk .asof')].map(x=>x.textContent).join(' / '));
  const fx=api.fx().FX,fxc=document.querySelector('.tk[data-tk="FX"]');
  ok('EUR/USD card: latest rate of the page, 3-month chart, 1 month and year to date',fxc&&txt('.tk[data-tk="FX"] .tk-price')===fx.slice(-1)[0][1].toLocaleString('en-GB',{minimumFractionDigits:4,maximumFractionDigits:4})&&!!fxc.querySelector('.tk-spark svg')&&/Year to date/.test(fxc.textContent)&&/dollars per euro/.test(fxc.textContent),fxc&&fxc.textContent);
  ok('EUR/USD chart still in Currency & Macroeconomics',!!document.querySelector('#ch-fx svg'));
  ok('snapshot layout: S&P 500 first, Bitcoin and EUR/USD full width',getComputedStyle($('.tk[data-tk="SXR8"]')).order==='-1'&&$('.tk[data-tk="BTC"]').classList.contains('tk-wide')&&fxc.classList.contains('tk-wide'));
  ok('portfolio is empty and says so',/Add your purchases/.test(txt('#pfKpis')));
  ok('no backup message for a new user',txt('#bkMsg')==='');
  ok('no backup status box when there is no data',txt('#bkState')==='');
  ok('prices table has 9 rows (8 assets + S&P 500)',$('#tbl-ret tbody').rows.length===9);
  ok('news list rendered',document.querySelectorAll('#newsList article').length>0);
  ok('currency is the Yahoo series',api.fx().FXSRC.startsWith('Yahoo'),api.fx().FXSRC);}},

 v1backup:{prep(d){d.backup=V1;},test(api){
  const b=get('buys')||[];
  ok('old version-1 backup loads into an empty browser',b.length===1&&b[0].a==='SXR8'&&b[0].q===1&&b[0].p===732.58,JSON.stringify(b));
  ok('migrated purchase gets its split reference (r, u)',b[0]&&Array.isArray(b[0].r)&&b[0].r[0]==='2026-10-01'&&/^\d{4}-\d{2}-\d{2}$/.test(b[0].u),JSON.stringify(b[0]));
  ok('browser and file marked as identical',get('savedAt')===V1.exported&&get('fileSaved')===V1.exported,get('savedAt'));
  ok('status: backup file has all the data',/has all your data/.test(txt('#bkState')),txt('#bkState'));
  ok('holdings: 1 SXR8 unit',txt('#pf-q-SXR8')==='1');}},

 reopen1:{prep(d){d.backup=V1;},test(){ok('first visit loads 1 purchase',(get('buys')||[]).length===1);}},
 reopen2:{prep(d){d.backup=V1;},test(){
  ok('reopening: still 1 purchase (no duplicates)',(get('buys')||[]).length===1,JSON.stringify(get('buys')));
  ok('reopening: message says browser matches the file',/matches/.test(txt('#bkMsg')),txt('#bkMsg'));}},

 legacy:{prep(d){d.backup=V1;seed({buys:[{id:'new1',a:'AAPL',d:'2026-09-01',q:2,p:200}],lots:[],savedAt:'2026-10-03T18:00:00.000Z'});},test(){
  const b=get('buys')||[];
  ok('legacy data: purchase deleted before tombstones existed does not come back',b.length===1&&b[0].id==='new1',JSON.stringify(b.map(x=>x.id)));
  ok('legacy data: the deleted id is now remembered','musjvneia0ov' in (get('deleted')||{}));
  ok('legacy data: asks to update the backup file',/not in the backup file/.test(txt('#bkState')),txt('#bkState'));}},

 merge:{prep(d){d.backup={app:'Bluechip Board',version:4,saved:'2026-10-03T12:00:00.000Z',buys:[{id:'A',a:'AAPL',d:'2026-09-01',q:1,p:200},{id:'B',a:'NVDA',d:'2026-09-02',q:3,p:150}],lots:[],sales:[],deleted:{}};
   seed({buys:[{id:'A',a:'AAPL',d:'2026-09-01',q:1,p:200}],lots:[],sales:[],deleted:{},savedAt:'2026-10-03T20:00:00.000Z'});},test(){
  const b=get('buys')||[];
  ok('two browsers: purchase added in the other browser is merged in',b.length===2&&b.some(x=>x.id==='B'),JSON.stringify(b.map(x=>x.id)));
  ok('two browsers: nothing local lost',b.some(x=>x.id==='A'));}},

 tomb:{prep(d){d.backup={app:'Bluechip Board',version:4,saved:'2026-10-03T12:00:00.000Z',buys:[],lots:[],sales:[],deleted:{A:'2026-10-03T11:00:00.000Z'}};
   seed({buys:[{id:'A',a:'AAPL',d:'2026-09-01',q:1,p:200},{id:'C',a:'GOOGL',d:'2026-09-03',q:1,p:150}],lots:[],sales:[],deleted:{},savedAt:'2026-10-02T20:00:00.000Z'});},test(){
  const b=get('buys')||[];
  ok('deleted in the other browser: removed here too',b.length===1&&b[0].id==='C',JSON.stringify(b.map(x=>x.id)));
  ok('deletion message shown',/removed/.test(txt('#bkMsg')),txt('#bkMsg'));}},

 wrongfolder:{prep(d){d.backup={app:'Bluechip Board',version:4,saved:'2026-09-01T10:00:00.000Z',buys:[{id:'A',a:'AAPL',d:'2026-08-01',q:1,p:200}],lots:[],sales:[],deleted:{}};
   seed({buys:[{id:'A',a:'AAPL',d:'2026-08-01',q:1,p:200},{id:'B',a:'AAPL',d:'2026-09-15',q:1,p:210}],lots:[],sales:[],deleted:{},savedAt:'2026-09-20T10:00:00.000Z',fileSaved:'2026-09-20T10:00:00.000Z',fileWrittenAt:'2026-09-20T10:00:05.000Z'});},test(){
  ok('saved file not found by the script: warns about the folder',/was not in the project folder/.test(txt('#bkMsg')),txt('#bkMsg'));}},

 restore:{prep(d){seed({buys:[{id:'A',a:'AAPL',d:'2026-08-01',q:1,p:200}],lots:[],sales:[],deleted:{},savedAt:'2026-09-20T10:00:00.000Z',fileSaved:'2026-09-20T10:00:00.000Z'});},test(api){
  const r=api.juntaBackup({app:'Bluechip Board',version:3,saved:'2026-01-01T00:00:00.000Z',buys:[{id:'old',a:'NVDA',d:'2025-12-01',q:2,p:150}],lots:[{id:'L',d:'2025-12-02',q:0.1,c:8000}]},false);
  const b=get('buys')||[];
  ok('restoring an older backup adds its entries and keeps the newer ones',b.length===2&&b.some(x=>x.id==='A')&&b.some(x=>x.id==='old')&&(get('lots')||[]).length===1,JSON.stringify(b.map(x=>x.id)));
  ok('restoring does not pretend the project file was written',get('fileSaved')==='2026-09-20T10:00:00.000Z',get('fileSaved'));
  ok('restored data is newer than the file (will be saved)',Date.parse(get('savedAt'))>Date.parse('2026-09-20T10:00:00.000Z'),get('savedAt'));}},

 migrate:{prep(d){G.p=histPrice(d,'AAPL','2026-09-01');seed({buys:[{id:'m1',a:'AAPL',d:'2026-09-01',q:2,p:200}],lots:[],savedAt:'2026-10-03T18:00:00.000Z'});},test(){
  const b=(get('buys')||[])[0]||{};
  ok('existing purchase gets base date = date of last save',b.u==='2026-10-03',JSON.stringify(b));
  ok('existing purchase gets the reference close of its date',Array.isArray(b.r)&&b.r[0]==='2026-09-01'&&near(b.r[1],G.p),JSON.stringify(b.r)+' vs '+G.p);
  ok('migration does not change the last-change time',get('savedAt')==='2026-10-03T18:00:00.000Z',get('savedAt'));}},

 split:{prep(d){
   /* simula um desdobramento futuro de 10:1 da NVIDIA: o Yahoo divide todo o histórico por 10 */
   const iso='2025-01-02';G.old=histPrice(d,'NVDA',iso);G.lastOld=+pair(d.ativos.find(a=>a.id==='NVDA').pontos.slice(-1)[0])[1];
   const div=L=>L.map(pair).map(p=>[p[0],+p[1]/10]);d.historico.NVDA=div(d.historico.NVDA);const a=d.ativos.find(x=>x.id==='NVDA');a.pontos=div(a.pontos);
   d.splits=d.splits||{};d.splits.NVDA=(d.splits.NVDA||[]).map(pair).concat([['2026-10-01',10]]);
   G.newLast=+a.pontos.slice(-1)[0][1];G.newRef=histPrice(d,'NVDA','2026-10-02');
   seed({deleted:{},lots:[],savedAt:'2026-09-30T10:00:00.000Z',buys:[
    {id:'n1',a:'NVDA',d:iso,q:2,p:100,r:[iso,G.old],u:'2026-09-30'},
    {id:'n2',a:'NVDA',d:iso,q:3,p:50,u:'2026-09-30'},
    {id:'n3',a:'NVDA',d:'2026-10-02',q:1,p:20,r:['2026-10-02',G.newRef],u:'2026-10-02'},
    {id:'n4',a:'NVDA',d:iso,q:1,p:10,r:[iso,G.old*1.5],u:'2026-10-02'}]});},test(api){
  const B=(get('buys')||[]).map(api.efetiva),f=id=>B.find(x=>x.id===id)||{};
  ok('split after entry, found through the reference price: q×10, p÷10',near(f('n1').q,20)&&near(f('n1').p,10),JSON.stringify(f('n1')));
  ok('split after entry, found through the base date: q×10, p÷10',near(f('n2').q,30)&&near(f('n2').p,5),JSON.stringify(f('n2')));
  ok('purchase entered after the split is not adjusted',near(f('n3').q,1)&&f('n3').k===1,JSON.stringify(f('n3')));
  ok('unexplained change in history: warning, no adjustment',f('n4').k===1&&/changed/.test(f('n4').aviso||''),JSON.stringify(f('n4')));
  const P=api.pfDados().find(x=>x.id==='NVDA');
  ok('amount invested unchanged by the split',near(P.c,200+150+20+10),P.c);
  const lp=api.ATIVOS.find(a=>a.id==='NVDA').pts.slice(-1)[0],fx=api.fxAt(lp[0]),antes=2*G.lastOld/fx,agora=f('n1').q*api.lastEur('NVDA');
  ok('value of the pre-split purchase is the same as before the split',near(agora,antes,1e-9),`${agora} vs ${antes}`);
  ok('purchases table notes the adjustment',/adjusted for a 10:1 split/.test(txt('#tbl-buys')));}},

 fxlong:{prep(d){d.fx.yahoo=[];},test(api){
  ok('Yahoo 1-year EUR/USD missing: long history used',api.fx().FXSRC==='long-term history',api.fx().FXSRC);
  const a=api.ATIVOS.find(x=>x.id==='AAPL'),e=api.inCur(a,'EUR');
  ok('euro series still complete',e.length===a.pts.length,e.length+' / '+a.pts.length);
  ok('alert says which rate is used',[...document.querySelectorAll('#alertas .alert-msg')].some(x=>/long-term history/.test(x.textContent)));}},

 fxnone:{prep(d){d.fx.yahoo=[];d.fx.bce=[];d.historico.FX=[];
   seed({deleted:{},lots:[],buys:[{id:'z',a:'AAPL',d:'2026-09-01',q:1,p:200,u:'2026-10-02'}],savedAt:'2026-10-03T10:00:00.000Z'});},test(api){
  const a=api.ATIVOS.find(x=>x.id==='AAPL');
  ok('no EUR/USD at all: no euro series (never dollars labelled as euros)',api.inCur(a,'EUR').length===0);
  ok('no EUR/USD: no euro price for AAPL',api.lastEur('AAPL')===null);
  ok('no EUR/USD: red alert',[...document.querySelectorAll('#alertas .alert.l-red .alert-msg')].some(x=>/No EUR\/USD/.test(x.textContent)));
  ok('no EUR/USD: portfolio value not shown for AAPL',/no price/.test(txt('#pf-v-AAPL')),txt('#pf-v-AAPL'));
  ok('no EUR/USD: prices table explains it',/No EUR\/USD rate/.test(txt('#tbl-ret')));
  ok('no EUR/USD: the snapshot card says so, no rate invented',/No EUR\/USD rate in this run/.test(txt('.tk[data-tk="FX"]'))&&!document.querySelector('.tk[data-tk="FX"] .tk-price'));}},

 fxecb:{prep(d){d.fx.yahoo=[];d.historico.FX=[];},test(api){
  ok('ECB fallback used',/^ECB/.test(api.fx().FXSRC),api.fx().FXSRC);
  const s=api.stats(api.inCur(api.ATIVOS.find(x=>x.id==='AAPL'),'EUR'));
  ok('ECB (90 days): no "1 year" or "52-week" figures in euros',s&&s.y1===null&&s.hi===null&&s.dist===null,JSON.stringify(s&&{y1:s.y1,hi:s.hi,dist:s.dist}));
  ok('ECB (90 days): short-period figures still shown',s&&s.d1!=null&&s.m1!=null);}},

 stale:{prep(d){const a=d.ativos.find(x=>x.id==='AAPL');a.pontos=a.pontos.slice(0,-3);a.parcial=false;},test(api){
  const a=api.ATIVOS.find(x=>x.id==='AAPL'),f=api.frescura('AAPL',a.pts);
  /* 3 pontos a menos = 3 sessões em falta se os dados acabavam numa sessão fechada, 2 se o último era intradiário (sessão aberta) */
  ok('price 2–3 sessions old is detected',f&&f.velho&&f.falta>=2&&f.falta<=3,JSON.stringify(f));
  ok('stale alert raised',[...document.querySelectorAll('#alertas .alert-msg')].some(x=>new RegExp(`Apple: the latest price is from .* ${f&&f.falta} sessions behind`).test(x.textContent)));
  ok('card shows the old date in warning colour',!!document.querySelector('.tk[data-tk="AAPL"] .asof.old'));
  const g=api.frescura('NVDA',api.ATIVOS.find(x=>x.id==='NVDA').pts);
  ok('current price is not flagged',g&&!g.velho,JSON.stringify(g));}},

 prev:{prep(d){d.ativos.find(x=>x.id==='SXR8').fonte='previous run (data up to 2026-10-02)';},test(){
  ok('previous-run prices are flagged',[...document.querySelectorAll('#alertas .alert-msg')].some(x=>/previous run's data/.test(x.textContent)));
  ok('card says previous run',/previous run/.test(txt('.tk[data-tk="SXR8"] .asof')));}},

 fifo:{prep(d){seed({deleted:{},savedAt:'2026-10-03T10:00:00.000Z',
   lots:[{id:'L1',d:'2024-01-10',q:0.5,c:15000},{id:'L2',d:'2026-06-01',q:0.3,c:20000}],
   buys:[{id:'b1',a:'SXR8',d:'2025-01-02',q:2,p:500,u:'2026-10-02'}],
   sales:[{id:'s1',a:'BTC',d:'2026-07-01',q:0.6,p:70000},{id:'s2',a:'SXR8',d:'2025-06-02',q:1,p:600,u:'2026-10-02'}]});},test(api){
  const C=api.carteira(),L=C.lotes.BTC,v=C.vendas.find(x=>x.id==='s1');
  ok('FIFO: oldest Bitcoin purchase used up first',near(L.find(x=>x.id==='L1').q,0)&&near(L.find(x=>x.id==='L2').q,0.2),JSON.stringify(L.map(x=>[x.id,x.q])));
  ok('FIFO: cost of the sale',near(v.custo,15000+0.1*20000/0.3),v.custo);
  ok('FIFO: realised gain',near(v.real,0.6*70000-(15000+0.1*20000/0.3)),v.real);
  const P=api.pfDados(),btc=P.find(x=>x.id==='BTC'),sx=P.find(x=>x.id==='SXR8');
  ok('holdings: 0.2 BTC left, cost of what is left',near(btc.q,0.2)&&near(btc.c,0.2*20000/0.3),JSON.stringify([btc.q,btc.c]));
  ok('holdings: 1 SXR8 unit left, realised +€100',near(sx.q,1)&&near(sx.c,500)&&near(sx.real,100),JSON.stringify([sx.q,sx.c,sx.real]));
  ok('tax counter: sold purchase is not counted as tax-free',/Sold/.test(txt('#tbl-lots'))&&/^0 BTC/.test(txt('#lotSum .kpi:nth-child(2) .v')),txt('#lotSum .kpi:nth-child(2) .v'));
  const x=api.carteira({id:'x',a:'BTC',d:'2026-07-02',q:5,p:1}).vendas.find(y=>y.id==='x');
  ok('selling more than held is detected',x.falta>0,x.falta);}},

 cal:{test(api){
  const hoje=new Intl.DateTimeFormat('en-CA',{timeZone:'Europe/Lisbon'}).format(new Date());
  ok('today (Lisbon) is 0 days away',api.diasAte(Date.parse(hoje+'T00:00:00Z'))===0);
  ok('Xetra: last trading day of 2026 is shortened to 14:00',api.regras('DE',2026).cur['2026-12-30']==='14:00');
  ok('Xetra: 31 Dec and 24–26 Dec are holidays',['2026-12-24','2026-12-25','2026-12-31'].every(x=>api.regras('DE',2026).fer.has(x)));
  ok('NYSE: 3 Jul 2026 holiday (4 Jul is a Saturday)',api.regras('US',2026).fer.has('2026-07-03'));
  ok('NYSE: Good Friday 2027 (26 Mar) computed by rule',api.regras('US',2027).fer.has('2027-03-26'));
  ok('NYSE: early close the day after Thanksgiving 2026',api.regras('US',2026).cur['2026-11-27']==='13:00');
  const b={id:'US',tz:'America/New_York',abre:'09:30',fecha:'16:00',feriados:[],curtos:{}};
  const s=api.sessao(b,Date.parse('2026-11-01T06:30:00Z'));   /* domingo em que acaba a hora de verão nos EUA */
  ok('NYSE session after the US clock change opens at 14:30 UTC',s&&new Date(s.a).toISOString()==='2026-11-02T14:30:00.000Z',s&&new Date(s.a).toISOString());
  const s2=api.sessao(b,Date.parse('2026-10-28T12:00:00Z'));   /* semana em que a Europa já mudou e os EUA ainda não */
  ok('NYSE session in the week Europe and the US differ opens at 13:30 UTC',s2&&new Date(s2.a).toISOString()==='2026-10-28T13:30:00.000Z',s2&&new Date(s2.a).toISOString());}},

 stats:{test(api){
  const pts=[];let t=Date.parse('2025-10-03T00:00:00Z'),i=0;while(t<=Date.parse('2026-10-02T00:00:00Z')){const w=new Date(t).getUTCDay();if(w!==0&&w!==6){pts.push([t,100*Math.pow(1.001,i)]);i++;}t+=864e5;}
  const s=api.stats(pts),n=pts.length;
  ok('1-year return',near(s.y1,(Math.pow(1.001,n-1)-1)*100,1e-9),s.y1);
  ok('1-session return',near(s.d1,0.1,1e-9),s.d1);
  ok('1-month return = 21 sessions',near(s.m1,(Math.pow(1.001,21)-1)*100,1e-9),s.m1);
  ok('50-session average',near(s.ma50,pts.slice(-50).reduce((a,p)=>a+p[1],0)/50,1e-12));
  ok('volatility of a constant-growth series is 0',near(s.vol,0,1e-9),s.vol);
  ok('52-week high is the last price',near(s.hi,pts[n-1][1])&&near(s.dist,0,1e-9));
  const y0=pts.findIndex(p=>new Date(p[0]).getUTCFullYear()===2026);
  ok('year to date from the last close of 2025',near(s.ytd,(pts[n-1][1]/pts[y0-1][1]-1)*100,1e-9),s.ytd);
  const dd=pts.map((p,j)=>[p[0],j===150?p[1]*0.7:p[1]]),s2=api.stats(dd);
  ok('max drawdown',s2.mdd<-29&&s2.mdd>-31,s2.mdd);
  const a=api.ATIVOS.find(x=>x.id==='AAPL'),last=a.pts[a.pts.length-1],f=api.fxAt(last[0]);
  ok('AAPL in euros = dollar close ÷ EUR/USD of the same day',near(api.lastEur('AAPL'),last[1]/f,1e-12),`${api.lastEur('AAPL')} vs ${last[1]/f}`);
  ok('AAPL card: euro price uses the same rate as the portfolio',txt('.tk[data-tk="AAPL"] .tk-eur').includes(api.lastEur('AAPL').toLocaleString('en-GB',{minimumFractionDigits:2,maximumFractionDigits:2})),txt('.tk[data-tk="AAPL"] .tk-eur'));}},

 xss:{prep(d){seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'h1',a:'SXR8',d:'2026-01-05',q:1,p:600,u:'2026-10-01'}],
   notes:{h1:{t:HOSTIL(5),at:'2026-10-03T10:00:00.000Z'}},policy:Object.assign(POLV(),{horizon:HOSTIL(6),drop20:HOSTIL(7),at:'2026-10-03T10:00:00.000Z'})});},test(){
  ok('notes and policy typed by the user: no script ran, no element injected, shown as text',window.__xss===undefined&&!document.querySelector('#tbl-buys img, #tbl-buys svg[onload], #polForm img, #alertas img')&&txt('#tbl-buys').includes('<img src=x')&&$('#pol-horizon').value.includes('<img src=x'),String(window.__xss));
  ok('no script from a headline ran',window.__xss===undefined,String(window.__xss));
  ok('no injected element',!document.querySelector('#newsList img, #newsList svg[onload], #newsList script'));
  ok('javascript: links neutralised',![...document.querySelectorAll('#newsList a, .dups a')].some(a=>/^javascript:/i.test(a.getAttribute('href')||'')));
  ok('the headline is shown as text',[...document.querySelectorAll('#newsList .news-t')].some(a=>a.textContent.includes('<img src=x')));}},

 /* ---------- dividendos ---------- */
 div:{prep(d){d.dividendos=DV3(d);seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',
   buys:[{id:'a1',a:'AAPL',d:'2026-01-05',q:10,p:200,u:'2026-10-02'},{id:'a2',a:'AAPL',d:'2026-03-02',q:5,p:210,u:'2026-10-02'},{id:'n1',a:'NVDA',d:'2026-02-02',q:2.5,p:150,u:'2026-10-02'},{id:'x1',a:'SXR8',d:'2026-02-02',q:2,p:600,u:'2026-10-02'}],
   sales:[{id:'s1',a:'AAPL',d:'2026-06-01',q:4,p:220,u:'2026-10-02'}]});},test(api){
  const X=api.dividendos(),f=id=>X.L.find(x=>x.id===id),P=api.pfDados(),fx=api.fx().FX.slice(-1)[0][1];
  ok('dividends: shares held are the Portfolio holdings (2 purchases − 1 partial sale = 11 AAPL)',near(f('AAPL').q,11)&&f('AAPL').q===P.find(x=>x.id==='AAPL').q,f('AAPL').q);
  ok('dividends: AAPL projected = 11 × $1.08',near(f('AAPL').usd,11.88),f('AAPL').usd);
  ok('dividends: fractional holding, 2.5 NVDA × $1.00 = $2.50',near(f('NVDA').usd,2.5),f('NVDA').usd);
  ok('dividends: euros at the latest EUR/USD of the page',near(f('AAPL').eur,Math.round(11*1.08/fx*100)/100,1e-12)&&X.fx[1]===fx,`${f('AAPL').eur} at ${fx}`);
  ok('dividends: zero holdings (GOOGL) shows "No holdings", not €0',f('GOOGL').usd===null&&/No holdings/.test(divRow('GOOGL').textContent)&&!/€0\.00/.test(divRow('GOOGL').textContent),divRow('GOOGL').textContent);
  ok('dividends: SXR8 is not in the table (accumulating ETF)',document.querySelectorAll('#tbl-div tbody tr').length===3&&!/ETF SXR8/.test(txt('#tbl-div')));
  ok('dividends: total = sum of the rows shown',near(X.totEur,Math.round((f('AAPL').eur+f('NVDA').eur)*100)/100,1e-12)&&txt('#div-teur')===fmtE(X.totEur),`${txt('#div-teur')} vs ${fmtE(X.totEur)}`);
  ok('dividends: dollar total = $14.38',txt('#div-tusd')==='$14.38',txt('#div-tusd'));
  ok('dividends: KPI shows the projected total and the rate used',txt('#divKpis').includes(fmtE(X.totEur))&&txt('#divKpis').includes(fx.toLocaleString('en-GB',{minimumFractionDigits:4,maximumFractionDigits:4})),txt('#divKpis'));
  ok('dividends: yield and latest payment shown',/0\.43%/.test(divRow('NVDA').textContent)&&/latest \$0\.2500/.test(divRow('NVDA').textContent),divRow('NVDA').textContent);}},

 divnone:{prep(d){d.dividendos=DV3(d);},test(api){
  ok('no holdings: says there is nothing to project',/You hold no Apple, NVIDIA or Alphabet/.test(txt('#divKpis')),txt('#divKpis'));
  ok('no holdings: no €0 total',txt('#div-teur')==='—'&&!/€0/.test(txt('#tbl-div')),txt('#div-teur'));
  ok('no holdings: dividend per share still listed',/\$1\.08/.test(txt('#tbl-div')));}},

 divfail:{prep(d){d.dividendos={AAPL:{id:'AAPL',estado:'error',erro:'blocked by test',obtidoEm:d.geradoEm},NVDA:DV(d,{id:'NVDA',anualPorAcao:'abc'}),GOOGL:DV(d,{id:'GOOGL',anualPorAcao:-1})};
   seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a',a:'AAPL',d:'2026-01-05',q:1,p:200,u:'2026-10-02'},{id:'n',a:'NVDA',d:'2026-01-05',q:1,p:100,u:'2026-10-02'},{id:'g',a:'GOOGL',d:'2026-01-05',q:1,p:100,u:'2026-10-02'}]});},test(api){
  ok('source failed / invalid values: every row "Unavailable"',['AAPL','NVDA','GOOGL'].every(id=>/Unavailable/.test(divRow(id).textContent)),txt('#tbl-div'));
  ok('the failure reason is shown',/Source failed: blocked by test/.test(divRow('AAPL').textContent));
  ok('no total and no €0 when nothing can be projected',txt('#div-teur')==='—'&&!/€0/.test(txt('#tbl-div'))&&/unavailable/.test(txt('#divKpis')),txt('#divKpis'));
  ok('portfolio still works',txt('#pf-q-AAPL')==='1'&&txt('#pf-q-NVDA')==='1');}},

 divbad:{prep(d){d.dividendos='garbage';seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a',a:'AAPL',d:'2026-01-05',q:1,p:200,u:'2026-10-02'}]});},test(){
  ok('malformed dividend data: page builds, rows unavailable',/Unavailable/.test(divRow('AAPL').textContent)&&txt('#pf-q-AAPL')==='1');}},

 divstale:{prep(d){const g=Date.parse(d.geradoEm),iso=n=>new Date(g-n*864e5).toISOString();
   d.dividendos={AAPL:DV(d,{estado:'previous run',fonte:'previous run (retrieved '+iso(30).slice(0,10)+')',obtidoEm:iso(30)}),NVDA:DV(d,{id:'NVDA',obtidoEm:iso(200)}),GOOGL:DV(d,{id:'GOOGL',anualPorAcao:0.88})};
   seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a',a:'AAPL',d:'2026-01-05',q:2,p:200,u:'2026-10-02'},{id:'n',a:'NVDA',d:'2026-01-05',q:1,p:100,u:'2026-10-02'}]});},test(api){
  const X=api.dividendos();
  ok('previous-run dividends are used and labelled with their date',X.L[0].usd!=null&&/previous run \(retrieved/.test(divRow('AAPL').textContent)&&!!divRow('AAPL').querySelector('.td-sub.old'),divRow('AAPL').textContent);
  ok('dividend data over 180 days old is not used',X.L[1].usd===null&&/too old/.test(divRow('NVDA').textContent),divRow('NVDA').textContent);
  ok('the stale asset is named in the KPI',/without NVIDIA/.test(txt('#divKpis')),txt('#divKpis'));}},

 divfx:{prep(d){d.dividendos=DV3(d);d.fx.yahoo=[];d.fx.bce=[];d.historico.FX=[];
   seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a',a:'AAPL',d:'2026-01-05',q:1,p:200,u:'2026-10-02'}]});},test(api){
  const X=api.dividendos();
  ok('no EUR/USD: dollars shown, euros "FX unavailable" (never dollars as euros)',X.L[0].usd===1.08&&X.L[0].eur===null&&/FX unavailable/.test(divRow('AAPL').textContent),divRow('AAPL').textContent);
  ok('no EUR/USD: euro total unavailable',txt('#div-teur')==='FX unavailable'&&/unavailable/.test(txt('#divKpis')),txt('#div-teur'));}},

 /* ---------- dividendos no anexo J (Quadro 8A) e líquido estimado ---------- */
 div8a:{prep(d){capDl();const S=o=>Object.keys(o).sort().map(k=>[k,o[k]]);
   d.historico.FX=S({'2025-09-08':1.17,'2025-11-14':1.16,'2025-12-08':1.165,'2026-02-09':1.18,'2026-03-09':1.17,'2026-05-11':1.125,'2026-08-11':1.16,'2026-10-01':1.17});
   d.dividendos={AAPL:DV(d,{pagamentos:[['2025-11-10',0.26],['2026-02-09',0.26],['2026-05-11',0.26],['2026-08-11',0.27],[diasAntes(-20),0.27]]}),
    NVDA:DV(d,{id:'NVDA',simbolo:'NVDA',pagamentos:[['2026-03-11',0.01]]}),
    GOOGL:DV(d,{id:'GOOGL',simbolo:'GOOGL',anualPorAcao:0.84,pagamentos:[['2025-09-08',0.21],['2025-11-16',0.21],['2025-12-08',0.21],['2026-03-09',0.21]]})};
   seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',
    buys:[{id:'a1',a:'AAPL',d:'2026-01-05',q:10,p:200,u:'2026-10-02'},{id:'a2',a:'AAPL',d:'2026-05-11',q:5,p:210,u:'2026-10-02'},{id:'g1',a:'GOOGL',d:'2025-06-02',q:4,p:150,u:'2026-10-02'}],
    sales:[{id:'s1',a:'AAPL',d:'2026-06-01',q:4,p:220,u:'2026-10-02'},{id:'g2',a:'GOOGL',d:'2026-03-09',q:4,p:170,u:'2026-10-02'}]});},test(api){
  const X=api.anexo8A(2026),r=(id,ex)=>X.L.find(x=>x.id===id&&x.ex===ex)||{},R=csvParse(X.csv),col=h=>csvCol(R,h),i=(id,ex)=>X.L.indexOf(r(id,ex));
  ok('one row per payment while held: AAPL 3, GOOGL 1; none before the first purchase or for a stock never held (NVDA)',X.L.length===4&&X.L.filter(x=>x.id==='AAPL').length===3&&!X.L.some(x=>x.id==='NVDA'||x.ex==='2025-11-10'),JSON.stringify(X.L.map(x=>[x.id,x.ex,x.q])));
  ok('shares entitled: bought before the ex-date count, bought on the ex-date do not (10 on 11 May, not 15)',near(r('AAPL','2026-02-09').q,10)&&near(r('AAPL','2026-05-11').q,10),JSON.stringify(X.L.map(x=>[x.ex,x.q])));
  ok('a sale after the ex-date does not reduce it; the next payment uses what was left (15 − 4 = 11 on 11 Aug)',near(r('AAPL','2026-08-11').q,11));
  ok('a sale on the ex-date still receives the dividend (GOOGL, 4 shares on 9 Mar)',near(r('GOOGL','2026-03-09').q,4));
  ok('a future ex-date is not exported',X.L.every(x=>x.ex<=diasAntes(0))&&!X.L.some(x=>x.ex===diasAntes(-20)));
  const a=r('AAPL','2026-08-11'),c2=v=>Math.round(v*100)/100;
  ok('gross in USD = shares × dividend; US tax withheld estimated at 15%',a.bruto===2.97&&a.ret===c2(2.97*0.15)&&r('AAPL','2026-02-09').bruto===2.6&&r('AAPL','2026-02-09').ret===c2(2.6*0.15),JSON.stringify(a));
  ok('euros only as a reference, at the EUR/USD of the ex-date (fxAt)',a.fx&&a.fx.d==='2026-08-11'&&a.fx.fx===api.fxAt(Date.parse('2026-08-11T00:00:00Z'),api.HIST.FX)&&a.eur===c2(2.97/1.16)&&a.retEur===c2(a.ret/1.16),JSON.stringify(a));
  ok('code and country from the configurable mapping: Quadro 8A, E11, 840',X.L.every(x=>x.codigo==='E11'&&x.pais==='840')&&col('Anexo J table').every(v=>v==='8A')&&col('Código rendim.').every(v=>v==='E11')&&api.ANEXO_J.q8a.codigo==='E11');
  ok('BROKER_FX on every row',X.L.every(x=>x.flags.includes('BROKER_FX')));
  const Y=api.anexo8A(2025),g=ex=>Y.L.find(x=>x.ex===ex)||{};
  ok('ex-date in December (last 45 days of the year): PAY_DATE_UNKNOWN and REVIEW',Y.L.length===3&&g('2025-12-08').flags.includes('PAY_DATE_UNKNOWN')&&g('2025-12-08').status==='REVIEW',JSON.stringify(Y.L.map(x=>[x.ex,x.status,x.flags])));
  ok('16 Nov (45 days before 31 Dec, outside the window) and September: no PAY_DATE_UNKNOWN',!g('2025-11-16').flags.includes('PAY_DATE_UNKNOWN')&&g('2025-09-08').status==='OK');
  ok('CSV: UTF-8 BOM, ";" separator, CRLF, every row has the header\'s columns',X.csv.charCodeAt(0)===0xFEFF&&X.csv.split('\r\n')[0].replace(/^﻿/,'').split(';')[0]==='Tax year'&&!/[^\r]\n/.test(X.csv)&&R.length===5&&R.every(x=>x.length===R[0].length));
  ok('CSV: decimal comma, no currency symbol',col('Gross amount (USD)')[i('AAPL','2026-08-11')]==='2,97'&&col('Dividend per share (USD)')[i('AAPL','2026-08-11')]==='0,270000'&&!/€|\$/.test(X.csv.split('\r\n').slice(1).join('')),col('Gross amount (USD)').join('|'));
  ok('CSV: Imposto retido em Portugal left blank, to fill in',col('Imposto retido em Portugal – Retenção na fonte (EUR)').every(v=>v==='')&&col('Imposto retido em Portugal – NIF da entidade retentora').every(v=>v===''));
  ok('CSV: deterministic (same data, same file)',api.anexo8A(2026).csv===X.csv);
  const P=api.dividendos(),f=P.L.find(x=>x.id==='AAPL');
  ok('projection: new "Net (est.)" column = gross × 0.72, with its own total',/Net \(est\.\)/.test(txt('#tbl-div thead'))&&f.liq===c2(f.eur*0.72)&&divRow('AAPL').textContent.includes(fmtE(f.liq))&&txt('#div-tliq')===fmtE(P.totLiq),txt('#tbl-div'));
  ok('projection: the existing columns and totals are unchanged',f.usd===11.88&&txt('#div-tusd')==='$11.88'&&txt('#div-teur')===fmtE(P.totEur)&&document.querySelectorAll('#tbl-div thead th').length===7);
  ok('the note explains the assumption: 15% withheld in the US plus the top-up to 28% in Portugal',/15%/.test(txt('#divNetNote'))&&/28%/.test(txt('#divNetNote'))&&/0\.72/.test(txt('#divNetNote')),txt('#divNetNote'));
  ok('dividend year list: years with payments while held, latest first',[...document.querySelectorAll('#divYear option')].map(o=>o.value).join(',')==='2026,2025');
  G.ls=lsSnap();$('#divYear').value='2026';$('#divExport').click();
  ok('separate button: one file, AnexoJ_Dividends_2026.csv',JSON.stringify(window.__dl)==='["AnexoJ_Dividends_2026.csv"]'&&/Downloaded AnexoJ_Dividends_2026\.csv/.test(txt('#divMsg'))&&window.__blobs.length===1,JSON.stringify(window.__dl)+' '+txt('#divMsg'));
  window.__dl=[];$('#taxYear').value='2026';$('#taxExport').click();},
  after(){ok('the existing export still downloads exactly its two files',JSON.stringify(window.__dl)==='["AnexoJ_Stocks_ETFs_2026.csv","AnexoJ_Crypto_2026.csv"]',JSON.stringify(window.__dl));
   ok('the dividend export changes nothing in the browser storage',lsSnap()===G.ls);}},

 div8afx:{prep(d){d.fx.yahoo=[];d.fx.bce=[];d.historico.FX=[];d.dividendos={AAPL:DV(d,{pagamentos:[['2026-02-09',0.26]]})};
   seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a1',a:'AAPL',d:'2026-01-05',q:10,p:200,u:'2026-10-02'}]});},test(api){
  const X=api.anexo8A(2026),r=X.L[0]||{},R=csvParse(X.csv);
  ok('no EUR/USD: dollars kept, euro columns blank (never dollars as euros), flagged',X.L.length===1&&r.bruto===2.6&&r.eur===null&&r.retEur===null&&r.flags.includes('NO_FX_REFERENCE')&&csvCol(R,'Rendimento bruto (EUR, reference)')[0]===''&&csvCol(R,'EUR/USD at the ex-date')[0]==='',JSON.stringify(r));
  ok('no EUR/USD: Net (est.) says FX unavailable, no €0',/FX unavailable/.test(divRow('AAPL').textContent)&&txt('#div-tliq')==='FX unavailable'&&!/€0/.test(txt('#tbl-div')),txt('#tbl-div'));}},

 div8anone:{prep(d){capDl();d.dividendos={AAPL:{id:'AAPL',estado:'error',erro:'blocked by test',obtidoEm:d.geradoEm},GOOGL:'garbage'};
   seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a1',a:'AAPL',d:'2026-01-05',q:10,p:200,u:'2026-10-02'},{id:'g1',a:'GOOGL',d:'2026-01-05',q:1,p:150,u:'2026-10-02'}]});},test(api){
  const X=api.anexo8A(2026);
  ok('dividend data unavailable: no rows invented, each reason kept',X.L.length===0&&X.falta.length===2&&/blocked by test/.test((X.falta.find(x=>x.id==='AAPL')||{}).motivo),JSON.stringify(X.falta));
  $('#divExport').click();
  ok('the export still works (headers only) and the message names what is missing',window.__dl[0]==='AnexoJ_Dividends_'+$('#divYear').value+'.csv'&&/unavailable/i.test(txt('#divMsg'))&&/Apple/.test(txt('#divMsg'))&&X.csv.split('\r\n').filter(Boolean).length===1,JSON.stringify(window.__dl)+' '+txt('#divMsg'));
  ok('Net (est.) unavailable too, no €0',/Unavailable/.test(divRow('AAPL').textContent)&&!/€0/.test(txt('#tbl-div')));}},

 /* ---------- notícias: "Relevance to me" e "Via top holding" ---------- */
 newsrel:{prep(d){tgtFix(d);const a=d.ativos.find(x=>x.id==='AAPL');a.pontos=[['2026-10-01',200]];a.parcial=false;
   const N=(t,emp,score,o)=>Object.assign({chave:t.toLowerCase(),titulo:t,link:'https://example.com/'+encodeURIComponent(t),fonte:'Reuters',dominio:'reuters.com',data:d.geradoEm,feed:'Google News: test',empresas:emp,temas:[],nivel:'yellow',score,severo:false,sentimento:'neutro',tier:'referencia',novo:false,outras:[],soFeed:false,viaPosicao:[]},o||{});
   d.noticias=[N('Story about Apple',['AAPL'],3.0),N('Story about the market',['MKT'],3.5),N('Story about Bitcoin',['BTC'],3.4,{nivel:'orange'}),N('Story about emerging markets',['IS3N'],3.6),
    N('Story about ASML',['EUNK'],3.1,{viaPosicao:['ASML']}),N('Story about Apple and NVIDIA',['AAPL','NVDA'],2.0,{nivel:'white'}),N('Story about the S&P 500 ETF',['SXR8','MKT'],3.3)];
   seed({deleted:{},sales:[],savedAt:'2026-10-03T10:00:00.000Z',lots:[{id:'L1',d:'2026-01-05',q:0.00375,c:250}],
    buys:[{id:'x1',a:'SXR8',d:'2026-01-05',q:4,p:200,u:'2026-10-02'},{id:'e1',a:'EUNK',d:'2026-01-05',q:10,p:60,u:'2026-10-02'},{id:'a1',a:'AAPL',d:'2026-01-05',q:1,p:150,u:'2026-10-02'}]});},test(api){
  const R=api.pfDados(),v=id=>(R.find(x=>x.id===id)||{}).v||0,tv=R.reduce((s,x)=>s+(x.v||0),0),lt=k=>api.ETF_IDS.reduce((s,id)=>s+v(id)*(+api.etfInfo(id)[k]||0)/100,0);
  const F={AAPL:(v('AAPL')+lt('AAPL'))/tv,NVDA:(v('NVDA')+lt('NVDA'))/tv,GOOGL:(v('GOOGL')+lt('GOOGL'))/tv,SXR8:v('SXR8')/tv,EUNK:v('EUNK')/tv,IS3N:0,EUNN:0,BTC:v('BTC')/tv,MKT:(tv-v('BTC'))/tv};
  const E=api.expoNoticias();
  ok('exposure: Apple direct + through the ETFs, an ETF and Bitcoin by value, Market = stocks and ETFs',E&&['AAPL','NVDA','SXR8','EUNK','BTC','MKT'].every(k=>near(E[k],F[k],1e-9))&&F.AAPL>v('AAPL')/tv&&!E.IS3N,JSON.stringify(E)+' vs '+JSON.stringify(F));
  const titles=()=>[...document.querySelectorAll('#newsList article .news-t')].map(a=>a.textContent.replace(' (opens in a new window)','').trim());
  const byScore=api.NEWS.slice().sort((a,b)=>b.score-a.score).map(n=>n.titulo);
  ok('default order unchanged: Potential impact (score from the script)',$('#f-ord').value==='score'&&JSON.stringify(titles())===JSON.stringify(byScore),titles().join(' | '));
  const rel=n=>{const ids=n.empresas.filter(c=>c!=='MKT'||n.empresas.length===1);return n.score+3*Math.min(1,ids.reduce((s,c)=>s+(F[c]||0),0));};
  const exp=api.NEWS.slice().sort((a,b)=>rel(b)-rel(a)||b.score-a.score).map(n=>n.titulo);
  ok('relevance = score + 3 × share of the portfolio exposed (sum of the story\'s assets, at most 1)',api.NEWS.every(n=>near(api.relevancia(n,api.expoNoticias()),rel(n),1e-9)));
  const nv=()=>[...document.querySelectorAll('#newsList article')].map(a=>a.className).join(',');const lv0=nv();
  $('#f-ord').value='rel';$('#f-ord').dispatchEvent(new Event('input'));
  ok('"Relevance to me" sorts by relevance',JSON.stringify(titles())===JSON.stringify(exp)&&exp[0]!==byScore[0],titles().join(' | ')+' vs '+exp.join(' | '));
  ok('levels and points still come from the script (unchanged)',[...document.querySelectorAll('#newsList article')].every(a=>{const t=a.querySelector('.news-t').textContent.replace(' (opens in a new window)','').trim(),n=api.NEWS.find(x=>x.titulo===t);return n&&a.classList.contains('l-'+n.nivel)&&a.textContent.includes(api.nf(n.score)+' points');})&&lv0.split(',').sort().join()===nv().split(',').sort().join());
  ok('the formula is explained while this order is chosen',/3 ×/.test(txt('#newsRelNote'))&&!$('#newsRelNote').hidden,txt('#newsRelNote'));
  const via=[...document.querySelectorAll('#newsList article')].find(a=>a.textContent.includes('Story about ASML'));
  ok('"Via top holding" tag on the story matched through a fund\'s holding, with the holding named',via&&/Via top holding/.test(via.textContent)&&/ASML/.test((via.querySelector('.tag[title*="ASML"]')||{}).title||'')&&document.querySelectorAll('#newsList .tag[title*="top holding"]').length===1,via&&via.innerHTML);}},

 newsrelempty:{prep(d){},test(api){
  const o=$('#f-ord option[value="rel"]');
  ok('no holdings: "Relevance to me" is disabled, with a note',o&&o.disabled&&/holdings/i.test(txt('#newsRelNote'))&&!$('#newsRelNote').hidden&&api.expoNoticias()===null,txt('#newsRelNote'));
  ok('no holdings: the default order is used',$('#f-ord').value==='score');}},

 /* ---------- Prices: eventos no gráfico, grandes movimentos e reação aos resultados ---------- */
 evts:{prep(d){
   /* AAPL sintética (USD): 200; +5% a 10 jun e de volta a 11 jun; +6% a 31 jul (resultados); 216 desde 5 ago; −5% a 15 set */
   const fer=['2026-06-19','2026-07-03','2026-09-07'],P=[];for(let t=Date.parse('2026-06-01T00:00:00Z');t<=Date.parse('2026-10-01T00:00:00Z');t+=864e5){const x=new Date(t),iso=x.toISOString().slice(0,10);if(x.getUTCDay()%6===0||fer.includes(iso))continue;
    const v=iso==='2026-06-10'?210:iso<'2026-07-31'?200:iso<'2026-08-05'?212:iso<'2026-09-15'?216:205.2;P.push([iso,v]);}
   const a=d.ativos.find(x=>x.id==='AAPL');a.pontos=P;a.parcial=false;a.fonte='Yahoo Finance';d.historico.AAPL=P.slice();
   const N=(t,data,emp,nivel,temas)=>({chave:t.toLowerCase(),titulo:t,link:'https://example.com/'+encodeURIComponent(t),fonte:'Reuters',data,empresas:emp,nivel,score:nivel==='red'?8:5,temas:temas||[]});
   d.historicoNoticias={inicio:'2026-07-01T00:00:00Z',dias:400,noticias:[N('Apple results beat <img src=x onerror="window.__xss=41">','2026-07-30T21:00:00Z',['AAPL'],'red',['Earnings']),N('Apple supplier warning','2026-07-30T19:00:00Z',['AAPL'],'orange'),
    N('Old Apple story','2026-07-29T15:00:00Z',['AAPL'],'orange'),N('Fed cuts rates by a quarter point','2026-09-16T18:00:00Z',['MKT'],'orange',['Macro and rates']),N('Weekend Apple story','2026-08-01T12:00:00Z',['AAPL'],'orange')]};
   d.resultadosSec={AAPL:{id:'AAPL',estado:'ok',fonte:'SEC EDGAR',obtidoEm:d.geradoEm,erro:'',nota:'',ultimoRelatorio:null,resultados:[{acc:'a1',entrega:'2026-07-30',aceite:'2026-07-30T20:30:28Z',horaNY:'2026-07-30 16:30',quando:'after',sessao:'2026-07-31',nota:''},
    {acc:'a0',entrega:'2026-04-30',aceite:null,horaNY:null,quando:null,sessao:null,nota:'Acceptance time unavailable'}]},NVDA:{id:'NVDA',estado:'error',erro:'blocked by test',resultados:[]}};
   d.calendario=[{d:'2026-09-17',e:'MKT',ev:'Fed decision (16–17 Sep meeting)',imp:'High',st:'C'}];},test(api){
  const T=s=>Date.parse(s.length===10?s+'T00:00:00Z':s),sess=api.HIST.AAPL.map(p=>p[0]);
  ok('session mapping in New York time: after the close (17:00 EDT) → next session; during (15:00 EDT) → that session',api.sessaoDe('AAPL',T('2026-07-30T21:00:00Z'),sess)===T('2026-07-31')&&api.sessaoDe('AAPL',T('2026-07-30T19:00:00Z'),sess)===T('2026-07-30'));
  ok('session mapping: a weekend story goes to Monday; US clock change (20:30 UTC = 15:30 EST on 6 Mar, 16:30 EDT on 9 Mar)',api.sessaoDe('AAPL',T('2026-08-01T12:00:00Z'),sess)===T('2026-08-03')&&api.sessaoDe('AAPL',T('2026-03-06T20:30:00Z'))===T('2026-03-06')&&api.sessaoDe('AAPL',T('2026-03-09T20:30:00Z'))===T('2026-03-10'));
  ok('session mapping: Bitcoin uses the UTC day',api.sessaoDe('BTC',T('2026-07-30T23:30:00Z'))===T('2026-07-30'));
  const M=api.grandesMovimentos().filter(m=>m.id==='AAPL'),m=M.find(x=>x.t===T('2026-07-31'))||{};
  ok('big moves: the sessions of 4% or more (10 Jun, 11 Jun, 31 Jul, 15 Sep), not the smaller ones',M.map(x=>api.isoU(x.t)).sort().join(',')==='2026-06-10,2026-06-11,2026-07-31,2026-09-15',M.map(x=>api.isoU(x.t)+' '+x.mv.toFixed(2)).join(' | '));
  ok('big move explained: the stories of that session and the previous one (not the day before)',near(m.mv,6,1e-9)&&m.news.map(n=>n.titulo).sort().join('|')==='Apple results beat <img src=x onerror="window.__xss=41">|Apple supplier warning',JSON.stringify(m.news&&m.news.map(n=>n.titulo)));
  const row=[...document.querySelectorAll('#tbl-moves tbody tr[data-co="AAPL"]')];
  ok('table: one row per big move, news named, a move before the history says so',row.length===4&&/Apple supplier warning/.test(txt('#tbl-moves'))&&/before the news history/i.test(row.find(r=>/10 Jun/.test(r.textContent)).textContent),txt('#tbl-moves').slice(0,600));
  ok('"News history since" with its start date',/News history since 1 Jul 2026/.test(txt('#movesNote')),txt('#movesNote'));
  const ch=$('#ch-perf'),k=x=>ch.querySelectorAll('.ev-mark[data-k="'+x+'"]').length;
  ok('chart markers: 2 earnings (reaction session; filing date when the time is unknown), 2 rate decisions (calendar and news), 1 material story',k('earn')===2&&k('rate')===2&&k('news')===1,[k('earn'),k('rate'),k('news')].join('/'));
  const mk=ch.querySelector('.ev-mark[data-k="news"]');mk.dispatchEvent(new PointerEvent('pointermove',{bubbles:true,clientX:20,clientY:20}));
  ok('a marker has a tooltip with the headline (as text)',/Apple results beat <img src=x/.test(txt('#tip'))&&!$('#tip').hidden&&/Apple: earnings/.test(ch.querySelector('.ev-mark[data-k="earn"]').getAttribute('data-tip')),txt('#tip'));
  ok('HTML in a headline is not run, in the chart, the tooltip or the table',window.__xss===undefined&&!document.querySelector('#ch-perf img, #tbl-moves img, #tip img, #tbl-earn img'));
  const E=api.reacoes('AAPL');
  ok('earnings reaction in USD: +6.00% in the reaction session, +8.00% after 5 sessions; 1 case (the one without a time is not counted)',E.ok&&E.C.length===1&&near(E.C[0].r1,6,1e-9)&&near(E.C[0].r5,8,1e-9)&&near(E.med,6,1e-9)&&E.sem===1,JSON.stringify(E));
  const er=[...document.querySelectorAll('#tbl-earn tbody tr')],er1=id=>er.find(r=>r.dataset.co===id)||{textContent:''};
  ok('earnings table: Apple with the reaction, the 5-session move, the median and the number of cases',/\+6\.00%/.test(er1('AAPL').textContent)&&/\+8\.00%/.test(er1('AAPL').textContent)&&/31 Jul/.test(er1('AAPL').textContent),er1('AAPL').textContent);
  ok('no SEC data: Unavailable with the reason (NVIDIA failed, Alphabet missing); the rest of the tab works',/Unavailable/.test(er1('NVDA').textContent)&&/blocked by test/.test(er1('NVDA').textContent)&&/Unavailable/.test(er1('GOOGL').textContent)&&$('#tbl-ret tbody').rows.length>0,er1('NVDA').textContent+' / '+er1('GOOGL').textContent);}},

 evtsempty:{prep(d){delete d.historicoNoticias;delete d.resultadosSec;},test(api){
  ok('no news history: big moves still listed, news Unavailable with the reason',/Unavailable/.test(txt('#movesNote'))&&/news history/i.test(txt('#movesNote'))&&!/News history since/.test(txt('#movesNote')),txt('#movesNote'));
  ok('no SEC data: every company Unavailable with the reason, no number invented',[...document.querySelectorAll('#tbl-earn tbody tr')].length===3&&[...document.querySelectorAll('#tbl-earn tbody tr')].every(r=>/Unavailable/.test(r.textContent)&&!/%/.test(r.textContent)),txt('#tbl-earn'));
  ok('no history: no news markers, the chart and the tables still work',$('#ch-perf').querySelectorAll('.ev-mark[data-k="news"]').length===0&&!!$('#ch-perf svg')&&$('#tbl-ret tbody').rows.length>0&&$('#tbl-ind tbody').rows.length>0);}},

 /* ---------- separador Fundamentals ---------- */
 fund:{prep(d){
   /* AAPL: preço constante de 100 (fecho no dia 28 de cada mês desde 2016, e 1 out 2026); EPS 1 por trimestre, exceto
      27 dez 2025 (publicado 1,0 a 30 jan 2026, reexpresso 1,5 a 31 jul 2026) e 28 mar 2026 (3,0, publicado a 1 mai 2026) */
   const H=[];for(let y=2016;y<=2026;y++)for(let m=1;m<=12;m++){if(y===2026&&m>9)break;H.push([`${y}-${String(m).padStart(2,'0')}-28`,100]);}H.push(['2026-10-01',100]);d.historico.AAPL=H;
   const a=d.ativos.find(x=>x.id==='AAPL');a.pontos=[['2026-09-30',100],['2026-10-01',100]];a.parcial=false;a.fonte='Yahoo Finance';
   const add=(iso,n)=>new Date(Date.parse(iso+'T00:00:00Z')+n*864e5).toISOString().slice(0,10),V=(v,f,o)=>Object.assign({v,f,p:f},o||{});
   const fins=['2023-12-30','2024-03-30','2024-06-29','2024-09-28','2024-12-28','2025-03-29','2025-06-28','2025-09-27','2025-12-27','2026-03-28','2026-06-27'];
   const Q=fins.map((fim,i)=>{const f=fim==='2026-03-28'?'2026-05-01':add(fim,35),eps=fim==='2025-12-27'?V(1.5,'2026-07-31',{p:'2026-01-30',v0:1.0}):V(fim==='2026-03-28'?3.0:1.0,f);
    return{fim,inicio:add(fim,-90),ano:'FY'+(+fim.slice(0,4)+(fim.slice(5)>'09-27'?1:0)),q:[1,2,3,4][(i+1)%4],m:{receita:V(100e9+i*1e9,f),lucroBruto:V(45e9,f),lucroOperacional:V(30e9,f),cfo:V(2.5e9,f,{d:true}),capex:V(0.5e9,f,{d:true}),fcf:V(2e9,f,{d:true}),eps,acoes:V(1e9-i*1e7,f)}};});
   const G=Q.map(q=>Object.assign({},q,{m:Object.assign({},q.m,{lucroBruto:null,acoes:q.fim<'2025-06-01'?null:q.m.acoes})}));
   d.fundamentais={AAPL:{id:'AAPL',estado:'ok',fonte:'SEC XBRL (companyfacts)',obtidoEm:d.geradoEm,relatorio:{form:'10-Q',data:'2026-07-31',periodo:'2026-06-27'},tags:{},faltam:{},trimestres:Q,ttm:null,erro:'',nota:''},
    NVDA:{id:'NVDA',estado:'error',erro:'blocked by test',trimestres:[]},
    GOOGL:{id:'GOOGL',estado:'previous run',fonte:'previous run (retrieved 2026-09-30)',obtidoEm:'2026-09-30T08:00:00Z',relatorio:null,tags:{},faltam:{lucroBruto:'no fact for GrossProfit'},trimestres:G,ttm:null,erro:'',nota:''}};},test(api){
  const tabs=[...document.querySelectorAll('.rail a')].map(a=>a.dataset.tab),ip=tabs.indexOf('prices');
  ok('navigation: Fundamentals between Prices and Currency & Macroeconomics',tabs[ip+1]==='fundamentals'&&tabs[ip+2]==='fx',tabs.join(','));
  location.hash='#fundamentals';window.dispatchEvent(new HashChangeEvent('hashchange'));
  ok('hash routing: #fundamentals shows the tab and marks the link',!$('[data-panel="fundamentals"]').hidden&&$('.rail a[data-tab="fundamentals"]').getAttribute('aria-current')==='page'&&$('[data-panel="prices"]').hidden);
  const T=s=>Date.parse(s+'T00:00:00Z'),P=api.fundHistorico('AAPL'),pe=iso=>(P.pe.find(p=>p[0]===T(iso))||[])[1];
  ok('no look-ahead: on 28 Apr 2026 the quarter published on 1 May is not used (P/E = 100 / 4 = 25)',near(pe('2026-04-28'),25,1e-9),pe('2026-04-28'));
  ok('…from 28 May it is (100 / (1 + 1 + 1 + 3) = 16.67)',near(pe('2026-05-28'),100/6,1e-9),pe('2026-05-28'));
  ok('no look-ahead: a restated value counts only from its filing date (1.0 until 31 Jul 2026, then 1.5)',near(pe('2026-06-28'),100/6,1e-9)&&near(pe('2026-08-28'),100/6.5,1e-9),[pe('2026-06-28'),pe('2026-08-28')].join(' / '));
  ok('history: one point per month-end, only once four quarters are published (from Nov 2024)',P.pe.length===24&&api.isoU(P.pe[0][0])==='2024-11-28'&&P.pfcf.length===24&&near(P.pfcf[0][1],100*1e9*(1-3e-2)/8e9,1e-9),P.pe.length+' '+api.isoU(P.pe[0][0])+' '+(P.pfcf[0]||[])[1]);
  ok('percentile: share of the history at or below the current value',api.percentil([1,2,3,4],3)===75&&api.percentil([1,2,3,4],0.5)===0&&near(api.percentil(P.pe.map(p=>p[1]),P.atual.pe),100*P.pe.filter(p=>p[1]<=P.atual.pe+1e-9).length/P.pe.length,1e-9));
  const card=id=>$(`#fundOut [data-co="${id}"]`),ct=id=>(card(id)||{}).textContent||'';
  ok('current valuation: P/E TTM = 100 / 6.5, P/FCF TTM = 100 × 0.9 bn / 8 bn, with the quarters used and the percentile',near(P.atual.pe,100/6.5,1e-9)&&near(P.atual.pfcf,11.25,1e-9)&&ct('AAPL').includes(api.nf(P.atual.pe,1))&&ct('AAPL').includes(api.nf(P.atual.pfcf,1))&&/27 Jun 2026/.test(ct('AAPL'))&&/higher than/i.test(ct('AAPL')),ct('AAPL').slice(0,500));
  ok('quarter table: revenue, year-on-year growth, margins, FCF, EPS, shares, with the quarter and the filing date',card('AAPL').querySelectorAll('tbody tr').length===8&&/Gross margin/.test(ct('AAPL'))&&/40\.9%/.test(ct('AAPL'))&&/1 Aug 2026/.test(ct('AAPL'))&&/\+3\.8%/.test(ct('AAPL'))&&/−4\.3%|-4\.3%/.test(ct('AAPL')),ct('AAPL').slice(0,900));
  ok('partial data: gross margin Unavailable with the reason; labelled previous run',/Unavailable/.test(ct('GOOGL'))&&/GrossProfit/.test(ct('GOOGL'))&&/previous run/.test(ct('GOOGL')),ct('GOOGL').slice(0,600));
  ok('a company without data: Unavailable with the reason, the others shown',/Unavailable/.test(ct('NVDA'))&&/blocked by test/.test(ct('NVDA'))&&!!card('AAPL').querySelector('svg'));
  ok('the note: valuation gives context, not a short-term forecast',/context/i.test(txt('[data-panel="fundamentals"]'))&&/short term/i.test(txt('[data-panel="fundamentals"]')));
  ok('no horizontal overflow in the tab',$('[data-panel="fundamentals"]').scrollWidth<=$('[data-panel="fundamentals"]').clientWidth+1,$('[data-panel="fundamentals"]').scrollWidth+' > '+$('[data-panel="fundamentals"]').clientWidth);
  $('.chip[data-co="AAPL"]').click();
  ok('asset filter: Apple only',!!card('AAPL')&&!card('NVDA')&&!card('GOOGL'));
  $('.chip[data-co="EUNK"]').click();
  ok('asset filter on an ETF: a note instead of the companies',!card('AAPL')&&/Apple, NVIDIA and Alphabet/.test(txt('#fundOut')),txt('#fundOut'));
  $('.chip[data-co="all"]').click();}},

 fundempty:{prep(d){delete d.fundamentais;},test(api){
  location.hash='#fundamentals';window.dispatchEvent(new HashChangeEvent('hashchange'));
  ok('no fundamentals in the data: every company Unavailable with the reason, no number',[...document.querySelectorAll('#fundOut [data-co]')].length===3&&[...document.querySelectorAll('#fundOut [data-co]')].every(c=>/Unavailable/.test(c.textContent)&&/EmailSEC/.test(c.textContent)),txt('#fundOut'));}},

 /* ---------- Currency & Macroeconomics: área do euro (BCE) ---------- */
 macro:{prep(d){d.macro={BCE_DFR:{id:'BCE_DFR',estado:'ok',fonte:'ECB Data Portal',obtidoEm:d.geradoEm,freq:'B',pontos:[['2014-09-10',-0.2],['2019-09-18',-0.5],['2022-07-27',0],['2022-09-14',0.75],['2023-09-20',4],['2024-06-12',3.75],['2026-09-16',2.5]]},
   HICP_EA:{id:'HICP_EA',estado:'previous run',fonte:'previous run (retrieved 2026-10-01)',obtidoEm:'2026-10-01T08:00:00Z',freq:'M',pontos:[['2025-10-01',2.1],['2025-11-01',2.1],['2025-12-01',1.9]]}};},test(api){
  ok('ECB deposit rate: latest value and the date it took effect, 10-year chart',/2\.50%/.test(txt('#k-dfr'))&&/16 Sep 2026/.test(txt('#k-dfr'))&&!!$('#ch-dfr svg'),txt('#k-dfr'));
  ok('euro area inflation: value and month, flagged when not recent, labelled previous run',/1\.9%/.test(txt('#k-hicp'))&&/Dec 2025/.test(txt('#k-hicp'))&&!!$('#k-hicp .old')&&/previous run/.test(txt('#k-hicp'))&&!!$('#ch-hicp svg'),txt('#k-hicp'));
  ok('honest notes: thermometers not clocks, inverted curve with false alarms, valuation over 10 years',/thermometers, not clocks/i.test(txt('[data-panel="fx"]'))&&/false alarms/.test(txt('#macroNao'))&&/10 years/.test(txt('#macroNao')),txt('#macroNao'));
  ok('no new alert in the Overview',!api.alerts().some(a=>/ECB|HICP|inflation|deposit/i.test(a.txt)),JSON.stringify(api.alerts().map(a=>a.txt)));}},

 macroempty:{prep(d){delete d.macro;},test(){
  ok('no ECB data: both Unavailable with the reason, the rest of the tab works',/Unavailable/.test(txt('#ch-dfr'))&&/Unavailable/.test(txt('#ch-hicp'))&&/no ECB data/.test(txt('#ch-dfr'))&&txt('#k-dfr')===''&&!!$('#ch-fx svg'),txt('#ch-dfr'));}},

 /* ---------- exportação para o anexo J ---------- */
 tax:{prep(d){capDl();seed(TAXSEED);},test(api){
  const X=api.anexoJ(2026),A=X.acoes,R=csvParse(X.csvAcoes),C=csvParse(X.csvCripto),row=(l,v)=>A.find(r=>r.lote===l&&r.venda===v);
  ok('FIFO: 120 AAPL sold from lots of 100 and 50 → two rows, 100 + 20',!!row('a1','s1')&&!!row('a2','s1')&&near(row('a1','s1').q,100)&&near(row('a2','s1').q,20),JSON.stringify(A.filter(r=>r.venda==='s1').map(r=>[r.lote,r.q])));
  ok('FIFO: acquisition values 100 × €100 and 20 × €120',row('a1','s1').aq===10000&&row('a2','s1').aq===2400);
  ok('user-entered sale price is used, not the market close',row('a1','s1').vd===15000&&row('a2','s1').vd===3000&&row('a1','s1').ganho===5000);
  ok('one purchase, two sales: two rows from the same lot, sale equal to the holding',A.filter(r=>r.lote==='g1').length===2&&near(A.filter(r=>r.lote==='g1').reduce((s,r)=>s+r.q,0),10)&&api.pfDados().find(x=>x.id==='GOOGL').q<1e-9);
  ok('unsold lots are not exported (NVDA)',!A.some(r=>r.id==='NVDA'));
  ok('earlier years are not mixed in (the 2025 sale is in the 2025 export only)',!A.some(r=>r.venda==='s0')&&api.anexoJ(2025).acoes.length===1&&api.anexoJ(2025).acoes[0].lote==='a0');
  ok('rows in the 2026 file: 5 stock/ETF, 2 crypto',A.length===5&&X.cripto.length===2,A.length+' / '+X.cripto.length);
  const a1=row('a1','s1');
  ok('acquisition and sale FX resolved separately, each for its own date',a1.fa&&a1.fv&&a1.fa.d<='2026-01-05'&&a1.fv.d<='2026-06-01'&&a1.fa.d!==a1.fv.d&&a1.fa.fx===api.fxAt(Date.parse('2026-01-05T00:00:00Z'),api.HIST.FX),JSON.stringify([a1.fa,a1.fv]));
  ok('dollar columns are EUR × that day\'s EUR/USD',csvCol(R,'Acquisition total (USD, derived)')[A.indexOf(a1)]===(Math.round(10000*a1.fa.fx*100)/100).toFixed(2).replace('.',','),csvCol(R,'Acquisition total (USD, derived)')[A.indexOf(a1)]);
  ok('mapping: AAPL G01 / 840, SXR8 G20 / 372, listed = Sim',a1.codigo==='G01'&&a1.pais==='840'&&a1.admitido==='Sim'&&A.find(r=>r.id==='SXR8').codigo==='G20'&&A.find(r=>r.id==='SXR8').pais==='372');
  ok('CSV: UTF-8 BOM, ";" separator, CRLF',X.csvAcoes.charCodeAt(0)===0xFEFF&&X.csvAcoes.split('\r\n')[0].split(';')[0].replace(/^﻿/,'')==='Tax year'&&!/[^\r]\n/.test(X.csvAcoes));
  ok('CSV: decimal comma, no thousands separator, no currency symbol',csvCol(R,'Aquisição Valor (EUR)')[A.indexOf(a1)]==='10000,00'&&csvCol(R,'Realização Valor (EUR)')[A.indexOf(a1)]==='15000,00'&&!/€|\$/.test(X.csvAcoes.split('\r\n').slice(1).join('')),csvCol(R,'Aquisição Valor (EUR)').join('|'));
  ok('CSV: dates split as in the form (2026 / 6 / 1)',csvCol(R,'Realização Ano')[A.indexOf(a1)]==='2026'&&csvCol(R,'Realização Mês')[A.indexOf(a1)]==='6'&&csvCol(R,'Realização Dia')[A.indexOf(a1)]==='1'&&csvCol(R,'Sale date')[A.indexOf(a1)]==='2026-06-01');
  ok('CSV: every row has the header\'s number of columns',R.every(r=>r.length===R[0].length)&&C.every(r=>r.length===C[0].length));
  ok('CSV: a reference with ; " and a leading = is escaped',X.csvAcoes.includes('"\'=x;""q"""')&&csvCol(R,'Sale reference').includes('\'=x;"q"'));
  ok('CSV: deterministic (same data, same file)',api.anexoJ(2026).csvAcoes===X.csvAcoes&&api.anexoJ(2026).csvCripto===X.csvCripto);
  ok('crypto: lot held 396 days is not Anexo J; lot held 59 days goes to Quadro 9.4A',X.cripto.find(r=>r.lote==='L1').quadro.startsWith('Not Anexo J')&&X.cripto.find(r=>r.lote==='L2').quadro==='9.4A'&&X.cripto.find(r=>r.lote==='L2').dias===59);
  ok('crypto: FIFO values (0.5 BTC for €15,000; 0.1 BTC for €6,666.67)',X.cripto.find(r=>r.lote==='L1').aq===15000&&X.cripto.find(r=>r.lote==='L2').aq===6666.67&&X.cripto.find(r=>r.lote==='L2').vd===7000);
  ok('crypto CSV has no stock code column and asks for the platform country',C[0].indexOf('Código')<0&&/País da fonte \(country of the platform\)/.test(X.csvCripto));
  ok('year list offers the years with sales, latest first',[...document.querySelectorAll('#taxYear option')].map(o=>o.value).join(',')==='2026,2025');
  G.ls=lsSnap();$('#taxYear').value='2026';$('#taxExport').click();
  ok('button: first file downloaded, message shown',window.__dl[0]==='AnexoJ_Stocks_ETFs_2026.csv'&&/Downloaded/.test(txt('#taxMsg')),JSON.stringify(window.__dl)+' '+txt('#taxMsg'));
  ok('CSV values: no number formatted by the browser locale',api.csvNum(1234567.891,2)==='1234567,89'&&api.csvNum(-12.5,2)==='-12,50'&&api.csvNum(-0.001,2)==='0,00'&&api.csvNum(null,2)==='');
  ok('CSV text: quotes, separators and formula characters',api.csvTxt('a;b')==='"a;b"'&&api.csvTxt('say "hi"')==='"say ""hi"""'&&api.csvTxt('=1+1')==="'=1+1"&&api.csvTxt('-x')==="'-x"&&api.csvTxt('plain')==='plain');},
  after(){ok('button: both files downloaded (stocks/ETF and crypto)',JSON.stringify(window.__dl)==='["AnexoJ_Stocks_ETFs_2026.csv","AnexoJ_Crypto_2026.csv"]',JSON.stringify(window.__dl));
   ok('export changes nothing in the browser storage',lsSnap()===G.ls);}},

 taxsplit:{prep(d){
   const iso='2025-01-02';G.old=histPrice(d,'NVDA',iso);
   const div=L=>L.map(pair).map(p=>[p[0],+p[1]/10]);d.historico.NVDA=div(d.historico.NVDA);const a=d.ativos.find(x=>x.id==='NVDA');a.pontos=div(a.pontos);
   d.splits=d.splits||{};d.splits.NVDA=(d.splits.NVDA||[]).map(pair).concat([['2026-10-01',10]]);G.newRef=histPrice(d,'NVDA','2026-10-02');
   seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[
    {id:'n1',a:'NVDA',d:iso,q:2,p:100,r:[iso,G.old],u:'2026-09-30'},{id:'n2',a:'NVDA',d:'2025-02-03',q:3,p:90,u:'2026-09-30'},
    {id:'n3',a:'NVDA',d:'2026-10-02',q:1,p:20,r:['2026-10-02',G.newRef],u:'2026-10-02'},{id:'a9',a:'AAPL',d:'2013-06-03',q:1,p:500,u:'2013-12-31'}],
    sales:[{id:'s0',a:'NVDA',d:'2025-03-03',q:1,p:95,u:'2026-09-30'},{id:'s1',a:'NVDA',d:'2026-10-02',q:15,p:15,r:['2026-10-02',G.newRef],u:'2026-10-02'},{id:'s2',a:'AAPL',d:'2026-03-02',q:10,p:200,u:'2026-10-02'}]});},test(api){
  const A5=api.anexoJ(2025).acoes,A6=api.anexoJ(2026).acoes,r=(L,l,v)=>L.find(x=>x.lote===l&&x.venda===v)||{};
  ok('bought and sold before a later 10:1 split: 1 old share = 10 of today, values unchanged',near(r(A5,'n1','s0').q,10)&&r(A5,'n1','s0').aq===100&&r(A5,'n1','s0').vd===95&&r(A5,'n1','s0').kA===10&&r(A5,'n1','s0').kV===10,JSON.stringify(r(A5,'n1','s0')));
  ok('bought before the split, sold after: today\'s shares, cost per share ÷ 10',near(r(A6,'n1','s1').q,10)&&r(A6,'n1','s1').aq===100&&r(A6,'n1','s1').vd===150&&r(A6,'n1','s1').kV===1,JSON.stringify(r(A6,'n1','s1')));
  ok('partial sale after the split spans two pre-split lots (10 + 5)',near(r(A6,'n2','s1').q,5)&&r(A6,'n2','s1').aq===45&&r(A6,'n2','s1').vd===75);
  ok('never adjusted twice: the lot\'s total cost across its rows is the €200 paid',r(A5,'n1','s0').aq+r(A6,'n1','s1').aq===200);
  ok('purchase after the split is not adjusted (and not sold here)',api.efetiva((JSON.parse(localStorage.getItem('bb.buys'))||[]).find(x=>x.id==='n3')).k===1&&!A6.some(x=>x.lote==='n3'));
  ok('two splits (7:1 and 4:1) since a 2013 purchase: ×28, cost unchanged',r(A6,'a9','s2').kA===28&&near(r(A6,'a9','s2').q,10)&&r(A6,'a9','s2').aq===178.57&&r(A6,'a9','s2').vd===2000,JSON.stringify(r(A6,'a9','s2')));}},

 taxcrypto:{prep(d){seed({deleted:{},buys:[],savedAt:'2026-10-03T10:00:00.000Z',lots:[{id:'C1',d:'2025-04-01',q:1,c:50000}],
   sales:[{id:'cs1',a:'BTC',d:'2026-03-31',q:0.25,p:80000},{id:'cs2',a:'BTC',d:'2026-04-01',q:0.25,p:80000},{id:'cs3',a:'BTC',d:'2026-09-01',q:0.25,p:90000}]});},test(api){
  const X=api.anexoJ(2026),r=v=>X.cripto.find(x=>x.venda===v)||{};
  ok('crypto: sold after 364 days → Quadro 9.4A, flagged for review (close to 365)',r('cs1').quadro==='9.4A'&&r('cs1').dias===364&&r('cs1').status==='REVIEW'&&r('cs1').flags.includes('NEAR_365_DAYS'));
  ok('crypto: sold on day 365 → not Anexo J (same rule as the 365-day counter), flagged',r('cs2').quadro.startsWith('Not Anexo J')&&r('cs2').dias===365&&r('cs2').status==='REVIEW');
  ok('crypto: held 518 days → not Anexo J, no review needed',r('cs3').quadro.startsWith('Not Anexo J')&&r('cs3').status==='OK');
  ok('crypto: values from the purchase cost (€50,000 for 1 BTC)',r('cs1').aq===12500&&r('cs1').vd===20000&&r('cs1').ganho===7500);
  ok('crypto CSV: 8-decimal quantity with a decimal comma',csvCol(csvParse(X.csvCripto),'Quantity (BTC)')[0]==='0,25000000');
  ok('a year without sales gives a header-only file',api.anexoJ(2025).csvCripto.split('\r\n').filter(Boolean).length===1&&api.anexoJ(2025).cripto.length===0);
  ok('the 365-day counter agrees (tax-free from 1 Apr 2026)',/1 Apr 2026/.test(txt('#tbl-lots')),txt('#tbl-lots'));}},

 taxold:{prep(d){d.backup=V1;},test(api){
  const X=api.anexoJ(2026);
  ok('old version-1 backup (no sales): export works and is empty',X.acoes.length===0&&X.cripto.length===0&&X.mal.length===0&&X.csvAcoes.split('\r\n').length===2);
  ok('year list falls back to the current year',$('#taxYear').options.length===1);
  $('#buyAsset').value='SXR8';$('#buyType').value='sell';$('#buyDate').value='2026-10-02';$('#buyQty').value='5';$('#buyPrice').value='700';
  $('#buyForm').dispatchEvent(new Event('submit',{cancelable:true}));
  ok('the form still rejects a sale larger than the holdings',/did not hold enough/.test(txt('#buyMsg'))&&(JSON.parse(localStorage.getItem('bb.sales'))||[]).length===0,txt('#buyMsg'));}},

 taxbad:{prep(d){seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'o1',a:'AAPL',d:'2026-01-05',q:2,p:200}],
   sales:[{id:'os1',a:'AAPL',d:'2026-02-02',q:1,p:210},{id:'bad1',a:'AAPL',d:'2026-13-45',q:1,p:1},{id:'bad2',a:'XYZ',d:'2026-02-02',q:1,p:1},{id:'bad3',a:'AAPL',d:'2026-02-03',q:'abc',p:1},{id:'ov',a:'GOOGL',d:'2026-03-03',q:2,p:100},{id:'np',a:'AAPL',d:'2026-02-04',q:0.5,p:0}]});},test(api){
  const X=api.anexoJ(2026),r=v=>X.acoes.find(x=>x.venda===v)||{};
  ok('entries without r/u (older versions) export normally',r('os1').aq===200&&r('os1').vd===210&&r('os1').status==='OK',JSON.stringify(r('os1')));
  ok('malformed sales (bad date, unknown asset, bad quantity) are left out and counted',X.mal.length===3&&!X.acoes.some(x=>/^bad/.test(x.venda)),JSON.stringify(X.mal));
  ok('sale larger than the holdings: REVIEW row, no invented purchase',r('ov').status==='REVIEW'&&r('ov').flags.includes('OVERSOLD_NO_PURCHASE')&&r('ov').aq===null&&csvCol(csvParse(X.csvAcoes),'Aquisição Valor (EUR)')[X.acoes.indexOf(r('ov'))]==='');
  ok('sale with no price: REVIEW',r('np').status==='REVIEW'&&r('np').flags.includes('NO_SALE_PRICE'));}},

 taxfx:{prep(d){d.fx.yahoo=[];d.fx.bce=[];d.historico.FX=(d.historico.FX||[]).map(pair).filter(p=>p[0]>='2026-03-01');capDl();
   seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a',a:'AAPL',d:'2026-01-05',q:2,p:200,u:'2026-10-02'}],sales:[{id:'s',a:'AAPL',d:'2026-06-01',q:2,p:250,u:'2026-10-02'}]});},test(api){
  const r=api.anexoJ(2026).acoes[0]||{},R=csvParse(api.anexoJ(2026).csvAcoes);
  ok('no EUR/USD for the purchase date: acquisition rate left blank, never today\'s rate',r.fa===null&&r.fv&&r.fv.d<='2026-06-01'&&csvCol(R,'Acquisition FX EUR/USD')[0]===''&&csvCol(R,'Acquisition total (USD, derived)')[0]==='',JSON.stringify([r.fa,r.fv]));
  ok('missing rate is flagged, euro values unaffected',r.flags.includes('NO_FX_REFERENCE')&&r.aq===400&&r.vd===500&&r.ganho===100);}},

 taxfx0:{prep(d){d.fx.yahoo=[];d.fx.bce=[];d.historico.FX=[];
   seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a',a:'AAPL',d:'2026-01-05',q:2,p:200,u:'2026-10-02'}],sales:[{id:'s',a:'AAPL',d:'2026-06-01',q:1,p:250,u:'2026-10-02'}]});},test(api){
  const r=api.anexoJ(2026).acoes[0]||{};
  ok('no EUR/USD at all: export still works, dollar columns blank',r.fa===null&&r.fv===null&&r.aq===200&&r.vd===250&&r.flags.includes('NO_FX_REFERENCE'));}},

 /* Não-regressão do anexo J (invariante 10): os CSV de 2025 e 2026 de um backup fixo têm de ser idênticos, byte a byte,
    aos ficheiros de referência em Tests\fixtures\taxBaseline (gerados uma vez com -WriteFixtures). O histórico, o EUR/USD
    e os splits usados vêm do próprio cenário, para o resultado não depender dos dados do dia. */
 taxBaseline:{prep(d){capDl();
   const S=(o)=>Object.keys(o).sort().map(k=>[k,o[k]]);
   d.historico.AAPL=S({'2024-11-04':220,'2025-02-03':225,'2025-06-02':230,'2026-05-04':250,'2026-10-01':255});
   d.historico.NVDA=S({'2024-03-01':82,'2026-02-02':120,'2026-10-01':180});
   d.historico.SXR8=S({'2025-01-06':600,'2026-03-02':650,'2026-10-01':700});
   d.historico.EUNK=S({'2025-05-05':80,'2026-04-01':85,'2026-10-01':90});
   d.historico.FX=S({'2024-03-01':1.0838,'2024-11-04':1.0883,'2025-01-06':1.0389,'2025-02-03':1.0302,'2025-05-05':1.1316,'2025-06-02':1.1442,'2025-09-01':1.1708,
    '2026-02-02':1.1795,'2026-03-02':1.1712,'2026-04-01':1.1633,'2026-05-04':1.1548,'2026-10-01':1.1734});
   d.splits={AAPL:[],NVDA:[['2024-06-10',10]],GOOGL:[],SXR8:[],EUNK:[],IS3N:[],EUNN:[]};
   d.backup={app:'Bluechip Board',version:4,saved:'2026-10-01T10:00:00.000Z',deleted:{},
    buys:[{id:'a1',a:'AAPL',d:'2024-11-04',q:10,p:200,r:['2024-11-04',220],u:'2026-10-01'},{id:'a2',a:'AAPL',d:'2025-02-03',q:5,p:210,r:['2025-02-03',225],u:'2026-10-01'},
     {id:'n1',a:'NVDA',d:'2024-03-01',q:2,p:750,r:['2024-03-01',820],u:'2024-03-01'},{id:'x1',a:'SXR8',d:'2025-01-06',q:3,p:600,r:['2025-01-06',600],u:'2026-10-01'},
     {id:'e1',a:'EUNK',d:'2025-05-05',q:4,p:80,r:['2025-05-05',80],u:'2026-10-01'}],
    sales:[{id:'v1',a:'AAPL',d:'2025-06-02',q:12,p:230,r:['2025-06-02',230],u:'2026-10-01'},{id:'v2',a:'NVDA',d:'2026-02-02',q:15,p:120,r:['2026-02-02',120],u:'2026-10-01'},
     {id:'v3',a:'SXR8',d:'2026-03-02',q:3,p:650,r:['2026-03-02',650],u:'2026-10-01'},{id:'v4',a:'EUNK',d:'2026-04-01',q:1.5,p:85,r:['2026-04-01',85],u:'2026-10-01'},
     {id:'v5',a:'AAPL',d:'2026-05-04',q:1,p:250,r:['2026-05-04',250],u:'2026-10-01'},{id:'c1',a:'BTC',d:'2025-06-02',q:0.1,p:95000},{id:'c2',a:'BTC',d:'2026-03-02',q:0.4,p:70000}],
    lots:[{id:'L1',d:'2024-03-01',q:0.4,c:20000},{id:'L2',d:'2025-09-01',q:0.2,c:18000}]};},test(api){
  const F=window.__FIX||{},b64=s=>{const u=new TextEncoder().encode(s);let t='';for(let i=0;i<u.length;i++)t+=String.fromCharCode(u[i]);return btoa(t);};
  const out={},X={2025:api.anexoJ(2025),2026:api.anexoJ(2026)};
  ok('test backup loaded: 5 purchases, 7 sales, 2 Bitcoin lots',(get('buys')||[]).length===5&&(get('sales')||[]).length===7&&(get('lots')||[]).length===2);
  ok('fixture is exercised: NVDA split ×10, AAPL FIFO over two lots, BTC on both sides of 365 days',X[2026].acoes.some(r=>r.id==='NVDA'&&r.kA===10)&&X[2025].acoes.filter(r=>r.venda==='v1').length===2&&X[2026].cripto.some(r=>r.quadro==='9.4A')&&X[2026].cripto.some(r=>/^Not Anexo J/.test(r.quadro)),JSON.stringify(X[2026].cripto.map(r=>[r.lote,r.quadro])));
  [2025,2026].forEach(y=>[['AnexoJ_Stocks_ETFs_'+y+'.csv',X[y].csvAcoes],['AnexoJ_Crypto_'+y+'.csv',X[y].csvCripto]].forEach(([n,csv])=>{
   const got=b64(csv);out[n]=got;
   if(!F[n]){ok(n+': reference file exists in Tests\\fixtures\\taxBaseline',false,'missing: run Test-Site.ps1 -Only taxBaseline -WriteFixtures once');return;}
   let i=0;const a=atob(got),e=atob(F[n]);while(i<a.length&&i<e.length&&a[i]===e[i])i++;
   ok(n+': identical byte by byte to the reference file ('+e.length+' bytes)',a===e,a===e?'':'first difference at byte '+i+' (got '+a.length+' bytes): …'+a.slice(Math.max(0,i-40),i+40)+'…');}));
  window.__OUT=out;G.n=[atob(F['AnexoJ_Stocks_ETFs_2026.csv']||''),atob(F['AnexoJ_Crypto_2026.csv']||'')].map(s=>s.length);
  $('#taxYear').value='2026';$('#taxExport').click();},
  after(){ok('export button: the two downloaded files have the reference sizes',JSON.stringify(window.__dl)==='["AnexoJ_Stocks_ETFs_2026.csv","AnexoJ_Crypto_2026.csv"]'&&window.__blobs.length===2&&window.__blobs[0].size===G.n[0]&&window.__blobs[1].size===G.n[1],JSON.stringify(window.__dl)+' '+window.__blobs.map(b=>b.size)+' vs '+G.n);}},

 /* ---------- rentabilidade (XIRR) e benchmark pessoal em SXR8 ---------- */
 ret:{prep(d){retFix(d);seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',
   buys:[{id:'x1',a:'SXR8',d:'2024-11-04',q:2,p:100,r:['2024-11-04',100],u:'2026-10-01'},{id:'x2',a:'SXR8',d:'2025-06-02',q:1,p:200,r:['2025-06-02',200],u:'2026-10-01'}]});},test(api){
  const R=api.rentab(),t=iso=>Date.parse(iso+'T00:00:00Z'),fl=[[t('2024-11-04'),-200],[t('2025-06-02'),-200],[t(G.fim),750]],exp=xirrInd(fl,365);
  ok('only purchases: XIRR matches an independent calculation',R.estado==='ok'&&R.anual&&near(R.r,exp,1e-7),R.r+' vs '+exp);
  ok('flows: purchases negative (q × p), closing value positive, dated on the build day in Lisbon',JSON.stringify(R.fl.map(x=>[api.isoU(x[0]),Math.round(x[1]*100)/100]))===JSON.stringify(fl.map(x=>[api.isoU(x[0]),x[1]])),JSON.stringify(R.fl));
  ok('benchmark: same money, same dates in SXR8 = the same portfolio (3 units, €750, same XIRR)',R.bench.estado==='ok'&&near(R.bench.unidades,3)&&near(R.bench.valor,750)&&near(R.bench.r,R.r,1e-9)&&near(R.dif,0,1e-9));
  ok('shown: return per year, benchmark value and the difference',/a year/.test(txt('#retKpis'))&&txt('#retKpis').includes(api.pct(exp*100))&&txt('#retKpis').includes('€750')&&/0\.0 pp/.test(txt('#retKpis')),txt('#retKpis'));
  ok('dates and sources are named',/1 Oct 2026/.test(txt('#retNote'))&&/4 Nov 2024/.test(txt('#retNote'))&&/SXR8/.test(txt('#retNote')),txt('#retNote'));
  ok('solver: no sign change → no result (Unavailable)',api.xirr([[t('2025-01-01'),-100],[t('2026-01-01'),-50]],365)===null);
  const hard=[[t('2026-01-01'),-100],[t('2026-04-01'),5000]];
  ok('solver: an extreme case still matches the independent calculation',near(api.xirr(hard,365),xirrInd(hard,365),1e-6),api.xirr(hard,365)+' vs '+xirrInd(hard,365));}},
 retsales:{prep(d){retFix(d);d.ativos.find(x=>x.id==='SXR8').parcial=true;seed({deleted:{},savedAt:'2026-10-03T10:00:00.000Z',
   buys:[{id:'e1',a:'EUNK',d:'2025-01-06',q:20,p:50,r:['2025-01-06',50],u:'2026-10-01'}],lots:[{id:'L1',d:'2025-06-02',q:0.01,c:500}],
   sales:[{id:'s1',a:'EUNK',d:'2026-01-05',q:15,p:60,r:['2026-01-05',60],u:'2026-10-01'}]});},test(api){
  const R=api.rentab(),t=iso=>Date.parse(iso+'T00:00:00Z'),fl=[[t('2025-01-06'),-1000],[t('2025-06-02'),-500],[t('2026-01-05'),900],[t(G.fim),1150]];
  ok('purchases (including Bitcoin at its total cost) and sales: XIRR matches an independent calculation',R.estado==='ok'&&near(R.r,xirrInd(fl,365),1e-7)&&near(R.valor,1150),R.r+' vs '+xirrInd(fl,365)+' / '+R.valor);
  ok('benchmark by hand: 10 + 2.5 − 4 = 8.5 SXR8 units × €250 = €2,125',R.bench.estado==='ok'&&near(R.bench.unidades,8.5)&&near(R.bench.valor,2125)&&!R.bench.curto,JSON.stringify(R.bench));
  const bf=[[t('2025-01-06'),-1000],[t('2025-06-02'),-500],[t('2026-01-05'),900],[t(G.fim),2125]];
  ok('benchmark XIRR and the difference',near(R.bench.r,xirrInd(bf,365),1e-7)&&near(R.dif,R.r-R.bench.r,1e-12)&&txt('#retKpis').includes('€2,125'),R.bench.r+' vs '+xirrInd(bf,365));
  ok('an intraday SXR8 price is named as such, never as a close',/latest price \(intraday, 1 Oct 2026\)/.test(txt('#retNote'))&&!/close of 1 Oct/.test(txt('#retNote')),txt('#retNote'));
  ok('the existing tables are unchanged by the new section',$('#tbl-pf')&&$('#tbl-div')&&$('#retKpis').compareDocumentPosition($('#tbl-div'))&Node.DOCUMENT_POSITION_FOLLOWING&&$('#tbl-pf').compareDocumentPosition($('#retKpis'))&Node.DOCUMENT_POSITION_FOLLOWING);}},
 retshort:{prep(d){retFix(d);seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',
   buys:[{id:'e1',a:'EUNK',d:'2026-01-05',q:20,p:50,r:['2026-01-05',60],u:'2026-10-01'}],
   sales:[{id:'s1',a:'EUNK',d:'2026-04-01',q:10,p:75,r:['2026-04-01',75],u:'2026-10-01'}]});},test(api){
  const R=api.rentab(),t=iso=>Date.parse(iso+'T00:00:00Z'),T=(t(G.fim)-t('2026-01-05'))/864e5,fl=[[t('2026-01-05'),-1000],[t('2026-04-01'),750],[t(G.fim),700]];
  ok('first flow under a year ago: return for the period, not annualised, with a note',R.estado==='ok'&&!R.anual&&near(R.r,xirrInd(fl,T),1e-7)&&/not annualised/i.test(txt('#retKpis')+txt('#retNote')),R.r+' vs '+xirrInd(fl,T)+' · '+txt('#retNote'));
  ok('benchmark without enough SXR8 units for a sale: sells all of it and says so',R.bench.estado==='ok'&&R.bench.curto&&near(R.bench.unidades,0)&&near(R.bench.valor,0)&&/not enough/i.test(txt('#retNote')),JSON.stringify(R.bench)+' · '+txt('#retNote'));
  const bf=[[t('2026-01-05'),-1000],[t('2026-04-01'),600],[t(G.fim),0]];
  ok('benchmark flows: only what the SXR8 units fetched (€600)',near(R.bench.r,xirrInd(bf,T),1e-7),R.bench.r+' vs '+xirrInd(bf,T));}},
 retmissing:{prep(d){retFix(d);const a=d.ativos.find(x=>x.id==='EUNN');a.pontos=[];
   seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',
   buys:[{id:'n1',a:'EUNN',d:'2025-01-06',q:1,p:70,u:'2026-10-01'},{id:'a1',a:'AAPL',d:'2009-06-01',q:1,p:100,u:'2026-10-01'}]});},test(api){
  const R=api.rentab();
  ok('an asset held without a current price: Unavailable, naming it, no number',R.estado==='noprice'&&R.r==null&&/Unavailable/.test(txt('#retKpis'))&&/Japan/.test(txt('#retKpis')+txt('#retNote'))&&!/%/.test(txt('#retKpis').replace(/[^%]*Unavailable[^%]*/,'')),txt('#retKpis')+' · '+txt('#retNote'));
  ok('a purchase before the SXR8 history: benchmark Unavailable, with the date',R.bench.estado==='early'&&R.bench.r==null&&/1 Jun 2009/.test(txt('#retNote')),JSON.stringify(R.bench)+' · '+txt('#retNote'));}},
 retempty:{prep(d){seed({deleted:{},buys:[],lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z'});},test(api){
  const R=api.rentab();
  ok('empty portfolio: no figures, a short explanation',R.estado==='empty'&&R.r==null&&R.bench.r==null&&!/%|€/.test(txt('#retKpis'))&&/purchases/i.test(txt('#retKpis')),txt('#retKpis'));}},

 /* ---------- alocação-alvo e próxima contribuição (backup v5: targets) ---------- */
 tgtsplit:{prep(d){tgtFix(d);seed(Object.assign(TGTSEED(),{targets:{weights:{SXR8:60,EUNK:30,BTC:10},band:5,monthly:300,at:'2026-10-02T10:00:00.000Z'}}));},test(api){
  const A=api.alocacao(),cur={SXR8:2000,EUNK:700,BTC:800},w={SXR8:.6,EUNK:.3,BTC:.1},exp=reparteInd(cur,w,300),L=id=>A.linhas.find(x=>x.id===id)||{};
  ok('current values and weights from Your holdings (total €3,500)',near(A.T,3500)&&near(L('SXR8').peso,2000/35)&&near(L('EUNK').peso,20)&&near(L('BTC').peso,800/35),JSON.stringify(A.linhas.map(x=>[x.id,x.v,x.peso])));
  ok('deficits larger than M: M split in proportion to the deficits (€116.67 / €183.33 / €0)',['SXR8','EUNK','BTC'].every(id=>near(L(id).aloc,exp[id],1e-9))&&near(exp.SXR8,300*280/720,1e-9)&&near(L('BTC').aloc,0),JSON.stringify(A.linhas.map(x=>[x.id,x.aloc])));
  ok('the split adds up to M',near(A.linhas.reduce((s,x)=>s+(x.aloc||0),0),300,1e-9));
  ok('deviation and band: EUNK (−10 pp) and Bitcoin (+12.9 pp) outside ±5 pp, SXR8 inside',L('EUNK').fora&&L('BTC').fora&&!L('SXR8').fora&&near(L('EUNK').desvio,-10,1e-9)&&$('#tbl-tgt').querySelectorAll('.tag.t-warn').length===2,txt('#tbl-tgt'));
  ok('months of contributions to get back inside the band, without selling (independent simulation)',A.meses===mesesInd(cur,w,300,5),A.meses+' vs '+mesesInd(cur,w,300,5));
  ok('shown: the split in euros and the months',txt('#tbl-tgt').includes('€116.67')&&txt('#tbl-tgt').includes('€183.33')&&new RegExp(A.meses+' months?').test(txt('#tgtOut')),txt('#tgtOut'));
  ok('fixed note: selling to rebalance realises taxable gains, contributions do not',/selling realises taxable capital gains/i.test(txt('#tgtNote'))&&/rebalancing with new contributions does not/i.test(txt('#tgtNote')),txt('#tgtNote'));
  ok('inputs show the saved targets',$('#tgt-w-SXR8').value==='60'&&$('#tgtBand').value==='5'&&$('#tgtMonthly').value==='300');}},
 tgtsplit2:{prep(d){tgtFix(d);seed(Object.assign(TGTSEED(),{targets:{weights:{SXR8:60,EUNK:30,BTC:10},band:5,monthly:10000,at:'2026-10-02T10:00:00.000Z'}}));},test(api){
  const A=api.alocacao(),L=id=>A.linhas.find(x=>x.id===id)||{};
  ok('every asset under its target after M: the deficits add up to M, so each gets exactly its deficit (€6,100 / €3,350 / €550)',near(L('SXR8').aloc,6100,1e-9)&&near(L('EUNK').aloc,3350,1e-9)&&near(L('BTC').aloc,550,1e-9),JSON.stringify(A.linhas.map(x=>[x.id,x.aloc])));
  ok('one month of contributions puts the portfolio back on target',A.meses===1,A.meses);
  ok('the split function also covers the case of deficits below M (cover them, the rest by the target weights)',(()=>{const r=api.reparte({A:100,B:0},{A:.5,B:.4},100);return near(r.B,88,1e-9)&&near(r.A,10,1e-9);})(),JSON.stringify(api.reparte({A:100,B:0},{A:.5,B:.4},100)));}},
 tgtnoprice:{prep(d){tgtFix(d);const a=d.ativos.find(x=>x.id==='EUNN');a.pontos=[];seed(Object.assign(TGTSEED(),{targets:{weights:{SXR8:50,EUNK:30,EUNN:20},band:5,monthly:1000,at:'2026-10-02T10:00:00.000Z'}}));},test(api){
  const A=api.alocacao(),L=id=>A.linhas.find(x=>x.id===id)||{},exp=reparteInd({SXR8:2000,EUNK:700,BTC:800},{SXR8:.625,EUNK:.375,BTC:0},1000);
  ok('an asset without a price is left out, named, and the other targets are scaled to 100%',A.excl.length===1&&/Japan/.test(A.excl[0])&&/Japan/.test(txt('#tgtOut'))&&L('EUNN').aloc==null&&near(L('SXR8').alvo,62.5),txt('#tgtOut'));
  ok('the split covers the other assets and adds up to M',['SXR8','EUNK','BTC'].every(id=>near(L(id).aloc,exp[id],1e-9))&&near(A.linhas.reduce((s,x)=>s+(x.aloc||0),0),1000,1e-9),JSON.stringify(A.linhas.map(x=>[x.id,x.aloc])));}},
 tgtsum:{prep(d){tgtFix(d);seed(TGTSEED());},test(api){
  const sv=get('savedAt'),f=(id,v)=>{$('#tgt-w-'+id).value=v;},sub=()=>$('#tgtForm').dispatchEvent(new Event('submit',{cancelable:true}));
  ok('no targets yet: a short prompt, current weights still shown',/Set a target/i.test(txt('#tgtOut'))&&/57\.1%/.test(txt('#tbl-tgt')),txt('#tgtOut'));
  f('SXR8','60');f('EUNK','30');$('#tgtBand').value='5';$('#tgtMonthly').value='250';sub();
  ok('targets that do not add up to 100%: error, nothing saved',/add up to 90/.test(txt('#tgtMsg'))&&get('targets')==null&&get('savedAt')===sv,txt('#tgtMsg'));
  f('BTC','10');$('#tgtBand').value='-1';sub();
  ok('a negative band is refused',/band/i.test(txt('#tgtMsg'))&&get('targets')==null,txt('#tgtMsg'));
  $('#tgtBand').value='';sub();const t=get('targets')||{};
  ok('valid targets are saved (band defaults to 5 pp) with their time',t.weights&&t.weights.SXR8===60&&t.weights.EUNK===30&&t.weights.BTC===10&&Object.keys(t.weights).length===3&&t.band===5&&t.monthly===250&&!isNaN(Date.parse(t.at))&&get('savedAt')!==sv,JSON.stringify(t));
  const b=api.dadosBackup();
  ok('the backup is version 5 and carries the targets',b.version===5&&JSON.stringify(b.targets)===JSON.stringify(t)&&Array.isArray(b.buys),JSON.stringify(b).slice(0,200));}},
 tgtmerge:{prep(d){tgtFix(d);seed(Object.assign(TGTSEED(),{targets:{weights:{SXR8:100},band:5,monthly:100,at:'2026-10-02T10:00:00.000Z'}}));
   d.backup={app:'Bluechip Board',version:5,saved:'2026-10-03T12:00:00.000Z',buys:[],lots:[],sales:[],deleted:{},targets:{weights:{SXR8:50,EUNK:50},band:3,monthly:400,at:'2026-10-03T12:00:00.000Z'}};},test(api){
  const t=get('targets')||{};
  ok('merge: the project file\'s newer targets win',t.weights&&t.weights.SXR8===50&&t.weights.EUNK===50&&t.band===3&&t.monthly===400,JSON.stringify(t));
  const r=api.juntaBackup({app:'Bluechip Board',version:5,saved:'2026-10-04T10:00:00.000Z',buys:[],lots:[],sales:[],deleted:{},targets:{weights:{BTC:100},band:5,monthly:0,at:'2026-10-01T10:00:00.000Z'}},false);
  ok('merge: older targets in a restored file do not replace newer ones here (and the file needs updating)',get('targets').weights.SXR8===50&&r.extraLocal,JSON.stringify([get('targets'),r]));
  api.juntaBackup({app:'Bluechip Board',version:5,saved:'2030-01-01T00:00:00.000Z',buys:[],lots:[],sales:[],deleted:{},targets:{weights:{'<img src=x onerror="window.__xss=9">':50,AAPL:50},band:5,monthly:0,at:'2030-01-01T00:00:00.000Z'}},false);
  ok('invalid targets in a file (unknown asset, not 100%) are ignored, nothing injected',get('targets').weights.SXR8===50&&!window.__xss&&api.limpaBackup({targets:{weights:{AAPL:'x'},at:'2030-01-01'}}).targets===null&&!document.querySelector('#tbl-tgt img'));}},
 tgtold:{prep(d){tgtFix(d);d.backup={app:'Bluechip Board',version:4,saved:'2026-10-03T12:00:00.000Z',buys:[{id:'A',a:'SXR8',d:'2026-09-01',q:1,p:200}],lots:[],sales:[],deleted:{}};},test(api){
  ok('version-4 backup without targets: loads, no targets',(get('buys')||[]).length===1&&get('targets')==null&&/Set a target/i.test(txt('#tgtOut')));
  api.juntaBackup({app:'Bluechip Board',version:3,saved:'2026-01-01T00:00:00.000Z',buys:[{id:'old',a:'NVDA',d:'2025-12-01',q:2,p:150}],lots:[{id:'L',d:'2025-12-02',q:0.1,c:8000}]},false);
  api.juntaBackup(V1,false);
  ok('version-3 and version-1 backups still merge, without targets',(get('buys')||[]).length===3&&(get('lots')||[]).length===1&&get('targets')==null,JSON.stringify(get('buys')));
  const b=api.dadosBackup();
  ok('saving without targets: version 5, no targets key',b.version===5&&!('targets' in b));}},

 /* ---------- política de investimento, notas por entrada e contexto nos alertas de queda ---------- */
 polctx:{prep(d){polFix(d);seed({deleted:{},buys:[],lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',
   policy:Object.assign(POLV(),{drop20:'Keep the monthly plan. '+HOSTIL(8),drop30:'Re-read why I bought it.',at:'2026-10-03T10:00:00.000Z'})});},test(api){
  const A=api.alerts(),a=A.find(x=>x.co==='AAPL'&&/below its 52-week high/.test(x.txt)),b=A.find(x=>x.co==='BTC'&&/below its 52-week high/.test(x.txt));
  ok('existing drop alerts keep their level and text (Apple −25%: important; Bitcoin −32%: important)',a&&a.l==='orange'&&a.txt==='Apple is 25.0% below its 52-week high. Check whether expected earnings have also fallen (volatility or deterioration?).'&&b&&b.l==='orange'&&b.txt==='Bitcoin is 32.0% below its 52-week high. Look for the cause (regulation, ETF flows, rates) in the news.',JSON.stringify([a,b]));
  const box=[...document.querySelectorAll('#alertas .alert')].find(x=>x.textContent.includes('Apple is 25.0% below')),c=box?txt2(box.querySelector('.alert-ctx')):'';
  ok('context line: the user\'s rule for a 20% drop, as text',/Your rule for a 20% drop: “Keep the monthly plan\. <img src=x/.test(c)&&!box.querySelector('.alert-ctx img')&&window.__xss===undefined,c);
  const rec=[[Date.parse('2011-03-01'),Date.parse('2011-12-01')],[Date.parse('2016-02-01'),Date.parse('2017-01-03')],[Date.parse('2020-03-23'),Date.parse('2020-08-03')]].map(x=>x[1]-x[0]).sort((x,y)=>x-y),med=Math.round(rec[1]/(30.44*864e5));
  ok('context: drops of 25% or more since 2010 counted with quedas() (4: 3 recovered, median time back to the peak, 1 not yet)',new RegExp('fell 25\\.0% or more from a peak 4 times').test(c)&&/3 recovered/.test(c)&&new RegExp('median of '+med+' months').test(c)&&/1 has not recovered yet/.test(c),c);
  ok('context: sample period and the rising-market caveat',/4 Jan 2010/.test(c)&&/1 Oct 2026/.test(c)&&/mostly a rising market/.test(c),c);
  const bx=[...document.querySelectorAll('#alertas .alert')].find(x=>x.textContent.includes('Bitcoin is 32.0% below')),cb=bx?txt2(bx.querySelector('.alert-ctx')):'';
  ok('Bitcoin −32%: the rule for a 30% drop is shown',/Your rule for a 30% drop: “Re-read why I bought it\.”/.test(cb)&&!/20% drop/.test(cb),cb);
  ok('only drop alerts get a context line',[...document.querySelectorAll('#alertas .alert')].filter(x=>x.querySelector('.alert-ctx')).every(x=>/below its 52-week high/.test(txt2(x.querySelector('.alert-msg'))))&&[...document.querySelectorAll('#alertas .alert-ctx')].length>=2);}},
 polnone:{prep(d){polFix(d);seed({deleted:{},buys:[],lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z'});},test(api){
  const a=api.alerts().find(x=>x.co==='AAPL'&&/below its 52-week high/.test(x.txt)),box=[...document.querySelectorAll('#alertas .alert')].find(x=>x.textContent.includes('Apple is 25.0% below')),c=box?txt2(box.querySelector('.alert-ctx')):'';
  ok('without a policy: same alert, context without a rule',a&&a.l==='orange'&&a.txt==='Apple is 25.0% below its 52-week high. Check whether expected earnings have also fallen (volatility or deterioration?).'&&!/Your rule/.test(c)&&/4 times/.test(c),c);}},
 polzero:{prep(d){polFix(d);d.historico.AAPL=[['2010-01-04',100],['2015-01-02',150],['2026-01-05',200]];seed({deleted:{},buys:[],lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z'});},test(api){
  const box=[...document.querySelectorAll('#alertas .alert')].find(x=>x.textContent.includes('Apple is 25.0% below')),c=box?txt2(box.querySelector('.alert-ctx')):'';
  ok('zero historical cases: said plainly, no median',/never fell 25\.0% or more from a peak/.test(c)&&!/median/.test(c)&&/4 Jan 2010/.test(c),c);}},
 polform:{prep(d){seed({deleted:{},buys:[],lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z'});},test(api){
  const sv=get('savedAt');[['horizon','15 years'],['allocation','Mostly ETFs'],['monthly','€400'],['drop20','Keep going'],['drop30','Review'],['sell','A broken thesis']].forEach(([k,v])=>{$('#pol-'+k).value=v;});
  $('#polForm').dispatchEvent(new Event('submit',{cancelable:true}));const p=get('policy')||{};
  ok('policy saved with its time, and it marks a change for the backup file',p.horizon==='15 years'&&p.allocation==='Mostly ETFs'&&p.monthly==='€400'&&p.drop20==='Keep going'&&p.drop30==='Review'&&p.sell==='A broken thesis'&&!isNaN(Date.parse(p.at))&&get('savedAt')!==sv,JSON.stringify(p));
  ok('the backup (version 5) carries the policy',api.dadosBackup().version===5&&JSON.stringify(api.dadosBackup().policy)===JSON.stringify(p)&&/Saved/.test(txt('#polMsg')),txt('#polMsg'));}},
 polnotes:{prep(d){window.prompt=()=>'Bought on the dip';seed({deleted:{},lots:[{id:'L1',d:'2026-02-02',q:0.01,c:700}],savedAt:'2026-10-03T10:00:00.000Z',
   buys:[{id:'n1',a:'SXR8',d:'2026-01-05',q:2,p:600,u:'2026-10-01'},{id:'n2',a:'EUNK',d:'2026-01-05',q:3,p:80,u:'2026-10-01'}],sales:[{id:'v1',a:'SXR8',d:'2026-03-02',q:1,p:650,u:'2026-10-01'}],
   notes:{n1:{t:'Core position',at:'2026-10-01T10:00:00.000Z'},v1:{t:'Paid the car',at:'2026-10-01T10:00:00.000Z'},gone:{t:'orphan',at:'2026-10-01T10:00:00.000Z'}}});},test(api){
  ok('notes are shown in the register, for purchases and sales',/Core position/.test(txt('#tbl-buys'))&&/Paid the car/.test(txt('#tbl-buys')),txt('#tbl-buys').slice(0,300));
  ok('a note without an entry is ignored and not saved in the backup',!/orphan/.test(txt('#tbl-buys'))&&!('gone' in (api.dadosBackup().notes||{}))&&'n1' in api.dadosBackup().notes);
  const sv=get('savedAt'),btn=document.querySelector('#tbl-buys [data-note="n2"]');if(btn)btn.click();const n=(get('notes')||{}).n2||{};
  ok('a note can be added from the register (and marks a change)',n.t==='Bought on the dip'&&!isNaN(Date.parse(n.at))&&get('savedAt')!==sv&&/Bought on the dip/.test(txt('#tbl-buys')),JSON.stringify(get('notes')));
  ok('the entries themselves are unchanged (notes live beside them)',get('buys').concat(get('sales'),get('lots')).every(x=>Object.keys(x).every(k=>['id','a','d','q','p','c','r','u'].includes(k))),JSON.stringify([get('buys'),get('sales'),get('lots')]));
  window.confirm=()=>true;const del=document.querySelector('#tbl-buys [data-del="n1"]');if(del)del.click();
  ok('deleting an entry removes its note',!('n1' in (get('notes')||{}))&&!('n1' in (api.dadosBackup().notes||{})),JSON.stringify(get('notes')));}},
 polmerge:{prep(d){seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-02T10:00:00.000Z',buys:[{id:'m1',a:'SXR8',d:'2026-01-05',q:1,p:600,u:'2026-10-01'},{id:'m2',a:'EUNK',d:'2026-01-05',q:1,p:80,u:'2026-10-01'}],
   policy:Object.assign(POLV(),{horizon:'local',at:'2026-10-02T10:00:00.000Z'}),notes:{m1:{t:'local m1',at:'2026-10-02T10:00:00.000Z'},m2:{t:'local m2',at:'2026-10-05T10:00:00.000Z'}}});
   d.backup={app:'Bluechip Board',version:5,saved:'2026-10-03T12:00:00.000Z',deleted:{},lots:[],sales:[],buys:[{id:'m1',a:'SXR8',d:'2026-01-05',q:1,p:600},{id:'m2',a:'EUNK',d:'2026-01-05',q:1,p:80}],
    policy:Object.assign(POLV(),{horizon:'file',at:'2026-10-03T12:00:00.000Z'}),notes:{m1:{t:'file m1',at:'2026-10-03T12:00:00.000Z'},m2:{t:'file m2',at:'2026-10-03T12:00:00.000Z'}}};},test(api){
  ok('merge: the newer policy wins (the file\'s)',(get('policy')||{}).horizon==='file');
  ok('merge: each note keeps its newer version (m1 from the file, m2 from this browser)',get('notes').m1.t==='file m1'&&get('notes').m2.t==='local m2',JSON.stringify(get('notes')));
  const r=api.juntaBackup({app:'Bluechip Board',version:5,saved:'2026-10-04T10:00:00.000Z',buys:[],lots:[],sales:[],deleted:{},policy:Object.assign(POLV(),{horizon:'older',at:'2026-09-01T10:00:00.000Z'})},false);
  ok('merge: an older policy in a restored file does not replace the newer one (the file needs updating)',get('policy').horizon==='file'&&r.extraLocal);
  ok('invalid policy or notes in a file are ignored',api.limpaBackup({policy:'x',notes:{a:{t:5}}}).policy===null&&Object.keys(api.limpaBackup({notes:{a:{t:5},b:'x'}}).notes).length===0);}},
 polold:{prep(d){d.backup={app:'Bluechip Board',version:4,saved:'2026-10-03T12:00:00.000Z',buys:[{id:'A',a:'SXR8',d:'2026-09-01',q:1,p:200}],lots:[],sales:[],deleted:{}};},test(api){G.rows=document.querySelectorAll('#tbl-buys tbody tr').length;
  api.juntaBackup({app:'Bluechip Board',version:3,saved:'2026-01-01T00:00:00.000Z',buys:[{id:'old',a:'NVDA',d:'2025-12-01',q:2,p:150}],lots:[{id:'L',d:'2025-12-02',q:0.1,c:8000}]},false);api.juntaBackup(V1,false);
  const b=api.dadosBackup();
  ok('version 1, 3 and 4 backups load without policy or notes, and the page works',(get('buys')||[]).length===3&&get('policy')==null&&get('notes')==null&&!('policy' in b)&&!('notes' in b)&&$('#pol-horizon').value===''&&G.rows===1,JSON.stringify([get('buys'),get('policy'),get('notes'),G.rows]));}},

 /* ---------- teste de stress (episódios do S&P 500) e comparação de estratégias ---------- */
 stress:{prep(d){stFix(d);seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'s1',a:'SXR8',d:'2025-01-06',q:4,p:200,u:'2026-10-01'},{id:'s2',a:'AAPL',d:'2025-01-06',q:10,p:150,u:'2026-10-01'},{id:'s3',a:'IS3N',d:'2025-01-06',q:10,p:40,u:'2026-10-01'}]});},test(api){
  const S=api.stress(),E=id=>S.eps.find(e=>e.id===id)||{},L=(e,id)=>(E(e).linhas||[]).find(x=>x.id===id)||{},eur=(id,iso)=>{const h=STH[id],k=Object.keys(h).filter(x=>x<=iso).sort().pop();return id==='AAPL'?h[k]/STFX[k]:h[k];};
  ok('episodes: the S&P 500 peak and trough dates',JSON.stringify(api.EPISODIOS.map(e=>[e.pk,e.tr]))==='[["2018-09-20","2018-12-24"],["2020-02-19","2020-03-23"],["2022-01-03","2022-10-12"]]');
  const vA=10*200/1.25,rA=eur('AAPL','2018-12-24')/eur('AAPL','2018-09-20')-1,rS=eur('SXR8','2018-12-24')/eur('SXR8','2018-09-20')-1;
  ok('2018 Q4: return in euros between the closes (AAPL at each day\'s EUR/USD; SXR8 uses 21 Dec, the last close before 24 Dec) applied to today\'s values',near(L('2018Q4','AAPL').r,rA,1e-12)&&near(L('2018Q4','AAPL').perda,vA*rA,1e-9)&&near(L('2018Q4','SXR8').r,rS,1e-12)&&near(L('2018Q4','SXR8').perda,1000*rS,1e-9)&&api.isoU(L('2018Q4','SXR8').tTr)==='2018-12-21',JSON.stringify(E('2018Q4').linhas));
  ok('an episode outside an asset\'s history: IS3N left out of 2018 Q4 and COVID and named, included in 2022',!L('2018Q4','IS3N').ok&&E('2018Q4').excl.some(n=>/IS3N|EM IMI/.test(n))&&!L('COVID','IS3N').ok&&L('2022','IS3N').ok&&near(L('2022','IS3N').r,26/33-1,1e-12),JSON.stringify([E('2018Q4').excl,E('2022').excl]));
  ok('total and % only over the assets with data',near(E('2018Q4').total,vA*rA+1000*rS,1e-9)&&near(E('2018Q4').pct,(vA*rA+1000*rS)/(vA+1000)*100,1e-9)&&near(E('2022').base,vA+1000+500,1e-9),JSON.stringify([E('2018Q4').total,E('2018Q4').pct,E('2022').base]));
  ok('time back to the starting value, after the trough (in euros)',L('2018Q4','AAPL').rec===Date.parse('2019-04-01')&&L('2018Q4','SXR8').rec===Date.parse('2019-06-03')&&L('COVID','AAPL').rec===Date.parse('2020-06-01')&&L('2022','IS3N').rec===Date.parse('2026-10-01'),JSON.stringify([L('2018Q4','AAPL').rec,L('2018Q4','SXR8').rec,L('COVID','AAPL').rec,L('2022','IS3N').rec]));
  const tb=txt('#tbl-stress');
  ok('shown: loss in euros and %, recovery time, no data for IS3N, a clear hypothetical label',tb.includes(api.eur(L('2018Q4','AAPL').perda,0))&&/back in \d+ months?/.test(tb)&&/no data for this period/i.test(tb)&&/hypothetical/i.test(txt('#stressLead')),tb.slice(0,400));}},
 stressnone:{prep(d){seed({deleted:{},buys:[],lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z'});},test(api){
  ok('no holdings: a short message, no figures',api.stress().held===0&&!/€/.test(txt('#stressOut'))&&/purchases/i.test(txt('#stressOut')),txt('#stressOut'));}},
 strat:{prep(d){stratFix(d);seed({deleted:{},buys:[],lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z','sim.btc':{amt:100,from:'2021-01'}});},test(api){
  const h=api.HIST.BTC,I=stratInd(h,100,'2021-01',10),R=api.estrategias(h,100,'2021-01',10);
  ok('existing monthly simulator: same numbers as before (independent calculation), and its saved inputs unchanged',/\€?Invested/.test(txt('#sim-btc-kpis'))&&txt('#sim-btc-kpis').includes(api.eur(I.inv))&&txt('#sim-btc-kpis').includes(api.eur(I.a))&&near(api.simular(h,100,'2021-01').u,I.uA,1e-12)&&JSON.stringify(get('sim.btc'))==='{"amt":100,"from":"2021-01"}',txt('#sim-btc-kpis'));
  ok('(a) monthly, (b) all at once, (c) buy the dip at 10%: final values match an independent calculation',R&&near(R.a.val,I.a,1e-9)&&near(R.b.val,I.b,1e-9)&&near(R.c.val,I.c,1e-9)&&near(R.tot,I.inv,1e-12)&&R.c.compras===I.buys&&near(R.c.cash,I.cash,1e-9),JSON.stringify([R&&[R.a.val,R.b.val,R.c.val,R.c.compras],I]));
  ok('same money and same period for the three',R.tot===I.inv&&R.n===I.n&&txt('#strat-btc').includes(api.eur(I.inv)));
  ok('shown: final value and gain of each strategy, and a chart with three lines',['#strat-btc'].every(s=>txt(s).includes(api.eur(I.a))&&txt(s).includes(api.eur(I.b))&&txt(s).includes(api.eur(I.c)))&&document.querySelectorAll('#strat-btc-ch svg path').length>=3,txt('#strat-btc').slice(0,300));
  const x=$('#strat-btc-dip');x.value='20';x.dispatchEvent(new Event('input'));const I2=stratInd(h,100,'2021-01',20);
  ok('X is adjustable (20%): the dip strategy is recalculated, the monthly simulator\'s inputs stay as they were',txt('#strat-btc').includes(api.eur(I2.c))&&JSON.stringify(get('sim.btc'))==='{"amt":100,"from":"2021-01"}'&&(get('strat.btc')||{}).dip===20,txt('#strat-btc').slice(0,300));
  ok('ETF tab: the same comparison under its monthly simulator',!!$('#strat-etf')&&/All at once/.test(txt('#strat-etf'))&&$('#sim-etf-outros').compareDocumentPosition($('#strat-etf'))&Node.DOCUMENT_POSITION_FOLLOWING);}},

 /* ---------- retornos rolantes, underwater e escala log ---------- */
 roll:{prep(d){rollFix(d);},test(api){
  const h=api.HIST.SXR8,R=[1,3,5].map(a=>api.rolar(h,a)),I=[1,3,5].map(a=>rollInd(h,a,252)),box=$('#roll-etf');
  ok('rolling returns are only calculated when the section is opened',box&&!box.open&&!$('#roll-etf-body svg')&&!/Worst/.test(txt('#roll-etf-body')));
  box.open=true;box.dispatchEvent(new Event('toggle'));
  ok('1, 3 and 5 years over every daily window (252 sessions a year): matches an independent calculation',R.every((r,i)=>r&&r.k===252*[1,3,5][i]&&r.n===I[i].n&&near(r.min,I[i].min,1e-12)&&near(r.med,I[i].med,1e-12)&&near(r.max,I[i].max,1e-12)&&near(r.neg,I[i].neg,1e-12)),JSON.stringify([R.map(r=>r&&[r.n,r.min,r.med,r.max,r.neg]),I.map(r=>[r.n,r.min,r.med,r.max,r.neg])]));
  ok('3 and 5 years also annualised (not for 1 year)',R[0].ann===null&&near(R[1].ann.med,I[1].ann.med,1e-12)&&near(R[2].ann.min,I[2].ann.min,1e-12)&&near(R[2].ann.max,I[2].ann.max,1e-12));
  const tb=txt('#roll-etf-body');
  ok('shown: worst, median, best, share of negative windows, number of windows, overlap warning',tb.includes(api.pct(I[0].min*100))&&tb.includes(api.pct(I[0].med*100))&&tb.includes(api.pct(I[2].ann.med*100))&&tb.includes(api.nf(I[0].n,0))&&/overlap/i.test(tb)&&/negative/i.test(tb),tb.slice(0,500));
  ok('a histogram (barChart) of the windows, by default 1 year',!!$('#roll-etf-ch svg')&&$('#roll-etf-ch svg').querySelectorAll('rect').length>=4&&/1 year/.test($('#roll-etf-ch svg').getAttribute('aria-label')),$('#roll-etf-ch svg')&&$('#roll-etf-ch svg').getAttribute('aria-label'));
  const b3=document.querySelector('#roll-etf-h button[data-y="3"]');if(b3)b3.click();
  ok('the histogram can show 3 years',/3 years/.test(($('#roll-etf-ch svg')||{getAttribute:()=>''}).getAttribute('aria-label')));
  const U=api.underwater(h);let mx=0;const ui=h.map(p=>{mx=Math.max(mx,p[1]);return p[1]/mx-1;});
  ok('underwater: price / running maximum − 1, next to the biggest-drops table',U.length===h.length&&U.every((p,i)=>p[0]===h[i][0]&&near(p[1],ui[i]*100,1e-9))&&!!$('#ch-etf-dd svg')&&$('#ch-etf-dd').closest('.dd-grid')===$('#tbl-drops-etf').closest('.dd-grid')&&!!$('#ch-etf-dd').closest('.dd-grid'));
  const lab=()=>($('#ch-etf-hist svg')||{getAttribute:()=>''}).getAttribute('aria-label');
  ok('long-term chart: linear by default, as before',lab()==='SXR8 long-term price in euros'&&document.querySelector('#f-elog button[data-log="0"]').getAttribute('aria-pressed')==='true',lab());
  document.querySelector('#f-elog button[data-log="1"]').click();const logL=lab();
  document.querySelector('#f-elog button[data-log="0"]').click();
  ok('log scale on and off',/logarithmic scale/.test(logL)&&lab()==='SXR8 long-term price in euros',logL+' / '+lab());}},
 rollshort:{prep(d){rollFix(d);},test(api){
  api.escolheEtf('IS3N');const box=$('#roll-etf');box.open=true;box.dispatchEvent(new Event('toggle'));const h=api.HIST.IS3N,I3=rollInd(h,3,252);
  ok('short history (IS3N since 2023 here): 5 years Unavailable, 1 and 3 years calculated',api.rolar(h,5)===null&&api.rolar(h,3)&&api.rolar(h,3).n===I3.n&&/Unavailable/.test(txt('#roll-etf-body'))&&near(api.rolar(h,3).med,I3.med,1e-12),txt('#roll-etf-body').slice(0,400));
  ok('the chosen fund is used (IS3N), with its own underwater chart',/IS3N/.test(txt('#roll-etf-body'))&&!!$('#ch-etf-dd svg'));}},
 rollbtc:{prep(d){rollFix(d);},test(api){
  const h=api.HIST.BTC,R=api.rolar(h,1),I=rollInd(h,1,365),box=$('#roll-btc');
  ok('Bitcoin: lazy too',box&&!$('#roll-btc-body svg'));box.open=true;box.dispatchEvent(new Event('toggle'));
  ok('Bitcoin: calendar days (365 a year), against an independent calculation',R&&R.k===365&&R.n===I.n&&near(R.med,I.med,1e-12)&&near(R.neg,I.neg,1e-12)&&!!$('#roll-btc-ch svg'),JSON.stringify([R&&[R.k,R.n,R.med],I&&[I.n,I.med]]));}},

 /* ---------- exposição agregada (país, setor, moeda) e concentração por empresa ---------- */
 expo:{prep(d){expoFix(d);seed(EXSEED());},test(api){
  const X=api.exposicao(),L=(dim,n)=>(X.dims[dim].L.find(x=>x.n===n)||{}).v,cob=3020-120;
  ok('values weighted by today\'s holdings: total €3,020 (SXR8 1,000, EUNK 700, EUNN 120, Apple 800, Bitcoin 400)',near(X.total,3020,1e-9),X.total);
  ok('countries: United States = SXR8 (index) + Apple, the EUNK countries from its file, Bitcoin as Crypto',near(L('paises','United States'),998+800,1e-9)&&near(L('paises','United Kingdom'),280,1e-9)&&near(L('paises','Switzerland'),171.5,1e-9)&&near(L('paises','Cash/Other'),2+3.5,1e-9)&&near(L('paises','Crypto'),400,1e-9),JSON.stringify(X.dims.paises.L));
  ok('a fund without aggregates (EUNN) is left out and the uncovered share of the portfolio is shown',near(X.dims.paises.sem,120,1e-9)&&X.dims.paises.semN.some(n=>/Japan|EUNN/.test(n))&&near(X.dims.paises.L.reduce((s,x)=>s+x.v,0),cob,1e-9)&&new RegExp(api.nf(120/3020*100,1)+'%').test(txt('#expoOut')),txt('#expoOut'));
  ok('percentages are of the covered value',near(X.dims.paises.L.find(x=>x.n==='United States').p,(998+800)/cob*100,1e-9));
  ok('sectors: Apple from the iShares file, the funds from their aggregates',near(L('setores','Information Technology'),400+800,1e-9)&&near(L('setores','Financials'),300+350,1e-9)&&near(L('setores','Crypto'),400,1e-9));
  ok('underlying currencies: SXR8 counts as USD, with the note',near(L('moedas','USD'),998+800,1e-9)&&near(L('moedas','GBP'),280,1e-9)&&/SXR8 is quoted in EUR, but its underlying currency exposure is USD/.test(txt('#expoNote')),txt('#expoNote'));
  ok('the country from the index is flagged as such',/index/i.test(txt('#expoNote'))&&/approximate/i.test(txt('#expoNote')));
  ok('three charts drawn',['#ch-expo-paises','#ch-expo-setores','#ch-expo-moedas'].every(s=>$(s+' svg')||$(s+' .hb')||$(s).children.length>0));}},
 expostocks:{prep(d){expoFix(d);d.etfs.SXR8=Object.assign({},d.etfs.SXR8,{top10:[]});seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a',a:'AAPL',d:'2025-01-06',q:5,p:150,u:'2026-10-01'},{id:'n',a:'NVDA',d:'2025-01-06',q:4,p:90,u:'2026-10-01'},{id:'g',a:'GOOGL',d:'2025-01-06',q:2,p:120,u:'2026-10-01'}]});},test(api){
  const X=api.exposicao(),L=(dim,n)=>(X.dims[dim].L.find(x=>x.n===n)||{}).v;
  ok('stocks only: 100% United States and USD, nothing uncovered',near(L('paises','United States'),800+320+240,1e-9)&&X.dims.paises.L.length===1&&near(L('moedas','USD'),1360,1e-9)&&X.dims.paises.sem===0&&!/not covered/i.test(txt('#expoOut')),JSON.stringify(X.dims));
  ok('sectors from the explicit configuration when the iShares file does not list the company',near(L('setores','Information Technology'),1120,1e-9)&&near(L('setores','Communication'),240,1e-9),JSON.stringify(X.dims.setores.L));}},
 expoold:{prep(d){expoFix(d);d.versao='1.3';['SXR8','EUNK','IS3N','EUNN'].forEach(id=>{if(d.etfs[id])delete d.etfs[id].agregados;});if(d.etf)delete d.etf.agregados;seed(EXSEED());},test(api){
  const X=api.exposicao();
  ok('a version-1.3 data file still loads: the funds are Unavailable, stocks and Bitcoin still counted',near(X.dims.paises.sem,1820,1e-9)&&near(X.dims.paises.L.reduce((s,x)=>s+x.v,0),1200,1e-9)&&/Unavailable/.test(txt('#expoOut'))&&!!$('#tbl-pf tbody tr'),txt('#expoOut'));}},
 conc:{prep(d){expoFix(d);seed(EXSEED());},test(api){
  const C=api.exposicao().conc,f=n=>C.filter(x=>x.n===n);
  ok('panel company: direct + through the ETF (Apple €800 + 7% of SXR8 €1,000 = €870, 28.8%), above the 10% threshold',f('Apple').length===1&&near(f('Apple')[0].v,870,1e-9)&&f('Apple')[0].acima&&api.CONC_LIMIAR===10);
  ok('top 10 of each fund on its own: ASML 44% of EUNK = €308 (10.2%) above, Microsoft below',near(f('ASML')[0].v,308,1e-9)&&f('ASML')[0].acima&&!f('Microsoft')[0].acima);
  ok('never added up by name across funds: Shell appears once per fund',f('Shell').length===2&&f('Shell').every(x=>!x.acima),JSON.stringify(f('Shell')));
  ok('informative note under the section, not a new Overview alert',/Apple/.test(txt('#concOut'))&&/ASML/.test(txt('#concOut'))&&/10%/.test(txt('#concOut'))&&!api.alerts().some(a=>/concentrat|threshold/i.test(a.txt)),txt('#concOut'));}},

 /* ---------- "Before you sell" e comissões ---------- */
 sellsim:{prep(d){expoFix(d);seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'a1',a:'AAPL',d:'2024-01-05',q:10,p:100,u:'2026-10-01'},{id:'a2',a:'AAPL',d:'2025-01-06',q:5,p:150,u:'2026-10-01'}],
   sales:[{id:'s1',a:'AAPL',d:'2025-06-02',q:4,p:160,u:'2026-10-01'}]});},test(api){
  const ls=lsSnap(),bk=semExport(api.dadosBackup()),S=api.simulaVenda('AAPL',8);
  ok('FIFO over several lots: 6 left from the first purchase, then 2 from the second',S&&!S.erro&&S.lotes.length===2&&S.lotes[0].id==='a1'&&near(S.lotes[0].q,6)&&S.lotes[1].id==='a2'&&near(S.lotes[1].q,2),JSON.stringify(S&&S.lotes));
  ok('gain at today\'s price (€160) and a 28% estimate on the positive gain',near(S.valor,1280,1e-9)&&near(S.custo,900,1e-9)&&near(S.ganho,380,1e-9)&&near(S.imposto,106.4,1e-9));
  $('#sellAsset').value='AAPL';$('#sellQty').value='8';$('#sellForm').dispatchEvent(new Event('submit',{cancelable:true}));
  ok('shown: lots used, gain, tax estimate and the simplification note',txt('#sellOut').includes('€380')&&txt('#sellOut').includes('€106')&&/28%/.test(txt('#sellOut'))&&/simplification/i.test(txt('#sellOut'))&&/loss/i.test(txt('#sellOut'))&&document.querySelectorAll('#tbl-sell tbody tr').length===2,JSON.stringify([txt('#sellOut').includes('€380'),txt('#sellOut').includes('€106'),/28%/.test(txt('#sellOut')),/simplification/i.test(txt('#sellOut')),/loss/i.test(txt('#sellOut')),document.querySelectorAll('#tbl-sell tbody tr').length]));
  $('#sellQty').value='20';$('#sellForm').dispatchEvent(new Event('submit',{cancelable:true}));
  ok('a sale larger than the holding is refused, as in the register',/do not hold enough/i.test(txt('#sellOut'))&&api.simulaVenda('AAPL',20).erro,txt('#sellOut'));
  ok('nothing is saved: browser storage and backup unchanged',lsSnap()===ls&&semExport(api.dadosBackup())===bk);}},
 sellbtc:{prep(d){expoFix(d);seed({deleted:{},buys:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',lots:[{id:'L1',d:diasAntes(400),q:0.1,c:3000},{id:'L2',d:diasAntes(350),q:0.1,c:5000}]});},test(api){
  const S=api.simulaVenda('BTC',0.15),L=id=>S.lotes.find(x=>x.id===id)||{};
  ok('Bitcoin: the lot held 400 days is exempt, the lot held 350 days is taxable',L('L1').isento&&!L('L2').isento&&near(L('L1').q,0.1)&&near(L('L2').q,0.05),JSON.stringify(S.lotes));
  ok('tax estimate only on the taxable part (€1,500 gain × 28%)',near(S.ganhoTrib,1500,1e-9)&&near(S.imposto,420,1e-9)&&near(S.ganho,6500,1e-9));
  $('#sellAsset').value='BTC';$('#sellQty').value='0.15';$('#sellForm').dispatchEvent(new Event('submit',{cancelable:true}));
  ok('warning for a lot less than 30 days from becoming exempt',L('L2').faltam===15&&/15 days/.test(txt('#sellOut'))&&/exempt/i.test(txt('#sellOut')),txt('#sellOut'));}},
 fees:{prep(d){window.prompt=()=>'12.5';seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',
   buys:[{id:'b1',a:'AAPL',d:'2025-01-06',q:10,p:100,u:'2026-10-01'},{id:'b2',a:'AAPL',d:'2025-02-03',q:10,p:110,u:'2026-10-01'},{id:'g1',a:'GOOGL',d:'2025-01-06',q:5,p:150,u:'2026-10-01'}],
   sales:[{id:'s1',a:'AAPL',d:'2026-03-02',q:15,p:150,u:'2026-10-01'},{id:'g2',a:'GOOGL',d:'2026-03-02',q:5,p:170,u:'2026-10-01'}],
   fees:{b1:{v:10,at:'2026-10-01T10:00:00.000Z'},b2:{v:7,at:'2026-10-01T10:00:00.000Z'},s1:{v:4.99,at:'2026-10-01T10:00:00.000Z'},g2:{v:2,at:'2026-10-01T10:00:00.000Z'},gone:{v:9,at:'2026-10-01T10:00:00.000Z'}}});},test(api){
  const X=api.anexoJ(2026),r=(l,v)=>X.acoes.find(x=>x.lote===l&&x.venda===v)||{},R=csvParse(X.csvAcoes),col=h=>csvCol(R,h);
  ok('fees split per pair: purchase fee × used/bought + sale fee × pair/sale (10 + 3.33; 3.50 + 1.66)',near(r('b1','s1').desp,13.33,1e-9)&&near(r('b2','s1').desp,5.16,1e-9),JSON.stringify([r('b1','s1').desp,r('b2','s1').desp]));
  ok('rounded to cents with the remainder on the sale\'s last row, so the total matches (€18.49)',near(r('b1','s1').desp+r('b2','s1').desp,18.49,1e-9));
  ok('CSV: Despesas e encargos filled, and no longer in "Fill in manually" when both fees are known',col('Despesas e encargos (EUR)')[X.acoes.indexOf(r('b1','s1'))]==='13,33'&&!/Despesas e encargos/.test(col('Fill in manually')[X.acoes.indexOf(r('b1','s1'))]),col('Fill in manually').join(' | '));
  ok('only one fee known (GOOGL: sale fee only): filled with it and flagged FEE_PARTIAL, still in "Fill in manually"',near(r('g1','g2').desp,2,1e-9)&&r('g1','g2').flags.includes('FEE_PARTIAL')&&/Despesas e encargos/.test(col('Fill in manually')[X.acoes.indexOf(r('g1','g2'))]),JSON.stringify(r('g1','g2')));
  ok('the realised gain shown in the register does not change; fees have their own column',/€10\.00/.test(txt('#tbl-buys'))&&/Fee/.test(txt('#tbl-buys thead'))&&r('b1','s1').ganho===500,txt('#tbl-buys thead'));
  const F=api.rentab().fl;
  ok('XIRR includes the fees (purchase −(1,000 + 10); sale +(2,250 − 4.99))',near(F[0][1],-1010,1e-9)&&F.some(x=>near(x[1],2250-4.99,1e-9)),JSON.stringify(F));
  ok('a fee without an entry is ignored and not saved in the backup',!('gone' in (api.dadosBackup().fees||{}))&&'b1' in api.dadosBackup().fees);
  const sv=get('savedAt'),b=document.querySelector('#tbl-buys [data-fee="g1"]');if(b)b.click();
  ok('a fee can be added from the register (saved, marks a change)',(get('fees')||{}).g1&&get('fees').g1.v===12.5&&get('savedAt')!==sv,JSON.stringify(get('fees')));
  const j=api.juntaBackup({app:'Bluechip Board',version:5,saved:'2026-10-05T10:00:00.000Z',buys:[],lots:[],sales:[],deleted:{},fees:{b1:{v:99,at:'2030-01-01T00:00:00.000Z'},b2:{v:1,at:'2020-01-01T00:00:00.000Z'}}},false);
  ok('merge: each fee keeps its most recent version',get('fees').b1.v===99&&get('fees').b2.v===7,JSON.stringify(get('fees')));}},
 feesold:{prep(d){d.backup={app:'Bluechip Board',version:4,saved:'2026-10-03T12:00:00.000Z',buys:[{id:'A',a:'SXR8',d:'2026-09-01',q:1,p:200}],lots:[],sales:[],deleted:{}};},test(api){
  api.juntaBackup({app:'Bluechip Board',version:3,saved:'2026-01-01T00:00:00.000Z',buys:[{id:'old',a:'NVDA',d:'2025-12-01',q:2,p:150}],lots:[{id:'L',d:'2025-12-02',q:0.1,c:8000}]},false);api.juntaBackup(V1,false);
  ok('version 1, 3 and 4 backups load without fees; no fees key when saving',(get('buys')||[]).length===3&&get('fees')==null&&!('fees' in api.dadosBackup())&&api.anexoJ(2026).acoes.every(r=>r.desp==null));}},

 /* ---------- os três ETF novos (EUNK, IS3N, EUNN) ---------- */
 etfui:{test(api){
  const chips=[...document.querySelectorAll('.chip[data-co]')].map(b=>b.dataset.co).join(',');
  ok('filter: one chip per new ETF, before Bitcoin',chips==='all,AAPL,NVDA,GOOGL,MKT,EUNK,IS3N,EUNN,BTC',chips);
  ok('ETFs come from the script configuration, SXR8 first',api.ETF_IDS.join(',')==='SXR8,EUNK,IS3N,EUNN',api.ETF_IDS.join(','));
  ok('new ETF cards: name, euro price and Xetra date',['EUNK','IS3N','EUNN'].every(x=>{const t=$(`.tk[data-tk="${x}"]`);return t&&/ETF MSCI/.test(t.textContent)&&/€/.test(txt(`.tk[data-tk="${x}"] .tk-price`))&&!!t.querySelector('.asof')&&!!t.querySelector('.tk-spark svg')&&!!t.querySelector('.range');}),txt('.tk[data-tk="EUNK"]'));
  const a=api.ATIVOS.find(x=>x.id==='EUNK');
  ok('euro-listed ETF: no EUR/USD conversion in euros, converted in dollars',api.inCur(a,'EUR')===a.pts&&near(api.inCur(a,'USD').slice(-1)[0][1],a.pts.slice(-1)[0][1]*api.fxAt(a.pts.slice(-1)[0][0]),1e-12));
  $('.chip[data-co="EUNK"]').click();
  ok('chip EUNK: prices table shows only EUNK',$('#tbl-ret tbody').rows.length===1&&txt('#tbl-ret tbody').startsWith('ETF EUNK'),txt('#tbl-ret tbody'));
  ok('chip EUNK: risk table and distance chart for EUNK',$('#tbl-ind tbody').rows.length===1&&/ETF EUNK/.test(txt('#ch-dist')));
  ok('chip EUNK: other cards dimmed',[...document.querySelectorAll('.tk')].filter(t=>!t.classList.contains('dim')).map(t=>t.dataset.tk).join(',')==='EUNK');
  const arts=[...document.querySelectorAll('#newsList article')];
  ok('chip EUNK: only EUNK news',arts.length>0&&arts.every(x=>x.textContent.includes('ETF EUNK')),arts.length);
  ok('chip EUNK: alerts only about EUNK',[...document.querySelectorAll('#alertas .alert')].every(x=>/ETF EUNK|You/.test(x.textContent)));
  ok('chip EUNK: calendar has no invented ETF events',[...$('#tbl-cal tbody').rows].every(r=>/You|No events/.test(r.textContent)),txt('#tbl-cal tbody'));
  ok('chip EUNK: the ETF tab shows EUNK',api.etfSel()==='EUNK'&&/IE00B4K48X80/.test(txt('#etfLead'))&&/MSCI Europe/.test(txt('#etfLead'))&&txt('#h-etf')==='iShares Core MSCI Europe',txt('#etfLead'));
  ok('EUNK: the "three companies" figure is hidden (it holds none of them)',$('#fig-etf-w').hidden);
  const w=api.etfInfo('EUNK');
  ok('EUNK: its own top 10 and sectors',arr(w.top10).length===10&&txt('#ch-top10').includes(w.top10[0].n)&&!/Your three companies/.test(txt('#top10Sub'))&&new RegExp('out of '+w.posicoes).test(txt('#top10Sub'))&&!!$('#ch-sect svg'),txt('#top10Sub'));
  ok('EUNK: long-term chart and drops table from its own history',txt('#etfHistTitle')==='EUNK price in euros'&&!!$('#ch-etf-hist svg')&&$('#tbl-drops-etf tbody').rows.length>0);
  const amt=+$('#sim-etf-EUNK-amt').value,from=$('#sim-etf-EUNK-from').value,r=api.simular(api.HIST.EUNK,amt,from),val=r.u*r.last;
  ok('EUNK: simulator on its own prices; the SXR8 simulator is hidden',!$('#sim-etf-EUNK').hidden&&$('#sim-etf').hidden&&txt('#sim-etf-EUNK-kpis').includes('€'+Math.round(val).toLocaleString('en-GB')),txt('#sim-etf-EUNK-kpis'));
  api.escolheEtf('IS3N');
  ok('IS3N: history from 2014, so "since 2014"',txt('#etfDropsTitle')==='Biggest drops since 2014'&&txt('#elp-max')==='Since 2014'&&!$('#sim-etf-IS3N').hidden);
  api.escolheEtf('SXR8');
  ok('SXR8 view as before',txt('#h-etf')==='iShares Core S&P 500'&&/^The iShares Core S&P 500 UCITS ETF \(Acc\), IE00B5BMR087, listed as SXR8 on Xetra in euros, holds about .*% in these three companies\./.test(txt('#etfLead'))&&!$('#fig-etf-w').hidden&&txt('#etfDropsTitle')==='Biggest drops since 2010'&&/Your three companies are highlighted/.test(txt('#top10Sub'))&&!$('#sim-etf').hidden,txt('#etfLead'));
  $('.chip[data-co="MKT"]').click();
  ok('"ETF & market" chip unchanged: SXR8 and the S&P 500',$('#tbl-ret tbody').rows.length===2&&/ETF SXR8/.test(txt('#tbl-ret tbody'))&&/S&P 500/.test(txt('#tbl-ret tbody')));
  ok('accumulating ETFs are not in Dividends',api.DIV_IDS.join(',')==='AAPL,NVDA,GOOGL'&&!/EUNK|IS3N|EUNN|MSCI/.test(txt('#tbl-div')));
  ok('the new ETFs follow the Xetra calendar',['EUNK','IS3N','EUNN'].every(x=>api.BOLSA_DE[x]==='DE')&&api.frescura('EUNK',a.pts).un==='session');
  ok('Bitcoin correlation includes the new ETFs',/ETF EUNK/.test(txt('#ch-corr'))&&/ETF EUNN/.test(txt('#ch-corr')));}},

 etfpf:{prep(d){capDl();seed({deleted:{},lots:[],savedAt:'2026-10-03T10:00:00.000Z',
   buys:[{id:'e1',a:'EUNK',d:'2026-01-05',q:10.5,p:100,u:'2026-10-02'},{id:'e2',a:'EUNK',d:'2026-02-02',q:5,p:105,u:'2026-10-02'},{id:'i1',a:'IS3N',d:'2026-01-05',q:20,p:45,u:'2026-10-02'},
    {id:'j1',a:'EUNN',d:'2026-03-02',q:3.25,p:70,u:'2026-10-02'},{id:'x1',a:'SXR8',d:'2026-01-05',q:1,p:700,u:'2026-10-02'}],
   sales:[{id:'es1',a:'EUNK',d:'2026-06-01',q:12,p:110,u:'2026-10-02'}]});},test(api){
  const P=api.pfDados(),f=id=>P.find(x=>x.id===id)||{};
  ok('portfolio lists the new ETFs',['EUNK','IS3N','EUNN'].every(x=>api.PF.some(p=>p[0]===x))&&[...$('#buyAsset').options].map(o=>o.value).join(',')==='AAPL,NVDA,GOOGL,SXR8,EUNK,IS3N,EUNN,BTC');
  ok('FIFO over two lots with a partial sale: 3.5 EUNK left, cost €367.50, realised €112.50',near(f('EUNK').q,3.5)&&near(f('EUNK').c,367.5)&&near(f('EUNK').real,112.5),JSON.stringify([f('EUNK').q,f('EUNK').c,f('EUNK').real]));
  ok('fractional holdings shown',txt('#pf-q-EUNK')==='3.5'&&txt('#pf-q-EUNN')==='3.25'&&txt('#pf-q-IS3N')==='20');
  ok('value at the latest euro price (no currency conversion)',near(f('IS3N').v,20*api.ATIVOS.find(a=>a.id==='IS3N').pts.slice(-1)[0][1],1e-12));
  ok('allocation shows each ETF',['ETF SXR8','ETF EUNK','ETF IS3N','ETF EUNN'].every(x=>txt('#ch-pf-alloc').includes(x)));
  const ex=txt('#ch-pf-exp');
  ok('look-through: each new ETF counted once, under its own index',/MSCI Europe \(EUNK\)/.test(ex)&&/MSCI EM IMI \(IS3N\)/.test(ex)&&/MSCI Japan IMI \(EUNN\)/.test(ex)&&/Rest of the S&P 500/.test(ex)&&ex.includes('€'+Math.round(f('EUNK').v).toLocaleString('en-GB')),ex);
  ok('look-through: the weights of each fund are named',/for SXR8 the iShares weights/.test(txt('#pfExpSub'))&&/for EUNK the iShares weights/.test(txt('#pfExpSub')),txt('#pfExpSub'));
  ok('no ETF in Dividends',[...document.querySelectorAll('#tbl-div tbody tr')].length===3&&!/EUNK|IS3N|EUNN/.test(txt('#tbl-div')));
  const X=api.anexoJ(2026),R=csvParse(X.csvAcoes),rows=X.acoes.filter(x=>x.id==='EUNK');
  ok('Anexo J: the EUNK sale gives two FIFO rows (10.5 + 1.5)',rows.length===2&&near(rows[0].q,10.5)&&near(rows[1].q,1.5)&&rows[0].aq===1050&&rows[0].vd===1155&&rows[1].aq===157.5&&rows[1].vd===165,JSON.stringify(rows.map(x=>[x.q,x.aq,x.vd])));
  ok('Anexo J: EUNK mapped as an Irish fund unit (G20, 372, listed, ISIN)',rows.every(x=>x.codigo==='G20'&&x.pais==='372'&&x.admitido==='Sim'&&x.isin==='IE00B4K48X80'&&x.status==='OK'&&!x.flags.length),JSON.stringify(rows[0]));
  ok('Anexo J: mapping for IS3N and EUNN in ANEXO_J',api.ANEXO_J.q92.ativos.IS3N.isin==='IE00BKM4GZ66'&&api.ANEXO_J.q92.ativos.EUNN.isin==='IE00B4L5YX21'&&['IS3N','EUNN'].every(x=>api.ANEXO_J.q92.ativos[x].codigo==='G20'&&api.ANEXO_J.q92.ativos[x].pais==='372'));
  ok('Anexo J CSV: euro-priced, no dollar columns for the ETF',csvCol(R,'Price basis').every(x=>x==='EUR as entered (priced in euros)')&&csvCol(R,'Acquisition FX EUR/USD').every(x=>x==='')&&csvCol(R,'Código').every(x=>x==='G20'),csvCol(R,'Price basis').join('|'));
  /* compra pelo formulário: preço do dia preenchido a partir do histórico do próprio ETF, quantidade fracionária */
  const sel=$('#buyAsset'),dt=$('#buyDate');sel.value='EUNN';sel.dispatchEvent(new Event('change'));$('#buyType').value='buy';dt.value='2026-03-02';dt.dispatchEvent(new Event('change'));
  const close=histPrice({historico:{EUNN:api.HIST.EUNN.map(p=>[new Date(p[0]).toISOString().slice(0,10),p[1]])}},'EUNN','2026-03-02');
  ok('form: the price fills in with the EUNN Xetra close of that day',+$('#buyPrice').value===+close.toFixed(2)&&/Xetra closing price/.test(txt('#buyMsg')),$('#buyPrice').value+' vs '+close+' '+txt('#buyMsg'));
  $('#buyQty').value='2.5';$('#buyForm').dispatchEvent(new Event('submit',{cancelable:true}));
  const nb=(JSON.parse(localStorage.getItem('bb.buys'))||[]).find(x=>x.a==='EUNN'&&x.q===2.5);
  ok('form: purchase saved in bb.buys with its split reference',!!nb&&Array.isArray(nb.r)&&nb.r[0]==='2026-03-02'&&/^\d{4}-\d{2}-\d{2}$/.test(nb.u),JSON.stringify(nb));
  sel.value='IS3N';sel.dispatchEvent(new Event('change'));$('#buyType').value='sell';$('#buyType').dispatchEvent(new Event('change'));dt.value='2026-07-01';dt.dispatchEvent(new Event('change'));$('#buyQty').value='5';$('#buyForm').dispatchEvent(new Event('submit',{cancelable:true}));
  ok('form: sale saved, holding reduced',(JSON.parse(localStorage.getItem('bb.sales'))||[]).some(x=>x.a==='IS3N'&&x.q===5)&&near(api.pfDados().find(x=>x.id==='IS3N').q,15));
  const b=nb&&document.querySelector(`#tbl-buys [data-del="${nb.id}"]`);if(b)b.click();
  ok('remove: the new purchase is deleted and remembered',!!nb&&!(JSON.parse(localStorage.getItem('bb.buys'))||[]).some(x=>x.id===nb.id)&&(nb.id in (JSON.parse(localStorage.getItem('bb.deleted'))||{})));}},

 etfbackup:{prep(d){d.backup={app:'Bluechip Board',version:4,saved:'2026-10-04T10:00:00.000Z',buys:[{id:'B1',a:'EUNK',d:'2026-01-05',q:2,p:100},{id:'B2',a:'IS3N',d:'2026-01-05',q:3.5,p:45},{id:'B3',a:'EUNN',d:'2026-01-05',q:1,p:70},{id:'B4',a:'SXR8',d:'2026-01-05',q:1,p:700}],
   lots:[],sales:[{id:'S1',a:'EUNN',d:'2026-04-01',q:0.5,p:75}],deleted:{}};},test(api){
  const b=JSON.parse(localStorage.getItem('bb.buys'))||[];
  ok('backup with the new ETFs loads into an empty browser',['B1','B2','B3','B4'].every(i=>b.some(x=>x.id===i))&&(JSON.parse(localStorage.getItem('bb.sales'))||[]).length===1,JSON.stringify(b.map(x=>x.id)));
  ok('entries from the file get their split reference (r, u)',b.filter(x=>['EUNK','IS3N','EUNN'].includes(x.a)).every(x=>/^\d{4}-\d{2}-\d{2}$/.test(x.u)),JSON.stringify(b));
  ok('holdings from the backup',near(api.pfDados().find(x=>x.id==='EUNN').q,0.5)&&near(api.pfDados().find(x=>x.id==='IS3N').q,3.5));
  const r=api.juntaBackup(V1,false),b2=JSON.parse(localStorage.getItem('bb.buys'))||[];
  ok('restoring an old version-1 backup keeps the ETF entries',b2.some(x=>x.id==='B1')&&b2.some(x=>x.id==='musjvneia0ov'&&x.a==='SXR8'),JSON.stringify(b2.map(x=>x.id)));
  ok('the backup the page writes keeps the new ETFs (same format)',['EUNK','IS3N','EUNN'].every(a=>b2.some(x=>x.a===a))&&Object.keys(b2[0]).every(k=>['id','a','d','q','p','r','u'].includes(k)));}},

 etfmissing:{prep(d){const a=d.ativos.find(x=>x.id==='EUNK');a.pontos=[];d.historico.EUNK=[];
   seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'e',a:'EUNK',d:'2026-01-05',q:1,p:100,u:'2026-10-02'},{id:'i',a:'IS3N',d:'2026-01-05',q:1,p:45,u:'2026-10-02'}]});},test(api){
  ok('no EUNK prices: the card says so',/No prices/.test(txt('.tk[data-tk="EUNK"]'))&&!!$('.tk[data-tk="IS3N"] .tk-price'));
  ok('no EUNK prices: alert, prices row "No data"',[...document.querySelectorAll('#alertas .alert-msg')].some(x=>/No prices for iShares Core MSCI Europe \(EUNK\)/.test(x.textContent))&&/ETF EUNKNo data/.test(txt('#tbl-ret tbody')));
  ok('no EUNK prices: portfolio shows "no price", never €0',txt('#pf-v-EUNK')==='no price'&&txt('#pf-q-EUNK')==='1'&&txt('#pf-v-IS3N')!=='no price');
  api.escolheEtf('EUNK');
  ok('no EUNK history: chart and simulator say so',/No long-term history/.test(txt('#ch-etf-hist'))&&/No long-term history/.test(txt('#sim-etf-EUNK-kpis')));}},

 etfstale:{prep(d){const i=d.ativos.find(x=>x.id==='IS3N');i.pontos=i.pontos.slice(0,-3);i.parcial=false;
   d.ativos.find(x=>x.id==='EUNN').fonte='previous run (data up to 2026-10-02)';
   const e=d.ativos.find(x=>x.id==='EUNK'),L=e.pontos.map(pair);G.prev=+L[L.length-2][1];L[L.length-1]=[L[L.length-1][0],G.prev*1.045];e.pontos=L;e.parcial=false;},test(api){
  const f=api.frescura('IS3N',api.ATIVOS.find(x=>x.id==='IS3N').pts);
  ok('IS3N behind its exchange (Xetra sessions): flagged',f&&f.velho&&f.falta>=2&&!!$('.tk[data-tk="IS3N"] .asof.old'),JSON.stringify(f));
  ok('IS3N: stale alert',[...document.querySelectorAll('#alertas .alert-msg')].some(x=>/The IS3N ETF: the latest price is from/.test(x.textContent)));
  ok('EUNN previous-run prices: alert and card label',[...document.querySelectorAll('#alertas .alert-msg')].some(x=>/The EUNN ETF: prices could not be downloaded/.test(x.textContent))&&/previous run/.test(txt('.tk[data-tk="EUNN"] .asof')));
  const s=api.stats(api.ATIVOS.find(x=>x.id==='EUNK').pts);
  ok('ETF daily move of 4.5% alerts (stock threshold 4%, not the Bitcoin 6%)',s.d1>4&&s.d1<6&&[...document.querySelectorAll('#alertas .alert-msg')].some(x=>/The EUNK ETF rose 4\.5%/.test(x.textContent)),s.d1);}},

 etfhold:{prep(d){d.etfs.EUNK=Object.assign({},d.etfs.EUNK,{aoVivo:false,top10:[],setores:[],posicoes:0,fonte:'Reference weights in the script (BlackRock, 2026-10-02)',data:'2026-10-02',dataIso:'2026-10-02'});
   seed({deleted:{},lots:[],sales:[],savedAt:'2026-10-03T10:00:00.000Z',buys:[{id:'e',a:'EUNK',d:'2026-01-05',q:1,p:100,u:'2026-10-02'},{id:'x',a:'SXR8',d:'2026-01-05',q:1,p:700,u:'2026-10-02'}]});},test(api){
  api.escolheEtf('EUNK');
  ok('EUNK holdings download failed: top 10 unavailable, nothing invented',/Unavailable: the iShares download failed/.test(txt('#top10Sub'))&&!$('#ch-top10 svg[role="img"]'));
  ok('look-through says EUNK uses reference weights, SXR8 the live ones',/for EUNK the reference weights stored in the script/.test(txt('#pfExpSub'))&&/for SXR8 the iShares weights/.test(txt('#pfExpSub')),txt('#pfExpSub'));
  api.escolheEtf('IS3N');
  ok('the other ETFs keep their own holdings',arr(api.etfInfo('IS3N').top10).length===10&&!/Unavailable/.test(txt('#top10Sub')));}},

 etfbad:{prep(d){d.etfs='garbage';},test(api){
  ok('malformed ETF data: ETFs still known from the price data, page works',api.ETF_IDS.join(',')==='SXR8,EUNK,IS3N,EUNN'&&$('#tbl-ret tbody').rows.length===9);
  api.escolheEtf('EUNK');
  ok('malformed ETF data: EUNK view shows unavailable holdings, no error',/Unavailable/.test(txt('#top10Sub'))&&txt('#h-etf')==='iShares Core MSCI Europe');}},

 /* ---------- remediação de 7 out 2026 ---------- */
 /* H4g: a pasta escolhida é confirmada logo (tem de ter o bluechip-board.html); uma pasta errada é recusada sem gravar nada */
 bkfolder:{wait:1000,prep(){seed({buys:[{id:'f1',a:'AAPL',d:'2026-09-01',q:1,p:200}],lots:[],sales:[],deleted:{},savedAt:'2026-10-03T10:00:00.000Z'});},test(api){
  const W=G.escrito=[];window.showSaveFilePicker=async()=>{throw new Error('the file picker should not be used');};
  /* IndexedDB em memória: o verdadeiro corre fora do tempo virtual do browser de teste (e não guarda objetos de teste) */
  const mem={};Object.defineProperty(window,'indexedDB',{configurable:true,value:{open:()=>{const q={};setTimeout(()=>{q.result={transaction:()=>{const tx={objectStore:()=>({get:k=>({result:mem[k]}),put:(v,k)=>{mem[k]=v;return{};}})};setTimeout(()=>tx.oncomplete&&tx.oncomplete(),0);return tx;},close(){}};if(q.onsuccess)q.onsuccess();},0);return q;}}});
  const ficheiro=n=>({kind:'file',name:n,queryPermission:async()=>'granted',requestPermission:async()=>'granted',createWritable:async()=>({write:async t=>{W.push(t);},close:async()=>{}})});
  const pasta=(n,tem)=>({kind:'directory',name:n,getFileHandle:async x=>{if(x==='bluechip-board.html'&&!tem)throw new DOMException('not found','NotFoundError');return ficheiro(x);}});
  window.showDirectoryPicker=async()=>pasta('Downloads',false);
  api.guardaFicheiro(true).then(()=>{G.msg1=txt('#bkMsg');G.n1=W.length;G.auto1=api.auto();
   window.showDirectoryPicker=async()=>pasta('BluechipBoard',true);return api.guardaFicheiro(true);}).then(()=>{G.msg2=txt('#bkMsg');G.auto2=api.auto();G.pend=api.pendente();}).catch(e=>{G.err=String(e);});},
  after(){ok('wrong folder (no bluechip-board.html in it): refused at once with a clear message, nothing saved',/has no bluechip-board\.html/.test(G.msg1||'')&&G.n1===0&&G.auto1===false,G.err||G.msg1);
   const b=G.escrito.length?JSON.parse(G.escrito[0]):null;
   ok('project folder: the backup is written there, automatic saving on, nothing left pending',G.escrito.length===1&&b&&b.app==='Bluechip Board'&&b.buys.length===1&&/Saved to bluechip-board-backup\.json/.test(G.msg2||'')&&G.auto2===true&&G.pend===false,JSON.stringify({m:G.msg2,a:G.auto2,p:G.pend,n:G.escrito.length,e:G.err,bk:txt('#bkMsg')}));}},
 /* H4a: depois de reabrir, a primeira alteração (um gesto do utilizador) pede logo a autorização do ficheiro ligado */
 bkperm:{wait:3000,prep(){seed({buys:[{id:'p1',a:'AAPL',d:'2026-09-01',q:1,p:200}],lots:[],sales:[],deleted:{},savedAt:'2026-10-03T10:00:00.000Z',fileSaved:'2026-10-03T10:00:00.000Z'});},test(api){
  let perm='prompt';const W=G.escrito=[];G.pedidos=0;
  api.setFH({kind:'file',name:'bluechip-board-backup.json',queryPermission:async()=>perm,requestPermission:async()=>{G.pedidos++;perm='granted';return'granted';},createWritable:async()=>({write:async t=>{W.push(t);},close:async()=>{}})});
  G.estado0=txt('#bkState');
  api.store.set('notes',{p1:{t:'why I bought it',at:new Date().toISOString()}});G.auto0=api.auto();
  setTimeout(()=>{G.auto1=api.auto();G.n1=W.length;api.store.set('notes',{p1:{t:'second change',at:new Date().toISOString()}});},300);
  setTimeout(()=>{G.n2=W.length;G.pend=api.pendente();G.estado2=txt('#bkState');},2500);},
  after(){ok('linked file and no pending change: no warning box before the first change',!/not in the backup file/.test(G.estado0),G.estado0);
   ok('first change after reopening: the browser is asked for permission once, then the change is saved to the file',G.pedidos===1&&G.auto1===true&&G.n1>=1&&JSON.parse(G.escrito[0]).notes.p1.t==='why I bought it',JSON.stringify({p:G.pedidos,a:G.auto1,n:G.n1}));
   ok('later changes are saved automatically, without asking again',G.pedidos===1&&G.n2>G.n1&&G.pend===false&&!/not in the backup file/.test(G.estado2),JSON.stringify({p:G.pedidos,n1:G.n1,n2:G.n2,pend:G.pend}));}},
 bkpermno:{wait:2500,prep(){seed({buys:[{id:'p1',a:'AAPL',d:'2026-09-01',q:1,p:200}],lots:[],sales:[],deleted:{},savedAt:'2026-10-03T10:00:00.000Z',fileSaved:'2026-10-03T10:00:00.000Z'});},test(api){
  G.pedidos=0;api.setFH({kind:'file',name:'bluechip-board-backup.json',queryPermission:async()=>'prompt',requestPermission:async()=>{G.pedidos++;return'denied';},createWritable:async()=>{throw new Error('no write without permission');}});
  api.store.set('notes',{p1:{t:'x',at:new Date().toISOString()}});setTimeout(()=>{api.store.set('notes',{p1:{t:'y',at:new Date().toISOString()}});},300);setTimeout(()=>{G.estado=txt('#bkState');G.auto=api.auto();},2000);},
  after(){ok('permission refused: asked only once on this page, automatic saving stays off and the box says what to do',G.pedidos===1&&G.auto===false&&/not in the backup file yet/.test(G.estado)&&/next change on this page|Save to project folder/.test(G.estado),JSON.stringify({p:G.pedidos})+' '+G.estado);}},
 /* backup malformado: só números verdadeiros (true, null, '' e [5] não passam a 1, 0, 0 e 5) */
 bkstrict:{test(api){const F=api.limpaBackup({app:'Bluechip Board',version:4,saved:'2026-10-01T00:00:00.000Z',deleted:{},sales:[],
   buys:[{id:'b1',a:'AAPL',d:'2025-01-02',q:[2],p:100},{id:'b2',a:'AAPL',d:'2025-01-02',q:true,p:100},{id:'b3',a:'AAPL',d:'2025-01-02',q:1,p:''},{id:'b4',a:'AAPL',d:'2025-01-02',q:'1.5',p:'200.25'}],
   lots:[{id:'l1',d:'2025-01-01',q:0.1,c:null},{id:'l2',d:'2025-01-01',q:0.1},{id:'l3',d:'2025-01-01',q:0.2,c:'3000'},{id:'l4',d:'2025-01-01',q:0.3,c:0}]});
  ok('malformed backup: quantities and prices that are not numbers are refused (arrays, booleans, empty text)',F.buys.map(x=>x.id).join(',')==='b4'&&F.buys[0].q===1.5&&F.buys[0].p===200.25,JSON.stringify(F.buys));
  ok('malformed backup: a Bitcoin purchase with a missing or null cost is refused (never a cost of €0); an explicit 0 is kept',F.lots.map(x=>x.id).join(',')==='l3,l4'&&F.lots[0].c===3000&&F.lots[1].c===0,JSON.stringify(F.lots));}},
 /* versão 1 sem ids: duas compras iguais no mesmo dia são duas, e juntar o mesmo ficheiro outra vez não duplica nada */
 v1dup:{prep(d){d.backup=V1DUP();},test(api){
  const b=get('buys')||[];
  ok('version-1 backup with two identical purchases (no ids): both kept, each with its own stable id',b.length===3&&new Set(b.map(x=>x.id)).size===3&&txt('#pf-q-SXR8')==='4',JSON.stringify(b.map(x=>x.id))+' '+txt('#pf-q-SXR8'));
  const r=api.juntaBackup(V1DUP(),false);
  ok('merging the same version-1 file again adds nothing',r.novos===0&&(get('buys')||[]).length===3,JSON.stringify(r));}},
 /* arredondamento aos cêntimos sobre o valor decimal (1.005 € → 1.01 €, não 1.00 €) */
 taxround:{prep(){seed({buys:[{id:'r1',a:'AAPL',d:'2025-01-02',q:1,p:10.005}],lots:[],sales:[{id:'r2',a:'AAPL',d:'2026-03-02',q:1,p:20.005}],deleted:{},savedAt:'2026-10-03T10:00:00.000Z'});},test(api){
  ok('rounding to cents: 1.005 → 1.01, 2.675 → 2.68, −1.005 → −1.01, 0.125 → 0.13 (half away from zero on the decimal value)',api.c2(1.005)===1.01&&api.c2(2.675)===2.68&&api.c2(-1.005)===-1.01&&api.c2(0.125)===0.13&&api.c2(10)===10&&api.c2(null)===null,[api.c2(1.005),api.c2(2.675),api.c2(-1.005),api.c2(0.125)].join(' '));
  const X=api.anexoJ(2026),r=X.acoes[0];
  ok('Anexo J: €10.005 and €20.005 entered give €10.01 and €20.01 (not €10.00 and €20.00), gain €10.00',r&&r.aq===10.01&&r.vd===20.01&&r.ganho===10,JSON.stringify(r&&[r.aq,r.vd,r.ganho]));}},
 /* preços atrasados: o alerta de movimento diz de quando é, e o "Latest daily move" da carteira não o soma como de hoje */
 stalemove:{prep(d){const a=d.ativos.find(x=>x.id==='NVDA'),P=a.pontos.map(pair).slice(0,-8);P[P.length-1]=[P[P.length-1][0],+(P[P.length-2][1]*1.07).toFixed(4)];a.pontos=P;a.parcial=false;a.fonte='Yahoo Finance';G.nvdaIso=P[P.length-1][0];
   seed({buys:[{id:'s1',a:'NVDA',d:'2026-01-05',q:2,p:150}],lots:[{id:'s2',d:'2026-01-05',q:0.01,c:600}],sales:[],deleted:{},savedAt:'2026-10-03T10:00:00.000Z'});},test(api){   /* Bitcoin (24/7) tem sempre o preço do dia */
  const al=api.alerts().filter(x=>x.co==='NVDA'&&/rose|fell/.test(x.txt));
  ok('stale prices: the daily-move alert says the date of that move ("not today") instead of "in the last session"',al.length===1&&al[0].txt.includes('(the latest price available, not today)')&&/the news of that day/.test(al[0].txt)&&!/today's news/.test(al[0].txt),al.map(x=>x.txt).join(' | '));
  const R=api.pfDados(),nv=R.find(x=>x.id==='NVDA'),sx=R.find(x=>x.id==='BTC');
  ok('portfolio: the latest daily move leaves out the asset whose price is not current, and names it',nv.velho===true&&sx.velho===false&&/without NVIDIA \(price not current\)/.test(txt('#pfKpis')),txt('#pfKpis'));}},
 staleall:{prep(d){const a=d.ativos.find(x=>x.id==='NVDA');a.pontos=a.pontos.map(pair).slice(0,-8);a.parcial=false;a.fonte='Yahoo Finance';
   seed({buys:[{id:'s1',a:'NVDA',d:'2026-01-05',q:2,p:150}],lots:[],sales:[],deleted:{},savedAt:'2026-10-03T10:00:00.000Z'});},test(){
  const k=[...document.querySelectorAll('#pfKpis .kpi')].find(x=>/Latest daily move/.test(x.textContent));
  ok('portfolio with no current price at all: the latest daily move is "—" (not €0)',k&&k.querySelector('.v').textContent.trim()==='—'&&/without NVIDIA/.test(k.textContent),k&&k.textContent);}},
 /* reação aos resultados: um preço intradiário (sessão aberta) não conta como fecho da sessão de reação */
 earnpartial:{prep(d){const a=d.ativos.find(x=>x.id==='AAPL'),H=d.historico.AAPL.map(pair),L=H[H.length-1][0],E=H[H.length-10][0];a.parcial=true;a.hora=d.geradoEm;
   const P=a.pontos.map(pair);if(P[P.length-1][0]!==L)P.push([L,H[H.length-1][1]]);a.pontos=P;G.L=L;G.E=E;
   d.resultadosSec={AAPL:{id:'AAPL',estado:'ok',fonte:'SEC EDGAR',obtidoEm:d.geradoEm,resultados:[{acc:'x1',entrega:L,aceite:L+'T12:00:00Z',horaNY:L+' 08:00',quando:'before',sessao:L,nota:''},{acc:'x2',entrega:E,aceite:E+'T12:00:00Z',horaNY:E+' 08:00',quando:'before',sessao:E,nota:''}],ultimoRelatorio:null,erro:'',nota:''}};},test(api){
  const E=api.reacoes('AAPL');
  ok('earnings reaction on a session still open (intraday price): not counted as a reaction; the earlier one is',E.ok&&E.C.length===1&&E.C[0].x.acc==='x2'&&E.sem===1,JSON.stringify({n:E.C.length,sem:E.sem}));}},
 /* um campo em falta (null) não é 0: a variação de 24 h da CoinGecko e um ponto de preço sem valor */
 btcnull:{prep(d){d.bitcoin=d.bitcoin||{};d.bitcoin.mercado=Object.assign({},d.bitcoin.mercado||{eur:60000},{var24:null});
   const a=d.ativos.find(x=>x.id==='AAPL'),P=a.pontos.map(pair);P[P.length-3]=[P[P.length-3][0],null];a.pontos=P;},test(api){
  const c=document.querySelector('.tk[data-tk="BTC"] .chg'),s=api.stats(api.ATIVOS.find(x=>x.id==='BTC').pts);
  ok('Bitcoin without the CoinGecko 24 h change: the card shows the change since 00:00 UTC, not +0.00%',c&&/since 00:00 UTC/.test(c.getAttribute('title')||'')&&(s.d1===0||!/^\+?0\.00%/.test(c.textContent.trim())),c&&(c.getAttribute('title')+' '+c.textContent));
  const aa=api.ATIVOS.find(x=>x.id==='AAPL');
  ok('a price point without a value is left out (never read as a price of 0)',aa.pts.every(p=>p[1]>0)&&!api.alerts().some(x=>x.co==='AAPL'&&/fell 100/.test(x.txt)),JSON.stringify(aa.pts.filter(p=>!(p[1]>0))));}},
 corrupt:{raw:'{"geradoEm": broken',after(){ok('unreadable data: the page says so instead of going blank',/could not be read/.test(txt('#main .callout')),txt('#main .callout'));}}
};
const sc=SC[S];
if(!sc)ok('unknown scenario '+S,false);
else if(sc.raw!=null)el.textContent=sc.raw;
else{let d=JSON.parse(el.textContent);if(sc.prep)sc.prep(d);el.textContent=JSON.stringify(d);}
window.__BB_TEST__=api=>{try{if(sc&&sc.test)sc.test(api);}catch(e){ok('exception: '+e.message,false,e.stack);}};
addEventListener('load',()=>setTimeout(()=>{try{if(sc&&sc.after)sc.after();}catch(e){ok('exception: '+e.message,false);}
 const p=document.createElement('pre');p.id='__res';p.textContent=JSON.stringify({s:S,errs:window.__errs,r:window.__R,out:window.__OUT||null});document.body.appendChild(p);},sc&&sc.wait||400));   /* wait: cenários com gravações assíncronas (backup) */
})();
