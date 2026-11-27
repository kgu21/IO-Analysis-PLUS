const $=id=>document.getElementById(id);
const fmt=x=>Number(x).toLocaleString('ko-KR',{maximumFractionDigits:3});
const fmt4=x=>Number(x).toLocaleString('ko-KR',{minimumFractionDigits:4,maximumFractionDigits:4});
const DATASETS={
 large:{file:'io-2023-large.xlsx',label:'대분류'},
 medium:{file:'io-2023-medium.xlsx',label:'중분류'},
 small:{file:'io-2023-small.xlsx',label:'소분류'},
 basic:{file:'io-2023-basic.xlsx',label:'기본부문'}
};
const LOCKABLE=['sheet','row','sector','add','run','row-search','col-search','sector-search','amount','kind-filter','struct-search','struct-sort','struct-download','demand-item','demand-sort','demand-png','demand-csv','demand-detail-csv'];
let worker, nextId=0, requests=new Map(), sectors=[], rowChoices=[], cells=[], basket=[], result=null, structRows=[], currentLabel='';
function text(tag,value){const e=document.createElement(tag);e.textContent=value;return e;}
function call(type,payload){return new Promise((resolve,reject)=>{const id=++nextId;requests.set(id,{resolve,reject});worker.postMessage({id,type,payload});});}
function startWorker(){worker=new Worker('./worker.js');worker.onmessage=({data:d})=>{if(d.type==='status'){$('status').textContent=d.text;return;}const p=requests.get(d.id);if(!p)return;requests.delete(d.id);d.ok?p.resolve(d.result):p.reject(Error(d.error));};worker.onerror=e=>{for(const p of requests.values())p.reject(Error(e.message||'Worker 실행 오류'));requests.clear();};}
function options(el,list){el.replaceChildren(...list.map(x=>{const o=document.createElement('option');o.value=x.code;o.textContent=x.code+' '+x.name;return o;}));}
function invalidate(){result=null;$('results').hidden=true;}
function lock(v){for(const id of LOCKABLE)$(id).disabled=v;}

async function load(buffer,label){
 invalidate();basket=[];renderBasket();lock(true);
 for(const p of requests.values())p.reject(Error('데이터 다시 로드'));requests.clear();
 if(!worker)startWorker();   // Python 환경은 한 번만 준비하고 재사용
 try{
  const info=await call('init',{buffer,label});
  currentLabel=label;sectors=info.sectors;options($('sector'),sectors);
  $('sector-search').value='';$('row-search').value='';$('col-search').value='';$('struct-search').value='';$('kind-filter').value='';
  $('sheet').replaceChildren(...info.sheets.map(s=>{const o=text('option',s);o.value=s;return o;}));
  $('validation').textContent=label+' · '+sectors.length+'개 부문\n'+JSON.stringify(info.residuals,null,2)+'\n최대 절대 잔차: 허용값 1e-8 미만';
  $('subtitle').textContent='2023 · 생산자가격 · '+label;
  $('lead').textContent=info.sheets.length+'개 원자료 시트 탐색, '+sectors.length+'개 부문 수요충격 분석, 수요 항목별 분석, 전·후방 연쇄효과 분석';
  $('level-now').textContent=label+' ('+sectors.length+'개 부문)';
  $('download').textContent='전체 '+sectors.length+'개 산업 CSV 다운로드';
  lock(false);
  await explore();
  await loadStructure();
  await loadDemand();
  $('status').textContent='준비 완료 · '+label+' · '+sectors.length+'개 부문 · '+info.sheets.length+'개 시트 · 계수 검증 통과';
 }catch(e){$('status').textContent='준비 실패: '+e.message+' — 새로고침 후 재시도하세요.';throw e;}
}

let exploreSeq=0;
async function explore(){const seq=++exploreSeq;try{const d=await call('explore',{sheet:$('sheet').value,row:$('row').value});if(seq!==exploreSeq)return;rowChoices=d.rows;filterRows(d.selected);cells=d.cells;$('unit').textContent=d.unit;$('explore-caption').textContent=currentLabel+' · '+$('sheet').value+' / 행: '+d.selected+' '+d.name;renderCells();}catch(e){$('status').textContent=e.message;}}
function filterRows(selected){const q=$('row-search').value.trim();options($('row'),rowChoices.filter(r=>(r.code+' '+r.name).includes(q)));if([...$('row').options].some(o=>o.value===selected))$('row').value=selected;}
function renderCells(){const q=$('col-search').value.trim();$('cells').replaceChildren(...cells.filter(c=>(c.code+' '+c.name).includes(q)).map(c=>{const tr=document.createElement('tr');for(const v of [c.code,c.name,c.value===null?'빈 셀':typeof c.value==='number'?fmt(c.value):c.value])tr.append(text('td',v));return tr;}));}
$('sheet').onchange=()=>{$('row-search').value='';$('row').replaceChildren();explore();};$('row').onchange=explore;
$('row-search').oninput=()=>{filterRows();if($('row').value)explore();else{$('cells').replaceChildren();$('explore-caption').textContent='일치하는 행이 없습니다.';}};
$('col-search').oninput=renderCells;
$('sector-search').oninput=()=>{const q=$('sector-search').value.trim();options($('sector'),sectors.filter(s=>(s.code+' '+s.name).includes(q)));};

function renderBasket(){$('basket').replaceChildren(...basket.map((s,i)=>{const li=text('li',s.code+' '+s.name+' / '+fmt(s.amount)+'백만원 ');const b=text('button','삭제');b.type='button';b.onclick=()=>{basket.splice(i,1);invalidate();renderBasket();};li.append(b);return li;}));$('sim-status').textContent=basket.length?basket.length+'개 입력 준비. 계산 버튼을 누르세요.':'시나리오를 추가하세요.';}
$('scenario').onsubmit=e=>{e.preventDefault();const amount=Number($('amount').value),s=sectors.find(s=>s.code===$('sector').value);if(!s||!Number.isFinite(amount)||amount<=0){$('sim-status').textContent='산업과 양의 증가액을 확인하세요.';return;}basket.push({...s,amount});invalidate();renderBasket();};
$('clear').onclick=()=>{basket=[];invalidate();renderBasket();};
$('run').onclick=async()=>{if(!basket.length){$('sim-status').textContent='시나리오를 먼저 추가하세요.';return;}const snapshot=basket.map(s=>({...s}));$('run').disabled=true;$('sim-status').textContent='계산 중…';try{const r=await call('simulate',snapshot);if(JSON.stringify(snapshot)!==JSON.stringify(basket)){$('sim-status').textContent='입력이 변경되었습니다. 다시 계산하세요.';return;}result=r;renderResult();$('sim-status').textContent='계산 완료. 아래 결과는 현재 시나리오('+currentLabel+')에 대응합니다.';}catch(e){invalidate();$('sim-status').textContent='계산 실패: '+e.message;}finally{$('run').disabled=false;}};
function renderResult(){$('results').hidden=false;$('cards').replaceChildren(...[['수요 증가액',result.demand],['생산유발액',result.production],['수입유발액',result.imports],['부가가치유발액',result.valueAdded]].map(([title,v])=>{const d=text('div',title);d.className='card';d.append(text('strong',fmt(v)),text('span','백만원'));return d;}),(()=>{const d=text('div','생산유발승수');d.className='card';d.append(text('strong',fmt4(result.multiplier)),text('span','배 (생산유발액 ÷ 수요 증가액)'));return d;})());const max=Math.max(...result.rows.map(r=>r.production),1);$('bars').replaceChildren(...result.rows.slice(0,15).map(r=>{const d=text('div',r.name+' · '+fmt(r.production)+'백만원');d.className='bar';const t=document.createElement('div');t.className='track';const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,r.production/max*100)+'%';t.append(f);d.append(t);return d;}));$('result-rows').replaceChildren(...result.rows.map(r=>{const tr=document.createElement('tr');[r.code,r.name,fmt(r.production),fmt(r.imports),fmt(r.valueAdded)].forEach(v=>tr.append(text('td',v)));return tr;}));}
function csvCell(v){let s=String(v);if(/^[=+@-]/.test(s)&&typeof v!=='number')s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
function saveCsv(rows,name){const url=URL.createObjectURL(new Blob(['\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}));const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
$('download').onclick=()=>{if(!result)return;saveCsv([['산업코드','산업명','생산유발액_백만원','수입유발액_백만원','부가가치유발액_백만원'],...result.rows.map(r=>[r.code,r.name,r.production,r.imports,r.valueAdded])],'io-2023-'+currentLabel+'-results.csv');};

// ---- 3. 산업 구조 분석 ----
async function loadStructure(){const d=await call('analysis',{});structRows=d.rows;renderStructure();}
function visibleStructure(){const q=$('struct-search').value.trim(),k=$('kind-filter').value,key=$('struct-sort').value;return structRows.filter(r=>(!k||r.kind===k)&&(r.code+' '+r.name).includes(q)).sort((a,b)=>b[key]-a[key]);}
function renderStructure(){const list=visibleStructure();$('struct-rows').replaceChildren(...list.map(r=>{const tr=document.createElement('tr');[r.code,r.name,fmt4(r.production),fmt4(r.imports),fmt4(r.valueAdded),fmt4(r.influence),fmt4(r.sensitivity),r.kind].forEach(v=>tr.append(text('td',v)));return tr;}));$('struct-status').textContent=currentLabel+' · '+list.length+'/'+structRows.length+'개 산업 표시 · 영향력·감응도 1 이상이 평균 이상';}
$('kind-filter').onchange=renderStructure;$('struct-search').oninput=renderStructure;$('struct-sort').onchange=renderStructure;
$('struct-download').onclick=()=>{const list=visibleStructure();if(!list.length)return;saveCsv([['산업코드','산업명','생산유발계수','수입유발계수','부가가치유발계수','영향력계수','감응도계수','유형'],...list.map(r=>[r.code,r.name,r.production,r.imports,r.valueAdded,r.influence,r.sensitivity,r.kind])],'io-2023-'+currentLabel+'-structure.csv');};


// ---- 3. 수요 항목별 분석 ----
const FONT="'Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif",COL_VA='#1b7f79',COL_IM='#d9822b';
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const shortName=n=>String(n).replace(/\s*\(.*$/,'').trim();
const pct=x=>(x*100).toFixed(1)+'%';
const SORT_LABEL={imports:'수입유발액',production:'생산유발액',valueAdded:'부가가치유발액'};
let demandData=null, demandSvgText='', demandSvgSize=[0,0];
async function loadDemand(){
 demandData=await call('demand',{});
 $('demand-item').replaceChildren(...demandData.items.map(i=>{const o=text('option',shortName(i.name));o.value=i.code;return o;}));
 if(demandData.items.some(i=>i.code==='9140'))$('demand-item').value='9140';
 renderDemand();
}
function demandSvg(){
 const items=demandData.items.filter(i=>i.coefValueAdded!==null&&i.coefImportIntermediate!==null);
 const W=780,left=190,right=30,top=84,rowH=38,H=top+items.length*rowH+34,bw=W-left-right;
 const clamp=v=>Math.max(0,Math.min(1,v));
 let s='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 '+W+' '+H+'" width="'+W+'" height="'+H+'" role="img">';
 s+='<rect width="'+W+'" height="'+H+'" fill="#ffffff"/>';
 s+='<text x="16" y="26" font-size="16" font-weight="700" fill="#1c2b36" font-family="'+FONT+'">'+esc('국산 최종수요 1원당 부가가치와 수입 중간재 ('+currentLabel+' 기준)')+'</text>';
 s+='<text x="16" y="46" font-size="12" fill="#5a6b78" font-family="'+FONT+'">부가가치 + 수입 중간재 = 100% · 자료: 한국은행 2023년 산업연관표(연장표)</text>';
 s+='<rect x="16" y="56" width="12" height="12" fill="'+COL_VA+'"/><text x="34" y="66" font-size="12" fill="#1c2b36" font-family="'+FONT+'">국내 부가가치로 남는 몫</text>';
 s+='<rect x="190" y="56" width="12" height="12" fill="'+COL_IM+'"/><text x="208" y="66" font-size="12" fill="#1c2b36" font-family="'+FONT+'">해외 수입 중간재로 새는 몫</text>';
 items.forEach((it,k)=>{
  const y=top+k*rowH,va=clamp(it.coefValueAdded),im=clamp(it.coefImportIntermediate),vw=va*bw,iw=im*bw;
  s+='<text x="'+(left-10)+'" y="'+(y+18)+'" text-anchor="end" font-size="13" fill="#1c2b36" font-family="'+FONT+'"'+(it.code==='9190'?' font-weight="700"':'')+'>'+esc(shortName(it.name))+'</text>';
  s+='<rect x="'+left+'" y="'+y+'" width="'+vw.toFixed(1)+'" height="26" fill="'+COL_VA+'"/>';
  s+='<rect x="'+(left+vw).toFixed(1)+'" y="'+y+'" width="'+iw.toFixed(1)+'" height="26" fill="'+COL_IM+'"/>';
  if(vw>48)s+='<text x="'+(left+vw/2).toFixed(1)+'" y="'+(y+17)+'" text-anchor="middle" font-size="12" fill="#ffffff" font-family="'+FONT+'">'+pct(va)+'</text>';
  if(iw>48)s+='<text x="'+(left+vw+iw/2).toFixed(1)+'" y="'+(y+17)+'" text-anchor="middle" font-size="12" fill="#ffffff" font-family="'+FONT+'">'+pct(im)+'</text>';
 });
 s+='</svg>';
 return [s,W,H];
}
function renderDemand(){
 if(!demandData)return;
 $('demand-rows').replaceChildren(...demandData.items.map(i=>{const tr=document.createElement('tr');
  const c4=v=>v===null?'-':fmt4(v);
  [shortName(i.name),fmt(i.domestic),fmt(i.imported),fmt(i.production),fmt(i.imports),fmt(i.valueAdded),c4(i.coefProduction),c4(i.coefImportIntermediate),c4(i.coefValueAdded),c4(i.coefImportPublished)].forEach(v=>tr.append(text('td',v)));return tr;}));
 const [svg,W,H]=demandSvg();demandSvgText=svg;demandSvgSize=[W,H];
 $('demand-chart').innerHTML=svg.replace('width="'+W+'" height="'+H+'"','style="width:100%;max-width:'+W+'px;height:auto"');
 renderDemandDetail();
}
function demandDetailRows(){
 const code=$('demand-item').value,key=$('demand-sort').value||'imports';
 return [...(demandData.detail[code]||[])].sort((a,b)=>b[key]-a[key]);
}
function renderDemandDetail(){
 if(!demandData)return;
 const code=$('demand-item').value,key=$('demand-sort').value||'imports',item=demandData.items.find(i=>i.code===code);
 if(!item)return;
 const rows=demandDetailRows(),nm=shortName(item.name);
 $('demand-detail-title').textContent=nm+' 최종수요가 유발하는 '+SORT_LABEL[key]+' 상위 15개 산업 ('+currentLabel+')';
 const max=Math.max(...rows.slice(0,15).map(r=>r[key]),1);
 $('demand-bars').replaceChildren(...rows.slice(0,15).map(r=>{const d=text('div',r.name+' · '+fmt(r[key])+'백만원');d.className='bar';const t=document.createElement('div');t.className='track';const f=document.createElement('div');f.className='fill';f.style.width=Math.max(0,r[key]/max*100)+'%';t.append(f);d.append(t);return d;}));
 $('demand-detail-caption').textContent=nm+' · 산업별 유발액 — 백만원, '+SORT_LABEL[key]+' 내림차순 (수입유발액은 해당 산업 생산물(상품)의 수입 증가분)';
 $('demand-detail-rows').replaceChildren(...rows.map(r=>{const tr=document.createElement('tr');[r.code,r.name,fmt(r.production),fmt(r.imports),fmt(r.valueAdded)].forEach(v=>tr.append(text('td',v)));return tr;}));
}
$('demand-item').onchange=renderDemandDetail;$('demand-sort').onchange=renderDemandDetail;
$('demand-csv').onclick=()=>{if(!demandData)return;saveCsv([['항목','국산최종수요_백만원','수입최종수요_백만원','생산유발액_백만원','수입유발액_백만원','부가가치유발액_백만원','생산유발계수','수입중간재유발계수','부가가치유발계수','공표수입유발계수'],...demandData.items.map(i=>[shortName(i.name),i.domestic,i.imported,i.production,i.imports,i.valueAdded,i.coefProduction,i.coefImportIntermediate,i.coefValueAdded,i.coefImportPublished])],'io-2023-'+currentLabel+'-demand-items.csv');};
$('demand-detail-csv').onclick=()=>{if(!demandData)return;const code=$('demand-item').value,item=demandData.items.find(i=>i.code===code);if(!item)return;saveCsv([['산업코드','산업명','생산유발액_백만원','수입유발액_백만원','부가가치유발액_백만원'],...demandDetailRows().map(r=>[r.code,r.name,r.production,r.imports,r.valueAdded])],'io-2023-'+currentLabel+'-'+shortName(item.name).replaceAll(' ','')+'-detail.csv');};
function savePng(svgText,w,h,name){
 const img=new Image(),url=URL.createObjectURL(new Blob([svgText],{type:'image/svg+xml;charset=utf-8'}));
 img.onload=()=>{const c=document.createElement('canvas');c.width=w*2;c.height=h*2;const x=c.getContext('2d');x.fillStyle='#ffffff';x.fillRect(0,0,c.width,c.height);x.drawImage(img,0,0,c.width,c.height);URL.revokeObjectURL(url);
  c.toBlob(b=>{const u=URL.createObjectURL(b),a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000);},'image/png');};
 img.onerror=()=>{URL.revokeObjectURL(url);$('status').textContent='차트 이미지 저장에 실패했습니다.';};
 img.src=url;
}
$('demand-png').onclick=()=>{if(demandSvgText)savePng(demandSvgText,demandSvgSize[0],demandSvgSize[1],'io-2023-'+currentLabel+'-demand-chart.png');};


// ---- 4. 분류 수준별 비교 ----
const LEVELS=['large','medium','small','basic'],LEVEL_COLOR={large:'#9fb7c9',medium:'#5c8fb0',small:'#2f6a8f',basic:'#123c52'};
const METRIC_LABEL={coefImportIntermediate:'수입(중간재)유발계수',coefValueAdded:'부가가치유발계수',coefProduction:'생산유발계수'};
let cmp=null, cmpIdx={}, cmpHier={}, cmpSvgText='', cmpSvgSize=[0,0];
async function loadComparison(){
 try{
  const r=await fetch('./data/comparison.json');
  if(!r.ok)throw Error('HTTP '+r.status);
  cmp=await r.json();
  for(const k of LEVELS)cmpIdx[k]=Object.fromEntries(cmp.levels[k].sectors.map(s=>[s.code,s]));
  cmpHier=Object.fromEntries(cmp.hierarchy.map(h=>[h.basic,h]));
  for(const id of ['cmp-metric','cmp-png','cmp-csv','cmp-search','cmp-sector','cmp-chain-csv'])$(id).disabled=false;
  $('cmp-search').value='';cmpFilterSectors();
  $('cmp-status').textContent='분류 비교 자료 준비 완료 · '+LEVELS.map(k=>cmp.levels[k].label+' '+cmp.levels[k].sectors.length+'개').join(' · ');
  renderCompare();renderChain();
 }catch(e){$('cmp-status').textContent='분류 비교 자료를 불러오지 못했습니다: '+e.message;}
}
function cmpItemRows(){
 const m=$('cmp-metric').value||'coefImportIntermediate';
 return cmp.levels.basic.items.map(i=>{const vals=LEVELS.map(k=>{const it=cmp.levels[k].items.find(x=>x.code===i.code);return it?it[m]:null;});return {code:i.code,name:shortName(i.name),vals};}).filter(r=>r.vals.every(v=>v!==null));
}
function cmpSvg(rows,m){
 const W=780,left=170,right=60,top=96,gh=82,H=top+rows.length*gh+20,bw=W-left-right;
 const maxv=m==='coefProduction'?2.5:1,ticks=[0,.25,.5,.75,1].map(t=>t*maxv);
 let s='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 '+W+' '+H+'" width="'+W+'" height="'+H+'" role="img">';
 s+='<rect width="'+W+'" height="'+H+'" fill="#ffffff"/>';
 s+='<text x="16" y="26" font-size="16" font-weight="700" fill="#1c2b36" font-family="'+FONT+'">'+esc('분류 수준별 '+METRIC_LABEL[m]+' (국산 최종수요 1원당)')+'</text>';
 s+='<text x="16" y="46" font-size="12" fill="#5a6b78" font-family="'+FONT+'">자료: 한국은행 2023년 산업연관표(연장표), 생산자가격</text>';
 LEVELS.forEach((k,i)=>{const x=16+i*110;s+='<rect x="'+x+'" y="58" width="12" height="12" fill="'+LEVEL_COLOR[k]+'"/><text x="'+(x+18)+'" y="68" font-size="12" fill="#1c2b36" font-family="'+FONT+'">'+esc(cmp.levels[k].label)+'</text>';});
 ticks.forEach(t=>{const x=left+t/maxv*bw;s+='<line x1="'+x.toFixed(1)+'" y1="'+(top-8)+'" x2="'+x.toFixed(1)+'" y2="'+(H-20)+'" stroke="#e1e8ee"/><text x="'+x.toFixed(1)+'" y="'+(top-14)+'" text-anchor="middle" font-size="11" fill="#5a6b78" font-family="'+FONT+'">'+t.toFixed(2)+'</text>';});
 rows.forEach((r,gi)=>{const y=top+gi*gh;
  s+='<text x="'+(left-10)+'" y="'+(y+36)+'" text-anchor="end" font-size="13" fill="#1c2b36" font-family="'+FONT+'"'+(r.code==='9190'?' font-weight="700"':'')+'>'+esc(r.name)+'</text>';
  r.vals.forEach((v,i)=>{const w=Math.max(0,Math.min(v/maxv,1))*bw,by=y+i*17;
   s+='<rect x="'+left+'" y="'+by+'" width="'+w.toFixed(1)+'" height="14" fill="'+LEVEL_COLOR[LEVELS[i]]+'"/><text x="'+(left+w+5).toFixed(1)+'" y="'+(by+11)+'" font-size="11" fill="#1c2b36" font-family="'+FONT+'">'+fmt4(v)+'</text>';});
 });
 s+='</svg>';
 return [s,W,H];
}
function renderCompare(){
 if(!cmp)return;
 const m=$('cmp-metric').value||'coefImportIntermediate',rows=cmpItemRows();
 $('cmp-caption').textContent='분류별 '+METRIC_LABEL[m]+' (국산 최종수요 1원당 원)';
 $('cmp-rows').replaceChildren(...rows.map(r=>{const tr=document.createElement('tr');
  [r.name,...r.vals.map(fmt4),(r.vals[3]-r.vals[0]>=0?'+':'')+fmt4(r.vals[3]-r.vals[0])].forEach(v=>tr.append(text('td',v)));return tr;}));
 const [svg,W,H]=cmpSvg(rows,m);cmpSvgText=svg;cmpSvgSize=[W,H];
 $('cmp-chart').innerHTML=svg.replace('width="'+W+'" height="'+H+'"','style="width:100%;max-width:'+W+'px;height:auto"');
}
function cmpFilterSectors(){
 const q=$('cmp-search').value.trim(),prev=$('cmp-sector').value;
 options($('cmp-sector'),cmp.levels.basic.sectors.filter(s=>(s.code+' '+s.name).includes(q)));
 if(prev&&[...$('cmp-sector').options].some(o=>o.value===prev))$('cmp-sector').value=prev;
}
function chainRows(){
 const code=$('cmp-sector').value,h=cmpHier[code];if(!h)return [];
 return LEVELS.map(k=>{const s=cmpIdx[k][h[k]];return {level:cmp.levels[k].label,...s};});
}
function renderChain(){
 if(!cmp)return;
 const rows=chainRows();
 if(!rows.length){$('cmp-chain-rows').replaceChildren();$('cmp-chain-caption').textContent='일치하는 산업이 없습니다.';return;}
 const last=rows[rows.length-1];
 $('cmp-chain-caption').textContent=last.code+' '+last.name+' — 대분류에서 기본부문까지 (국산 최종수요 1원당)';
 $('cmp-chain-rows').replaceChildren(...rows.map(r=>{const tr=document.createElement('tr');
  [r.level,r.code,r.name,fmt4(r.production),fmt4(r.imports),fmt4(r.valueAdded),fmt4(r.influence),fmt4(r.sensitivity)].forEach(v=>tr.append(text('td',v)));return tr;}));
}
$('cmp-metric').onchange=renderCompare;
$('cmp-search').oninput=()=>{cmpFilterSectors();renderChain();};
$('cmp-sector').onchange=renderChain;
$('cmp-png').onclick=()=>{if(cmpSvgText)savePng(cmpSvgText,cmpSvgSize[0],cmpSvgSize[1],'io-2023-compare-'+($('cmp-metric').value||'metric')+'.png');};
$('cmp-csv').onclick=()=>{if(!cmp)return;const m=$('cmp-metric').value||'coefImportIntermediate';saveCsv([['항목',...LEVELS.map(k=>cmp.levels[k].label+'_'+METRIC_LABEL[m])],...cmpItemRows().map(r=>[r.name,...r.vals])],'io-2023-compare-items-'+m+'.csv');};
$('cmp-chain-csv').onclick=()=>{if(!cmp)return;const rows=chainRows();if(!rows.length)return;saveCsv([['분류','코드','산업명','생산유발계수','수입유발계수','부가가치유발계수','영향력계수','감응도계수'],...rows.map(r=>[r.level,r.code,r.name,r.production,r.imports,r.valueAdded,r.influence,r.sensitivity])],'io-2023-compare-'+$('cmp-sector').value+'.csv');};
loadComparison();

// ---- 분류별 엑셀 불러오기 ----
async function loadSelected(){
 const ds=DATASETS[$('level').value];
 $('status').textContent=ds.label+' Excel 내려받는 중…';
 const response=await fetch('./data/'+ds.file);
 if(!response.ok)throw Error(ds.label+' Excel 다운로드 실패 (HTTP '+response.status+')');
 await load(await response.arrayBuffer(),ds.label);
}
$('load-data').onclick=async()=>{
 const button=$('load-data');
 button.disabled=true;$('level').disabled=true;
 button.textContent='불러오는 중…';
 try{await loadSelected();button.textContent='다시 불러오기';}
 catch(error){$('status').textContent='준비 실패: '+error.message+' — 연결 상태를 확인하고 새로고침하세요.';button.textContent='불러오기';}
 finally{button.disabled=false;$('level').disabled=false;}
};
