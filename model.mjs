// Entirely invented demo fixtures. No production workbook or report is imported.
export const months = Array.from({length:8},(_,i)=>`2026-${String(i+1).padStart(2,'0')}`);
export const rates = [1310,1335,1320,1360,1340,1385,1370,1400];
export const modules = [
 {id:'receivables',title:'채권관리',en:'RECEIVABLES',color:'#30343b',icon:'ledger',description:'거래처별 채권과 회수지연 위험을 확인합니다.',tabs:['통합 현황','회수지연 위험','회수 시뮬레이션']},
 {id:'fx',title:'외화관리',en:'FOREIGN EXCHANGE',color:'#2457d6',icon:'exchange',description:'외화잔액, 평가·실현손익과 환율 영향을 살펴봅니다.',tabs:['외화잔액','평가손익','환차손익','Net 손익','환율 시뮬레이션']},
 {id:'cip',title:'건설중인자산',en:'CAPITAL PROJECTS',color:'#c96826',icon:'building',description:'진행 중인 투자와 본자산 대체·상각을 검토합니다.',tabs:['투자 현황','대체 검토','감가상각 예측']},
 {id:'expenses',title:'비용증감',en:'EXPENSE REVIEW',color:'#147a83',icon:'chart',description:'전월 대비 비용 변동을 분석하고 검토 의견을 정리합니다.',tabs:['비용 비교','검토 보고서']}
];
export function customers(m=7){return ['가상 오로라','가상 블루리프','가상 코스모','가상 델타웍스','가상 에버필드','가상 포레스트'].map((name,i)=>{const balance=Number((8+i*2.1+m*.43).toFixed(2)),overdue=Number((balance*[.06,.18,.42,.09,.32,.22][i]).toFixed(2));return {id:`DEMO-C${i+1}`,name,balance,overdue,days:[8,27,112,15,78,45][i]+m,over90:i===2?Number((overdue*.72).toFixed(2)):0};});}
export function fxRows(m=7){return [{name:'외화보통예금',type:'자산',base:4.2,growth:.18},{name:'매출채권',type:'자산',base:3.8,growth:.12},{name:'미수금',type:'자산',base:.8,growth:.03},{name:'단기차입금',type:'부채',base:-5,growth:.1},{name:'매입채무',type:'부채',base:-1.4,growth:-.02}].map((r,i)=>{const usd=r.base+r.growth*m,prev=m?rates[m-1]:1300,valuation=usd*(rates[m]-1300)/100,realized=(i%2?-.07:.09)*(m+1)+(rates[m]-prev)*usd*.04/100;return {...r,usd,krw:usd*rates[m]/100,valuation,realized,net:valuation+realized};});}
export function projects(m=7){return ['가상 생산라인 A','가상 물류센터 B','가상 연구시설 C','가상 설비개선 D','가상 에너지설비 E'].map((name,i)=>({id:`DEMO-P${i+1}`,name,amount:6+i*3.4+m*.35,days:[12,48,125,210,31][i],life:[10,20,8,5,12][i],status:['진행','준공 검토','대체 검토','장기 미발생','진행'][i]}));}
export function expenses(m=7){return ['인건비','운송비','지급수수료','수선비','감가상각비'].map((name,i)=>{const previous=3+i*1.3+m*.08,current=previous+[.36,-.18,.22,1.15,.12][i];return {name,previous,current,difference:current-previous};});}
export const sum=(rows,key)=>rows.reduce((s,r)=>s+r[key],0);
export const fxImpact=(m,rate,volume=100)=>sum(fxRows(m),'usd')*volume/100*(rate-rates[m])/100;
export const recovery=(m,percent)=>sum(customers(m),'overdue')*percent/100;
