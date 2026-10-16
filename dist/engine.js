/* XLSX package editing: preserve every untargeted cell as its original XML. */
const EM = (()=>{
const MAP={'静设备':'JSB','屏蔽泵':'D-PBB','反应设备':'J-REA','离心泵':'D-LXB','列管式换热器':'J-LGH','贮槽/罐':'J-POT','板式换热器':'J-BSH','离心式风机':'D-LXF','其他静设备':'J-QTJ','自吸泵':'D-ZXB','搅拌机':'D-JBF'};
const DATE_FIELDS=['出厂日期','启用日期','设备制造日期'];
// Calendar-year distance, inclusive at the boundary; shared by all date fields.
const MAX_DATE_AGE_YEARS=100;
const FIELDS=['设备小类名称','设备小类编码',...DATE_FIELDS,'设备代码'];
const ASSET_FIELD='资产ABC';
// Match a self-closing cell/row by itself, without consuming its next sibling.
const CELL_XML=/<(?:\w+:)?c\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?c>)/g;
const ROW_XML=/<(?:\w+:)?row\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?row>)/g;
const parse=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw Error('工作簿 XML 无法解析');return d};
const nodes=(d,n)=>Array.from(d.getElementsByTagNameNS('*',n));
const first=(d,n)=>nodes(d,n)[0];
const txt=(d,n)=>first(d,n)?.textContent||'';
const esc=s=>s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
function dateYear(now=new Date()){
 return Number(new Intl.DateTimeFormat('en',{timeZone:'Asia/Shanghai',year:'numeric'}).format(now));
}
function calendarError(y,m,d){
 if(y!==null&&(!Number.isInteger(y)||y<1||y>9999))return'日期年份必须为 1 至 9999 的整数';
 if(m<1||m>12)return`日期不存在：月份 ${m} 不在 1 至 12 范围内`;
 const leap=y===null||y%4===0&&(y%100!==0||y%400===0);
 const last=[31,leap?29:28,31,30,31,30,31,31,30,31,30,31][m-1];
 if(d<1||d>last)return`日期不存在：${m} 月的日期须为 1 至 ${last} 日，实际为 ${d} 日`;
 return'';
}
function date(raw,native=false,epoch1904=false,currentYear=dateYear(),formattedDate=false){
 const block=reason=>({block:true,reason});
 const s=raw.trim();if(!s)return{value:''};
 if(!formattedDate&&/^\d{4}(?:年)?$/.test(s))return block('仅填写年份，无法确定完整日期');
 let y,m,d;
 const calendar=s.match(/^(\d{4})[-./年](\d{1,2})(?:[-./月](\d{1,2})日?|月)?$/);
 const numericSerial=native&&/^[+-]?\d+(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(s);
 // Keep existing YYYY.MM month handling, including numeric cells such as 2026.03.
 if(calendar){
  [y,m,d]=[+calendar[1],+calendar[2],+(calendar[3]??1)];
 }else if(numericSerial){
  const v=Number(s),serial=Math.floor(v);
  if(!Number.isFinite(v)||v<0||serial>(epoch1904?2957003:2958465))return block('Excel 日期序列值超出有效范围');
  if(!epoch1904&&serial===60)return block('日期不存在：Excel 序列值 60 对应 1900-02-29，1900 年不是闰年');
  const dt=new Date(Date.UTC(epoch1904?1904:1899,epoch1904?0:11,epoch1904?1:31)+(serial-(epoch1904?0:serial>=60?1:0))*86400000);
  y=dt.getUTCFullYear();m=dt.getUTCMonth()+1;d=dt.getUTCDate();
 }else{
  // Also catch impossible month/day examples without inventing a missing year.
  const monthDay=s.match(/^(\d{1,2})[-./月](\d{1,2})日?$/);
  if(monthDay){const reason=calendarError(null,+monthDay[1],+monthDay[2]);if(reason)return block(reason)}
  if(/^[+-]?\d{4,}[-./年]/.test(s))return block('日期格式或数值不合法，请填写 YYYY-MM-DD 或 YYYY-MM');
  return{value:'',cleared:true};
 }
 const reason=calendarError(y,m,d);if(reason)return block(reason);
 const earliestYear=currentYear-MAX_DATE_AGE_YEARS;
 if(y>currentYear)return block(`日期年份 ${y} 大于当年 ${currentYear}；允许年份为 ${earliestYear} 至 ${currentYear}`);
 if(y<earliestYear)return block(`日期年份 ${y} 距当年 ${currentYear} 超过 ${MAX_DATE_AGE_YEARS} 年；允许年份为 ${earliestYear} 至 ${currentYear}`);
 return{value:`${String(y).padStart(4,'0')}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`};
}
function expandDecimal(s){ // Decimal string arithmetic only; never Number(device code).
 const m=s.match(/^([+-]?)(\d+)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/);if(!m)return s;
 const fraction=m[3]||'',exponent=parseInt(m[4]||'0',10);if(Math.abs(exponent)>1000)throw Error('设备代码指数过大');
 let digits=m[2]+fraction,pos=m[2].length+exponent;
 let out=pos<=0?'0.'+'0'.repeat(-pos)+digits:pos>=digits.length?digits+'0'.repeat(pos-digits.length):digits.slice(0,pos)+'.'+digits.slice(pos);
 if(out.includes('.'))out=out.replace(/0+$/,'').replace(/\.$/,'');return(m[1]==='-'?'-':'')+out;
}
function outputFilename(filename,now=new Date()){
 const parts=new Intl.DateTimeFormat('en',{timeZone:'Asia/Shanghai',month:'2-digit',day:'2-digit'}).formatToParts(now);
 const stamp=parts.find(p=>p.type==='month').value+parts.find(p=>p.type==='day').value;
 return filename.replace(/\.xlsx$/i,'')+'_导入版_'+stamp+'.xlsx';
}
async function process(buffer, customMappings={},options={}){
 const mappings=Object.assign(Object.create(null),MAP,customMappings);
 const currentYear=dateYear(options.now??new Date());
 const zip=await JSZip.loadAsync(buffer);const read=async p=>{const f=zip.file(p);if(!f)throw Error('无效 XLSX：缺少 '+p);return f.async('string')};
 const wb=parse(await read('xl/workbook.xml')),rels=parse(await read('xl/_rels/workbook.xml.rels'));
 const epoch=['1','true'].includes(first(wb,'workbookPr')?.getAttribute('date1904'));
 // A four-digit date serial is not a year-only entry when Excel marks it as a date.
 const dateStyles=new Set();
 if(zip.file('xl/styles.xml')){
  const styles=parse(await read('xl/styles.xml'));
  const formats=new Map(nodes(styles,'numFmt').map(f=>[Number(f.getAttribute('numFmtId')),f.getAttribute('formatCode')]));
  nodes(first(styles,'cellXfs')||styles,'xf').forEach((xf,index)=>{
   const id=Number(xf.getAttribute('numFmtId'));
   const format=(formats.get(id)||'').replace(/"[^"]*"|\\.|\[[^\]]*\]/g,'');
   if(id>=14&&id<=22||id>=27&&id<=36||id>=45&&id<=47||id>=50&&id<=58||/[ymd]/i.test(format))dateStyles.add(index);
  });
 }
 const shared=zip.file('xl/sharedStrings.xml')?nodes(parse(await read('xl/sharedStrings.xml')),'si').map(si=>nodes(si,'t').map(t=>t.textContent).join('')):[];
 const value=c=>{if(!c)return'';const t=c.getAttribute('t');return t==='s'?(shared[parseInt(txt(c,'v'),10)]??''):t==='inlineStr'?nodes(c,'t').map(t=>t.textContent).join(''):txt(c,'v')};
 const typedValue=c=>{const v=value(c),t=c?.getAttribute('t');if(t==='b')return v==='1';if(c&&(!t||t==='n')&&v!==''&&Number.isFinite(Number(v)))return Number(v);return v};
 const standardized={equipment:[],technicalParameters:[],schemas:{equipment:[],technicalParameters:[]}};
 options.onRead?.();
 const issues=[],plans=[];let total=0,cleared=0,filled=0,numericCodes=0;
 const issue=(sheet,row,field,raw,reason)=>issues.push({sheet,row,field,raw,reason});
 for(const sheet of nodes(wb,'sheet')){
  if(['hidden','veryHidden'].includes(sheet.getAttribute('state')) || /^hidden@/i.test(sheet.getAttribute('name')||''))continue;
  const name=sheet.getAttribute('name'),rel=nodes(rels,'Relationship').find(r=>r.getAttribute('Id')===sheet.getAttribute('r:id'));if(!rel)continue;
  const target=rel.getAttribute('Target');const path=target.startsWith('/')?target.slice(1):'xl/'+target.replace(/^\.\//,'');
  const xml=await read(path),doc=parse(xml),rows=nodes(doc,'row');if(!rows.some(r=>nodes(r,'c').some(c=>value(c)!=='')))continue;
  const header=rows.find(r=>r.getAttribute('r')==='1'),cols={};
  const allHeaders=header?nodes(header,'c').map(c=>({field:value(c).trim(),column:c.getAttribute('r').replace(/\d+$/,'')})).filter(h=>h.field):[];
  const schema={sheet:name,headers:allHeaders.map(h=>h.field),duplicates:allHeaders.filter((h,i,a)=>a.findIndex(x=>x.field===h.field)!==i).map(h=>h.field)};
  const record=row=>{const values=Object.create(null),rn=row.getAttribute('r'),cells=nodes(row,'c');for(const h of allHeaders){const c=cells.find(c=>c.getAttribute('r')===h.column+rn);values[h.field]=/设备编码|设备位号|人员编码|部门编码|注册代码/.test(h.field)?value(c):typedValue(c)}return{values,source:{sheet:name,row:Number(rn)}}};
  if(name==='技术参数'||(schema.headers.includes('参数')&&schema.headers.includes('设备档案.设备编码'))){
   standardized.schemas.technicalParameters.push(schema);
   for(const row of rows)if(+row.getAttribute('r')>1&&nodes(row,'c').some(c=>value(c)!==''))standardized.technicalParameters.push(record(row));
   continue;
  }
  if(header)for(const c of nodes(header,'c')){const v=value(c).trim();if(FIELDS.includes(v)||v===ASSET_FIELD){if(cols[v])issue(name,'表头',v,v,'字段名重复');cols[v]=c.getAttribute('r').replace(/\d+$/,'')}}
  if(!FIELDS.some(f=>cols[f]))continue;
  standardized.schemas.equipment.push(schema);
  for(const f of FIELDS)if(!cols[f])issue(name,'表头',f,'','缺少所需字段');
  const replacements=new Map(),additions=new Map(),recordByRow=new Map();
  const set=(row,field,v)=>{if(!cols[field])return;const addr=cols[field]+row.getAttribute('r');const original=nodes(row,'c').find(c=>c.getAttribute('r')===addr);const style=original?.getAttribute('s');const cell=`<c r="${addr}"${style?` s="${style}"`:''} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;const rec=recordByRow.get(row.getAttribute('r'));if(rec)rec.values[field]=v;if(original)replacements.set(addr,cell);else{const n=row.getAttribute('r');if(!additions.has(n))additions.set(n,[]);additions.get(n).push({addr,cell})}};
  for(const row of rows){const rn=row.getAttribute('r');if(+rn<=1)continue;const cells=nodes(row,'c');if(!cells.some(c=>value(c)!==''))continue;total++;
   const rec=record(row);recordByRow.set(rn,rec);standardized.equipment.push(rec);
   const get=f=>cells.find(c=>c.getAttribute('r')===cols[f]+rn);
   const rawName=value(get(FIELDS[0])),nameValue=rawName.trim();if(nameValue&&!Object.hasOwn(mappings,nameValue))issue(name,rn,FIELDS[0],rawName,'未配置对应编码');else{set(row,FIELDS[1],mappings[nameValue]||'');if(nameValue)filled++}
   for(const f of DATE_FIELDS){const c=get(f),raw=value(c);let result;
    if(c?.getAttribute('t')==='d'){result=date(raw.split('T')[0],false,epoch,currentYear);}else{const native=!!c&&(!c.getAttribute('t')||c.getAttribute('t')==='n');result=date(raw,native,epoch,currentYear,native&&dateStyles.has(Number(c.getAttribute('s')||0)))}
    if(result.block)issue(name,rn,f,raw,result.reason);else{set(row,f,result.value);if(result.cleared)cleared++}
   }
   const codeCell=get(FIELDS[5]);let code=value(codeCell);if(code.trim()===''||code.trim()==='/')code='';
   if(code&&(!codeCell.getAttribute('t')||codeCell.getAttribute('t')==='n')){numericCodes++;code=expandDecimal(code);}
   set(row,FIELDS[5],code?JSON.stringify({'注册代码':code}):'');
   if(cols[ASSET_FIELD]){const raw=value(get(ASSET_FIELD));
    if(raw.trim()===''||raw==='/')set(row,ASSET_FIELD,'');
    else if(['A','B','C','a','b','c'].includes(raw))set(row,ASSET_FIELD,raw.toUpperCase());
    else issue(name,rn,ASSET_FIELD,raw,'空值或 / 会置空；其余仅允许单个 A、B、C、a、b 或 c，其他内容均须修正');
   }
  }
  plans.push({name,path,xml,replacements,additions});
 }
 if(!plans.length)issue('工作簿','—','—','','未识别到设备台账：请确认台账第 1 行包含设备小类名称、设备小类编码、出厂日期、启用日期、设备制造日期、设备代码。辅助工作表不会参与处理。');
 if(issues.length)return{issues,total,standardized};
 const colIndex=a=>{let n=0;for(const c of a.replace(/\d/g,''))n=n*26+c.charCodeAt(0)-64;return n};
 for(const p of plans){const written=new Set();let xml=p.xml.replace(ROW_XML,rowXml=>{
  const n=rowXml.match(/\br=["'](\d+)["']/)?.[1];let out=rowXml.replace(CELL_XML,c=>{const a=c.match(/\br=["']([^"']+)["']/)?.[1];if(!p.replacements.has(a))return c;written.add(a);return p.replacements.get(a)});
  for(const add of p.additions.get(n)||[]){let inserted=false;out=out.replace(CELL_XML,c=>{const a=c.match(/\br=["']([^"']+)["']/)?.[1];if(!inserted&&colIndex(a)>colIndex(add.addr)){inserted=true;return add.cell+c}return c});if(!inserted)out=out.replace(/<\/(?:\w+:)?row>$/,close=>{inserted=true;return add.cell+close});if(inserted)written.add(add.addr)}
  return out;
 });
 for(const addr of [...p.replacements.keys(),...Array.from(p.additions.values()).flat().map(c=>c.addr)])if(!written.has(addr))throw Error('工作表“'+p.name+'”的单元格 '+addr+' 未成功写入，已终止导出');
 zip.file(p.path,xml)}
 const output=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});return{issues:[],total,cleared,filled,numericCodes,output,standardized};
}
return{MAP,FIELDS,DATE_FIELDS,MAX_DATE_AGE_YEARS,dateYear,date,expandDecimal,outputFilename,process};})();
if(typeof module!=='undefined')module.exports=EM;
