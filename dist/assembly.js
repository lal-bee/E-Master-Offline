/* Stage two consumes stage one's records; it never reparses the uploaded file. */
const EM_ASSEMBLY=(()=>{
 const engine=typeof EM!=='undefined'?EM:require('./engine.js');
 const rules=typeof EM_IMPORT_RULES!=='undefined'?EM_IMPORT_RULES:require('./assembly-rules.js');
 const rowPattern=/<(?:\w+:)?row\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?row>)/g;
 const nodes=(d,n)=>Array.from(d.getElementsByTagNameNS('*',n));
 const first=(d,n)=>nodes(d,n)[0];
 const parse=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(nodes(d,'parsererror').length)throw Error('工作簿 XML 无法解析');return d};
 const blank=v=>v==null||String(v).trim()==='';
 const sourceNames=field=>rules.aliases[field]||[field];
 const structureError=detail=>Object.assign(Error('目标模板结构发生变化：'+detail),{code:'TEMPLATE_STRUCTURE'});
 const escape=v=>String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
 function resolve(record,field){
  const candidates=sourceNames(field).filter(f=>Object.hasOwn(record.values,f));
  const populated=candidates.filter(f=>!blank(record.values[f]));
  const selected=populated[0]||candidates[0];
  return{value:selected?record.values[selected]:'',sourceField:selected||field,conflict:populated.length>1&&new Set(populated.map(f=>String(record.values[f]))).size>1};
 }
 function issue(record,field,raw,reason,kind='validation'){
  return{sheet:record?.source?.sheet||'源文件',row:record?.source?.row??'—',field,raw:raw??'',reason,kind};
 }
 function validate(data){
  const issues=[],duplicates=[],equipmentIssues=[],parameterIssues=[],warnings=[];
  const equipment=data?.equipment||[],parameters=data?.technicalParameters||[];
  const addEquipment=i=>{issues.push(i);equipmentIssues.push(i)};
  const addParameter=i=>{issues.push(i);parameterIssues.push(i)};
  for(const [kind,list] of Object.entries(data?.schemas||{}))for(const schema of list){
   const add=kind==='equipment'?addEquipment:addParameter;
   for(const field of schema.duplicates||[])add({sheet:schema.sheet,row:'表头',field,raw:field,reason:'源文件字段名重复，无法按字段名称确定映射',kind:'source_header'});
   const required=kind==='equipment'?rules.requiredEquipment:rules.requiredParameters;
   for(const field of required)if(!sourceNames(field).some(f=>schema.headers.includes(f)))add({sheet:schema.sheet,row:'表头',field,raw:'',reason:'源文件缺少必填字段',kind:'source_header'});
  }
  if(!equipment.length)addEquipment(issue(null,'设备编码','','没有可写入的设备档案数据'));
  const devices=new Map();
  for(const record of equipment){
   if(!record.source?.sheet||!Number.isInteger(record.source.row)||record.source.row<2)addEquipment(issue(record,'源行号','','原始 Excel 行号元数据缺失，不能生成导入文件'));
   for(const field of rules.requiredEquipment){const r=resolve(record,field);if(blank(r.value))addEquipment(issue(record,r.sourceField,r.value,'必填字段不能为空'));if(r.conflict)addEquipment(issue(record,field,r.value,'同一目标字段对应多个不同的源值，请修正源文件'))}
   const code=String(resolve(record,'设备编码').value??'');
   if(!blank(code)){if(!devices.has(code))devices.set(code,[]);devices.get(code).push(record)}
   for(const field of Object.keys(rules.aliases))if(resolve(record,field).conflict&&!rules.requiredEquipment.includes(field))addEquipment(issue(record,field,resolve(record,field).value,'同一目标字段对应多个不同的源值，请修正源文件'));
  }
  for(const [code,records] of devices)if(records.length>1)addEquipment({sheet:[...new Set(records.map(r=>r.source?.sheet||'源文件'))].join('、'),row:records.map(r=>r.source?.row??'—').join('、'),field:'设备编码',raw:code,reason:'设备编码重复，请修正源文件后重新上传',kind:'duplicate_equipment'});
  const grouped=new Map();
  for(const record of parameters){
   if(!record.source?.sheet||!Number.isInteger(record.source.row)||record.source.row<2)addParameter(issue(record,'源行号','','原始 Excel 行号元数据缺失，不能报告真实重复位置'));
   for(const field of rules.requiredParameters)if(blank(record.values[field]))addParameter(issue(record,field,record.values[field],'必填字段不能为空'));
   const code=String(record.values['设备档案.设备编码']??''),parameter=String(record.values['参数']??'');
   if(!blank(code)&&!devices.has(code))addParameter(issue(record,'设备档案.设备编码',code,'技术参数关联不到设备档案：'+code,'orphan_parameter'));
   if(blank(code)||blank(parameter))continue;
   if(!grouped.has(code))grouped.set(code,new Map());
   const group=grouped.get(code);if(!group.has(parameter))group.set(parameter,[]);group.get(parameter).push(record);
  }
  for(const [deviceCode,group] of grouped)for(const [parameter,records] of group)if(records.length>1){
   const locations=records.map(r=>({sheet:r.source?.sheet||'源文件',row:r.source?.row??'—',value:r.values['值范围']??''}));
   duplicates.push({deviceCode,parameter,locations});
   addParameter({sheet:[...new Set(locations.map(l=>l.sheet))].join('、'),row:locations.map(l=>l.row).join('、'),field:'参数',raw:deviceCode+' / '+parameter,reason:'同一设备的同名技术参数重复，请修改源文件后重新上传',kind:'duplicate_parameter'});
  }
  if(!(data?.schemas?.technicalParameters||[]).length)warnings.push({sheet:'技术参数',field:'',message:'源文件未提供技术参数数据，最终技术参数表只保留模板表头。'});
  return{issues,duplicates,equipmentIssues,parameterIssues,warnings};
 }
 function targetPath(target){
  const result=[];for(const part of (target.startsWith('/')?target.slice(1):'xl/'+target).split('/')){if(part==='..')result.pop();else if(part&&part!=='.')result.push(part)}return result.join('/');
 }
 async function open(buffer){
  const zip=await JSZip.loadAsync(buffer);
  const read=async p=>{const f=zip.file(p);if(!f)throw structureError('缺少 '+p);return f.async('string')};
  const workbookXml=await read('xl/workbook.xml'),workbook=parse(workbookXml),rels=parse(await read('xl/_rels/workbook.xml.rels'));
  const shared=zip.file('xl/sharedStrings.xml')?nodes(parse(await read('xl/sharedStrings.xml')),'si').map(si=>nodes(si,'t').map(t=>t.textContent).join('')):[];
  const value=c=>{if(!c)return'';const t=c.getAttribute('t');return t==='s'?(shared[Number(first(c,'v')?.textContent)]??''):t==='inlineStr'?nodes(c,'t').map(t=>t.textContent).join(''):(first(c,'v')?.textContent||'')};
  const sheets=[];
  for(const sheet of nodes(workbook,'sheet')){
   const rel=nodes(rels,'Relationship').find(r=>r.getAttribute('Id')===sheet.getAttribute('r:id'));
   if(!rel||rel.getAttribute('TargetMode')==='External')throw structureError('工作表关系异常');
   const path=targetPath(rel.getAttribute('Target')),xml=await read(path),doc=parse(xml),rows=nodes(doc,'row');
   const headerXml=(xml.match(rowPattern)||[]).find(r=>r.match(/\br=["']([^"']+)["']/)?.[1]==='1');
   if(!headerXml)throw structureError('工作表 '+sheet.getAttribute('name')+' 缺少表头');
   const header=rows.find(r=>r.getAttribute('r')==='1');
   const headers=nodes(header,'c').map(c=>({column:c.getAttribute('r').replace(/\d+$/,''),field:value(c),formula:first(c,'f')?.textContent??null}));
   sheets.push({name:sheet.getAttribute('name'),state:sheet.getAttribute('state')||'visible',path,headerXml,headers,xml,doc});
  }
  return{zip,workbookXml,sheets,value};
 }
 function snapshot(context){return{workbookXml:context.workbookXml,sheets:context.sheets.map(({name,state,path,headerXml,headers})=>({name,state,path,headerXml,headers}))}}
 function assertStructure(context,contract){
  if(!contract?.sheets||JSON.stringify(snapshot(context))!==JSON.stringify({workbookXml:contract.workbookXml,sheets:contract.sheets}))throw structureError('Sheet 数量、名称、顺序、隐藏状态、表头或表头公式不一致');
  for(const [name,required] of Object.entries(rules.criticalHeaders)){
   const sheet=context.sheets.find(s=>s.name===name);if(!sheet||sheet.state!=='visible')throw structureError('缺少关键工作表 '+name);
   const fields=sheet.headers.filter(h=>h.field).map(h=>h.field);
   if(new Set(fields).size!==fields.length)throw structureError(name+' 表头字段名重复');
   for(const field of required)if(!fields.includes(field))throw structureError(name+' 缺少关键表头 '+field);
  }
 }
 function mappingWarnings(data,context){
  const warnings=[];
  for(const [kind,sheetName] of [['equipment','设备档案'],['technicalParameters','技术参数']]){
   const target=context.sheets.find(s=>s.name===sheetName),schemas=data.schemas?.[kind]||[];
   const incoming=[...new Set(schemas.flatMap(s=>s.headers))],columns=target.headers.filter(h=>h.field);
   const used=new Set(columns.flatMap(h=>kind==='equipment'?sourceNames(h.field):[h.field]));
   for(const field of incoming)if(!used.has(field))warnings.push({sheet:sheetName,field,message:'源字段在目标模板中没有对应列，已忽略。'});
   for(const h of columns)if(!h.formula&&!(kind==='equipment'?sourceNames(h.field):[h.field]).some(f=>incoming.includes(f)))warnings.push({sheet:sheetName,field:h.field,message:'源文件未提供对应字段，目标数据列留空。'});
  }
  return warnings;
 }
 function cell(address,value,style){
  const attr=' r="'+address+'"'+(style?' s="'+escape(style)+'"':'');
  if(typeof value==='number'&&Number.isFinite(value))return'<c'+attr+'><v>'+value+'</v></c>';
  if(typeof value==='boolean')return'<c'+attr+' t="b"><v>'+(value?'1':'0')+'</v></c>';
  const text=value==null?'':String(value);
  if(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(text))throw Error('字段内容包含 Excel 不支持的控制字符，请修正源文件');
  return'<c'+attr+' t="inlineStr"><is><t xml:space="preserve">'+escape(text)+'</t></is></c>';
 }
 function writeSheet(sheet,records,isEquipment){
  const example=nodes(sheet.doc,'row').find(r=>Number(r.getAttribute('r'))>1),styleByColumn=new Map();
  if(example)for(const c of nodes(example,'c'))styleByColumn.set(c.getAttribute('r').replace(/\d+$/,''),c.getAttribute('s'));
  const fields=sheet.headers.filter(h=>h.field),expected=[];
  const rows=records.map((record,i)=>{const rn=i+2,values=[];
   const cells=fields.map(h=>{const v=isEquipment?resolve(record,h.field).value:(record.values[h.field]??'');values.push(v);return cell(h.column+rn,v,styleByColumn.get(h.column))}).join('');
   expected.push(values);return'<row r="'+rn+'">'+cells+'</row>';
  });
  let xml=sheet.xml.replace(/(<(?:\w+:)?sheetData\b[^>]*>)[\s\S]*?(<\/(?:\w+:)?sheetData>)/,(whole,start,end)=>start+sheet.headerXml+rows.join('')+end);
  if(xml===sheet.xml&&records.length)throw structureError(sheet.name+' 无法写入数据区');
  const dim=sheet.xml.match(/<(?:\w+:)?dimension\b[^>]*\bref=["']([^"']+)["']/)?.[1];
  if(dim){const lastColumn=dim.split(':').at(-1).replace(/\d+$/,'');xml=xml.replace(/(<(?:\w+:)?dimension\b[^>]*\bref=["'])[^"']+(["'])/,(whole,start,end)=>start+'A1:'+lastColumn+Math.max(1,records.length+1)+end)}
  return{xml,fields,expected};
 }
 async function writeValidated(data,template,contract,validation,progress){
  const context=await open(template);assertStructure(context,contract);
  const writes=[];
  for(const [name,records,isEquipment] of [['设备档案',data.equipment,true],['技术参数',data.technicalParameters,false]]){
   const sheet=context.sheets.find(s=>s.name===name),written=writeSheet(sheet,records,isEquipment);context.zip.file(sheet.path,written.xml);writes.push({name,path:sheet.path,...written});
  }
  progress?.('assemble','done');progress?.('generate','running');
  const output=await context.zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  const final=await open(output);assertStructure(final,contract);
  const before=await JSZip.loadAsync(template),changed=new Set(writes.map(w=>w.path));
  const parts=z=>Object.keys(z.files).filter(p=>!z.files[p].dir).sort();
  if(JSON.stringify(parts(before))!==JSON.stringify(parts(final.zip)))throw structureError('工作簿内部文件清单发生变化');
  for(const name of parts(before))if(!changed.has(name)){
   const a=await before.file(name).async('uint8array'),b=await final.zip.file(name).async('uint8array');
   if(a.length!==b.length||a.some((v,i)=>v!==b[i]))throw structureError('非写入内容发生变化：'+name);
  }
  for(const write of writes){const sheet=final.sheets.find(s=>s.name===write.name),rows=nodes(sheet.doc,'row').filter(r=>Number(r.getAttribute('r'))>1);
   if(rows.length!==write.expected.length)throw structureError(write.name+' 数据行数不一致');
   rows.forEach((row,i)=>write.fields.forEach((h,j)=>{const c=nodes(row,'c').find(c=>c.getAttribute('r')===h.column+(i+2)),v=write.expected[i][j];const expected=typeof v==='boolean'?(v?'1':'0'):String(v??'');if(final.value(c)!==expected)throw Error('最终文件回读校验失败：'+write.name+' '+h.column+(i+2))}));
  }
  progress?.('generate','done');
  return{...validation,warnings:[...validation.warnings,...mappingWarnings(data,context)],output,equipmentCount:data.equipment.length,parameterCount:data.technicalParameters.length,sheetCount:final.sheets.length,templateProtected:true};
 }
 async function assemble(data,template,contract){const validation=validate(data);if(validation.issues.length)return validation;return writeValidated(data,template,contract,validation)}
 async function loadTemplate(){
  const responses=await Promise.all([fetch(rules.templateUrl,{cache:'no-store'}),fetch(rules.contractUrl,{cache:'no-store'})]);
  if(responses.some(r=>!r.ok))throw Error('固定目标模板加载失败，请刷新后重试');
  return{template:await responses[0].arrayBuffer(),contract:await responses[1].json()};
 }
 async function run(buffer,mappings={},options={}){
  const statuses=[],progress=(id,state)=>{const found=statuses.find(s=>s.id===id);if(found)found.state=state;else statuses.push({id,state});options.onProgress?.(id,state)};
  let current='read';
  try{
   progress('read','running');
   const firstStage=await engine.process(buffer,mappings,{now:options.now,onRead:()=>{progress('read','done');current='standardize';progress(current,'running')}});
   if(firstStage.issues.length){progress('standardize','error');return{issues:firstStage.issues,duplicates:[],statuses,stage:'standardize'}}
   progress('standardize','done');
   const data=firstStage.standardized,validation=validate(data);current='equipment';progress(current,validation.equipmentIssues.length?'error':'done');
   current='parameters';progress(current,validation.parameterIssues.length?'error':'done');
   if(validation.issues.length)return{...validation,statuses,stage:validation.equipmentIssues.length?'equipment':'parameters'};
   current='assemble';progress(current,'running');const fixed=await(options.loadTemplate||loadTemplate)();
   const result=await writeValidated(data,fixed.template,fixed.contract,validation,(id,state)=>{current=id;progress(id,state)});
   return{...result,statuses,firstStage:{total:firstStage.total,filled:firstStage.filled,cleared:firstStage.cleared,numericCodes:firstStage.numericCodes},stage:'complete'};
  }catch(error){progress(current,'error');return{issues:[{sheet:current==='assemble'||current==='generate'?'目标模板':'工作簿',row:'—',field:'—',raw:'',reason:error.message,kind:error.code||'processing'}],duplicates:[],statuses,stage:current}}
 }
 return{validate,assemble,run,open,snapshot,rules};
})();
if(typeof module!=='undefined')module.exports=EM_ASSEMBLY;
