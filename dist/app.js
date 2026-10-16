const $=id=>document.getElementById(id);let selected=null,url=null,outputName='',busy=false;
const STORAGE_KEY='e-master-custom-mappings-v1';let customMappings=Object.create(null);
try{const saved=JSON.parse(localStorage.getItem(STORAGE_KEY)||'{}');if(saved&&typeof saved==='object'&&!Array.isArray(saved))for(const [n,c] of Object.entries(saved)){if(n.trim()&&typeof c==='string'&&c.trim()&&!Object.hasOwn(EM.MAP,n))customMappings[n]=c;}}catch{$('mappingFeedback').textContent='本地映射读取失败，可重新添加映射。'}
function renderMappings(){
 $('mapping').replaceChildren();const all={...EM.MAP,...customMappings};$('mappingCount').textContent=Object.keys(all).length;
 for(const [n,c] of Object.entries(all)){const e=document.createElement('span');e.textContent=n+' → '+c;if(Object.hasOwn(customMappings,n)){const b=document.createElement('button');b.type='button';b.className='remove-mapping';b.textContent='删除';b.setAttribute('aria-label','删除映射 '+n);b.onclick=()=>{if(busy)return;delete customMappings[n];mappingChanged('已删除“'+n+'”的自定义映射。')};e.append(b)}$('mapping').append(e)}
}
function mappingChanged(message){let persisted=true;try{localStorage.setItem(STORAGE_KEY,JSON.stringify(customMappings))}catch{persisted=false}reset();renderMappings();$('mappingFeedback').textContent=message+(persisted?' 已保存在当前浏览器。':' 当前浏览器无法保存，关闭页面后将丢失。')+(selected?' 请手动点击“开始处理”重新检查。':'');}
$('mappingForm').onsubmit=e=>{e.preventDefault();if(busy)return;const n=$('mappingName').value.trim(),c=$('mappingCode').value.trim();if(!n||!c){$('mappingFeedback').textContent='设备小类名称和编码均不能为空。';return}if(Object.hasOwn(EM.MAP,n)||Object.hasOwn(customMappings,n)){$('mappingFeedback').textContent='该设备小类名称已存在映射，请勿重复添加。自定义映射可删除后重新添加。';return}customMappings[n]=c;mappingChanged('已添加“'+n+' → '+c+'”。');$('mappingForm').reset();};
renderMappings();
const progressLabels={read:'文件读取完成',standardize:'数据标准化完成',equipment:'设备档案校验完成',parameters:'技术参数唯一性与关联校验通过',assemble:'目标模板装配完成',generate:'最终文件生成成功'};
let progressStates={};
function renderProgress(){
 $('progress').hidden=false;$('progress').replaceChildren();
 for(const [id,label] of Object.entries(progressLabels)){const state=progressStates[id]||'waiting',line=document.createElement('div');line.className='progress-line '+state;
  line.textContent=(state==='done'?'✓ ':state==='error'?'✕ ':state==='running'?'… ':'— ')+(state==='error'?(id==='generate'?'最终文件未生成':label.replace(/完成$|通过$/,'')+'失败'):state==='running'?label.replace(/完成$|通过$|成功$/,'')+'中…':state==='waiting'?label.replace(/完成$|通过$|成功$/,'')+'（未执行）':label);$('progress').append(line)}
}
function setProgress(id,state){progressStates[id]=state;renderProgress()}
function reset(){if(url)URL.revokeObjectURL(url);url=null;outputName='';progressStates={};for(const id of ['message','stats','issues','duplicateDetails','warnings','progress','downloadArea'])$(id).hidden=true;$('idle').hidden=false;$('statusBadge').textContent='待处理'}
function select(file){if(busy)return;reset();selected=null;$('start').disabled=true;if(!file)return;if(!/\.xlsx$/i.test(file.name)){show('请选择 .xlsx 文件。旧版 .xls 请先在 Excel 中另存为 .xlsx。',true);return}selected=file;$('fileTitle').textContent=file.name;$('fileHint').textContent=(file.size/1024).toFixed(1)+' KB · 已就绪';$('ready').textContent='文件已选择，可开始预检查';$('start').disabled=false}
function show(message,error=false){$('idle').hidden=true;$('message').hidden=false;$('message').className=error?'error':'';$('message').textContent=message}
$('file').onchange=e=>select(e.target.files[0]);$('drop').onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('file').click()}};for(const type of ['dragenter','dragover'])$('drop').addEventListener(type,e=>{e.preventDefault();$('drop').classList.add('drag')});$('drop').ondragleave=()=>$('drop').classList.remove('drag');$('drop').ondrop=e=>{e.preventDefault();$('drop').classList.remove('drag');select(e.dataTransfer.files[0])};
function table(headers,rows){const t=document.createElement('table'),thead=document.createElement('thead'),head=document.createElement('tr');for(const label of headers){const th=document.createElement('th');th.textContent=label;head.append(th)}thead.append(head);t.append(thead);const body=document.createElement('tbody');for(const values of rows){const tr=document.createElement('tr');for(const value of values){const td=document.createElement('td');td.textContent=value??'';tr.append(td)}body.append(tr)}t.append(body);return t}
function showProblems(result){
 const groups=result.duplicates||[];
 if(groups.length){const area=$('duplicateDetails');area.hidden=false;area.replaceChildren();const title=document.createElement('h3');title.textContent='共发现 '+groups.length+' 组重复技术参数';area.append(title);
  groups.forEach((group,index)=>{const block=document.createElement('div'),label=document.createElement('strong'),locations=document.createElement('p');block.className='duplicate-group';label.textContent=(index+1)+'. '+group.deviceCode+' / '+group.parameter;locations.textContent='重复位置：'+group.locations.map(l=>l.sheet+'：源文件第'+l.row+'行').join('、');block.append(label,locations,table(['源工作表','源 Excel 行号','对应值'],group.locations.map(l=>[l.sheet,l.row,l.value])));area.append(block)})}
 const other=result.issues.filter(i=>i.kind!=='duplicate_parameter');if(other.length){$('issues').hidden=false;$('issues').replaceChildren(table(['源工作表','源 Excel 行号','字段名称','原始值 / 设备编码','阻断原因'],other.map(i=>[i.sheet,i.row,i.field,i.raw,i.reason])))}
}
$('start').onclick=async()=>{
 if(!selected||busy)return;reset();busy=true;$('addMapping').disabled=true;$('start').disabled=true;$('file').disabled=true;$('statusBadge').textContent='处理中';show('正在进行第一阶段标准化和第二阶段校验，请稍候…');
 try{
  await new Promise(r=>setTimeout(r,30));
  const result=await EM_ASSEMBLY.run(await selected.arrayBuffer(),{...customMappings},{onProgress:setProgress});
  if(result.issues.length){setProgress('generate','error');$('statusBadge').textContent='处理失败';show('处理失败，文件未生成。'+(result.duplicates?.length?'共发现 '+result.duplicates.length+' 组重复技术参数。':'共发现 '+result.issues.length+' 个阻断问题。')+'请修正源文件或补充映射后，重新上传处理。',true);showProblems(result);return}
  $('statusBadge').textContent='处理完成';show('处理完成：已将 '+result.equipmentCount+' 条设备档案和 '+result.parameterCount+' 条技术参数写入固定目标模板。'+result.sheetCount+' 个 Sheet、表头及表头公式保护校验通过，最终 XLSX 文件已生成。'+(result.firstStage.numericCodes?' 部分设备代码原为数值单元格，文件原有精度或前导零损失无法恢复。':''));
  $('stats').hidden=false;$('stats').replaceChildren();for(const [label,value] of [['设备档案',result.equipmentCount],['技术参数',result.parameterCount],['日期置空',result.firstStage.cleared],['编码填充',result.firstStage.filled]]){const d=document.createElement('div'),b=document.createElement('b');d.textContent=label;b.textContent=value.toLocaleString()+' 项';d.append(b);$('stats').append(d)}
  if(result.warnings.length){$('warnings').hidden=false;$('warnings').replaceChildren();const details=document.createElement('details'),summary=document.createElement('summary');summary.textContent='非阻断提示（'+result.warnings.length+' 项）：查看忽略字段和空缺字段';details.append(summary,table(['工作表','字段','说明'],result.warnings.map(w=>[w.sheet,w.field,w.message])));$('warnings').append(details)}
  url=URL.createObjectURL(new Blob([result.output],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));outputName=EM.outputFilename(selected.name);$('outputFileName').textContent='最终文件：'+outputName;$('downloadArea').hidden=false;
 }catch(e){setProgress('generate','error');$('statusBadge').textContent='处理失败';show('处理失败，文件未生成：'+e.message,true)}
 finally{busy=false;$('addMapping').disabled=false;$('start').disabled=false;$('file').disabled=false}
};
$('download').onclick=()=>{if(!url)return;const a=document.createElement('a');a.href=url;a.download=outputName;a.click()};
