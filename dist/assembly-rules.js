/* Field names, never source filenames or fixed Excel column numbers. */
const EM_IMPORT_RULES=Object.freeze({
 templateUrl:'templates/设备档案导入JLTY0714.xls',
 contractUrl:'templates/template-contract.json',
 sheets:{equipment:'设备档案',parameters:'技术参数'},
 aliases:{
  '设备编码':['设备编码（设备位号）','设备编码'],
  '注册代码':['设备代码','注册代码'],
  // These display-name aliases share the same field identifiers in the supplied templates.
  '规格型号':['设备型号','规格型号'],
  '区域位置.区域编码':['位置编码','区域位置.区域编码'],
  '区域位置.区域名称':['位置编码名称','区域位置.区域名称'],
  '详细位置':['设备详细位置','详细位置'],
  '制造厂':['制造商名称','制造厂'],
  '供应商':['供应商名称','供应商'],
  '使用年限':['设备使用年限','使用年限'],
  '专业':['设备大类名称','专业']
 },
 requiredEquipment:['设备编码'],
 requiredParameters:['设备档案.设备编码','参数'],
 criticalHeaders:{'设备档案':['设备编码','注册代码'],'技术参数':['设备档案.设备编码','参数','值范围']}
});
if(typeof module!=='undefined')module.exports=EM_IMPORT_RULES;
