const assert=require('node:assert/strict');
const {test}=require('node:test');
const EM=require('../dist/engine.js');
const year=2026;
const parse=(value,native=false,epoch1904=false,formattedDate=false)=>EM.date(value,native,epoch1904,year,formattedDate);

test('all three equipment date fields share the same policy',()=>{
 assert.deepEqual(EM.DATE_FIELDS,['出厂日期','启用日期','设备制造日期']);
 assert.equal(EM.MAX_DATE_AGE_YEARS,100);
});

for(const value of ['2026-02-30','2026-02-00','2026-00-00','2026/13/01','2026年4月31日','2025-02-29','1900-02-29','2026-01-32','0000-01-01','2026-00','2026年0月','2月30日','2月0日','0月0日','2/30','2026-02--1','2026-002-01']){
 test('impossible calendar date blocks: '+value,()=>{
  const result=parse(value);assert.equal(result.block,true);assert(result.reason);assert(!Object.hasOwn(result,'value'));
 });
}

for(const [value,expected] of [
 ['2024-02-29','2024-02-29'],['2000年2月29日','2000-02-29'],['2026/2/28','2026-02-28'],
 ['2026.04.30','2026-04-30'],['2026-2','2026-02-01'],['2026年2月','2026-02-01'],
 ['1926-01-01','1926-01-01'],['2026-12-31','2026-12-31']
])test('valid date normalizes: '+value,()=>assert.deepEqual(parse(value),{value:expected}));

for(const value of ['2027-01-01','2027-02','9999年12月31日']){
 test('future year blocks: '+value,()=>assert.match(parse(value).reason,/大于当年 2026/));
}
for(const value of ['1925-12-31','1900-01','1800年1月1日']){
 test('year older than 100 years blocks: '+value,()=>assert.match(parse(value).reason,/超过 100 年.*1926 至 2026/));
}

test('year limit follows the China calendar at New Year',()=>{
 assert.equal(EM.dateYear(new Date('2026-12-31T15:59:59Z')),2026);
 assert.equal(EM.dateYear(new Date('2026-12-31T16:00:00Z')),2027);
 assert(EM.date('2027-01-01',false,false,2026).block);
 assert.deepEqual(EM.date('2027-01-01',false,false,2027),{value:'2027-01-01'});
 assert(EM.date('1926-12-31',false,false,2027).block);
});

test('existing blank, slash, non-date text and year-only rules remain',()=>{
 for(const value of ['','  ','/','非日期'])assert.equal(parse(value).value,'');
 for(const value of ['2026','2026年'])assert.equal(parse(value).reason,'仅填写年份，无法确定完整日期');
});

test('Excel serial dates use the same calendar and year checks',()=>{
 assert.deepEqual(parse('46309',true),{value:'2026-10-14'});
 assert.deepEqual(parse('46309.75',true),{value:'2026-10-14'});
 assert.deepEqual(parse('4.6309e4',true),{value:'2026-10-14'});
 assert.deepEqual(parse('44847',true,true),{value:'2026-10-14'});
 for(const value of ['60','60.5'])assert.match(parse(value,true).reason,/1900-02-29/);
 for(const value of ['-1','9999999','1e999'])assert(parse(value,true).block);
 assert.match(parse('1',true).reason,/超过 100 年/);
 assert.match(parse('50000',true).reason,/大于当年/);
});

test('four-digit Excel date serials at the oldest allowed year are valid',()=>{
 assert.deepEqual(parse('9498',true,false,true),{value:'1926-01-01'});
 assert.deepEqual(parse('9498e0',true,false,true),{value:'1926-01-01'});
 assert.deepEqual(parse('2026.03',true,false,true),{value:'2026-03-01'});
 assert.match(parse('9497',true,false,true).reason,/超过 100 年/);
 assert.equal(parse('2026',true).reason,'仅填写年份，无法确定完整日期');
});

test('Gregorian century leap rule and both Excel epochs',()=>{
 assert(EM.date('2100-02-29',false,false,2100).block);
 assert.deepEqual(EM.date('2000-02-29',false,false,2100),{value:'2000-02-29'});
 assert.deepEqual(EM.date('59',true,false,2000),{value:'1900-02-28'});
 assert.deepEqual(EM.date('61',true,false,2000),{value:'1900-03-01'});
 assert.deepEqual(EM.date('60',true,true,2000),{value:'1904-03-01'});
});
