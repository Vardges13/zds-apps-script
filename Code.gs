const API_KEY=***
const SHEETS={employees:'zp_employees',daily:'zp_daily',advances:'zp_advances',schedule:'zp_schedule',settings:'zp_settings'};

function checkKey(e){return(e.parameter.key||'')===API_KEY}

function getOrCreateSheet(name){
var ss=SpreadsheetApp.getActiveSpreadsheet();
var sheet=ss.getSheetByName(name);
if(!sheet){
sheet=ss.insertSheet(name);
var h=getSheetHeaders(name);
if(h)sheet.getRange(1,1,1,h.length).setValues([h]);
}
return sheet;
}

function getSheetHeaders(n){
if(n==='zp_employees')return['id','name','position','calcType','fixedDay','fixedMonth','percent','active','appRole','login','pass','permissions'];
if(n==='zp_daily')return['date','emp_id','worked','revenue','patients'];
if(n==='zp_advances')return['emp_id','month','amount'];
if(n==='zp_schedule')return['emp_id','date','status'];
if(n==='zp_settings')return['key','value'];
return null;
}

function readSheet(n){
var s=getOrCreateSheet(n);
var d=s.getDataRange().getValues();
if(d.length<=1)return[];
var h=d[0];
var r=[];
for(var i=1;i<d.length;i++){
if(!d[i][0]&&!d[i][1])continue;
var o={};
for(var j=0;j<h.length&&j<d[i].length;j++){
o[h[j]]=d[i][j];
}
r.push(o);
}
return r;
}

function writeSheet(n,data){
var s=getOrCreateSheet(n);
var h=getSheetHeaders(n);
if(!h)return;
s.getRange(2,1,s.getMaxRows(),s.getMaxColumns()).clearContent();
if(!data||!data.length)return;
var rows=data.map(function(item){return h.map(function(k){return item[k]===undefined?'':item[k]})});
s.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

function upsertSheet(n,data){
if(!data||!data.length)return;
var h=getSheetHeaders(n);
var ex=readSheet(n);
var kf;
if(n==='zp_employees')kf=['id'];
else if(n==='zp_daily'||n==='zp_schedule')kf=['date','emp_id'];
else if(n==='zp_advances')kf=['month','emp_id'];
else kf=[];
for(var i=0;i<data.length;i++){
var item=data[i];
var k=kf.map(function(f){return String(item[f]||'')}).join('|');
var found=false;
for(var j=0;j<ex.length;j++){
var ek=kf.map(function(f){return String(ex[j][f]||'')}).join('|');
if(ek===k){
for(var p=0;p<h.length;p++){
if(item[h[p]]!==undefined)ex[j][h[p]]=item[h[p]];
}
found=true;
break;
}
}
if(!found){
var nr={};
for(var p=0;p<h.length;p++)nr[h[p]]=item[h[p]]!==undefined?item[h[p]]:'';
ex.push(nr);
}
}
writeSheet(n,ex);
}

function jr(st,d,msg){
var r={status:st};
if(d!==undefined)r.data=d;
if(msg)r.message=msg;
return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function pb(e){
try{return JSON.parse(e.postData.contents)}catch(x){return null}
}

function doGet(e){
try{
if(!checkKey(e))return jr('error',null,'Неверный ключ');
var a=e.parameter.action||'';

// READ операции
if(a==='ping')return jr('ok',{pong:true});
if(a==='read_employees')return jr('ok',readSheet(SHEETS.employees));
if(a==='read_daily'){
var d=readSheet(SHEETS.daily);
if(e.parameter.date)d=d.filter(function(r){return String(r.date)===String(e.parameter.date)});
if(e.parameter.emp_id)d=d.filter(function(r){return String(r.emp_id)===String(e.parameter.emp_id)});
return jr('ok',d);
}
if(a==='read_advances'){
var d=readSheet(SHEETS.advances);
if(e.parameter.month)d=d.filter(function(r){return String(r.month)===String(e.parameter.month)});
return jr('ok',d);
}
if(a==='read_schedule'){
var d=readSheet(SHEETS.schedule);
if(e.parameter.month)d=d.filter(function(r){return String(r.date).startsWith(String(e.parameter.month))});
return jr('ok',d);
}
if(a==='read_settings'){
var d=readSheet(SHEETS.settings);
var o={};
for(var i=0;i<d.length;i++)if(d[i].key)o[d[i].key]=d[i].value;
return jr('ok',o);
}
if(a==='read_all')return jr('ok',{employees:readSheet(SHEETS.employees),daily:readSheet(SHEETS.daily),advances:readSheet(SHEETS.advances),schedule:readSheet(SHEETS.schedule)});

// WRITE операции через GET (data как JSON строка в параметре)
if(a==='write_daily'||a==='write_employees'||a==='write_advances'||a==='write_schedule'||a==='write_settings'){
var sheetName=a.replace('write_','zp_');
var dataStr=e.parameter.data||'';
var rep=(e.parameter.replace==='true');
try{var data=JSON.parse(dataStr)}catch(x){return jr('error',null,'Неверный JSON')}
if(!data||!data.length)return jr('error',null,'Нет данных');
if(rep)writeSheet(sheetName,data);else upsertSheet(sheetName,data);
return jr('ok',{written:data.length});
}

// Булковая запись через GET
if(a==='bulk_write'){
var dataStr=e.parameter.data||'';
try{var d=JSON.parse(dataStr)}catch(x){return jr('error',null,'Неверный JSON')}
if(d.employees)writeSheet(SHEETS.employees,d.employees);
if(d.daily)writeSheet(SHEETS.daily,d.daily);
if(d.advances)writeSheet(SHEETS.advances,d.advances);
if(d.schedule)writeSheet(SHEETS.schedule,d.schedule);
var t=0;if(d.employees)t+=d.employees.length;if(d.daily)t+=d.daily.length;if(d.advances)t+=d.advances.length;if(d.schedule)t+=d.schedule.length;
return jr('ok',{written:t});
}

return jr('error',null,'Неизвестное: '+a);
}catch(x){return jr('error',null,'Ошибка: '+x.message)}
}

function doPost(e){
try{
if(!checkKey(e))return jr('error',null,'Неверный ключ');
var a=e.parameter.action||'';
var b=pb(e);
if(!b)return jr('error',null,'Нет данных');
var rep=(b.replace===true);
var d=b.data;
if(!d)return jr('error',null,'Нет данных');
if(a==='write_employees'){if(rep)writeSheet(SHEETS.employees,d);else upsertSheet(SHEETS.employees,d);return jr('ok',{written:d.length})}
if(a==='write_daily'){if(rep)writeSheet(SHEETS.daily,d);else upsertSheet(SHEETS.daily,d);return jr('ok',{written:d.length})}
if(a==='write_advances'){if(rep)writeSheet(SHEETS.advances,d);else upsertSheet(SHEETS.advances,d);return jr('ok',{written:d.length})}
if(a==='write_schedule'){if(rep)writeSheet(SHEETS.schedule,d);else upsertSheet(SHEETS.schedule,d);return jr('ok',{written:d.length})}
if(a==='write_settings'){
var rows=[];
if(Array.isArray(d))rows=d;
else{var keys=Object.keys(d);for(var i=0;i<keys.length;i++)rows.push({key:keys[i],value:String(d[keys[i]])})}
if(rep)writeSheet(SHEETS.settings,rows);else upsertSheet(SHEETS.settings,rows);
return jr('ok',{written:rows.length});
}
if(a==='bulk_write'){
if(d.employees)writeSheet(SHEETS.employees,d.employees);
if(d.daily)writeSheet(SHEETS.daily,d.daily);
if(d.advances)writeSheet(SHEETS.advances,d.advances);
if(d.schedule)writeSheet(SHEETS.schedule,d.schedule);
var t=0;if(d.employees)t+=d.employees.length;if(d.daily)t+=d.daily.length;if(d.advances)t+=d.advances.length;if(d.schedule)t+=d.schedule.length;
return jr('ok',{written:t});
}
return jr('error',null,'Неизвестное: '+a);
}catch(x){return jr('error',null,'Ошибка: '+x.message)}
}
