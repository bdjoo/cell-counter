/* Minimal .xlsx writer: plain values only (numbers and text), one or more sheets, no compression.
   XlsxLite.build([{name:'sheet', aoa:[[...],[...]]}]) -> Uint8Array */
(function(root){
'use strict';
var enc=new TextEncoder();
var NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
var REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
var PKG='http://schemas.openxmlformats.org/package/2006';
var HEAD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

var CRC=(function(){
  var t=new Uint32Array(256);
  for(var n=0;n<256;n++){var c=n;for(var k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);t[n]=c>>>0;}
  return t;
})();
function crc32(u8){
  var c=0xFFFFFFFF;
  for(var i=0;i<u8.length;i++)c=CRC[(c^u8[i])&0xFF]^(c>>>8);
  return (c^0xFFFFFFFF)>>>0;
}
function esc(s){
  return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'')
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function colName(i){
  var s='';i++;
  while(i>0){var m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26);}
  return s;
}
function sheetXml(aoa){
  var out=HEAD+'<worksheet xmlns="'+NS+'"><sheetData>';
  for(var r=0;r<aoa.length;r++){
    out+='<row r="'+(r+1)+'">';
    for(var c=0;c<aoa[r].length;c++){
      var v=aoa[r][c];
      if(v===''||v===null||v===undefined)continue;
      var ref=colName(c)+(r+1);
      if(typeof v==='number'&&isFinite(v))out+='<c r="'+ref+'"><v>'+v+'</v></c>';
      else out+='<c r="'+ref+'" t="inlineStr"><is><t xml:space="preserve">'+esc(v)+'</t></is></c>';
    }
    out+='</row>';
  }
  return out+'</sheetData></worksheet>';
}
function sheetName(s,i){
  s=String(s||'').replace(/[\[\]\*\?\/\\:]/g,' ').slice(0,31).trim();
  return s||('Sheet'+(i+1));
}
function zip(files){
  var d=new Date();
  var time=(d.getHours()<<11)|(d.getMinutes()<<5)|(d.getSeconds()>>1);
  var date=((d.getFullYear()-1980)<<9)|((d.getMonth()+1)<<5)|d.getDate();
  var parts=[],central=[],offset=0;
  files.forEach(function(f){
    var name=enc.encode(f.name),data=f.data,crc=crc32(data);
    var lh=new DataView(new ArrayBuffer(30));
    lh.setUint32(0,0x04034b50,true);lh.setUint16(4,20,true);lh.setUint16(6,0x0800,true);lh.setUint16(8,0,true);
    lh.setUint16(10,time,true);lh.setUint16(12,date,true);lh.setUint32(14,crc,true);
    lh.setUint32(18,data.length,true);lh.setUint32(22,data.length,true);
    lh.setUint16(26,name.length,true);lh.setUint16(28,0,true);
    parts.push(new Uint8Array(lh.buffer),name,data);
    var ch=new DataView(new ArrayBuffer(46));
    ch.setUint32(0,0x02014b50,true);ch.setUint16(4,20,true);ch.setUint16(6,20,true);ch.setUint16(8,0x0800,true);ch.setUint16(10,0,true);
    ch.setUint16(12,time,true);ch.setUint16(14,date,true);ch.setUint32(16,crc,true);
    ch.setUint32(20,data.length,true);ch.setUint32(24,data.length,true);
    ch.setUint16(28,name.length,true);ch.setUint16(30,0,true);ch.setUint16(32,0,true);
    ch.setUint16(34,0,true);ch.setUint16(36,0,true);ch.setUint32(38,0,true);ch.setUint32(42,offset,true);
    central.push(new Uint8Array(ch.buffer),name);
    offset+=30+name.length+data.length;
  });
  var csize=0;central.forEach(function(p){csize+=p.length;});
  var end=new DataView(new ArrayBuffer(22));
  end.setUint32(0,0x06054b50,true);end.setUint16(4,0,true);end.setUint16(6,0,true);
  end.setUint16(8,files.length,true);end.setUint16(10,files.length,true);
  end.setUint32(12,csize,true);end.setUint32(16,offset,true);end.setUint16(20,0,true);
  var all=parts.concat(central,[new Uint8Array(end.buffer)]),total=0;
  all.forEach(function(p){total+=p.length;});
  var out=new Uint8Array(total),pos=0;
  all.forEach(function(p){out.set(p,pos);pos+=p.length;});
  return out;
}
function build(sheets){
  var n=sheets.length,i;
  var types=HEAD+'<Types xmlns="'+PKG+'/content-types">'
    +'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    +'<Default Extension="xml" ContentType="application/xml"/>'
    +'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    +'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
  var wb=HEAD+'<workbook xmlns="'+NS+'" xmlns:r="'+REL+'"><sheets>';
  var wbRels=HEAD+'<Relationships xmlns="'+PKG+'/relationships">';
  for(i=0;i<n;i++){
    types+='<Override PartName="/xl/worksheets/sheet'+(i+1)+'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    wb+='<sheet name="'+esc(sheetName(sheets[i].name,i))+'" sheetId="'+(i+1)+'" r:id="rId'+(i+1)+'"/>';
    wbRels+='<Relationship Id="rId'+(i+1)+'" Type="'+REL+'/worksheet" Target="worksheets/sheet'+(i+1)+'.xml"/>';
  }
  types+='</Types>';
  wb+='</sheets></workbook>';
  wbRels+='<Relationship Id="rId'+(n+1)+'" Type="'+REL+'/styles" Target="styles.xml"/></Relationships>';
  var rootRels=HEAD+'<Relationships xmlns="'+PKG+'/relationships">'
    +'<Relationship Id="rId1" Type="'+REL+'/officeDocument" Target="xl/workbook.xml"/></Relationships>';
  var styles=HEAD+'<styleSheet xmlns="'+NS+'">'
    +'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    +'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    +'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    +'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    +'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    +'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    +'</styleSheet>';
  var files=[
    {name:'[Content_Types].xml',data:enc.encode(types)},
    {name:'_rels/.rels',data:enc.encode(rootRels)},
    {name:'xl/workbook.xml',data:enc.encode(wb)},
    {name:'xl/_rels/workbook.xml.rels',data:enc.encode(wbRels)},
    {name:'xl/styles.xml',data:enc.encode(styles)}
  ];
  for(i=0;i<n;i++)files.push({name:'xl/worksheets/sheet'+(i+1)+'.xml',data:enc.encode(sheetXml(sheets[i].aoa))});
  return zip(files);
}
root.XlsxLite={build:build};
if(typeof module!=='undefined'&&module.exports)module.exports=root.XlsxLite;
})(typeof self!=='undefined'?self:this);
