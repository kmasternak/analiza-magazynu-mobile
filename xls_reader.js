/* MaterialControl: offline BIFF8 (.xls) reader for schematic BOMs. */
(function(root){'use strict';
function parse(buffer){
 const d=new DataView(buffer), u=new Uint8Array(buffer), decoder=new TextDecoder('utf-16le');
 const U16=(p)=>d.getUint16(p,true),U32=(p)=>d.getUint32(p,true);
 if(u.length<512||U32(0)!==0xe011cfd0||U32(4)!==0xe11ab1a1)throw Error('To nie jest poprawny plik XLS (OLE).');
 const sectorSize=1<<U16(30), miniSize=1<<U16(32), sector=(n)=>512+n*sectorSize;
 const END=0xfffffffe, FREE=0xffffffff;
 const difat=[];for(let i=0;i<109;i++){const v=U32(76+4*i);if(v!==FREE)difat.push(v)}
 let ds=U32(68);for(let j=0;j<U32(72)&&ds!==END&&ds!==FREE;j++){
  const p=sector(ds);for(let k=0;k<sectorSize/4-1;k++){const v=U32(p+4*k);if(v!==FREE)difat.push(v)}ds=U32(p+sectorSize-4)
 }
 const fat=[];for(const s of difat){const p=sector(s);for(let k=0;k<sectorSize/4;k++)fat.push(U32(p+k*4))}
 function chain(start,limit=100000){const parts=[];let n=start;const seen=new Set();while(n!==END&&n!==FREE&&n<fat.length&&!seen.has(n)&&parts.length<limit){seen.add(n);parts.push(n);n=fat[n]}return parts}
 function concat(parts,size){const out=new Uint8Array(size);let pos=0;for(const part of parts){if(pos>=size)break;const slice=part.subarray(0,Math.min(part.length,size-pos));out.set(slice,pos);pos+=slice.length}return out}
 const dir=concat(chain(U32(48)).map(s=>u.subarray(sector(s),sector(s)+sectorSize)),(U32(40)||100)*sectorSize);
 const dv=new DataView(dir.buffer);const dirs=[];
 for(let p=0;p+128<=dir.length;p+=128){const n=dv.getUint16(p+64,true);if(n<2||n>64)continue;const name=decoder.decode(dir.subarray(p,p+n-2));const size=dv.getUint32(p+120,true);dirs.push({name,type:dir[p+66],start:dv.getUint32(p+116,true),size})}
 const rootEntry=dirs.find(x=>x.type===5),book=dirs.find(x=>x.type===2&&/^(Workbook|Book)$/i.test(x.name));
 if(!book)throw Error('Nie znaleziono strumienia Workbook w pliku XLS.');
 let bytes;
 if(book.size<U32(56)){
  if(!rootEntry)throw Error('Brak głównego strumienia XLS.');
  const mini=concat(chain(rootEntry.start).map(s=>u.subarray(sector(s),sector(s)+sectorSize)),rootEntry.size);
  const mf=[];for(const s of chain(U32(60)))for(let p=sector(s);p<sector(s)+sectorSize;p+=4)mf.push(U32(p));
  const parts=[];let n=book.start;const seen=new Set();while(n!==END&&n!==FREE&&n<mf.length&&!seen.has(n)){seen.add(n);parts.push(mini.subarray(n*miniSize,(n+1)*miniSize));n=mf[n]}bytes=concat(parts,book.size)
 }else bytes=concat(chain(book.start).map(s=>u.subarray(sector(s),sector(s)+sectorSize)),book.size);
 const b=new DataView(bytes.buffer), B16=p=>b.getUint16(p,true),B32=p=>b.getUint32(p,true);
 const recs=[];for(let p=0;p+4<=bytes.length;){const type=B16(p),len=B16(p+2);if(p+4+len>bytes.length)break;recs.push({type,start:p+4,end:p+4+len,offset:p});p+=4+len}
 const bounds=[];let sstRecord=-1;
 for(let i=0;i<recs.length;i++){const r=recs[i];if(r.type===0x0085){const len=bytes[r.start+6];const flags=bytes[r.start+7];const raw=bytes.subarray(r.start+8,r.start+8+len*(flags&1?2:1));bounds.push({offset:B32(r.start),name:(flags&1)?decoder.decode(raw):new TextDecoder('latin1').decode(raw)})}if(r.type===0x00fc)sstRecord=i}
 if(!bounds.length)throw Error('Nie znaleziono arkuszy w XLS.');
 const strings=[];
 if(sstRecord>=0){const chunks=[recs[sstRecord]];for(let i=sstRecord+1;i<recs.length&&recs[i].type===0x003c;i++)chunks.push(recs[i]);
  let ix=0,p=chunks[0].start;
  function advance(){while(ix<chunks.length&&p>=chunks[ix].end){ix++;if(ix<chunks.length)p=chunks[ix].start}if(ix>=chunks.length)throw Error('Niekompletna tabela tekstów XLS.')}
  function byte(){advance();return bytes[p++]}
  function word(){return byte()|(byte()<<8)}
  function double(){return word()|(word()<<16)}
  function chars(n,unicode){let out='';for(let j=0;j<n;j++){
   if(p>=chunks[ix].end){ix++;if(ix>=chunks.length)throw Error('Niekompletny tekst XLS.');p=chunks[ix].start;unicode=!!(byte()&1)}
   out+=String.fromCharCode(unicode?word():byte());
  }return out}
  const total=double(),unique=double();for(let j=0;j<unique&&j<1000000;j++){
   const n=word(),flags=byte(),rich=flags&8?word():0,ext=flags&4?double():0;
   strings.push(chars(n,!!(flags&1)));for(let k=0;k<rich*4+ext;k++)byte();
  }
 }
 const sheets=[];
 for(const sh of bounds.slice(0,2)){
  const rows=[];let active=false;
  for(const r of recs){if(r.offset===sh.offset){active=true;continue}if(!active)continue;if(r.type===0x000a)break;
   const p=r.start;if(r.end-p<6)continue;const row=B16(p),col=B16(p+2);if(row>65535||col>255)continue;
   let val;
   if(r.type===0x00fd&&r.end-p>=10)val=strings[B32(p+6)]||'';
   else if(r.type===0x0203&&r.end-p>=14)val=b.getFloat64(p+6,true);
   else if(r.type===0x027e&&r.end-p>=10){const raw=B32(p+6);val=raw&2?raw>>2:((raw&0xfffffffc) * 4294967296 + 0);if(!(raw&2)){const ab=new ArrayBuffer(8),dv2=new DataView(ab);dv2.setUint32(4,raw&0xfffffffc,true);val=dv2.getFloat64(0,true)}if(raw&1)val/=100}
   else if(r.type===0x0204&&r.end-p>=8){const n=B16(p+6);val=new TextDecoder('latin1').decode(bytes.subarray(p+8,p+8+n))}
   else if(r.type===0x00bd&&r.end-p>=12){const last=B16(r.end-2);for(let c=col;c<=last;c++){const pos=p+4+(c-col)*6;if(pos+6>r.end-2)break;const raw=B32(pos+2);let v;if(raw&2)v=raw>>2;else{const ab=new ArrayBuffer(8),dv2=new DataView(ab);dv2.setUint32(4,raw&0xfffffffc,true);v=dv2.getFloat64(0,true)}if(raw&1)v/=100;(rows[row]??=[])[c]=v}continue}
   if(val!==undefined)(rows[row]??=[])[col]=val;
  }
  sheets.push({name:sh.name,rows});
 }
 return sheets;
}
root.MaterialControlXLS={parse};
})(typeof window!=='undefined'?window:globalThis);
