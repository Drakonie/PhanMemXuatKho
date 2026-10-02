'use strict';
// Image preparation is local to the browser. Coordinates always refer to the
// rotated source image, so TSV positions and the chosen reading region agree.
window.ImageRegions = (() => {
  const clamp = (value,low,high) => Math.min(high,Math.max(low,Number(value) || 0));
  function normalized(region) {
    const x = clamp(region?.x,0,.995), y = clamp(region?.y,0,.995);
    return {x,y,width:clamp(region?.width ?? 1,.005,1-x),height:clamp(region?.height ?? 1,.005,1-y)};
  }
  function pixels(region,width,height) {
    const value = normalized(region);
    const x = Math.min(width-1,Math.floor(value.x*width)),y = Math.min(height-1,Math.floor(value.y*height));
    return {x,y,width:Math.max(1,Math.min(width-x,Math.ceil(value.width*width))),height:Math.max(1,Math.min(height-y,Math.ceil(value.height*height)))};
  }
  function detect(canvas) {
    // Analyse a bounded thumbnail rather than allocating another full photo.
    const ratio = Math.min(1,1400/Math.max(canvas.width,canvas.height));
    const sample = document.createElement('canvas');sample.width=Math.max(1,Math.round(canvas.width*ratio));sample.height=Math.max(1,Math.round(canvas.height*ratio));
    const context = sample.getContext('2d',{willReadFrequently:true});context.drawImage(canvas,0,0,sample.width,sample.height);
    const {width,height} = sample, data = context.getImageData(0,0,width,height).data, ink = new Uint8Array(width*height);
    for (let index=0;index<ink.length;index++) {
      const offset=index*4;
      // Dark lettering survives without forcing yellow/cream cell fills to ink.
      ink[index] = data[offset+3]>30 && (.2126*data[offset]+.7152*data[offset+1]+.0722*data[offset+2])<220 ? 1 : 0;
    }
    const original = ink.slice();
    // Long spreadsheet rules are not data. Remove their pixels only in the
    // detection mask; the OCR image itself is never whitened or sharpened.
    const horizontal = Math.max(40,Math.round(width*.12)),vertical = Math.max(40,Math.round(height*.16));
    for (let y=0;y<height;y++) {
      let start=-1;
      for (let x=0;x<=width;x++) {
        if (x<width && original[y*width+x]) {if(start<0)start=x;}
        else if(start>=0) {if(x-start>=horizontal)ink.fill(0,y*width+start,y*width+x);start=-1;}
      }
    }
    for (let x=0;x<width;x++) {
      let start=-1;
      for (let y=0;y<=height;y++) {
        if(y<height && original[y*width+x]) {if(start<0)start=y;}
        else if(start>=0) {if(y-start>=vertical)for(let yy=start;yy<y;yy++)ink[yy*width+x]=0;start=-1;}
      }
    }
    let left=width,top=height,right=-1,bottom=-1,count=0;
    for(let y=0;y<height;y++)for(let x=0;x<width;x++)if(ink[y*width+x]) {left=Math.min(left,x);right=Math.max(right,x);top=Math.min(top,y);bottom=Math.max(bottom,y);count++;}
    sample.width=1;sample.height=1;
    if(count<12 || right<left || bottom<top)return {x:0,y:0,width:1,height:1,empty:true};
    // Keep a safe margin around every detected mark, including metadata above
    // the table. Internal blank cells never cause populated columns to vanish.
    const padding=Math.max(8,Math.round(Math.max(width,height)*.012));
    left=Math.max(0,left-padding);top=Math.max(0,top-padding);right=Math.min(width,right+padding+1);bottom=Math.min(height,bottom+padding+1);
    return {x:left/width,y:top/height,width:(right-left)/width,height:(bottom-top)/height,empty:false};
  }
  function prepare(canvas,region={mode:'auto'},upscale=true) {
    const chosen = region?.mode==='full' ? {x:0,y:0,width:1,height:1} : region?.mode==='manual' ? normalized(region) : detect(canvas);
    const bounds = pixels(chosen,canvas.width,canvas.height);
    // Small screenshots need larger characters. Large photos keep their native
    // resolution, bounded to 3400px and nine million pixels for browser memory.
    const edge=Math.max(bounds.width,bounds.height),wanted=upscale && Math.max(canvas.width,canvas.height)<1600?2.5:1;
    const scale=Math.min(wanted,3400/edge,Math.sqrt(9000000/(bounds.width*bounds.height)));
    const output=document.createElement('canvas');output.width=Math.max(1,Math.round(bounds.width*scale));output.height=Math.max(1,Math.round(bounds.height*scale));
    const context=output.getContext('2d');context.fillStyle='#fff';context.fillRect(0,0,output.width,output.height);context.imageSmoothingEnabled=true;context.imageSmoothingQuality='high';
    context.drawImage(canvas,bounds.x,bounds.y,bounds.width,bounds.height,0,0,output.width,output.height);
    return {canvas:output,region:{...bounds,mode:region?.mode || 'auto',source_width:canvas.width,source_height:canvas.height,scale_x:output.width/bounds.width,scale_y:output.height/bounds.height},empty:chosen.empty || false};
  }
  return {normalized,pixels,detect,prepare};
})();

// A spreadsheet screenshot can be read cell by cell without treating grid
// lines, empty cells or adjacent customer columns as characters in one line.
Object.assign(window.ImageRegions, (() => {
  const norm = text => String(text || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/đ/g,'d').replace(/[^a-z0-9]/g,'');
  function clusters(values) {
    const groups=[];values.forEach(value=>{if(groups.length && value-groups.at(-1).at(-1)<=2)groups.at(-1).push(value);else groups.push([value]);});
    return groups.map(group=>group.reduce((a,b)=>a+b,0)/group.length);
  }
  function grid(canvas) {
    const w=canvas.width,h=canvas.height;
    if(w<500 || h<180 || w*h>12000000)return null;
    const data=canvas.getContext('2d',{willReadFrequently:true}).getImageData(0,0,w,h).data;
    const dark=new Uint8Array(w*h),yhits=[],strength=[];
    for(let y=0;y<h;y++){let count=0;for(let x=0;x<w;x++){const i=(y*w+x)*4,r=data[i],g=data[i+1],b=data[i+2];dark[y*w+x]=(.2126*r+.7152*g+.0722*b<220 && Math.max(r,g,b)-Math.min(r,g,b)<70)?1:0;count+=dark[y*w+x];}strength[y]=count/w;if(count/w>.55)yhits.push(y);}
    const ys=clusters(yhits);if(ys.length<12)return null;
    let start=-1;
    for(let i=0;i<ys.length-8;i++){const gaps=ys.slice(i+1,i+9).map((v,j)=>v-ys[i+j]),sorted=[...gaps].sort((a,b)=>a-b),mid=sorted[4];if(mid>=8 && mid<=h*.08 && gaps.every(g=>g>mid*.72 && g<mid*1.28)){start=i;break;}}
    if(start<2)return null;
    const top=Math.ceil(ys[start]),bottom=Math.floor(ys.at(-1)),xhits=[],runRequired=Math.max(25,Math.round((bottom-top)*.10));
    for(let x=0;x<w;x++){let run=0,longest=0;for(let y=top;y<bottom;y++){run=dark[y*w+x]?run+1:0;longest=Math.max(longest,run);}if(longest>=runRequired)xhits.push(x);}
    let xs=clusters(xhits);if(xs.length<12 || xs.length>80)return null;
    if(xs[0]>3)xs.unshift(0);if(xs.at(-1)<w-3)xs.push(w-1);
    xs=xs.filter((v,i)=>!i || v-xs[i-1]>8);
    const strongest=Math.max(...strength);let end=ys.length-1;const ruleStrength=index=>Math.max(...strength.slice(Math.max(0,Math.floor(ys[index])-2),Math.min(h,Math.ceil(ys[index])+3)));let weak=0;for(let i=start+1;i<ys.length;i++){if(ruleStrength(i)<strongest*.94)weak++;else if(weak>=5){end=i;break;}}
    return {xs,ys,start,end,headerTop:ys[Math.max(0,start-3)],headerMiddle:ys[start-2],headerBottom:ys[start-1],bodyTop:ys[start]};
  }
  function cell(canvas,left,top,right,bottom,numeric=false) {
    const inset=Math.max(1.7,(bottom-top)*.12),x=Math.max(0,Math.ceil(left+inset)),y=Math.max(0,Math.ceil(top+inset)),w=Math.max(1,Math.floor(right-inset)-x),h=Math.max(1,Math.floor(bottom-inset)-y);
    const source=document.createElement('canvas');source.width=w;source.height=h;
    const ctx=source.getContext('2d',{willReadFrequently:true});ctx.drawImage(canvas,x,y,w,h,0,0,w,h);
    const pixels=ctx.getImageData(0,0,w,h);let ink=0,min=255,max=0;
    for(let i=0;i<pixels.data.length;i+=4){const v=.2126*pixels.data[i]+.7152*pixels.data[i+1]+.0722*pixels.data[i+2];min=Math.min(min,v);max=Math.max(max,v);if(v<170)ink++;}
    if(ink<3 || max-min<35){source.width=1;source.height=1;return null;}
    for(let i=0;i<pixels.data.length;i+=4){let v=.2126*pixels.data[i]+.7152*pixels.data[i+1]+.0722*pixels.data[i+2];v=(v-min)*255/Math.max(1,max-min);v=numeric?(v<155?0:255):Math.min(255,v);pixels.data[i]=pixels.data[i+1]=pixels.data[i+2]=v;}
    ctx.putImageData(pixels,0,0);
    const out=document.createElement('canvas'),scale=Math.min(6,Math.max(2,60/h),1800/Math.max(w,h)),pad=8;
    out.width=Math.ceil((w+2*pad)*scale);out.height=Math.ceil((h+2*pad)*scale);
    const target=out.getContext('2d');target.fillStyle='white';target.fillRect(0,0,out.width,out.height);target.imageSmoothingQuality='high';target.drawImage(source,pad*scale,pad*scale,w*scale,h*scale);source.width=1;source.height=1;
    return out;
  }
  async function readTable(canvas,recognize,onProgress=()=>{}) {
    const layout=grid(canvas);if(!layout)return null;
    const {xs,ys,start,end}=layout,totals=[],headers=[];
    async function read(l,t,r,b,numeric=false){const prepared=cell(canvas,l,t,r,b,numeric);if(!prepared)return {text:'',confidence:100};try{const data=await recognize(prepared,numeric);return {text:String(data.text||'').trim().replace(/\s+/g,' ') || '[OCR: chưa đọc được]',confidence:Number(data.confidence||0)};}finally{prepared.width=1;prepared.height=1;}}
    for(let col=0;col<xs.length-1;col++){
      const label=await read(xs[col],layout.headerBottom,xs[col+1],layout.bodyTop);
      headers.push(label.text);if(['tong','t0ng','total','long','t6ng'].includes(norm(label.text)))totals.push(col);
    }
    window.ImageRegions.lastScan={layout,headers,totals};if(totals.length<3)return null;
    const sharedTotal=totals[0],priceColumn=sharedTotal+1;
    const priceHeader=await read(xs[priceColumn],layout.headerTop,xs[priceColumn+1],layout.bodyTop);
    window.ImageRegions.lastScan.priceHeader=priceHeader;const priceInferred=!/dongi[ad]|donggia|unitprice|price/.test(norm(priceHeader.text));
    const candidates=totals.slice(1).filter(col=>col>=priceColumn+3 && col+2<xs.length);
    if(candidates.length<2 || candidates.length>30)return null;
    // Require the repeated NT/MG/Tổng/amount pattern; incomplete groups require
    // ordinary OCR review rather than exporting a subset as a complete table.
    if(candidates.some((col,i)=>i && col-candidates[i-1]!==4))return null;
    const widest=xs.slice(0,sharedTotal).map((x,i)=>({i,w:xs[i+1]-x})).sort((a,b)=>b.w-a.w)[0];
    if(!widest)return null;
    const nameColumn=widest.i,unitColumn=nameColumn+1,groups=[];
    for(let i=0;i<candidates.length;i++){
      const col=candidates[i],name=await read(xs[col-2],layout.headerMiddle,xs[col+2],layout.headerBottom);
      groups.push({name:name.text,index:i+1,quantity_column:col,amount_column:col+1,confidence:name.confidence});
    }
    const rows=[];
    for(let row=start;row<end;row++){
      onProgress(row-start,end-start);
      const name=await read(xs[nameColumn],ys[row],xs[nameColumn+1],ys[row+1]);
      if(/^(tongcong|cong|luyke|ton|tongtien|total|subtotal)/.test(norm(name.text)))break;
      if(!name.text)continue;
      const unit=await read(xs[unitColumn],ys[row],xs[unitColumn+1],ys[row+1]);
      const price=await read(xs[priceColumn],ys[row],xs[priceColumn+1],ys[row+1],true);
      const values=[];
      for(const group of groups){const q=group.quantity_column,a=group.amount_column;
        const quantity=await read(xs[q],ys[row],xs[q+1],ys[row+1],true),amount=await read(xs[a],ys[row],xs[a+1],ys[row+1],true);
        values.push({quantity:quantity.text,amount:amount.text,confidence:Math.min(quantity.confidence,amount.confidence)});
      }
      rows.push({name:name.text,unit:unit.text,price:price.text,confidence:Math.min(name.confidence,unit.confidence,price.confidence),values,source_row:row-start+1});
      if(rows.length>=200)throw new Error('Bảng ảnh quá dài; hãy chọn vùng nhỏ hơn.');
    }
    if(!rows.length)return null;
    return {groups,rows,price_inferred:priceInferred,header_inferred:headers.some(label=>['long','t6ng'].includes(norm(label))),method:'grid-cell-ocr',bbox:{left:xs[0],top:layout.headerTop,width:xs.at(-1)-xs[0],height:ys.at(-1)-layout.headerTop}};
  }
  return {grid,readTable,cell};
})());
