'use strict';
// Conservative local corrections. Keep the original photo for review and
// leave already-clear, straight images untouched.
Object.assign(window.ImageRegions, (() => {
  function inspect(canvas) {
    const ratio = Math.min(1, 900 / Math.max(canvas.width, canvas.height));
    const sample = document.createElement('canvas');
    sample.width = Math.max(1, Math.round(canvas.width * ratio));
    sample.height = Math.max(1, Math.round(canvas.height * ratio));
    const context = sample.getContext('2d', {willReadFrequently:true});
    context.fillStyle = 'white';context.fillRect(0,0,sample.width,sample.height);
    context.drawImage(canvas,0,0,sample.width,sample.height);
    const pixels = context.getImageData(0,0,sample.width,sample.height).data;
    const gray = new Uint8Array(sample.width * sample.height), histogram = new Uint32Array(256);
    for (let i=0;i<gray.length;i++) {
      gray[i] = Math.round(.2126*pixels[4*i]+.7152*pixels[4*i+1]+.0722*pixels[4*i+2]);
      histogram[gray[i]]++;
    }
    function percentile(fraction) {
      let count=0;
      for(let v=0;v<256;v++) {count+=histogram[v];if(count>=gray.length*fraction)return v;}
      return 255;
    }
    const low=percentile(.005),high=percentile(.90),points=[];
    const threshold = high - Math.max(25,(high-low)*.45);
    for(let y=0;y<sample.height;y++)for(let x=0;x<sample.width;x+=2) {
      if(gray[y*sample.width+x]<threshold)points.push([x-sample.width/2,y-sample.height/2]);
    }
    const height = sample.height+Math.ceil(sample.width*.3)+8;
    const histogramRows=new Uint32Array(height);
    function score(degrees) {
      histogramRows.fill(0);
      const radians=degrees*Math.PI/180,c=Math.cos(radians),s=Math.sin(radians);
      for(const [x,y] of points)histogramRows[Math.round(y*c+x*s+height/2)]++;
      let result=0;for(const n of histogramRows)result+=n*n;
      return result;
    }
    let angle=0, improvement=1;
    if(high-low>=25 && points.length>=150 && points.length<gray.length*.18) {
      const baseline=score(0);let best=baseline;
      for(let candidate=-7;candidate<=7;candidate+=.5) {
        const value=score(candidate);if(value>best){best=value;angle=candidate;}
      }
      const coarse=angle;
      for(let tenth=-4;tenth<=4;tenth++) {
        const candidate=coarse+tenth/10;if(Math.abs(candidate)>7)continue;
        const value=score(candidate);if(value>best){best=value;angle=candidate;}
      }
      improvement=best/Math.max(1,baseline);
      if(Math.abs(angle)<.4 || Math.abs(angle)>6.8 || improvement<1.18)angle=0;
    }
    sample.width=sample.height=1;
    return {angle:Math.round(angle*10)/10,low,high,
      contrast:high-low>=25 && (high<235 || high-low<140),improvement};
  }
  function enhance(canvas) {
    const quality=inspect(canvas),changes=[];
    if(!quality.angle && !quality.contrast)return {canvas,angle:0,contrast:false,changes};
    let corrected=canvas;
    if(quality.contrast) {
      corrected=document.createElement('canvas');corrected.width=canvas.width;corrected.height=canvas.height;
      const context=corrected.getContext('2d',{willReadFrequently:true});context.drawImage(canvas,0,0);
      const pixels=context.getImageData(0,0,corrected.width,corrected.height),range=quality.high-quality.low;
      for(let i=0;i<pixels.data.length;i+=4) {
        const v=(.2126*pixels.data[i]+.7152*pixels.data[i+1]+.0722*pixels.data[i+2]-quality.low)*255/range;
        pixels.data[i]=pixels.data[i+1]=pixels.data[i+2]=Math.min(255,Math.max(0,v));
      }
      context.putImageData(pixels,0,0);changes.push('Đã tăng tương phản để đọc chữ');
    }
    if(quality.angle) {
      const radians=quality.angle*Math.PI/180,c=Math.abs(Math.cos(radians)),s=Math.abs(Math.sin(radians));
      const width=canvas.width*c+canvas.height*s+8,height=canvas.height*c+canvas.width*s+8;
      const scale=Math.min(1,3400/Math.max(width,height),Math.sqrt(9000000/(width*height)));
      const rotated=document.createElement('canvas');rotated.width=Math.ceil(width*scale);rotated.height=Math.ceil(height*scale);
      const context=rotated.getContext('2d');context.fillStyle='white';context.fillRect(0,0,rotated.width,rotated.height);
      context.translate(rotated.width/2,rotated.height/2);context.scale(scale,scale);context.rotate(radians);
      context.imageSmoothingQuality='high';context.drawImage(corrected,-corrected.width/2,-corrected.height/2);
      if(corrected!==canvas)corrected.width=corrected.height=1;
      corrected=rotated;changes.push(`Đã chỉnh nghiêng ${Math.abs(quality.angle).toLocaleString('vi-VN')}°`);
    }
    return {canvas:corrected,angle:quality.angle,contrast:quality.contrast,changes};
  }
  return {inspect,enhance};
})());
