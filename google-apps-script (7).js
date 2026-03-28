const SHEET_ID = 'ضع_ID_الشيت_هنا';

function doGet(e) {
  const params   = e.parameter || {};
  const action   = params.action || '';
  const callback = params.callback || '';

  let result;
  try {
    if (action === 'ping') {
      result = {success: true, message: 'pong'};

    } else if (params.payload) {
      const payload = JSON.parse(decodeURIComponent(params.payload));
      if (payload.action === 'sync') {
        result = handleSync(payload.data);
      } else {
        result = {success: false, error: 'unknown action'};
      }

    } else if (action === 'get') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sh = ss.getSheetByName('_meta');
      if (!sh) { result = {success: false, error: 'no data'}; }
      else {
        const raw = sh.getRange('A1').getValue();
        result = raw ? {success: true, data: JSON.parse(raw)} : {success: false, error: 'empty'};
      }
    } else {
      result = {success: false, error: 'unknown action: ' + action};
    }
  } catch(err) {
    result = {success: false, error: err.toString()};
  }

  // JSONP — wraps response in callback to bypass CORS
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) { return doGet(e); }

function handleSync(data) {
  if (!data || !data.notes) return {success: false, error: 'incomplete data'};
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let meta = ss.getSheetByName('_meta') || ss.insertSheet('_meta');
  meta.getRange('A1').setValue(JSON.stringify(data));
  meta.getRange('A2').setValue('Last updated: ' + new Date().toLocaleString('ar-EG'));
  meta.hideSheet();
  writeTeamSummary(ss, data);
  (data.members||[]).forEach(m => {
    const name = typeof m==='string'?m:m.name;
    const role = typeof m==='string'?'mod':(m.role||'mod');
    writeMemberSheet(ss, data, name, role);
  });
  reorderSheets(ss, data.members||[]);
  return {success: true, count: data.notes.length};
}

function writeTeamSummary(ss, data) {
  let sh = ss.getSheetByName('📊 ملخص الفريق') || ss.insertSheet('📊 ملخص الفريق');
  sh.clearContents(); sh.clearFormats();
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.setRightToLeft(true);
  sh.getRange('A1:H1').merge();
  sh.getRange('A1').setValue('📊 ملخص فريق المخبر الاقتصادي — ' + new Date().toLocaleDateString('ar-EG'));
  sh.getRange('A1').setBackground('#1a1008').setFontColor('#e8b84a').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(1,44);
  sh.getRange('A2:H2').merge();
  sh.getRange('A2').setValue('إجمالي الملاحظات: ' + data.notes.length + '   |   الأعضاء: ' + (data.members||[]).length);
  sh.getRange('A2').setBackground('#221608').setFontColor('#887858').setFontSize(11).setHorizontalAlignment('center');
  const roleLabels = {mod:'مودريشن',design:'جرافيك',video:'فيديو'};
  const headers = ['العضو','الوظيفة','الملاحظات','متوسط التقييم','أعلى معيار','إيجابي','تحسين','آخر تقييم'];
  const hr = sh.getRange(4,1,1,headers.length);
  hr.setValues([headers]).setBackground('#b85840').setFontColor('#f2e8d8').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sh.setRowHeight(4,30);
  const members = data.members||[];
  const rows = members.map(m => {
    const name=typeof m==='string'?m:m.name;
    const role=typeof m==='string'?'mod':(m.role||'mod');
    const notes=data.notes.filter(n=>n.member===name);
    const scored=notes.filter(n=>n.stars>0);
    const avg=scored.length?Math.round(scored.reduce((a,n)=>a+n.stars,0)/scored.length*10)/10:0;
    const last=notes[0]?new Date(notes[0].date).toLocaleDateString('ar-EG'):'—';
    return [name,roleLabels[role]||role,notes.length,avg,getTopCriteria(notes),
            notes.filter(n=>n.type==='pos').length,notes.filter(n=>n.type==='neg').length,last];
  }).sort((a,b)=>b[3]-a[3]);
  if(rows.length){
    sh.getRange(5,1,rows.length,headers.length).setValues(rows).setHorizontalAlignment('center').setFontSize(12);
    rows.forEach((row,i)=>{
      const rr=sh.getRange(5+i,1,1,headers.length);
      rr.setBackground(i%2===0?'#1a1008':'#221608').setFontColor('#e8e0d4');
      rr.setBorder(true,true,true,true,true,true,'#3a2a10',SpreadsheetApp.BorderStyle.SOLID);
      const sc=row[3];const scC=sh.getRange(5+i,4);
      if(sc>=4.5){scC.setBackground('#1a3a1a');scC.setFontColor('#80b08a');scC.setFontWeight('bold');}
      else if(sc>=3.5){scC.setBackground('#0a2a3a');scC.setFontColor('#b8d8f0');scC.setFontWeight('bold');}
      else if(sc>=2.5){scC.setBackground('#3a2a0a');scC.setFontColor('#c8a050');scC.setFontWeight('bold');}
      else if(sc>0){scC.setBackground('#3a0a0a');scC.setFontColor('#cc7a60');scC.setFontWeight('bold');}
    });
    const cr=5+rows.length+2;
    sh.getRange(cr,1).setValue('العضو');sh.getRange(cr,2).setValue('المتوسط');
    rows.forEach((r,i)=>{sh.getRange(cr+1+i,1).setValue(r[0]);sh.getRange(cr+1+i,2).setValue(r[3]);});
    const chart=sh.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(cr,1,rows.length+1,2)).setPosition(5+rows.length+2,4,0,0)
      .setOption('title','مقارنة متوسط التقييم').setOption('titleTextStyle',{color:'#e8b84a',fontSize:13,bold:true})
      .setOption('backgroundColor',{fill:'#1a1008'}).setOption('legend',{position:'none'})
      .setOption('hAxis',{minValue:0,maxValue:5,textStyle:{color:'#887858'}})
      .setOption('vAxis',{textStyle:{color:'#e8e0d4'}}).setOption('colors',['#b85840'])
      .setOption('width',460).setOption('height',260).build();
    sh.insertChart(chart);
  }
  [140,100,100,120,180,80,80,110].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(4);
}

function writeMemberSheet(ss, data, member, role) {
  const shName='👤 '+member;
  let sh=ss.getSheetByName(shName)||ss.insertSheet(shName);
  sh.clearContents();sh.clearFormats();
  sh.getCharts().forEach(c=>sh.removeChart(c));
  sh.setRightToLeft(true);
  const roleColors={mod:'#b85840',design:'#5a9068',video:'#4878a0'};
  const roleLabels={mod:'مودريشن / Community Manager',design:'Graphic Designer',video:'Video Editor'};
  const rColor=roleColors[role]||'#c8902a';
  const notes=data.notes.filter(n=>n.member===member);
  const scored=notes.filter(n=>n.stars>0);
  const avg=scored.length?Math.round(scored.reduce((a,n)=>a+n.stars,0)/scored.length*10)/10:0;
  const pos=notes.filter(n=>n.type==='pos').length;
  const neg=notes.filter(n=>n.type==='neg').length;
  const neu=notes.filter(n=>n.type==='neu').length;
  sh.getRange('A1:J1').merge();
  sh.getRange('A1').setValue('👤 '+member+'  ·  '+(roleLabels[role]||role));
  sh.getRange('A1').setBackground('#1a1008').setFontColor(rColor).setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(1,42);
  sh.getRange('A2:J2').merge();
  sh.getRange('A2').setValue('إجمالي: '+notes.length+' ملاحظة  |  متوسط: '+(avg||'—')+'/5  |  '+new Date().toLocaleDateString('ar-EG'));
  sh.getRange('A2').setBackground('#221608').setFontColor('#887858').setFontSize(11).setHorizontalAlignment('center');
  const sr=4;
  [{v:avg?avg+'/5':'—',l:'متوسط التقييم',bg:avg>=4.5?'#1a3a1a':avg>=3.5?'#0a2a3a':avg>0?'#3a2a0a':'#221608',tc:avg>=4.5?'#80b08a':avg>=3.5?'#b8d8f0':avg>0?'#c8a050':'#887858'},
   {v:notes.length,l:'إجمالي الملاحظات',bg:'#2a1a08',tc:'#e8b84a'},
   {v:pos,l:'إيجابي 👍',bg:'#0a2a10',tc:'#80b08a'},
   {v:neg,l:'يحتاج تحسين ⚠️',bg:'#2a0a08',tc:'#cc7a60'},
   {v:neu,l:'ملاحظة 📌',bg:'#08182a',tc:'#b8d8f0'}
  ].forEach((s,i)=>{
    const col=1+i*2;
    sh.getRange(sr,col,1,2).merge().setValue(s.v).setBackground(s.bg).setFontColor(s.tc).setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(sr+1,col,1,2).merge().setValue(s.l).setBackground(s.bg).setFontColor('#887858').setFontSize(10).setHorizontalAlignment('center');
    sh.setRowHeight(sr,42);sh.setRowHeight(sr+1,22);
  });
  const cr=sr+3;
  sh.getRange(cr,1,1,5).merge().setValue('📋 تفصيل المعايير');
  sh.getRange(cr,1).setBackground('#2a1808').setFontColor('#e8b84a').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sh.setRowHeight(cr,32);
  sh.getRange(cr+1,1,1,5).setValues([['المعيار','عدد التقييمات','متوسط النجوم','إيجابي','يحتاج تحسين']]);
  sh.getRange(cr+1,1,1,5).setBackground('#b85840').setFontColor('#f2e8d8').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  const uCrits=[...new Set(notes.map(n=>n.critLabel))].filter(Boolean);
  const cData=uCrits.map(c=>{
    const cn=notes.filter(n=>n.critLabel===c);const cs=cn.filter(n=>n.stars>0);
    const ca=cs.length?Math.round(cs.reduce((a,n)=>a+n.stars,0)/cs.length*10)/10:0;
    return[c,cn.length,ca,cn.filter(n=>n.type==='pos').length,cn.filter(n=>n.type==='neg').length];
  });
  if(cData.length){
    sh.getRange(cr+2,1,cData.length,5).setValues(cData).setHorizontalAlignment('center').setFontSize(11);
    cData.forEach((row,i)=>{
      const rr=sh.getRange(cr+2+i,1,1,5);
      rr.setBackground(i%2===0?'#1a1008':'#221608').setFontColor('#e8e0d4');
      rr.setBorder(true,true,true,true,true,true,'#3a2a10',SpreadsheetApp.BorderStyle.SOLID);
      const sc=row[2];const scC=sh.getRange(cr+2+i,3);
      if(sc>=4.5){scC.setBackground('#1a3a1a');scC.setFontColor('#80b08a');scC.setFontWeight('bold');}
      else if(sc>=3.5){scC.setBackground('#0a2a3a');scC.setFontColor('#b8d8f0');scC.setFontWeight('bold');}
      else if(sc>=2.5){scC.setBackground('#3a2a0a');scC.setFontColor('#c8a050');scC.setFontWeight('bold');}
      else if(sc>0){scC.setBackground('#3a0a0a');scC.setFontColor('#cc7a60');scC.setFontWeight('bold');}
    });
    sh.getRange(cr+1,7).setValue('المعيار');sh.getRange(cr+1,8).setValue('المتوسط');
    cData.forEach((r,i)=>{sh.getRange(cr+2+i,7).setValue(r[0]);sh.getRange(cr+2+i,8).setValue(r[2]);});
    const bc=sh.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(cr+1,7,cData.length+1,2)).setPosition(cr+1,9,20,0)
      .setOption('title','تقييم المعايير').setOption('titleTextStyle',{color:'#e8b84a',fontSize:13,bold:true})
      .setOption('backgroundColor',{fill:'#1a1008'}).setOption('legend',{position:'none'})
      .setOption('hAxis',{minValue:0,maxValue:5,textStyle:{color:'#887858'}})
      .setOption('vAxis',{textStyle:{color:'#e8e0d4'}}).setOption('colors',[rColor])
      .setOption('width',420).setOption('height',Math.max(200,cData.length*46)).build();
    sh.insertChart(bc);
  }
  const nr=cr+cData.length+4;
  sh.getRange(nr,1,1,7).merge().setValue('🗒️ سجل الملاحظات');
  sh.getRange(nr,1).setBackground('#2a1808').setFontColor('#e8b84a').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sh.setRowHeight(nr,32);
  sh.getRange(nr+1,1,1,7).setValues([['التاريخ','الفترة','المعيار','النجوم','النوع','الملاحظة','الشهر']]);
  sh.getRange(nr+1,1,1,7).setBackground('#b85840').setFontColor('#f2e8d8').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  const pm={daily:'يومي',weekly:'أسبوعي',monthly:'شهري'};
  const tm={pos:'إيجابي 👍',neg:'تحسين ⚠️',neu:'ملاحظة 📌'};
  const tc={pos:'#0a2a10',neg:'#2a0808',neu:'#08182a'};
  if(notes.length){
    const nR=notes.map(n=>[n.date?new Date(n.date).toLocaleDateString('ar-EG'):'',pm[n.period]||n.period||'',n.critLabel||'',n.stars?'★'.repeat(n.stars):'—',tm[n.type]||n.type||'',n.text||'',n.monthKey||'']);
    sh.getRange(nr+2,1,nR.length,7).setValues(nR).setHorizontalAlignment('center').setFontSize(11);
    notes.forEach((n,i)=>{
      const rr=sh.getRange(nr+2+i,1,1,7);
      rr.setBackground(tc[n.type]||'#1a1008').setFontColor('#e8e0d4');
      rr.setBorder(true,true,true,true,true,true,'#3a2a10',SpreadsheetApp.BorderStyle.SOLID);
    });
    sh.getRange(nr+2,6,nR.length,1).setHorizontalAlignment('right');
  }
  if(notes.length>0){
    const pr=nr;const pc=9;
    sh.getRange(pr,pc).setValue('النوع');sh.getRange(pr,pc+1).setValue('العدد');
    sh.getRange(pr+1,pc).setValue('إيجابي');sh.getRange(pr+1,pc+1).setValue(pos);
    sh.getRange(pr+2,pc).setValue('تحسين');sh.getRange(pr+2,pc+1).setValue(neg);
    sh.getRange(pr+3,pc).setValue('ملاحظة');sh.getRange(pr+3,pc+1).setValue(neu);
    const pie=sh.newChart().setChartType(Charts.ChartType.PIE)
      .addRange(sh.getRange(pr,pc,4,2)).setPosition(pr,7,20,0)
      .setOption('title','توزيع الملاحظات').setOption('titleTextStyle',{color:'#e8b84a',fontSize:13,bold:true})
      .setOption('backgroundColor',{fill:'#1a1008'}).setOption('legend',{textStyle:{color:'#e8e0d4'}})
      .setOption('colors',['#5a9068','#b85840','#4878a0']).setOption('pieSliceBorderColor','#1a1008')
      .setOption('width',360).setOption('height',270).build();
    sh.insertChart(pie);
  }
  [120,90,160,90,110,260,90].forEach((w,i)=>sh.setColumnWidth(i+1,w));
  sh.setFrozenRows(2);
}

function reorderSheets(ss,members){
  try{
    const s=ss.getSheetByName('📊 ملخص الفريق');
    if(s){ss.setActiveSheet(s);ss.moveActiveSheet(1);}
    members.forEach((m,i)=>{
      const name=typeof m==='string'?m:m.name;
      const sh=ss.getSheetByName('👤 '+name);
      if(sh){ss.setActiveSheet(sh);ss.moveActiveSheet(i+2);}
    });
  }catch(e){}
}

function getTopCriteria(notes){
  if(!notes.length) return '—';
  const c={};
  notes.forEach(n=>{if(n.critLabel)c[n.critLabel]=(c[n.critLabel]||0)+1;});
  return Object.entries(c).sort((a,b)=>b[1]-a[1])[0]?.[0]||'—';
}

function out(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function testWrite(){
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sh=ss.getSheetByName('TEST')||ss.insertSheet('TEST');
  sh.getRange('A1').setValue('working — '+new Date());
}
