/**
 * Penerima MetalVerseFery: satu tab per materi, log lengkap, dan rekap siswa.
 * Tempel sebagai pengganti kode penerima pada project Apps Script Spreadsheet ini.
 * Perbarui penerapan Web App yang SUDAH ADA ke versi baru agar URL tetap sama.
 * Jalankan siapkanSheetMateri() sekali untuk membuat tab dan mengelompokkan log lama.
 */
var METALVERSE_SHEET_ID = '13VeZKXkum2BdIfBe4dkkjKoP_M5VP7siv793mHOe9Zs';
var METALVERSE_PROTOCOL = 3;
var METALVERSE_TOPICS = [
  'K3LH_Budaya_Industri','Pengukuran_Marking','Pemotongan_Mekanik',
  'Pemotongan_Termal','Pemotongan_Mekanik_Termal','Pembentukan_Brazing',
  'OAW','SMAW','GMAW','GTAW','WPS_Perakitan_QC','Asesmen_Akhir',
  'Plasma_Cutting','Las_Titik_RSW','Proyek_Serok_Sampah','Proyek_Toolbox','Tugas_Lainnya','Uji_Sistem'
];
var METALVERSE_HEADERS = ['Timestamp','No Absen','Nama Siswa','Kelas','Asal Sekolah','Kategori Aktivitas','Rincian Tugas & Hasil','Skor / Nilai','XP','Status','Topik Materi','ID Kiriman'];

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
function text_(value) { return value === null || value === undefined ? '' : String(value).trim(); }
function safeCell_(value) {
  if (typeof value === 'number') return value;
  var text = text_(value);
  // Student answers must remain text, including answers starting with '='.
  return /^[=+@]/.test(text) ? "'" + text : text;
}
function attendance_(value) {
  var text = text_(value);
  return /^\d+$/.test(text) ? String(Number(text)) : text;
}
function studentKey_(row) {
  return JSON.stringify([attendance_(row[1]),text_(row[2]).toLowerCase(),text_(row[3]).toLowerCase(),text_(row[4]).toLowerCase()]);
}
function inferTopic_(data) {
  var name = text_(data.nama).toLowerCase();
  var category = text_(data.kategori).toLowerCase();
  if (/\buji\s+(sistem|coba)\b|\btest\b|^tes_/.test(name) || /uji sistem|uji integrasi|koneksi tes/.test(category)) return 'Uji_Sistem';
  if (METALVERSE_TOPICS.indexOf(data.topik) !== -1) return data.topik;
  var romans = {I:1,II:2,III:3,IV:4,V:5,VI:6,VII:7,VIII:8,IX:9,X:10,XI:11};
  var chapterTopics = ['K3LH_Budaya_Industri','Pengukuran_Marking','Pemotongan_Mekanik','Pemotongan_Termal','Pembentukan_Brazing','OAW','SMAW','GMAW','GTAW','WPS_Perakitan_QC','Asesmen_Akhir'];
  var chapter = category.toUpperCase().match(/\bBAB\s+(XI|IX|VIII|VII|VI|IV|III|II|X|V|I|\d+)\b/);
  if (chapter) { var index = romans[chapter[1]] || Number(chapter[1]); if(index>=1 && index<=11)return chapterTopics[index-1]; }
  if (/plasma|\bpac\b/.test(category)) return 'Plasma_Cutting';
  if (/\brsw\b|spot welding|las titik/.test(category)) return 'Las_Titik_RSW';
  if (/pemotongan.*mekanik.*termal|pemotongan.*serok/.test(category)) return 'Pemotongan_Mekanik_Termal';
  if (/oxy|\bofc\b|pemotongan.*panas|pemotongan.*termal|gouging/.test(category)) return 'Pemotongan_Termal';
  if (/guillotine|pemotongan|shearing|potong/.test(category)) return 'Pemotongan_Mekanik';
  if (/gmaw|\bmig\b|mag/.test(category)) return 'GMAW';
  if (/gtaw|\btig\b|tungsten/.test(category)) return 'GTAW';
  if (/smaw|busur|arc master/.test(category)) return 'SMAW';
  if (/\boaw\b|oksi.asetilen/.test(category)) return 'OAW';
  if (/k3|safety|budaya industri|latihan stop/.test(category)) return 'K3LH_Budaya_Industri';
  if (/marking|pengukuran|bentangan/.test(category)) return 'Pengukuran_Marking';
  if (/bending|pembentukan|brazing|penekukan/.test(category)) return 'Pembentukan_Brazing';
  if (/toolbox|kotak perkakas/.test(category)) return 'Proyek_Toolbox';
  if (/serok|pengki/.test(category)) return 'Proyek_Serok_Sampah';
  if (/qc|cacat|defect|wps|perakitan/.test(category)) return 'WPS_Perakitan_QC';
  if (/asesmen akhir|bank soal/.test(category)) return 'Asesmen_Akhir';
  return 'Tugas_Lainnya';
}
function ensureSheet_(ss,name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (!sheet.getLastRow()) {
    sheet.getRange(1,1,1,METALVERSE_HEADERS.length).setValues([METALVERSE_HEADERS]);
    sheet.getRange(1,1,1,METALVERSE_HEADERS.length).setBackground('#568A7D').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    var currentHeaders=sheet.getRange(1,1,1,10).getValues()[0];
    if(currentHeaders.join('|')!==METALVERSE_HEADERS.slice(0,10).join('|'))throw Error('Tab '+name+' sudah berisi format lain. Ubah nama tab lama agar data lama tetap aman.');
    // Preserve the first ten established columns and every historical row.
    sheet.getRange(1,11,1,2).setValues([METALVERSE_HEADERS.slice(10)]);
  }
  return sheet;
}
function findId_(sheet,id) {
  if (!id || sheet.getLastRow()<2) return 0;
  var result = sheet.getRange(2,12,sheet.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();
  return result ? result.getRow() : 0;
}
function rekapStudent_(ss,log,userRow) {
  var headers = ['No Absen','Nama Siswa','Kelas','Asal Sekolah','Skor Guillotine (Potong)','Skor SMAW Arc (Las)','Skor QC Cacat (AWS)','Kuis Interaktif','Job Sheet / Tugas','Total XP','Status Terakhir','Waktu Update'];
  var sheet = ss.getSheetByName('Rekap_Nilai_Siswa') || ss.insertSheet('Rekap_Nilai_Siswa');
  if(!sheet.getLastRow()){sheet.getRange(1,1,1,headers.length).setValues([headers]);sheet.setFrozenRows(1);}
  var all = log.getDataRange().getValues().slice(1), key=studentKey_(userRow);
  var values=[userRow[1],userRow[2],userRow[3],userRow[4],'-','-','-','-','-',0,'',''];
  var seen={};
  all.forEach(function(row){
    if(studentKey_(row)!==key)return;
    var id=text_(row[11]); if(id && seen[id])return; if(id)seen[id]=true;
    var category=text_(row[5]).toLowerCase(),column=8;
    if(/kuis|quiz|asesmen|essay|esai/.test(category))column=7;
    else if(/qc|cacat|defect/.test(category))column=6;
    else if(/guillotine|pemotongan|potong|shear/.test(category))column=4;
    else if(/smaw|\barc\b/.test(category))column=5;
    if(row[7]!=='' && row[7]!=='-')values[column]=row[7];
    values[9]+=Number(row[8])||0;values[10]=row[9];values[11]=row[0];
  });
  var target=sheet.getLastRow()+1;
  if(sheet.getLastRow()>1){
    var identities=sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
    for(var i=0;i<identities.length;i++)if(studentKey_(['',identities[i][0],identities[i][1],identities[i][2],identities[i][3]])===key){target=i+2;break;}
  }
  sheet.getRange(target,1,1,12).setValues([values.map(safeCell_)]);
}
function doPost(e) {
  var lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))return json_({status:'error',message:'Penerima sedang sibuk. Kiriman dapat dicoba kembali.'});
  try {
    if(!e || !e.postData || !e.postData.contents)throw Error('Isi kiriman tidak ada.');
    var data=JSON.parse(e.postData.contents);
    if(!text_(data.nama)||!text_(data.noAbsen)||!text_(data.kelas)||!text_(data.asalSekolah)||!text_(data.kategori))throw Error('Identitas atau kategori tugas belum lengkap.');
    if(text_(data.rincianTugas).length>45000)throw Error('Jawaban terlalu panjang untuk satu sel; kirim sebagai beberapa bagian.');
    var ss=SpreadsheetApp.openById(METALVERSE_SHEET_ID),topic=inferTopic_(data);
    var id=text_(data.submissionId)||Utilities.getUuid();
    var row=[data.timestamp||Utilities.formatDate(new Date(),'Asia/Jakarta','dd/MM/yyyy HH:mm:ss'),data.noAbsen,data.nama,data.kelas,data.asalSekolah,data.kategori,data.rincianTugas||'',data.skor!==undefined?data.skor:'-',data.xp||0,data.status||'Selesai',topic,id];
    row=row.map(safeCell_);
    var log=ensureSheet_(ss,'Log_Aktivitas_Detail'),logRow=findId_(log,id);
    if(!logRow){log.appendRow(row);logRow=log.getLastRow();}
    var topicSheet=ensureSheet_(ss,topic),topicRow=findId_(topicSheet,id);
    if(!topicRow){topicSheet.appendRow(row);topicRow=topicSheet.getLastRow();}
    rekapStudent_(ss,log,row);
    SpreadsheetApp.flush();
    return json_({status:'success',protocolVersion:METALVERSE_PROTOCOL,spreadsheetId:METALVERSE_SHEET_ID,submissionId:id,topicSheet:topic,topicSheetId:topicSheet.getSheetId(),topicRow:topicRow,logRow:logRow,timestamp:new Date().toISOString()});
  } catch(error) {
    return json_({status:'error',message:error.message||String(error)});
  } finally {lock.releaseLock();}
}
function doGet(e) {
  try {
    var ss=SpreadsheetApp.openById(METALVERSE_SHEET_ID);
    if(e && e.parameter && e.parameter.action==='receipt'){
      var id=text_(e.parameter.id),log=ss.getSheetByName('Log_Aktivitas_Detail');
      var found=log?findId_(log,id):0;
      if(!found)return json_({status:'pending',submissionId:id});
      var topic=text_(log.getRange(found,11).getValue()),sheet=ss.getSheetByName(topic),row=sheet?findId_(sheet,id):0;
      return json_({status:row?'success':'pending',protocolVersion:METALVERSE_PROTOCOL,submissionId:id,spreadsheetId:METALVERSE_SHEET_ID,topicSheet:topic,topicRow:row,logRow:found});
    }
    return json_({status:'success',protocolVersion:METALVERSE_PROTOCOL,spreadsheetId:METALVERSE_SHEET_ID,topicRouting:true,message:'Penerima aktif. Setiap materi memiliki tab sendiri.'});
  }catch(error){return json_({status:'error',message:error.message||String(error)});}
}
function siapkanSheetMateri() {
  var lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    var ss=SpreadsheetApp.openById(METALVERSE_SHEET_ID);
    METALVERSE_TOPICS.forEach(function(topic){ensureSheet_(ss,topic);});
    var log=ensureSheet_(ss,'Log_Aktivitas_Detail');
    if(log.getLastRow()<2)return;
    var rows=log.getRange(2,1,log.getLastRow()-1,12).getValues();
    var groups={},ids={};
    METALVERSE_TOPICS.forEach(function(topic){
      groups[topic]=[];ids[topic]={};
      var target=ss.getSheetByName(topic);
      if(target.getLastRow()>1)target.getRange(2,12,target.getLastRow()-1,1).getValues().forEach(function(row){ids[topic][text_(row[0])]=true;});
    });
    var assignments=[];
    rows.forEach(function(row,index){
      var topic=inferTopic_({nama:row[2],kategori:row[5],topik:row[10]});
      var id=text_(row[11])||'historis-'+(index+2);
      row[10]=topic;row[11]=id;
      if(!ids[topic][id]){groups[topic].push(row.map(safeCell_));ids[topic][id]=true;}
      assignments.push([topic,id]);
    });
    METALVERSE_TOPICS.forEach(function(topic){var target=ss.getSheetByName(topic);if(groups[topic].length)target.getRange(target.getLastRow()+1,1,groups[topic].length,12).setValues(groups[topic]);});
    log.getRange(2,11,assignments.length,2).setValues(assignments);
    SpreadsheetApp.flush();
  }finally{lock.releaseLock();}
}
