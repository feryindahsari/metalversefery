(function () {
  'use strict';
  window.metalverseSubmit = async function (payload, buttonId, silent) {
    var button = buttonId ? document.getElementById(buttonId) : null;
    if (button) { button.disabled=true; button.textContent='Mengirim ke Google Sheets...'; }
    try {
      var result=await window.MetalVerseSync.send(payload);
      if(button)button.textContent=result.status==='success'?'Tersimpan di Google Sheets':'Menunggu pengiriman — coba lagi';
      window.MetalVerseSync.feedback(result,null,silent);
      return result;
    } catch(error) {
      if(button)button.textContent='Belum terkirim — coba lagi';
      if(!silent)alert(error.message);
      console.warn('Pengiriman nilai:',error);
      return {status:'error',error:error.message};
    } finally { if(button)button.disabled=false; }
  };
  try {
    if (window.parent !== window && window.parent.location.origin === location.origin && window.parent.MetalVerseSync) {
      window.MetalVerseSync = window.parent.MetalVerseSync;
      return;
    }
  } catch (error) {}
  var ENDPOINT = 'https://script.google.com/macros/s/AKfycbxfZJ7pdx4BcJ9JjNkkX_KeQzz4ieFH65Xwz5DPMEYID41gYfgfzKoZr3U2Yg06_sM11Q/exec';
  var STORE = 'metalverse_delivery_v2';
  var memoryUser = null, memoryExpires = 0, active = {}, queue = [], durable = true;
  function readStore() {
    try { return JSON.parse(localStorage.getItem(STORE) || '[]'); }
    catch (error) { durable = false; return []; }
  }
  function merge() {
    var records = readStore();
    records.forEach(function (saved) {
      var current = queue.find(function (item) { return item.id === saved.id; });
      if (!current) queue.push(saved);
      else if (saved.state === 'sent') Object.assign(current, saved);
    });
  }
  function persist() {
    merge();
    // Keep every pending item; retain a bounded list of confirmed receipts.
    var sent = queue.filter(function (item) { return item.state === 'sent'; }).slice(-300);
    queue = queue.filter(function (item) { return item.state !== 'sent'; }).concat(sent);
    try { localStorage.setItem(STORE, JSON.stringify(queue)); durable = true; }
    catch (error) { durable = false; }
  }
  var TOPICS_BY_CHAPTER = { 'bab-1':'K3LH_Budaya_Industri','bab-2':'Pengukuran_Marking','bab-3':'Pemotongan_Mekanik','bab-4':'Pemotongan_Termal','bab-5':'Pembentukan_Brazing','bab-6':'OAW','bab-7':'SMAW','bab-8':'GMAW','bab-9':'GTAW','bab-10':'WPS_Perakitan_QC','bab-11':'Asesmen_Akhir' };
  var TOPICS_BY_JOB = {'jobsheet-1':'Pemotongan_Mekanik','jobsheet-2':'Pemotongan_Termal','jobsheet-3':'SMAW','jobsheet-4':'SMAW','jobsheet-5':'SMAW','jobsheet-6':'GMAW','jobsheet-7':'GTAW','jobsheet-8':'Proyek_Serok_Sampah','jobsheet-9':'Proyek_Toolbox'};
  function topicFor(title, id) {
    if(TOPICS_BY_CHAPTER[id])return TOPICS_BY_CHAPTER[id];
    if(TOPICS_BY_JOB[id])return TOPICS_BY_JOB[id];
    var s=String(title||'').toLowerCase();
    if(/plasma|\bpac\b/.test(s))return 'Plasma_Cutting';
    if(/\brsw\b|spot welding|las titik/.test(s))return 'Las_Titik_RSW';
    if(/pemotongan.*mekanik.*termal|serok.*pemotongan|pemotongan.*serok/.test(s))return 'Pemotongan_Mekanik_Termal';
    if(/oxy|ofc|pemotongan.*panas|pemotongan.*termal|gouging/.test(s))return 'Pemotongan_Termal';
    if(/guillotine|pemotongan|shearing|potong/.test(s))return 'Pemotongan_Mekanik';
    if(/gmaw|mig|mag/.test(s))return 'GMAW';
    if(/gtaw|\btig\b|tungsten/.test(s))return 'GTAW';
    if(/smaw|busur|arc master/.test(s))return 'SMAW';
    if(/\boaw\b|oksi.asetilen/.test(s))return 'OAW';
    if(/k3|safety|budaya industri|latihan stop/.test(s))return 'K3LH_Budaya_Industri';
    if(/marking|pengukuran|bentangan/.test(s))return 'Pengukuran_Marking';
    if(/bending|pembentukan|brazing|penekukan/.test(s))return 'Pembentukan_Brazing';
    if(/toolbox|kotak perkakas/.test(s))return 'Proyek_Toolbox';
    if(/serok|pengki/.test(s))return 'Proyek_Serok_Sampah';
    if(/qc|cacat|defect|wps|perakitan/.test(s))return 'WPS_Perakitan_QC';
    if(/asesmen akhir|bank soal/.test(s))return 'Asesmen_Akhir';
    return 'Tugas_Lainnya';
  }

  function validUser(user) {
    if (!user || typeof user.name !== 'string' || !user.name.trim()) return false;
    if (user.role === 'teacher') return user.teacherAccess === 'shared-password-v1';
    return user.role === 'student' && /^\d+$/.test(String(user.noAbsen)) && Number(user.noAbsen) > 0 &&
      ['kelas', 'asalSekolah'].every(function (key) { return typeof user[key] === 'string' && !!user[key].trim(); });
  }
  function getUser() {
    try {
      var saved = JSON.parse(sessionStorage.getItem('metalverse_identity_v1') || 'null');
      if (saved && saved.expiresAt > Date.now() && validUser(saved.user)) return saved.user;
      return !saved && memoryExpires>Date.now() && validUser(memoryUser) ? memoryUser : null;
    } catch (error) { return memoryExpires>Date.now() && validUser(memoryUser) ? memoryUser : null; }
  }
  function setIdentity(user) {
    if (!validUser(user)) throw Error('Lengkapi nama, absen, kelas, dan sekolah melalui halaman login.');
    memoryUser = user; memoryExpires=Date.now()+43200000;
    try { sessionStorage.setItem('metalverse_identity_v1', JSON.stringify({user:user, expiresAt:Date.now()+43200000})); } catch (error) {}
  }
  function requireUser() {
    var user = getUser();
    if (!user) throw Error('Identitas belum lengkap. Masuk melalui halaman utama MetalVerseFery sebelum mengirim pekerjaan.');
    return user;
  }
  function clearIdentity() { memoryUser = null; memoryExpires=0; try { sessionStorage.removeItem('metalverse_identity_v1'); } catch (error) {} }
  function storageKey(base) {
    var user = getUser();
    return base + ':' + (user ? encodeURIComponent(JSON.stringify([user.role,user.name,user.noAbsen,user.kelas,user.asalSekolah])) : 'belum-login');
  }
  function endpoint() {
    try { var custom = localStorage.getItem('metalverse_sheets_url'); if (custom && /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(custom)) return custom; } catch (error) {}
    return ENDPOINT;
  }
  function notify(record) {
    try { window.dispatchEvent(new CustomEvent('metalverse:submission', {detail:{payload:record.payload,state:record.state,id:record.id,receipt:record.receipt}})); } catch (error) {}
    paint(record.state === 'sent' ? 'Nilai tersimpan di Google Sheets.' : null);
  }
  function paint(message) {
    if (!document.body) return;
    var pending = queue.filter(function (item) { return item.state !== 'sent'; }).length;
    var box = document.getElementById('metalverse-delivery-status');
    if (!box && (pending || message)) {
      box = document.createElement('div'); box.id = 'metalverse-delivery-status'; box.setAttribute('role','status');
      box.style.cssText='position:fixed;bottom:12px;left:12px;right:12px;max-width:470px;z-index:20000;padding:10px 14px;border-radius:12px;font:14px/1.4 system-ui;box-shadow:0 3px 16px #0002';
      document.body.appendChild(box);
    }
    if (!box) return;
    box.style.background = pending ? '#fff3cd' : '#dcfce7'; box.style.color = pending ? '#713f12' : '#14532d';
    box.textContent = pending ? pending+' kiriman menunggu Google Sheets. '+(durable ? 'Tersimpan di perangkat; akan dicoba lagi.' : 'Biarkan halaman ini terbuka; penyimpanan perangkat diblokir.') : (message || 'Semua kiriman sudah diterima Google Sheets.');
    if (pending) { var retry=document.createElement('button'); retry.textContent='Coba kirim lagi'; retry.style.cssText='margin-left:10px;text-decoration:underline;cursor:pointer'; retry.onclick=flush; box.appendChild(retry); }
    else setTimeout(function(){if(box.parentNode && !queue.some(function(x){return x.state!=='sent';}))box.remove();},8000);
  }
  function deliver(item) {
    if (item.state === 'sent') return Promise.resolve({status:'success',id:item.id,receipt:item.receipt});
    if (navigator.onLine === false) return Promise.resolve({status:'pending',id:item.id,durable:durable});
    if (active[item.id]) return active[item.id];
    active[item.id] = (async function () {
      item.attempts = (item.attempts || 0) + 1;
      var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
      var timeout = controller ? setTimeout(function(){controller.abort();},45000) : null;
      try {
        var response=await fetch(item.url,{method:'POST',mode:'cors',credentials:'omit',redirect:'follow',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(item.payload),signal:controller ? controller.signal : undefined});
        if (!response.ok || response.type === 'opaque') throw Error('Google belum mengonfirmasi penerimaan.');
        var receipt=await response.json();
        if (!receipt || receipt.status !== 'success') throw Error(receipt && receipt.message || 'Google Sheets belum menerima kiriman.');
        if (receipt.submissionId && receipt.submissionId!==item.id) throw Error('Balasan tidak cocok dengan kiriman ini.');
        if (receipt.spreadsheetId && receipt.spreadsheetId!=='13VeZKXkum2BdIfBe4dkkjKoP_M5VP7siv793mHOe9Zs') throw Error('Alamat penerima mengarah ke Spreadsheet lain.');
        item.state='sent'; item.receipt=receipt; item.confirmedAt=new Date().toISOString(); delete item.error;
        persist(); notify(item);
        return {status:'success',id:item.id,receipt:receipt};
      } catch(error) {
        item.state='pending'; item.error=error.message; persist(); notify(item);
        return {status:'pending',id:item.id,durable:durable,error:error.message};
      } finally { if(timeout)clearTimeout(timeout); delete active[item.id]; }
    })();
    return active[item.id];
  }
  async function send(data) {
    var user=requireUser();
    var payload=Object.assign({},data,{nama:user.name,noAbsen:user.noAbsen,kelas:user.kelas,asalSekolah:user.asalSekolah,role:user.role});
    payload.timestamp=payload.timestamp || new Date().toISOString();
    payload.topik=payload.topik || topicFor(payload.kategori);
    payload.jenisTugas=payload.jenisTugas || payload.kategori;
    // The deployed receiver stores rincianTugas; include the complete answers there.
    if (payload.catatan) payload.rincianTugas=(payload.rincianTugas||'')+'\n'+payload.catatan;
    var text=payload.rincianTugas || '';
    var parts=Math.max(1,Math.ceil(text.length/40000)), last;
    for(var i=0;i<parts;i++) {
      var part=Object.assign({},payload,{rincianTugas:parts>1?'Bagian '+(i+1)+'/'+parts+'\n'+text.slice(i*40000,(i+1)*40000):text});
      if(i>0){part.skor='-';part.xp=0;}
      var key=JSON.stringify([user.role,user.name,user.noAbsen,user.kelas,user.asalSekolah,part.kategori,part.rincianTugas,part.skor,part.xp]);
      merge();
      var item=queue.find(function(entry){return entry.key===key;});
      if(!item) {
        var id='mv-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2);
        part.submissionId=id;
        item={id:id,key:key,payload:part,state:'pending',url:endpoint(),createdAt:new Date().toISOString()};
        queue.push(item);persist();notify(item);
      }
      var result=await deliver(item);
      if(!last || result.status!=='success')last=result;
    }
    return last;
  }
  async function flush() {
    merge();
    var pending=queue.filter(function(item){return item.state!=='sent';});
    for(var i=0;i<pending.length;i++) { if(navigator.onLine===false)break; var result=await deliver(pending[i]); if(result.status!=='success')break; }
    paint();
  }
  function feedback(result,buttonId,silent) {
    var ok=result.status==='success';
    var message=ok?('Tersimpan di Google Sheets'+(result.receipt && result.receipt.topicSheet?' — tab '+result.receipt.topicSheet:'')+'.'):'Belum diterima Google Sheets. '+(result.durable?'Kiriman disimpan di perangkat dan akan dicoba lagi.':'Biarkan halaman ini terbuka agar kiriman dapat dicoba lagi.');
    if(buttonId){var button=document.getElementById(buttonId);if(button){button.disabled=false;button.textContent=ok?'Tersimpan di Google Sheets':'Menunggu pengiriman — coba lagi';}}
    if(!silent)alert(message);
    return result;
  }
  queue=readStore();
  window.MetalVerseSync={send:send,flush:flush,getUser:getUser,requireUser:requireUser,setIdentity:setIdentity,clearIdentity:clearIdentity,storageKey:storageKey,topicFor:topicFor,feedback:feedback,
    readUser:function(){var u=getUser();return u?JSON.stringify(u):null;},
    saveUser:function(raw){setIdentity(JSON.parse(raw));},
    records:function(){merge();return JSON.parse(JSON.stringify(queue));}};
  window.addEventListener('online',flush);
  window.addEventListener('storage',function(event){if(event.key===STORE){merge();paint();}});
  window.addEventListener('DOMContentLoaded',function(){paint();if(queue.some(function(x){return x.state!=='sent';}))flush();});
  setInterval(function(){if(queue.some(function(x){return x.state!=='sent';}))flush();},30000);
})();
