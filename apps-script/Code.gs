/**
 * برج الأندلس — سكريبت البيانات المحمي (Google Apps Script)
 * ----------------------------------------------------------------
 * البرنامج (index.html) بيقرا ويحفظ من هنا.
 *
 *  - القراءة (GET ?key=data): مفتوحة، عشان السكان يشوفوا الحسابات. الرمز السري مش جوّه البيانات.
 *  - الحفظ / رفع المرفقات / تغيير الرمز: لازم الرمز السري الصح، والرمز محفوظ هنا بس
 *    (Script Properties) — مش في البرنامج ولا في البيانات.
 *  - بعد 10 محاولات غلط في الرمز: أي محاولة بتترفض لمدة 15 دقيقة.
 *  - لو جهاز تاني حفظ في النص (البيانات اتغيّرت من ورا الجهاز ده)، الحفظ بيترفض (CONFLICT)
 *    والبرنامج بيحدّث بدل ما يكتب فوق تعديل حد تاني.
 *  - لو البيانات الجديدة أصغر بكتير فجأة، الحفظ بيترفض إلا لو اتأكد (SHRINK_GUARD).
 *
 * التشغيل أول مرة: شغّل الدالة setup مرة واحدة من المحرر (Run)، وبعدين Deploy كـ Web app.
 * setup بتنقل البيانات والرمز السري الحالي من الرابط القديم لوحدها.
 */

// الرابط القديم (اللي البيانات عليه دلوقتي) — setup بتقرا منه مرة واحدة بس
var OLD_URL = 'https://script.google.com/macros/s/AKfycbz1_5VB0BEdJt6w0GiajPA2Rwdyx63CWSaYhZ2EYfu9EG038vZmsgDleAYxZrW1-Ch3aA/exec';

var CHUNK = 40000;            // الخلية في جوجل شيت بتشيل 50,000 حرف — بنقسّم البيانات على كذا خلية
var MAX_FAILS = 10;           // محاولات الرمز الغلط المسموحة
var LOCK_SECONDS = 15 * 60;   // مدة القفل بعدها

function props_(){ return PropertiesService.getScriptProperties(); }

/* ------------------------- التشغيل أول مرة ------------------------- */
function setup(){
  var p = props_();
  // ملف جوجل شيت اللي البيانات هتتحفظ فيه (بيتعمل لوحده في الدرايف بتاعك)
  if(!p.getProperty('SHEET_ID')){
    var ss = SpreadsheetApp.create('برج الأندلس — بيانات البرنامج (متلمسوش)');
    p.setProperty('SHEET_ID', ss.getId());
  }
  // فولدر المرفقات (فواتير/إيصالات)
  if(!p.getProperty('FOLDER_ID')){
    var folder = DriveApp.createFolder('برج الأندلس — مرفقات');
    p.setProperty('FOLDER_ID', folder.getId());
  }
  // نقل البيانات والرمز السري من الرابط القديم (لو لسه ما اتنقلتش)
  if(!readData_()){
    var res = UrlFetchApp.fetch(OLD_URL + '?key=data', {muteHttpExceptions: true, followRedirects: true});
    var body = JSON.parse(res.getContentText());
    if(!body || !body.value) throw new Error('الرابط القديم ما رجّعش بيانات — مفيش حاجة اتنقلت.');
    var state = JSON.parse(body.value);
    if(!state || !state.apartments || !state.apartments.length) throw new Error('البيانات القديمة مش سليمة — مفيش حاجة اتنقلت.');
    var pin = state.settings && state.settings.editPin;
    if(pin && !p.getProperty('ADMIN_PIN')) p.setProperty('ADMIN_PIN', String(pin));
    if(state.settings) delete state.settings.editPin;
    writeData_(JSON.stringify(state));
  }
  if(!p.getProperty('ADMIN_PIN')) p.setProperty('ADMIN_PIN', '1234');
  Logger.log('تمام ✅ — البيانات اتنقلت (' + readData_().length + ' حرف). اعمل Deploy دلوقتي.');
}

/* ------------------------- التخزين ------------------------- */
function sheet_(){
  var ss = SpreadsheetApp.openById(props_().getProperty('SHEET_ID'));
  var sh = ss.getSheetByName('data');
  if(!sh){ sh = ss.getSheets()[0]; sh.setName('data'); }
  return sh;
}
function readData_(){
  if(!props_().getProperty('SHEET_ID')) return '';
  var sh = sheet_();
  var n = sh.getLastRow();
  if(!n) return '';
  return sh.getRange(1, 1, n, 1).getValues().map(function(r){ return String(r[0]); }).join('');
}
function writeData_(raw){
  var sh = sheet_();
  var rows = [];
  for(var i = 0; i < raw.length; i += CHUNK) rows.push([raw.slice(i, i + CHUNK)]);
  sh.getRange('A:A').setNumberFormat('@');          // نص عادي، عشان الشيت مايغيّرش حاجة
  sh.clearContents();
  if(rows.length) sh.getRange(1, 1, rows.length, 1).setValues(rows);
  SpreadsheetApp.flush();
}
// نفس الدالة بالظبط اللي في البرنامج — بتطلّع "بصمة" للبيانات عشان نعرف لو حد غيّرها
function hash_(s){
  var h = 5381;
  for(var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36) + ':' + s.length;
}

/* ------------------------- الرمز السري ------------------------- */
function checkPin_(pin){
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pinFails') || 0);
  if(fails >= MAX_FAILS) return 'LOCKED';
  if(pin != null && String(pin) === props_().getProperty('ADMIN_PIN')) return 'OK';
  cache.put('pinFails', String(fails + 1), LOCK_SECONDS);
  return 'BAD';
}

/* ------------------------- الردود ------------------------- */
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  var key = e && e.parameter && e.parameter.key;
  if(key !== 'data') return json_({ok: false, error: 'مفيش حاجة هنا'});
  return json_({key: 'data', value: readData_() || null});
}

function doPost(e){
  var body;
  try{ body = JSON.parse(e.postData.contents); }catch(err){ return json_({ok: false, error: 'طلب مش مفهوم'}); }
  var auth = checkPin_(body.pin);
  if(auth === 'LOCKED') return json_({ok: false, code: 'LOCKED', error: 'محاولات كتير غلط — جرّب تاني بعد ربع ساعة'});
  if(auth !== 'OK') return json_({ok: false, code: 'AUTH', error: 'الرمز السري غلط'});

  var action = body.action || 'save';
  if(action === 'check') return json_({ok: true});

  if(action === 'setPin'){
    var np = String(body.newPin || '').trim();
    if(np.length < 4) return json_({ok: false, error: 'الرمز لازم يكون 4 حروف/أرقام على الأقل'});
    props_().setProperty('ADMIN_PIN', np);
    return json_({ok: true});
  }

  if(action === 'upload'){
    var folder = DriveApp.getFolderById(props_().getProperty('FOLDER_ID'));
    var blob = Utilities.newBlob(Utilities.base64Decode(body.data), body.type || 'application/octet-stream', body.name || 'file');
    var file = folder.createFile(blob);
    try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(err){}
    return json_({ok: true, id: file.getId(), url: file.getUrl()});
  }

  if(action === 'save'){
    if(body.key !== 'data' || typeof body.value !== 'string') return json_({ok: false, error: 'طلب حفظ ناقص'});
    var state;
    try{ state = JSON.parse(body.value); }catch(err){ return json_({ok: false, error: 'البيانات مش سليمة'}); }
    if(!state || !state.apartments || !state.apartments.length) return json_({ok: false, error: 'اتمنع حفظ بيانات ناقصة'});
    var raw = body.value;
    if(state.settings && state.settings.editPin !== undefined){  // الرمز مايتحفظش جوّه البيانات أبدًا
      delete state.settings.editPin;
      raw = JSON.stringify(state);
    }
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try{
      var current = readData_();
      if(!body.force){
        if(body.base && current && body.base !== hash_(current)){
          return json_({ok: false, code: 'CONFLICT', error: 'فيه تعديل اتحفظ من جهاز تاني'});
        }
        if(current && current.length > 2000 && raw.length < current.length * 0.6){
          return json_({ok: false, code: 'SHRINK_GUARD', oldLength: current.length, newLength: raw.length,
                        error: 'البيانات الجديدة أصغر بكتير من القديمة'});
        }
      }
      writeData_(raw);
      return json_({ok: true, hash: hash_(raw)});
    } finally {
      lock.releaseLock();
    }
  }
  return json_({ok: false, error: 'طلب مش معروف'});
}
