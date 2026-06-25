/* ============================================================
   كود Apps Script — مخزن سحابى لتطبيق "سجل الترع المتعبة"
   ------------------------------------------------------------
   الوظيفة: استقبال بيانات الترع من التطبيق وحفظها فى جوجل شيت،
   وإرجاعها عند الطلب، مع فلترة بالمحافظة/الإدارة.

   طريقة التركيب (خطوة بخطوة فى ردى بعد الكود):
   1) افتح جوجل شيت جديد فاضى.
   2) Extensions ▸ Apps Script.
   3) امسح اللى فيه والصق الكود ده كامل.
   4) عدّل SHEET_ID و API_KEY بالقيم بتاعتك (موضّح تحت).
   5) Deploy ▸ New deployment ▸ Web app ▸ Execute as: Me ▸
      Who has access: Anyone ▸ Deploy ▸ انسخ رابط الـWeb app.
   6) حط الرابط فى إعدادات التطبيق.
   ============================================================ */

// ===== إعدادات لازم تعدّلها =====
var SHEET_ID = '';            // اتركه فاضى لو الكود داخل نفس الشيت، أو حط الـID
var SHEET_NAME = 'الترع';     // اسم التبويب اللى هيتخزّن فيه
var API_KEY = 'CHANGE_ME_2026'; // كلمة سر بسيطة تتطابق مع التطبيق (غيّرها)

// الأعمدة المقروءة (للعرض البشرى) + عمود data_json للسجل الكامل
var HEAD = ['id','الاسم','الإدارة','الهندسة','إدارة التوسع','الزمام','الطول كم','المتعب كم','المقترح','آخر تعديل','معدّل بواسطة','data_json'];

function getSheet_(){
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1,1,1,HEAD.length).setValues([HEAD]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,HEAD.length).setFontWeight('bold').setBackground('#0e3b5c').setFontColor('#ffffff');
  }
  return sh;
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* قراءة: GET ?key=...&admin=اسم الإدارة (admin اختيارى للفلترة) */
function doGet(e){
  try{
    var p = e.parameter || {};
    if(p.key !== API_KEY) return jsonOut_({ok:false, error:'مفتاح غير صحيح'});
    var sh = getSheet_();
    var rng = sh.getDataRange().getValues();
    if(rng.length < 2) return jsonOut_({ok:true, canals:[]});
    var head = rng[0];
    var jsonCol = head.indexOf('data_json');
    var adminCol = head.indexOf('الإدارة');
    var out = [];
    for(var i=1;i<rng.length;i++){
      var row = rng[i];
      if(!row[0]) continue; // لا id
      if(p.admin && String(row[adminCol]).trim() !== String(p.admin).trim()) continue;
      var rec = {};
      if(jsonCol>=0 && row[jsonCol]){ try{ rec = JSON.parse(row[jsonCol]); }catch(err){ rec = {}; } }
      rec.id = row[0];
      out.push(rec);
    }
    return jsonOut_({ok:true, canals:out, count:out.length});
  }catch(err){ return jsonOut_({ok:false, error:String(err)}); }
}

/* كتابة: POST {key, action:'upsert'|'delete', user, canals:[...] } */
function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var body = JSON.parse(e.postData.contents);
    if(body.key !== API_KEY) return jsonOut_({ok:false, error:'مفتاح غير صحيح'});
    var sh = getSheet_();
    var rng = sh.getDataRange().getValues();
    var head = rng[0];
    var idCol = 0;
    var idIndex = {};                       // id -> رقم الصف الفعلى
    for(var i=1;i<rng.length;i++){ if(rng[i][idCol]) idIndex[rng[i][idCol]] = i+1; }

    var user = body.user || '';
    var action = body.action || 'upsert';
    var canals = body.canals || [];
    var done = 0;

    canals.forEach(function(c){
      if(!c.id) return;
      if(action === 'delete'){
        if(idIndex[c.id]){ sh.deleteRow(idIndex[c.id]); 
          // إعادة بناء الفهرس بعد الحذف
          rng = sh.getDataRange().getValues(); idIndex = {};
          for(var k=1;k<rng.length;k++){ if(rng[k][idCol]) idIndex[rng[k][idCol]] = k+1; }
        }
        done++; return;
      }
      var rowVals = [
        c.id, c.f8_name||'', c.f8_admin||'', c.f8_eng||'', c.f8_expansion||'',
        c.f8_zam||'', c.f8_len||'', c.f8_tired_len||'', c.f8_proposal||'',
        new Date(), user, JSON.stringify(c)
      ];
      if(idIndex[c.id]){
        sh.getRange(idIndex[c.id], 1, 1, rowVals.length).setValues([rowVals]);
      } else {
        sh.appendRow(rowVals);
        idIndex[c.id] = sh.getLastRow();
      }
      done++;
    });
    return jsonOut_({ok:true, done:done, action:action});
  }catch(err){
    return jsonOut_({ok:false, error:String(err)});
  }finally{
    lock.releaseLock();
  }
}
