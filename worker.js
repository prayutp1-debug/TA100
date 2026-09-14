/* worker.js — อ่านไฟล์ Excel และคำนวณคะแนนในเธรดแยก เพื่อไม่ให้หน้าเว็บค้าง */
importScripts("vendor/xlsx.full.min.js", "engine.js");

var NEEDED = ["Punch", "จดทะเบียน", "การันตี", "เวลาทำงาน", "วันทำงาน", "คะแนน-Punch"];

function say(stage, pct){ self.postMessage({type:"progress", stage:stage, pct:pct}); }
function boom(msg){ self.postMessage({type:"error", message:msg}); }
function mb(n){ return (n/1048576).toFixed(1); }

function run(buf){
  try {
    say("กำลังเปิดไฟล์ Excel", .18);
    var wb = XLSX.read(new Uint8Array(buf), {type:"array", cellDates:false});
    buf = null;

    var missing = NEEDED.filter(function(n){ return wb.SheetNames.indexOf(n) < 0; });
    if (missing.length) return boom("ไฟล์ Excel ไม่มีชีต: " + missing.join(", "));

    say("กำลังอ่านข้อมูล", .3);
    var sheets = {};
    NEEDED.forEach(function(name){
      var ws = wb.Sheets[name];
      sheets[name] = (name === "เวลาทำงาน")
        ? XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null})   // หัวตาราง 2 ชั้น
        : XLSX.utils.sheet_to_json(ws, {raw:true, defval:null});
    });
    wb = null;

    var db = buildDB(sheets, function(t, p){ say(t, .35 + p * .6); });
    sheets = null;
    self.postMessage({type:"done", db:db});
  } catch (err){
    boom(err && err.message ? err.message : String(err));
  }
}

self.onmessage = function(e){
  var msg = e.data;
  if (msg.buffer) return run(msg.buffer);

  var xhr = new XMLHttpRequest();
  xhr.open("GET", msg.url, true);
  xhr.responseType = "arraybuffer";
  xhr.onprogress = function(ev){
    if (ev.lengthComputable)
      say("กำลังดาวน์โหลดไฟล์ Excel " + mb(ev.loaded) + " / " + mb(ev.total) + " MB",
          .02 + (ev.loaded / ev.total) * .14);
    else
      say("กำลังดาวน์โหลดไฟล์ Excel " + mb(ev.loaded) + " MB", .08);
  };
  xhr.onload = function(){
    if (xhr.status && xhr.status >= 400) return boom("โหลดไฟล์ไม่ได้ (HTTP " + xhr.status + ")");
    if (!xhr.response || !xhr.response.byteLength) return boom("ไฟล์ Excel ว่างหรือโหลดไม่สำเร็จ");
    run(xhr.response);
  };
  xhr.onerror = function(){ boom("เชื่อมต่อเพื่อโหลดไฟล์ Excel ไม่ได้"); };
  xhr.send(null);
};
