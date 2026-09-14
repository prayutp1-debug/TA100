/* engine.js — คำนวณคะแนน TA จากไฟล์ TA100.xlsx (ใช้ร่วมกันระหว่าง worker และหน้าเว็บ)
   แก้เกณฑ์ที่ไฟล์นี้ที่เดียว แล้วเพิ่มเลข ENGINE_VERSION เพื่อล้าง cache */

var ENGINE_VERSION = 3;

var TA_TIERS = [
  {name:"TA 1", min:95, inc: 6, gpct:1.000},
  {name:"TA 2", min:90, inc: 4, gpct:0.965},
  {name:"TA 3", min:85, inc: 2, gpct:0.930},
  {name:"TA 4", min:75, inc: 0, gpct:0},
  {name:"TA 5", min:50, inc:-3, gpct:0},
  {name:"TA 6", min:-1, inc:-6, gpct:0}
];
var LATE_GRACE_MIN = 30;   // check_time ต้องไม่เกินเวลาเข้างานโรงงาน + 30 นาที
var REST_HOURS     = 13;   // ดีพีใบสุดท้ายวันก่อน + 13 ชม. = สิทธิ์พักของคนขับ
var MONTH_EN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

/* ── ตัวช่วยอ่านค่าจากเซลล์ ─────────────────────────────────── */
function S(v){ return v==null ? "" : String(v).trim(); }

function findCol(keys, part){
  for (var i=0;i<keys.length;i++) if (keys[i].indexOf(part) !== -1) return keys[i];
  return null;
}

/** serial ของ Excel -> {y, m, d} (คิดแบบ UTC จึงไม่เพี้ยนตาม timezone) */
function excelDate(v){
  if (v instanceof Date) return {y:v.getFullYear(), m:v.getMonth()+1, d:v.getDate()};
  if (typeof v === "number"){
    var ms = Date.UTC(1899,11,30) + Math.round(v) * 86400000;
    var t = new Date(ms);
    return {y:t.getUTCFullYear(), m:t.getUTCMonth()+1, d:t.getUTCDate()};
  }
  var s = S(v), m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return {y:+m[1], m:+m[2], d:+m[3]};
  var t2 = new Date(s);
  return isNaN(t2) ? null : {y:t2.getFullYear(), m:t2.getMonth()+1, d:t2.getDate()};
}

/** ค่าเวลา -> จำนวนวินาทีนับจากเที่ยงคืน (null ถ้าไม่ใช่เวลา) */
function timeSecs(v){
  if (v == null || v === "") return null;
  if (v instanceof Date) return v.getHours()*3600 + v.getMinutes()*60 + v.getSeconds();
  if (typeof v === "number"){
    var f = v % 1; if (f < 0) f += 1;
    // ปัดเป็นมิลลิวินาทีก่อน แล้วตัดเศษวินาที — กันคลาดเคลื่อนจากทศนิยมของ Excel
    return Math.floor(Math.round(f * 86400 * 1000) / 1000);
  }
  var s = S(v);
  if (!s || s === "ไม่มีดีพี") return null;
  var m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  return m ? (+m[1])*3600 + (+m[2])*60 + (+(m[3]||0)) : null;
}

function hhmm(secs){
  var h = Math.floor(secs/3600) % 24, m = Math.floor(secs/60) % 60;
  return (h<10?"0":"")+h + ":" + (m<10?"0":"")+m;
}

/* ── ตัวคำนวณหลัก ──────────────────────────────────────────── */
/**
 * @param sheets  { 'Punch':[{...}], 'จดทะเบียน':[...], ... }  (แต่ละชีตเป็น array ของ object)
 *                ยกเว้น 'เวลาทำงาน' ที่ส่งมาเป็น array ของ array (เพราะหัวตาราง 2 ชั้น)
 * @param report  ฟังก์ชันรายงานความคืบหน้า (ไม่บังคับ)
 */
function buildDB(sheets, report){
  report = report || function(){};
  var punch = sheets["Punch"], reg = sheets["จดทะเบียน"], gar = sheets["การันตี"],
      wtimeRows = sheets["เวลาทำงาน"], wdays = sheets["วันทำงาน"], rulesRaw = sheets["คะแนน-Punch"];

  if (!punch || !punch.length) throw new Error("ไม่พบข้อมูลในชีต Punch");

  var pk = Object.keys(punch[0]);
  var C_PREV = findCol(pk, "ดีพีใบสุดท้ายวันก่อน");

  /* เดือน/ปีของข้อมูล — ใช้ค่าที่พบบ่อยที่สุดใน work_date */
  var ymCount = {}, dates = new Array(punch.length);
  for (var i=0;i<punch.length;i++){
    var dt = excelDate(punch[i].work_date);
    dates[i] = dt;
    if (dt){ var k = dt.y+"-"+dt.m; ymCount[k] = (ymCount[k]||0)+1; }
  }
  var best = null;
  for (var k2 in ymCount) if (!best || ymCount[k2] > ymCount[best]) best = k2;
  if (!best) throw new Error("อ่านคอลัมน์ work_date ไม่ได้");
  var YEAR = +best.split("-")[0], MONTH = +best.split("-")[1];
  var monthEn = MONTH_EN[MONTH-1];
  var daysInMonth = new Date(Date.UTC(YEAR, MONTH, 0)).getUTCDate();

  /* วันทำงานของเดือน */
  var workDays = daysInMonth, workNote = "";
  (wdays||[]).forEach(function(r){
    if (+r["ปี"] === YEAR && S(r["เดือน"]) === monthEn){
      workDays = +r["วันทำงาน"] || daysInMonth;
      workNote = S(r["หมายเหตุ"]);
    }
  });

  /* ── ตารางเกณฑ์คะแนน ── */
  var SC = {punch_in_status:{}, punch_out_status:{}, device_no:{}, check_time:{}};
  var FINE = {}, FINE_ALL = {}, MEANING = {}, rules = [];
  (rulesRaw||[]).forEach(function(r){
    var grp = S(r["In/Out"]), code = S(r["รหัส"]);
    if (!grp || !code) return;
    var pts = parseFloat(r["คะแนนรถ"]); if (isNaN(pts)) pts = 0;
    var fine = S(r["หัก จบส.(บาท)"]);
    rules.push({grp:grp, code:code, meaning:S(r["ความหมายจากระบบ"]), pts:pts, fine:fine});
    if (code === "XX") return;                 // เกณฑ์ที่บันทึกเพิ่มเอง ไม่มีในข้อมูล Punch
    if (!SC[grp]) SC[grp] = {};
    SC[grp][code] = pts < 0 ? pts : 0;         // คะแนน 1 = ไม่หัก, ค่าลบ = จำนวนที่หัก
    MEANING[grp+"|"+code] = S(r["ความหมายจากระบบ"]);
    if (grp === "punch_out_status" && fine){
      if (fine === "ตัดทั้งโครงการ") FINE_ALL[code] = 1;
      else { var f = parseFloat(fine.replace(/,/g,"")); if (!isNaN(f)) FINE[code] = f; }
    }
  });
  var DED_NO_CHECK = SC.check_time["Blanks"] != null ? SC.check_time["Blanks"] : -1;
  var DED_LATE     = SC.check_time[""] != null ? SC.check_time[""] : -0.25;
  var DED_OUT_BLANK= SC.punch_out_status["Blanks"] != null ? SC.punch_out_status["Blanks"] : -1;

  /* ── เวลาเข้างานของแต่ละโรงงาน (หัวตาราง 2 ชั้น) ── */
  var plantIn = {};
  if (wtimeRows && wtimeRows.length){
    var head = wtimeRows[0].map(S);
    var iCode = head.indexOf("รหัส"), iIn = head.indexOf("Now");
    if (iCode < 0) iCode = 3;
    if (iIn   < 0) iIn   = 5;
    for (var w=2; w<wtimeRows.length; w++){
      var row = wtimeRows[w]; if (!row) continue;
      var code = S(row[iCode]); if (!code) continue;
      var t = timeSecs(row[iIn]);
      plantIn[code] = (t == null) ? 8*3600 : t;
    }
  }

  /* ── ยอดการันตีรายโรงงาน ── */
  var garantee = {};
  if (gar && gar.length){
    var gk = Object.keys(gar[0]);
    var cOld = findCol(gk, "รถเก่า"), cNew = findCol(gk, "FLEET_TYPE_NO");
    gar.forEach(function(r){
      var code = S(r["รหัส"]); if (!code) return;
      garantee[code] = {old:+(r[cOld]||0), t012:+(r[cNew]||0)};
    });
  }

  /* ── ทะเบียนรถ ── */
  var REG = {}, reportYM = YEAR*12 + MONTH;
  (reg||[]).forEach(function(r){
    var t = S(r["เบอร์รถ"]); if (!t) return;
    var gt = excelDate(r["วันหมดอายุ GT รถใหม่"]);
    REG[t] = {
      division:S(r["กิจการ"]), dept:S(r["ภาค"]), sect:S(r["แผนก"]),
      plant_no:S(r["รหัสโรงงาน"]), plant:S(r["โรงงาน"]),
      vendor:S(r["ผู้รับเหมา"]), vendor_id:S(r["Vendor"]),
      fleet:S(r["FLEET_TYPE_NO"]), size:S(r["ขนาดรถ"]),
      gt: gt ? (gt.y+"-"+(gt.m<10?"0":"")+gt.m+"-"+(gt.d<10?"0":"")+gt.d) : null,
      gt_ym: gt ? gt.y*12 + gt.m : null
    };
  });

  /* ── หา driver ที่ Punch เกิน 1 คันในวันเดียวกัน ── */
  report("ตรวจ driver ที่ Punch หลายคัน", .35);
  var seen = {}, multi = {};
  for (var a=0; a<punch.length; a++){
    var dt0 = dates[a]; if (!dt0) continue;
    var key = S(punch[a].driver_no) + "|" + dt0.d;
    var tno = S(punch[a].truck_no);
    if (!seen[key]) seen[key] = {};
    seen[key][tno] = 1;
  }
  for (var kk in seen){
    var n = 0; for (var z in seen[kk]) n++;
    if (n > 1) multi[kk] = n;
  }

  /* ── คำนวณรายแถว ── */
  report("คำนวณคะแนนรายรายการ", .45);
  var detail = new Array(punch.length), dayAgg = {}, fineRows = {}, drvInfo = {};

  for (var r0=0; r0<punch.length; r0++){
    var p = punch[r0], dtt = dates[r0];
    if (!dtt) { detail[r0] = null; continue; }
    var day = dtt.d;
    var truck = S(p.truck_no), drv = S(p.driver_no);
    var pin = S(p.punch_in_status), pout = S(p.punch_out_status), dev = S(p.device_no);
    var plantNo = S(p.plant_no);

    var dIn  = SC.punch_in_status[pin]  || 0;
    var dDev = SC.device_no[dev]        || 0;
    var dOut = (pout === "") ? DED_OUT_BLANK : (SC.punch_out_status[pout] || 0);

    /* เส้นตายตรวจรถ = เวลาที่ช้ากว่าระหว่าง (เข้างาน+30 น.) กับ (ดีพีวันก่อน+13 ชม.) */
    var dCt = 0, ctStat = "on_time", note = "";
    var ct = timeSecs(p.check_time);
    if (ct == null){
      dCt = DED_NO_CHECK; ctStat = "no_check"; note = "ไม่ตรวจรถทั้งวัน";
    } else {
      var shiftDl = day*86400 + (plantIn[plantNo] != null ? plantIn[plantNo] : 8*3600) + LATE_GRACE_MIN*60;
      var deadline = shiftDl, why = "shift";
      var prev = C_PREV ? timeSecs(p[C_PREV]) : null;
      if (prev != null){
        var restDl = (day-1)*86400 + prev + REST_HOURS*3600;
        if (restDl > deadline){ deadline = restDl; why = "rest"; }
      }
      // เทียบระดับนาที เพราะ check_time มีความละเอียดแค่นาที
      if (Math.floor((day*86400 + ct)/60) > Math.floor(deadline/60)){
        dCt = DED_LATE;
        if (why === "rest"){
          ctStat = "late_rest";
          note = "พักครบ " + REST_HOURS + " ชม. เวลา " + hhmm(deadline % 86400) +
                 " แต่ตรวจรถ " + hhmm(ct);
        } else {
          ctStat = "late_shift";
          note = "ตรวจรถเกิน 30 นาทีไม่ทันเวลาเข้างาน";
        }
      }
    }

    var pre = 1 + dIn + dDev + dOut + dCt;
    var net = Math.min(1, Math.max(0, pre));

    var rmk = [];
    var nMulti = multi[drv + "|" + day];
    if (nMulti) rmk.push("Driver Punch " + nMulti + " คัน");
    if (note) rmk.push(note);
    if (pout && FINE_ALL[pout]) rmk.push("ตัดทั้งโครงการ");

    detail[r0] = [day, plantNo, truck, drv, dev, pin, pout, ctStat,
      ct == null ? "" : hhmm(ct),
      +dIn.toFixed(2), +dDev.toFixed(2), +dOut.toFixed(2), +dCt.toFixed(2),
      +Math.max(0, Math.min(1, pre)).toFixed(2), +net.toFixed(2), rmk.join(" / ")];

    /* รวมเป็นรายคัน-รายวัน (หลายแถวใน 1 วัน = รวมยอดหัก แล้วตัดที่ -1) */
    var ak = truck + "|" + day, agg = dayAgg[ak];
    if (!agg) agg = dayAgg[ak] = {ded:0, rmk:[], qty:0, cost:0};
    agg.ded += dIn + dDev + dOut + dCt;
    agg.qty += (+p.dp_dispatch_qty || 0);
    agg.cost += (+p.partialtcost || 0);
    for (var m3=0; m3<rmk.length; m3++) if (agg.rmk.indexOf(rmk[m3]) < 0) agg.rmk.push(rmk[m3]);

    /* หักเงิน จบส. */
    if (FINE[pout] > 0){
      if (!fineRows[drv]) fineRows[drv] = {};
      fineRows[drv][day] = (fineRows[drv][day] || 0) + FINE[pout];
    }
    if (!drvInfo[drv]) drvInfo[drv] = {
      division:S(p.division_name), dept:S(p.dept_name), sect:S(p.sect_name),
      plant_no:plantNo, plant:S(p.plant_name), vendor_id:S(p.vendor_id), vendor:S(p.vendor_name)
    };
  }
  detail = detail.filter(Boolean);

  /* ── สรุปรายคัน ── */
  report("สรุปคะแนนรายคัน", .8);
  var trucksMap = {};
  for (var ak2 in dayAgg){
    var parts = ak2.split("|"), tn = parts[0], dd = +parts[1], ag = dayAgg[ak2];
    var t = trucksMap[tn];
    if (!t) t = trucksMap[tn] = {days:{}, rmk:{}, qty:0, cost:0};
    t.days[dd] = +Math.min(1, Math.max(0, 1 + ag.ded)).toFixed(2);
    t.qty += ag.qty; t.cost += ag.cost;
    if (ag.rmk.length) t.rmk[dd] = ag.rmk.join(" / ");
  }

  function tierOf(pct){
    for (var i=0;i<TA_TIERS.length;i++){
      var t = TA_TIERS[i];
      if (i === 0 ? pct > t.min : pct >= t.min) return t;
    }
    return TA_TIERS[TA_TIERS.length-1];
  }

  var trucks = Object.keys(trucksMap).sort().map(function(tn){
    var t = trucksMap[tn], info = REG[tn] || {};
    var score = 0; for (var d in t.days) score += t.days[d];
    score = +score.toFixed(2);
    var pct = workDays ? +(score / workDays * 100).toFixed(2) : 0;
    var tier = tierOf(pct);

    var g = garantee[info.plant_no || ""] || {old:0, t012:0};
    var isT012 = info.fleet === "T012" && info.gt_ym != null && info.gt_ym >= reportYM;
    var gBase = isT012 ? g.t012 : g.old;
    var gAmount = +(gBase * tier.gpct).toFixed(2);
    var revenue = +t.cost.toFixed(2);
    var incentive = +(tier.inc * t.qty).toFixed(2);
    var gTopup = +Math.max(0, gAmount - revenue).toFixed(2);

    return {
      truck: tn,
      division: info.division||"", dept: info.dept||"", sect: info.sect||"",
      plant_no: info.plant_no||"", plant: info.plant||"",
      vendor: info.vendor||"", vendor_id: info.vendor_id||"",
      fleet: info.fleet||"", size: info.size||"", gt: info.gt||null,
      gt_ok: !!isT012,
      g_missing: !(info.plant_no in garantee),
      days: t.days, rmk: t.rmk,
      score: score, pct: pct, tier: tier.name,
      qty: +t.qty.toFixed(2), revenue: revenue,
      inc_rate: tier.inc, incentive: incentive,
      g_base: gBase, g_pct: +(tier.gpct*100).toFixed(1),
      g_amount: gAmount, g_topup: gTopup,
      net: +(incentive + gTopup).toFixed(2)
    };
  });

  var drivers = Object.keys(drvInfo).sort().map(function(dn){
    var days = {}, total = 0, src = fineRows[dn] || {};
    for (var d in src){ days[d] = +src[d].toFixed(2); total += src[d]; }
    var o = drvInfo[dn];
    return {driver:dn, division:o.division, dept:o.dept, sect:o.sect,
            plant_no:o.plant_no, plant:o.plant, vendor_id:o.vendor_id, vendor:o.vendor,
            days:days, total:+total.toFixed(2)};
  });

  report("เสร็จ", 1);
  return {
    meta: {
      year: YEAR, month: MONTH, month_en: monthEn, days_in_month: daysInMonth,
      work_days: workDays, work_note: workNote,
      generated: new Date().toISOString().slice(0,16).replace("T"," "),
      punch_rows: detail.length, rest_hours: REST_HOURS, late_grace: LATE_GRACE_MIN,
      engine: ENGINE_VERSION
    },
    tiers: TA_TIERS.map(function(t){ return {name:t.name, min:t.min, inc:t.inc, gpct:+(t.gpct*100).toFixed(1)}; }),
    rules: rules,
    trucks: trucks,
    drivers: drivers,
    detail: {
      cols: ["day","plant_no","truck","driver","device","pin","pout","ct_stat","check_time",
             "d_in","d_dev","d_out","d_ct","pre","net","remark"],
      meaning: MEANING,
      rows: detail
    }
  };
}

if (typeof module !== "undefined") module.exports = {buildDB:buildDB, ENGINE_VERSION:ENGINE_VERSION};
