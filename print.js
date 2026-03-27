// =====================================================
// print.js — PDF / Print Functions (Frontend-only)
// ทุก PDF สร้างจาก App.students ที่โหลดอยู่แล้ว
// ไม่ต้อง call GAS เพิ่มเลย
// =====================================================

// ── HELPERS ──────────────────────────────────────────

function writeToPrintWindow(win, html, title = 'พิมพ์') {
  if (!win) return;
  try {
    // inject print script ถ้ายังไม่มี
    const PRINT_SCRIPT = '<script>window.onload=()=>setTimeout(()=>window.print(),800);<\/script>';
    if (!html.includes('window.print()')) {
      html = html.replace('</body>', PRINT_SCRIPT + '</body>');
    }
    win.document.open();
    win.document.write(html.includes('<html') ? html
      : `<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8"><title>${title}</title></head><body>${html}</body></html>`);
    win.document.close();
  } catch (e) { Utils.toast('เปิดหน้าพิมพ์ไม่ได้', 'error'); }
}

function escH(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function scoreToGrade(score) {
  const s = Number(score) || 0;
  if (s >= 80) return '4';   if (s >= 75) return '3.5';
  if (s >= 70) return '3';   if (s >= 65) return '2.5';
  if (s >= 60) return '2';   if (s >= 55) return '1.5';
  if (s >= 50) return '1';   return '0';
}

/**
 * คำนวณเกรดจากคะแนนจริง normalize เป็น 0–100 ก่อน
 * @param {number|string} score    คะแนนที่ได้
 * @param {number}        maxScore คะแนนเต็ม (50=ประถม/เทอม, 100=ประถมรวม/มัธยม)
 * @param {boolean}       isR      ติด ร
 */
function calcGradeNorm(score, maxScore, isR) {
  if (isR) return 'ร';
  const s = Number(score);
  if (score === '' || score == null || isNaN(s)) return '';
  if (s === 0) return '0';
  const max = Number(maxScore) || 100;
  const pct = max === 100 ? s : Math.round((s / max) * 100);
  return scoreToGrade(pct);
}

function isAddSubject(name) {
  return ['เพิ่มเติม','คอมพิวเตอร์','ต้านทุจริต','หน้าที่พลเมือง','บูรณาการ','โครงงาน','อาเซียน'].some(k => String(name||'').includes(k));
}

function normHolistic(val) {
  const s = String(val || '');
  if (s.includes('3')) return 'ดีเยี่ยม (3)';
  if (s.includes('2')) return 'ดี (2)';
  if (s.includes('1')) return 'ผ่าน (1)';
  if (s.includes('0')) return 'ไม่ผ่าน (0)';
  return val || 'ดีเยี่ยม (3)';
}

function normBehaviorLevel(txt) {
  if (!txt) return '';
  const t = String(txt).trim();
  if (t.includes('ดีเยี่ยม')) return 'ดีเยี่ยม';
  if (t.includes('ดี'))       return 'ดี';
  if (t.includes('ผ่าน'))     return 'ผ่าน';
  return t;
}

function classLabel(cls) {
  return cls
    .replace('ม.1 เทอม 1','มัธยมศึกษาปีที่ 1 ภาคเรียนที่ 1')
    .replace('ม.1 เทอม 2','มัธยมศึกษาปีที่ 1 ภาคเรียนที่ 2')
    .replace('ม.2 เทอม 1','มัธยมศึกษาปีที่ 2 ภาคเรียนที่ 1')
    .replace('ม.2 เทอม 2','มัธยมศึกษาปีที่ 2 ภาคเรียนที่ 2')
    .replace('ม.3 เทอม 1','มัธยมศึกษาปีที่ 3 ภาคเรียนที่ 1')
    .replace('ม.3 เทอม 2','มัธยมศึกษาปีที่ 3 ภาคเรียนที่ 2')
    .replace('ป.1','ประถมศึกษาปีที่ 1').replace('ป.2','ประถมศึกษาปีที่ 2')
    .replace('ป.3','ประถมศึกษาปีที่ 3').replace('ป.4','ประถมศึกษาปีที่ 4')
    .replace('ป.5','ประถมศึกษาปีที่ 5').replace('ป.6','ประถมศึกษาปีที่ 6');
}

function safeArr(v) {
  return Array.isArray(v) ? v : [];
}

function safeObj(v) {
  return v && typeof v === 'object' ? v : {};
}

function toNum(v, d = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : d;
}

function getSchoolProfileSafe() {
  return safeObj(App.schoolProfile);
}

function getCourseInfoSafe() {
  if (App.courseInfo && typeof App.courseInfo === 'object') return App.courseInfo;
  if (typeof getCourseInfoForm === 'function') {
    try { return getCourseInfoForm() || {}; } catch (e) {}
  }
  return {};
}

function countHolisticLevel(rows, key, lv) {
  return rows.filter(r => String(r[key] || '').includes(`(${lv})`)).length;
}

// ── PDF BASE STYLES ──────────────────────────────────

function buildPdfHead(opts = {}) {
  const pageSize = opts.pageSize || 'portrait';
  const margin   = opts.margin   || (pageSize === 'landscape' ? '12mm' : '12mm 15mm');
  const extra    = opts.extra    || '';
  const FONT = "'Sarabun','Tahoma',sans-serif";
  return `
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
<style>
@page { size: A4 ${pageSize}; margin: ${margin}; }
*,html,body,table,th,td,div,span,p,b,strong { font-family:${FONT}!important; font-style:normal!important; }
body { color:#000; font-size:17px; line-height:1.35; margin:0; padding:0; }
b,strong,th,.bold { font-weight:700!important; }
h1,h2,h3,h4,p { margin:2px 0; }
.text-center{text-align:center;} .bg-gray{background:#e2e2e2;}
.page{page-break-after:always;box-sizing:border-box;}
.page:last-child{page-break-after:auto;}
table{border-collapse:collapse;width:100%;}
th,td{border:1px solid #000;padding:4px 6px;font-family:${FONT}!important;}
th{background:#e5e7eb;font-weight:700;text-align:center;}
.center{text-align:center;} .right{text-align:right;} .left{text-align:left;}
.page-no-pad{padding:0!important;}
${extra}
</style>`;
}

// ── ปพ.6 รายบุคคล ────────────────────────────────────

function buildPp6StudentData(stu, type) {
  const classroom = $('gClass').value;
  const year      = $('gYear').value;
  const subject   = $('gSubj').value;
  const ignoreR   = App.ignoreR;
  const isSem     = App.isSemMode;

  const holH = stu.holistic_homeroom || {};
  const grades = stu.grades || {};
  const subjectNames = Object.keys(App.subs || {}).filter(Boolean);

  // ถ้าไม่มี subs ให้ดึงจาก DOM
  const allSubjects = subjectNames.length ? subjectNames
    : [...$$('#gtBody tr[data-sid]')].length
      ? [...new Set([...$$('#gtBody tr[data-sid]')].map(r => r.getAttribute('data-subj') || subject))]
      : [subject];

  const basicSubjects = [], addSubjects = [];
  let totalGrade = 0, gradeCount = 0, hasR = false;

  // ดึงข้อมูลคะแนนจาก DOM ตาราง (เหมือนที่ getTableDataForSummary ทำ)
  const rowEl = document.querySelector(`#gtBody tr[data-sid="${stu.studentId}"]`);
  const gradeRows = rowEl ? [] : []; // fallback

  // สร้างจาก grades ใน App.students
  allSubjects.forEach(subj => {
    const g = stu.grades || {};
    let score, grade;

    if (isSem) {
      // โหมด sem (มัธยม): ดูจาก type ว่าเทอมไหน
      const tNum = type === 'term2' ? 2 : 1;
      score = Number(g[`t${tNum}_total`]) || 0;
      grade = String(g[`t${tNum}_grade`] || scoreToGrade(score));
    } else if (type === 'term1') {
      score = Number(g.t1_total) || 0;
      grade = scoreToGrade(score);
    } else if (type === 'term2') {
      score = Number(g.t2_total) || 0;
      grade = scoreToGrade(score);
    } else {
      score = Number(g.grand_total) || 0;
      grade = scoreToGrade(score);
    }

    if (ignoreR && grade === 'ร') grade = scoreToGrade(score);
    if (grade === 'ร') hasR = true;
    else if (grade) { totalGrade += Number(grade) || 0; gradeCount++; }

    const obj = { name: subj, score: score === 0 ? '' : score, grade };
    (isAddSubject(subj) ? addSubjects : basicSubjects).push(obj);
  });

  return {
    id: stu.studentId, name: stu.name, classroom, year,
    attendance: stu.stats || {},
    basicSubjects, addSubjects, hasR,
    holistic: {
      char: holH.char || 'ดีเยี่ยม (3)',
      read: holH.read || 'ดีเยี่ยม (3)',
      skill: holH.skill || 'ดีเยี่ยม (3)'
    },
    gpa: hasR ? 'รอผล' : (gradeCount > 0 ? (totalGrade / gradeCount).toFixed(2) : '')
  };
}

function buildPp6HTML(d, type) {
  const OBEC = 'https://raw.githubusercontent.com/Bk14School/easygrademain/refs/heads/main/logo-OBEC.png';
  const isSem  = App.isSemMode;
  const isMath = ['ม.1','ม.2','ม.3'].some(c =>
    String(d.classroom || '').split(' เทอม')[0].trim().startsWith(c)
  ) || !!(App && App._pp6MathMode);
  // fallback chain: parameter → d._printType → App state → 'year'
  const resolvedType = type || d._printType || App._pp6PrintType || 'year';
  console.log('[ปพ.6 DEBUG] buildPp6HTML type=', type, '| d._printType=', d._printType, '| App._pp6PrintType=', App._pp6PrintType, '| resolvedType=', resolvedType);
  const termNo = resolvedType === 'term2' ? '2' : '1';
  const fullCls = classLabel(d.classroom).replace(/ภาคเรียนที่\s*[12]/g,'').trim();
  const yr = String(d.year || '').trim() || $('gYear')?.value || '';
  const yearLbl = resolvedType === 'term1' ? `ภาคเรียนที่ 1 ปีการศึกษา ${yr}`
    : resolvedType === 'term2' ? `ภาคเรียนที่ 2 ปีการศึกษา ${yr}`
    : isSem ? `ภาคเรียนที่ ${termNo} ปีการศึกษา ${yr}`
    : `ปีการศึกษา ${yr}`;

  // มัธยม: คะแนนเต็ม 100 เสมอ (ทั้งเทอม 1 และเทอม 2)
  // ประถมเทอม 1/2: 50 | ประถมทั้งปี: 100
  const maxScore = isMath ? 100
    : (resolvedType === 'term1' || resolvedType === 'term2') ? 50
    : 100;

  // ประถมเทอม 1: แสดงแค่คะแนน ไม่มีเกรด
  // ประถมเทอม 1: แสดงแค่คะแนน ไม่มีเกรด
  const isPrathumT1 = !isMath && resolvedType === 'term1';
  const _isRGrade = (g) => {
    const s = String(g||'').trim();
    return s === 'ร' || s.includes('ติด') || (!['4','3.5','3','2.5','2','1.5','1','0',''].includes(s) && isNaN(Number(s)));
  };
  const mkRow = (s) => {
    const scoreIsR = _isRGrade(s.score);
    const gradeIsR = _isRGrade(s.grade) || scoreIsR; // ถ้า score ติด ร → grade ก็ติด ร
    const scoreDisplay = scoreIsR ? 'ติด ร'
      : (s.score === '' || s.score == null) ? '' : s.score;
    const gradeDisplay = gradeIsR ? 'ร'
      : (s.grade === '' || s.grade == null) ? '' : s.grade;
    if (isPrathumT1) {
      return `<tr><td>${escH(s.name)}</td>
         <td class="center">${scoreDisplay}</td></tr>`;
    }
    return `<tr><td>${escH(s.name)}</td>
       <td class="center">${scoreDisplay}</td>
       <td class="center" style="font-weight:700;${gradeIsR?'color:#dc2626;':''}">${escH(gradeDisplay)}</td></tr>`;
  };
  const emptyRow = isPrathumT1
    ? `<tr><td>-</td><td></td></tr>`
    : `<tr><td>-</td><td></td><td></td></tr>`;
  // คำนวณ hasR และ GPA จาก grade ที่มีใน items โดยตรง (ไม่พึ่ง GAS)
  // _isR เช็คทุกรูปแบบ: 'ร', 'ติด ร', 'ติด s', score=0 แต่ไม่มีเกรด
  const _isR = (g) => {
    const s = String(g||'').trim();
    if (s === '') return false;
    if (s === 'ร' || s.includes('ติด')) return true;
    // encoding ผิด — ไม่ใช่ตัวเลขและไม่ใช่เกรดปกติ
    const validGrades = ['4','3.5','3','2.5','2','1.5','1','0'];
    if (!validGrades.includes(s) && isNaN(Number(s))) return true;
    return false;
  };
  const _allItems = [...(d.basicSubjects||[]), ...(d.addSubjects||[])];
  // เช็ค hasR จากทั้ง grade และ score (กรณี grade ว่างแต่ score = 'ติด ร')
  const _actualHasR = _allItems.some(s => _isR(s.grade) || _isR(s.score));
  const _gradeVals = _allItems
    .map(s => {
      if (_isR(s.grade) || _isR(s.score)) return NaN;
      const g = String(s.grade||'').trim();
      if (g === '') return NaN;
      return Number(g);
    })
    .filter(n => !isNaN(n) && n >= 0);
  const _computedGpa = _actualHasR ? '' :
    (_gradeVals.length > 0
      ? (_gradeVals.reduce((a,b) => a+b, 0) / _gradeVals.length).toFixed(2)
      : '');

  const basicRows = d.basicSubjects.map(mkRow).join('') || emptyRow;
  const addRows   = d.addSubjects.map(mkRow).join('')   || emptyRow;

  return `<div class="page">
  <table style="width:100%;border:none;margin-bottom:10px;">
    <tr>
      <td style="width:15%;text-align:center;border:none;"><img src="${OBEC}" style="width:58px;"></td>
      <td style="border:none;text-align:center;line-height:1.4;">
        <div style="font-size:19px;font-weight:700;">แบบรายงานผลการพัฒนาคุณภาพผู้เรียนรายบุคคล (ปพ.6)</div>
        <div style="font-size:16px;font-weight:700;">${yearLbl}</div>
        <div style="font-size:15px;">โรงเรียนบ้านคลอง 14 สำนักงานเขตพื้นที่การศึกษาประถมศึกษานครนายก</div>
      </td>
    </tr>
  </table>
  <table style="width:100%;border:none;margin-bottom:10px;font-size:15px;">
    <tr>
      <td style="border:none;"><b>ชื่อ-สกุล</b> ${escH(d.name)} &nbsp; <b>รหัส</b> ${escH(d.id)}</td>
      <td style="border:none;text-align:right;"><b>ชั้น</b> ${escH(fullCls)}</td>
    </tr>
  </table>
  <table style="width:100%;border-collapse:collapse;table-layout:fixed;font-size:14px;margin-bottom:10px;">
    <thead>
      ${isPrathumT1 ? `
      <tr><th colspan="2">หลักสูตรการศึกษาขั้นพื้นฐาน</th></tr>
      <tr>
        <th style="width:75%;">ชื่อวิชา</th>
        <th style="width:25%;">รวม ${maxScore} คะแนน</th>
      </tr>` : `
      <tr><th colspan="3">หลักสูตรการศึกษาขั้นพื้นฐาน</th></tr>
      <tr>
        <th style="width:60%;">ชื่อวิชา</th>
        <th style="width:20%;">รวม ${maxScore} คะแนน</th>
        <th style="width:20%;">ผลการเรียน</th>
      </tr>`}
    </thead>
    <tbody>
      ${isPrathumT1
        ? `<tr><td colspan="2" style="background:#e5e7eb;font-weight:700;text-align:center;">วิชาพื้นฐาน</td></tr>`
        : `<tr><td colspan="3" style="background:#e5e7eb;font-weight:700;text-align:center;">วิชาพื้นฐาน</td></tr>`}
      ${basicRows}
      ${isPrathumT1
        ? `<tr><td colspan="2" style="background:#e5e7eb;font-weight:700;text-align:center;">วิชาเพิ่มเติม</td></tr>`
        : `<tr><td colspan="3" style="background:#e5e7eb;font-weight:700;text-align:center;">วิชาเพิ่มเติม</td></tr>`}
      ${addRows}
      ${isPrathumT1 ? '' : `
      <tr style="font-weight:700;">
        <td colspan="2" class="center">เกรดเฉลี่ย (GPA)</td>
        <td class="center">${_actualHasR ? '' : escH(_computedGpa)}</td>
      </tr>`}
    </tbody>
  </table>
  <div style="display:flex;gap:12px;margin-top:8px;">
    <div style="width:50%;">
      <table style="font-size:13px;width:100%;">
        <tr><th colspan="2">กิจกรรมพัฒนาผู้เรียน</th></tr>
        <tr><td>แนะแนว</td><td class="center">ผ่าน</td></tr>
        <tr><td>ลูกเสือ-เนตรนารี</td><td class="center">ผ่าน</td></tr>
        <tr><td>ชุมนุม</td><td class="center">ผ่าน</td></tr>
        <tr><td>กิจกรรมเพื่อสังคมฯ</td><td class="center">ผ่าน</td></tr>
      </table>
    </div>
    <div style="width:50%;">
      <table style="font-size:13px;width:100%;margin-bottom:8px;">
        <tr><td>คุณลักษณะอันพึงประสงค์</td><td class="center"><b>${escH(normHolistic(d.holistic.char))}</b></td></tr>
        <tr><td>การอ่าน คิดวิเคราะห์ และเขียน</td><td class="center"><b>${escH(normHolistic(d.holistic.read))}</b></td></tr>
      </table>
      <table style="border:none;width:100%;font-size:13px;margin-top:20px;">
        <tr><td style="border:none;text-align:center;padding-top:30px;">
          ลงชื่อ .............................................<br>
          <span style="font-size:12px;">( ครูประจำชั้น )</span>
        </td></tr>
        <tr><td style="border:none;text-align:center;padding-top:30px;">
          ลงชื่อ .............................................<br>
          <span style="font-size:12px;">( ผู้บริหารสถานศึกษา )</span>
        </td></tr>
      </table>
    </div>
  </div>
</div>`;
}

// ── ปพ.6 รายบุคคล (single) ───────────────────────────
async function fetchPp6SummaryData(type = 'year') {
  return await api('getPp6FromSummary', {
    year: $('gYear').value,
    classroom: $('gClass').value,
    mode: type
  });
}

async function printRowPDF(sid, type = 'year') {
  const cls = App.loadedClass || $('gClass')?.value || '';
  const isMathCls = ['ม.1','ม.2','ม.3'].some(c => cls.startsWith(c));
  if (isMathCls && type !== 'term1' && type !== 'term2') {
    App._pp6MathMode = true;
    App._pp6PendingSid = sid; // เก็บ sid ไว้รอ
    return _showPp6TermModalSingle(sid);
  }

  const win = window.open('', '_blank');
  if (!win) return Utils.toast('กรุณาอนุญาต Popup ใน browser ด้วยครับ', 'error');
  win.document.write('<html><body style="font-family:sans-serif;padding:20px;"><p>⏳ กำลังโหลดข้อมูล...</p></body></html>');

  Utils.showLoading('กำลังเตรียม ปพ.6...');
  try {
    const res = await fetchPp6SummaryData(type);
    const stu = (res.students || []).find(s => String(s.id) === String(sid));

    if (!stu) {
      Utils.hideLoading();
      win.close();
      return Utils.toast('ไม่พบข้อมูลนักเรียน', 'error');
    }

    const body = buildPp6HTML(stu, type);
    const extra = `
      .page{min-height:96vh;}
      table{font-size:14px;}
    `;

    writeToPrintWindow(win, `<!DOCTYPE html><html lang="th"><head>
      <meta charset="UTF-8"><title>ปพ.6 ${stu.name}</title>
      ${buildPdfHead({ extra })}
      </head><body>${body}
      <script>window.onload=()=>setTimeout(()=>window.print(),600);<\/script>
      </body></html>`);
  } catch (e) {
    Utils.toast(e.message || 'พิมพ์ ปพ.6 ไม่สำเร็จ', 'error');
  }
  Utils.hideLoading();
}

// ── ปพ.6 ทั้งห้อง ────────────────────────────────────

// ── router: เลือก mode ตามชั้น ──
async function printPp6Auto(type = 'year') {
  const cls = App.loadedClass || $('gClass')?.value || '';
  const isMath = ['ม.1','ม.2','ม.3'].some(c => cls.startsWith(c));
  App._pp6MathMode = isMath;

  if (isMath) {
    // มัธยม: ให้เลือกเทอมก่อน
    return _showPp6TermModal();
  }
  // ประถม: term1 = แสดงคะแนน, year = แสดงรวม+เกรด
  return printClassPDF(type);
}

// ── modal เลือกเทอม ปพ.6 มัธยม (รายบุคคล) ──
function _showPp6TermModalSingle(sid) {
  const old = document.getElementById('_pp6TermModal');
  if (old) old.remove();
  const modal = document.createElement('div');
  modal.id = '_pp6TermModal';
  modal.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:9500;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(4px);';
  modal.innerHTML = `
    <div style="background:#fff;border-radius:16px;padding:28px 32px;width:95%;max-width:360px;box-shadow:0 20px 50px rgba(0,0,0,.15);text-align:center;">
      <div style="font-size:1.5rem;margin-bottom:6px;">📋</div>
      <div style="font-weight:800;font-size:1rem;color:#1e293b;margin-bottom:4px;">พิมพ์ ปพ.6 มัธยม</div>
      <div style="font-size:.83rem;color:#64748b;margin-bottom:20px;">เลือกภาคเรียนที่ต้องการพิมพ์</div>
      <div style="display:flex;flex-direction:column;gap:10px;">
        <button onclick="document.getElementById('_pp6TermModal').remove(); printRowPDF('${sid}','term1')"
          style="padding:11px;border:none;border-radius:10px;background:linear-gradient(135deg,#6d28d9,#5b21b6);color:#fff;font-size:.95rem;font-weight:700;cursor:pointer;font-family:inherit;">
          📘 ภาคเรียนที่ 1
        </button>
        <button onclick="document.getElementById('_pp6TermModal').remove(); printRowPDF('${sid}','term2')"
          style="padding:11px;border:none;border-radius:10px;background:linear-gradient(135deg,#0369a1,#075985);color:#fff;font-size:.95rem;font-weight:700;cursor:pointer;font-family:inherit;">
          📗 ภาคเรียนที่ 2
        </button>
        <button onclick="document.getElementById('_pp6TermModal').remove()"
          style="padding:9px;border:1.5px solid #e2e8f0;border-radius:10px;background:#f8fafc;color:#64748b;font-size:.88rem;font-weight:700;cursor:pointer;font-family:inherit;">
          ยกเลิก
        </button>
      </div>
    </div>`;
  document.body.appendChild(modal);
}

// ── modal เลือกเทอม ปพ.6 มัธยม (ทั้งห้อง) ──
function _showPp6TermModal() {
  // ลบ modal เก่าถ้ามี
  const old = document.getElementById('_pp6TermModal');
  if (old) old.remove();

  const modal = document.createElement('div');
  modal.id = '_pp6TermModal';
  modal.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:9500;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(4px);';
  modal.innerHTML = `
    <div style="background:#fff;border-radius:16px;padding:28px 32px;width:95%;max-width:360px;box-shadow:0 20px 50px rgba(0,0,0,.15);text-align:center;">
      <div style="font-size:1.5rem;margin-bottom:6px;">📋</div>
      <div style="font-weight:800;font-size:1rem;color:#1e293b;margin-bottom:4px;">พิมพ์ ปพ.6 มัธยม</div>
      <div style="font-size:.83rem;color:#64748b;margin-bottom:20px;">เลือกภาคเรียนที่ต้องการพิมพ์</div>
      <div style="display:flex;flex-direction:column;gap:10px;">
        <button onclick="document.getElementById('_pp6TermModal').remove(); printClassPDF('term1')"
          style="padding:11px;border:none;border-radius:10px;background:linear-gradient(135deg,#6d28d9,#5b21b6);color:#fff;font-size:.95rem;font-weight:700;cursor:pointer;font-family:inherit;">
          📘 ภาคเรียนที่ 1
        </button>
        <button onclick="document.getElementById('_pp6TermModal').remove(); printClassPDF('term2')"
          style="padding:11px;border:none;border-radius:10px;background:linear-gradient(135deg,#0369a1,#075985);color:#fff;font-size:.95rem;font-weight:700;cursor:pointer;font-family:inherit;">
          📗 ภาคเรียนที่ 2
        </button>
        <button onclick="document.getElementById('_pp6TermModal').remove()"
          style="padding:9px;border:1.5px solid #e2e8f0;border-radius:10px;background:#f8fafc;color:#64748b;font-size:.88rem;font-weight:700;cursor:pointer;font-family:inherit;">
          ยกเลิก
        </button>
      </div>
    </div>`;
  document.body.appendChild(modal);
}

async function printClassPDF(type = 'year') {
  const cls = App.loadedClass || $('gClass')?.value || '';
  const isMathCls = ['ม.1','ม.2','ม.3'].some(c => cls.startsWith(c));

  // ถ้าเป็นมัธยมและยังไม่ได้เลือกเทอมที่ถูกต้อง → เปิด modal เลือกเทอม
  if (isMathCls && type !== 'term1' && type !== 'term2') {
    App._pp6MathMode = true;
    return _showPp6TermModal();
  }

  const termLabel = type === 'term1' ? 'ภาคเรียนที่ 1'
    : type === 'term2' ? 'ภาคเรียนที่ 2'
    : 'ทั้งปี';
  if (!confirm(`พิมพ์ ปพ.6 ${termLabel} ทั้งห้อง?`)) return;

  // เก็บ type ไว้ใน App state เป็น fallback
  App._pp6PrintType = type;

  const win = window.open('', '_blank');
  if (!win) return Utils.toast('กรุณาอนุญาต Popup ใน browser ด้วยครับ', 'error');
  win.document.write('<html><body style="font-family:sans-serif;padding:20px;"><p>⏳ กำลังโหลดข้อมูล...</p></body></html>');

  Utils.showLoading('กำลังเตรียม ปพ.6 ทั้งห้อง...');
  try {
    const res = await fetchPp6SummaryData(type);
    const students = res.students || [];
    // ใช้ mode ที่ GAS ยืนยันกลับมา เพื่อป้องกัน type หาย
    const confirmedType = res.mode || type;
    console.log('[ปพ.6 DEBUG] type=', type, '| res.mode=', res.mode, '| confirmedType=', confirmedType);

    if (!students.length) {
      Utils.hideLoading();
      win.close();
      return Utils.toast('ไม่พบข้อมูลนักเรียน', 'error');
    }

    const pages = students.map(stu => {
      stu._printType = confirmedType; // inject type เข้า object โดยตรง
      return buildPp6HTML(stu, confirmedType);
    }).join('\n');
    const extra = `.page{min-height:96vh;} table{font-size:14px;}`;

    writeToPrintWindow(win, `<!DOCTYPE html><html lang="th"><head>
      <meta charset="UTF-8"><title>ปพ.6 ทั้งห้อง</title>
      ${buildPdfHead({ extra })}
      </head><body>${pages}
      <script>window.onload=()=>setTimeout(()=>window.print(),800);<\/script>
      </body></html>`);
  } catch (e) {
    Utils.toast(e.message || 'พิมพ์ ปพ.6 ไม่สำเร็จ', 'error');
  }
  Utils.hideLoading();
}

// ── ตารางสรุปผลการเรียน ──────────────────────────────

async function printSummaryPDF(type = 'year') {
  closeSummaryModeModal();

  const classroom = $('gClass').value;
  const year      = $('gYear').value;
  const profile   = getSchoolProfileSafe();
  const schoolName = escH((profile && profile.school_name) || 'โรงเรียนบ้านคลอง 14');

  const termLabel = type === 'term1' ? `ภาคเรียนที่ 1 ปีการศึกษา ${year}`
    : type === 'term2' ? `ภาคเรียนที่ 2 ปีการศึกษา ${year}`
    : `ปีการศึกษา ${year}`;

  // เปิด window ทันที (ก่อน await) เพื่อหลีกเลี่ยง popup blocker
  const win = window.open('', '_blank');
  if (!win) return Utils.toast('กรุณาอนุญาต Popup ใน browser ด้วยครับ', 'error');
  win.document.write('<html><body style="font-family:sans-serif;padding:20px;"><p>⏳ กำลังโหลดข้อมูล...</p></body></html>');

  Utils.showLoading('กำลังโหลดข้อมูลสรุปเกรด...');
  let res;
  try {
    res = await api('getPp6FromSummary', { year, classroom, mode: type });
  } catch(e) {
    Utils.hideLoading();
    win.close();
    return Utils.toast(e.message || 'โหลดข้อมูลไม่สำเร็จ', 'error');
  }
  Utils.hideLoading();

  const students = res.students || [];
  if (!students.length) { win.close(); return Utils.toast('ไม่พบข้อมูลนักเรียน', 'error'); }

  // รวบรวมชื่อวิชาทั้งหมด (basic + add) จากนักเรียนทุกคน
  const basicNames = [], addNames = [], seenB = new Set(), seenA = new Set();
  students.forEach(stu => {
    (stu.basicSubjects || []).forEach(s => { if (!seenB.has(s.name)) { seenB.add(s.name); basicNames.push(s.name); } });
    (stu.addSubjects   || []).forEach(s => { if (!seenA.has(s.name)) { seenA.add(s.name); addNames.push(s.name); } });
  });
  const allSubjects = [...basicNames, ...addNames];

  // สร้าง map ของแต่ละนักเรียน
  const getSubMap = (stu) => {
    const m = {};
    [...(stu.basicSubjects||[]), ...(stu.addSubjects||[])].forEach(s => { m[s.name] = s; });
    return m;
  };

  // แบ่งวิชาออกเป็นกลุ่ม กลุ่มละไม่เกิน COLS_PER_PAGE วิชา
  const COLS_PER_PAGE = Math.ceil(allSubjects.length / Math.ceil(allSubjects.length / 6));
  const subjChunks = [];
  for (let i = 0; i < allSubjects.length; i += COLS_PER_PAGE) {
    subjChunks.push(allSubjects.slice(i, i + COLS_PER_PAGE));
  }

  const buildTable = (subjGroup, pageNum, totalPages) => {
    const hd1 = subjGroup.map(n =>
      `<th colspan="4" style="background:#1e3a5f;color:#fff;white-space:nowrap;padding:4px 6px;font-size:11px;">${escH(n)}</th>`
    ).join('');
    const hd2 = subjGroup.map(() =>
      `<th style="width:46px;font-size:10px;">เก็บ</th><th style="width:46px;font-size:10px;">สอบ</th><th style="width:52px;font-size:10px;">รวม</th><th style="width:48px;background:#fef3c7;color:#92400e;font-size:10px;">เกรด</th>`
    ).join('');
    const rows = students.map((stu, i) => {
      const sm = getSubMap(stu);
      const cells = subjGroup.map(n => {
        const s = sm[n];
        if (!s) return `<td></td><td></td><td></td><td></td>`;
        const grade = s.grade || '';
        const gradeColor = grade === 'ร' ? 'color:#dc2626;font-weight:700;'
          : grade === '4' || grade === '3.5' || grade === '3' ? 'color:#166534;font-weight:700;'
          : grade === '0' ? 'color:#dc2626;'
          : 'color:#1e293b;font-weight:700;';
        return `<td style="text-align:center;">${s.keep === '' || s.keep == null ? '' : s.keep}</td>` +
               `<td style="text-align:center;">${s.exam === '' || s.exam == null ? '' : s.exam}</td>` +
               `<td style="text-align:center;font-weight:600;">${s.score === '' ? '' : s.score}</td>` +
               `<td style="text-align:center;background:#fefce8;${gradeColor}">${escH(grade)}</td>`;
      }).join('');
      const gpaColor = stu.hasR ? 'color:#dc2626;font-weight:700;' : 'color:#166534;font-weight:700;';
      return `<tr style="${i%2===1?'background:#f8fafc;':''}">
        <td style="text-align:center;">${i+1}</td>
        <td style="text-align:center;font-size:10px;">${escH(stu.id)}</td>
        <td style="text-align:left;white-space:nowrap;padding:3px 6px;">${escH(stu.name)}</td>
        ${cells}
        ${pageNum === totalPages ? `<td style="text-align:center;${gpaColor}">${escH(stu.hasR ? 'รอผล' : stu.gpa)}</td>` : ''}
      </tr>`;
    }).join('');

    const gpaHeader = pageNum === totalPages
      ? `<th rowspan="2" style="background:#0f172a;color:#fff;width:54px;font-size:11px;">GPA</th>` : '';

    return `<div class="page-block">
      <div style="text-align:center;font-size:13px;font-weight:700;margin-bottom:2px;">
        ตารางสรุปผลการเรียน — ${escH(classroom)} | ${termLabel}
      </div>
      <div style="text-align:center;font-size:11px;color:#64748b;margin-bottom:6px;">${schoolName}
        &nbsp;|&nbsp; หน้า ${pageNum}/${totalPages}
      </div>
      <div class="tbl-wrap">
      <table>
        <thead>
          <tr>
            <th rowspan="2" style="background:#1e293b;color:#fff;width:32px;font-size:10px;">ที่</th>
            <th rowspan="2" style="background:#1e293b;color:#fff;width:66px;font-size:10px;">รหัส</th>
            <th rowspan="2" style="background:#1e293b;color:#fff;min-width:140px;font-size:10px;">ชื่อ-นามสกุล</th>
            ${hd1}
            ${gpaHeader}
          </tr>
          <tr>${hd2}</tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
      </div>
    </div>`;
  };

  const pages = subjChunks.map((chunk, idx) =>
    buildTable(chunk, idx + 1, subjChunks.length)
  ).join('\n');

  const html = `<!DOCTYPE html><html lang="th"><head>
  <meta charset="UTF-8">
  <title>สรุปเกรด ${escH(classroom)} ${termLabel}</title>
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
  <style>
    *{font-family:'Sarabun','Tahoma',sans-serif;box-sizing:border-box;}
    body{margin:0;padding:12px;font-size:13px;color:#1e293b;background:#f8fafc;}
    .tbl-wrap{overflow-x:auto;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.08);background:#fff;margin-bottom:8px;}
    table{border-collapse:collapse;font-size:12px;width:100%;}
    th,td{border:1px solid #e2e8f0;padding:4px 5px;vertical-align:middle;}
    thead th{background:#334155;color:#fff;font-weight:700;text-align:center;position:sticky;top:0;z-index:10;}
    thead tr:nth-child(2) th{top:30px;}
    .page-block{margin-bottom:20px;}
    .print-btn{
      display:block;margin:0 auto 14px;padding:10px 28px;
      background:linear-gradient(135deg,#6366f1,#4338ca);color:#fff;
      border:none;border-radius:8px;font-size:.95rem;font-weight:700;
      cursor:pointer;font-family:inherit;
    }
    .print-btn:hover{opacity:.88;}
    @media print{
      .print-btn{display:none!important;}
      body{padding:0;margin:0;background:#fff;}
      .tbl-wrap{box-shadow:none;border-radius:0;overflow:visible;}
      table{font-size:9.5px;width:100%;}
      th,td{border:0.5px solid #999;padding:2.5px 3px;}
      thead th{position:static!important;font-size:9px;}
      .page-block{page-break-after:always;}
      .page-block:last-child{page-break-after:auto;}
      @page{size:A4 landscape;margin:6mm 7mm;}
    }
  </style>
  </head><body>
  <button class="print-btn" onclick="window.print()">🖨️ พิมพ์ / บันทึก PDF</button>
  ${pages}
  </body></html>`;

  writeToPrintWindow(win, html);
}

// ── ปพ.5 รายวิชา ─────────────────────────────────────

function buildPp5Styles() {
  const F = "'Sarabun','Tahoma',sans-serif";
  const B = '1px solid #000';
  return `
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
<style>
@page{size:A4 portrait;margin:7mm 10mm;}
*,html,body,table,th,td,div,span,p,b,strong{font-family:${F}!important;font-style:normal!important;box-sizing:border-box;}
body{color:#000;font-size:14px;line-height:1.4;margin:0;padding:0;}
b,strong,th,.bold{font-weight:700!important;}
p,h1,h2,h3,h4{margin:2px 0;font-size:14px;}
.page{page-break-after:always!important;}
.page:last-child{page-break-after:auto!important;}
table{border-collapse:collapse;width:100%;font-size:14px;}
th,td{border:${B};padding:3px 3px;text-align:center;vertical-align:middle;}
th{background:#f2f2f2;font-weight:700;}
.left{text-align:left!important;}.right{text-align:right!important;}.center{text-align:center!important;}
.name-cell{text-align:left!important;padding-left:5px;white-space:nowrap;overflow:hidden;}
/* ── ปก ── */
.cv-wrap{padding:2mm 1mm;min-height:271mm;display:flex;flex-direction:column;justify-content:space-between;}
.cv-logo{text-align:center;margin-bottom:5px;}
.cv-title{text-align:center;font-size:20px;font-weight:700;margin:0 0 2px;}
.cv-sub{text-align:center;font-size:16px;font-weight:700;color:#222;margin-bottom:8px;}
.cv-box{border:1.5px solid #555;border-radius:4px;padding:7px 11px;margin-bottom:8px;}
.cv-row{display:flex;align-items:flex-end;margin-bottom:5px;flex-wrap:nowrap;}
.cv-row:last-child{margin-bottom:0;}
.cv-lbl{font-weight:700;white-space:nowrap;font-size:15px;flex-shrink:0;}
.cv-val{flex:1;border-bottom:1.5px solid #000;min-height:19px;padding:1px 5px;font-size:15px;margin-left:4px;}
.cv-sep{font-size:15px;padding:0 5px;flex-shrink:0;color:#555;padding-bottom:2px;}
.cv-sec{font-weight:700;font-size:14px;margin:6px 0 3px;border-left:4px solid #555;padding-left:7px;}
.pp5-summary th{background:#e5e7eb;padding:3px 2px;font-size:13px;}
.pp5-summary td{padding:3px 2px;font-size:13px;}
.pp5-mini{width:49%;font-size:13px;}
.pp5-mini th{background:#e5e7eb;padding:3px 2px;}
.pp5-mini td{padding:3px 2px;}
.approve-wrap{margin-top:4px;}
.approve-row{display:flex;align-items:flex-end;gap:8px;margin-bottom:9px;}
.approve-chk{white-space:nowrap;font-size:15px;padding-bottom:2px;flex-shrink:0;}
.sign-area{flex:1;border-bottom:1.5px solid #000;min-height:24px;}
.role-lbl{white-space:nowrap;font-size:15px;padding-bottom:2px;flex-shrink:0;min-width:170px;text-align:right;}
.box{display:inline-block;width:13px;height:13px;border:${B};vertical-align:middle;margin-right:3px;}
.director-wrap{text-align:center;margin-top:5px;font-size:15px;line-height:1.9;}
.dir-sign{display:inline-block;border-bottom:1.5px solid #000;min-width:200px;padding-bottom:2px;}
/* ── หน้าคำอธิบาย ── */
.info-tbl{width:100%;border-collapse:collapse;margin-bottom:8px;}
.info-tbl td{border:none;padding:3px 5px;font-size:14px;}
.info-lbl{font-weight:700;white-space:nowrap;width:1%;}
.info-val{border-bottom:1px solid #777;}
.info-rhs{font-weight:700;white-space:nowrap;width:1%;padding-left:14px!important;}
.info-rval{border-bottom:1px solid #777;white-space:nowrap;width:1%;}
.desc-box{border:1.2px solid #444;min-height:570px;padding:12px 15px;line-height:1.85;white-space:pre-wrap;text-indent:28px;font-size:14px;}
/* ── ตารางคะแนน ── */
.pp5-title{text-align:center;font-size:18px;font-weight:700;margin:4px 0 7px;}
.score-table th{background:#e5e7eb;}
.vertical{writing-mode:vertical-rl;transform:rotate(180deg);white-space:normal;font-size:12px;height:120px;padding:4px 1px;line-height:1.2;}
/* ── ตารางเช็คชื่อ ── */
.att-table th{background:#e5e7eb;}
.att-name{text-align:left!important;padding-left:5px;white-space:nowrap;overflow:hidden;}
.col-name{text-align:left!important;padding-left:5px;white-space:nowrap;}
.col-date{width:18px;font-size:11px;padding:2px 1px;}
.col-dow{background:#f9fafb;font-size:11px;color:#555;}
.col-sum{width:24px;font-size:11px;font-weight:700;}
.att-legend{font-size:11px;margin-top:4px;color:#555;}
.att-head-title{font-size:15px;font-weight:700;text-align:center;}
.att-head-sub{font-size:14px;text-align:center;margin-bottom:4px;}
</style>`;
}

// ── SAFE CONFIG (สำคัญมาก) ──────────────────────────
function getPrintConfigSafe() {
  if (App.config && typeof App.config === 'object') return App.config;

  if (typeof buildConfigPayload === 'function') {
    try {
      return buildConfigPayload();
    } catch (e) {
      console.warn('buildConfigPayload failed:', e);
    }
  }

  return {
    t1_acc: 0,
    t1_exam: 0,
    t2_acc: 0,
    t2_exam: 0,
    units: {
      t1: Array.isArray(App.units?.[1]) ? App.units[1] : [],
      t2: Array.isArray(App.units?.[2]) ? App.units[2] : []
    },
    subItems: { t1: [], t2: [] },
    rawMax: { t1: 0, t2: 0 },
    courseInfo: App.courseInfo || {}
  };
}
// CLASS_SEQ_NUMS (mirror จาก subject-manager.js)
const _PP5_SEQ_NUMS = {
  'ม.1': { t1: 1, t2: 2 },
  'ม.2': { t1: 3, t2: 4 },
  'ม.3': { t1: 5, t2: 6 }
};

function getPp5Data(forceTermNum) {
  const classroom  = $('gClass')?.value || '';
  const year       = $('gYear')?.value || '';
  const subjectRaw = $('gSubj')?.value || '';
  const profile    = getSchoolProfileSafe();
  const courseInfo = getCourseInfoSafe();
  const config     = getPrintConfigSafe();
  const students   = safeArr(App.students);
  const termDates  = safeObj(App.termDates);
  const holidays   = safeArr(App.holidays);

  const isSemMode = !!App.isSemMode;
  const isTerm2   = classroom.includes('เทอม 2');
  // forceTermNum ใช้เมื่อเรียกจาก _doPrintPp5(1 หรือ 2)
  const termNum   = forceTermNum ? String(forceTermNum) : (isSemMode ? (isTerm2 ? '2' : '1') : 'ปี');
  const tk        = 't' + (forceTermNum || (isTerm2 ? 2 : 1)); // 't1' หรือ 't2'

  // ── resolve ชื่อวิชาต่อเนื่อง (isSeq) ──
  // ถ้าวิชานี้เป็น seq → ต่อเลขท้ายตามชั้นและเทอม
  // เช่น ม.3 เทอม 1 → "เทคโนโลยี(วิทยาการคำนวณ) 5"
  const clsBase   = classroom.split(' เทอม')[0].trim();
  const seqNums   = _PP5_SEQ_NUMS[clsBase];
  const isSeq     = !!(config && config.isSeq);
  let subject = subjectRaw;
  if (isSeq && seqNums) {
    const seqNum = tk === 't2' ? seqNums.t2 : seqNums.t1;
    subject = subjectRaw + ' ' + seqNum;
  }

  const sch = safeObj(courseInfo.schedule);
  const weeklyHours =
    toNum(sch.mon) +
    toNum(sch.tue) +
    toNum(sch.wed) +
    toNum(sch.thu) +
    toNum(sch.fri);

  const totalHours = toNum(sch.totalHours);

  const rows = students.map((stu, i) => {
    const g = safeObj(stu.grades);
    const hs = safeObj(stu.holistic_subject);

    const tNum  = forceTermNum || (isTerm2 ? 2 : 1);
    const totalScore = isSemMode
      ? toNum(g[`t${tNum}_total`])
      : toNum(g.grand_total);

    let grade = '';
    if (isSemMode) {
      const rStatus = String(g.t2_acc || '').trim();
      grade = (rStatus === 'ติด ร' && !App.ignoreR)
        ? 'ร'
        : (totalScore === 0 ? '' : scoreToGrade(totalScore));
    } else {
      grade = totalScore === 0 ? '' : scoreToGrade(totalScore);
    }

    return {
      no: i + 1,
      studentId: stu.studentId,
      name: stu.name,
      keepScore: isSemMode ? toNum(g[`t${tNum}_acc`]) : (toNum(g.t1_acc) + toNum(g.t2_acc)),
      examScore: isSemMode ? toNum(g[`t${tNum}_exam`]) : (toNum(g.t1_exam) + toNum(g.t2_exam)),
      totalScore,
      grade,
      char: hs.char || '',
      read: hs.read || '',
      skill: hs.skill || '',
      char_scores: safeArr(hs.char_scores).length ? hs.char_scores : Array(8).fill(''),
      read_scores: safeArr(hs.read_scores).length ? hs.read_scores : Array(5).fill(''),
      skill_scores: safeArr(hs.skill_scores).length ? hs.skill_scores : Array(5).fill(''),
      char_total: hs.char_total || '',
      read_total: hs.read_total || '',
      skill_total: hs.skill_total || ''
    };
  }).filter(r => !String(r.name || '').includes('ย้ายออก'));

  const gb = { '4':0, '3.5':0, '3':0, '2.5':0, '2':0, '1.5':0, '1':0, '0':0, 'ร':0 };
  rows.forEach(r => {
    if (gb[r.grade] !== undefined) gb[r.grade]++;
  });

  const pct = n => rows.length > 0 ? ((n * 100) / rows.length).toFixed(2) : '0.00';

  const char3 = countHolisticLevel(rows, 'char', 3);
  const char2 = countHolisticLevel(rows, 'char', 2);
  const char1 = countHolisticLevel(rows, 'char', 1);
  const char0 = countHolisticLevel(rows, 'char', 0);

  const read3 = countHolisticLevel(rows, 'read', 3);
  const read2 = countHolisticLevel(rows, 'read', 2);
  const read1 = countHolisticLevel(rows, 'read', 1);
  const read0 = countHolisticLevel(rows, 'read', 0);

  return {
    classroom,
    year,
    subject,
    profile,
    courseInfo,
    config,
    students,
    rows,
    termDates,
    holidays,
    isSemMode,
    isTerm2,
    termNum,
    tk,
    weeklyHours,
    totalHours,
    gb,
    pct,
    char3, char2, char1, char0,
    read3, read2, read1, read0,
    gradeStats: {
      g4: gb['4'], g35: gb['3.5'], g3: gb['3'], g25: gb['2.5'],
      g2: gb['2'], g15: gb['1.5'], g1: gb['1'], g0: gb['0'], gR: gb['ร'],
      p4: pct(gb['4']), p35: pct(gb['3.5']), p3: pct(gb['3']),
      p25: pct(gb['2.5']), p2: pct(gb['2']), p15: pct(gb['1.5']),
      p1: pct(gb['1']), p0: pct(gb['0']), pR: pct(gb['ร'])
    }
  };
}

function buildPp5Cover(d) {
  const OBEC = 'https://raw.githubusercontent.com/Bk14School/easygrademain/refs/heads/main/logo-OBEC.png';
  const sp   = safeObj(d.profile);
  const ci   = safeObj(d.courseInfo);
  const gb   = safeObj(d.gb);
  const pct  = typeof d.pct === 'function' ? d.pct : (() => '0.00');
  const rows = safeArr(d.rows);
  const n    = rows.length;
  const passCount = rows.filter(r => Number(r.totalScore) >= 50).length;

  const clsShort = String(d.classroom || '').split(' เทอม')[0].trim()
    .replace('ม.1','มัธยมศึกษาปีที่ 1')
    .replace('ม.2','มัธยมศึกษาปีที่ 2')
    .replace('ม.3','มัธยมศึกษาปีที่ 3')
    .replace('ป.1','ประถมศึกษาปีที่ 1')
    .replace('ป.2','ประถมศึกษาปีที่ 2')
    .replace('ป.3','ประถมศึกษาปีที่ 3')
    .replace('ป.4','ประถมศึกษาปีที่ 4')
    .replace('ป.5','ประถมศึกษาปีที่ 5')
    .replace('ป.6','ประถมศึกษาปีที่ 6');

  const r1 = (l,v) => `<div class="cv-row"><span class="cv-lbl">${l}</span><span class="cv-val">${escH(v??'')}</span></div>`;
  const r2 = (l1,v1,l2,v2) => `<div class="cv-row"><span class="cv-lbl">${l1}</span><span class="cv-val">${escH(v1??'')}</span><span class="cv-sep">|</span><span class="cv-lbl">${l2}</span><span class="cv-val">${escH(v2??'')}</span></div>`;
  const r3 = (l1,v1,l2,v2,w2,l3,v3,w3) => `<div class="cv-row"><span class="cv-lbl">${l1}</span><span class="cv-val">${escH(v1??'')}</span><span class="cv-sep">|</span><span class="cv-lbl">${l2}</span><span class="cv-val"${w2?` style="max-width:${w2}"`:''}>${escH(v2??'')}</span><span class="cv-sep">|</span><span class="cv-lbl">${l3}</span><span class="cv-val"${w3?` style="max-width:${w3}"`:''}>${escH(v3??'')}</span></div>`;

  return `<div class="page"><div class="cv-wrap">
  <div>
    <div class="cv-logo"><img src="${OBEC}" style="width:68px;"></div>
    <div class="cv-title">แบบบันทึกผลการเรียนประจำรายวิชา</div>
    <div class="cv-sub">${escH(sp.school_name||'โรงเรียนบ้านคลอง 14')}</div>
    <div class="cv-box">
      ${r2('อำเภอ/เขต','องครักษ์','จังหวัด','นครนายก')}
      ${r3('ชั้น',clsShort,'ห้อง','1','52px','ภาคเรียนที่',d.termNum,'48px')}
      ${r2('ปีการศึกษา',d.year,'รหัสวิชา',ci.code||'')}
      ${r1('รายวิชา',d.subject)}
      ${r3('หน่วยกิต',ci.credit||'','เวลาเรียน (ชม.)',d.totalHours,'62px','ชม./สัปดาห์',d.weeklyHours,'58px')}
      ${r1('ครูผู้สอน',ci.teacherName||'')}
      ${r1('ครูประจำชั้น',App.homeroomTeacher||'-')}
    </div>
    <div class="cv-sec">สรุปผลการประเมิน</div>
    <table class="pp5-summary">
      <tr><th rowspan="2" style="width:12%">จำนวน<br>นักเรียน</th><th colspan="9">ระดับผลการเรียน</th><th colspan="4">ผลการประเมิน</th></tr>
      <tr><th>4</th><th>3.5</th><th>3</th><th>2.5</th><th>2</th><th>1.5</th><th>1</th><th>0</th><th>ร</th><th>มส</th><th>มผ</th><th>ผ</th><th>อื่นๆ</th></tr>
      <tr><td><b>${n}</b></td><td>${gb['4']||0}</td><td>${gb['3.5']||0}</td><td>${gb['3']||0}</td><td>${gb['2.5']||0}</td><td>${gb['2']||0}</td><td>${gb['1.5']||0}</td><td>${gb['1']||0}</td><td>${gb['0']||0}</td><td>${gb['ร']||0}</td><td>0</td><td>0</td><td>${passCount}</td><td>0</td></tr>
      <tr><td>ร้อยละ</td><td>${pct(gb['4']||0)}</td><td>${pct(gb['3.5']||0)}</td><td>${pct(gb['3']||0)}</td><td>${pct(gb['2.5']||0)}</td><td>${pct(gb['2']||0)}</td><td>${pct(gb['1.5']||0)}</td><td>${pct(gb['1']||0)}</td><td>${pct(gb['0']||0)}</td><td>${pct(gb['ร']||0)}</td><td>0</td><td>0</td><td>${pct(passCount)}</td><td>0</td></tr>
    </table>
    <div style="display:flex;gap:6px;margin-top:4px;">
      <table class="pp5-mini"><tr><th colspan="4">คุณลักษณะอันพึงประสงค์</th></tr><tr><th>3=ดีเยี่ยม</th><th>2=ดี</th><th>1=ผ่าน</th><th>0=ไม่ผ่าน</th></tr><tr><td>${d.char3||0}</td><td>${d.char2||0}</td><td>${d.char1||0}</td><td>${d.char0||0}</td></tr><tr><td>${pct(d.char3||0)}</td><td>${pct(d.char2||0)}</td><td>${pct(d.char1||0)}</td><td>${pct(d.char0||0)}</td></tr></table>
      <table class="pp5-mini"><tr><th colspan="4">การอ่าน คิดวิเคราะห์ และเขียน</th></tr><tr><th>3=ดีเยี่ยม</th><th>2=ดี</th><th>1=ผ่าน</th><th>0=ไม่ผ่าน</th></tr><tr><td>${d.read3||0}</td><td>${d.read2||0}</td><td>${d.read1||0}</td><td>${d.read0||0}</td></tr><tr><td>${pct(d.read3||0)}</td><td>${pct(d.read2||0)}</td><td>${pct(d.read1||0)}</td><td>${pct(d.read0||0)}</td></tr></table>
    </div>
  </div>
  <div class="approve-wrap">
    <div class="cv-sec">การอนุมัติผลการเรียน</div>
    ${['ครูผู้สอน','หัวหน้ากลุ่มสาระการเรียนรู้',sp.academic_head_position||'หัวหน้าวิชาการ','รองผู้อำนวยการ'].map(role=>`
    <div class="approve-row"><span class="approve-chk"><span class="box"></span>อนุมัติ &nbsp;<span class="box"></span>ไม่อนุมัติ</span><span class="sign-area"></span><span class="role-lbl">(${role})</span></div>`).join('')}
    <div class="director-wrap">
      <div style="margin-bottom:4px;"><span class="box"></span>อนุมัติ &nbsp;<span class="box"></span>ไม่อนุมัติ</div>
            <div><span class="dir-sign"></span></div>
      <div>(${escH(sp.director_name||'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;')})</div>
      <div>${escH(sp.director_position||'ผู้อำนวยการโรงเรียน')}</div>
      <div style="margin-top:4px;">วันที่ .......... เดือน ........................ พ.ศ. ............</div>
    </div>
  </div>
</div></div>`;
}

function buildPp5DescPages(d) {
  const ci = d.courseInfo || {};
  const clsLvl = d.classroom.split(' เทอม')[0].trim()
    .replace('ม.1','มัธยมศึกษาปีที่ 1/1').replace('ม.2','มัธยมศึกษาปีที่ 2/1').replace('ม.3','มัธยมศึกษาปีที่ 3/1');

  const infoHdr = `<table class="info-tbl">
    <tr><td class="info-lbl">วิชา</td><td class="info-val">${escH(d.subject)}</td><td class="info-rhs">รหัสวิชา</td><td class="info-rval" style="width:88px;">${escH(ci.code||'')}</td></tr>
    <tr><td class="info-lbl">ชั้น</td><td class="info-val">${escH(clsLvl)}</td><td class="info-rhs">ภาคเรียน</td><td class="info-rval" style="width:38px;">${d.termNum}</td></tr>
    <tr><td class="info-lbl">ครูผู้สอน</td><td colspan="3" class="info-val">${escH(ci.teacherName||'')}</td></tr>
  </table>`;

  const desc = String(ci.description||'').trim();
  const indi = String(ci.indicators||'').trim();

  const makePage = (title, text) => `<div class="page">
    <div class="pp5-title">${title}</div>
    ${infoHdr}
    <div class="desc-box${text?'':' placeholder'}">${text ? escH(text).replace(/\n/g,'<br>') : `กรุณากรอก${title}`}</div>
  </div>`;

  return makePage('คำอธิบายรายวิชา', desc) + makePage('ตัวชี้วัด / ผลการเรียนรู้', indi);
}

function buildPp5ScorePage(d) {
  const config   = d.config || {};
  const tk       = d.tk || 't1'; // 't1' หรือ 't2'
  const units    = (config.units && config.units[tk]) || [];
  const accTotal = Number(config[tk + '_acc']) || 0;
  const examTotal= Number(config[tk + '_exam']) || 0;

  const bodyRows = units.map((u, i) => {
    const unitScore = (u.items||[]).reduce((s,it)=>s+(Number(it.max)||0),0);
    // indicators เป็น string[] — แสดงแต่ละตัวบนบรรทัดใหม่
    const inds = Array.isArray(u.indicators) && u.indicators.length
      ? u.indicators.map(s => escH(s)).join('<br>')
      : (typeof u.indicators === 'string' && u.indicators.trim() ? escH(u.indicators) : '-');
    return `<tr>
      <td class="center">${i+1}</td>
      <td style="text-align:left;">${escH(u.name||`หน่วยที่ ${i+1}`)}</td>
      <td style="text-align:left;font-size:12px;line-height:1.6;">${inds}</td>
      <td class="center">${escH(u.hours||'')}</td>
      <td class="center">${unitScore||''}</td>
      <td class="center"></td>
      <td class="center">${unitScore||''}</td>
    </tr>`;
  }).join('') || `<tr><td class="center">1</td><td>-</td><td class="center">-</td><td></td><td></td><td></td><td></td></tr>`;

  return `<div class="page">
    <div class="pp5-title">สัดส่วนคะแนน</div>
    <div style="text-align:center;font-size:16px;margin-bottom:8px;">
      ระหว่างภาค : ปลายภาค = ${accTotal} : ${examTotal}
    </div>
    <table class="score-table">
      <thead><tr>
        <th style="width:6%;">หน่วยที่</th>
        <th style="width:26%;">ชื่อหน่วย</th>
        <th style="width:30%;">ตัวชี้วัด</th>
        <th style="width:7%;">ชั่วโมง</th>
        <th style="width:11%;">ระหว่างภาค</th>
        <th style="width:10%;">ปลายภาค</th>
        <th style="width:10%;">รวม</th>
      </tr></thead>
      <tbody>
        <tr><td colspan="7" style="background:#e5e7eb;font-weight:700;text-align:center;">สัดส่วนคะแนนก่อนปลายภาค</td></tr>
        ${bodyRows}
        <tr style="font-weight:700;"><td colspan="4" class="right">รวมก่อนปลายภาค</td><td class="center">${accTotal}</td><td class="center">0</td><td class="center">${accTotal}</td></tr>
        <tr><td colspan="7" style="background:#e5e7eb;font-weight:700;text-align:center;">สัดส่วนคะแนนปลายภาค</td></tr>
        <tr><td class="center">-</td><td style="text-align:left;">คะแนนสอบปลายภาค</td><td class="center">-</td><td></td><td class="center">0</td><td class="center">${examTotal}</td><td class="center">${examTotal}</td></tr>
        <tr style="font-weight:700;"><td colspan="4" class="right">รวมทั้งหมด</td><td class="center">${accTotal}</td><td class="center">${examTotal}</td><td class="center">${accTotal+examTotal}</td></tr>
      </tbody>
    </table>
  </div>`;
}

function buildPp5ScoreSummaryPage(d) {
  const config   = d.config || {};
  const tk       = d.tk || 't1'; // 't1' หรือ 't2'
  const tNum     = tk === 't2' ? 2 : 1;
  const units    = (config.units && config.units[tk]) || [];
  const maxCols  = Math.min(units.length, 8);
  const accTotal = Number(config[tk + '_acc']) || 0;
  const examMax  = Number(config[tk + '_exam']) || 0;

  const unitThs = units.slice(0,maxCols).map((u,i)=>`<th style="width:44px;">${i+1}</th>`).join('');
  const unitMaxThs = units.slice(0,maxCols).map(u=>{
    const mx=(u.items||[]).reduce((s,it)=>s+(Number(it.max)||0),0);
    return `<th>${mx||''}</th>`;
  }).join('');

  const bodyRows = d.rows.map(r => {
    const g = r;
    const unitScores = Array(maxCols).fill('').map((_,i) => {
      const stu = d.students.find(s=>s.studentId===r.studentId);
      const unitKey = (tk === 't2') ? 't2_units' : 't1_units';
      const units_t1 = stu?.grades?.[unitKey] || [];
      const v = units_t1[i];
      return v === '' || v == null ? '' : v;
    });
    return `<tr>
      <td>${r.no}</td><td>${escH(r.studentId)}</td>
      <td class="name-cell">${escH(r.name)}</td>
      ${unitScores.map(v=>`<td>${v}</td>`).join('')}
      <td>${r.keepScore}</td><td>${r.examScore}</td>
      <td>${r.totalScore}</td><td style="font-weight:700;">${escH(r.grade)}</td>
    </tr>`;
  }).join('');

  return `<div class="page">
    <div style="text-align:center;font-size:16px;font-weight:700;margin-bottom:6px;">
      ตารางคะแนนระหว่างภาคเรียน<br>
      รายวิชา ${escH(d.subject)} &nbsp;|&nbsp; ชั้น ${escH(d.classroom.split(' เทอม')[0])}
    </div>
    <table class="score-table">
      <thead>
        <tr>
          <th rowspan="3" style="width:34px;">ที่</th>
          <th rowspan="3" style="width:70px;">รหัส</th>
          <th rowspan="3" style="width:200px;">ชื่อ - สกุล</th>
          <th colspan="${maxCols}">คะแนนรายหน่วย</th>
          <th rowspan="3" style="width:60px;">รวมระหว่างภาค</th>
          <th rowspan="3" style="width:52px;">ปลายภาค</th>
          <th rowspan="3" style="width:58px;">รวมคะแนน</th>
          <th rowspan="3" style="width:52px;">ผลการเรียน</th>
        </tr>
        <tr>${unitThs}</tr>
        <tr>${unitMaxThs}</tr>
      </thead>
      <tbody>
        <tr style="font-weight:700;"><td colspan="3">คะแนนเต็ม</td>${unitMaxThs}<td>${accTotal}</td><td>${examMax}</td><td>${accTotal+examMax}</td><td></td></tr>
        ${bodyRows}
      </tbody>
    </table>
  </div>`;
}

function buildPp5BehaviorPage(d) {
  const rows = d.rows.map(r => `<tr>
    <td>${r.no}</td><td>${escH(r.studentId)}</td><td class="name-cell">${escH(r.name)}</td>
    ${(r.char_scores||Array(8).fill('')).map(v=>`<td>${escH(v)}</td>`).join('')}
    <td>${escH(r.char_total)}</td><td style="font-size:13px;">${escH(normBehaviorLevel(r.char))}</td>
    ${(r.read_scores||Array(5).fill('')).map(v=>`<td>${escH(v)}</td>`).join('')}
    <td>${escH(r.read_total)}</td><td style="font-size:13px;">${escH(normBehaviorLevel(r.read))}</td>
  </tr>`).join('');

  return `<div class="page">
    <div style="text-align:center;font-size:15px;font-weight:700;margin-bottom:6px;">
      แบบประเมินคุณลักษณะอันพึงประสงค์ และการอ่าน คิดวิเคราะห์ และเขียน<br>
      รายวิชา ${escH(d.subject)} &nbsp;|&nbsp; ชั้น ${escH(d.classroom.split(' เทอม')[0])}
    </div>
    <table class="score-table" style="font-size:13px;">
      <thead>
        <tr>
          <th rowspan="3" style="width:30px;">ที่</th>
          <th rowspan="3" style="width:60px;">รหัส</th>
          <th rowspan="3" style="width:180px;">ชื่อ-นามสกุล</th>
          <th colspan="10">คุณลักษณะอันพึงประสงค์</th>
          <th colspan="7">อ่าน คิดวิเคราะห์ และเขียน</th>
        </tr>
        <tr>
          ${['รักชาติ','ซื่อสัตย์','มีวินัย','ใฝ่เรียนรู้','พอเพียง','มุ่งมั่น','รักความเป็นไทย','จิตสาธารณะ','รวม','ระดับ'].map(t=>`<th><div class="vertical">${t}</div></th>`).join('')}
          ${['จับใจความ','วิเคราะห์','วิจารณ์','เขียนสื่อ','ใช้ภาษา','รวม','ระดับ'].map(t=>`<th><div class="vertical">${t}</div></th>`).join('')}
        </tr>
        <tr>${[...Array(8)].map((_,i)=>`<th>${i+1}</th>`).join('')}<th></th><th></th>${[...Array(5)].map((_,i)=>`<th>${i+1}</th>`).join('')}<th></th><th></th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

function buildPp5SkillPage(d) {
  const rows = d.rows.map(r => `<tr>
    <td>${r.no}</td><td>${escH(r.studentId)}</td><td class="name-cell">${escH(r.name)}</td>
    ${(r.skill_scores||Array(5).fill('')).map(v=>`<td>${escH(v)}</td>`).join('')}
    <td>${escH(r.skill_total)}</td><td style="font-size:13px;">${escH(normBehaviorLevel(r.skill))}</td>
  </tr>`).join('');

  return `<div class="page" style="page-break-after:auto!important;">
    <div style="text-align:center;font-size:15px;font-weight:700;margin-bottom:6px;">
      ผลการประเมินสมรรถนะสำคัญของผู้เรียน<br>
      รายวิชา ${escH(d.subject)} &nbsp;|&nbsp; ชั้น ${escH(d.classroom.split(' เทอม')[0])}
    </div>
    <table class="score-table" style="font-size:13px;">
      <thead>
        <tr>
          <th rowspan="3" style="width:30px;">ที่</th>
          <th rowspan="3" style="width:60px;">รหัส</th>
          <th rowspan="3" style="width:200px;">ชื่อ-นามสกุล</th>
          <th colspan="7">สมรรถนะสำคัญของผู้เรียน</th>
        </tr>
        <tr>${['การสื่อสาร','การคิด','การแก้ปัญหา','ทักษะชีวิต','การใช้เทคโนโลยี','รวม','ระดับ'].map(t=>`<th><div class="vertical" style="height:160px;">${t}</div></th>`).join('')}</tr>
        <tr>${[...Array(5)].map((_,i)=>`<th>${i+1}</th>`).join('')}<th></th><th></th></tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

// ── ปพ.5 รายวิชา (main) ──────────────────────────────
function buildPp5AttendancePage(d) {
  const rows = d.rows.map(r => {
    const att = d.students.find(s => s.studentId === r.studentId)?.stats || {};
    const present  = Number(att.present)  || 0;
    const late     = Number(att.late)     || 0;
    const leave    = Number(att.leave)    || 0;
    const absent   = Number(att.absent)   || 0;
    const total    = present + late + leave + absent;
    return `<tr>
      <td>${r.no}</td>
      <td>${escH(r.studentId)}</td>
      <td class="name-cell">${escH(r.name)}</td>
      <td>${total || ''}</td>
      <td>${present || ''}</td>
      <td>${late || ''}</td>
      <td>${leave || ''}</td>
      <td>${absent || ''}</td>
      <td></td>
    </tr>`;
  }).join('');

  return `<div class="page">
    <div style="text-align:center;font-size:15px;font-weight:700;margin-bottom:6px;">
      ตารางการมาเรียน<br>
      รายวิชา ${escH(d.subject)} &nbsp;|&nbsp; ชั้น ${escH(d.classroom.split(' เทอม')[0])}
    </div>
    <table class="att-table">
      <thead>
        <tr>
          <th style="width:30px;">ที่</th>
          <th style="width:60px;">รหัส</th>
          <th class="col-name" style="width:200px;">ชื่อ-นามสกุล</th>
          <th style="width:50px;">เวลาเรียน<br>(ชม.)</th>
          <th style="width:50px;">มาเรียน<br>(ชม.)</th>
          <th style="width:50px;">มาสาย<br>(ชม.)</th>
          <th style="width:50px;">ลา<br>(ชม.)</th>
          <th style="width:50px;">ขาด<br>(ชม.)</th>
          <th style="width:50px;">หมายเหตุ</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

// ── ปพ.5 ตารางการมาเรียนรายเดือน (1 เดือน = 1 หน้า) ──

// ── กรอง attData เฉพาะเดือนในช่วงเทอมที่เลือก ──
function _filterAttDataByTerm(attData, termNum, termDates) {
  if (!attData || !attData.months) return attData;

  // parse วันจาก "dd/mm/yyyy_BE" เป็น Date (CE)
  const parseThDate = (str) => {
    if (!str) return null;
    const p = String(str).split('/');
    if (p.length !== 3) return null;
    const dd = parseInt(p[0]), mm = parseInt(p[1]), yy = parseInt(p[2]);
    return new Date(yy > 2500 ? yy - 543 : yy, mm - 1, dd);
  };

  const tk = termNum === 2 ? 2 : 1;
  const startDate = parseThDate(termDates[`t${tk}_start`]);
  const endDate   = parseThDate(termDates[`t${tk}_end`]);

  if (!startDate || !endDate) return attData; // ไม่มีวันเทอม → ไม่กรอง

  // monthKey format: "2568-05"  (BE year)
  const filteredMonths = attData.months.filter(mo => {
    // เช็คว่าเดือนนั้น overlap กับช่วงเทอม
    // (วันแรกของเดือน <= endDate) AND (วันสุดท้ายของเดือน >= startDate)
    const [byStr, mStr] = mo.key.split('-');
    const by = parseInt(byStr), m = parseInt(mStr);
    const cy = by > 2500 ? by - 543 : by;
    const monthStart = new Date(cy, m - 1, 1);
    const monthEnd   = new Date(cy, m, 0); // วันสุดท้ายของเดือน
    return monthStart <= endDate && monthEnd >= startDate;
  });

  // กรอง byMonth ของแต่ละนักเรียนด้วย
  const filteredStudents = attData.students.map(stu => {
    const byMonth = {};
    filteredMonths.forEach(mo => {
      byMonth[mo.key] = stu.byMonth[mo.key] || [];
    });
    return { ...stu, byMonth };
  });

  return { ...attData, months: filteredMonths, students: filteredStudents };
}

function buildPp5AttendanceMonthPages(attData, d) {
  if (!attData || !attData.months || !attData.months.length) {
    return `<div class="page"><div style="text-align:center;padding:40px;color:#888;">ไม่พบข้อมูลการมาเรียน</div></div>`;
  }
  const SIGN  = {'ม':'✓','ข':'ข','ล':'ล','ป':'ป','ลก':'ลก'};
  const COLOR = {'ข':'color:#dc2626;font-weight:700;','ล':'color:#d97706;','ป':'color:#d97706;','ลก':'color:#7c3aed;'};
  const clsShort   = d.classroom.split(' เทอม')[0];
  const schoolName = escH((App.schoolProfile||{}).school_name||'โรงเรียนบ้านคลอง 14');

  return attData.months.map(mo => {
    const dayThs = mo.days.map(dc=>`<th class="col-date">${dc.day}</th>`).join('');
    const dowThs = mo.days.map(dc=>`<th class="col-date col-dow">${dc.dow}</th>`).join('');
    const bodyRows = attData.students.map((stu,idx) => {
      const daily = stu.byMonth[mo.key]||[];
      let present=0,absent=0,leave=0;
      const cells = daily.map(cell => {
        const s=cell.status;
        if(s==='ม') present++; else if(s==='ข') absent++; else if(s==='ล'||s==='ป'||s==='ลก') leave++;
        return `<td style="${COLOR[s]||''}">${SIGN[s]||s||''}</td>`;
      }).join('');
      return `<tr><td>${idx+1}</td><td class="att-name">${escH(stu.name)}</td>${cells}<td class="col-sum">${present||''}</td><td class="col-sum" style="color:#dc2626;">${absent||''}</td><td class="col-sum" style="color:#d97706;">${leave||''}</td><td></td></tr>`;
    }).join('');
    return `<div class="page">
      <div class="att-head-title">บันทึกเวลาเรียนรายเดือน — ${mo.monthName} ${mo.bYear}</div>
      <div class="att-head-sub">รายวิชา ${escH(d.subject)} | ชั้น ${escH(clsShort)} | ปีการศึกษา ${d.year} | ${schoolName}</div>
      <table class="att-table">
        <thead>
          <tr><th rowspan="2" style="width:24px;">ที่</th><th rowspan="2" style="min-width:155px;" class="col-name">ชื่อ-นามสกุล</th>${dayThs}<th rowspan="2" class="col-sum">ม</th><th rowspan="2" class="col-sum" style="color:#dc2626;">ข</th><th rowspan="2" class="col-sum" style="color:#d97706;">ล/ป</th><th rowspan="2" style="width:46px;">หมายเหตุ</th></tr>
          <tr>${dowThs}</tr>
        </thead>
        <tbody>${bodyRows}</tbody>
      </table>
      <div class="att-legend">สัญลักษณ์: ✓ = มาเรียน &nbsp; ข = ขาด &nbsp; ล = ลา &nbsp; ป = ป่วย &nbsp; ลก = ลากิจ</div>
    </div>`;
  }).join('\n');
}
// ── เลือกเทอมก่อนพิมพ์ปพ.5 ──
function printPp5SubjectPDF() {
  const cls = $('gClass')?.value || '';
  const isMathayom = ['ม.1','ม.2','ม.3'].some(c => cls.startsWith(c));
  if (!isMathayom)
    return Utils.toast('ปพ.5 เวอร์ชันนี้ใช้กับมัธยมก่อน', 'warning');
  const students = App.students || [];
  if (!students.length) return Utils.toast('ไม่พบข้อมูลนักเรียน', 'error');

  const modal = document.createElement('div');
  modal.id = '_pp5TermModal';
  modal.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:9500;display:flex;align-items:center;justify-content:center;backdrop-filter:blur(4px);';
  modal.innerHTML = `
    <div style="background:#fff;border-radius:16px;padding:28px 32px;width:95%;max-width:360px;box-shadow:0 20px 50px rgba(0,0,0,.15);text-align:center;">
      <div style="font-size:1.5rem;margin-bottom:6px;">🖨️</div>
      <div style="font-weight:800;font-size:1rem;color:#1e293b;margin-bottom:4px;">พิมพ์ ปพ.5</div>
      <div style="font-size:.83rem;color:#64748b;margin-bottom:20px;">เลือกภาคเรียนที่ต้องการพิมพ์</div>
      <div style="display:flex;flex-direction:column;gap:10px;">
        <button onclick="window._doPrintPp5(1)" style="padding:11px;border:none;border-radius:10px;background:linear-gradient(135deg,#6d28d9,#5b21b6);color:#fff;font-size:.95rem;font-weight:700;cursor:pointer;font-family:inherit;">
          📘 ภาคเรียนที่ 1
        </button>
        <button onclick="window._doPrintPp5(2)" style="padding:11px;border:none;border-radius:10px;background:linear-gradient(135deg,#0369a1,#075985);color:#fff;font-size:.95rem;font-weight:700;cursor:pointer;font-family:inherit;">
          📗 ภาคเรียนที่ 2
        </button>
        <button onclick="document.getElementById('_pp5TermModal').remove()" style="padding:9px;border:1.5px solid #e2e8f0;border-radius:10px;background:#f8fafc;color:#64748b;font-size:.88rem;font-weight:700;cursor:pointer;font-family:inherit;">
          ยกเลิก
        </button>
      </div>
    </div>`;
  document.body.appendChild(modal);
}

async function _doPrintPp5(termNum) {
  const m = document.getElementById('_pp5TermModal');
  if (m) m.remove();

  const students = App.students || [];
  if (!students.length) return Utils.toast('ไม่พบข้อมูลนักเรียน', 'error');

  const win = window.open('', '_blank');
  if (!win) return Utils.toast('กรุณาอนุญาต Popup ใน browser ด้วยครับ', 'error');
  win.document.write('<html><body style="font-family:sans-serif;padding:20px;"><p>⏳ กำลังโหลดข้อมูล...</p></body></html>');

  Utils.showLoading('กำลังโหลดข้อมูลการมาเรียน...');
  let attData = null;
  try {
    const sch = {
      mon: +($('sch_mon')?.value) || 0,
      tue: +($('sch_tue')?.value) || 0,
      wed: +($('sch_wed')?.value) || 0,
      thu: +($('sch_thu')?.value) || 0,
      fri: +($('sch_fri')?.value) || 0
    };
    attData = await api('getAttendanceForPp5', {
      year:      $('gYear').value,
      classroom: $('gClass').value,
      schedule:  sch
    });
  } catch (e) {
    Utils.toast('โหลดข้อมูลการมาเรียนไม่ได้: ' + e.message, 'error');
  }
  Utils.hideLoading();

  const d = getPp5Data(termNum);

  // ── กรอง attData เฉพาะเดือนในเทอมที่เลือก ──
  const filteredAttData = _filterAttDataByTerm(attData, termNum, App.termDates || {});

  const pages = [
    buildPp5Cover(d),
    buildPp5DescPages(d),
    buildPp5ScorePage(d),
    buildPp5AttendanceMonthPages(filteredAttData, d),
    buildPp5AttendancePage(d),
    buildPp5ScoreSummaryPage(d),
    buildPp5BehaviorPage(d),
    buildPp5SkillPage(d)
  ].join('\n');

  writeToPrintWindow(win, `<!DOCTYPE html><html lang="th"><head>
    <meta charset="UTF-8"><title>ปพ.5 ${d.subject} ${d.classroom} ภาคเรียนที่ ${termNum}</title>
    ${buildPp5Styles()}
    </head><body>${pages}
    <script>window.onload=()=>setTimeout(()=>window.print(),1200);<\/script>
    </body></html>`);
}

// =====================================================
// ส่วนที่คงไว้ — ไม่เกี่ยวกับ GAS PDF
// =====================================================

function getTableDataForSummary(type) {
  const rows = [];
  $$('#gtBody tr[data-sid]').forEach((tr, index) => {
    const id = tr.querySelector('.s-c1')?.textContent.trim() || '';
    const name = tr.querySelector('.s-c2 span')?.textContent.trim() || '';
    
    let keep = '', exam = '', total = '', grade = '';
    
    if (App.isSemMode) {
      // โหมดมัธยม (รายเทอม)
      keep = tr.querySelector('.t1sc')?.textContent.trim() || '0';
      exam = tr.querySelector('.s-t1e')?.value || '0';
      total = tr.querySelector('.t1tot')?.textContent.trim() || '0';
      grade = tr.querySelector('.final-grade')?.textContent.trim() || '';
    } else {
      // โหมดประถม (รายปี)
      if (type === 'term1') {
        keep = tr.querySelector('.t1sc')?.textContent.trim() || '0';
        exam = tr.querySelector('.s-t1e')?.value || '0';
        total = tr.querySelector('.t1tot')?.textContent.trim() || '0';
        grade = '-'; // เทอม 1 ยังไม่ตัดเกรด
      } else if (type === 'term2') {
        keep = tr.querySelector('.t2sc')?.textContent.trim() || '0';
        exam = tr.querySelector('.s-t2e')?.value || '0';
        total = tr.querySelector('.t2tot')?.textContent.trim() || '0';
        grade = '-'; // เทอม 2 ย่อย ยังไม่ตัดเกรด
      } else {
        // ทั้งปี (year)
        keep = tr.querySelector('.sum-keep')?.textContent.trim() || '0';
        exam = tr.querySelector('.sum-exam')?.textContent.trim() || '0';
        total = tr.querySelector('.gtot')?.textContent.trim() || '0';
        grade = tr.querySelector('.final-grade')?.textContent.trim() || '';
      }
    }
    
    rows.push({ no: index + 1, id, name, keep, exam, total, grade });
  });
  return rows;
}

// =====================================================
// พิมพ์สรุปผลการเรียน (ดึงจาก HTML)
// =====================================================

async function exportSummaryCSV() {
  if (!confirm(`ดาวน์โหลดไฟล์ Excel (CSV) สรุปคะแนนวิชา ${$('gSubj').value}?`)) return; 
  
  // ใช้ Type = year เป็นค่าเริ่มต้นสำหรับดึงคอลัมน์ (ถ้าเป็นมัธยมมันจะจัดการให้อัตโนมัติ)
  const dataRows = getTableDataForSummary('year');
  if (!dataRows.length) return Utils.toast('ไม่พบข้อมูลนักเรียน', 'error');

  // ใส่ BOM (\uFEFF) เพื่อให้ Excel อ่านภาษาไทยได้ถูกต้อง
  let csvContent = "\uFEFFที่,รหัสประจำตัว,ชื่อ-นามสกุล,คะแนนเก็บ,คะแนนสอบ,รวม,เกรด\n";
  
  dataRows.forEach(r => {
    // ใส่ " คลุมชื่อเผื่อมีช่องว่างหรือเครื่องหมายคอมม่า
    csvContent += `${r.no},${r.id},"${r.name}",${r.keep},${r.exam},${r.total},${r.grade}\n`;
  });

  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `สรุปคะแนน_${$('gClass').value}_${$('gSubj').value}_ปี${$('gYear').value}.csv`;
  
  document.body.appendChild(link);
  link.click();
  link.remove();
  
  Utils.toast('✅ ดาวน์โหลดไฟล์ Excel สำเร็จ');
}

function printHolisticReport(mode) {
  const rows =[...$$(`${mode === 'homeroom' ? '#hrAssBody' : '#assBody'} tr[data-hsid]`)].map((tr, i) => ({
    no: i + 1,
    name: tr.querySelector('.ass-name')?.textContent || '',
    cS:[...tr.querySelectorAll('.c-inp')].map(i => i.value),
    cT: tr.querySelector('.c-total')?.textContent || '-',
    cR: tr.querySelector('.c-result')?.textContent || '-',
    rS: [...tr.querySelectorAll('.r-inp')].map(i => i.value),
    rT: tr.querySelector('.r-total')?.textContent || '-',
    rR: tr.querySelector('.r-result')?.textContent || '-',
    sS: [...tr.querySelectorAll('.s-inp')].map(i => i.value),
    sT: tr.querySelector('.s-total')?.textContent || '-',
    sR: tr.querySelector('.s-result')?.textContent || '-'
  }));

  if (!rows.length) return Utils.toast('ไม่พบข้อมูล', 'error');

  // ตรวจสอบว่าเป็นรายงานแบบรายวิชา หรือ ประจำชั้น
  const metaText = mode === 'homeroom' 
    ? `ชั้น ${$('gClass').value} | สรุปประเมินประจำชั้น` 
    : `ชั้น ${$('gClass').value} | รายวิชา ${$('gSubj').value}`;

  const mkTbl = (tit, hdrs, kS, kT, kR, cls) => {
    // นับจำนวนสรุปผลการประเมิน
    const counts = { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 };
    rows.forEach(r => {
      let res = r[kR];
      if (counts[res] !== undefined) counts[res]++;
    });

    return `<section class="print-page">
      <div class="report-head">
        <div class="report-school">โรงเรียนบ้านคลอง 14</div>
        <div class="report-title">${tit}</div>
        <div class="report-meta">${metaText}</div>
      </div>
      <table class="print-tbl ${cls}">
        <thead>
          <tr>
            <th class="col-no">ที่</th>
            <th class="col-name">ชื่อ-นามสกุล</th>
            ${hdrs.map(h => `<th>${h}</th>`).join('')}
            <th class="col-total">รวม</th>
            <th class="col-result">ผล</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => `<tr>
            <td>${r.no}</td>
            <td style="text-align:left">${escHtml(r.name)}</td>
            ${r[kS].map(v => `<td>${v || '-'}</td>`).join('')}
            <td>${r[kT]}</td>
            <td><b>${r[kR]}</b></td>
          </tr>`).join('')}
        </tbody>
      </table>
      <div class="report-summary">
        <b>สรุปผลการประเมิน:</b> 
        <span class="sum-item sum-3">ดีเยี่ยม: ${counts['ดีเยี่ยม']} คน</span> |
        <span class="sum-item sum-2">ดี: ${counts['ดี']} คน</span> |
        <span class="sum-item sum-1">ผ่าน: ${counts['ผ่าน']} คน</span> |
        <span class="sum-item sum-0">ไม่ผ่าน: ${counts['ไม่ผ่าน']} คน</span>
      </div>
    </section>`;
  };

  const win = window.open('', '_blank');
  if (!win) return Utils.toast('กรุณาอนุญาตให้เปิดหน้าต่างใหม่ (Popup)', 'error');

  const html = `<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="utf-8">
<title>รายงานประเมิน</title>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
  @page { size: A4 landscape; margin: 10mm; }
  body { font-family: "Sarabun", sans-serif; margin: 0; color: #111827; }
  .print-page { page-break-after: always; padding: 4mm 2mm; }
  .report-head { text-align: center; margin-bottom: 12px; }
  .report-school { font-size: 16px; font-weight: bold; }
  .report-title { font-size: 18px; font-weight: bold; margin: 4px 0; }
  .report-meta { font-size: 14px; color: #4b5563; }
  .print-tbl { width: 100%; border-collapse: collapse; font-size: 13px; }
  .print-tbl th, .print-tbl td { border: 1px solid #9ca3af; padding: 6px; text-align: center; }
  .print-tbl .col-no { width: 40px; }
  .print-tbl .col-name { width: 220px; }
  .theme-char thead th { background: #fef08a; }
  .theme-read thead th { background: #fdba74; }
  .theme-skill thead th { background: #d1d5db; }
  .report-summary { margin-top: 15px; text-align: right; font-size: 15px; padding: 10px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 6px; }
  .sum-item { margin: 0 8px; font-weight: bold; }
  .sum-3 { color: #166534; }
  .sum-2 { color: #075985; }
  .sum-1 { color: #92400e; }
  .sum-0 { color: #991b1b; }
</style>
</head>
<body>
  ${mkTbl('ประเมินคุณลักษณะอันพึงประสงค์',['1','2','3','4','5','6','7','8'], 'cS', 'cT', 'cR', 'theme-char')}
  ${mkTbl('ประเมินการอ่าน คิด วิเคราะห์ เขียน', ['1','2','3','4','5'], 'rS', 'rT', 'rR', 'theme-read')}
  ${mkTbl('ประเมินสมรรถนะสำคัญของผู้เรียน', ['1','2','3','4','5'], 'sS', 'sT', 'sR', 'theme-skill')}
</body>
</html>`;

  writeToPrintWindow(win, html, 'รายงานประเมิน');
}

// =====================================================
// 13. UI NAVIGATION & KEYBOARD EVENTS
// =====================================================
function switchHolisticTab(mode, btn) { $$('#holisticTabs .ttab').forEach(t => t.classList.remove('on')); if (btn) btn.classList.add('on'); $('holisticSubjectPanel').style.display = mode === 'subject' ? '' : 'none'; $('holisticHomeroomPanel').style.display = mode === 'homeroom' ? '' : 'none'; }
function showScoreMain(btn) { if($('scoreMainArea')) $('scoreMainArea').style.display='block'; if($('settingsContentArea')) $('settingsContentArea').style.display='none'; $$('.settings-panel').forEach(p=>p.classList.remove('active')); $$('#settingsTabs .settings-tab').forEach(t=>t.classList.remove('active')); if(btn) btn.classList.add('active'); }
function switchSettingsTab(tid, btn) { if($('scoreMainArea')) $('scoreMainArea').style.display='none'; if($('settingsContentArea')) $('settingsContentArea').style.display='block'; $$('.settings-panel').forEach(p=>p.classList.remove('active')); $$('#settingsTabs .settings-tab').forEach(t=>t.classList.remove('active')); if($('panel_'+tid)) $('panel_'+tid).classList.add('active'); if(btn) btn.classList.add('active'); }
function openSummaryModeModal() { if($('gClass').value.includes('เทอม')) return printSummaryPDF('term1'); $('summaryModeModal').style.display='flex'; }
function closeSummaryModeModal() { $('summaryModeModal').style.display='none'; }
$('summaryModeModal')?.addEventListener('click', function(e) { if(e.target===this) closeSummaryModeModal(); });

function updateScoreInputState(inp) {
  if (!inp) return; inp.classList.remove('score-empty', 'score-partial', 'score-full');
  const v = Number(inp.value), mx = Number(inp.getAttribute('max')) || 0;
  if (inp.value.trim() === '' || isNaN(v) || v <= 0) inp.classList.add('score-empty');
  else if (mx > 0 && v >= mx) inp.classList.add('score-full');
  else inp.classList.add('score-partial');
}
function refreshAllScoreInputStates(s = document) { s.querySelectorAll('.sub1, .sub2, .s-t1e, .s-t2e').forEach(updateScoreInputState); }

document.addEventListener('keydown', e => {
  if (!['Enter','ArrowDown','ArrowUp'].includes(e.key) || !document.activeElement.classList.contains('sinput')) return;
  e.preventDefault(); const tr = document.activeElement.closest('tr'), rows = Array.from(tr.closest('tbody').querySelectorAll('tr'));
  const target = rows[rows.indexOf(tr) + (e.key === 'ArrowUp' ? -1 : 1)]?.cells[document.activeElement.closest('td').cellIndex]?.querySelector('.sinput');
  if (target) { target.focus(); target.select(); }
});

  // =====================================================
async function fetchPp6Data() {
  return await api('getPp6FromSummary', {
    year: $('gYear').value,
    classroom: $('gClass').value
  });
}

window._doPrintPp5 = _doPrintPp5;

// =====================================================
// สรุปเกรดทุกชั้น (Admin)
// =====================================================
async function openAllGradesSummary() {
  const year = $('ag_year')?.value || $('gYear')?.value || '2568';
  const type = $('ag_term')?.value || 'year';
  const profile    = getSchoolProfileSafe();
  const schoolName = escH((profile && profile.school_name) || 'โรงเรียนบ้านคลอง 14');
  const termLabel  = type === 'term1' ? 'ภาคเรียนที่ 1'
    : type === 'term2' ? 'ภาคเรียนที่ 2' : 'ทั้งปี';

  const win = window.open('', '_blank');
  if (!win) return Utils.toast('กรุณาอนุญาต Popup ใน browser ด้วยครับ', 'error');
  win.document.write('<html><body style="font-family:sans-serif;padding:20px;"><p>⏳ กำลังโหลดข้อมูลทุกชั้น...</p></body></html>');

  const allCls  = CONFIG.ALL_CLS || [];
  const isMath  = c =>['ม.1','ม.2','ม.3'].some(m => String(c).startsWith(m));
  const FS      = 9; // font-size ใน table (px)

  Utils.showLoading('กำลังโหลดข้อมูลทุกชั้น...');

  // ดึง t1 + t2 ทุกชั้นพร้อมกัน
  const [resT1, resT2] = await Promise.all([
    Promise.allSettled(allCls.map(cls => api('getPp6FromSummary', { year, classroom: cls, mode: 'term1' }))),
    Promise.allSettled(allCls.map(cls => api('getPp6FromSummary', { year, classroom: cls, mode: 'term2' })))
  ]);
  Utils.hideLoading();

  // helper สร้าง map นักเรียน { id -> stuObj }
  const mkStuMap = (res, idx) => {
    const r = res[idx];
    if (r.status !== 'fulfilled') return {};
    const m = {};
    (r.value?.students || []).forEach(s => { m[String(s.id)] = s; });
    return m;
  };

  // helper รวบรวมชื่อวิชาจาก stuList
  const collectSubjs = (stuList) => {
    const names =[], seen = new Set();
    stuList.forEach(s => {
      [...(s.basicSubjects||[]), ...(s.addSubjects||[])].forEach(sb => {
        // FIX 1: ตัด suffix เลข 1-9 ออก เพื่อ merge วิชาเดียวกัน (รองรับ ม.1-ม.3)
        const normName = sb.name.replace(/\s+[1-9]$/, '').trim();
        if (!seen.has(normName)) { seen.add(normName); names.push(normName); }
      });
    });
    return names;
  };

  // สร้าง HTML ทุกชั้น
  const sections = allCls.map((cls, ci) => {
    const isMathCls = isMath(cls);
    const mapT1 = mkStuMap(resT1, ci);
    const mapT2 = mkStuMap(resT2, ci);

    // รวมนักเรียนจากทั้งสอง term
    const allIds  = [...new Set([...Object.keys(mapT1), ...Object.keys(mapT2)])];
    if (!allIds.length) return '';

    // ชื่อวิชา (normalized — ไม่มี suffix ตัวเลข)
    const allStuT1 = Object.values(mapT1);
    const allStuT2 = Object.values(mapT2);
    const subjNames = collectSubjs([...allStuT1, ...allStuT2]);
    if (!subjNames.length) return '';

    // COLS_PER_PAGE
    const COLS = Math.ceil(subjNames.length / Math.ceil(subjNames.length / 5));
    const chunks =[];
    for (let i = 0; i < subjNames.length; i += COLS) chunks.push(subjNames.slice(i, i+COLS));

    // FIX 2: ดึงข้อมูลแยกเทอมด้วยตรรกะ เลขคี่=เทอม 1, เลขคู่=เทอม 2
    const getMap = (stu, termNum) => {
      const m = {};
      [...(stu?.basicSubjects||[]), ...(stu?.addSubjects||[])].forEach(s => {
        const normName = s.name.replace(/\s+[1-9]$/, '').trim();
        const match    = s.name.match(/\s+([1-9])$/);
        const suffix   = match ? parseInt(match[1]) : null;

        let isMatchTerm = false;
        if (!suffix) {
          isMatchTerm = true; // ไม่มีเลข ถือว่าใช้ได้ทุกเทอม
        } else {
          // เช็คว่าเลขคี่(เทอม1) หรือ เลขคู่(เทอม2)
          const isOdd = (suffix % 2 !== 0);
          if (termNum === 1 && isOdd) isMatchTerm = true;
          if (termNum === 2 && !isOdd) isMatchTerm = true;
        }

        if (isMatchTerm) {
          m[normName] = s;
        } else if (!m[normName] && !suffix) {
          m[normName] = s;
        }
      });
      return m;
    };

    const gradeColor = g =>
      g === 'ร' ? 'color:#dc2626;font-weight:700;'
      : (g==='4'||g==='3.5'||g==='3') ? 'color:#166534;font-weight:700;'
      : g === '0' ? 'color:#dc2626;' : 'font-weight:700;';

    // เรียง id
    const sortedIds = allIds.sort((a,b) => String(a).localeCompare(String(b), undefined, {numeric:true}));

    return chunks.map((chunk, chunkIdx) => {
      const isLast = chunkIdx === chunks.length - 1;

      if (!isMathCls) {
        // ── ประถม ──
        const COL_PER_SUBJ = 10; 
        const hd1 = chunk.map(n =>
          '<th colspan="' + COL_PER_SUBJ + '" style="background:#1e3a5f;color:#fff;white-space:nowrap;padding:2px 4px;font-size:' + FS + 'px;">' + escH(n) + '</th>'
        ).join('') + (isLast ? '<th rowspan="3" style="background:#0f172a;color:#fff;width:40px;font-size:' + FS + 'px;">GPA</th>' : '');

        const hd2 =
          '<th colspan="3" style="background:#1d4ed8;color:#fff;font-size:' + FS + 'px;">เทอม 1</th>' +
          '<th colspan="3" style="background:#0369a1;color:#fff;font-size:' + FS + 'px;">เทอม 2</th>' +
          '<th colspan="3" style="background:#065f46;color:#fff;font-size:' + FS + 'px;">รวม 2 เทอม</th>' +
          '<th style="background:#166534;color:#fff;font-size:' + FS + 'px;">เกรด</th>';
        const hd2row = '<tr>' + chunk.map(() => hd2).join('') + '</tr>';

        const hd3 = chunk.map(() =>
          '<th style="font-size:' + (FS-1) + 'px;">เก็บ</th><th style="font-size:' + (FS-1) + 'px;">สอบ</th><th style="background:#dbeafe;font-size:' + (FS-1) + 'px;">รวม</th>' +
          '<th style="font-size:' + (FS-1) + 'px;">เก็บ</th><th style="font-size:' + (FS-1) + 'px;">สอบ</th><th style="background:#e0f2fe;font-size:' + (FS-1) + 'px;">รวม</th>' +
          '<th style="background:#d1fae5;font-size:' + (FS-1) + 'px;">เก็บ</th><th style="background:#d1fae5;font-size:' + (FS-1) + 'px;">สอบ</th><th style="background:#a7f3d0;font-weight:700;font-size:' + (FS-1) + 'px;">รวม</th>' +
          '<th style="background:#dcfce7;font-size:' + (FS-1) + 'px;color:#166534;">ผล</th>'
        ).join('');

        const bodyRows = sortedIds.map((id, ri) => {
          const s1 = mapT1[id], s2 = mapT2[id];
          const name = (s1||s2)?.name || id;
          const isR  = !!(s1||s2)?.hasR;
          const cells = chunk.map(n => {
            const d1 = getMap(s1, 1)[n], d2 = getMap(s2, 2)[n];
            const k1 = d1?.keep ?? '', ex1 = d1?.exam ?? '', tot1 = d1?.score ?? '';
            const k2 = d2?.keep ?? '', ex2 = d2?.exam ?? '', tot2 = d2?.score ?? '';
            
            // FIX 3: ซ่อนคำว่า "ติด ร" จากคอลัมน์คะแนน
            const dTot1 = String(tot1).includes('ร') ? '' : tot1;
            const dTot2 = String(tot2).includes('ร') ? '' : tot2;

            const sumK   = (k1  !== '' || k2  !== '') ? ((Number(k1 )||0) + (Number(k2 )||0)) : '';
            const sumEx  = (ex1 !== '' || ex2 !== '') ? ((Number(ex1)||0) + (Number(ex2)||0)) : '';
            let sumTot = (dTot1!== '' || dTot2!== '') ? ((Number(dTot1)||0) + (Number(dTot2)||0)) : '';
            
            const grade  = (isR || String(tot1).includes('ร') || String(tot2).includes('ร')) ? 'ร' : calcGradeNorm(sumTot, 100, false);
            if (grade === 'ร') sumTot = '';
            const gc = gradeColor(grade);

            return '<td>' + k1 + '</td><td>' + ex1 + '</td><td style="background:#dbeafe;font-weight:600;">' + dTot1 + '</td>' +
                   '<td>' + k2 + '</td><td>' + ex2 + '</td><td style="background:#e0f2fe;font-weight:600;">' + dTot2 + '</td>' +
                   '<td style="background:#d1fae5;">' + sumK + '</td><td style="background:#d1fae5;">' + sumEx + '</td>' +
                   '<td style="background:#a7f3d0;font-weight:700;">' + sumTot + '</td>' +
                   '<td style="background:#dcfce7;' + gc + '">' + escH(grade) + '</td>';
          }).join('');
          const gpa = (s1||s2)?.hasR ? 'รอผล' : ((s1||s2)?.gpa || '');
          const gpaStyle = (s1||s2)?.hasR ? 'color:#dc2626;font-weight:700;' : 'color:#166534;font-weight:700;';
          return '<tr style="' + (ri%2===1?'background:#f8fafc;':'') + '">' +
            '<td style="font-size:' + FS + 'px;">' + (ri+1) + '</td>' +
            '<td style="text-align:left;white-space:nowrap;padding:2px 4px;font-size:' + FS + 'px;">' + escH(name) + '</td>' +
            cells +
            (isLast ? '<td style="font-size:' + FS + 'px;' + gpaStyle + '">' + escH(gpa) + '</td>' : '') +
            '</tr>';
        }).join('');

        return '<div class="page-block">' +
          '<div style="text-align:center;font-size:12px;font-weight:700;margin-bottom:4px;">ชั้น ' + escH(cls) + ' | ปีการศึกษา ' + year + (chunks.length>1?' ('+(chunkIdx+1)+'/'+chunks.length+')':'') + '</div>' +
          '<table><thead>' +
          '<tr><th rowspan="3" style="background:#1e293b;color:#fff;width:26px;font-size:' + FS + 'px;">ที่</th>' +
          '<th rowspan="3" style="background:#1e293b;color:#fff;min-width:120px;font-size:' + FS + 'px;">ชื่อ-นามสกุล</th>' +
          hd1 + '</tr>' +
          hd2row +
          '<tr>' + hd3 + '</tr>' +
          '</thead><tbody>' + bodyRows + '</tbody></table></div>';

      } else {
        // ── มัธยม ──
        const COL_MATH = 8;
        const mathIds = [...new Set([...Object.keys(mapT1), ...Object.keys(mapT2)])]
          .sort((a,b) => String(a).localeCompare(String(b),undefined,{numeric:true}));
        if (!mathIds.length) return '';

        const hd1Math = chunk.map(n =>
          '<th colspan="' + COL_MATH + '" style="background:#1e3a5f;color:#fff;white-space:nowrap;padding:2px 4px;font-size:' + FS + 'px;">' + escH(n) + '</th>'
        ).join('') + (isLast ? '<th rowspan="3" style="background:#0f172a;color:#fff;width:40px;font-size:' + FS + 'px;">GPA</th>' : '');

        const hd2Math = chunk.map(() =>
          '<th colspan="4" style="background:#1d4ed8;color:#fff;font-size:' + FS + 'px;">เทอม 1</th>' +
          '<th colspan="4" style="background:#0369a1;color:#fff;font-size:' + FS + 'px;">เทอม 2</th>'
        ).join('');

        const hd3Math = chunk.map(() =>
          '<th style="font-size:' + (FS-1) + 'px;">เก็บ</th><th style="font-size:' + (FS-1) + 'px;">สอบ</th>' +
          '<th style="background:#dbeafe;font-size:' + (FS-1) + 'px;">รวม</th>' +
          '<th style="background:#fef3c7;color:#92400e;font-size:' + (FS-1) + 'px;">เกรด</th>' +
          '<th style="font-size:' + (FS-1) + 'px;">เก็บ</th><th style="font-size:' + (FS-1) + 'px;">สอบ</th>' +
          '<th style="background:#e0f2fe;font-size:' + (FS-1) + 'px;">รวม</th>' +
          '<th style="background:#fef9c3;color:#92400e;font-size:' + (FS-1) + 'px;">เกรด</th>'
        ).join('');

        const bodyRowsMath = mathIds.map((id, ri) => {
          const s1 = mapT1[id], s2 = mapT2[id];
          const name = (s1||s2)?.name || id;
          const isR  = !!(s1||s2)?.hasR;
          const sm1 = getMap(s1, 1), sm2 = getMap(s2, 2);
          
          const cells = chunk.map(n => {
            const d1 = sm1[n], d2 = sm2[n];
            const k1 = d1?.keep??'', ex1 = d1?.exam??'', tot1 = d1?.score??'';
            const k2 = d2?.keep??'', ex2 = d2?.exam??'', tot2 = d2?.score??'';
            
            // FIX 3: ซ่อนคำว่า "ติด ร" จากคอลัมน์คะแนนรวมให้เหลือว่างๆ
            const dTot1 = String(tot1).includes('ร') ? '' : tot1;
            const dTot2 = String(tot2).includes('ร') ? '' : tot2;

            // ดันเกรดเป็น ร ถ้าพบคำว่า ร ในคะแนน หรือ isR เป็นจริง
            const gr1 = (isR || String(tot1).includes('ร')) ? 'ร' : calcGradeNorm(dTot1, 100, false);
            const gr2 = (isR || String(tot2).includes('ร')) ? 'ร' : calcGradeNorm(dTot2, 100, false);
            const gc1 = gradeColor(gr1), gc2 = gradeColor(gr2);

            return '<td>' + k1 + '</td><td>' + ex1 + '</td>' +
                   '<td style="background:#dbeafe;font-weight:600;">' + dTot1 + '</td>' +
                   '<td style="background:#fef3c7;' + gc1 + '">' + escH(gr1) + '</td>' +
                   '<td>' + k2 + '</td><td>' + ex2 + '</td>' +
                   '<td style="background:#e0f2fe;font-weight:600;">' + dTot2 + '</td>' +
                   '<td style="background:#fef9c3;' + gc2 + '">' + escH(gr2) + '</td>';
          }).join('');
          
          const gpa = (s1||s2)?.hasR ? 'รอผล' : ((s2||s1)?.gpa || '');
          const gpaStyle = (s1||s2)?.hasR ? 'color:#dc2626;font-weight:700;' : 'color:#166534;font-weight:700;';
          
          return '<tr style="' + (ri%2===1?'background:#f8fafc;':'') + '">' +
            '<td style="font-size:' + FS + 'px;">' + (ri+1) + '</td>' +
            '<td style="text-align:left;white-space:nowrap;padding:2px 4px;font-size:' + FS + 'px;">' + escH(name) + '</td>' +
            cells +
            (isLast ? '<td style="font-size:' + FS + 'px;' + gpaStyle + '">' + escH(gpa) + '</td>' : '') +
            '</tr>';
        }).join('');

        return '<div class="page-block">' +
          '<div style="text-align:center;font-size:12px;font-weight:700;margin-bottom:4px;">ชั้น ' + escH(cls) + ' | ปีการศึกษา ' + year + (chunks.length>1?' ('+(chunkIdx+1)+'/'+chunks.length+')':'') + '</div>' +
          '<table><thead>' +
          '<tr><th rowspan="3" style="background:#1e293b;color:#fff;width:26px;font-size:' + FS + 'px;">ที่</th>' +
          '<th rowspan="3" style="background:#1e293b;color:#fff;min-width:120px;font-size:' + FS + 'px;">ชื่อ-นามสกุล</th>' +
          hd1Math + '</tr>' +
          '<tr>' + hd2Math + '</tr>' +
          '<tr>' + hd3Math + '</tr>' +
          '</thead><tbody>' + bodyRowsMath + '</tbody></table></div>';
      }
    }).join('');
  }).join('');

  const html = `<!DOCTYPE html><html lang="th"><head>
  <meta charset="UTF-8">
  <title>สรุปเกรดทุกชั้น ${termLabel} ปี${year}</title>
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
  <style>
    *{font-family:'Sarabun','Tahoma',sans-serif;box-sizing:border-box;}
    body{margin:0;padding:12px;font-size:10px;background:#f8fafc;}
    .page-block{margin-bottom:18px;background:#fff;border-radius:8px;box-shadow:0 1px 6px rgba(0,0,0,.08);padding:10px;overflow-x:auto;}
    table{border-collapse:collapse;width:100%;}
    th,td{border:1px solid #e2e8f0;padding:2px 3px;text-align:center;vertical-align:middle;}
    thead th{background:#334155;color:#fff;font-weight:700;position:sticky;top:0;z-index:5;}
    .print-btn{display:block;margin:0 auto 14px;padding:10px 28px;background:linear-gradient(135deg,#6366f1,#4338ca);color:#fff;border:none;border-radius:8px;font-size:.95rem;font-weight:700;cursor:pointer;font-family:inherit;}
    h2{text-align:center;font-size:14px;margin:0 0 4px;}
    h3{text-align:center;font-size:11px;font-weight:400;color:#64748b;margin:0 0 12px;}
    @media print{
      .print-btn{display:none!important;}
      body{padding:0;background:#fff;}
      .page-block{box-shadow:none;page-break-after:always;border-radius:0;overflow:visible;}
      .page-block:last-child{page-break-after:auto;}
      thead th{position:static!important;}
      @page{size:A4 landscape;margin:5mm 6mm;}
    }
  </style>
  </head><body>
  <button class="print-btn" onclick="window.print()">🖨️ พิมพ์ / บันทึก PDF</button>
  <h2>ตารางสรุปผลการเรียนทุกชั้น</h2>
  <h3>${termLabel} ปีการศึกษา ${year} | ${schoolName}</h3>
  ${sections || '<p style="text-align:center;color:#94a3b8;">ไม่พบข้อมูล</p>'}
  </body></html>`;

  writeToPrintWindow(win, html);
}

// =====================================================
// =====================================================
// Export เกรด CSV — เปิด modal เรียงวิชาก่อน
// =====================================================
let _csvExportData = null;

async function exportGradeCSV() {
  const year = $('ag_year')?.value || '2568';
  const type = $('ag_term')?.value || 'year';
  const cls  = $('ag_export_cls')?.value;
  if (!cls) return Utils.toast('กรุณาเลือกชั้นเรียน', 'error');

  const MATH_CLS2 = ['ม.1','ม.2','ม.3'];
  const isMath2   = MATH_CLS2.some(m => String(cls).startsWith(m));
  const mode      = isMath2 ? (type === 'year' ? 'term1' : type) : 'year';
  const termLabel = mode === 'term1' ? 'เทอม1' : mode === 'term2' ? 'เทอม2' : 'ทั้งปี';

  Utils.showLoading('กำลังโหลดข้อมูล...');
  let res;
  try {
    res = await api('getPp6FromSummary', { year, classroom: cls, mode });
  } catch(e) {
    Utils.hideLoading();
    return Utils.toast(e.message || 'โหลดข้อมูลไม่สำเร็จ', 'error');
  }
  Utils.hideLoading();

  const students = res.students || [];
  if (!students.length) return Utils.toast('ไม่พบข้อมูลนักเรียน', 'error');

  const subjNames = [], seen = new Set();
  students.forEach(stu => {
    [...(stu.basicSubjects||[]), ...(stu.addSubjects||[])].forEach(s => {
      if (!seen.has(s.name)) { seen.add(s.name); subjNames.push(s.name); }
    });
  });

  _csvExportData = { students, subjNames, cls, termLabel, year };
  _openCsvSortModal([...subjNames]);
}

function _openCsvSortModal(subjNames) {
  if (_csvExportData) _csvExportData.subjNames = subjNames;
  const list = document.getElementById('csvSortList');
  if (!list) return;

  list.innerHTML = subjNames.map((n, i) =>
    '<div class="csv-sort-item" draggable="true" data-idx="' + i + '"' +
    ' style="display:flex;align-items:center;gap:8px;background:#f8fafc;border:1.5px solid #e2e8f0;' +
    'border-radius:8px;padding:8px 10px;cursor:grab;user-select:none;">' +
    '<span style="color:#94a3b8;font-size:1rem;">⠿</span>' +
    '<span style="flex:1;font-size:.88rem;font-weight:600;">' + escH(n) + '</span>' +
    '<div style="display:flex;flex-direction:column;gap:2px;">' +
    '<button onclick="_csvMove(' + i + ',-1)" style="background:#f1f5f9;border:1px solid #cbd5e1;border-radius:4px;width:24px;height:20px;cursor:pointer;font-size:11px;line-height:1;">▲</button>' +
    '<button onclick="_csvMove(' + i + ',1)" style="background:#f1f5f9;border:1px solid #cbd5e1;border-radius:4px;width:24px;height:20px;cursor:pointer;font-size:11px;line-height:1;">▼</button>' +
    '</div></div>'
  ).join('');

  _initCsvDrag(list);
  document.getElementById('csvSortModal').style.display = 'flex';
}

function _csvMove(idx, dir) {
  if (!_csvExportData) return;
  const arr = _csvExportData.subjNames;
  const ni  = idx + dir;
  if (ni < 0 || ni >= arr.length) return;
  [arr[idx], arr[ni]] = [arr[ni], arr[idx]];
  _openCsvSortModal(arr);
}

function _initCsvDrag(list) {
  let dragIdx = null;
  list.querySelectorAll('.csv-sort-item').forEach(item => {
    item.addEventListener('dragstart', () => {
      dragIdx = parseInt(item.dataset.idx);
      setTimeout(() => item.style.opacity = '0.4', 0);
    });
    item.addEventListener('dragend',  () => { item.style.opacity = '1'; });
    item.addEventListener('dragover', e => { e.preventDefault(); item.style.borderColor = '#6366f1'; });
    item.addEventListener('dragleave',() => { item.style.borderColor = '#e2e8f0'; });
    item.addEventListener('drop', e => {
      e.preventDefault();
      item.style.borderColor = '#e2e8f0';
      const dropIdx = parseInt(item.dataset.idx);
      if (dragIdx === null || dragIdx === dropIdx) return;
      const arr = _csvExportData.subjNames;
      const [moved] = arr.splice(dragIdx, 1);
      arr.splice(dropIdx, 0, moved);
      _openCsvSortModal(arr);
    });
  });
}

function _doExportCSV() {
  if (!_csvExportData) return;
  const { students, subjNames, cls, termLabel, year } = _csvExportData;
  document.getElementById('csvSortModal').style.display = 'none';

  const sorted = [...students].sort((a, b) =>
    String(a.id).localeCompare(String(b.id), undefined, { numeric: true })
  );

  const csvQ = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  const headers = ['รหัสประจำตัว', 'ชื่อ-นามสกุล', ...subjNames, 'GPA'];
  const rows = sorted.map(stu => {
    const sm = {};
    [...(stu.basicSubjects||[]), ...(stu.addSubjects||[])].forEach(s => { sm[s.name] = s; });
    const grades = subjNames.map(n => sm[n]?.grade ?? '');
    return [stu.id, stu.name, ...grades, stu.hasR ? 'รอผล' : stu.gpa].map(csvQ).join(',');
  });

  const csv  = '\uFEFF' + [headers.map(csvQ).join(','), ...rows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = 'เกรด_' + cls + '_' + termLabel + '_' + year + '.csv';
  document.body.appendChild(a);
  a.click();
  a.remove();
  Utils.toast('✅ Export ' + sorted.length + ' คน สำเร็จ');
  _csvExportData = null;
}
