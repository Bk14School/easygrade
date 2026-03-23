// =====================================================
// MODULE: การอ่าน คิดวิเคราะห์ และเขียน (RTW)
// =====================================================

// 1. ฟังก์ชันคำนวณคะแนนและเลื่อนเคอร์เซอร์อัตโนมัติ (แบบเลื่อนลงเมื่อพิมพ์ครบ)
function _rtwInput(el, mx) {
  // ควบคุมค่าสูงสุด-ต่ำสุด
  if (+el.value > mx) { 
    el.value = mx; 
    Utils.toast('เต็ม ' + mx + ' คะแนน', 'warning'); 
  }
  if (+el.value < 0) el.value = 0;

  // คำนวณผลรวมของแถวปัจจุบันทันที
  calcRTWRow(el);

  // ลอจิกเลื่อนช่องอัตโนมัติ (Vertical Auto-tab)
  let shouldTab = false;
  const valStr = el.value.toString();
  const valNum = parseInt(valStr, 10);

  if (valStr.length > 0) {
    if (mx < 10) {
      // กรณีคะแนนเต็มหลักเดียว พิมพ์ปุ๊บเลื่อนลงเลย
      shouldTab = true;
    } else {
      // กรณีคะแนนเต็ม 2 หลัก (เช่น 10, 15, 20)
      if (valStr.length >= 2) {
        shouldTab = true; 
      } else {
        // พิมพ์เพิ่งได้ 1 หลัก จะเลื่อนก็ต่อเมื่อพิมพ์ 0 หรือเลขที่เกินหลักแรกของคะแนนเต็ม
        const firstDigitOfMax = Math.floor(mx / 10);
        if (valNum === 0 || valNum > firstDigitOfMax) {
          shouldTab = true;
        }
      }
    }
  }

  // ทำการย้ายโฟกัสไปยังแถวถัดไป
  if (shouldTab && document.activeElement === el) {
    const tbody = el.closest('tbody');
    if (!tbody) return;
    const rows = [...tbody.querySelectorAll('tr[data-rtwsid]')];
    const curTr = el.closest('tr');
    const curIdx = rows.indexOf(curTr);
    
    // หาคลาสของช่องปัจจุบัน (เช่น rtw-c1-3) เพื่อโฟกัสช่องเดียวกันในแถวล่าง
    const currentClass = Array.from(el.classList).find(c => c.startsWith('rtw-'));
    
    const nextTr = rows[curIdx + 1];
    if (nextTr && currentClass) {
      const nextInput = nextTr.querySelector('.' + currentClass);
      if (nextInput) { 
        nextInput.focus(); 
        nextInput.select(); 
      }
    }
  }
}

// 2. ฟังก์ชันคำนวณคะแนนและตัดเกรดในแต่ละแถว
function calcRTWRow(input) {
  if (!input) return;
  const tr = input.closest('tr');
  const max = parseFloat(input.getAttribute('max'));
  if (parseFloat(input.value) > max) { 
    input.value = max; 
    Utils.toast(`คะแนนเต็ม ${max}`, 'warning'); 
  }
  
  const getV = (sel) => parseFloat(tr.querySelector(sel).value) || 0;
  
  // รวมการอ่าน
  const r1 = getV('.rtw-r1-1') + getV('.rtw-r1-2');
  const r2 = getV('.rtw-r2-1') + getV('.rtw-r2-2');
  tr.querySelector('.rtw-r1-sum').textContent = r1;
  tr.querySelector('.rtw-r2-sum').textContent = r2;

  // รวมการคิด
  const c1 = getV('.rtw-c1-3') + getV('.rtw-c1-4');
  const c2 = getV('.rtw-c2-3') + getV('.rtw-c2-4');
  tr.querySelector('.rtw-c1-sum').textContent = c1;
  tr.querySelector('.rtw-c2-sum').textContent = c2;

  // การเขียน
  const w1 = getV('.rtw-w1-5'), w2 = getV('.rtw-w2-5');

  // รวมทั้งหมด
  const grandTotal = r1 + r2 + c1 + c2 + w1 + w2;
  tr.querySelector('.rtw-grand-total').textContent = grandTotal;

  // ตัดเกรด
  const gBadge = tr.querySelector('.rtw-grade');
  let level = "ไม่ผ่าน (0)", cls = "res-0";
  if (grandTotal >= 86) { level = "ดีเยี่ยม (3)"; cls = "res-3"; }
  else if (grandTotal >= 70) { level = "ดี (2)"; cls = "res-2"; }
  else if (grandTotal >= 50) { level = "ผ่าน (1)"; cls = "res-1"; }
  gBadge.textContent = level; 
  gBadge.className = `res-badge rtw-grade ${cls}`;
}


// 3. ฟังก์ชันบันทึกข้อมูล (เพิ่มการจำข้อมูลลง Memory ทันที)
async function saveRTWData() {
  const records = [];
  $$('#rtwBody tr[data-rtwsid]').forEach(tr => {
    const getV = sel => parseFloat(tr.querySelector(sel)?.value) || 0;
    records.push({
      studentId : tr.getAttribute('data-rtwsid'),
      r1_1  : getV('.rtw-r1-1'),
      r1_2  : getV('.rtw-r1-2'),
      r2_1  : getV('.rtw-r2-1'),
      r2_2  : getV('.rtw-r2-2'),
      c1_3  : getV('.rtw-c1-3'),
      c1_4  : getV('.rtw-c1-4'),
      c2_3  : getV('.rtw-c2-3'),
      c2_4  : getV('.rtw-c2-4'),
      w1_5  : getV('.rtw-w1-5'),
      w2_5  : getV('.rtw-w2-5'),
      total : tr.querySelector('.rtw-grand-total')?.textContent || '0',
      grade : tr.querySelector('.rtw-grade')?.textContent || ''
    });
  });

  if (!records.length) return Utils.toast('ไม่พบข้อมูล', 'error');
  Utils.showLoading('กำลังบันทึก...');
  try {
    const res = await api('saveRTW', {
      year      : $('gYear').value,
      classroom : $('gClass').value,
      subject   : $('gSubj').value,
      records   : records
    });

    // 🌟 สิ่งที่เพิ่มเข้ามา: นำข้อมูลที่เพิ่งพิมพ์เสร็จ ยัดกลับเข้าไปในหน่วยความจำของระบบทันที
    // ทำให้เวลาสลับแท็บไปมา หรือสั่งพิมพ์ ข้อมูลจะถูกดึงมาโชว์โดยไม่ต้องโหลดหน้าเว็บใหม่
    records.forEach(rec => {
      const stu = App.students.find(s => s.studentId === rec.studentId);
      if (stu) {
        stu.rtw_data = rec; 
      }
    });

    Utils.toast('✅ ' + res);
  } catch (e) { 
    Utils.toast(e.message, 'error'); 
  }
  Utils.hideLoading();
}

// 4. ฟังก์ชันพิมพ์รายงาน (PDF)
function printRTWReport() {
  const cls = $('gClass').value;
  const subj = $('gSubj').value;
  const directorName = "นางสาวสู่ขวัญ ตลับนาค";
  const academicHead = $('sp_academic_head_name').value || '........................................................';

  // ดึงชื่อครูประจำชั้นจากช่องในแท็บนี้โดยตรง
  let rawTeacherName = $('rtw_homeroom_teacher').value.trim();
  if (!rawTeacherName) rawTeacherName = '........................................................';
  
  let teachers = rawTeacherName.split(/\s*(?:และ|\/|,)\s*/).filter(t => t);
  if (teachers.length === 0) teachers = ['........................................................'];
  const displayTeacherName = teachers.join(' และ '); 

  const rows = [...$$('#rtwBody tr[data-rtwsid]')].map((tr, i) => {
    const getV = (s) => parseFloat(tr.querySelector(s).value) || 0;
    const r1 = getV('.rtw-r1-1') + getV('.rtw-r1-2');
    const r2 = getV('.rtw-r2-1') + getV('.rtw-r2-2');
    const c1 = getV('.rtw-c1-3') + getV('.rtw-c1-4');
    const c2 = getV('.rtw-c2-3') + getV('.rtw-c2-4');
    const w1 = getV('.rtw-w1-5');
    const w2 = getV('.rtw-w2-5');

    return {
      no: i + 1, 
      name: tr.querySelector('.ass-name').textContent,
      r1, r2, r_total: r1 + r2,
      c1, c2, c_total: c1 + c2,
      w1, w2, w_total: w1 + w2,
      grand_total: tr.querySelector('.rtw-grand-total').textContent,
      grade: tr.querySelector('.rtw-grade').textContent
    };
  });

  if (!rows.length) return Utils.toast('ไม่พบข้อมูลนักเรียน', 'error');

  const getLvl = (score, max) => {
    const p = (score / max) * 100;
    if (p >= 86) return 'ex'; if (p >= 70) return 'g'; if (p >= 50) return 'p'; return 'i';
  };

  const stats = { r: { ex:0, g:0, p:0, i:0 }, c: { ex:0, g:0, p:0, i:0 }, w: { ex:0, g:0, p:0, i:0 } };
  rows.forEach(row => {
    stats.r[getLvl(row.r_total, 30)]++;
    stats.c[getLvl(row.c_total, 40)]++;
    stats.w[getLvl(row.w_total, 30)]++;
  });

  let indicators =[];
  if (cls.match(/[ป].[1-3]/)) {
    indicators =["1. อ่านและหาประสบการณ์จากสื่อที่หลากหลาย", "2. จับประเด็นสำคัญ ข้อเท็จจริง ความคิดเห็น", "3. เปรียบเทียบแง่มุมต่างๆ เช่น ข้อดี-เสีย ประโยชน์-โทษ", "4. แสดงความคิดเห็นต่อเรื่องที่อ่านอย่างมีเหตุผล", "5. ถ่ายทอดความรู้สึกจากเรื่องที่อ่านโดยการเขียน"];
  } else if (cls.match(/[ป].[4-6]/)) {
    indicators =["1. อ่านเพื่อหาข้อมูลสารสนเทศเสริมประสบการณ์", "2. จับประเด็นสำคัญ เชื่อมโยงความเป็นเหตุเป็นผล", "3. เชื่อมโยงความสัมพันธ์ของเรื่องราวและเหตุการณ์", "4. แสดงความคิดเห็นต่อเรื่องที่อ่านโดยมีเหตุผลสนับสนุน", "5. ถ่ายทอดความเข้าใจ ความคิดเห็น คุณค่าโดยการเขียน"];
  } else {
    indicators =["1. คัดสรรสื่อเพื่อหาข้อมูลสารสนเทศตามวัตถุประสงค์", "2. จับประเด็นสำคัญ ประเด็นสนับสนุน และโต้แย้ง", "3. วิเคราะห์ วิจารณ์ความสมเหตุสมผลน่าเชื่อถือ", "4. สรุปคุณค่า แนวคิด แง่คิดที่ได้จากการอ่าน", "5. สรุป อภิปราย ขยายความ แสดงความเห็นโต้แย้งโดยการเขียน"];
  }

  const win = window.open('', '_blank');
  const html = `
    <!DOCTYPE html>
    <html lang="th">
    <head>
      <meta charset="utf-8">
      <title>รายงานประเมินการอ่านคิดวิเคราะห์และเขียน</title>
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
      <style>
        @page { size: A4 portrait; margin: 15mm; }
        * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
        body { font-family: 'Sarabun', sans-serif; font-size: 15px; line-height: 1.5; color: #000; margin: 0; }
        .page { page-break-after: always; min-height: 260mm; position: relative; padding: 10px 20px; }
        
        .text-center { text-align: center; }
        .text-left { text-align: left; }
        .fw-bold { font-weight: bold; }
        
        .cover-logo { width: 95px; height: auto; margin-bottom: 15px; }
        .cover-title { font-size: 22px; font-weight: bold; margin-bottom: 5px; letter-spacing: 0.5px; }
        .cover-school { font-size: 20px; font-weight: bold; margin-bottom: 20px; }
        .cover-divider { border-top: 1px dashed #666; width: 60%; margin: 0 auto 25px; }
        
        .info-row { font-size: 16px; margin-bottom: 12px; }
        .fill-line { display: inline-block; border-bottom: 1.5px dotted #000; padding: 0 15px; color: #1e3a8a; font-weight: bold; }
        
        .summary-table { width: 95%; margin: 25px auto; border-collapse: collapse; box-shadow: 0 0 0 1px #000; }
        .summary-table th, .summary-table td { border: 1px solid #000; padding: 6px; text-align: center; font-size: 15px; }
        .summary-table th { background: #f8fafc; font-weight: bold; }
        
        .sign-flex { display: flex; justify-content: center; gap: 40px; margin-top: 20px; text-align: center; font-size: 15px; flex-wrap: wrap; }
        .sign-box { min-width: 250px; margin-bottom: 10px; }
        .sign-name { margin-top: 15px; }
        
        .approval-card { border: 2px solid #000; border-radius: 8px; width: 85%; margin: 30px auto 0; padding: 15px; text-align: center; }
        .check-box { display: inline-block; width: 14px; height: 14px; border: 1px solid #000; vertical-align: middle; margin-right: 8px; margin-top: -2px; }

        .score-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        .score-table th, .score-table td { border: 1px solid #000; padding: 6px 4px; text-align: center; font-size: 12px; }
        .score-table th { background: #f2f2f2; font-size: 11px; }
        .score-table .col-name { text-align: left; padding-left: 5px; width: 160px; }
      </style>
    </head>
    <body>

      <!-- ================= หน้า 1: ปก ================= -->
      <div class="page" style="padding:0;display:flex;flex-direction:column;min-height:260mm;">
        <div style="background:linear-gradient(135deg,#1e3a5f 0%,#2563eb 100%);color:#fff;padding:22px 30px;display:flex;align-items:center;gap:20px;">
          <img src="https://raw.githubusercontent.com/Bk14School/easygrademain/refs/heads/main/logo-OBEC.png"
               style="width:72px;height:auto;filter:brightness(0) invert(1);">
          <div>
            <div style="font-size:11px;letter-spacing:2px;opacity:.8;margin-bottom:4px;">สำนักงานคณะกรรมการการศึกษาขั้นพื้นฐาน</div>
            <div style="font-size:20px;font-weight:800;line-height:1.3;">แบบบันทึกผลการประเมิน</div>
            <div style="font-size:16px;font-weight:700;opacity:.9;">การอ่าน คิดวิเคราะห์ และเขียน</div>
          </div>
        </div>

        <div style="padding:20px 28px;flex:1;">
          <div style="background:#f0f9ff;border:1.5px solid #bae6fd;border-radius:10px;padding:14px 20px;margin-bottom:16px;">
            <div style="font-size:17px;font-weight:800;color:#0c4a6e;margin-bottom:10px;">🏫 โรงเรียนบ้านคลอง 14</div>
            <div style="display:flex;gap:12px;flex-wrap:wrap;">
              <div style="flex:1;min-width:180px;">
                <div style="font-size:11px;color:#64748b;margin-bottom:2px;">ชั้นเรียน</div>
                <div style="font-size:15px;font-weight:700;color:#1e293b;border-bottom:2px solid #0ea5e9;padding-bottom:3px;">${cls}</div>
              </div>
              <div style="flex:2;min-width:220px;">
                <div style="font-size:11px;color:#64748b;margin-bottom:2px;">ครูประจำชั้น / ครูที่ปรึกษา</div>
                <div style="font-size:15px;font-weight:700;color:#1e293b;border-bottom:2px solid #0ea5e9;padding-bottom:3px;">${displayTeacherName}</div>
              </div>
            </div>
          </div>

          <div style="font-size:13px;font-weight:800;color:#1e293b;margin-bottom:8px;">📊 สรุปผลการประเมิน</div>
          <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">
            <thead>
              <tr>
                <th style="border:1px solid #cbd5e1;padding:7px 6px;background:#1e3a5f;color:#fff;width:14%;">จำนวนนักเรียน</th>
                <th style="border:1px solid #cbd5e1;padding:7px 6px;background:#1e3a5f;color:#fff;width:26%;">มาตรฐาน</th>
                <th style="border:1px solid #cbd5e1;padding:7px 6px;background:#166534;color:#fff;">ดีเยี่ยม</th>
                <th style="border:1px solid #cbd5e1;padding:7px 6px;background:#1d4ed8;color:#fff;">ดี</th>
                <th style="border:1px solid #cbd5e1;padding:7px 6px;background:#d97706;color:#fff;">ผ่าน</th>
                <th style="border:1px solid #cbd5e1;padding:7px 6px;background:#dc2626;color:#fff;">ควรปรับปรุง</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td rowspan="3" style="border:1px solid #cbd5e1;text-align:center;font-size:22px;font-weight:800;color:#1e3a5f;">${rows.length}<br><span style="font-size:12px;font-weight:400;color:#64748b;">คน</span></td>
                <td style="border:1px solid #cbd5e1;padding:7px 10px;font-weight:700;">📖 การอ่าน (30)</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#f0fdf4;font-weight:700;color:#166534;">${stats.r.ex}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#eff6ff;font-weight:700;color:#1d4ed8;">${stats.r.g}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#fffbeb;font-weight:700;color:#d97706;">${stats.r.p}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#fff1f2;font-weight:700;color:#dc2626;">${stats.r.i}</td>
              </tr>
              <tr>
                <td style="border:1px solid #cbd5e1;padding:7px 10px;font-weight:700;">🧠 การคิดวิเคราะห์ (40)</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#f0fdf4;font-weight:700;color:#166534;">${stats.c.ex}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#eff6ff;font-weight:700;color:#1d4ed8;">${stats.c.g}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#fffbeb;font-weight:700;color:#d97706;">${stats.c.p}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#fff1f2;font-weight:700;color:#dc2626;">${stats.c.i}</td>
              </tr>
              <tr>
                <td style="border:1px solid #cbd5e1;padding:7px 10px;font-weight:700;">✍️ การเขียน (30)</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#f0fdf4;font-weight:700;color:#166534;">${stats.w.ex}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#eff6ff;font-weight:700;color:#1d4ed8;">${stats.w.g}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#fffbeb;font-weight:700;color:#d97706;">${stats.w.p}</td>
                <td style="border:1px solid #cbd5e1;text-align:center;background:#fff1f2;font-weight:700;color:#dc2626;">${stats.w.i}</td>
              </tr>
            </tbody>
          </table>

          <div style="font-size:13px;font-weight:800;color:#1e293b;margin-bottom:10px;">✍️ การรับรองผลการประเมิน</div>
          <div style="display:flex;gap:16px;flex-wrap:wrap;margin-bottom:14px;">
            ${teachers.map(t => `
              <div style="flex:1;min-width:160px;text-align:center;border:1px solid #e2e8f0;border-radius:8px;padding:12px;">
                <div style="border-bottom:1.5px solid #1e3a5f;margin:0 10px 8px;padding-bottom:2px;">&nbsp;</div>
                <div style="font-size:13px;font-weight:700;">( ${t} )</div>
                <div style="font-size:11px;color:#64748b;">ครูประจำชั้น</div>
              </div>`).join('')}
            <div style="flex:1;min-width:160px;text-align:center;border:1px solid #e2e8f0;border-radius:8px;padding:12px;">
              <div style="border-bottom:1.5px solid #1e3a5f;margin:0 10px 8px;padding-bottom:2px;">&nbsp;</div>
              <div style="font-size:13px;font-weight:700;">( ${academicHead} )</div>
              <div style="font-size:11px;color:#64748b;">ประธานกรรมการประเมินฯ</div>
            </div>
          </div>

          <div style="border:2px solid #1e3a5f;border-radius:10px;padding:14px 20px;text-align:center;">
            <div style="font-weight:800;font-size:14px;color:#1e3a5f;margin-bottom:10px;">การอนุมัติผลการประเมิน</div>
            <div style="display:flex;justify-content:center;gap:50px;margin-bottom:14px;font-size:14px;">
              <span><span style="display:inline-block;width:14px;height:14px;border:1.5px solid #000;border-radius:2px;vertical-align:middle;margin-right:6px;"></span>อนุมัติ</span>
              <span><span style="display:inline-block;width:14px;height:14px;border:1.5px solid #000;border-radius:2px;vertical-align:middle;margin-right:6px;"></span>ไม่อนุมัติ</span>
            </div>
            <div style="border-bottom:1.5px solid #1e3a5f;margin:0 40px 6px;padding-bottom:2px;">&nbsp;</div>
            <div style="font-size:13px;font-weight:700;">( ${directorName} )</div>
            <div style="font-size:11px;color:#64748b;">ผู้อำนวยการโรงเรียนบ้านคลอง 14</div>
          </div>

        </div>
      </div>

      <!-- ================= หน้า 2: ตารางคะแนนแยกครั้ง ================= -->
      <div class="page">
        <div class="text-center">
          <div class="fw-bold" style="font-size:18px;">สรุปคะแนนการประเมินการอ่าน คิดวิเคราะห์ และเขียน (แยกตามครั้ง)</div>
          <p style="font-size:15px; margin-top:5px;">วิชา ${subj} | ชั้น ${cls}</p>
        </div>
        <table class="score-table">
          <thead>
            <tr>
              <th rowspan="2" style="width:25px;">ที่</th><th rowspan="2" class="col-name">ชื่อ-นามสกุล</th>
              <th colspan="3" style="background:#eff6ff;">การอ่าน (30)</th>
              <th colspan="3" style="background:#fff7ed;">การคิดวิเคราะห์ (40)</th>
              <th colspan="3" style="background:#f0fdf4;">การเขียน (30)</th>
              <th rowspan="2" style="width:40px;">รวม<br>(100)</th><th rowspan="2" style="width:60px;">ระดับผล</th>
            </tr>
            <tr>
              <th style="background:#eff6ff; width:28px;">ค.1</th><th style="background:#eff6ff; width:28px;">ค.2</th><th style="background:#eff6ff; width:28px;">รวม</th>
              <th style="background:#fff7ed; width:28px;">ค.1</th><th style="background:#fff7ed; width:28px;">ค.2</th><th style="background:#fff7ed; width:28px;">รวม</th>
              <th style="background:#f0fdf4; width:28px;">ค.1</th><th style="background:#f0fdf4; width:28px;">ค.2</th><th style="background:#f0fdf4; width:28px;">รวม</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(r => `
              <tr>
                <td>${r.no}</td><td class="col-name">${r.name}</td>
                <td style="background:#eff6ff;">${r.r1}</td><td style="background:#eff6ff;">${r.r2}</td><td class="fw-bold" style="background:#eff6ff;">${r.r_total}</td>
                <td style="background:#fff7ed;">${r.c1}</td><td style="background:#fff7ed;">${r.c2}</td><td class="fw-bold" style="background:#fff7ed;">${r.c_total}</td>
                <td style="background:#f0fdf4;">${r.w1}</td><td style="background:#f0fdf4;">${r.w2}</td><td class="fw-bold" style="background:#f0fdf4;">${r.w_total}</td>
                <td class="fw-bold" style="font-size:13px;">${r.grand_total}</td><td class="fw-bold">${r.grade}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>

      <!-- ================= หน้า 3: เกณฑ์และตัวชี้วัด ================= -->
      <div class="page">
        <div class="text-center">
          <div class="fw-bold" style="font-size:18px;">เกณฑ์การประเมินและตัวชี้วัด</div>
          <p>ความสามารถในการอ่าน คิดวิเคราะห์ และเขียน (ชั้น ${cls})</p>
        </div>
        <div style="margin-left: 20px; margin-right: 20px;">
          <div class="fw-bold" style="text-decoration: underline;">ตัวชี้วัดที่ประเมิน:</div>
          <div style="margin: 10px 0 20px 10px; line-height: 1.8;">
            ${indicators.map((txt) => `<div>${txt}</div>`).join('')}
          </div>
          
          <div class="fw-bold" style="text-decoration: underline; margin-top: 20px;">สัดส่วนคะแนน (100 คะแนน):</div>
          <div style="margin: 10px 0 20px 10px; line-height: 1.8;">
            <div><b>1. การอ่าน (30 คะแนน)</b> ประเมินจากตัวชี้วัดข้อที่ 1 และ 2<br><span style="color:#555; font-size:13px; margin-left:15px;">- ครั้งที่ 1 (15 คะแนน) / ครั้งที่ 2 (15 คะแนน)</span></div>
            <div><b>2. การคิดวิเคราะห์ (40 คะแนน)</b> ประเมินจากตัวชี้วัดข้อที่ 3 และ 4<br><span style="color:#555; font-size:13px; margin-left:15px;">- ครั้งที่ 1 (20 คะแนน) / ครั้งที่ 2 (20 คะแนน)</span></div>
            <div><b>3. การเขียน (30 คะแนน)</b> ประเมินจากตัวชี้วัดข้อที่ 5<br><span style="color:#555; font-size:13px; margin-left:15px;">- ครั้งที่ 1 (15 คะแนน) / ครั้งที่ 2 (15 คะแนน)</span></div>
          </div>

          <div class="fw-bold" style="text-decoration: underline; margin-top: 30px;">เกณฑ์การตัดสินคุณภาพ:</div>
          <table style="width: 80%; margin: 10px auto; border-collapse: collapse;">
            <tr style="background:#f8fafc;">
              <th style="border:1px solid #000; padding:8px;">ช่วงคะแนนร้อยละ</th>
              <th style="border:1px solid #000; padding:8px;">ระดับคุณภาพ</th>
              <th style="border:1px solid #000; padding:8px;">ความหมาย</th>
            </tr>
            <tr><td style="border:1px solid #000; text-align:center;">86 - 100</td><td style="border:1px solid #000; text-align:center;">ดีเยี่ยม (3)</td><td class="text-left" style="border:1px solid #000; padding-left:10px;">มีผลงานแสดงถึงความสามารถที่มีคุณภาพดีเลิศอยู่เสมอ</td></tr>
            <tr><td style="border:1px solid #000; text-align:center;">70 - 85</td><td style="border:1px solid #000; text-align:center;">ดี (2)</td><td class="text-left" style="border:1px solid #000; padding-left:10px;">มีผลงานแสดงถึงความสามารถที่มีคุณภาพเป็นที่ยอมรับ</td></tr>
            <tr><td style="border:1px solid #000; text-align:center;">50 - 69</td><td style="border:1px solid #000; text-align:center;">ผ่าน (1)</td><td class="text-left" style="border:1px solid #000; padding-left:10px;">มีผลงานแสดงถึงความสามารถที่มีข้อบกพร่องบางประการ</td></tr>
            <tr><td style="border:1px solid #000; text-align:center;">0 - 49</td><td style="border:1px solid #000; text-align:center;">ไม่ผ่าน (0)</td><td class="text-left" style="border:1px solid #000; padding-left:10px;">ไม่มีผลงาน หรือผลงานต้องได้รับการปรับปรุงหลายประการ</td></tr>
          </table>
          
          <div style="margin-top: 30px; font-size: 13px; font-style: italic; color: #555; text-align:center;">
            * การประเมินอ้างอิงตามหลักสูตรแกนกลางการศึกษาขั้นพื้นฐาน พุทธศักราช 2551
          </div>
        </div>
      </div>
    </body>
    </html>
  `;
  win.document.open(); win.document.write(html); win.document.close();
}

// 5. เรนเดอร์ตารางประเมินลงในหน้าเว็บ
function renderRTWTable() {
  const container = $('rtwContainer');
  if (!App.students.length) return;

  const cls = $('gClass').value.trim();
  
  // ระบบจดจำชื่อครูประจำชั้นแยกตามห้องเรียน
  App.hrMap = App.hrMap || {};
  let hrTeacherStr = App.hrMap[cls];
  if (hrTeacherStr === undefined) {
    hrTeacherStr = (App.user && App.user.classroom === cls) ? App.user.name : "";
  }

  let indicators =[];
  let title = "";
  if (cls.match(/[ป].[1-3]/)) {
    title = "ระดับชั้นประถมศึกษาปีที่ 1-3";
    indicators =["1. อ่านและหาประสบการณ์จากสื่อที่หลากหลาย", "2. จับประเด็นสำคัญ ข้อเท็จจริง ความคิดเห็น", "3. เปรียบเทียบแง่มุมต่างๆ (ข้อดี-เสีย/ประโยชน์-โทษ)", "4. แสดงความคิดเห็นต่อเรื่องที่อ่านอย่างมีเหตุผล", "5. ถ่ายทอดความรู้สึกจากเรื่องที่อ่านโดยการเขียน"];
  } else if (cls.match(/[ป].[4-6]/)) {
    title = "ระดับชั้นประถมศึกษาปีที่ 4-6";
    indicators =["1. อ่านเพื่อหาข้อมูลสารสนเทศเสริมประสบการณ์", "2. จับประเด็นสำคัญ เชื่อมโยงความเป็นเหตุผล", "3. เชื่อมโยงความสัมพันธ์ของเรื่องราวและเหตุการณ์", "4. แสดงความคิดเห็นโดยมีเหตุผลสนับสนุน", "5. ถ่ายทอดความคิดเห็นและคุณค่าโดยการเขียน"];
  } else {
    title = "ระดับชั้นมัธยมศึกษาปีที่ 1-3";
    indicators =["1. คัดสรรสื่อเพื่อหาข้อมูลสารสนเทศตามวัตถุประสงค์", "2. จับประเด็นสำคัญ ประเด็นสนับสนุน และโต้แย้ง", "3. วิเคราะห์ วิจารณ์ความสมเหตุสมผลน่าเชื่อถือ", "4. สรุปคุณค่า แนวคิด แง่คิดจากการอ่าน", "5. สรุป อภิปราย ขยายความ แสดงความเห็นโต้แย้งโดยการเขียน"];
  }

  const criteriaHtml = `
    <div style="background:#f8fafc; border:1px solid #cbd5e1; border-radius:8px; padding:15px; margin-bottom:15px; box-shadow:0 2px 4px rgba(0,0,0,0.02);">
      
      <div style="display:flex; align-items:center; gap:10px; margin-bottom:15px; background:#fff; padding:10px 15px; border-radius:6px; border:1px solid #e2e8f0;">
        <label style="font-weight:bold; color:#1e40af; margin:0; white-space:nowrap; font-size:0.9rem;">👨‍🏫 ครูประจำชั้น:</label>
        <input type="text" id="rtw_homeroom_teacher" value="${hrTeacherStr}" 
               oninput="App.hrMap[document.getElementById('gClass').value.trim()] = this.value; if(document.getElementById('guidance_teacher')) document.getElementById('guidance_teacher').value = this.value;" 
               style="flex:1; max-width:400px; padding:6px 10px; border:1px solid #94a3b8; border-radius:4px; font-weight:bold; font-family:inherit; font-size:0.9rem;" placeholder="นาย ก / นาง ข">
        <span style="font-size:0.75rem; color:#64748b;">(ระบบจำชื่อแยกตามชั้นอัตโนมัติ)</span>
      </div>

      <div style="font-weight:bold; color:#0f172a; margin-bottom:10px; font-size:0.95rem; border-bottom:1px solid #e2e8f0; padding-bottom:8px;">
        📖 เกณฑ์การประเมินและตัวชี้วัด (${title})
      </div>
      <div style="display:flex; flex-wrap:wrap; gap:20px; font-size:0.85rem;">
          <div style="flex:1.5; min-width:300px;">
              <div style="font-weight:bold; color:#334155; margin-bottom:5px;">📌 ตัวชี้วัดที่ประเมิน:</div>
              <div style="color:#475569; line-height:1.6; margin-left:10px;">
                  ${indicators.map(i => `<div>${i}</div>`).join('')}
              </div>
          </div>
          <div style="flex:1; min-width:260px;">
              <div style="font-weight:bold; color:#334155; margin-bottom:5px;">📊 สัดส่วนคะแนน (100 คะแนน):</div>
              <div style="color:#475569; line-height:1.6; margin-left:10px; margin-bottom:12px;">
                  <div><b style="color:#1d4ed8;">การอ่าน (30):</b> วัดข้อ 1-2 (ครั้งละ 15)</div>
                  <div><b style="color:#c2410c;">การคิด (40):</b> วัดข้อ 3-4 (ครั้งละ 20)</div>
                  <div><b style="color:#15803d;">การเขียน (30):</b> วัดข้อ 5 (ครั้งละ 15)</div>
              </div>
          </div>
      </div>
    </div>
  `;
  
  $('rtwIndicatorLabel').innerHTML = criteriaHtml;
  $('rtwIndicatorLabel').className = "";
  $('rtwIndicatorLabel').style.border = "none";
  $('rtwIndicatorLabel').style.background = "transparent";

  const htmlRows = App.students.map((s, idx) => {
    const r = s.rtw_data || {};
    return `
      <tr data-rtwsid="${s.studentId}">
        <td class="ass-no">${idx + 1}</td>
        <td class="ass-name" style="text-align:left; background:#fff; position:sticky; left:0; z-index:10;">${s.name}</td>
        <td><input type="number" class="inp-ass rtw-r1-1" min="0" max="8" value="${r.r1_1||''}" oninput="_rtwInput(this,8)"></td>
        <td><input type="number" class="inp-ass rtw-r1-2" min="0" max="7" value="${r.r1_2||''}" oninput="_rtwInput(this,7)"></td>
        <td class="bg-sum rtw-r1-sum">-</td>
        <td><input type="number" class="inp-ass rtw-r2-1" min="0" max="8" value="${r.r2_1||''}" oninput="_rtwInput(this,8)"></td>
        <td><input type="number" class="inp-ass rtw-r2-2" min="0" max="7" value="${r.r2_2||''}" oninput="_rtwInput(this,7)"></td>
        <td class="bg-sum rtw-r2-sum">-</td>
        <td><input type="number" class="inp-ass rtw-c1-3" min="0" max="10" value="${r.c1_3||''}" oninput="_rtwInput(this,10)"></td>
        <td><input type="number" class="inp-ass rtw-c1-4" min="0" max="10" value="${r.c1_4||''}" oninput="_rtwInput(this,10)"></td>
        <td class="bg-sum rtw-c1-sum">-</td>
        <td><input type="number" class="inp-ass rtw-c2-3" min="0" max="10" value="${r.c2_3||''}" oninput="_rtwInput(this,10)"></td>
        <td><input type="number" class="inp-ass rtw-c2-4" min="0" max="10" value="${r.c2_4||''}" oninput="_rtwInput(this,10)"></td>
        <td class="bg-sum rtw-c2-sum">-</td>
        <td><input type="number" class="inp-ass rtw-w1-5" min="0" max="15" value="${r.w1_5||''}" oninput="_rtwInput(this,15)"></td>
        <td><input type="number" class="inp-ass rtw-w2-5" min="0" max="15" value="${r.w2_5||''}" oninput="_rtwInput(this,15)"></td>
        <td class="bg-sum rtw-grand-total" style="font-weight:bold; background:#e0f2fe!important; color:#0369a1;">-</td>
        <td><span class="res-badge rtw-grade">-</span></td>
      </tr>`;
  }).join('');

  container.innerHTML = `
    <div class="ass-wrap">
      <table class="ass-tbl" id="rtwTable" style="font-size:12px;">
        <thead>
          <tr style="background:#f8fafc;">
            <th rowspan="3">ที่</th><th rowspan="3" style="min-width:160px; z-index:20;">ชื่อ-นามสกุล</th>
            <th colspan="6" style="background:#dbeafe;">1. การอ่าน (30)</th>
            <th colspan="6" style="background:#ffedd5;">2. การคิดวิเคราะห์ (40)</th>
            <th colspan="2" style="background:#dcfce7;">3. การเขียน (30)</th>
            <th rowspan="3">รวม</th><th rowspan="3">ระดับ</th>
          </tr>
          <tr>
            <th colspan="3">ครั้งที่ 1 (15)</th><th colspan="3">ครั้งที่ 2 (15)</th>
            <th colspan="3">ครั้งที่ 1 (20)</th><th colspan="3">ครั้งที่ 2 (20)</th>
            <th>ค1(15)</th><th>ค2(15)</th>
          </tr>
          <tr style="font-size:10px;">
            <th style="background:#eff6ff;">ข้อ 1(8)</th><th style="background:#eff6ff;">ข้อ 2(7)</th><th class="bg-sum">รวม</th>
            <th style="background:#eff6ff;">ข้อ 1(8)</th><th style="background:#eff6ff;">ข้อ 2(7)</th><th class="bg-sum">รวม</th>
            <th style="background:#fff7ed;">ข้อ 3(10)</th><th style="background:#fff7ed;">ข้อ 4(10)</th><th class="bg-sum">รวม</th>
            <th style="background:#fff7ed;">ข้อ 3(10)</th><th style="background:#fff7ed;">ข้อ 4(10)</th><th class="bg-sum">รวม</th>
            <th style="background:#f0fdf4;">ข้อ 5(15)</th><th style="background:#f0fdf4;">ข้อ 5(15)</th>
          </tr>
        </thead>
        <tbody id="rtwBody">${htmlRows}</tbody>
      </table>
    </div>`;
    
  $$('#rtwBody tr[data-rtwsid]').forEach(tr => calcRTWRow(tr.querySelector('.rtw-r1-1')));
}


// =====================================================
// ระบบเลื่อนช่องกรอกคะแนนด้วยลูกศรคีย์บอร์ด (ซ้าย ขวา ขึ้น ลง)
// ครอบคลุมทั้งหน้าตารางคะแนนหลัก และหน้าประเมิน/RTW
// =====================================================
document.addEventListener('keydown', e => {
  const activeEl = document.activeElement;
  
  // เช็คว่าช่องที่กำลังพิมพ์อยู่เป็นช่องกรอกคะแนนหรือไม่ (.sinput = เกรดหลัก, .inp-ass = RTW/ประเมิน)
  if (!activeEl || (!activeEl.classList.contains('sinput') && !activeEl.classList.contains('inp-ass'))) return;

  // ดักจับเฉพาะปุ่มลูกศรและปุ่ม Enter
  const keys =['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Enter'];
  if (!keys.includes(e.key)) return;

  e.preventDefault(); // ป้องกันหน้าจอเลื่อนตามลูกศร

  const tr = activeEl.closest('tr');
  const tbody = tr.closest('tbody');
  if (!tr || !tbody) return;

  const rows = Array.from(tbody.querySelectorAll('tr'));
  const rowIndex = rows.indexOf(tr);
  
  // ค้นหาช่องกรอกคะแนนทั้งหมดในแถวเดียวกัน (จะกระโดดข้ามช่องผลรวมให้อัตโนมัติ)
  const rowInputs = Array.from(tr.querySelectorAll('.sinput, .inp-ass'));
  const colIndex = rowInputs.indexOf(activeEl);

  let targetInput = null;

  if (e.key === 'ArrowUp') {
    const prevRow = rows[rowIndex - 1];
    if (prevRow) targetInput = prevRow.querySelectorAll('.sinput, .inp-ass')[colIndex];
  } 
  else if (e.key === 'ArrowDown' || e.key === 'Enter') {
    const nextRow = rows[rowIndex + 1];
    if (nextRow) targetInput = nextRow.querySelectorAll('.sinput, .inp-ass')[colIndex];
  } 
  else if (e.key === 'ArrowLeft') {
    if (colIndex > 0) targetInput = rowInputs[colIndex - 1];
  } 
  else if (e.key === 'ArrowRight') {
    if (colIndex < rowInputs.length - 1) targetInput = rowInputs[colIndex + 1];
  }

  if (targetInput) {
    targetInput.focus();
    targetInput.select();
  }
});
