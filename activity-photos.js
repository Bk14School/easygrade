// =====================================================
// activity-photos.js — UI + Logic จัดการภาพกิจกรรม
// ใช้ร่วมกับ แนะแนว / ลูกเสือ / ชุมนุม
// =====================================================

// เก็บภาพที่โหลดมาแล้วใน memory: { 'แนะแนว_1': [...], ... }
const ActivityPhotos = {
  _cache: {},

  _key(actType, term) { return actType + '_' + term; },

  get(actType, term) { return this._cache[this._key(actType, term)] || []; },

  set(actType, term, photos) { this._cache[this._key(actType, term)] = photos || []; },

  add(actType, term, photo) {
    const key = this._key(actType, term);
    if (!this._cache[key]) this._cache[key] = [];
    this._cache[key].push(photo);
  },

  remove(actType, term, fileId) {
    const key = this._key(actType, term);
    this._cache[key] = (this._cache[key] || []).filter(p => p.fileId !== fileId);
  }
};

// ── แสดง UI อัปโหลดภาพทันที (ไม่รอ API) ─────────────
function initPhotoGrid(actType, term) {
  const containerId = _getPhotoContainerId(actType, term);
  const container   = document.getElementById(containerId);
  if (!container) return;
  // render shell พร้อมปุ่มอัปโหลดทันที ภาพจะโหลดทีหลัง
  _renderPhotoGrid(container, actType, term, [], false);
}

// ── โหลดภาพจาก Drive ────────────────────────────────
async function loadActivityPhotos(actType, term) {
  const containerId = _getPhotoContainerId(actType, term);
  const container   = document.getElementById(containerId);
  if (!container) return;

  // render shell ก่อน (มีปุ่มอัปโหลดให้ใช้ได้ทันที)
  if (!ActivityPhotos.get(actType, term).length) {
    _renderPhotoGrid(container, actType, term, [], false);
  }

  // เช็คว่ามีข้อมูลพร้อมก่อน call API
  const yr  = $('gYear')?.value;
  const cls = $('gClass')?.value;
  if (!yr || !cls) return; // ยังไม่พร้อม — ไม่ call API

  try {
    const photos = await api('getActivityPhotos', { year: yr, classroom: cls, actType, term });
    ActivityPhotos.set(actType, term, photos);
    _renderPhotoGrid(container, actType, term, photos, false);
  } catch(e) {
    // API ยังไม่มี route หรือ fail — ใช้งานได้ปกติ แค่ไม่โหลดภาพเก่า
    console.warn('loadActivityPhotos:', e.message);
  }
}

// ── อัปโหลดภาพ ───────────────────────────────────────
async function uploadActivityPhoto(actType, term, file) {
  if (!file) return;
  if (file.size > 5 * 1024 * 1024) return Utils.toast('ภาพใหญ่เกิน 5MB', 'error');

  const containerId = _getPhotoContainerId(actType, term);
  const container   = document.getElementById(containerId);

  // อ่านเป็น base64
  const base64 = await new Promise((res, rej) => {
    const r = new FileReader();
    r.onload  = () => res(r.result);
    r.onerror = rej;
    r.readAsDataURL(file);
  });

  Utils.showLoading('กำลังอัปโหลดภาพ...');
  try {
    const photo = await api('saveActivityPhoto', {
      year:      $('gYear').value,
      classroom: $('gClass').value,
      actType, term,
      filename: file.name,
      mimeType: file.type,
      base64: base64.split(',')[1]  // ตัด data:image/...;base64, ออก
    });
    ActivityPhotos.add(actType, term, photo);
    _renderPhotoGrid(container, actType, term, ActivityPhotos.get(actType, term), false);
    Utils.toast('✅ อัปโหลดภาพสำเร็จ');
  } catch(e) {
    Utils.toast(e.message || 'อัปโหลดไม่สำเร็จ', 'error');
  }
  Utils.hideLoading();
}

// ── ลบภาพ ────────────────────────────────────────────
async function deleteActivityPhoto(actType, term, fileId) {
  if (!confirm('ลบภาพนี้ออกจาก Google Drive?')) return;
  Utils.showLoading('กำลังลบ...');
  try {
    await api('deleteActivityPhoto', { fileId });
    ActivityPhotos.remove(actType, term, fileId);
    const containerId = _getPhotoContainerId(actType, term);
    const container   = document.getElementById(containerId);
    _renderPhotoGrid(container, actType, term, ActivityPhotos.get(actType, term), false);
    Utils.toast('✅ ลบภาพแล้ว');
  } catch(e) {
    Utils.toast(e.message || 'ลบไม่สำเร็จ', 'error');
  }
  Utils.hideLoading();
}

// ── helper: หา container id ──────────────────────────
function _getPhotoContainerId(actType, term) {
  const prefix = actType === 'แนะแนว' ? 'guidance'
               : actType === 'ลูกเสือ' ? 'scout'
               : 'club';
  return `photoGrid_${prefix}_${term}`;
}

// ── render grid ─────────────────────────────────────
function _renderPhotoGrid(container, actType, term, photos, loading) {
  if (!container) return;

  if (loading) {
    container.innerHTML = `<div style="color:#64748b;font-size:.8rem;padding:8px;">⏳ กำลังโหลดภาพ...</div>`;
    return;
  }

  const gridId = container.id;
  const photoItems = (photos || []).map(p => {
    const bigUrl = p.thumbUrl.replace('sz=w400', 'sz=w1600');
    return '<div style="position:relative;display:inline-block;margin:4px;">' +
      '<img src="' + p.thumbUrl + '" alt="ภาพกิจกรรม"' +
      ' style="width:120px;height:90px;object-fit:cover;border-radius:8px;border:2px solid #e2e8f0;cursor:pointer;"' +
      ' onclick="window.open(\'' + bigUrl + '\',\'_blank\')"' +
      ' onerror="this.style.opacity=0.3">' +
      '<button onclick="deleteActivityPhoto(\'' + actType + '\',\'' + term + '\',\'' + p.fileId + '\')"' +
      ' title="ลบภาพ"' +
      ' style="position:absolute;top:2px;right:2px;width:20px;height:20px;border-radius:50%;background:#ef4444;color:#fff;border:none;cursor:pointer;font-size:11px;line-height:1;display:flex;align-items:center;justify-content:center;padding:0;">✕</button>' +
      '</div>';
  }).join('');

  container.innerHTML = `
    <div style="background:#f8fafc;border:1.5px dashed #cbd5e1;border-radius:10px;padding:12px;margin-top:10px;">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;flex-wrap:wrap;gap:6px;">
        <span style="font-weight:700;font-size:.82rem;color:#475569;">📸 ภาพกิจกรรม (${(photos||[]).length} ภาพ)</span>
        <label style="cursor:pointer;padding:5px 12px;background:linear-gradient(135deg,#6366f1,#4338ca);color:#fff;border-radius:7px;font-size:.78rem;font-weight:700;">
          ➕ เพิ่มภาพ
          <input type="file" accept="image/*" multiple style="display:none;"
            onchange="_handlePhotoFiles(this,'${actType}','${term}')">
        </label>
      </div>
      ${photoItems
        ? `<div style="display:flex;flex-wrap:wrap;gap:4px;">${photoItems}</div>`
        : `<div style="color:#94a3b8;font-size:.78rem;text-align:center;padding:8px;">ยังไม่มีภาพ — กดเพิ่มภาพได้เลย</div>`}
    </div>`;
}

// ── handle multiple files ────────────────────────────
async function _handlePhotoFiles(input, actType, term) {
  const files = Array.from(input.files);
  for (const file of files) {
    await uploadActivityPhoto(actType, term, file);
  }
  input.value = '';
}

// ── สร้าง HTML section ภาพสำหรับ print ─────────────
function buildActivityPhotosHTML(actType, term) {
  const photos = ActivityPhotos.get(actType, term);
  if (!photos.length) return '';

  // thumbUrl ขนาด w800 สำหรับพิมพ์ (ไม่ต้อง login เหมือน url ปกติ)
  const bigThumb = (p) => p.thumbUrl.replace('sz=w400', 'sz=w800');

  // 2 ภาพต่อแถว
  const rows = [];
  for (let i = 0; i < photos.length; i += 2) {
    const pair = photos.slice(i, i + 2);
    const cells = pair.map(p =>
      `<td style="width:50%;padding:8px;text-align:center;vertical-align:top;">
        <img src="${bigThumb(p)}" style="width:100%;max-width:320px;height:240px;object-fit:cover;border:1px solid #ccc;border-radius:6px;">
      </td>`
    ).join('');
    rows.push(`<tr>${cells}</tr>`);
  }

  return `<div class="page" style="page-break-before:always;padding:16px;">
    <div style="text-align:center;font-weight:700;font-size:16px;margin-bottom:14px;">
      ภาพกิจกรรม${actType} ภาคเรียนที่ ${term}
    </div>
    <table style="width:100%;border-collapse:collapse;">
      ${rows.join('')}
    </table>
  </div>`;
}