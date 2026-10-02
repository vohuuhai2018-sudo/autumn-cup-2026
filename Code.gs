/**
 * AUTUMN CUP 2026 – Máy chủ nhận phiếu "VĐV được yêu thích nhất".
 * Chạy trên Google Apps Script, phiếu lưu vào một Google Sheet trong Drive của bạn.
 *
 * Làm 1 lần:
 *  1. Chọn hàm caiDat ở thanh trên cùng → bấm Chạy → cấp quyền.
 *     Xem "Nhật ký thực thi" để lấy link Google Sheet vừa tạo.
 *  2. Triển khai → Tùy chọn triển khai mới → loại: Ứng dụng web
 *       Thực thi dưới dạng: Tôi
 *       Người có quyền truy cập: Bất kỳ ai
 *  3. Chép URL ứng dụng web (đuôi /exec) dán vào API_URL trong index.html.
 *
 * Trong Google Sheet:
 *  - Tab VanDongVien: sửa tên, dán link ảnh (Google Drive để "Bất kỳ ai có đường liên kết"). Không sửa cột Mã.
 *  - Tab PhieuBau: mỗi dòng là 1 phiếu. Xoá dòng nào thì phiếu đó mất và thiết bị đó được bầu lại.
 *  - Tab CaiDat: bỏ tick "Mở bình chọn" để khoá; khi khoá ai vào link cũng xem được kết quả.
 *  - Muốn xoá sạch phiếu thử trước khi gửi link: chạy hàm xoaHetPhieu.
 */

const TEN_FILE = 'AUTUMN CUP 2026 – Bình chọn VĐV yêu thích';
const TAB_VDV = 'VanDongVien';
const TAB_PHIEU = 'PhieuBau';
const TAB_CAIDAT = 'CaiDat';
const GIAY_CACHE = 300;   // bảng tổng hợp giữ trong bộ nhớ tạm tối đa 5 phút
const GIAY_DOC_LAI = 45;  // quá 45 giây thì đọc lại Sheet để nhận phần sửa tay

const VDV_MAC_DINH = [
  'A. Tùng (dầu)', 'A. Tân', 'A. Hưng Motor', 'A. Hoàng', 'A. Quốc', 'A. Lành',
  'A. Hội', 'A. Khoa IT', 'A. Hải Landscape', 'A. Hiếu', 'A. Serge', 'A. Quang',
  'A. Đào', 'A. Hiệp', 'A. An', 'A. Phương APEC', 'A. Long Nam Long', 'A. Hải Interlab',
  'A. Lực', 'A. Thạch', 'A. Cường', 'A. Toàn', 'A. Dương', 'A. Thành'
];

/* ===================== CÀI ĐẶT ===================== */

function caiDat() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SHEET_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(TEN_FILE);
    props.setProperty('SHEET_ID', ss.getId());
    ss.getSheets()[0].setName(TAB_VDV);
  }

  const shV = ss.getSheetByName(TAB_VDV) || ss.insertSheet(TAB_VDV);
  if (shV.getLastRow() < 2) {
    shV.getRange('A:A').setNumberFormat('@');
    shV.getRange(1, 1, 1, 3).setValues([['Mã (không sửa)', 'Tên vận động viên', 'Link ảnh (không bắt buộc)']]);
    shV.getRange(2, 1, VDV_MAC_DINH.length, 2)
      .setValues(VDV_MAC_DINH.map((ten, i) => [String(i + 1).padStart(2, '0'), ten]));
    shV.setFrozenRows(1);
    shV.getRange(1, 1, 1, 3).setFontWeight('bold');
    shV.setColumnWidth(2, 220);
    shV.setColumnWidth(3, 420);
  }

  const shP = ss.getSheetByName(TAB_PHIEU) || ss.insertSheet(TAB_PHIEU);
  if (shP.getLastRow() < 1) {
    shP.getRange('B:C').setNumberFormat('@');
    shP.getRange(1, 1, 1, 5).setValues([['Thời gian', 'Mã thiết bị', 'Mã VĐV', 'Tên VĐV', 'Trình duyệt']]);
    shP.setFrozenRows(1);
    shP.getRange(1, 1, 1, 5).setFontWeight('bold');
    shP.setColumnWidth(1, 150);
    shP.setColumnWidth(2, 290);
    shP.setColumnWidth(4, 180);
  }

  const shC = ss.getSheetByName(TAB_CAIDAT) || ss.insertSheet(TAB_CAIDAT);
  if (shC.getLastRow() < 1) {
    shC.getRange(1, 1, 2, 3).setValues([
      ['Cài đặt', 'Giá trị', 'Ghi chú'],
      ['Mở bình chọn', '', 'Bỏ tick để khoá bình chọn. Khi khoá, ai vào link cũng xem được kết quả.']
    ]);
    shC.getRange('B2').insertCheckboxes().setValue(true);
    shC.getRange(1, 1, 1, 3).setFontWeight('bold');
    shC.setColumnWidth(1, 160);
    shC.setColumnWidth(3, 520);
  }

  xoaCache_();
  Logger.log('Đã cài đặt xong. Google Sheet: ' + ss.getUrl());
}

/** Xoá toàn bộ phiếu (dùng sau khi bầu thử). */
function xoaHetPhieu() {
  const sh = moSheet_().getSheetByName(TAB_PHIEU);
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.deleteRows(2, n);
  xoaCache_();
  Logger.log('Đã xoá ' + Math.max(n, 0) + ' phiếu.');
}

/* ===================== API ===================== */

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const device = hopLeThietBi_(p.device) ? p.device : '';
    return json_(traLoi_(layTongHop_(), device));
  } catch (err) {
    return json_({ ok: false, loi: 'may_chu', thongBao: 'Máy chủ gặp lỗi: ' + (err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action !== 'vote') return json_({ ok: false, loi: 'sai_yeu_cau', thongBao: 'Yêu cầu không hợp lệ.' });

    const device = String(body.device || '');
    if (!hopLeThietBi_(device)) {
      return json_({ ok: false, loi: 'sai_thiet_bi', thongBao: 'Không nhận ra thiết bị. Tải lại trang rồi thử lại.' });
    }
    const ma = chuanMa_(body.ma);

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) {
      return json_({ ok: false, loi: 'ban', thongBao: 'Đang đông người gửi, bấm gửi lại sau vài giây.' });
    }
    try {
      const d = layTongHop_();
      if (!d.mo) return json_(traLoi_(d, device, { ok: false, loi: 'dong', thongBao: 'Bình chọn đã đóng.' }));
      if (d.thietBi[device]) return json_(traLoi_(d, device, { ok: false, loi: 'da_bau', thongBao: 'Thiết bị này đã bình chọn rồi.' }));

      const vdv = d.vdv.find(v => v.ma === ma);
      if (!vdv) return json_({ ok: false, loi: 'sai_vdv', thongBao: 'Không tìm thấy vận động viên này. Tải lại trang rồi chọn lại.' });

      const ua = sachChu_(String(body.ua || '').slice(0, 200));
      moSheet_().getSheetByName(TAB_PHIEU).appendRow([new Date(), device, ma, vdv.ten, ua]);
      SpreadsheetApp.flush();

      d.thietBi[device] = ma;
      d.phieu[ma] = (d.phieu[ma] || 0) + 1;
      d.tong++;
      const cache = CacheService.getScriptCache();
      cache.put('phienban', String(Date.now()), 21600);
      try { cache.put('tonghop', JSON.stringify(d), GIAY_CACHE); } catch (e2) { cache.remove('tonghop'); }

      return json_(traLoi_(d, device));
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, loi: 'may_chu', thongBao: 'Máy chủ gặp lỗi: ' + (err && err.message || err) });
  }
}

/* ===================== NỘI BỘ ===================== */

function traLoi_(d, device, them) {
  const cuaToi = (device && d.thietBi[device]) || null;
  return Object.assign({
    ok: true,
    mo: d.mo,
    vdv: d.vdv,
    cuaToi: cuaToi,
    ketQua: (cuaToi || !d.mo) ? { phieu: d.phieu, tong: d.tong } : null,
    capNhat: new Date().toISOString()
  }, them || {});
}

/** Bản tổng hợp trong bộ nhớ tạm nếu còn mới, không thì đọc lại Sheet. */
function layTongHop_() {
  const c = CacheService.getScriptCache().get('tonghop');
  if (c) {
    try {
      const d = JSON.parse(c);
      if (d && d.docLuc && Date.now() - d.docLuc < GIAY_DOC_LAI * 1000) return d;
    } catch (e) {}
  }
  return docSheet_();
}

function docSheet_() {
  const cache = CacheService.getScriptCache();
  const phienBan = cache.get('phienban');
  const ss = moSheet_();

  const shV = ss.getSheetByName(TAB_VDV);
  const nV = shV.getLastRow() - 1;
  const vdv = nV > 0 ? shV.getRange(2, 1, nV, 3).getValues()
    .map(r => ({ ma: chuanMa_(r[0]), ten: String(r[1]).trim(), anh: String(r[2]).trim() }))
    .filter(v => v.ma && v.ten) : [];

  const shC = ss.getSheetByName(TAB_CAIDAT);
  const giaTri = shC ? shC.getRange('B2').getValue() : true;
  const mo = !(giaTri === false || String(giaTri).trim().toUpperCase() === 'FALSE');

  const hopLe = {};
  vdv.forEach(v => { hopLe[v.ma] = true; });
  const shP = ss.getSheetByName(TAB_PHIEU);
  const nP = shP.getLastRow() - 1;
  const rows = nP > 0 ? shP.getRange(2, 2, nP, 2).getValues() : [];
  const phieu = {}, thietBi = {};
  let tong = 0;
  rows.forEach(r => {
    const tb = String(r[0]).trim();
    const ma = chuanMa_(r[1]);
    if (!tb || !ma || thietBi[tb] || !hopLe[ma]) return;
    thietBi[tb] = ma;
    phieu[ma] = (phieu[ma] || 0) + 1;
    tong++;
  });

  const d = { vdv: vdv, mo: mo, phieu: phieu, tong: tong, thietBi: thietBi, docLuc: Date.now() };
  // Không ghi đè cache nếu có phiếu mới vừa vào trong lúc đang đọc
  if (cache.get('phienban') === phienBan) {
    try { cache.put('tonghop', JSON.stringify(d), GIAY_CACHE); } catch (e) {}
  }
  return d;
}

function moSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Chưa chạy hàm caiDat.');
  return SpreadsheetApp.openById(id);
}

function xoaCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove('tonghop');
  cache.put('phienban', String(Date.now()), 21600);
}

function chuanMa_(x) {
  let s = String(x == null ? '' : x).trim();
  if (/^\d+$/.test(s)) s = String(parseInt(s, 10)).padStart(2, '0');
  return s.slice(0, 20);
}

function hopLeThietBi_(s) {
  return typeof s === 'string' && /^[A-Za-z0-9-]{8,64}$/.test(s);
}

function sachChu_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
