// ===== CONFIG =====
const API = 'https://phan-d-default-rtdb.asia-southeast1.firebasedatabase.app/don_doi_ca';

// ===== STATE =====
window.rec = [];

// ===== CLOCK =====
function updateClock() {
  document.getElementById('clock').textContent = new Date().toLocaleTimeString('vi-VN');
}
updateClock();
setInterval(updateClock, 1000);

// ===== HELPERS =====
function fd(s) {
  if (!s) return '';
  const p = s.split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}

function tc(s) {
  return s.toLowerCase().split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}

function cl(s) {
  if (!s) return '............................';
  const t = s.replace(/\./g, '').trim();
  return t || '.............................';
}

window.fd = fd;
window.tc = tc;

// ===== LOAD DATA =====
async function load() {
  try {
    const r = await fetch(API + '.json');
    const d = await r.json();
    let a = [];
    if (d && typeof d === 'object') {
      for (let k in d) a.push({ id: k, ...d[k] });
    }
    a.sort((x, y) => new Date(y.createdAt || 0) - new Date(x.createdAt || 0));
    window.rec = a;

    const tb = document.getElementById('tbody');
    if (!a.length) {
      tb.innerHTML = '<tr><td colspan="5" class="empty">Chưa có đơn nào</td></tr>';
      return;
    }

    tb.innerHTML = '';
    a.forEach(i => {
      const lb = i.loai === 'truongca'
        ? '<span class="badge badge-tc">⭐ Trưởng ca</span>'
        : '<span class="badge badge-nv">👤 Nhân viên</span>';

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td data-label="Ngày tạo">
          ${new Date(i.createdAt).toLocaleDateString('vi-VN')}<br>
          <small style="color:var(--muted)">${new Date(i.createdAt).toLocaleTimeString('vi-VN')}</small>
        </td>
        <td data-label="Người xin đổi"><strong>${i.nguoiXin}</strong><br>${lb}</td>
        <td data-label="Trực thay"><strong>${i.nguoiDoi}</strong></td>
        <td data-label="Chi tiết ca">
          <div class="ca-box"><span style="color:var(--red);font-size:.75rem">▲</span> ${i.caDi}</div>
          <div class="ca-box"><span style="color:var(--green);font-size:.75rem">▼</span> ${i.caVe}</div>
        </td>
        <td data-label="Thao tác" class="td-act">
          <button class="bw" onclick="expOld('${i.id}')">📄 Word</button>
          <button class="bd" onclick="del('${i.id}')">🗑</button>
        </td>`;
      tb.appendChild(tr);
    });
  } catch (e) {
    document.getElementById('tbody').innerHTML =
      '<tr><td colspan="5" class="empty" style="color:var(--red)">Lỗi kết nối!</td></tr>';
  }
}

// ===== BIND PRINT PREVIEW =====
function bind() {
  const dots = '................................................................';
  const dotsCa = '.................................................';
  const nx = document.getElementById('nx').value || dots;
  const nd = document.getElementById('nd').value || dots;
  const loai = document.querySelector('input[name="loai"]:checked').value;

  document.getElementById('p_nx').textContent = nx;
  document.getElementById('p_cd').textContent = loai === 'truongca' ? 'trưởng ca' : 'nhân viên';
  ['p_nd', 'p_nd2', 'p_nd3'].forEach(id => document.getElementById(id).textContent = nd);

  const cdi = document.getElementById('cdi').value;
  const ddi = document.getElementById('ddi').value;
  const cve = document.getElementById('cve').value;
  const dve = document.getElementById('dve').value;

  document.getElementById('p_c1').textContent = (cdi && ddi) ? cdi + ' ngày ' + fd(ddi) : dotsCa;
  document.getElementById('p_c2').textContent = (cve && dve) ? cve + ' ngày ' + fd(dve) : dotsCa;
  document.getElementById('p_sl').textContent = document.getElementById('sl').value || '......';
  document.getElementById('p_ld').textContent = document.getElementById('ld').value || '...............................................................................';
  document.getElementById('p_nxk').textContent = document.getElementById('nx').value
    ? tc(document.getElementById('nx').value) : '............................';
  document.getElementById('p_ndk').textContent = document.getElementById('nd').value
    ? tc(document.getElementById('nd').value) : '............................';

  const now = new Date();
  document.getElementById('p_ng').textContent = now.getDate().toString().padStart(2, '0');
  document.getElementById('p_th').textContent = (now.getMonth() + 1).toString().padStart(2, '0');
  document.getElementById('p_na').textContent = now.getFullYear();
}

// ===== LISTEN INPUTS =====
document.querySelectorAll('input, select, textarea').forEach(e => {
  e.addEventListener('input', bind);
  e.addEventListener('change', bind);
});

// ===== PRINT =====
document.getElementById('btnPrint').onclick = () => {
  bind();
  window.print();
};

// ===== EXPORT WORD =====
document.getElementById('btnWord').onclick = () => {
  bind();
  const now = new Date();
  expDocx({
    nx: document.getElementById('nx').value || '...',
    nd: document.getElementById('nd').value || '...',
    cdi: document.getElementById('p_c1').textContent,
    cve: document.getElementById('p_c2').textContent,
    sl: document.getElementById('sl').value || '...',
    ld: document.getElementById('ld').value || '...',
    loai: document.querySelector('input[name="loai"]:checked').value,
    ng: now.getDate().toString().padStart(2, '0'),
    th: (now.getMonth() + 1).toString().padStart(2, '0'),
    na: now.getFullYear().toString()
  }, 'Don_Doi_Ca_' + tc(document.getElementById('nx').value || 'Moi').replace(/ /g, '_'));
};

// ===== SAVE FORM =====
document.getElementById('frm').addEventListener('submit', async function (e) {
  e.preventDefault();
  const o = {
    nguoiXin: document.getElementById('nx').value,
    nguoiDoi: document.getElementById('nd').value,
    caDi: document.getElementById('cdi').value + ' ngày ' + fd(document.getElementById('ddi').value),
    caVe: document.getElementById('cve').value + ' ngày ' + fd(document.getElementById('dve').value),
    // Note: option values already include 'Ca 1', 'Ca 2', etc.
    soLan: document.getElementById('sl').value,
    lyDo: document.getElementById('ld').value,
    loai: document.querySelector('input[name="loai"]:checked').value,
    createdAt: new Date().toISOString()
  };
  try {
    const r = await fetch(API + '.json', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(o)
    });
    if (!r.ok) throw 0;
    this.reset();
    bind();
    load();
    alert('✅ Đã lưu thành công!');
  } catch {
    alert('❌ Lưu thất bại!');
  }
});

// ===== EXPORT OLD RECORD =====
window.expOld = function (id) {
  const r = window.rec.find(x => x.id === id);
  if (!r) return;
  const d = new Date(r.createdAt);
  expDocx({
    nx: r.nguoiXin, nd: r.nguoiDoi,
    cdi: r.caDi, cve: r.caVe,
    sl: r.soLan, ld: r.lyDo,
    loai: r.loai || 'nhanvien',
    ng: d.getDate().toString().padStart(2, '0'),
    th: (d.getMonth() + 1).toString().padStart(2, '0'),
    na: d.getFullYear().toString()
  }, 'Don_Doi_Ca_' + tc(r.nguoiXin || 'Moi').replace(/ /g, '_'));
};

// ===== DELETE RECORD =====
window.del = async function (id) {
  if (!confirm('Xóa đơn này?')) return;
  try {
    await fetch(API + '/' + id + '.json', { method: 'DELETE' });
    load();
  } catch {
    alert('Lỗi xóa!');
  }
};

// ===== EXPORT DOCX =====
async function expDocx(d, fn) {
  if (typeof JSZip === 'undefined') { alert('JSZip chưa tải!'); return; }

  function esc(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function run(t, o) {
    o = o || {};
    return '<w:r><w:rPr>'
      + '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
      + (o.bold ? '<w:b/><w:bCs/>' : '')
      + (o.italic ? '<w:i/><w:iCs/>' : '')
      + (o.ul ? '<w:u w:val="single"/>' : '')
      + '<w:sz w:val="' + (o.sz || 28) + '"/>'
      + '<w:szCs w:val="' + (o.sz || 28) + '"/>'
      + '</w:rPr><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
  }

  function para(a, o) {
    o = o || {};
    return '<w:p><w:pPr>'
      + '<w:jc w:val="' + (o.jc || 'both') + '"/>'
      + (o.ind ? '<w:ind w:firstLine="709"/>' : '')
      + '<w:spacing w:after="0" w:line="276" w:lineRule="auto"/>'
      + '</w:pPr>' + a.map(r => run(r.t, r)).join('') + '</w:p>';
  }

  function el() {
    return '<w:p><w:pPr><w:spacing w:after="0" w:line="180" w:lineRule="exact"/></w:pPr></w:p>';
  }

  const nb = '<w:tblBorders>'
    + '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '</w:tblBorders>';

  function cell(c, w) {
    return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>'
      + '<w:tcBorders>'
      + '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
      + '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
      + '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
      + '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
      + '</w:tcBorders></w:tcPr>' + c + '</w:tc>';
  }

  function cp(a) { return para(a, { jc: 'center' }); }

  const H = 4513;
  const sig =
    '<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/>' + nb + '</w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="' + H + '"/><w:gridCol w:w="' + H + '"/></w:tblGrid>'
    + '<w:tr>' + cell(cp([{ t: 'Người đổi' }]), H) + cell(cp([{ t: 'Người viết đơn' }]), H) + '</w:tr>'
    + '<w:tr>' + cell(el(), H) + cell(el(), H) + '</w:tr>'
    + '<w:tr>' + cell(el(), H) + cell(el(), H) + '</w:tr>'
    + '<w:tr>' + cell(cp([{ t: cl(d.nd) }]), H) + cell(cp([{ t: cl(d.nx) }]), H) + '</w:tr>'
    + '</w:tbl>'
    + '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>'
    + '<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/>' + nb + '</w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="9026"/></w:tblGrid>'
    + '<w:tr>' + cell(cp([{ t: 'Đội trưởng' }]), 9026) + '</w:tr>'
    + '<w:tr>' + cell(el(), 9026) + '</w:tr>'
    + '<w:tr>' + cell(el(), 9026) + '</w:tr>'
    + '<w:tr>' + cell(cp([{ t: 'Nguyễn Văn Trung', bold: true }]), 9026) + '</w:tr>'
    + '</w:tbl>';

  const cd = d.loai === 'truongca' ? 'trưởng ca' : 'nhân viên';
  const body =
    para([{ t: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', bold: true }], { jc: 'center' }) +
    para([{ t: 'Độc lập – Tự do – Hạnh phúc', bold: true, ul: true }], { jc: 'center' }) +
    para([{ t: 'TP. Hồ Chí Minh, ngày ' + d.ng + ' tháng ' + d.th + ' năm ' + d.na + '.', italic: true }], { jc: 'right' }) +
    el() +
    para([{ t: 'ĐƠN XIN ĐỔI CA LÀM VIỆC', bold: true, sz: 32 }], { jc: 'center' }) +
    el() +
    para([{ t: 'Kính gửi: ', bold: true }, { t: 'Đội trưởng Đội Vận hành, bảo trì đường hầm.' }], { jc: 'center' }) +
    el() +
    para([{ t: 'Tôi tên là: ' }, { t: d.nx }, { t: ', ' + cd + ' Đội Vận hành, bảo trì đường hầm.' }], { ind: true }) +
    para([{ t: 'Nay tôi viết đơn này xin phép cho tôi đổi ca với ông ' }, { t: d.nd }, { t: '.' }], { ind: true }) +
    para([{ t: d.cdi }, { t: ', ông ' }, { t: d.nd }, { t: ' sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.' }], { ind: true }) +
    para([{ t: d.cve }, { t: ', tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông ' }, { t: d.nd }, { t: '.' }], { ind: true }) +
    para([{ t: 'Số lần đã đổi ca trong tháng: ' }, { t: d.sl }, { t: '.' }], { ind: true }) +
    para([{ t: 'Lý do: ' }, { t: d.ld }], { ind: true }) +
    el() +
    para([{ t: 'Rất mong được sự chấp thuận của Đội trưởng.' }], { ind: true }) +
    para([{ t: 'Trân trọng cảm ơn.' }], { ind: true }) +
    el() + sig +
    '<w:sectPr>'
    + '<w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>'
    + '</w:sectPr>';

  const docX = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:body>' + body + '</w:body></w:document>';

  const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '</Types>';

  const rm = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships>';

  const rd = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

  const zip = new JSZip();
  zip.file('[Content_Types].xml', ct);
  zip.file('_rels/.rels', rm);
  zip.file('word/document.xml', docX);
  zip.file('word/_rels/document.xml.rels', rd);

  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  document.body.appendChild(a);
  a.style.display = 'none';
  a.href = url;
  a.download = fn + '.docx';
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 800);
}

// ===== INIT =====
load();
bind();
