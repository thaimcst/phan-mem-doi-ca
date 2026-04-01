// ===== CONFIG =====
const API = 'https://phan-d-default-rtdb.asia-southeast1.firebasedatabase.app/don_doi_ca';

// ===== STATE =====
window.rec = [];

// ===== CLOCK =====
function updateClock() {
  document.getElementById('clock').textContent = new Date().toLocaleTimeString('vi-VN');
  const db = document.getElementById('dateBadge');
  if (db) {
    const now = new Date();
    db.textContent = now.toLocaleDateString('vi-VN', { weekday:'long', day:'2-digit', month:'2-digit', year:'numeric' });
  }
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
window.fd = fd; window.tc = tc;

// ===== TOGGLE NGƯỜI ĐỔI 2 =====
document.getElementById('chk2').addEventListener('change', function () {
  const block = document.getElementById('block2');
  block.style.display = this.checked ? 'block' : 'none';

  // toggle required trên các field người đổi 2
  ['nd2', 'cdi2', 'ddi2', 'cve2', 'dve2'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.required = this.checked;
  });

  bind();
});

// ===== BIND PRINT PREVIEW =====
function bind() {
  const dots = '................................................................';
  const dotsCa = '.................................................';
  const has2 = document.getElementById('chk2').checked;

  const nx = document.getElementById('nx').value || dots;
  const nd = document.getElementById('nd').value || dots;
  const loai = document.querySelector('input[name="loai"]:checked').value;

  // Người viết & loại
  document.getElementById('p_nx').textContent = nx;
  document.getElementById('p_cd').textContent = loai === 'truongca' ? 'trưởng ca' : 'nhân viên';

  // Người đổi 1
  ['p_nd', 'p_nd2p', 'p_nd3p'].forEach(id => document.getElementById(id).textContent = nd);

  // "và ông X" khi có người đổi 2
  const andSpan = document.getElementById('p_nd_and2');
  if (has2) {
    const nd2val = document.getElementById('nd2').value;
    andSpan.textContent = nd2val ? ' và ông ' + nd2val : '';
  } else {
    andSpan.textContent = '';
  }

  const cdi = document.getElementById('cdi').value;
  const ddi = document.getElementById('ddi').value;
  const cve = document.getElementById('cve').value;
  const dve = document.getElementById('dve').value;
  document.getElementById('p_c1').textContent = (cdi && ddi) ? cdi + ' ngày ' + fd(ddi) : dotsCa;
  document.getElementById('p_c2').textContent = (cve && dve) ? cve + ' ngày ' + fd(dve) : dotsCa;

  // Người đổi 2
  const block2Print = document.getElementById('p_block2_print');
  const sigWrap1 = document.getElementById('sig_wrap_1');
  const sigWrap2 = document.getElementById('sig_wrap_2');

  if (has2) {
    const nd2 = document.getElementById('nd2').value || dots;
    const cdi2 = document.getElementById('cdi2').value;
    const ddi2 = document.getElementById('ddi2').value;
    const cve2 = document.getElementById('cve2').value;
    const dve2 = document.getElementById('dve2').value;

    ['p_nd2_b', 'p_nd3_b'].forEach(id => document.getElementById(id).textContent = nd2);
    document.getElementById('p_c1_b').textContent = (cdi2 && ddi2) ? cdi2 + ' ngày ' + fd(ddi2) : dotsCa;
    document.getElementById('p_c2_b').textContent = (cve2 && dve2) ? cve2 + ' ngày ' + fd(dve2) : dotsCa;

    // Chữ ký 2 người
    document.getElementById('p_ndk_b1').textContent = nd ? tc(nd) : '............................';
    document.getElementById('p_ndk_b2').textContent = nd2 !== dots ? tc(nd2) : '............................';
    document.getElementById('p_nxk2').textContent = nx !== dots ? tc(nx) : '............................';

    block2Print.style.display = 'inline';
    sigWrap1.classList.remove('active'); sigWrap1.style.display = 'none';
    sigWrap2.classList.add('active');    sigWrap2.style.display = 'block';
  } else {
    block2Print.style.display = 'none';
    sigWrap1.classList.add('active');    sigWrap1.style.display = 'block';
    sigWrap2.classList.remove('active'); sigWrap2.style.display = 'none';

    // Chữ ký 1 người
    document.getElementById('p_nxk').textContent = nx !== dots ? tc(nx) : '............................';
    document.getElementById('p_ndk').textContent = nd !== dots ? tc(nd) : '............................';
  }

  document.getElementById('p_sl').textContent = document.getElementById('sl').value || '......';
  document.getElementById('p_ld').textContent = document.getElementById('ld').value || '...............................................................................';

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

  const has2 = document.getElementById('chk2').checked;
  const now = new Date();

  // Lấy dữ liệu từ form
  const nx  = document.getElementById('nx').value  || '............................';
  const nd  = document.getElementById('nd').value  || '............................';
  const nd2 = has2 ? (document.getElementById('nd2').value || '............................') : '';
  const cdi = document.getElementById('cdi').value;
  const ddi = document.getElementById('ddi').value;
  const cve = document.getElementById('cve').value;
  const dve = document.getElementById('dve').value;
  const cdi2 = has2 ? document.getElementById('cdi2').value : '';
  const ddi2 = has2 ? document.getElementById('ddi2').value : '';
  const cve2 = has2 ? document.getElementById('cve2').value : '';
  const dve2 = has2 ? document.getElementById('dve2').value : '';
  const sl  = document.getElementById('sl').value  || '......';
  const ld  = document.getElementById('ld').value  || '...............................................................................';
  const loai = document.querySelector('input[name="loai"]:checked').value;
  const cd  = loai === 'truongca' ? 'trưởng ca' : 'nhân viên';

  const dotsCa = '.................................................';
  const c1 = (cdi && ddi) ? cdi + ' ngày ' + fd(ddi) : dotsCa;
  const c2 = (cve && dve) ? cve + ' ngày ' + fd(dve) : dotsCa;
  const c1b = has2 && cdi2 && ddi2 ? cdi2 + ' ngày ' + fd(ddi2) : dotsCa;
  const c2b = has2 && cve2 && dve2 ? cve2 + ' ngày ' + fd(dve2) : dotsCa;

  const ng = now.getDate().toString().padStart(2,'0');
  const th = (now.getMonth()+1).toString().padStart(2,'0');
  const na = now.getFullYear();

  // Phần người đổi 2 (nếu có)
  const para2 = has2 ? `
    <p>${c1b}, ông ${nd2} sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.</p>
    <p>${c2b}, tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông ${nd2}.</p>
  ` : '';

  const ndLine = has2 ? `${nd} và ông ${nd2}` : nd;

  // Phần chữ ký
  const spaceRows = `
      <tr class="sig-space"><td></td><td></td></tr>
      <tr class="sig-space"><td></td><td></td></tr>
      <tr class="sig-space"><td></td><td></td></tr>
      <tr class="sig-space"><td></td><td></td></tr>
      <tr class="sig-space"><td></td><td></td></tr>`;
  const spaceRows3 = `
      <tr class="sig-space"><td></td><td></td><td></td></tr>
      <tr class="sig-space"><td></td><td></td><td></td></tr>
      <tr class="sig-space"><td></td><td></td><td></td></tr>
      <tr class="sig-space"><td></td><td></td><td></td></tr>
      <tr class="sig-space"><td></td><td></td><td></td></tr>`;
  const dtBlock = `
    <table class="sig-tbl" style="margin-top:16px">
      <tr><td style="text-align:center">Đội trưởng</td></tr>
      <tr class="sig-space"><td></td></tr>
      <tr class="sig-space"><td></td></tr>
      <tr class="sig-space"><td></td></tr>
      <tr class="sig-space"><td></td></tr>
      <tr class="sig-space"><td></td></tr>
      <tr><td style="text-align:center;font-weight:bold">Nguyễn Văn Trung</td></tr>
    </table>`;

  const sigHTML = has2 ? `
    <table class="sig-tbl">
      <tr><td>Người đổi (1)</td><td>Người đổi (2)</td><td>Người viết đơn</td></tr>
      ${spaceRows3}
      <tr><td>${tc(nd)}</td><td>${tc(nd2)}</td><td>${tc(nx)}</td></tr>
    </table>
    ${dtBlock}
  ` : `
    <table class="sig-tbl">
      <tr><td>Người đổi</td><td>Người viết đơn</td></tr>
      ${spaceRows}
      <tr><td>${tc(nd)}</td><td>${tc(nx)}</td></tr>
    </table>
    ${dtBlock}
  `;

  const html = `<!DOCTYPE html><html lang="vi"><head>
    <meta charset="UTF-8">
    <title>Đơn Xin Đổi Ca</title>
    <style>
      @page { size: A4; margin: 20mm 18mm 20mm 25mm; }
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { font-family: 'Times New Roman', Times, serif; font-size: 13pt; line-height: 1.6; color: #000; background: #fff; }
      .ph { text-align: center; margin-bottom: 2px; }
      .ph .ul { font-weight: bold; text-decoration: underline; }
      .pd { text-align: right; margin: 4px 0 14px; font-style: italic; font-size: 12pt; }
      .pt { text-align: center; font-size: 15pt; font-weight: bold; margin: 10px 0; }
      .pkg { text-align: center; margin-bottom: 14px; font-size: 12pt; }
      .pc p { text-indent: 10mm; margin-bottom: 3px; text-align: justify; font-size: 12pt; }
      .sig-tbl { width: 100%; border-collapse: collapse; margin-top: 20px; }
      .sig-tbl td { text-align: center; font-size: 12pt; padding: 0; border: none; }
      .sig-space td { height: 14px; }
    </style>
  </head><body>
    <div class="ph"><b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br><span class="ul">Độc lập – Tự do – Hạnh phúc</span></div>
    <div class="pd"><i>TP. Hồ Chí Minh, ngày ${ng} tháng ${th} năm ${na}.</i></div>
    <div class="pt">ĐƠN XIN ĐỔI CA LÀM VIỆC</div>
    <div class="pkg"><b>Kính gửi:</b> Đội trưởng Đội Vận hành, bảo trì đường hầm.</div>
    <div class="pc">
      <p>Tôi tên là: ${nx}, ${cd} Đội Vận hành, bảo trì đường hầm.</p>
      <p>Nay tôi viết đơn này xin phép cho tôi đổi ca với ông ${ndLine}.</p>
      <p>${c1}, ông ${nd} sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.</p>
      <p>${c2}, tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông ${nd}.</p>
      ${para2}
      <p>Số lần đã đổi ca trong tháng: ${sl}.</p>
      <p>Lý do: ${ld}.</p>
      <p>Rất mong được sự chấp thuận của Đội trưởng.</p>
      <p>Trân trọng cảm ơn.</p>
    </div>
    ${sigHTML}
    <script>window.onload = function(){ window.print(); window.onafterprint = function(){ window.close(); }; }<\/script>
  </body></html>`;

  const w = window.open('', '_blank', 'width=800,height=600');
  w.document.write(html);
  w.document.close();
};

// ===== EXPORT WORD =====
document.getElementById('btnWord').onclick = () => {
  bind();
  const has2 = document.getElementById('chk2').checked;
  const now = new Date();
  const base = {
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
  };
  if (has2) {
    base.nd2 = document.getElementById('nd2').value || '...';
    base.cdi2 = document.getElementById('p_c1_b').textContent;
    base.cve2 = document.getElementById('p_c2_b').textContent;
  }
  expDocx(base, 'Don_Doi_Ca_' + tc(base.nx || 'Moi').replace(/ /g, '_'));
};

// ===== SAVE FORM =====
document.getElementById('frm').addEventListener('submit', async function (e) {
  e.preventDefault();
  const has2 = document.getElementById('chk2').checked;
  const o = {
    nguoiXin: document.getElementById('nx').value,
    nguoiDoi: document.getElementById('nd').value,
    caDi: document.getElementById('cdi').value + ' ngày ' + fd(document.getElementById('ddi').value),
    caVe: document.getElementById('cve').value + ' ngày ' + fd(document.getElementById('dve').value),
    soLan: document.getElementById('sl').value,
    lyDo: document.getElementById('ld').value,
    loai: document.querySelector('input[name="loai"]:checked').value,
    createdAt: new Date().toISOString()
  };
  if (has2) {
    o.nguoiDoi2 = document.getElementById('nd2').value;
    o.caDi2 = document.getElementById('cdi2').value + ' ngày ' + fd(document.getElementById('ddi2').value);
    o.caVe2 = document.getElementById('cve2').value + ' ngày ' + fd(document.getElementById('dve2').value);
  }
  try {
    const r = await fetch(API + '.json', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(o)
    });
    if (!r.ok) throw 0;
    this.reset();
    document.getElementById('chk2').checked = false;
    document.getElementById('block2').style.display = 'none';
    bind(); load();
    alert('✅ Đã lưu thành công!');
  } catch { alert('❌ Lưu thất bại!'); }
});

// ===== LOAD DATA =====
async function load() {
  try {
    const r = await fetch(API + '.json'), d = await r.json();
    let a = [];
    if (d && typeof d === 'object') { for (let k in d) a.push({ id: k, ...d[k] }); }
    a.sort((x, y) => new Date(y.createdAt || 0) - new Date(x.createdAt || 0));
    window.rec = a;

    // Update stats
    const now2 = new Date();
    const thisMonth = a.filter(i => { const d = new Date(i.createdAt); return d.getMonth() === now2.getMonth() && d.getFullYear() === now2.getFullYear(); });
    const today = a.filter(i => { const d = new Date(i.createdAt); return d.toDateString() === now2.toDateString(); });
    document.getElementById('stat-total').textContent = a.length;
    document.getElementById('stat-month').textContent = thisMonth.length;
    document.getElementById('stat-today').textContent = today.length;
    const tb = document.getElementById('tbody');
    if (!a.length) { tb.innerHTML = '<tr><td colspan="5" class="empty">Chưa có đơn nào</td></tr>'; return; }
    tb.innerHTML = '';
    a.forEach(i => {
      const lb = i.loai === 'truongca'
        ? '<span class="badge badge-tc">⭐ Trưởng ca</span>'
        : '<span class="badge badge-nv">👤 Nhân viên</span>';
      const nguoiDoi = i.nguoiDoi2
        ? `<strong>${i.nguoiDoi}</strong><br><small style="color:var(--green)">+ ${i.nguoiDoi2}</small>`
        : `<strong>${i.nguoiDoi}</strong>`;
      const caDetail = i.nguoiDoi2
        ? `<div class="ca-box"><span style="color:var(--red);font-size:.75rem">▲</span> ${i.caDi}</div>
           <div class="ca-box"><span style="color:var(--green);font-size:.75rem">▼</span> ${i.caVe}</div>
           <div class="ca-box" style="opacity:.7"><span style="color:var(--red);font-size:.75rem">▲</span> ${i.caDi2}</div>
           <div class="ca-box" style="opacity:.7"><span style="color:var(--green);font-size:.75rem">▼</span> ${i.caVe2}</div>`
        : `<div class="ca-box"><span style="color:var(--red);font-size:.75rem">▲</span> ${i.caDi}</div>
           <div class="ca-box"><span style="color:var(--green);font-size:.75rem">▼</span> ${i.caVe}</div>`;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td data-label="Ngày tạo">${new Date(i.createdAt).toLocaleDateString('vi-VN')}<br><small style="color:var(--muted)">${new Date(i.createdAt).toLocaleTimeString('vi-VN')}</small></td>
        <td data-label="Người xin đổi"><strong>${i.nguoiXin}</strong><br>${lb}</td>
        <td data-label="Trực thay">${nguoiDoi}</td>
        <td data-label="Chi tiết ca">${caDetail}</td>
        <td data-label="Thao tác" class="td-act">
          <button class="bw" onclick="expOld('${i.id}')">📄 Word</button>
          <button class="bd" onclick="del('${i.id}')">🗑</button>
        </td>`;
      tb.appendChild(tr);
    });
  } catch (e) {
    document.getElementById('tbody').innerHTML = '<tr><td colspan="5" class="empty" style="color:var(--red)">Lỗi kết nối!</td></tr>';
  }
}

// ===== EXPORT OLD =====
window.expOld = function (id) {
  const r = window.rec.find(x => x.id === id); if (!r) return;
  const d = new Date(r.createdAt);
  const data = {
    nx: r.nguoiXin, nd: r.nguoiDoi,
    cdi: r.caDi, cve: r.caVe,
    sl: r.soLan, ld: r.lyDo,
    loai: r.loai || 'nhanvien',
    ng: d.getDate().toString().padStart(2, '0'),
    th: (d.getMonth() + 1).toString().padStart(2, '0'),
    na: d.getFullYear().toString()
  };
  if (r.nguoiDoi2) {
    data.nd2 = r.nguoiDoi2;
    data.cdi2 = r.caDi2;
    data.cve2 = r.caVe2;
  }
  expDocx(data, 'Don_Doi_Ca_' + tc(r.nguoiXin || 'Moi').replace(/ /g, '_'));
};

// ===== DELETE =====
window.del = async function (id) {
  if (!confirm('Xóa đơn này?')) return;
  try { await fetch(API + '/' + id + '.json', { method: 'DELETE' }); load(); }
  catch { alert('Lỗi xóa!'); }
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
      + (o.bold ? '<w:b/><w:bCs/>' : '') + (o.italic ? '<w:i/><w:iCs/>' : '') + (o.ul ? '<w:u w:val="single"/>' : '')
      + '<w:sz w:val="' + (o.sz || 28) + '"/><w:szCs w:val="' + (o.sz || 28) + '"/>'
      + '</w:rPr><w:t xml:space="preserve">' + esc(t) + '</w:t></w:r>';
  }
  function para(a, o) {
    o = o || {};
    return '<w:p><w:pPr><w:jc w:val="' + (o.jc || 'both') + '"/>'
      + (o.ind ? '<w:ind w:firstLine="709"/>' : '')
      + '<w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr>'
      + a.map(r => run(r.t, r)).join('') + '</w:p>';
  }
  function el() { return '<w:p><w:pPr><w:spacing w:after="0" w:line="180" w:lineRule="exact"/></w:pPr></w:p>'; }

  const nb = '<w:tblBorders>'
    + '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    + '</w:tblBorders>';
  function cell(c, w) {
    return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>'
      + '<w:tcBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
      + '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
      + '</w:tcBorders></w:tcPr>' + c + '</w:tc>';
  }
  function cp(a) { return para(a, { jc: 'center' }); }

  const cd = d.loai === 'truongca' ? 'trưởng ca' : 'nhân viên';
  const has2 = !!d.nd2;

  // Chữ ký — 1 bảng 3 cột đều nhau, khớp đúng vùng nội dung A4
  // Content width = 11906 (A4) - 1701 (lề trái) - 851 (lề phải) = 9354 dxa
  const TW = 9354; const W = 3118; const WM = 3118; // 3118 * 3 = 9354
  let sigCol1Label, sigCol1Name, sigCol2Label, sigCol2Name, sigCol3Label, sigCol3Name;
  if (!has2) {
    sigCol1Label = 'Người đổi';      sigCol1Name = cl(d.nd);
    sigCol2Label = 'Đội trưởng';     sigCol2Name = 'Nguyễn Văn Trung';
    sigCol3Label = 'Người viết đơn'; sigCol3Name = cl(d.nx);
  } else {
    sigCol1Label = 'Người đổi (1)';  sigCol1Name = cl(d.nd);
    sigCol2Label = 'Người đổi (2)';  sigCol2Name = cl(d.nd2);
    sigCol3Label = 'Người viết đơn'; sigCol3Name = cl(d.nx);
  }

  const sig = '<w:tbl><w:tblPr><w:tblW w:w="' + TW + '" w:type="dxa"/>' + nb + '</w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="' + W + '"/><w:gridCol w:w="' + WM + '"/><w:gridCol w:w="' + W + '"/></w:tblGrid>'
    + '<w:tr>' + cell(cp([{ t: sigCol1Label }]), W) + cell(cp([{ t: sigCol2Label }]), WM) + cell(cp([{ t: sigCol3Label }]), W) + '</w:tr>'
    + '<w:tr>' + cell(el(), W) + cell(el(), WM) + cell(el(), W) + '</w:tr>'
    + '<w:tr>' + cell(el(), W) + cell(el(), WM) + cell(el(), W) + '</w:tr>'
    + '<w:tr>' + cell(el(), W) + cell(el(), WM) + cell(el(), W) + '</w:tr>'
    + '<w:tr>' + cell(el(), W) + cell(el(), WM) + cell(el(), W) + '</w:tr>'
    + '<w:tr>' + cell(el(), W) + cell(el(), WM) + cell(el(), W) + '</w:tr>'
    + '<w:tr>'
      + cell(cp([{ t: sigCol1Name }]), W)
      + cell(cp([{ t: sigCol2Name, bold: !has2 }]), WM)
      + cell(cp([{ t: sigCol3Name }]), W)
    + '</w:tr>'
    + '</w:tbl>';

  // sigDT không còn cần thiết nữa (đã gộp vào sig)
  const sigDT = '';

  let bodyParas =
    para([{ t: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', bold: true }], { jc: 'center' }) +
    para([{ t: 'Độc lập – Tự do – Hạnh phúc', bold: true, ul: true }], { jc: 'center' }) +
    para([{ t: 'TP. Hồ Chí Minh, ngày ' + d.ng + ' tháng ' + d.th + ' năm ' + d.na + '.', italic: true }], { jc: 'right' }) +
    el() +
    para([{ t: 'ĐƠN XIN ĐỔI CA LÀM VIỆC', bold: true, sz: 32 }], { jc: 'center' }) +
    el() +
    para([{ t: 'Kính gửi: ', bold: true }, { t: 'Đội trưởng Đội Vận hành, bảo trì đường hầm.' }], { jc: 'center' }) +
    el() +
    para([{ t: 'Tôi tên là: ' }, { t: d.nx }, { t: ', ' + cd + ' Đội Vận hành, bảo trì đường hầm.' }], { ind: true }) +
    (has2
      ? para([{ t: 'Nay tôi viết đơn này xin phép cho tôi đổi ca với ông ' }, { t: d.nd }, { t: ' và ông ' }, { t: d.nd2 }, { t: '.' }], { ind: true })
      : para([{ t: 'Nay tôi viết đơn này xin phép cho tôi đổi ca với ông ' }, { t: d.nd }, { t: '.' }], { ind: true })
    ) +
    para([{ t: d.cdi }, { t: ', ông ' }, { t: d.nd }, { t: ' sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.' }], { ind: true }) +
    para([{ t: d.cve }, { t: ', tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông ' }, { t: d.nd }, { t: '.' }], { ind: true });

  if (has2) {
    bodyParas +=
      para([{ t: d.cdi2 }, { t: ', ông ' }, { t: d.nd2 }, { t: ' sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.' }], { ind: true }) +
      para([{ t: d.cve2 }, { t: ', tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông ' }, { t: d.nd2 }, { t: '.' }], { ind: true });
  }

  bodyParas +=
    para([{ t: 'Số lần đã đổi ca trong tháng: ' }, { t: d.sl }, { t: '.' }], { ind: true }) +
    para([{ t: 'Lý do: ' }, { t: d.ld }], { ind: true }) +
    el() +
    para([{ t: 'Rất mong được sự chấp thuận của Đội trưởng.' }], { ind: true }) +
    para([{ t: 'Trân trọng cảm ơn.' }], { ind: true }) +
    el() + sig + sigDT +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>'
    + '</w:sectPr>';

  const docX = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    + '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:body>' + bodyParas + '</w:body></w:document>';

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

  const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  document.body.appendChild(a);
  a.style.display = 'none'; a.href = url;
  a.download = fn + '.docx'; a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 800);
}

// ===== INIT =====
load();
bind();
