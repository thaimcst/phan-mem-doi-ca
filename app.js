// ===== CONFIG =====
const API = 'https://phan-d-default-rtdb.asia-southeast1.firebasedatabase.app/don_doi_ca';
const GIOI_HAN_THANG = 4;

// ===== LỊCH KÍP VẬN HÀNH =====
// Load từ file lich_kip.json
let LICH_KIP = null;
fetch('lich_kip.json')
  .then(r => r.json())
  .then(data => {
    LICH_KIP = data;
    console.log('✅ Đã tải lịch kíp:', Object.keys(data).length, 'nhân viên');
    // Cập nhật gợi ý nếu form đã có giá trị
    capNhatGoiY();
  })
  .catch(e => console.warn('⚠️ Không tải được lịch kíp:', e));

// Lấy ca làm việc của một người trong ngày
function getCaNgay(tenNguoi, dateStr) {
  if (!LICH_KIP || !LICH_KIP[tenNguoi]) return null;
  return LICH_KIP[tenNguoi].lich[dateStr] || null;
}

// Lấy kíp của một người
function getKip(tenNguoi) {
  if (!LICH_KIP || !LICH_KIP[tenNguoi]) return null;
  return LICH_KIP[tenNguoi].kip;
}

// Kiểm tra đổi ca có hợp lệ không
// Điều kiện: 2 người phải trực 2 ca KHÁC nhau trong 2 ngày được chọn
function kiemTraHopLe(tenA, caChuyenDi_date, tenB, caNhanVe_date) {
  if (!LICH_KIP) return { ok: true, msg: '' };

  const caA_homdo = getCaNgay(tenA, caChuyenDi_date);  // ca A đang trực ngày cần đổi
  const caB_ngaydo = getCaNgay(tenB, caChuyenDi_date); // ca B ngày đó
  const caA_ngaynhan = getCaNgay(tenA, caNhanVe_date); // ca A ngày nhận
  const caB_ngaynhan = getCaNgay(tenB, caNhanVe_date); // ca B ngày nhận

  const msgs = [];

  // Nếu không có trong lịch (ngoài tháng 5-12/2026)
  if (!caA_homdo && !caB_ngaynhan) return { ok: true, msg: 'Ngày ngoài phạm vi lịch' };

  // Kiểm tra A có đang nghỉ không
  if (caA_homdo === 'N') msgs.push(`⚠️ ${tenA} đang nghỉ ngày ${caChuyenDi_date}`);
  // Kiểm tra B có đang nghỉ không
  if (caB_ngaynhan === 'N') msgs.push(`⚠️ ${tenB} đang nghỉ ngày ${caNhanVe_date}`);
  // Kiểm tra B ngày A chuyển đi có trực không (phải khác ca A)
  if (caA_homdo && caB_ngaydo && caA_homdo !== 'N' && caB_ngaydo !== 'N' && caA_homdo === caB_ngaydo) {
    msgs.push(`⚠️ ${tenA} và ${tenB} cùng trực ${caA_homdo} ngày ${caChuyenDi_date} — không thể đổi`);
  }
  // Kiểm tra A ngày nhận về có trực không
  if (caA_ngaynhan && caA_ngaynhan !== 'N' && caB_ngaynhan && caB_ngaynhan !== 'N' && caA_ngaynhan === caB_ngaynhan) {
    msgs.push(`⚠️ ${tenA} và ${tenB} cùng trực ${caA_ngaynhan} ngày ${caNhanVe_date} — không thể đổi`);
  }

  return { ok: msgs.length === 0, msg: msgs.join('\n'), caA: caA_homdo, caB_nhan: caB_ngaynhan };
}

// Tìm người CÓ THỂ đổi ca cho tenNguoi vào ngày đó
function timNguoiCoTheDoiCa(tenNguoi, ngayChuyenDi) {
  if (!LICH_KIP || !ngayChuyenDi) return [];

  const caCanDoi = getCaNgay(tenNguoi, ngayChuyenDi);
  if (!caCanDoi || caCanDoi === 'N') return [];

  const kipNguoi = getKip(tenNguoi);
  const goiY = [];

  for (const [ten, info] of Object.entries(LICH_KIP)) {
    if (ten === tenNguoi) continue;
    // Chỉ gợi ý người kíp khác (khác kíp mới có thể đổi)
    if (info.kip === kipNguoi) continue;

    const caNgay = info.lich[ngayChuyenDi];
    // Người đó phải đang trực ca khác vào ngày cần đổi
    if (!caNgay || caNgay === 'N' || caNgay === caCanDoi) continue;

    // Tìm ngày B có thể trực thay (ngày B đang trực caCanDoi, A nghỉ)
    const ngayBuLai = timNgayBuLai(ten, caCanDoi, tenNguoi, ngayChuyenDi);

    goiY.push({
      ten,
      kip: info.kip,
      caHienTai: caNgay,
      caCanDoi,
      ngayBuLai: ngayBuLai?.ngay || null,
      caBuLai: ngayBuLai?.ca || null,
      hopLe: !!ngayBuLai
    });
  }

  return goiY.sort((a, b) => (b.hopLe ? 1 : 0) - (a.hopLe ? 1 : 0));
}

// Tìm ngày B có thể trực thay cho A (ngày B đang trực caCanDoi, còn A nghỉ)
function timNgayBuLai(tenB, caCanTimB, tenA, ngayGoc) {
  if (!LICH_KIP) return null;
  const lichA = LICH_KIP[tenA]?.lich || {};
  const lichB = LICH_KIP[tenB]?.lich || {};

  // Tìm trong vòng 30 ngày từ ngày gốc
  const base = new Date(ngayGoc);
  for (let delta = 1; delta <= 30; delta++) {
    for (const sign of [1, -1]) {
      const d = new Date(base);
      d.setDate(d.getDate() + delta * sign);
      const ds = d.toISOString().split('T')[0];
      // B phải trực caCanTimB ngày đó, và A phải nghỉ hoặc trực ca khác
      if (lichB[ds] === caCanTimB && lichA[ds] !== caCanTimB) {
        return { ngay: ds, ca: caCanTimB };
      }
    }
  }
  return null;
}

// Gợi ý khi chọn người và ngày
function capNhatGoiY() {
  const nx = document.getElementById('nx')?.value;
  const ddi = document.getElementById('ddi')?.value;
  const cdi = document.getElementById('cdi')?.value;
  const el = document.getElementById('goi-y-doi-ca');
  if (!el) return;
  if (!nx || !ddi || !LICH_KIP) { el.style.display = 'none'; return; }

  const goiY = timNguoiCoTheDoiCa(nx, ddi);
  if (!goiY.length) { el.style.display = 'none'; return; }

  const caCanDoi = getCaNgay(nx, ddi);
  const kipNx = getKip(nx);

  el.style.display = 'block';
  const hopLe = goiY.filter(g => g.hopLe).slice(0, 6);
  const khac = goiY.filter(g => !g.hopLe).slice(0, 3);

  const caMap = { C1: 'Ca 1', C2: 'Ca 2', C3: 'Ca 3', N: 'Nghỉ' };
  const kipColor = { 'Kíp 1': '#00b4ff', 'Kíp 2': '#00d68f', 'Kíp 3': '#f0a500', 'Kíp 4': '#c084fc' };

  let html = `<div style="font-size:.68rem;color:var(--muted);margin-bottom:8px;text-transform:uppercase;letter-spacing:1px">
    💡 Gợi ý đổi ca — ${nx} <span style="color:var(--accent)">(${kipNx})</span> đang trực <strong style="color:var(--red)">${caMap[caCanDoi]||caCanDoi}</strong> ngày ${fd(ddi)}
  </div>`;

  if (hopLe.length) {
    html += `<div style="font-size:.7rem;color:var(--green);margin-bottom:6px">✅ Có thể đổi:</div>`;
    html += hopLe.map(g => {
      const kc = kipColor[g.kip] || '#fff';
      return `<div class="goi-y-item" onclick="chonGoiY('${g.ten}','${g.ngayBuLai||''}','${g.caBuLai||''}','${caCanDoi}','${ddi}')"
        style="background:rgba(0,214,143,.06);border:1px solid rgba(0,214,143,.25);border-radius:6px;padding:8px 10px;margin-bottom:5px;cursor:pointer;transition:all .2s"
        onmouseover="this.style.background='rgba(0,214,143,.12)'" onmouseout="this.style.background='rgba(0,214,143,.06)'">
        <div style="display:flex;align-items:center;gap:8px">
          <span style="width:6px;height:6px;border-radius:50%;background:${kc};flex-shrink:0"></span>
          <strong style="font-size:.85rem;color:var(--text)">${g.ten}</strong>
          <span style="font-size:.65rem;color:${kc};border:1px solid ${kc};padding:1px 6px;border-radius:10px">${g.kip}</span>
          <span style="margin-left:auto;font-size:.72rem;color:var(--green)">← Bấm để chọn</span>
        </div>
        <div style="font-size:.72rem;color:var(--muted);margin-top:4px;padding-left:14px">
          Ngày ${fd(ddi)}: đang trực <strong>${caMap[g.caHienTai]||g.caHienTai}</strong>
          ${g.ngayBuLai ? `→ Bù lại: <strong>${caMap[g.caBuLai]||g.caBuLai}</strong> ngày <strong>${fd(g.ngayBuLai)}</strong>` : ''}
        </div>
      </div>`;
    }).join('');
  }

  if (khac.length) {
    html += `<div style="font-size:.7rem;color:var(--muted);margin:6px 0 4px">⚠️ Có thể xem xét:</div>`;
    html += khac.map(g => {
      const kc = kipColor[g.kip] || '#888';
      return `<div style="background:rgba(255,255,255,.03);border:1px solid var(--border);border-radius:6px;padding:7px 10px;margin-bottom:4px">
        <div style="display:flex;align-items:center;gap:7px">
          <span style="width:6px;height:6px;border-radius:50%;background:${kc};flex-shrink:0"></span>
          <strong style="font-size:.82rem;color:var(--muted)">${g.ten}</strong>
          <span style="font-size:.65rem;color:${kc};border:1px solid ${kc};padding:1px 6px;border-radius:10px">${g.kip}</span>
        </div>
        <div style="font-size:.7rem;color:var(--muted);margin-top:3px;padding-left:13px">Ngày ${fd(ddi)}: trực <strong>${caMap[g.caHienTai]||g.caHienTai}</strong></div>
      </div>`;
    }).join('');
  }

  el.innerHTML = html;
}

// Khi bấm chọn gợi ý → tự điền form
window.chonGoiY = function(ten, ngayBuLai, caBuLai, caCanDoi, ngayChuyenDi) {
  // Điền người đổi
  const ndEl = document.getElementById('nd');
  if (ndEl) {
    ndEl.value = ten;
    // Nếu option không tồn tại, thêm vào
    if (![...ndEl.options].find(o => o.value === ten)) {
      const opt = new Option(ten, ten);
      ndEl.appendChild(opt);
    }
    ndEl.value = ten;
  }
  // Điền ca chuyển đi
  const caMap2 = { C1: 'Ca 1', C2: 'Ca 2', C3: 'Ca 3' };
  const cdiEl = document.getElementById('cdi');
  if (cdiEl && caCanDoi) cdiEl.value = caMap2[caCanDoi] || caCanDoi;
  // Điền ngày chuyển đi (đã có)
  // Điền ca nhận về & ngày
  if (ngayBuLai) {
    const cveEl = document.getElementById('cve');
    const dveEl = document.getElementById('dve');
    if (cveEl && caBuLai) cveEl.value = caMap2[caBuLai] || caBuLai;
    if (dveEl) dveEl.value = ngayBuLai;
  }
  bind();
  // Flash xác nhận
  const el = document.getElementById('goi-y-doi-ca');
  if (el) {
    el.style.borderColor = 'rgba(0,214,143,.6)';
    setTimeout(() => { el.style.borderColor = 'rgba(0,180,255,.15)'; }, 800);
  }
};



// ===== STATE =====
window.rec = [];

// ===== CLOCK =====
function updateClock() {
  document.getElementById('clock').textContent = new Date().toLocaleTimeString('vi-VN');
  const db = document.getElementById('dateBadge');
  if (db) db.textContent = new Date().toLocaleDateString('vi-VN', { weekday:'long', day:'2-digit', month:'2-digit', year:'numeric' });
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

// ===== GIỚI HẠN ĐỔI CA =====
function demSoLanThang(records, tenNguoi) {
  const now = new Date();
  return records.filter(r => {
    if (r.nguoiXin !== tenNguoi) return false;
    const d = new Date(r.createdAt);
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  }).length;
}
function kiemTraGioiHan(tenNguoi) {
  const soLan = demSoLanThang(window.rec, tenNguoi);
  return { soLan, vuotGioiHan: soLan >= GIOI_HAN_THANG, conLai: Math.max(0, GIOI_HAN_THANG - soLan) };
}

// ===== CẢNH BÁO GIỚI HẠN =====
function capNhatCanhBao() {
  const nx = document.getElementById('nx').value;
  let el = document.getElementById('gioi-han-warning');
  if (!el) {
    el = document.createElement('div');
    el.id = 'gioi-han-warning';
    el.style.cssText = 'margin:6px 0 0;padding:8px 12px;border-radius:4px;font-size:0.78rem;font-weight:600;display:none;';
    document.getElementById('nx').parentNode.appendChild(el);
  }
  if (!nx) { el.style.display = 'none'; return; }
  const { soLan, vuotGioiHan, conLai } = kiemTraGioiHan(nx);
  if (vuotGioiHan) {
    el.style.display = 'block';
    el.style.background = 'rgba(225,29,72,0.12)';
    el.style.border = '1px solid rgba(225,29,72,0.3)';
    el.style.color = '#fda4af';
    el.innerHTML = `🚫 <strong>${nx}</strong> đã đổi ca <strong>${soLan}/${GIOI_HAN_THANG} lần</strong> trong tháng này. Không thể đổi thêm!`;
  } else if (soLan >= GIOI_HAN_THANG - 1) {
    el.style.display = 'block';
    el.style.background = 'rgba(245,158,11,0.12)';
    el.style.border = '1px solid rgba(245,158,11,0.3)';
    el.style.color = '#fcd34d';
    el.innerHTML = `⚠️ <strong>${nx}</strong> đã đổi <strong>${soLan}/${GIOI_HAN_THANG} lần</strong> — còn <strong>${conLai} lần</strong> trong tháng.`;
  } else if (soLan > 0) {
    el.style.display = 'block';
    el.style.background = 'rgba(0,180,255,0.07)';
    el.style.border = '1px solid rgba(0,180,255,0.2)';
    el.style.color = '#7ec8ff';
    el.innerHTML = `ℹ️ <strong>${nx}</strong> đã đổi <strong>${soLan}/${GIOI_HAN_THANG} lần</strong> — còn <strong>${conLai} lần</strong>.`;
  } else {
    el.style.display = 'none';
  }
}
document.getElementById('nx').addEventListener('change', capNhatCanhBao);

// ===== TOGGLE NGƯỜI ĐỔI 2 =====
document.getElementById('chk2').addEventListener('change', function () {
  document.getElementById('block2').style.display = this.checked ? 'block' : 'none';
  ['nd2','cdi2','ddi2','cve2','dve2'].forEach(id => {
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

  document.getElementById('p_nx').textContent = nx;
  document.getElementById('p_cd').textContent = loai === 'truongca' ? 'trưởng ca' : 'nhân viên';
  ['p_nd','p_nd2p','p_nd3p'].forEach(id => document.getElementById(id).textContent = nd);

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

  const block2Print = document.getElementById('p_block2_print');
  const sigWrap1 = document.getElementById('sig_wrap_1');
  const sigWrap2 = document.getElementById('sig_wrap_2');

  if (has2) {
    const nd2 = document.getElementById('nd2').value || dots;
    const cdi2 = document.getElementById('cdi2').value;
    const ddi2 = document.getElementById('ddi2').value;
    const cve2 = document.getElementById('cve2').value;
    const dve2 = document.getElementById('dve2').value;
    ['p_nd2_b','p_nd3_b'].forEach(id => document.getElementById(id).textContent = nd2);
    document.getElementById('p_c1_b').textContent = (cdi2 && ddi2) ? cdi2 + ' ngày ' + fd(ddi2) : dotsCa;
    document.getElementById('p_c2_b').textContent = (cve2 && dve2) ? cve2 + ' ngày ' + fd(dve2) : dotsCa;
    document.getElementById('p_ndk_b1').textContent = nd !== dots ? tc(nd) : '............................';
    document.getElementById('p_ndk_b2').textContent = nd2 !== dots ? tc(nd2) : '............................';
    document.getElementById('p_nxk2').textContent = nx !== dots ? tc(nx) : '............................';
    block2Print.style.display = 'inline';
    sigWrap1.classList.remove('active'); sigWrap1.style.display = 'none';
    sigWrap2.classList.add('active');    sigWrap2.style.display = 'block';
  } else {
    block2Print.style.display = 'none';
    sigWrap1.classList.add('active');    sigWrap1.style.display = 'block';
    sigWrap2.classList.remove('active'); sigWrap2.style.display = 'none';
    document.getElementById('p_nxk').textContent = nx !== dots ? tc(nx) : '............................';
    document.getElementById('p_ndk').textContent = nd !== dots ? tc(nd) : '............................';
  }

  document.getElementById('p_sl').textContent = document.getElementById('sl').value || '......';
  document.getElementById('p_ld').textContent = document.getElementById('ld').value || '...............................................................................';
  const now = new Date();
  document.getElementById('p_ng').textContent = now.getDate().toString().padStart(2,'0');
  document.getElementById('p_th').textContent = (now.getMonth()+1).toString().padStart(2,'0');
  document.getElementById('p_na').textContent = now.getFullYear();
}

document.querySelectorAll('input, select, textarea').forEach(e => {
  e.addEventListener('input', bind);
  e.addEventListener('change', bind);
});

// Trigger gợi ý khi thay đổi người hoặc ngày
['nx', 'ddi', 'cdi'].forEach(id => {
  const el = document.getElementById(id);
  if (el) el.addEventListener('change', capNhatGoiY);
});

// ===== COLLECT FORM DATA =====
function collectFormData() {
  const has2 = document.getElementById('chk2').checked;
  const now = new Date();
  const dots = '.................................................';
  const nx  = document.getElementById('nx').value  || '............................';
  const nd  = document.getElementById('nd').value  || '............................';
  const nd2 = has2 ? (document.getElementById('nd2').value || '............................') : null;
  const cdi = document.getElementById('cdi').value;
  const ddi = document.getElementById('ddi').value;
  const cve = document.getElementById('cve').value;
  const dve = document.getElementById('dve').value;
  return {
    nx, nd, nd2,
    c1:  (cdi&&ddi) ? cdi+' ngày '+fd(ddi) : dots,
    c2:  (cve&&dve) ? cve+' ngày '+fd(dve) : dots,
    c1b: (has2&&document.getElementById('cdi2').value&&document.getElementById('ddi2').value)
      ? document.getElementById('cdi2').value+' ngày '+fd(document.getElementById('ddi2').value) : dots,
    c2b: (has2&&document.getElementById('cve2').value&&document.getElementById('dve2').value)
      ? document.getElementById('cve2').value+' ngày '+fd(document.getElementById('dve2').value) : dots,
    sl:  document.getElementById('sl').value || '......',
    ld:  document.getElementById('ld').value || '...............................................................................',
    cd:  document.querySelector('input[name="loai"]:checked').value==='truongca' ? 'trưởng ca' : 'nhân viên',
    ng:  now.getDate().toString().padStart(2,'0'),
    th:  (now.getMonth()+1).toString().padStart(2,'0'),
    na:  now.getFullYear()
  };
}

// ===== BUILD HTML ĐƠN =====
const DON_CSS = `
  @page{size:A4;margin:20mm 18mm 20mm 25mm}
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Times New Roman',Times,serif;font-size:13pt;line-height:1.6;color:#000;background:#fff}
  .ph{text-align:center;margin-bottom:2px}
  .ul{font-weight:bold;text-decoration:underline}
  .pd{text-align:right;margin:4px 0 14px;font-style:italic;font-size:12pt}
  .pt{text-align:center;font-size:15pt;font-weight:bold;margin:10px 0}
  .pkg{text-align:center;margin-bottom:14px;font-size:12pt}
  .pc p{text-indent:10mm;margin-bottom:3px;text-align:justify;font-size:12pt}
  .sig-tbl{width:100%;border-collapse:collapse;margin-top:20px}
  .sig-tbl td{text-align:center;font-size:12pt;padding:0;border:none}
  .sig-space td{height:14px}
`;

function buildDonHTML(d) {
  const has2 = !!d.nd2;
  const sp2 = '<tr class="sig-space"><td></td><td></td></tr>'.repeat(5);
  const sp3 = '<tr class="sig-space"><td></td><td></td><td></td></tr>'.repeat(5);
  const sp1 = '<tr class="sig-space"><td></td></tr>'.repeat(5);
  const dtBlock = `<table class="sig-tbl" style="margin-top:16px">
    <tr><td style="text-align:center">Đội trưởng</td></tr>${sp1}
    <tr><td style="text-align:center;font-weight:bold">Nguyễn Văn Trung</td></tr>
  </table>`;
  const sigHTML = has2 ? `
    <table class="sig-tbl"><tr><td>Người đổi (1)</td><td>Người đổi (2)</td><td>Người viết đơn</td></tr>
    ${sp3}<tr><td>${tc(d.nd)}</td><td>${tc(d.nd2)}</td><td>${tc(d.nx)}</td></tr></table>${dtBlock}
  ` : `
    <table class="sig-tbl"><tr><td>Người đổi</td><td>Người viết đơn</td></tr>
    ${sp2}<tr><td>${tc(d.nd)}</td><td>${tc(d.nx)}</td></tr></table>${dtBlock}
  `;
  const ndLine = has2 ? `${d.nd} và ông ${d.nd2}` : d.nd;
  const para2 = has2 ? `
    <p>${d.c1b}, ông ${d.nd2} sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.</p>
    <p>${d.c2b}, tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông ${d.nd2}.</p>` : '';
  return `
    <div class="ph"><b>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</b><br><span class="ul">Độc lập – Tự do – Hạnh phúc</span></div>
    <div class="pd"><i>TP. Hồ Chí Minh, ngày ${d.ng} tháng ${d.th} năm ${d.na}.</i></div>
    <div class="pt">ĐƠN XIN ĐỔI CA LÀM VIỆC</div>
    <div class="pkg"><b>Kính gửi:</b> Đội trưởng Đội Vận hành, bảo trì đường hầm.</div>
    <div class="pc">
      <p>Tôi tên là: ${d.nx}, ${d.cd} Đội Vận hành, bảo trì đường hầm.</p>
      <p>Nay tôi viết đơn này xin phép cho tôi đổi ca với ông ${ndLine}.</p>
      <p>${d.c1}, ông ${d.nd} sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.</p>
      <p>${d.c2}, tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông ${d.nd}.</p>
      ${para2}
      <p>Số lần đã đổi ca trong tháng: ${d.sl}.</p>
      <p>Lý do: ${d.ld}.</p>
      <p>Rất mong được sự chấp thuận của Đội trưởng.</p>
      <p>Trân trọng cảm ơn.</p>
    </div>${sigHTML}`;
}

// ===== TỰ ĐỘNG LƯU ĐƠN =====
async function autoSaveDon(d, nguon) {
  try {
    const o = {
      nguoiXin: d.nx,
      nguoiDoi: d.nd,
      caDi: d.c1,
      caVe: d.c2,
      soLan: d.sl,
      lyDo: d.ld,
      loai: d.cd === 'trưởng ca' ? 'truongca' : 'nhanvien',
      nguon: nguon, // 'in' | 'pdf' | 'word'
      createdAt: new Date().toISOString()
    };
    if (d.nd2) {
      o.nguoiDoi2 = d.nd2;
      o.caDi2 = d.c1b;
      o.caVe2 = d.c2b;
    }
    const r = await fetch(API + '.json', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(o)
    });
    if (r.ok) { load(); return true; }
  } catch(e) { console.error('autoSave error:', e); }
  return false;
}

// ===== PRINT =====
document.getElementById('btnPrint').onclick = async () => {
  bind();
  const d = collectFormData();

  // Kiểm tra giới hạn
  const { soLan, vuotGioiHan } = kiemTraGioiHan(d.nx);
  if (vuotGioiHan) {
    alert(`🚫 ${d.nx} đã đổi ca ${soLan}/${GIOI_HAN_THANG} lần tháng này!\nKhông thể in thêm đơn.`);
    return;
  }

  const html = `<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8">
    <title>Don_Doi_Ca_${tc(d.nx||'Moi').replace(/ /g,'_')}</title>
    <style>${DON_CSS}</style></head><body>${buildDonHTML(d)}
    <script>window.onload=function(){window.print();window.onafterprint=function(){window.close();};};<\/script>
  </body></html>`;
  const w = window.open('', '_blank', 'width=900,height=700');
  if (!w) { alert('Trình duyệt chặn popup! Vui lòng cho phép popup cho trang này.'); return; }
  w.document.write(html);
  w.document.close();

  // Tự động lưu sau khi in
  await autoSaveDon(d, 'in');
};

// ===== EXPORT PDF =====
document.getElementById('btnPdf').onclick = async () => {
  bind();
  const d = collectFormData();

  const { soLan, vuotGioiHan } = kiemTraGioiHan(d.nx);
  if (vuotGioiHan) {
    alert(`🚫 ${d.nx} đã đổi ca ${soLan}/${GIOI_HAN_THANG} lần tháng này!`);
    return;
  }

  const html = `<!DOCTYPE html><html lang="vi"><head><meta charset="UTF-8">
    <title>Don_Doi_Ca_${tc(d.nx||'Moi').replace(/ /g,'_')}</title>
    <style>${DON_CSS}</style></head><body>${buildDonHTML(d)}
    <script>window.onload=function(){window.print();window.onafterprint=function(){window.close();};};<\/script>
  </body></html>`;
  const w = window.open('', '_blank', 'width=900,height=700');
  if (!w) { alert('Trình duyệt chặn popup!'); return; }
  w.document.write(html);
  w.document.close();

  // Tự động lưu
  await autoSaveDon(d, 'pdf');
};

// ===== EXPORT WORD =====
document.getElementById('btnWord').onclick = async () => {
  bind();
  const has2 = document.getElementById('chk2').checked;
  const d = collectFormData();

  const { soLan, vuotGioiHan } = kiemTraGioiHan(d.nx);
  if (vuotGioiHan) {
    alert(`🚫 ${d.nx} đã đổi ca ${soLan}/${GIOI_HAN_THANG} lần tháng này!`);
    return;
  }

  const base = {
    nx: d.nx, nd: d.nd,
    cdi: document.getElementById('p_c1').textContent,
    cve: document.getElementById('p_c2').textContent,
    sl: d.sl, ld: d.ld,
    loai: document.querySelector('input[name="loai"]:checked').value,
    ng: d.ng, th: d.th, na: d.na
  };
  if (has2) {
    base.nd2 = d.nd2;
    base.cdi2 = document.getElementById('p_c1_b').textContent;
    base.cve2 = document.getElementById('p_c2_b').textContent;
  }
  expDocx(base, 'Don_Doi_Ca_' + tc(base.nx||'Moi').replace(/ /g,'_'));

  // Tự động lưu
  await autoSaveDon(d, 'word');
};

// ===== SAVE FORM (nút LƯU ĐƠN) =====
document.getElementById('frm').addEventListener('submit', async function(e) {
  e.preventDefault();
  const tenNguoi = document.getElementById('nx').value;
  const has2 = document.getElementById('chk2').checked;

  const { soLan, vuotGioiHan } = kiemTraGioiHan(tenNguoi);
  if (vuotGioiHan) {
    alert(`🚫 ${tenNguoi} đã đổi ca ${soLan}/${GIOI_HAN_THANG} lần tháng này!\nKhông thể đổi ca thêm.`);
    return;
  }

  const o = {
    nguoiXin: tenNguoi,
    nguoiDoi: document.getElementById('nd').value,
    caDi: document.getElementById('cdi').value + ' ngày ' + fd(document.getElementById('ddi').value),
    caVe: document.getElementById('cve').value + ' ngày ' + fd(document.getElementById('dve').value),
    soLan: document.getElementById('sl').value,
    lyDo: document.getElementById('ld').value,
    loai: document.querySelector('input[name="loai"]:checked').value,
    nguon: 'luu',
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
    const conLai = GIOI_HAN_THANG - soLan - 1;
    alert(`✅ Đã lưu thành công!\n📊 ${tenNguoi} còn ${conLai} lần đổi ca trong tháng này.`);
  } catch { alert('❌ Lưu thất bại!'); }
});

// ===== LOAD + FILTER + CHART =====
let chartInstance = null;

async function load() {
  try {
    const r = await fetch(API + '.json'), d = await r.json();
    let a = [];
    if (d && typeof d === 'object') { for (let k in d) a.push({ id: k, ...d[k] }); }
    a.sort((x, y) => new Date(y.createdAt||0) - new Date(x.createdAt||0));
    window.rec = a;

    // Stats
    const now2 = new Date();
    const thisMonth = a.filter(i => { const d=new Date(i.createdAt); return d.getMonth()===now2.getMonth()&&d.getFullYear()===now2.getFullYear(); });
    const today = a.filter(i => { const d=new Date(i.createdAt); return d.toDateString()===now2.toDateString(); });
    document.getElementById('stat-total').textContent = a.length;
    document.getElementById('stat-month').textContent = thisMonth.length;
    document.getElementById('stat-today').textContent = today.length;

    capNhatCanhBao();
    renderTable(a);
    renderChart(a);
  } catch(e) {
    document.getElementById('tbody').innerHTML = '<tr><td colspan="6" class="empty" style="color:var(--red)">Lỗi kết nối!</td></tr>';
  }
}

// ===== RENDER TABLE =====
function renderTable(a) {
  const filterTen = document.getElementById('filterTen')?.value || '';
  const filterThang = document.getElementById('filterThang')?.value || '';

  let filtered = a;
  if (filterTen) filtered = filtered.filter(i => i.nguoiXin === filterTen);
  if (filterThang) {
    const [yyyy, mm] = filterThang.split('-');
    filtered = filtered.filter(i => {
      const d = new Date(i.createdAt);
      return d.getMonth()+1 === +mm && d.getFullYear() === +yyyy;
    });
  }

  const tb = document.getElementById('tbody');
  if (!filtered.length) { tb.innerHTML = '<tr><td colspan="6" class="empty">Không có đơn nào</td></tr>'; return; }
  tb.innerHTML = '';
  filtered.forEach(i => {
    const lb = i.loai==='truongca'
      ? '<span class="badge badge-tc">⭐ Trưởng ca</span>'
      : '<span class="badge badge-nv">👤 Nhân viên</span>';
    const nguoiDoi = i.nguoiDoi2
      ? `<strong>${i.nguoiDoi}</strong><br><small style="color:var(--green)">+ ${i.nguoiDoi2}</small>`
      : `<strong>${i.nguoiDoi}</strong>`;
    const caDetail = i.nguoiDoi2
      ? `<div class="ca-box"><span style="color:var(--red);font-size:.75rem">▲</span> ${i.caDi}</div>
         <div class="ca-box"><span style="color:var(--green);font-size:.75rem">▼</span> ${i.caVe}</div>
         <div class="ca-box" style="opacity:.7"><span style="color:var(--red);font-size:.75rem">▲</span> ${i.caDi2||''}</div>
         <div class="ca-box" style="opacity:.7"><span style="color:var(--green);font-size:.75rem">▼</span> ${i.caVe2||''}</div>`
      : `<div class="ca-box"><span style="color:var(--red);font-size:.75rem">▲</span> ${i.caDi}</div>
         <div class="ca-box"><span style="color:var(--green);font-size:.75rem">▼</span> ${i.caVe}</div>`;

    const thangDon = new Date(i.createdAt);
    const soLanThang = a.filter(x =>
      x.nguoiXin===i.nguoiXin &&
      new Date(x.createdAt).getMonth()===thangDon.getMonth() &&
      new Date(x.createdAt).getFullYear()===thangDon.getFullYear()
    ).length;
    const limitColor = soLanThang>=GIOI_HAN_THANG ? '#f43f5e' : soLanThang>=GIOI_HAN_THANG-1 ? '#f59e0b' : '#00d68f';
    const limitBadge = `<span style="font-size:.65rem;color:${limitColor};margin-left:4px">[${soLanThang}/${GIOI_HAN_THANG}]</span>`;

    // Nguồn đơn
    const nguonMap = { in:'🖨 In', pdf:'📑 PDF', word:'📄 Word', luu:'💾 Lưu' };
    const nguonBadge = i.nguon ? `<br><span style="font-size:.6rem;color:var(--muted)">${nguonMap[i.nguon]||''}</span>` : '';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="Ngày tạo">${new Date(i.createdAt).toLocaleDateString('vi-VN')}<br>
        <small style="color:var(--muted)">${new Date(i.createdAt).toLocaleTimeString('vi-VN')}</small>${nguonBadge}</td>
      <td data-label="Người xin đổi"><strong>${i.nguoiXin}</strong>${limitBadge}<br>${lb}</td>
      <td data-label="Trực thay">${nguoiDoi}</td>
      <td data-label="Chi tiết ca">${caDetail}</td>
      <td data-label="Thao tác" class="td-act">
        <button class="bw" onclick="expOld('${i.id}')">📄 Word</button>
        <button class="bd" onclick="del('${i.id}')">🗑</button>
      </td>`;
    tb.appendChild(tr);
  });
}

// ===== CHART: Đổi ca theo tháng =====
function renderChart(a) {
  const canvasEl = document.getElementById('chartDoiCa');
  if (!canvasEl) return;
  if (typeof Chart === 'undefined') return;

  // Tạo dữ liệu 6 tháng gần nhất
  const now = new Date();
  const labels = [];
  const dataTong = [];
  const dataNV = [];
  const dataTC = [];

  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const mm = d.getMonth();
    const yy = d.getFullYear();
    labels.push(`${d.toLocaleString('vi-VN',{month:'short'})} ${yy}`);
    const thang = a.filter(x => {
      const xd = new Date(x.createdAt);
      return xd.getMonth()===mm && xd.getFullYear()===yy;
    });
    dataTong.push(thang.length);
    dataNV.push(thang.filter(x=>x.loai!=='truongca').length);
    dataTC.push(thang.filter(x=>x.loai==='truongca').length);
  }

  if (chartInstance) chartInstance.destroy();
  chartInstance = new Chart(canvasEl.getContext('2d'), {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label:'Nhân viên', data:dataNV, backgroundColor:'rgba(0,180,255,0.6)', borderColor:'#00b4ff', borderWidth:1.5, borderRadius:3 },
        { label:'Trưởng ca', data:dataTC, backgroundColor:'rgba(240,165,0,0.6)', borderColor:'#f0a500', borderWidth:1.5, borderRadius:3 }
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{ labels:{ color:'#4a7fa8', font:{size:11}, boxWidth:12 } },
        tooltip:{ callbacks:{ footer: (items) => `Tổng: ${items.reduce((s,i)=>s+i.parsed.y,0)} đơn` } }
      },
      scales:{
        x:{ stacked:true, ticks:{color:'#2a5080',font:{size:10}}, grid:{color:'rgba(0,40,80,0.3)'} },
        y:{ stacked:true, ticks:{color:'#2a5080',font:{size:10},stepSize:1}, grid:{color:'rgba(0,40,80,0.3)'}, beginAtZero:true }
      }
    }
  });

  // Render bảng thống kê nhân viên tháng này
  renderThongKeNhanVien(a);
}

// ===== THỐNG KÊ TỪNG NHÂN VIÊN =====
function renderThongKeNhanVien(a) {
  const el = document.getElementById('thongKeNV');
  if (!el) return;
  const now = new Date();
  const thangNay = a.filter(i => {
    const d = new Date(i.createdAt);
    return d.getMonth()===now.getMonth() && d.getFullYear()===now.getFullYear();
  });

  // Đếm theo từng người
  const counter = {};
  thangNay.forEach(i => { counter[i.nguoiXin] = (counter[i.nguoiXin]||0)+1; });
  const sorted = Object.entries(counter).sort((a,b)=>b[1]-a[1]);

  if (!sorted.length) { el.innerHTML = '<div style="color:var(--muted);font-size:.8rem;text-align:center;padding:1rem">Chưa có dữ liệu tháng này</div>'; return; }

  el.innerHTML = sorted.map(([ten,so]) => {
    const pct = Math.round(so/GIOI_HAN_THANG*100);
    const color = so>=GIOI_HAN_THANG ? '#f43f5e' : so>=GIOI_HAN_THANG-1 ? '#f59e0b' : '#00d68f';
    return `<div style="margin-bottom:10px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px">
        <span style="font-size:.8rem;color:var(--text)">${ten}</span>
        <span style="font-size:.75rem;font-weight:700;color:${color}">${so}/${GIOI_HAN_THANG}</span>
      </div>
      <div style="background:rgba(0,10,30,.5);border-radius:3px;height:5px;overflow:hidden">
        <div style="width:${Math.min(pct,100)}%;height:100%;background:${color};border-radius:3px;transition:width .5s"></div>
      </div>
      ${so>=GIOI_HAN_THANG ? '<div style="font-size:.65rem;color:#f43f5e;margin-top:2px">⛔ Đã hết lượt đổi ca tháng này</div>' : ''}
    </div>`;
  }).join('');
}

// ===== FILTER =====
function applyFilter() {
  renderTable(window.rec);
}
window.applyFilter = applyFilter;

function clearFilter() {
  document.getElementById('filterTen').value = '';
  document.getElementById('filterThang').value = '';
  renderTable(window.rec);
}
window.clearFilter = clearFilter;

// ===== EXPORT OLD =====
window.expOld = function(id) {
  const r = window.rec.find(x=>x.id===id); if(!r) return;
  const d = new Date(r.createdAt);
  const data = {
    nx:r.nguoiXin, nd:r.nguoiDoi, cdi:r.caDi, cve:r.caVe,
    sl:r.soLan, ld:r.lyDo, loai:r.loai||'nhanvien',
    ng:d.getDate().toString().padStart(2,'0'),
    th:(d.getMonth()+1).toString().padStart(2,'0'),
    na:d.getFullYear().toString()
  };
  if(r.nguoiDoi2){ data.nd2=r.nguoiDoi2; data.cdi2=r.caDi2; data.cve2=r.caVe2; }
  expDocx(data, 'Don_Doi_Ca_'+tc(r.nguoiXin||'Moi').replace(/ /g,'_'));
};

// ===== DELETE =====
window.del = async function(id) {
  if(!confirm('Xóa đơn này?')) return;
  try { await fetch(API+'/'+id+'.json',{method:'DELETE'}); load(); }
  catch { alert('Lỗi xóa!'); }
};

// ===== EXPORT DOCX =====
async function expDocx(d, fn) {
  if(typeof JSZip==='undefined'){alert('JSZip chưa tải!');return;}
  function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
  function run(t,o){o=o||{};return'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'+(o.bold?'<w:b/><w:bCs/>':'')+(o.italic?'<w:i/><w:iCs/>':'')+(o.ul?'<w:u w:val="single"/>':'')+'<w:sz w:val="'+(o.sz||28)+'"/><w:szCs w:val="'+(o.sz||28)+'"/></w:rPr><w:t xml:space="preserve">'+esc(t)+'</w:t></w:r>';}
  function para(a,o){o=o||{};return'<w:p><w:pPr><w:jc w:val="'+(o.jc||'both')+'"/>'+(o.ind?'<w:ind w:firstLine="709"/>':'')+'<w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr>'+a.map(r=>run(r.t,r)).join('')+'</w:p>';}
  function el(){return'<w:p><w:pPr><w:spacing w:after="0" w:line="180" w:lineRule="exact"/></w:pPr></w:p>';}
  const nb='<w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>';
  function cell(c,w){return'<w:tc><w:tcPr><w:tcW w:w="'+w+'" w:type="dxa"/><w:tcBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tcBorders></w:tcPr>'+c+'</w:tc>';}
  function cp(a){return para(a,{jc:'center'});}

  const TW=9354,W=3118,WM=3118;
  const has2=!!d.nd2;
  const cd=d.loai==='truongca'?'trưởng ca':'nhân viên';
  const emptyRow3='<w:tr>'+cell(el(),W)+cell(el(),WM)+cell(el(),W)+'</w:tr>';
  const emptyRow2='<w:tr>'+cell(el(),W)+cell(el(),WM)+'</w:tr>';
  const emptyRow1='<w:tr>'+cell(el(),TW)+'</w:tr>';

  let sig;
  if(!has2){
    sig='<w:tbl><w:tblPr><w:tblW w:w="'+TW+'" w:type="dxa"/>'+nb+'</w:tblPr><w:tblGrid><w:gridCol w:w="'+W+'"/><w:gridCol w:w="'+WM+'"/><w:gridCol w:w="'+W+'"/></w:tblGrid>'
      +'<w:tr>'+cell(cp([{t:'Người đổi'}]),W)+cell(cp([{t:'Đội trưởng'}]),WM)+cell(cp([{t:'Người viết đơn'}]),W)+'</w:tr>'
      +emptyRow3+emptyRow3+emptyRow3+emptyRow3+emptyRow3
      +'<w:tr>'+cell(cp([{t:cl(d.nd)}]),W)+cell(cp([{t:'Nguyễn Văn Trung',bold:true}]),WM)+cell(cp([{t:cl(d.nx)}]),W)+'</w:tr>'
      +'</w:tbl>';
  } else {
    sig='<w:tbl><w:tblPr><w:tblW w:w="'+TW+'" w:type="dxa"/>'+nb+'</w:tblPr><w:tblGrid><w:gridCol w:w="'+W+'"/><w:gridCol w:w="'+WM+'"/><w:gridCol w:w="'+W+'"/></w:tblGrid>'
      +'<w:tr>'+cell(cp([{t:'Người đổi (1)'}]),W)+cell(cp([{t:'Người đổi (2)'}]),WM)+cell(cp([{t:'Người viết đơn'}]),W)+'</w:tr>'
      +emptyRow3+emptyRow3+emptyRow3+emptyRow3+emptyRow3
      +'<w:tr>'+cell(cp([{t:cl(d.nd)}]),W)+cell(cp([{t:cl(d.nd2)}]),WM)+cell(cp([{t:cl(d.nx)}]),W)+'</w:tr>'
      +'</w:tbl>'
      +'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>'
      +'<w:tbl><w:tblPr><w:tblW w:w="'+TW+'" w:type="dxa"/>'+nb+'</w:tblPr><w:tblGrid><w:gridCol w:w="'+TW+'"/></w:tblGrid>'
      +'<w:tr>'+cell(cp([{t:'Đội trưởng'}]),TW)+'</w:tr>'
      +emptyRow1+emptyRow1+emptyRow1+emptyRow1+emptyRow1
      +'<w:tr>'+cell(cp([{t:'Nguyễn Văn Trung',bold:true}]),TW)+'</w:tr>'
      +'</w:tbl>';
  }

  let body=para([{t:'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',bold:true}],{jc:'center'})
    +para([{t:'Độc lập – Tự do – Hạnh phúc',bold:true,ul:true}],{jc:'center'})
    +para([{t:'TP. Hồ Chí Minh, ngày '+d.ng+' tháng '+d.th+' năm '+d.na+'.',italic:true}],{jc:'right'})
    +el()+para([{t:'ĐƠN XIN ĐỔI CA LÀM VIỆC',bold:true,sz:32}],{jc:'center'})+el()
    +para([{t:'Kính gửi: ',bold:true},{t:'Đội trưởng Đội Vận hành, bảo trì đường hầm.'}],{jc:'center'})+el()
    +para([{t:'Tôi tên là: '},{t:d.nx},{t:', '+cd+' Đội Vận hành, bảo trì đường hầm.'}],{ind:true})
    +(has2
      ? para([{t:'Nay tôi viết đơn này xin phép cho tôi đổi ca với ông '},{t:d.nd},{t:' và ông '},{t:d.nd2},{t:'.'}],{ind:true})
      : para([{t:'Nay tôi viết đơn này xin phép cho tôi đổi ca với ông '},{t:d.nd},{t:'.'}],{ind:true}))
    +para([{t:d.cdi},{t:', ông '},{t:d.nd},{t:' sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.'}],{ind:true})
    +para([{t:d.cve},{t:', tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông '},{t:d.nd},{t:'.'}],{ind:true});
  if(has2){
    body+=para([{t:d.cdi2},{t:', ông '},{t:d.nd2},{t:' sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.'}],{ind:true})
      +para([{t:d.cve2},{t:', tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông '},{t:d.nd2},{t:'.'}],{ind:true});
  }
  body+=para([{t:'Số lần đã đổi ca trong tháng: '},{t:d.sl},{t:'.'}],{ind:true})
    +para([{t:'Lý do: '},{t:d.ld}],{ind:true})+el()
    +para([{t:'Rất mong được sự chấp thuận của Đội trưởng.'}],{ind:true})
    +para([{t:'Trân trọng cảm ơn.'}],{ind:true})+el()+sig
    +'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';

  const docX='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+body+'</w:body></w:document>';
  const ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>';
  const rm='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const rd='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const zip=new JSZip();
  zip.file('[Content_Types].xml',ct);zip.file('_rels/.rels',rm);zip.file('word/document.xml',docX);zip.file('word/_rels/document.xml.rels',rd);
  const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  const url=URL.createObjectURL(blob),a=document.createElement('a');
  document.body.appendChild(a);a.style.display='none';a.href=url;a.download=fn+'.docx';a.click();
  setTimeout(()=>{document.body.removeChild(a);URL.revokeObjectURL(url);},800);
}

// ===== INIT =====
load();
bind();

// ===== AI ASSISTANT =====
const AI_HISTORY = [];

function buildContext() {
  const now = new Date();
  const thangNay = window.rec.filter(i => {
    const d = new Date(i.createdAt);
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  });

  // Thống kê ai đổi nhiều nhất
  const counter = {};
  window.rec.forEach(i => { counter[i.nguoiXin] = (counter[i.nguoiXin] || 0) + 1; });
  const topDoi = Object.entries(counter).sort((a,b) => b[1]-a[1]).slice(0,5);

  // Thống kê theo tháng
  const byMonth = {};
  window.rec.forEach(i => {
    const d = new Date(i.createdAt);
    const key = `${d.getMonth()+1}/${d.getFullYear()}`;
    byMonth[key] = (byMonth[key] || 0) + 1;
  });

  // Cảnh báo giới hạn tháng này
  const canh_bao = [];
  thangNay.forEach(i => {
    const so = thangNay.filter(x => x.nguoiXin === i.nguoiXin).length;
    if (so >= GIOI_HAN_THANG && !canh_bao.find(c => c.ten === i.nguoiXin)) {
      canh_bao.push({ ten: i.nguoiXin, so });
    }
  });

  // Xung đột ca (2 đơn cùng ngày cùng ca)
  const xungDot = [];
  const allDon = window.rec.filter(i => {
    const d = new Date(i.createdAt);
    return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
  });

  // Lịch kíp tóm tắt
  const lichTomTat = LICH_KIP ? Object.keys(LICH_KIP).reduce((acc, ten) => {
    acc[ten] = LICH_KIP[ten].kip;
    return acc;
  }, {}) : {};

  return `Bạn là trợ lý AI của hệ thống Quản lý Đổi Ca UTMC Đường hầm TP.HCM.
Hệ thống 3 ca 4 kíp vận hành thiết bị:
- Ca 1: 6h-14h | Ca 2: 14h-22h | Ca 3: 22h-6h
- Mỗi kíp: 2 ngày Ca1 → 2 ngày Ca2 → 2 ngày Ca3 → nghỉ 2 ngày
- Giới hạn đổi ca: ${GIOI_HAN_THANG} lần/người/tháng
- Kíp VH thiết bị: Kíp 1 (Đặng Thái Nguyên, Phan Hoàng Thanh, Hồng Mạnh Thái, Bùi Văn Hùng, Đặng Ngọc Hòa, Nguyễn Thành Long, Trần Hải Dương, Ngụy Huỳnh Trung), Kíp 2 (Võ Tấn Nghĩa, Nguyễn Văn Hậu, Võ Văn Hoài, Nguyễn Quốc Thái, Phạm Bá Quỳnh, Trần Minh Tân, Nguyễn Hoài Nam, Hồ Trung Hiếu), Kíp 3 (Nguyễn Phương, Nguyễn Nhật Tiến, Đoàn Văn Hạnh, Nguyễn Minh Thảo, Dương Đại Nghĩa, Phan Quốc Nhật, Nguyễn Minh Trung, Nguyễn Thanh Thuận), Kíp 4 (Phạm Anh Tuấn, Võ Đông Đức, Huỳnh Hiếu Thịnh, Bùi Đình Khánh, Nguyễn Thành Tâm, Trần Quang Vinh, Nguyễn Huỳnh, Nguyễn Ngọc Duy)

DỮ LIỆU HIỆN TẠI:
- Tổng đơn đã lưu: ${window.rec.length}
- Đơn tháng ${now.getMonth()+1}/${now.getFullYear()}: ${thangNay.length}
- Top 5 người đổi ca nhiều nhất: ${topDoi.map(([t,s]) => `${t}(${s})`).join(', ')}
- Thống kê theo tháng: ${Object.entries(byMonth).map(([k,v]) => `${k}:${v}đơn`).join(', ')}
- Người đã hết lượt tháng này: ${canh_bao.length ? canh_bao.map(c => `${c.ten}(${c.so}/${GIOI_HAN_THANG})`).join(', ') : 'Không có'}
- 10 đơn gần nhất: ${window.rec.slice(0,10).map(i => `${i.nguoiXin}→${i.nguoiDoi}(${new Date(i.createdAt).toLocaleDateString('vi-VN')})`).join('; ')}

Trả lời bằng tiếng Việt, ngắn gọn, rõ ràng. Dùng emoji phù hợp. Định dạng đẹp.`;
}

function aiQuick(type) {
  const prompts = {
    'phan-tich': 'Phân tích xu hướng đổi ca: ai đổi nhiều nhất, tháng nào cao nhất, nhận xét tổng thể?',
    'canh-bao': 'Phát hiện bất thường và cảnh báo: ai sắp hết lượt, có xung đột ca nào không, điều gì cần chú ý?',
    'goi-y': 'Dựa vào lịch kíp và lịch sử, gợi ý cặp đổi ca phù hợp nhất hiện tại là ai với ai?',
    'tom-tat': 'Tóm tắt tình hình đổi ca tháng này: số lượng, ai đổi, kíp nào nhiều nhất, nhận xét?'
  };
  document.getElementById('ai-input').value = prompts[type];
  sendAI();
}

async function sendAI() {
  const input = document.getElementById('ai-input');
  const msg = input.value.trim();
  if (!msg) return;

  const btn = document.getElementById('ai-send-btn');
  btn.disabled = true;
  btn.textContent = '...';
  input.value = '';

  // Thêm tin nhắn user
  appendAIMsg('user', msg);

  // Thêm typing indicator
  const typingId = appendTyping();

  // Build messages
  AI_HISTORY.push({ role: 'user', content: msg });
  if (AI_HISTORY.length > 10) AI_HISTORY.splice(0, 2);

  try {
    const res = await fetch('/api/ai', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        system: buildContext(),
        messages: AI_HISTORY
      })
    });

    const data = await res.json();
    const reply = data.content?.[0]?.text || 'Xin lỗi, có lỗi xảy ra!';

    removeTyping(typingId);
    AI_HISTORY.push({ role: 'assistant', content: reply });
    appendAIMsg('ai', reply);

  } catch(e) {
    removeTyping(typingId);
    appendAIMsg('ai', '❌ Không kết nối được AI. Kiểm tra lại kết nối mạng.');
  }

  btn.disabled = false;
  btn.textContent = 'GỬI ➤';
  input.focus();
}

function appendAIMsg(role, text) {
  const chat = document.getElementById('ai-chat');
  const div = document.createElement('div');
  div.className = role === 'ai' ? 'ai-msg-ai' : 'ai-msg-user';

  // Convert markdown-like to HTML
  let html = text
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.*?)\*/g, '<em>$1</em>')
    .replace(/`(.*?)`/g, '<code style="background:rgba(0,180,255,.1);padding:1px 5px;border-radius:3px;font-size:.8em">$1</code>')
    .replace(/^#{1,3}\s(.+)$/gm, '<div style="font-weight:700;color:#00b4ff;margin:4px 0">$1</div>')
    .replace(/^[-•]\s(.+)$/gm, '<div style="padding-left:10px">• $1</div>')
    .replace(/\n/g, '<br>');

  div.innerHTML = `
    <div class="ai-msg-icon">${role === 'ai' ? '🤖' : '👤'}</div>
    <div class="ai-msg-text">${html}</div>`;
  chat.appendChild(div);
  chat.scrollTop = chat.scrollHeight;
}

function appendTyping() {
  const chat = document.getElementById('ai-chat');
  const id = 'typing-' + Date.now();
  const div = document.createElement('div');
  div.id = id;
  div.className = 'ai-msg-ai';
  div.innerHTML = `<div class="ai-msg-icon">🤖</div>
    <div class="ai-msg-text">
      <div class="ai-typing"><span></span><span></span><span></span></div>
    </div>`;
  chat.appendChild(div);
  chat.scrollTop = chat.scrollHeight;
  return id;
}

function removeTyping(id) {
  const el = document.getElementById(id);
  if (el) el.remove();
}
