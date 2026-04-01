// =============================================
// 1. ĐỒNG HỒ THỜI GIAN THỰC
// =============================================
function updateClock() {
    document.getElementById('clock').textContent = new Date().toLocaleTimeString('vi-VN');
}
setInterval(updateClock, 1000);
updateClock();

// =============================================
// 2. FIREBASE URL
// =============================================
const API_URL = 'https://phan-d-default-rtdb.asia-southeast1.firebasedatabase.app/don_doi_ca';
window.allRecords = [];

// =============================================
// 3. TẢI DỮ LIỆU TỪ SERVER
// =============================================
async function loadData() {
    try {
        const res = await fetch(`${API_URL}.json`);
        const data = await res.json();
        let arr = [];
        if (data && typeof data === 'object') {
            for (let key in data) arr.push({ id: key, ...data[key] });
        }
        arr.sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
        window.allRecords = arr;

        const tbody = document.getElementById('dataBody');
        tbody.innerHTML = '';
        arr.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td data-label="Lập Đơn Lúc">${new Date(item.createdAt).toLocaleDateString('vi-VN')} <small style="color:#94a3b8">${new Date(item.createdAt).toLocaleTimeString('vi-VN')}</small></td>
                <td data-label="Bạn Xin Đổi"><strong style="font-size:1.1rem">${item.nguoiXin}</strong></td>
                <td data-label="Trực Thay Bạn"><strong style="font-size:1.1rem; color:#e2e8f0">${item.nguoiDoi}</strong></td>
                <td data-label="Thông Tin Ca Trực">
                    <div style="background:rgba(255,255,255,0.05); padding:6px; border-radius:4px; margin-bottom:4px; width:100%;">
                        <small style="color:#ef4444">Để lại ca:</small> <span style="font-weight:600; color:#e2e8f0">${item.caDi}</span>
                    </div>
                    <div style="background:rgba(255,255,255,0.05); padding:6px; border-radius:4px; width:100%;">
                        <small style="color:#22c55e">Nhận về ca:</small> <span style="font-weight:600; color:#e2e8f0">${item.caVe}</span>
                    </div>
                </td>
                <td class="td-actions" style="display:flex; width:100%">
                    <button class="btn btn-secondary btn-sm" onclick="exportWordOldRecord('${item.id}')" style="flex:1">📄 Tải Word Về Zalo</button>
                    <button class="btn btn-sm" style="background:#ef4444;color:white;width:44px;display:flex;align-items:center;justify-content:center;" onclick="deleteRecord('${item.id}')">🗑️</button>
                </td>`;
            tbody.appendChild(tr);
        });
    } catch (e) {
        document.getElementById('dataBody').innerHTML = '<tr><td colspan="5" style="text-align:center;color:#ef4444">Lỗi kết nối Máy chủ Đám mây!</td></tr>';
    }
}

// =============================================
// 4. ĐỒNG BỘ FORM → KHUNG IN
// =============================================
function bindPrintData() {
    var dots = '................................................................';
    var dotsCa = '.................................................';
    var nguoiXin = document.getElementById('inp_nguoi_xin').value || dots;
    var nguoiDoi = document.getElementById('inp_nguoi_doi').value || dots;

    document.getElementById('p_nguoi_xin').textContent = nguoiXin;
    ['p_nguoi_doi','p_nguoi_doi_2','p_nguoi_doi_3'].forEach(function(id) {
        document.getElementById(id).textContent = nguoiDoi;
    });

    var caDiVal   = document.getElementById('sel_ca_di').value;
    var dateDiVal = document.getElementById('date_di').value;
    var caVeVal   = document.getElementById('sel_ca_ve').value;
    var dateVeVal = document.getElementById('date_ve').value;

    document.getElementById('p_ca_1').textContent = (caDiVal && dateDiVal) ? caDiVal + ' ngày ' + formatDateVN(dateDiVal) : dotsCa;
    document.getElementById('p_ca_2').textContent = (caVeVal && dateVeVal) ? caVeVal + ' ngày ' + formatDateVN(dateVeVal) : dotsCa;
    document.getElementById('p_so_lan').textContent = document.getElementById('inp_so_lan').value || '......';
    document.getElementById('p_ly_do').textContent  = document.getElementById('inp_ly_do').value  || '....';

    document.getElementById('p_nguoi_xin_ky').textContent = document.getElementById('inp_nguoi_xin').value ? toTitleCase(document.getElementById('inp_nguoi_xin').value) : '............................';
    document.getElementById('p_nguoi_doi_ky').textContent = document.getElementById('inp_nguoi_doi').value ? toTitleCase(document.getElementById('inp_nguoi_doi').value) : '............................';

    var now = new Date();
    document.getElementById('p_ngay').textContent  = now.getDate().toString().padStart(2,'0');
    document.getElementById('p_thang').textContent = (now.getMonth()+1).toString().padStart(2,'0');
    document.getElementById('p_nam').textContent   = now.getFullYear();
}

function formatDateVN(dateStr) {
    if (!dateStr) return '';
    var p = dateStr.split('-');
    return p[2] + '/' + p[1] + '/' + p[0];
}
window.formatDateVN = formatDateVN;

function toTitleCase(str) {
    return str.toLowerCase().split(' ').map(function(w) {
        return w.charAt(0).toUpperCase() + w.slice(1);
    }).join(' ');
}
window.toTitleCase = toTitleCase;

document.querySelectorAll('input, select, textarea').forEach(function(el) {
    el.addEventListener('input', bindPrintData);
    el.addEventListener('change', bindPrintData);
});

// =============================================
// NÚT IN ĐƠN
// =============================================
document.getElementById('btnPrint').addEventListener('click', function() {
    bindPrintData();
    window.print();
});


document.getElementById('btnPrintPreview').addEventListener('click', function() {
    bindPrintData();
    var now = new Date();
    exportDocx({
        nguoiXin : document.getElementById('inp_nguoi_xin').value || '................................................................',
        nguoiDoi : document.getElementById('inp_nguoi_doi').value || '................................................................',
        caDi     : document.getElementById('p_ca_1').textContent,
        caVe     : document.getElementById('p_ca_2').textContent,
        soLan    : document.getElementById('inp_so_lan').value || '......',
        lyDo     : document.getElementById('inp_ly_do').value  || '....',
        loai     : document.querySelector('input[name="loai_don"]:checked').value,
        ngay     : now.getDate().toString().padStart(2,'0'),
        thang    : (now.getMonth()+1).toString().padStart(2,'0'),
        nam      : now.getFullYear().toString()
    }, 'Don_Doi_Ca_' + toTitleCase(document.getElementById('inp_nguoi_xin').value || 'Moi').replace(/ /g,'_'));
});

// =============================================
// 6. LƯU ĐƠN LÊN FIREBASE
// =============================================
document.getElementById('handoverForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    var newRecord = {
        nguoiXin  : document.getElementById('inp_nguoi_xin').value,
        nguoiDoi  : document.getElementById('inp_nguoi_doi').value,
        caDi      : document.getElementById('sel_ca_di').value + ' ngày ' + formatDateVN(document.getElementById('date_di').value),
        caVe      : document.getElementById('sel_ca_ve').value + ' ngày ' + formatDateVN(document.getElementById('date_ve').value),
        soLan     : document.getElementById('inp_so_lan').value,
        lyDo      : document.getElementById('inp_ly_do').value,
        loai      : document.querySelector('input[name="loai_don"]:checked').value,
        createdAt : new Date().toISOString()
    };
    try {
        var response = await fetch(API_URL + '.json', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(newRecord)
        });
        if (!response.ok) throw new Error();
        document.getElementById('handoverForm').reset();
        bindPrintData();
        loadData();
        alert('✅ Đã lưu lên Đám Mây thành công!');
    } catch(err) {
        alert('❌ Lưu gặp sự cố! Kiểm tra lại quyền Firebase!');
    }
});

// =============================================
// 7. TẢI WORD TỪ LỊCH SỬ
// =============================================
window.exportWordOldRecord = function(id) {
    var r = window.allRecords.find(function(x) { return x.id === id; });
    if (!r) return;
    var d = new Date(r.createdAt);
    exportDocx({
        nguoiXin : r.nguoiXin,
        nguoiDoi : r.nguoiDoi,
        caDi     : r.caDi,
        caVe     : r.caVe,
        soLan    : r.soLan,
        lyDo     : r.lyDo,
        loai     : r.loai || 'nhanvien',
        ngay     : d.getDate().toString().padStart(2,'0'),
        thang    : (d.getMonth()+1).toString().padStart(2,'0'),
        nam      : d.getFullYear().toString()
    }, 'Don_Doi_Ca_' + toTitleCase(r.nguoiXin || 'Moi').replace(/ /g,'_'));
};

// =============================================
// 8. XÓA BẢN GHI
// =============================================
window.deleteRecord = async function(id) {
    if (confirm('⚠️ Bạn có chắc chắn muốn XÓA bản ghi này không?')) {
        try {
            await fetch(API_URL + '/' + id + '.json', { method: 'DELETE' });
            loadData();
        } catch(e) { alert('Lỗi xóa. Vui lòng thử lại!'); }
    }
};

// =============================================
// TẠO FILE .DOCX THUẦN OOXML — CHỈ CẦN JSZIP
// Không phụ thuộc thư viện tạo Word nào khác
// =============================================
async function exportDocx(d, filename) {
    if (typeof JSZip === 'undefined') {
        alert('JSZip chưa tải xong. Vui lòng kiểm tra mạng và thử lại!');
        return;
    }

    // Escape XML
    function esc(str) {
        return String(str || '')
            .replace(/&/g,'&amp;')
            .replace(/</g,'&lt;')
            .replace(/>/g,'&gt;')
            .replace(/"/g,'&quot;');
    }

    // Tạo <w:r> (run) với định dạng
    function run(text, opts) {
        opts = opts || {};
        var bold = opts.bold ? '<w:b/><w:bCs/>' : '';
        var italic = opts.italic ? '<w:i/><w:iCs/>' : '';
        var ul = opts.underline ? '<w:u w:val="single"/>' : '';
        var sz = '<w:sz w:val="' + (opts.size || 28) + '"/><w:szCs w:val="' + (opts.size || 28) + '"/>';
        return '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' +
               bold + italic + ul + sz +
               '</w:rPr><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
    }

    // Tạo <w:p> (đoạn văn)
    function para(runs_arr, opts) {
        opts = opts || {};
        var jc = opts.align || 'both';
        var indent = opts.indent ? '<w:ind w:firstLine="709"/>' : '';
        var spacing = '<w:spacing w:after="0" w:line="276" w:lineRule="auto"/>';
        var runsXml = runs_arr.map(function(r) {
            return run(r.text, r);
        }).join('');
        return '<w:p><w:pPr><w:jc w:val="' + jc + '"/>' + indent + spacing + '</w:pPr>' + runsXml + '</w:p>';
    }

    function emptyLine() {
        return '<w:p><w:pPr><w:spacing w:after="0" w:line="180" w:lineRule="exact"/></w:pPr></w:p>';
    }

    // Ô bảng không viền
    function cell(content, width) {
        var nb = '<w:val w:val="none" w:sz="0" w:space="0" w:color="auto"/>';
        return '<w:tc><w:tcPr><w:tcW w:w="' + width + '" w:type="dxa"/>' +
               '<w:tcBorders>' +
               '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
               '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
               '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
               '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
               '</w:tcBorders></w:tcPr>' + content + '</w:tc>';
    }

    function centerPara(runs_arr) { return para(runs_arr, {align:'center'}); }

    var W1 = 3008, W2 = 3010;

    // Trả về tên sạch (Title Case), bỏ qua nếu là dấu chấm hoặc rỗng
    function cleanName(str) {
        if (!str) return '............................';
        var trimmed = str.replace(/[.]/g, '').trim();
        if (trimmed.length === 0) return '............................';
        return toTitleCase(trimmed);
    }


    // Bảng ký tên đúng mẫu: 3 cột
    // Cột 1: Người đổi (trái) | Cột 2: Đội trưởng (giữa) | Cột 3: Người viết đơn (phải)
    // Hàng 1: tiêu đề 3 cột
    // Hàng 2-4: trống (khoảng ký)
    // Hàng 5: tên người đổi | Nguyễn Văn Trung (bold) | tên người viết
    var W1 = 3000, W2 = 3026, W3 = 3000;
    var tb3 = '<w:tblBorders>' +
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>' +
        '</w:tblBorders>';
    function emptyRow3() {
        return '<w:tr>' + cell(emptyLine(),W1) + cell(emptyLine(),W2) + cell(emptyLine(),W3) + '</w:tr>';
    }
    var sigTable =
        '<w:tbl>' +
        '<w:tblPr><w:tblW w:w="9026" w:type="dxa"/>' + tb3 + '</w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="'+W1+'"/><w:gridCol w:w="'+W2+'"/><w:gridCol w:w="'+W3+'"/></w:tblGrid>' +
        // Hàng tiêu đề
        '<w:tr>' +
            cell(centerPara([{text:'Người đổi', bold:true}]), W1) +
            cell(centerPara([{text:'Đội trưởng', bold:true}]), W2) +
            cell(centerPara([{text:'Người viết đơn', bold:true}]), W3) +
        '</w:tr>' +
        // Khoảng trống ký tên
        emptyRow3() + emptyRow3() + emptyRow3() +
        // Hàng tên
        '<w:tr>' +
            cell(centerPara([{text: cleanName(d.nguoiDoi)}]), W1) +
            cell(centerPara([{text:'Nguyễn Văn Trung', bold:true}]), W2) +
            cell(centerPara([{text: cleanName(d.nguoiXin)}]), W3) +
        '</w:tr>' +
        '</w:tbl>';

    var chucDanh = d.loai === 'truongca' ? 'trưởng ca' : 'nhân viên';
    var body =
        para([{text:'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', bold:true}], {align:'center'}) +
        para([{text:'Độc lập – Tự do – Hạnh phúc', bold:true, underline:true}], {align:'center'}) +
        para([{text:'TP. Hồ Chí Minh, ngày ' + d.ngay + ' tháng ' + d.thang + ' năm ' + d.nam + '.', italic:true}], {align:'right'}) +
        emptyLine() +
        para([{text:'ĐƠN XIN ĐỔI CA LÀM VIỆC', bold:true, size:32}], {align:'center'}) +
        emptyLine() +
        para([{text:'Kính gửi: ', bold:true}, {text:'Đội trưởng Đội Vận hành, bảo trì đường hầm.'}], {align:'center'}) +
        emptyLine() +
        para([{text:'Tôi tên là ' + d.nguoiXin + ', ' + chucDanh + ' Đội Vận hành, bảo trì đường hầm.'}], {indent:true}) +
        para([{text:'Nay tôi viết đơn này xin phép cho tôi đổi ca với ông ' + d.nguoiDoi + '.'}], {indent:true}) +
        para([{text:'Ca ' + d.caDi + ', ông ' + d.nguoiDoi + ' sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.'}], {indent:true}) +
        para([{text:'Ca ' + d.caVe + ', tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông ' + d.nguoiDoi + '.'}], {indent:true}) +
        para([{text:'Số lần đã đổi ca trong tháng: ' + d.soLan + '.'}], {indent:true}) +
        para([{text:'Lý do: ' + d.lyDo}], {indent:true}) +
        emptyLine() +
        para([{text:'Rất mong được sự chấp thuận của Đội trưởng.'}], {indent:true}) +
        para([{text:'Trân trọng cảm ơn.'}], {indent:true}) +
        emptyLine() +
        sigTable +
        '<w:sectPr>' +
        '<w:pgSz w:w="11906" w:h="16838"/>' +
        '<w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>' +
        '</w:sectPr>';

    var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"\n' +
        '  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n' +
        '<w:body>' + body + '</w:body></w:document>';

    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>';

    var relsMain = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>';

    var relsDoc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';

    var zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.file('_rels/.rels', relsMain);
    zip.file('word/document.xml', documentXml);
    zip.file('word/_rels/document.xml.rels', relsDoc);

    var blob = await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    document.body.appendChild(a);
    a.style.display = 'none';
    a.href = url;
    a.download = filename + '.docx';
    a.click();
    setTimeout(function() { document.body.removeChild(a); URL.revokeObjectURL(url); }, 800);
}

// =============================================
// KHỞI TẠO
// =============================================
loadData();
bindPrintData();
