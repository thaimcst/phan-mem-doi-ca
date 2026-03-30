// 1. Đồng hồ thời gian thực
function updateClock() {
    const now = new Date();
    document.getElementById('clock').textContent = now.toLocaleTimeString('vi-VN');
}
setInterval(updateClock, 1000);
updateClock();

// 2. URL tới Database Đám mây Firebase
const API_URL = 'https://phan-d-default-rtdb.asia-southeast1.firebasedatabase.app/don_doi_ca';
window.allRecords = [];

// 3. Tải Dữ liệu từ Server và hiển thị vào Table
async function loadData() {
    try {
        const res = await fetch(`${API_URL}.json`);
        const data = await res.json();
        
        let arr = [];
        if (data && typeof data === 'object') {
            for (let key in data) {
                arr.push({ id: key, ...data[key] });
            }
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
                    <div style="background:rgba(255,255,255,0.05); padding:6px; border-radius:4px; margin-bottom:4px; width: 100%;">
                        <small style="color:#ef4444">Để lại ca:</small> <span style="font-weight:600; color:#e2e8f0">${item.caDi}</span>
                    </div>
                    <div style="background:rgba(255,255,255,0.05); padding:6px; border-radius:4px; width: 100%;">
                        <small style="color:#22c55e">Nhận về ca:</small> <span style="font-weight:600; color:#e2e8f0">${item.caVe}</span>
                    </div>
                </td>
                <td class="td-actions" style="display:flex; width:100%">
                    <button class="btn btn-secondary btn-sm" onclick="exportWordOldRecord('${item.id}')" title="Tải File Word Bản Này" style="flex:1">📄 Tải Word Về Zalo</button>
                    <button class="btn btn-sm" style="background:#ef4444;color:white; width: 44px; display:flex; align-items:center; justify-content:center;" onclick="deleteRecord('${item.id}')" title="Xóa">🗑️</button>
                </td>
            `;
            tbody.appendChild(tr);
        });
    } catch (e) {
        console.error('Lỗi kết nối Firebase', e);
        document.getElementById('dataBody').innerHTML = `<tr><td colspan="5" style="text-align:center; color:#ef4444">Lỗi kết nối Máy chủ Đám mây! Vui lòng kiểm tra quyền.</td></tr>`;
    }
}

// 4. Đồng bộ dữ liệu Nhập ở Form sang Khung In (Print Area)
function bindPrintData() {
    const defaultDotNames = '........................................................................';
    const defaultDotCa = '.................................................';
    
    const nguoi_xin = document.getElementById('inp_nguoi_xin').value || defaultDotNames;
    const nguoi_doi = document.getElementById('inp_nguoi_doi').value || defaultDotNames;
    
    document.getElementById('p_nguoi_xin').textContent = nguoi_xin;
    document.getElementById('p_nguoi_doi').textContent = nguoi_doi;
    document.getElementById('p_nguoi_doi_2').textContent = nguoi_doi;
    document.getElementById('p_nguoi_doi_3').textContent = nguoi_doi;
    
    let strCaDi = document.getElementById('sel_ca_di').value && document.getElementById('date_di').value 
        ? `${document.getElementById('sel_ca_di').value} ngày ${window.formatDateVN(document.getElementById('date_di').value)}` : '';
    let strCaVe = document.getElementById('sel_ca_ve').value && document.getElementById('date_ve').value 
        ? `${document.getElementById('sel_ca_ve').value} ngày ${window.formatDateVN(document.getElementById('date_ve').value)}` : '';
        
    document.getElementById('p_ca_1').textContent = strCaDi || defaultDotCa;
    document.getElementById('p_ca_2').textContent = strCaVe || defaultDotCa;
    document.getElementById('p_so_lan').textContent = document.getElementById('inp_so_lan').value || '......';
    document.getElementById('p_ly_do').textContent = document.getElementById('inp_ly_do').value || '....................................................................................................................................';
    
    document.getElementById('p_nguoi_xin_ky').textContent = document.getElementById('inp_nguoi_xin').value ? window.toTitleCase(document.getElementById('inp_nguoi_xin').value) : '............................';
    document.getElementById('p_nguoi_doi_ky').textContent = document.getElementById('inp_nguoi_doi').value ? window.toTitleCase(document.getElementById('inp_nguoi_doi').value) : '............................';
    
    const now = new Date();
    document.getElementById('p_ngay').textContent = now.getDate().toString().padStart(2, '0');
    document.getElementById('p_thang').textContent = (now.getMonth() + 1).toString().padStart(2, '0');
    document.getElementById('p_nam').textContent = now.getFullYear();
}

window.formatDateVN = function(dateStr) {
    if (!dateStr) return "";
    const parts = dateStr.split('-');
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

window.toTitleCase = function(str) {
  return str.toLowerCase().split(' ').map(function(word) {
    return (word.charAt(0).toUpperCase() + word.slice(1));
  }).join(' ');
}

document.querySelectorAll('input, select, textarea').forEach(el => {
    el.addEventListener('input', bindPrintData);
    el.addEventListener('change', bindPrintData);
});

// 5. Tải file Word trực tiếp từ form
document.getElementById('btnPrintPreview').addEventListener('click', () => {
    bindPrintData();
    const data = collectFormData();
    let tenNguoi = window.toTitleCase(document.getElementById('inp_nguoi_xin').value || "Moi");
    buildAndDownloadDocx(data, 'Don_Doi_Ca_' + tenNguoi.replace(/ /g, '_'));
});

// Thu thập dữ liệu hiện tại từ form
function collectFormData() {
    const now = new Date();
    return {
        nguoiXin: document.getElementById('inp_nguoi_xin').value || '................................................................',
        nguoiDoi: document.getElementById('inp_nguoi_doi').value || '................................................................',
        caDi: document.getElementById('p_ca_1').textContent,
        caVe: document.getElementById('p_ca_2').textContent,
        soLan: document.getElementById('inp_so_lan').value || '......',
        lyDo: document.getElementById('inp_ly_do').value || '....',
        ngay: now.getDate().toString().padStart(2, '0'),
        thang: (now.getMonth() + 1).toString().padStart(2, '0'),
        nam: now.getFullYear().toString(),
    };
}

// =========================================================
// HÀM TẠO FILE .DOCX CHUẨN OOXML — KHÔNG DÙNG html-docx-js
// Dùng thư viện docx.js (CDN) để tạo file Word thật sự
// =========================================================
async function buildAndDownloadDocx(d, filename) {
    // Đảm bảo thư viện docx đã load
    if (typeof docx === 'undefined') {
        alert('Thư viện tạo Word chưa tải xong. Vui lòng thử lại sau vài giây!');
        return;
    }

    const {
        Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, WidthType, BorderStyle, VerticalAlign
    } = docx;

    const fontName = 'Times New Roman';
    const fontSize = 28; // 14pt = 28 half-points

    // Helper: tạo đoạn văn thụt đầu dòng, justify
    function para(children, opts = {}) {
        return new Paragraph({
            children,
            alignment: opts.align || AlignmentType.JUSTIFIED,
            indent: opts.indent ? { firstLine: 567 } : undefined, // ~1cm
            spacing: { after: 80 },
        });
    }

    // Helper: tạo TextRun font chung
    function run(text, opts = {}) {
        return new TextRun({
            text,
            font: fontName,
            size: opts.size || fontSize,
            bold: opts.bold || false,
            italics: opts.italic || false,
            underline: opts.underline ? {} : undefined,
        });
    }

    // Đường viền vô hình cho bảng ký tên
    const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

    // Bảng 3 ô chữ ký
    const sigTable = new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3008, 3010, 3008],
        borders: { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder, insideH: noBorder, insideV: noBorder },
        rows: [
            // Hàng 1: tiêu đề 3 cột
            new TableRow({
                children: [
                    new TableCell({
                        borders: noBorders,
                        width: { size: 3008, type: WidthType.DXA },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('Người đổi', { bold: true })] })],
                    }),
                    new TableCell({
                        borders: noBorders,
                        width: { size: 3010, type: WidthType.DXA },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('Người viết đơn', { bold: true })] })],
                    }),
                    new TableCell({
                        borders: noBorders,
                        width: { size: 3008, type: WidthType.DXA },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('Đội trưởng', { bold: true })] })],
                    }),
                ]
            }),
            // Hàng 2: khoảng trắng ký tên (4 dòng trống)
            ...[1,2,3,4].map(() => new TableRow({
                children: [
                    new TableCell({ borders: noBorders, width: { size: 3008, type: WidthType.DXA }, children: [new Paragraph({ children: [run('')] })] }),
                    new TableCell({ borders: noBorders, width: { size: 3010, type: WidthType.DXA }, children: [new Paragraph({ children: [run('')] })] }),
                    new TableCell({ borders: noBorders, width: { size: 3008, type: WidthType.DXA }, children: [new Paragraph({ children: [run('')] })] }),
                ]
            })),
            // Hàng 3: tên người ký
            new TableRow({
                children: [
                    new TableCell({
                        borders: noBorders,
                        width: { size: 3008, type: WidthType.DXA },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(window.toTitleCase(d.nguoiDoi), { bold: true })] })],
                    }),
                    new TableCell({
                        borders: noBorders,
                        width: { size: 3010, type: WidthType.DXA },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(window.toTitleCase(d.nguoiXin), { bold: true })] })],
                    }),
                    new TableCell({
                        borders: noBorders,
                        width: { size: 3008, type: WidthType.DXA },
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run('Nguyễn Văn Trung', { bold: true })] })],
                    }),
                ]
            }),
        ]
    });

    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    size: { width: 11906, height: 16838 }, // A4
                    margin: { top: 1134, right: 850, bottom: 1134, left: 1701 } // trên/dưới 2cm, trái 3cm, phải 1.5cm
                }
            },
            children: [
                // QUỐC HIỆU
                para([run('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { bold: true, size: 28 })], { align: AlignmentType.CENTER }),
                para([run('Độc lập – Tự do – Hạnh phúc', { bold: true, underline: true, size: 28 })], { align: AlignmentType.CENTER }),

                // Dòng trống
                para([run('')]),

                // Ngày tháng năm
                para([run(`TP. Hồ Chí Minh, ngày ${d.ngay} tháng ${d.thang} năm ${d.nam}.`, { italic: true })], { align: AlignmentType.RIGHT }),

                // Dòng trống
                para([run('')]),

                // TIÊU ĐỀ
                para([run('ĐƠN XIN ĐỔI CA LÀM VIỆC', { bold: true, size: 32 })], { align: AlignmentType.CENTER }),

                // Dòng trống
                para([run('')]),

                // Kính gửi
                para([run('Kính gửi: ', { bold: true }), run('Đội trưởng Đội Vận hành, bảo trì đường hầm.')], { align: AlignmentType.CENTER }),

                // Dòng trống
                para([run('')]),

                // Nội dung
                para([
                    run('Tôi tên là: '),
                    run(d.nguoiXin, { bold: true }),
                    run(', nhân viên Đội Vận hành, bảo trì đường hầm.'),
                ], { indent: true }),

                para([
                    run('Nay tôi viết đơn này xin phép cho tôi đổi ca với ông '),
                    run(d.nguoiDoi, { bold: true }),
                    run('.'),
                ], { indent: true }),

                para([
                    run('Ca '),
                    run(d.caDi, { bold: true }),
                    run(', ông '),
                    run(d.nguoiDoi, { bold: true }),
                    run(' sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của tôi.'),
                ], { indent: true }),

                para([
                    run('Ca '),
                    run(d.caVe, { bold: true }),
                    run(', tôi sẽ chịu trách nhiệm hoàn toàn vào ca làm việc của ông '),
                    run(d.nguoiDoi, { bold: true }),
                    run('.'),
                ], { indent: true }),

                para([
                    run('Số lần xin đổi ca trong tháng: '),
                    run(d.soLan, { bold: true }),
                    run('.'),
                ], { indent: true }),

                para([
                    run('Lý do: '),
                    run(d.lyDo),
                ], { indent: true }),

                para([run('')]),

                para([run('Rất mong được sự chấp thuận của Đội trưởng.')], { indent: true }),
                para([run('Trân trọng cảm ơn.')], { indent: true }),

                // Dòng trống trước bảng ký
                para([run('')]),
                para([run('')]),

                // Bảng chữ ký
                sigTable,
            ]
        }]
    });

    try {
        const buffer = await Packer.toBlob(doc);
        const url = URL.createObjectURL(buffer);
        const a = document.createElement('a');
        document.body.appendChild(a);
        a.style.display = 'none';
        a.href = url;
        a.download = filename + '.docx';
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 600);
    } catch(err) {
        console.error('Lỗi tạo docx:', err);
        alert('Lỗi tạo file Word: ' + err.message);
    }
}

// 6. Gửi Đơn lên Server (Lưu Form)
document.getElementById('handoverForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const newRecord = {
        nguoiXin: document.getElementById('inp_nguoi_xin').value,
        nguoiDoi: document.getElementById('inp_nguoi_doi').value,
        caDi: `${document.getElementById('sel_ca_di').value} ngày ${window.formatDateVN(document.getElementById('date_di').value)}`,
        caVe: `${document.getElementById('sel_ca_ve').value} ngày ${window.formatDateVN(document.getElementById('date_ve').value)}`,
        soLan: document.getElementById('inp_so_lan').value,
        lyDo: document.getElementById('inp_ly_do').value,
        createdAt: new Date().toISOString()
    };
    
    try {
        const response = await fetch(`${API_URL}.json`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(newRecord)
        });
        
        if (!response.ok) throw new Error('Bạn quên cấp quyền Test Mode');
        
        document.getElementById('handoverForm').reset();
        bindPrintData();
        loadData();
        alert('✅ Đã ném thẳng lên Đám Mây Google thành công!');
        
    } catch(err) {
        alert('❌ Lưu gặp sự cố! Hãy kiểm tra lại Quy tắc Quyền truy cập Test Mode trên Firebase!');
    }
});

// 7. Tải lại 1 bản ghi cũ ra Word
window.exportWordOldRecord = function(id) {
    const record = window.allRecords.find(x => x.id === id);
    if (record) {
        const date = new Date(record.createdAt);
        const d = {
            nguoiXin: record.nguoiXin,
            nguoiDoi: record.nguoiDoi,
            caDi: record.caDi,
            caVe: record.caVe,
            soLan: record.soLan,
            lyDo: record.lyDo,
            ngay: date.getDate().toString().padStart(2, '0'),
            thang: (date.getMonth() + 1).toString().padStart(2, '0'),
            nam: date.getFullYear().toString(),
        };
        let tenNguoi = window.toTitleCase(record.nguoiXin || "Moi");
        buildAndDownloadDocx(d, 'Don_Doi_Ca_' + tenNguoi.replace(/ /g, '_'));
    }
}

// 8. Xóa 1 bản ghi
window.deleteRecord = async function(id) {
    if (confirm('⚠️ Bạn có chắc chắn muốn XÓA bản ghi này khỏi máy chủ không? Hành động này không thể hoàn tác.')) {
        try {
            await fetch(`${API_URL}/${id}.json`, { method: 'DELETE' });
            loadData();
        } catch(e) {
            alert('Lỗi xóa. Vui lòng thử lại!');
        }
    }
}

// KHỞI TẠO
loadData();
bindPrintData();
