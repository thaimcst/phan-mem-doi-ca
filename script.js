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
        // Sắp xếp giảm dần theo thời gian tạo
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
    
    // Tên người ký (bỏ dấu chấm nếu đang trống)
    document.getElementById('p_nguoi_xin_ky').textContent = document.getElementById('inp_nguoi_xin').value ? window.toTitleCase(document.getElementById('inp_nguoi_xin').value) : '............................';
    document.getElementById('p_nguoi_doi_ky').textContent = document.getElementById('inp_nguoi_doi').value ? window.toTitleCase(document.getElementById('inp_nguoi_doi').value) : '............................';
    
    
    // Cập nhật Ngày tháng năm hiện tại trên văn bản In
    const now = new Date();
    document.getElementById('p_ngay').textContent = now.getDate().toString().padStart(2, '0');
    document.getElementById('p_thang').textContent = (now.getMonth() + 1).toString().padStart(2, '0');
    document.getElementById('p_nam').textContent = now.getFullYear();
}

// Hàm chuẩn hóa định dạng ngày từ kiểu (YYYY-MM-DD của HTML) sang hiển thị VN (DD/MM/YYYY)
window.formatDateVN = function(dateStr) {
    if (!dateStr) return "";
    const parts = dateStr.split('-');
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

// Chuẩn hóa tên viết hoa chữ cái đầu (VD: nguyễn văn a -> Nguyễn Văn A)
window.toTitleCase = function(str) {
  return str.toLowerCase().split(' ').map(function(word) {
    return (word.charAt(0).toUpperCase() + word.slice(1));
  }).join(' ');
}

// Lắng nghe sự kiện người dùng Gõ phím vào Form để tự đồng bộ Data in
document.querySelectorAll('input, select, textarea').forEach(el => {
    el.addEventListener('input', bindPrintData);
    el.addEventListener('change', bindPrintData);
});

// 5. Tải file Word trực tiếp từ form
document.getElementById('btnPrintPreview').addEventListener('click', () => {
    bindPrintData(); // Đồng bộ dữ liệu
    let tenNguoi = window.toTitleCase(document.getElementById('inp_nguoi_xin').value || "Moi");
    exportHTMLToWord('printArea', 'Don_Doi_Ca_' + tenNguoi.replace(/ /g, '_'));
});

// Hàm tạo và tải file Word (doc) từ HTML
function exportHTMLToWord(elementId, filename) {
    const printArea = document.getElementById(elementId);
    let htmlContent = printArea.innerHTML;
    
    // Khởi tạo HTML chuẩn cho MS Word
    const preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export Word Document</title><style>body { font-family: 'Times New Roman', Times, serif; font-size: 14pt;} .print-header { text-align: center; } .print-date { text-align: right; font-style: italic; } .print-title { text-align: center; font-size: 16pt; font-weight: bold; } .print-kto { text-align: center; } .print-content p { text-indent: 1cm; text-align: justify; line-height: 1.2; margin: 4px 0; } </style></head><body>";
    const postHtml = "</body></html>";
    
    // Đã chuyển sang dùng Table ở HTML nên không cần can thiệp JS
    const htmlExport = preHtml + htmlContent + postHtml;
    // Tạo Blob (có ký tự Byte Order Mark cho tiếng Việt chuẩn UTF-8)
    const blob = new Blob(['\ufeff', htmlExport], { type: 'application/msword;charset=utf-8' });
    
    const url = URL.createObjectURL(blob);
    const downloadLink = document.createElement("a");
    downloadLink.href = url;
    downloadLink.download = filename ? filename + '.doc' : 'Don_Xin_Doi_Ca.doc';
    downloadLink.click();
    URL.revokeObjectURL(url);
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
        createdAt: new Date().toISOString() // Thêm thời gian tạo bằng tay
    };
    
    try {
        const response = await fetch(`${API_URL}.json`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(newRecord)
        });
        
        if (!response.ok) {
            throw new Error('Bạn quên cấp quyền Test Mode');
        }
        
        // Sau khi lưu xong
        document.getElementById('handoverForm').reset();
        bindPrintData(); // Reset lại bản in
        loadData();      // Tải lại bảng lịch sử
        
        alert('✅ Đã ném thẳng lên Đám Mây Google thành công!');
        
    } catch(err) {
        alert('❌ Lưu gặp sự cố! Hãy kiểm tra lại Quy tắc Quyền truy cập Test Mode trên Firebase!');
    }
});

// 7. Tải lại 1 bản thu đã có trong Lịch Sử ra Word
window.exportWordOldRecord = function(id) {
    const record = window.allRecords.find(x => x.id === id);
    if(record) {
        // Render dữ liệu record cũ vào Format Print
        document.getElementById('p_nguoi_xin').textContent = record.nguoiXin;
        document.getElementById('p_nguoi_doi').textContent = record.nguoiDoi;
        document.getElementById('p_nguoi_doi_2').textContent = record.nguoiDoi;
        document.getElementById('p_nguoi_doi_3').textContent = record.nguoiDoi;
        
        document.getElementById('p_ca_1').textContent = record.caDi;
        document.getElementById('p_ca_2').textContent = record.caVe;
        document.getElementById('p_so_lan').textContent = record.soLan;
        document.getElementById('p_ly_do').textContent = record.lyDo;
        document.getElementById('p_nguoi_xin_ky').textContent = window.toTitleCase(record.nguoiXin);
        document.getElementById('p_nguoi_doi_ky').textContent = window.toTitleCase(record.nguoiDoi);
        
        
        // Thời gian in là thời gian tạo đơn ban đầu
        const date = new Date(record.createdAt);
        document.getElementById('p_ngay').textContent = date.getDate().toString().padStart(2, '0');
        document.getElementById('p_thang').textContent = (date.getMonth() + 1).toString().padStart(2, '0');
        document.getElementById('p_nam').textContent = date.getFullYear();
        
        // Lệnh xuất file Word thay vì in PDF
        let tenNguoi = window.toTitleCase(record.nguoiXin || "Moi");
        exportHTMLToWord('printArea', 'Don_Doi_Ca_' + tenNguoi.replace(/ /g, '_'));
        
        // Format xong phải gọi bindPrintData để reset PrintArea theo Input Form hiện tại
        bindPrintData();
    }
}

// 8. Xóa 1 bản ghi
window.deleteRecord = async function(id) {
    if(confirm('⚠️ Bạn có chắc chắn muốn XÓA bản ghi này khỏi máy chủ không? Hành động này không thể hoàn tác.')) {
        try {
            await fetch(`${API_URL}/${id}.json`, { method: 'DELETE' });
            loadData();
        } catch(e) {
            alert('Lỗi xóa. Vui lòng thử lại!');
        }
    }
}

// CHẠY HÀM KHỞI TẠO BAN ĐẦU
loadData();      // Load lịch sử Server
bindPrintData(); // Bind lần đầu (để format dấu ........)
