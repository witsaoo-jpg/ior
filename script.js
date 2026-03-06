// ==========================================
// 1. DATA STRUCTURES & CONFIG
// ==========================================
// URL ของ Google Apps Script ที่เชื่อมต่อกับ Sheet
const scriptURL = 'https://script.google.com/macros/s/AKfycbycxJxkj-DiDglZt3Swv2RBWc1swLDawEy9WFGOJQNxpUO1m1AaG149Q7Tm51mMsJDp/exec';

// ฐานข้อมูลกลุ่มงานและหน่วยงาน (อ้างอิงจาก CSV)
const hospitalDepts = {
    "กลุ่มงานการพยาบาลผู้ป่วยอายุรกรรม": [
        { id: "13", name: "หอผู้ป่วย สก.3" }, { id: "14", name: "หอผู้ป่วย สก.4" },
        { id: "15", name: "หอผู้ป่วย สก.5" }, { id: "16", name: "หอผู้ป่วย สก.6" },
        { id: "90", name: "หอผู้ป่วยสงฆ์อาพาธ" }, { id: "91", name: "หอผู้ป่วยสงฆ์พิเศษ 2+3" },
        { id: "11", name: "หอผู้ป่วยสามัญติดเชื้อ ธารน้ำใจ 1" }, { id: "18", name: "หอผู้ป่วยสามัญติดเชื้อ ธารน้ำใจ 2" },
        { id: "40", name: "หอผู้ป่วยพิเศษอายุรกรรมชลาทรล่าง" }, { id: "17", name: "หอผู้ป่วยพิเศษอายุรกรรมชลาทรบน" },
        { id: "29", name: "หอผู้ป่วย Low Immune ธนจ.4" }
    ],
    "กลุ่มงานการพยาบาลผู้ป่วยศัลยกรรม": [
        { id: "22", name: "หอผู้ป่วยชลาทิศ 1" }, { id: "24", name: "หอผู้ป่วยชลาทิศ 2" },
        { id: "39", name: "หอผู้ป่วยชลาทิศ 3" }, { id: "25", name: "หอผู้ป่วยชลาทิศ 4" },
        { id: "27", name: "หอผู้ป่วยพิเศษศัลยกรรม ฉ.7" }, { id: "28", name: "หอผู้ป่วยพิเศษศัลยกรรม ฉ.8" },
        { id: "94", name: "หอผู้ป่วยพิเศษศัลยกรรม Ex.9" }, { id: "23", name: "หอผู้ป่วยแผลไหม้" },
        { id: "34", name: "หอผู้ป่วยคมีบำบัด" }
    ],
    "กลุ่มงานการพยาบาลผู้ป่วยสูติ-นรีเวช": [
        { id: "30", name: "หอผู้ป่วยหลังคลอด" }, { id: "41", name: "หอผู้ป่วยนรีเวช ชลารักษ์4" },
        { id: "42", name: "หอผู้ป่วยพิเศษนรีเวช ชลารักษ์4" }
    ],
    "กลุ่มงานการพยาบาลผู้ป่วยออร์โธปิดิกส์": [
        { id: "81", name: "หอผู้ป่วยกระดูกชาย" }, { id: "82", name: "หอผู้ป่วยศัลยกรรมอุบัติเหตุและกระดูกหญิง" },
        { id: "44", name: "หอผู้ป่วยพิเศษศัลยกรรม Ex.8" }
    ],
    "กลุ่มงานการพยาบาลผู้ป่วย โสต ศอ นาสิก จักษุ": [
        { id: "61", name: "หอผู้ป่วยสามัญ EENT และศัลยกรรมเด็ก ชว.3" }, { id: "10", name: "หอผู้ป่วยพิเศษ EENT" }
    ]
};

// หมวดหมู่และหัวข้อย่อยอุบัติการณ์
const incidentCategories = {
    "1. ความคลาดเคลื่อนทางยา (Medication Error)": ["สั่งใช้ยาผิดพลาด (Prescribing)", "คัดลอกคำสั่งการรักษาผิด (Transcription)", "จัด/จ่ายยาผิดพลาด (Dispensing)", "บริหารยา/ให้ยาผิดพลาด (Administration)"],
    "2. อุบัติเหตุและพลัดตกหกล้ม (Patient Fall)": ["พลัดตกเตียง", "ลื่นล้ม/หกล้มในห้องน้ำ", "หกล้มขณะเดินหรือทำกิจกรรม", "พลัดตกขณะเคลื่อนย้ายผู้ป่วย"],
    "3. สายสวน ท่อ และอุปกรณ์เลื่อนหลุด (Tube & Line)": ["สายน้ำเกลือเลื่อนหลุด/อุดตัน/Phlebitis", "ท่อช่วยหายใจเลื่อนหลุด (Unplanned Extubation)", "สายสวนปัสสาวะ (Foley catheter) เลื่อนหลุด", "สายให้อาหาร (NG/OG Tube) เลื่อนหลุด"],
    "4. เครื่องมือและอุปกรณ์การแพทย์ขัดข้อง": ["เครื่องมือทำงานผิดปกติขณะใช้งาน", "อุปกรณ์ไม่พร้อมใช้งานในภาวะฉุกเฉิน", "เครื่องมือชำรุดเสียหาย"],
    "5. การระบุตัวผู้ป่วยผิดพลาด (Patient ID Error)": ["ทำหัตถการผิดคน / ผิดตำแหน่ง", "เจาะเลือด / เก็บสิ่งส่งตรวจผิดคน", "ติดป้ายข้อมือ (Wristband) สลับคน", "ส่งมอบเวร / เคลื่อนย้ายผู้ป่วยผิดคน"],
    "6. การติดเชื้อในโรงพยาบาล (HAI)": ["ปอดอักเสบ (VAP)", "การติดเชื้อในระบบทางเดินปัสสาวะ (CAUTI)", "การติดเชื้อในกระแสเลือด (CABSI)", "แผลผ่าตัดติดเชื้อ (SSI)"],
    "7. การให้เลือดและส่วนประกอบของเลือด": ["ให้เลือดผิดกรุ๊ป / ผิดคน / ผิดชนิด", "เกิดปฏิกิริยาแพ้เลือด (Transfusion Reaction)"],
    "8. พฤติกรรมและสิทธิผู้ป่วย": ["ผู้ป่วยหลบหนี (Abscond)", "พฤติกรรมก้าวร้าว / ทำร้ายร่างกาย", "การปฏิเสธการรักษา", "ข้อร้องเรียนด้านพฤติกรรมบริการ"],
    "9. ระบบสิ่งแวดล้อมและเทคโนโลยี": ["ระบบไฟฟ้าขัดข้อง / ไฟตก", "ระบบก๊าซทางการแพทย์ขัดข้อง", "ระบบ IT / เครือข่ายล่ม (HIS Down)", "อัคคีภัย / ภัยพิบัติ"]
};

let incidents = JSON.parse(localStorage.getItem('HIRS_Data')) || [];
let currentUser = JSON.parse(sessionStorage.getItem('HIRS_User')) || null;

// ==========================================
// 2. AUTHENTICATION & ROLE MANAGEMENT
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
    if (currentUser) {
        showApp();
    } else {
        document.getElementById('login-page').classList.remove('hidden');
    }
});

// ตรวจสอบ Login ผ่าน Google Sheet
document.getElementById('loginForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const user = document.getElementById('loginUsername').value.trim();
    const pass = document.getElementById('loginPassword').value;
    const btn = this.querySelector('button[type="submit"]');
    
    // UI Loading state
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> กำลังเข้าสู่ระบบ...';
    document.getElementById('loginError').classList.add('hidden');

    // ส่ง Request ไปหา Google Sheet เพื่อตรวจสอบ User
    let formData = new FormData();
    formData.append('action', 'login');
    formData.append('username', user);
    formData.append('password', pass);

    fetch(scriptURL, { method: 'POST', body: formData })
        .then(response => response.json())
        .then(data => {
            if (data.result === 'success') {
                currentUser = data.user;
                sessionStorage.setItem('HIRS_User', JSON.stringify(currentUser));
                showApp();
            } else {
                document.getElementById('loginError').classList.remove('hidden');
            }
        })
        .catch(err => {
            console.error(err);
            alert("ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้ โปรดตรวจสอบอินเทอร์เน็ต");
        })
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = 'เข้าสู่ระบบ';
        });
});

function showApp() {
    document.getElementById('login-page').classList.add('hidden');
    document.getElementById('app-container').classList.remove('hidden');
    document.getElementById('displayUsername').innerHTML = `<i class="fas fa-user-circle"></i> ${currentUser.displayName}`;
    
    const reporterInput = document.getElementById('reporter');
    if (reporterInput) {
        reporterInput.value = currentUser.displayName;
        reporterInput.readOnly = true;
    }
    
    applyRolePermissions();
    populateDropdowns();
    
    // สั่งให้ดึงข้อมูลจาก Sheet ทันทีที่เข้าแอป
    loadDataFromSheet();
    
    // ผู้ใช้ทั่วไปให้เข้าหน้า Form ก่อน / Admin เข้า Dashboard ก่อน
    if(currentUser.role === 'admin') {
        showPage('dashboard');
    } else {
        showPage('report');
    }
}

function applyRolePermissions() {
    const adminNav = document.getElementById('nav-admin');
    const dashboardNav = document.getElementById('nav-dashboard');

    if (currentUser.role === 'user') {
        adminNav.classList.add('hidden');
        dashboardNav.classList.add('hidden'); 
    } else {
        adminNav.classList.remove('hidden');
        dashboardNav.classList.remove('hidden');
    }
}

function logout() {
    sessionStorage.removeItem('HIRS_User');
    currentUser = null;
    document.getElementById('app-container').classList.add('hidden');
    document.getElementById('login-page').classList.remove('hidden');
    document.getElementById('loginForm').reset();
}

// ==========================================
// 3. UI POPULATION & CASCADING DROPDOWNS
// ==========================================
function populateDropdowns() {
    const mainTypeSelect = document.getElementById('mainType');
    const groupSelects = ['departmentGroup', 'dashDeptGroup']; 

    groupSelects.forEach(id => {
        const sel = document.getElementById(id);
        if(sel && sel.options.length <= 1) {
            for (const group in hospitalDepts) {
                sel.add(new Option(group, group));
            }
        }
    });

    if(mainTypeSelect && mainTypeSelect.options.length <= 1) {
        for (const type in incidentCategories) {
            mainTypeSelect.add(new Option(type, type));
        }
        mainTypeSelect.add(new Option("อื่นๆ (โปรดระบุ)", "Other"));
    }
}

function handleGroupChange(groupId, targetDeptId) {
    const groupVal = document.getElementById(groupId).value;
    const deptSelect = document.getElementById(targetDeptId);
    
    deptSelect.innerHTML = `<option value="">-- ${targetDeptId==='dashDept'?'ทุกหน่วยงาน':'เลือกรหัสและหน่วยงาน'} --</option>`; 
    
    if (groupVal) {
        deptSelect.disabled = false;
        hospitalDepts[groupVal].forEach(dept => {
            const optionText = `${dept.id} - ${dept.name}`;
            deptSelect.add(new Option(optionText, optionText));
        });
    } else {
        deptSelect.disabled = true;
    }
}

document.getElementById('departmentGroup')?.addEventListener('change', () => handleGroupChange('departmentGroup', 'department'));
document.getElementById('dashDeptGroup')?.addEventListener('change', () => handleGroupChange('dashDeptGroup', 'dashDept'));

document.getElementById('mainType')?.addEventListener('change', function() {
    const subTypeSelect = document.getElementById('subType');
    const otherInput = document.getElementById('otherType');
    
    subTypeSelect.innerHTML = '<option value="">-- เลือกหัวข้อย่อย --</option>'; 
    otherInput.classList.add('hidden');
    otherInput.required = false;

    if (this.value === "Other") {
        subTypeSelect.disabled = true;
        otherInput.classList.remove('hidden');
        otherInput.required = true;
    } else if (this.value) {
        subTypeSelect.disabled = false;
        incidentCategories[this.value].forEach(sub => subTypeSelect.add(new Option(sub, sub)));
        subTypeSelect.add(new Option("อื่นๆ ในหมวดนี้ (โปรดระบุด้านล่าง)", "OtherSub"));
    } else {
        subTypeSelect.disabled = true;
    }
});

document.getElementById('subType')?.addEventListener('change', function() {
    const otherInput = document.getElementById('otherType');
    if (this.value === "OtherSub") {
        otherInput.classList.remove('hidden');
        otherInput.required = true;
    } else {
        otherInput.classList.add('hidden');
        otherInput.required = false;
        otherInput.value = '';
    }
});

// ==========================================
// 4. NAVIGATION & HELPERS
// ==========================================
function showPage(pageId) {
    document.querySelectorAll('[id^="page-"]').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.sidebar a').forEach(el => el.classList.remove('active'));
    
    document.getElementById(`page-${pageId}`).classList.remove('hidden');
    document.getElementById(`nav-${pageId}`)?.classList.add('active');
    
    const titles = { 'dashboard': 'Dashboard', 'report': 'รายงานความเสี่ยงและอุบัติการณ์', 'list': 'รายการเหตุการณ์', 'admin': 'Admin Panel' };
    document.getElementById('page-title').innerText = titles[pageId] || '';

    if(pageId === 'dashboard') updateDashboard();
    if(pageId === 'list') renderTable();
    if(pageId === 'admin') renderAdminTable();
}

function getSeverityColor(severity) {
    if (['A', 'B'].includes(severity)) return 'badge-severity-A';
    if (['C', 'D'].includes(severity)) return 'badge-severity-C';
    if (['E', 'F'].includes(severity)) return 'badge-severity-E';
    return 'badge-severity-G'; 
}

// สร้างรหัสอัจฉริยะ
function generateSmartID(dateStr, deptStr) {
    let datePart = "";
    if (dateStr) {
        const parts = dateStr.split('-');
        const year = parts[0].substring(2); 
        const month = parts[1];
        const day = parts[2];
        datePart = `${year}${month}${day}`;
    }

    let deptPart = "XX";
    if (deptStr) {
        deptPart = deptStr.split(' ')[0]; 
    }

    const randomPart = Math.random().toString(36).substring(2, 6).toUpperCase();
    return `IR-${datePart}-${deptPart}-${randomPart}`;
}

// ==========================================
// 5. FORM SUBMISSION
// ==========================================
document.getElementById('incidentForm')?.addEventListener('submit', function(e) {
    e.preventDefault();
    const submitBtn = document.getElementById('submitBtn');
    const originalBtnText = submitBtn.innerHTML;
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> กำลังบันทึกข้อมูล...';

    let mainType = document.getElementById('mainType').value;
    let subType = document.getElementById('subType').value;
    let otherDetail = document.getElementById('otherType').value;
    
    let finalMainType = mainType === "Other" ? otherDetail : mainType;
    let finalSubType = subType === "OtherSub" ? otherDetail : (subType || "-");

    const inputDate = document.getElementById('date').value;
    const inputDept = document.getElementById('department').value;
    const smartID = generateSmartID(inputDate, inputDept);

    let formData = new FormData();
    formData.append('action', 'submit'); // บอก GAS ว่านี่คือการ Submit ข้อมูล
    formData.append('incidentId', smartID);
    formData.append('date', inputDate);
    formData.append('time', document.getElementById('time').value);
    formData.append('deptGroup', document.getElementById('departmentGroup').value);
    formData.append('dept', inputDept);
    formData.append('mainType', finalMainType);
    formData.append('subType', finalSubType);
    formData.append('severity', document.getElementById('severity').value);
    formData.append('details', document.getElementById('details').value);
    formData.append('action', document.getElementById('action').value);
    formData.append('reporter', document.getElementById('reporter').value);
    formData.append('status', 'Pending'); // สถานะตั้งต้น

    const newIncident = {
        id: smartID, 
        date: inputDate,
        time: document.getElementById('time').value,
        group: document.getElementById('departmentGroup').value,
        dept: inputDept,
        mainType: finalMainType,
        subType: finalSubType,
        severity: document.getElementById('severity').value,
        details: document.getElementById('details').value,
        action: document.getElementById('action').value,
        reporter: currentUser.displayName,
        status: 'Pending'
    };
    
    incidents.unshift(newIncident); 
    localStorage.setItem('HIRS_Data', JSON.stringify(incidents));

    // ส่งเข้า GAS
    fetch(scriptURL, { method: 'POST', body: formData })
        .then(response => response.json())
        .then(data => {
            if(data.result === 'success') {
                alert(`บันทึกข้อมูลสำเร็จ!\nรหัสอุบัติการณ์ของคุณคือ: ${smartID}`);
            } else {
                alert(`เกิดข้อผิดพลาดในการบันทึกบนคลาวด์ แต่บันทึกในเครื่องแล้ว\nรหัส: ${smartID}`);
            }
        })
        .catch(err => {
            alert(`ระบบออฟไลน์: บันทึกในเบราว์เซอร์แล้ว\nรหัสอุบัติการณ์: ${smartID}`);
        })
        .finally(() => {
            completeSubmit(this, submitBtn, originalBtnText);
        });
});

function completeSubmit(form, btn, originalText) {
    form.reset();
    document.getElementById('reporter').value = currentUser.displayName;
    document.getElementById('otherType').classList.add('hidden');
    document.getElementById('department').disabled = true;
    document.getElementById('subType').disabled = true;
    btn.disabled = false;
    btn.innerHTML = originalText;
    showPage('list');
}

// ==========================================
// ดึงข้อมูลทั้งหมดจาก Google Sheet
// ==========================================
function loadDataFromSheet() {
    // ปรับ UI เล็กน้อยให้รู้ว่ากำลังโหลดข้อมูล
    const originalTitle = document.getElementById('page-title').innerText;
    document.getElementById('page-title').innerHTML = `${originalTitle} <span class="fs-6 text-warning"><i class="fas fa-spinner fa-spin"></i> กำลังซิงค์ข้อมูล...</span>`;

    fetch(scriptURL) // ส่ง Request ไปแบบ GET
        .then(response => response.json())
        .then(data => {
            if (data.result === 'success') {
                // นำข้อมูลจาก Sheet มาแทนที่ในตัวแปร incidents
                incidents = data.data; 
                
                // อัปเดตข้อมูลลง LocalStorage ด้วย
                localStorage.setItem('HIRS_Data', JSON.stringify(incidents));
                
                // สั่งให้รีเฟรชกราฟและตารางใหม่ทั้งหมดด้วยข้อมูลล่าสุด
                updateDashboard();
                renderTable();
                renderAdminTable();
                
                document.getElementById('page-title').innerText = originalTitle; // คืนค่าชื่อ Title
            } else {
                console.error("Sheet Error:", data.message);
                document.getElementById('page-title').innerText = originalTitle;
            }
        })
        .catch(err => {
            console.error("Fetch Error:", err);
            document.getElementById('page-title').innerHTML = `${originalTitle} <span class="fs-6 text-danger"><i class="fas fa-wifi"></i> ออฟไลน์ (ใช้ข้อมูลในเครื่อง)</span>`;
        });
}

// ==========================================
// 6. DASHBOARD CHARTS & FILTERS
// ==========================================
let chartInstances = {};
function updateDashboard() {
    const sDate = document.getElementById('dashStartDate').value;
    const eDate = document.getElementById('dashEndDate').value;
    const gFilter = document.getElementById('dashDeptGroup').value;
    const dFilter = document.getElementById('dashDept').value;

    let filteredData = incidents.filter(item => {
        let pass = true;
        if (sDate && item.date < sDate) pass = false;
        if (eDate && item.date > eDate) pass = false;
        if (gFilter && item.group !== gFilter) pass = false;
        if (dFilter && item.dept !== dFilter) pass = false;
        return pass;
    });

    document.getElementById('stat-total').innerText = filteredData.length;
    document.getElementById('stat-low').innerText = filteredData.filter(i => ['A','B'].includes(i.severity)).length;
    document.getElementById('stat-med').innerText = filteredData.filter(i => ['C','D'].includes(i.severity)).length;
    document.getElementById('stat-high').innerText = filteredData.filter(i => ['E','F','G','H','I'].includes(i.severity)).length;

    const monthlyCounts = new Array(12).fill(0);
    const deptCounts = {};
    const sevCounts = {};

    filteredData.forEach(i => {
        if(i.date) monthlyCounts[new Date(i.date).getMonth()]++;
        let targetArea = gFilter && !dFilter ? i.dept.split(' - ')[0] : (i.group || 'อื่นๆ');
        if(dFilter) targetArea = i.dept.split(' - ')[0]; 
        deptCounts[targetArea] = (deptCounts[targetArea] || 0) + 1;
        if(i.severity) sevCounts[i.severity] = (sevCounts[i.severity] || 0) + 1;
    });

    renderChart('monthlyChart', 'bar', ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'], monthlyCounts, 'จำนวนอุบัติการณ์', '#004b93');
    renderChart('deptChart', 'doughnut', Object.keys(deptCounts), Object.values(deptCounts), 'พื้นที่');
    renderChart('severityChart', 'pie', Object.keys(sevCounts), Object.values(sevCounts), 'ระดับความรุนแรง');
}

function renderChart(canvasId, type, labels, data, label, color = null) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    if (chartInstances[canvasId]) chartInstances[canvasId].destroy();

    let bgColors = color;
    if (type === 'pie' || type === 'doughnut') {
        bgColors = labels.map(l => ['A','B'].includes(l) ? '#28a745' : ['C','D'].includes(l) ? '#ffc107' : ['E','F'].includes(l) ? '#fd7e14' : '#dc3545');
    }

    chartInstances[canvasId] = new Chart(ctx, {
        type: type,
        data: { labels: labels, datasets: [{ label: label, data: data, backgroundColor: bgColors, borderWidth: 1 }] },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

// ==========================================
// 7. LIST & ADMIN VIEWS
// ==========================================
function renderTable() {
    const tbody = document.getElementById('incidentTableBody');
    tbody.innerHTML = '';
    const search = document.getElementById('searchInput').value.toLowerCase();
    const filterSev = document.getElementById('filterSeverity').value;

    const filtered = incidents.filter(item => {
        const matchSearch = (item.details||'').toLowerCase().includes(search) || (item.mainType||'').toLowerCase().includes(search) || (item.dept||'').toLowerCase().includes(search) || (item.id||'').toLowerCase().includes(search);
        let matchSev = true;
        if (filterSev === 'Low') matchSev = ['A','B'].includes(item.severity);
        else if (filterSev === 'Medium') matchSev = ['C','D'].includes(item.severity);
        else if (filterSev === 'High') matchSev = ['E','F','G','H','I'].includes(item.severity);
        return matchSearch && matchSev;
    });

    filtered.forEach(item => {
        let statusBadge = '<span class="badge bg-secondary">ไม่มีสถานะ</span>';
        if(item.status === 'Pending' || !item.status) statusBadge = '<span class="badge bg-warning text-dark"><i class="fas fa-clock"></i> รอทบทวน</span>';
        if(item.status === 'Reviewed') statusBadge = '<span class="badge bg-info text-dark"><i class="fas fa-search"></i> กำลังดำเนินการ</span>';
        if(item.status === 'Resolved') statusBadge = '<span class="badge bg-success"><i class="fas fa-check-circle"></i> ปิดความเสี่ยง</span>';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${item.date}<br><small class="text-muted">${item.time}</small></td>
            <td><span class="badge bg-light text-dark border">${(item.dept||'').split(' - ')[0]}</span><br><small class="text-muted" style="font-size:0.75rem;">${item.id}</small></td>
            <td><strong>${item.mainType}</strong><br><small class="text-muted">${item.subType}</small></td>
            <td>
                <span class="badge ${getSeverityColor(item.severity)}">Level ${item.severity}</span><br>
                ${statusBadge}
            </td>
            <td><button class="btn btn-sm btn-outline-primary" onclick="showDetail('${item.id}')"><i class="fas fa-search"></i></button></td>
        `;
        tbody.appendChild(tr);
    });
}

document.getElementById('searchInput')?.addEventListener('keyup', renderTable);
document.getElementById('filterSeverity')?.addEventListener('change', renderTable);

function showDetail(id) {
    const item = incidents.find(i => i.id === id);
    if(!item) return;
    document.getElementById('modal-date').innerText = `${item.date} ${item.time}`;
    document.getElementById('modal-group').innerText = item.group;
    document.getElementById('modal-dept').innerText = item.dept;
    document.getElementById('modal-type').innerText = `${item.mainType} > ${item.subType}`;
    document.getElementById('modal-severity').innerHTML = `<span class="badge ${getSeverityColor(item.severity)}">Level ${item.severity}</span>`;
    document.getElementById('modal-reporter').innerText = item.reporter;
    document.getElementById('modal-action').innerText = item.action || '-';
    document.getElementById('modal-details').innerHTML = `<strong>รหัสอ้างอิง:</strong> ${item.id} <br><br> ${item.details}`;

    // จัดการส่วนของการตอบกลับ (Reply) ตาม Role
    const replySection = document.getElementById('reply-section');
    if (currentUser.role === 'admin') {
        // เพิ่ม Dropdown ให้เลือกสถานะขณะตอบกลับ
        replySection.innerHTML = `
            <div class="mb-2">
                <label class="form-label small fw-bold text-dark">ปรับเปลี่ยนสถานะเคสนี้:</label>
                <select class="form-select form-select-sm mb-2" id="adminReplyStatus" style="width: auto;">
                    <option value="Pending" ${item.status === 'Pending' || !item.status ? 'selected' : ''}>รอทบทวน</option>
                    <option value="Reviewed" ${item.status === 'Reviewed' ? 'selected' : ''}>กำลังดำเนินการ</option>
                    <option value="Resolved" ${item.status === 'Resolved' ? 'selected' : ''}>ปิดความเสี่ยง</option>
                </select>
            </div>
            <textarea class="form-control mb-2" id="adminReplyText" rows="3" placeholder="พิมพ์การตอบกลับ / ข้อเสนอแนะเพื่อพัฒนาความปลอดภัย...">${item.reply || ''}</textarea>
            <button class="btn btn-success btn-sm" onclick="saveReply('${item.id}')" id="btnSaveReply">
                <i class="fas fa-save"></i> บันทึกการตอบกลับและสถานะ
            </button>
        `;
    } else {
        const replyText = item.reply ? item.reply : '<span class="text-muted">ยังไม่มีการตอบกลับจาก Admin</span>';
        replySection.innerHTML = `
            <div class="p-3 rounded border" style="background-color: #d4edda; color: #155724;">
                ${replyText.replace(/\n/g, '<br>')}
            </div>
        `;
    }

    new bootstrap.Modal(document.getElementById('detailModal')).show();
}

function renderAdminTable() {
    const tbody = document.getElementById('adminTableBody');
    if(!tbody) return;
    tbody.innerHTML = '';
    incidents.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${item.id}</td><td>${item.date}</td><td>${item.group}</td><td>${item.dept}</td><td>${item.mainType}</td><td>${item.severity}</td><td>${item.reporter}</td>`;
        tbody.appendChild(tr);
    });
}

function exportCSV() {
    let csv = "ID,Date,Time,Group,Department,MainType,SubType,Severity,Details,Action,Reporter,Status\n";
    incidents.forEach(r => {
        csv += `${r.id},${r.date},${r.time},${r.group},${r.dept},${r.mainType},${r.subType},${r.severity},"${(r.details||'').replace(/"/g, '""')}","${(r.action||'').replace(/"/g, '""')}",${r.reporter},${r.status||'Pending'}\n`;
    });
    const link = document.createElement("a");
    link.href = encodeURI("data:text/csv;charset=utf-8,\uFEFF" + csv); 
    link.download = "RIRS_Data.csv";
    document.body.appendChild(link);
    link.click();
    link.remove();
}

// ฟังก์ชันสำหรับ Admin บันทึกการตอบกลับและสถานะ
function saveReply(id) {
    const replyText = document.getElementById('adminReplyText').value;
    const newStatus = document.getElementById('adminReplyStatus').value; // ดึงค่าสถานะจาก Dropdown
    const btn = document.getElementById('btnSaveReply');
    const oldText = btn.innerHTML;
    
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> กำลังบันทึก...';

    // 1. อัปเดตข้อมูลและ "สถานะ" ในเครื่องทันที
    const index = incidents.findIndex(i => i.id === id);
    if(index !== -1) {
        incidents[index].reply = replyText;
        incidents[index].status = newStatus; // ใช้สถานะที่ Admin เลือก
        localStorage.setItem('HIRS_Data', JSON.stringify(incidents));
        renderTable();      // อัปเดตตารางหลัก
        renderAdminTable(); // อัปเดตตารางหน้า Admin
    }

    // 2. ส่งข้อมูลไปอัปเดตใน Google Sheet
    let formData = new FormData();
    formData.append('action', 'reply'); 
    formData.append('incidentId', id);
    formData.append('reply', replyText);
    formData.append('status', newStatus); // ส่งสถานะใหม่ไปบันทึกด้วย

    fetch(scriptURL, { method: 'POST', body: formData })
        .then(response => response.json())
        .then(data => {
            if (data.result === 'success') {
                alert('บันทึกการตอบกลับและอัปเดตสถานะสำเร็จ!');
                // ปิด Modal อัตโนมัติเมื่อบันทึกเสร็จ
                bootstrap.Modal.getInstance(document.getElementById('detailModal')).hide();
            } else {
                alert(`เกิดข้อผิดพลาดจากเซิร์ฟเวอร์: ${data.message}`);
            }
        })
        .catch(err => {
            console.error(err);
            alert("ระบบเครือข่ายมีปัญหา แต่บันทึกการตอบกลับและสถานะไว้ในเครื่องแล้ว");
        })
        .finally(() => {
            btn.disabled = false;
            btn.innerHTML = oldText;
        });
}
