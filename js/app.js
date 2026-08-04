// --- State ---
let rawData = [];
let dataList = [];
let chartClassInstance = null;
let chartTrendInstance = null;
let importRawData = []; // Store Excel data here
let isAdmin = false;
let currentIndicatorConfig = null; // Will load from server
let isDarkMode = false; // State for dark mode
let selectedYear = "2568";
let availableYears = ["2568"];

// Criteria State
let criteriaDataCoop = {}; // Store coop criteria { '1.1': 'html...', ... }
let criteriaDataFG = {}; // Store farmer group criteria
let currentCriteriaCode = null; // Current indicator code being viewed/edited
let currentCriteriaType = null; // 'coop' or 'farmer_group'
let quillEditor = null; // Quill editor instance

// Initialize with default placeholders for Cooperatives (สหกรณ์)
let INDICATOR_INFO = {
    // Dimension 1
    'd1_1': { code: '1.1', desc: 'ผลการดำเนินงาน' },
    'd1_2': { code: '1.2', desc: 'การดำเนินธุรกิจ' },
    'd1_3': { code: '1.3', desc: 'การจัดการทางการเงิน' },
    // Dimension 2
    'd2_1': { code: '2.1', desc: 'ความสามารถในการชำระหนี้' },
    'd2_2': { code: '2.2', desc: 'ความสามารถในการทำกำไร' },
    'd2_3': { code: '2.3', desc: 'ประสิทธิภาพการดำเนินงาน' },
    'd2_4': { code: '2.4', desc: 'ผลตอบแทนต่อส่วนของเจ้าของ' },
    'd2_5': { code: '2.5', desc: 'อัตราการหมุนเวียนของสินทรัพย์' },
    'd2_6': { code: '2.6', desc: 'การดำรงเงินกองทุน' },
    // Dimension 3
    'd3_1': { code: '3.1', desc: 'การควบคุมภายใน' },
    'd3_2': { code: '3.2', desc: 'การจัดทำบัญชี' },
    'd3_3': { code: '3.3', desc: 'การตรวจสอบบัญชี' },
    'd3_4': { code: '3.4', desc: 'การรายงานผลการดำเนินงาน' },
    'd3_5': { code: '3.5', desc: 'การปฏิบัติตามกฎหมาย' },
    // Dimension 4
    'd4_1': { code: '4.1', desc: 'ความพอเพียงของเงินทุน' },
    'd4_2': { code: '4.2', desc: 'คุณภาพสินทรัพย์' },
    'd4_3': { code: '4.3', desc: 'ความสามารถในการทำกำไร (มิติ 4)' },
    'd4_4': { code: '4.4', desc: 'สภาพคล่อง' }
};

// Separate indicator descriptions for Farmer Groups (กลุ่มเกษตรกร)
let FARMER_GROUP_INDICATOR_INFO = {
    // Dimension 1
    'd1_1': { code: '1.1', desc: 'การผลิต/บริการของกลุ่ม' },
    'd1_2': { code: '1.2', desc: 'การดำเนินกิจกรรมกลุ่ม' },
    'd1_3': { code: '1.3', desc: 'การบริหารจัดการกลุ่ม' },
    // Dimension 2
    'd2_1': { code: '2.1', desc: 'ความสามารถในการชำระหนี้' },
    'd2_2': { code: '2.2', desc: 'ความสามารถในการทำกำไร' },
    'd2_3': { code: '2.3', desc: 'ประสิทธิภาพการดำเนินงาน' },
    'd2_4': { code: '2.4', desc: 'ผลตอบแทนต่อส่วนของสมาชิก' },
    'd2_5': { code: '2.5', desc: 'อัตราการหมุนเวียนของสินทรัพย์' },
    'd2_6': { code: '2.6', desc: 'การดำรงเงินกองทุน' },
    // Dimension 3 (กลุ่มเกษตรกรไม่มี 3.5)
    'd3_1': { code: '3.1', desc: 'การควบคุมภายในกลุ่ม' },
    'd3_2': { code: '3.2', desc: 'การจัดทำบัญชี' },
    'd3_3': { code: '3.3', desc: 'การตรวจสอบบัญชี' },
    'd3_4': { code: '3.4', desc: 'การรายงานผลการดำเนินงาน' },
    'd3_5': { code: '3.5', desc: '-' }, // Not applicable for farmer groups
    // Dimension 4
    'd4_1': { code: '4.1', desc: 'ความพอเพียงของเงินทุน' },
    'd4_2': { code: '4.2', desc: 'คุณภาพสินทรัพย์' },
    'd4_3': { code: '4.3', desc: 'ความสามารถในการทำกำไร' },
    'd4_4': { code: '4.4', desc: 'สภาพคล่อง' }
};

// Helper to get correct indicator info based on name
function getIndicatorInfo(name) {
    if (name && name.includes('กลุ่มเกษตรกร')) {
        return FARMER_GROUP_INDICATOR_INFO;
    }
    return INDICATOR_INFO;
}

// --- Admin Logic ---
function adminLogin() {
    Swal.fire({
        title: 'เข้าสู่ระบบผู้ดูแลระบบ',
        input: 'password',
        inputLabel: 'กรุณากรอกรหัสผ่าน',
        inputPlaceholder: 'Password',
        showCancelButton: true,
        confirmButtonText: 'เข้าสู่ระบบ',
        cancelButtonText: 'ยกเลิก'
    }).then((result) => {
        if (result.isConfirmed) {
            // ส่งรหัสผ่านไปตรวจสอบที่ Backend
            google.script.run
                .withSuccessHandler((response) => {
                    if (response.success) {
                        isAdmin = true;
                        updateAdminState();
                        Swal.fire('สำเร็จ', response.message, 'success');
                    } else {
                        Swal.fire('ผิดพลาด', response.message, 'error');
                    }
                })
                .withFailureHandler((err) => {
                    Swal.fire('ผิดพลาด', 'เกิดข้อผิดพลาด: ' + err.message, 'error');
                })
                .verifyAdminPassword(result.value);
        }
    });
}

function adminLogout() {
    isAdmin = false;
    updateAdminState();
    Swal.fire('ออกจากระบบ', 'ออกจากระบบเรียบร้อยแล้ว', 'success');
}

function updateAdminState() {
    // Toggle Admin Button is now handled by Secret Trigger (Hidden by default)
    // if (!isAdmin) { ... } -> We keep it hidden.

    // Show Logout if admin
    const logoutBtn = document.getElementById('admin-logout-btn');
    if (logoutBtn) {
        if (isAdmin) {
            logoutBtn.classList.remove('hidden');
        } else {
            logoutBtn.classList.add('hidden');
        }
    }

    // Toggle Add/Import Buttons
    const btnContainer = document.querySelector('.flex.gap-2'); // Buttons container
    if (btnContainer) {
        // We can toggle children but easier to re-render or toggle specific IDs if we added them.
        // Let's modify the buttons to have IDs or classes.
        // For now, let's just re-render table to toggle Edit/Delete
    }
    // Toggle Add buttons specific logic
    const addBtn = document.querySelector('button[onclick="openModal()"]');
    const importBtn = document.querySelector('button[onclick*="excel-file"]');
    if (addBtn) addBtn.style.display = isAdmin ? 'flex' : 'none';
    if (importBtn) importBtn.style.display = isAdmin ? 'flex' : 'none';

    renderTable();
}

// --- Init ---
window.onload = function () {
    fetchData();
    setupLiveCalc();
    updateAdminState(); // Init restricted state
    setupSecretTrigger();
    initDarkMode(); // Init Theme
};

// --- Dark Mode Logic ---
function initDarkMode() {
    // Check local storage or system preference
    if (localStorage.getItem('theme') === 'dark' ||
        (!('theme' in localStorage) && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
        isDarkMode = true;
        document.documentElement.classList.add('dark');
        document.getElementById('theme-icon').innerText = '☀️';
    } else {
        isDarkMode = false;
        document.documentElement.classList.remove('dark');
        document.getElementById('theme-icon').innerText = '🌙';
    }
}

function toggleDarkMode() {
    isDarkMode = !isDarkMode;
    if (isDarkMode) {
        document.documentElement.classList.add('dark');
        localStorage.setItem('theme', 'dark');
        document.getElementById('theme-icon').innerText = '☀️';
    } else {
        document.documentElement.classList.remove('dark');
        localStorage.setItem('theme', 'light');
        document.getElementById('theme-icon').innerText = '🌙';
    }
    // Re-render charts to update text colors
    if (chartClassInstance || chartTrendInstance) {
        // Simple re-render if data exists
        renderTable();
    }
}

function setupSecretTrigger() {
    let clicks = 0;
    let timer = null;
    const trigger = document.getElementById('secret-login-trigger');

    if (trigger) {
        trigger.addEventListener('click', () => {
            clicks++;
            if (clicks === 1) {
                timer = setTimeout(() => {
                    clicks = 0;
                }, 1000); // Reset after 1 second
            } else if (clicks === 3) {
                clearTimeout(timer);
                clicks = 0;
                if (!isAdmin) adminLogin();
            }
        });
    }
}


function showLoader(show) {
    document.getElementById('loader').style.display = show ? 'flex' : 'none';
}

// --- Data Fetching ---
function fetchData() {
    showLoader(true);
    google.script.run
        .withSuccessHandler(onDataSuccess)
        .withFailureHandler(onDataError)
        .getData(selectedYear);
}

function updateFormLabels(year) {
    const yearNum = parseInt(year) || 2568;
    const prevYearNum = yearNum - 1;
    
    const lblGrade = document.getElementById('label-grade');
    if (lblGrade) lblGrade.textContent = `ผลการจัดชั้นปี ${yearNum}`;
    
    const lblGrade67 = document.getElementById('label-grade_67');
    if (lblGrade67) lblGrade67.textContent = `ผลการจัดชั้นปีก่อนหน้า (${prevYearNum})`;
    
    const lblAgm68 = document.getElementById('label-agm_68');
    if (lblAgm68) lblAgm68.textContent = `การประชุมใหญ่ปีปัจจุบัน (${yearNum})`;
}

function onDataSuccess(response) {
    rawData = response.data; // [{id, no, values:[]}]

    if (response.availableYears) {
        availableYears = response.availableYears;
    }
    if (response.selectedYear) {
        selectedYear = response.selectedYear;
    }

    // Update Coop Indicator Config if provided
    if (response.indicatorConfig) {
        INDICATOR_INFO = response.indicatorConfig;
        currentIndicatorConfig = response.indicatorConfig;
    }

    // Update Farmer Group Indicator Config if provided
    if (response.farmerGroupIndicatorConfig) {
        FARMER_GROUP_INDICATOR_INFO = response.farmerGroupIndicatorConfig;
    }

    // Map array values to easier object structure for frontend usage if needed, or index
    // Mapping Indices based on Code.gs (shifted by 1 for Eval Year in col index 1):
    // 0:No (Ignored from value, use item.no), 1:Year, 2:Agency, 3:Name, 4:Code, 5:Type, 31:Total, 32:Grade, 42:Trend

    dataList = rawData.map(item => {
        const v = item.values;
        return {
            id: item.id,
            no: item.no, // Use dynamic No from backend
            year: v[1],
            agency: v[2],
            name: v[3],
            code: v[4],
            type: v[5],
            total: parseFloat(v[31]) || 0,
            grade: v[32],
            trend: v[42],
            // Advice JSON at Index 43
            advice: JSON.parse(v[43] || '{}'),
            // Store full row for editing
            fullRow: v
        };
    });

    // Populate Year Filter Dropdown
    const yearSelect = document.getElementById('filter-year');
    if (yearSelect) {
        yearSelect.innerHTML = '';
        availableYears.forEach(yr => {
            const opt = new Option('ปีงบประมาณ ' + yr, yr);
            yearSelect.add(opt);
        });
        yearSelect.value = selectedYear;
        
        // Add change listener
        yearSelect.onchange = function() {
            selectedYear = this.value;
            fetchData();
        };
    }
    
    // Update display-year header
    const displayYearSpan = document.getElementById('display-year');
    if (displayYearSpan) {
        displayYearSpan.textContent = selectedYear;
    }

    // Populate Year Select in the Form
    const formYearSelect = document.getElementById('form-eval-year');
    if (formYearSelect) {
        formYearSelect.innerHTML = '';
        const formYears = [...availableYears];
        const nextYear = String(parseInt(availableYears[0] || "2568") + 1);
        if (!formYears.includes(nextYear)) {
            formYears.unshift(nextYear);
        }
        formYears.forEach(yr => {
            formYearSelect.add(new Option(yr, yr));
        });
        formYearSelect.value = selectedYear;
    }

    // Update form labels relative to selected year
    updateFormLabels(selectedYear);

    populateFilters();
    renderTable();
    renderDashboard();
    showLoader(false);

    // Load criteria in background after main data is shown
    loadAllCriteriaData();
}

function onDataError(err) {
    showLoader(false);
    Swal.fire('Error', 'Failed to load data: ' + err.message, 'error');
}

// --- Rendering ---
function populateFilters() {
    // Unique Agencies and Types
    const agencies = [...new Set(dataList.map(d => d.agency).filter(Boolean))].sort();
    const types = [...new Set(dataList.map(d => d.type).filter(Boolean))].sort();

    const agencySelect = document.getElementById('filter-agency');
    const typeSelect = document.getElementById('filter-type');

    // Reset but keep first option
    agencySelect.length = 1;
    typeSelect.length = 1;

    agencies.forEach(a => agencySelect.add(new Option(a, a)));
    types.forEach(t => typeSelect.add(new Option(t, t)));

    // Add Listeners
    agencySelect.onchange = renderTable;
    typeSelect.onchange = renderTable;
    document.getElementById('filter-grade').onchange = renderTable;
    document.getElementById('filter-sector').onchange = renderTable;
    document.getElementById('filter-trend').onchange = renderTable;
    document.getElementById('search-box').addEventListener('keyup', renderTable);
}

function getFilteredData() {
    const agency = document.getElementById('filter-agency').value;
    const type = document.getElementById('filter-type').value;
    const grade = document.getElementById('filter-grade').value;
    const sector = document.getElementById('filter-sector').value;
    const trend = document.getElementById('filter-trend').value;
    const seed = document.getElementById('search-box').value.toLowerCase();

    // Normalize grade for comparison (handle "ชั้น 1" and "ชั้น1")
    const normalizeGrade = (g) => {
        if (!g) return '';
        if (g.includes('1')) return '1';
        if (g.includes('2')) return '2';
        if (g.includes('3')) return '3';
        return g;
    };

    return dataList.filter(d => {
        const matchAgency = !agency || d.agency === agency;
        const matchType = !type || d.type === type;
        const matchGrade = !grade || normalizeGrade(d.grade) === normalizeGrade(grade);
        const matchTrend = !trend || d.trend === trend;
        const matchSearch = !seed ||
            (d.name && d.name.toLowerCase().includes(seed)) ||
            (d.code && d.code.toString().includes(seed));

        let matchSector = true;
        if (sector === 'agro') {
            // สหกรณ์ในภาคการเกษตร
            const agroTypes = ['สหกรณ์การเกษตร', 'สหกรณ์ประมง', 'สหกรณ์นิคม'];
            matchSector = agroTypes.includes(d.type);
        } else if (sector === 'non-agro') {
            // สหกรณ์นอกภาคการเกษตร
            const nonAgroTypes = ['สหกรณ์ออมทรัพย์', 'สหกรณ์บริการ', 'สหกรณ์ร้านค้า', 'สหกรณ์เครดิตยูเนี่ยน'];
            matchSector = nonAgroTypes.includes(d.type);
        } else if (sector === 'farmer_group') {
            // กลุ่มเกษตรกร (Check Name)
            matchSector = d.name.includes('กลุ่มเกษตรกร');
        }

        return matchAgency && matchType && matchGrade && matchTrend && matchSearch && matchSector;
    });
}

function renderTable() {
    const displayData = getFilteredData();

    // Sort by Total Score Descending (มากไปน้อย)
    displayData.sort((a, b) => b.total - a.total);

    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    // Toggle Header Visibility
    const thCode = document.getElementById('th-code');
    if (thCode) {
        if (isAdmin) {
            thCode.classList.remove('hidden');
        } else {
            thCode.classList.add('hidden');
        }
    }

    displayData.forEach((d, index) => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-accent-light/10 dark:hover:bg-claude-darkBorder/20 transition border-b border-claude-border dark:border-claude-darkBorder';

        // Color badge for grade (normalize for both "ชั้น 1" and "ชั้น1")
        let gradeBadge = 'bg-gray-50 text-gray-600 border border-gray-200 dark:bg-gray-900/40 dark:text-gray-300 dark:border-gray-800';
        if (d.grade && d.grade.includes('1')) {
            gradeBadge = 'bg-emerald-50 text-emerald-800 border border-emerald-200/50 dark:bg-emerald-950/30 dark:text-emerald-300 dark:border-emerald-800/40';
        } else if (d.grade && d.grade.includes('2')) {
            gradeBadge = 'bg-amber-50 text-amber-800 border border-amber-200/50 dark:bg-amber-950/30 dark:text-amber-300 dark:border-amber-800/40';
        } else if (d.grade && d.grade.includes('3')) {
            gradeBadge = 'bg-rose-50 text-rose-800 border border-rose-200/50 dark:bg-rose-950/30 dark:text-rose-300 dark:border-rose-800/40';
        }

        // Trend Icon (Elegant Badges with direction arrow)
        let trendIcon = `
            <div class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-semibold bg-gray-50 text-gray-500 border border-gray-200/50 dark:bg-gray-900/30 dark:text-gray-400 dark:border-gray-800/40">
                <span class="text-xs">→</span><span>คงเดิม</span>
            </div>
        `;
        if (d.trend === 'ดีขึ้น') {
            trendIcon = `
                <div class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200/40 dark:bg-emerald-950/30 dark:text-emerald-300 dark:border-emerald-800/40">
                    <span class="text-xs">↑</span><span>ดีขึ้น</span>
                </div>
            `;
        } else if (d.trend === 'แย่ลง') {
            trendIcon = `
                <div class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200/40 dark:bg-rose-950/30 dark:text-rose-300 dark:border-rose-800/40">
                    <span class="text-xs">↓</span><span>แย่ลง</span>
                </div>
            `;
        }

        let actionHtml = `
             <button onclick="viewDetails(${d.id})" class="action-btn text-claude-muted dark:text-claude-darkMuted hover:text-accent dark:hover:text-accent-light p-1.5 rounded-md hover:bg-accent-light/30 dark:hover:bg-claude-darkBorder/40 transition" title="ดูรายละเอียด">
                 <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg>
             </button>
        `;

        if (isAdmin) {
            actionHtml += `
                 <button onclick="editData(${d.id})" class="action-btn text-amber-600 dark:text-amber-400 hover:text-amber-800 dark:hover:text-amber-200 p-1.5 rounded-md hover:bg-amber-50 dark:hover:bg-amber-950/30 transition" title="แก้ไข">
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>
                 </button>
                 <button onclick="deleteData(${d.id})" class="action-btn text-rose-600 dark:text-rose-400 hover:text-rose-800 dark:hover:text-rose-200 p-1.5 rounded-md hover:bg-rose-50 dark:hover:bg-rose-950/30 transition" title="ลบ">
                     <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
                 </button>
            `;
        }

        // Conditionally render Code Cell
        let codeCell = '';
        if (isAdmin) {
            codeCell = `<td class="px-4 py-3 text-sm text-claude-muted dark:text-claude-darkMuted font-mono">${d.code}</td>`;
        }

        tr.innerHTML = `
            <td class="px-4 py-3 text-sm font-medium text-claude-text dark:text-claude-darkText">${index + 1}</td>
            ${codeCell}
            <td class="px-6 py-4 font-semibold text-sm text-claude-text dark:text-claude-darkText">${d.name}</td>
            <td class="px-6 py-4 text-claude-muted dark:text-claude-darkMuted text-xs font-medium">${d.agency}</td>
            <td class="px-6 py-4 text-claude-muted dark:text-claude-darkMuted text-xs font-medium">${d.type}</td>
            <td class="px-6 py-4 text-center font-bold text-base text-claude-text dark:text-claude-darkText">${d.total.toFixed(2)}</td>
            <td class="px-6 py-4 text-center"><span class="badge-grade ${gradeBadge}">${d.grade}</span></td>
            <td class="px-6 py-4 text-center">${trendIcon}</td>
            <td class="px-6 py-4 text-center">
                <div class="flex items-center justify-center gap-1.5">
                    ${actionHtml}
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
    document.getElementById('row-count').innerText = `แสดงทั้งหมด ${displayData.length} รายการ`;
    updateCharts(displayData);
}

function renderDashboard() {
    // Initial Dashboard Render based on full data or triggered by updateCharts
    // We'll let updateCharts handle it
}

function updateCharts(data) {
    // Normalize grade for comparison (handle "ชั้น 1" and "ชั้น1" formats)
    const normalizeGrade = (g) => {
        if (!g) return '';
        if (g.includes('1')) return 'ชั้น 1';
        if (g.includes('2')) return 'ชั้น 2';
        if (g.includes('3')) return 'ชั้น 3';
        return g;
    };

    // Stats
    document.getElementById('stat-total').innerText = data.length;
    document.getElementById('stat-c1').innerText = data.filter(d => normalizeGrade(d.grade) === 'ชั้น 1').length;
    document.getElementById('stat-c2').innerText = data.filter(d => normalizeGrade(d.grade) === 'ชั้น 2').length;
    document.getElementById('stat-c3').innerText = data.filter(d => normalizeGrade(d.grade) === 'ชั้น 3').length;

    // Chart Class
    const counts = { 'ชั้น 1': 0, 'ชั้น 2': 0, 'ชั้น 3': 0 };
    data.forEach(d => {
        const normalized = normalizeGrade(d.grade);
        if (counts[normalized] !== undefined) counts[normalized]++;
    });

    const ctxClass = document.getElementById('chartClass').getContext('2d');
    if (chartClassInstance) chartClassInstance.destroy();

    // Determine Design colors based on theme
    const textColor = isDarkMode ? '#9B978E' : '#7E7970'; // Claude muted colors
    const gridColor = isDarkMode ? '#3A3835' : '#E5E2DD';  // Claude border colors
    const surfaceBorderColor = isDarkMode ? '#242321' : '#ffffff';

    chartClassInstance = new Chart(ctxClass, {
        type: 'doughnut',
        data: {
            labels: ['ชั้น 1', 'ชั้น 2', 'ชั้น 3'],
            datasets: [{
                data: [counts['ชั้น 1'], counts['ชั้น 2'], counts['ชั้น 3']],
                backgroundColor: ['#859F84', '#E2B383', '#DE8A8A'], // Claude Sage, Amber, Rose
                borderColor: surfaceBorderColor,
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: textColor,
                        boxWidth: 12,
                        padding: 15,
                        font: { family: 'Inter, Noto Sans Thai, sans-serif', size: 11 }
                    }
                }
            },
            cutout: '72%'
        }
    });

    // Trend Chart
    const trends = { 'ดีขึ้น': 0, 'คงเดิม': 0, 'แย่ลง': 0 };
    data.forEach(d => { if (trends[d.trend] !== undefined) trends[d.trend]++; });

    const ctxTrend = document.getElementById('chartTrend').getContext('2d');
    if (chartTrendInstance) chartTrendInstance.destroy();

    const trendBgColor = isDarkMode 
        ? ['#859F84', '#3A3835', '#DE8A8A'] 
        : ['#859F84', '#E2E0D9', '#DE8A8A'];

    chartTrendInstance = new Chart(ctxTrend, {
        type: 'bar',
        data: {
            labels: ['ดีขึ้น', 'คงเดิม', 'แย่ลง'],
            datasets: [{
                label: 'จำนวนสหกรณ์',
                data: [trends['ดีขึ้น'], trends['คงเดิม'], trends['แย่ลง']],
                backgroundColor: trendBgColor,
                borderRadius: 6,
                barThickness: 32
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: { 
                        color: textColor,
                        font: { family: 'Inter, Noto Sans Thai, sans-serif', size: 10 }
                    },
                    grid: { color: gridColor }
                },
                x: {
                    ticks: { 
                        color: textColor,
                        font: { family: 'Inter, Noto Sans Thai, sans-serif', size: 11 }
                    },
                    grid: { display: false }
                }
            },
            plugins: {
                legend: {
                    display: false
                }
            }
        }
    });
}

// --- Forms & Modals ---
function openModal() {
    document.getElementById('dataModal').classList.remove('hidden');
    document.getElementById('mainForm').reset();
    document.getElementById('rowId').value = ""; // Mode Add
    document.getElementById('modal-title').innerText = "📝 เพิ่มข้อมูลใหม่";
    updateLiveScore();
    document.body.classList.add('overflow-hidden');
}

function closeModal() {
    document.getElementById('dataModal').classList.add('hidden');
    document.body.classList.remove('overflow-hidden');
}

// --- Details View ---
function viewDetails(id) {
    const item = dataList.find(d => d.id === id);
    if (!item) return;

    document.getElementById('detail-title').innerText = item.name;
    document.getElementById('detail-title').className = "text-2xl font-serif font-bold text-claude-text dark:text-white";

    let subtitle = `<span class="font-medium">สังกัด:</span> ${item.agency} | <span class="font-medium">คะแนนรวม:</span> ${item.total.toFixed(2)} (${item.grade})`;
    if (isAdmin) {
        subtitle = `<span class="font-medium">รหัส:</span> ${item.code} | ` + subtitle;
    }
    subtitle += ` <span id="detail-modal-id" data-id="${item.id}" class="hidden"></span>`;

    document.getElementById('detail-subtitle').innerHTML = subtitle;
    document.getElementById('detail-subtitle').className = "text-sm text-claude-muted dark:text-gray-400 mt-1";

    const v = item.fullRow;
    // Map values similarly to editData (shifted by 1 due to Year column at index 1)
    const dataMap = {
        1: [
            { key: 'd1_1', val: v[6] }, { key: 'd1_2', val: v[7] }, { key: 'd1_3', val: v[8] }
        ],
        2: [
            { key: 'd2_1', val: v[10] }, { key: 'd2_2', val: v[11] }, { key: 'd2_3', val: v[12] },
            { key: 'd2_4', val: v[13] }, { key: 'd2_5', val: v[14] }, { key: 'd2_6', val: v[15] }
        ],
        3: [
            { key: 'd3_1', val: v[17] }, { key: 'd3_2', val: v[18] }, { key: 'd3_3', val: v[19] },
            { key: 'd3_4', val: v[20] }, { key: 'd3_5', val: v[21] }
        ],
        4: [
            { key: 'd4_1', val: v[24] }, { key: 'd4_2', val: v[25] }, { key: 'd4_3', val: v[26] }, { key: 'd4_4', val: v[27] }
        ]
    };

    // Render Tabs
    for (let dim = 1; dim <= 4; dim++) {
        const container = document.getElementById(`tab-content-${dim}`);
        let html = `
            <div class="overflow-hidden rounded-2xl border border-claude-border dark:border-gray-800">
            <table class="w-full text-sm text-left">
                <thead class="text-xs font-bold text-claude-muted dark:text-gray-400 uppercase bg-gray-50 dark:bg-gray-800/50 border-b border-claude-border dark:border-gray-800">
                    <tr>
                        <th class="px-4 py-3 w-24 text-center">ตัวชี้วัด</th>
                        <th class="px-4 py-3">คำอธิบาย</th>
                        <th class="px-4 py-3 text-right w-24">คะแนน</th>
                    </tr>
                </thead>
            <tbody class="divide-y divide-claude-border dark:divide-gray-800">`;

        dataMap[dim].forEach(d => {
            // Get correct indicator info based on name (สหกรณ์ vs กลุ่มเกษตรกร)
            const indicatorSet = getIndicatorInfo(item.name);
            const info = indicatorSet[d.key] || { code: d.key, desc: '-' };
            // Local Advice for this Coop
            const coopAdvice = (item.advice && item.advice[d.key]) ? item.advice[d.key] : "";
            const val = parseFloat(d.val) || 0;

            // Determine criteria type based on item name
            const criteriaType = (item.name && item.name.includes('กลุ่มเกษตรกร')) ? 'farmer_group' : 'coop';

            let descDisplay = info.desc;
            let adviceInput = "";

            if (isAdmin) {
                // Global Desc (Admin) - Still updates global
                descDisplay = `<input type="text" class="border rounded p-1 w-full text-xs bg-yellow-50 focus:bg-white dark:bg-gray-800 dark:border-gray-600 dark:text-white dark:focus:bg-gray-700" 
                    value="${info.desc}" onchange="updateIndicatorField('${d.key}', 'desc', this.value)">`;

                // Local Advice (Admin) - Updates local coop advice
                adviceInput = `<div class="mt-1">
                        <label class="text-[10px] text-gray-400">คำแนะนำ:</label>
                        <textarea class="w-full border rounded p-1 text-xs bg-blue-50 focus:bg-white resize-y" rows="2" placeholder="เพิ่มคำแนะนำเฉพาะสหกรณ์นี้..."
                        onchange="updateLocalAdvice('${d.key}', this.value)">${coopAdvice}</textarea>
                    </div>`;
            } else {
                // Read Only Advice
                if (coopAdvice) {
                    adviceInput = `<div class="mt-2 p-2 bg-gray-50 dark:bg-gray-700 rounded text-xs text-gray-600 dark:text-gray-300 border border-gray-100 dark:border-gray-600">
                        <strong class="text-primary dark:text-primary-400">💡 คำแนะนำ:</strong> ${coopAdvice}
                     </div>`;
                }
            }

            html += `
                <tr class="bg-white dark:bg-gray-800 border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-700 align-top">
                    <td class="px-2 py-3 text-center align-top">
                        <div class="flex items-center justify-center gap-1">
                            <span class="font-bold text-gray-900 dark:text-white">${info.code}</span>
                            <button onclick="viewIndicatorCriteria('${info.code}', '${criteriaType}')" 
                                class="text-info hover:text-blue-700 text-sm p-1 rounded hover:bg-blue-50 dark:hover:bg-blue-900/30 transition" 
                                title="ดูรายละเอียดเกณฑ์">ℹ️</button>
                        </div>
                    </td>
                    <td class="px-2 py-3 text-wrap pr-2">
                        <div>${descDisplay}</div>
                        ${adviceInput}
                    </td>
                    <td class="px-2 py-3 text-right font-medium align-top ${val > 0 ? 'text-primary dark:text-primary-400' : 'text-gray-400 dark:text-gray-500'}">${val.toFixed(2)}</td>
                </tr>
            `;
        });

        // Add Summary Row for each dim
        // Note: v indices for sums: D1(9), D2(16), D3(23), D4(28) (shifted by 1)
        let sum = 0;
        if (dim === 1) sum = v[9];
        if (dim === 2) sum = v[16];
        if (dim === 3) sum = v[23];
        if (dim === 4) sum = v[28];
        sum = parseFloat(sum) || 0;

        html += `
                <tr class="bg-gray-50/50 dark:bg-gray-800/50 font-bold text-claude-text dark:text-white">
                    <td class="px-6 py-4" colspan="2">รวมมิติที่ ${dim}</td>
                    <td class="px-6 py-4 text-right">${sum.toFixed(2)}</td>
                </tr>
            </tbody>
        </table > 
        </div>`;

        container.innerHTML = html;
    }

    // Special Case for Dim 3 Fail Text - render as proper indicator row like other indicators
    // v index shifted by 1: failText at index 22
    const failText = v[22];
    if (failText) {
        const indicatorSet = getIndicatorInfo(item.name);
        const failInfo = indicatorSet['d3_fail'] || { code: 'ภาพรวม', desc: 'ภาพรวมไม่เข้าเกณฑ์' };
        const failAdvice = (item.advice && item.advice['d3_fail']) ? item.advice['d3_fail'] : "";
        const failVal = parseFloat(failText) || 0;
        const criteriaType = (item.name && item.name.includes('กลุ่มเกษตรกร')) ? 'farmer_group' : 'coop';

        let failDescDisplay = failInfo.desc;
        let failAdviceInput = "";

        if (isAdmin) {
            failDescDisplay = `<input type="text" class="border rounded p-1 w-full text-xs bg-yellow-50 focus:bg-white dark:bg-gray-800 dark:border-gray-600 dark:text-white dark:focus:bg-gray-700" 
                value="${failInfo.desc}" onchange="updateIndicatorField('d3_fail', 'desc', this.value)">`;
            failAdviceInput = `<div class="mt-1">
                    <label class="text-[10px] text-gray-400">คำแนะนำ:</label>
                    <textarea class="w-full border rounded p-1 text-xs bg-blue-50 focus:bg-white resize-y" rows="2" placeholder="เพิ่มคำแนะนำเฉพาะสหกรณ์นี้..."
                    onchange="updateLocalAdvice('d3_fail', this.value)">${failAdvice}</textarea>
                </div>`;
        } else {
            if (failAdvice) {
                failAdviceInput = `<div class="mt-2 p-2 bg-gray-50 dark:bg-gray-700 rounded text-xs text-gray-600 dark:text-gray-300 border border-gray-100 dark:border-gray-600">
                    <strong class="text-primary dark:text-primary-400">💡 คำแนะนำ:</strong> ${failAdvice}
                 </div>`;
            }
        }

        // Add as a proper table row matching other indicators
        const failRowHtml = `
            <table class="w-full text-sm text-left text-gray-500 dark:text-gray-400 border-separate border-spacing-y-2 mt-4">
                <thead class="text-xs text-gray-700 dark:text-gray-300 uppercase bg-amber-50 dark:bg-amber-900/30">
                    <tr>
                        <th class="px-2 py-2 w-24 text-center">หมายเหตุ</th>
                        <th class="px-2 py-2">คำอธิบาย</th>
                        <th class="px-2 py-2 text-right w-24">คะแนน</th>
                    </tr>
                </thead>
                <tbody>
                    <tr class="bg-white dark:bg-gray-800 border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-700 align-top">
                        <td class="px-2 py-3 text-center align-top">
                            <div class="flex items-center justify-center gap-1">
                                <span class="font-bold text-amber-600 dark:text-amber-400 text-xs">⚠️ ภาพรวม</span>
                                <button onclick="viewIndicatorCriteria('overview', '${criteriaType}')" 
                                    class="text-info hover:text-blue-700 text-sm p-1 rounded hover:bg-blue-50 dark:hover:bg-blue-900/30 transition" 
                                    title="ดูรายละเอียดเกณฑ์">ℹ️</button>
                            </div>
                        </td>
                        <td class="px-2 py-3 text-wrap pr-2">
                            <div>${failDescDisplay}</div>
                            ${failAdviceInput}
                        </td>
                        <td class="px-2 py-3 text-right font-medium align-top ${failVal > 0 ? 'text-amber-600 dark:text-amber-400' : 'text-gray-400 dark:text-gray-500'}">${failVal.toFixed(2)}</td>
                    </tr>
                </tbody>
            </table>`;

        document.getElementById('tab-content-3').innerHTML += failRowHtml;
    }

    // Reset Tabs to 1
    // Admin Save Button Injection
    const modalFooter = document.getElementById('detail-modal-footer') || document.querySelector('#detailModal .flex.justify-end');
    if (modalFooter) {
        if (isAdmin) {
            modalFooter.innerHTML = `
                <div class="mr-auto flex gap-2">
                     <button type="button" onclick="saveIndicatorConfig()" class="inline-flex justify-center rounded-md border border-transparent shadow-sm px-4 py-2 bg-yellow-500 text-white font-medium hover:bg-yellow-600 focus:outline-none sm:text-sm">💾 บันทึกคำอธิบาย (ทั้งหมด)</button>
                     <button type="button" onclick="saveLocalAdvice()" class="inline-flex justify-center rounded-md border border-transparent shadow-sm px-4 py-2 bg-green-600 text-white font-medium hover:bg-green-700 focus:outline-none sm:text-sm">📨 บันทึกคำแนะนำ (เฉพาะราย)</button>
                </div>
                <button type="button" onclick="closeDetailModal()" class="w-full inline-flex justify-center rounded-md border border-gray-300 shadow-sm px-4 py-2 bg-white text-base font-medium text-gray-700 hover:bg-gray-50 focus:outline-none sm:w-auto sm:text-sm">ปิด</button>
            `;
        } else {
            modalFooter.innerHTML = `<button type="button" onclick="closeDetailModal()" class="w-full inline-flex justify-center rounded-md border border-gray-300 shadow-sm px-4 py-2 bg-white text-base font-medium text-gray-700 hover:bg-gray-50 focus:outline-none sm:mt-0 sm:w-auto sm:text-sm">ปิด</button>`;
        }
    }

    switchTab(1);
    document.getElementById('detailModal').classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
}

function closeDetailModal() {
    document.getElementById('detailModal').classList.add('hidden');
    document.body.classList.remove('overflow-hidden');
}

function switchTab(dim) {
    // Hide all
    document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.tab-btn').forEach(el => {
        el.classList.remove('border-primary', 'text-primary');
        el.classList.add('border-transparent', 'text-gray-500');
    });

    // Show current
    // Show current
    document.getElementById(`tab-content-${dim}`).classList.remove('hidden');
    const btn = document.getElementById(`tab-${dim}`);
    btn.classList.remove('border-transparent', 'text-gray-500');
    btn.classList.add('border-primary', 'text-primary');
}

function editData(id) {
    const item = dataList.find(d => d.id === id);
    if (!item) return;

    openModal();
    document.getElementById('modal-title').innerText = "✏️ แก้ไขข้อมูล: " + item.name;
    const f = document.getElementById('mainForm');
    const v = item.fullRow;

    // Map values back to inputs by index (shifted by 1 for multi-year)
    // Indices: 0:No, 1:EvalYear, 2:Agency, 3:Name, 4:Code, 5:Type
    // 6-8: d1.. 9:Sum1
    // 10-15: d2.. 16:Sum2
    // 17-21: d3.. 22:fail_text 23:Sum3
    // 24-27: d4.. 28:Sum4
    // 33: Remark
    // 34: NoBalanceYears, 35: EvalYearRound (Sheet EvalYear, col index 35), 36: AGM68, 37: Fin68, 38: Acc68
    // 39: Acc67, 40: ChangeNote, 41: Grade67 

    document.getElementById('rowId').value = id;
    f.eval_year.value = v[1] || selectedYear;
    f.agency.value = v[2];
    f.coop_name.value = v[3];
    f.coop_code.value = v[4].replace("'", ""); // Remove escape char
    f.coop_type.value = v[5];

    f.d1_1.value = v[6]; f.d1_2.value = v[7]; f.d1_3.value = v[8];

    f.d2_1.value = v[10]; f.d2_2.value = v[11]; f.d2_3.value = v[12];
    f.d2_4.value = v[13]; f.d2_5.value = v[14]; f.d2_6.value = v[15];

    f.d3_1.value = v[17]; f.d3_2.value = v[18]; f.d3_3.value = v[19];
    f.d3_4.value = v[20]; f.d3_5.value = v[21]; f.d3_fail_text.value = v[22];

    f.d4_1.value = v[24]; f.d4_2.value = v[25]; f.d4_3.value = v[26]; f.d4_4.value = v[27];

    f.remark.value = v[33] || ""; 
    f.no_balance_years.value = v[34];
    f.agm_68.value = v[36];
    f.change_year_note.value = v[40];
    f.grade_67.value = v[41];

    updateLiveScore();
}

function setupLiveCalc() {
    const inputs = document.querySelectorAll('.score-input');
    inputs.forEach(input => {
        input.addEventListener('input', updateLiveScore);
    });
}

function updateLiveScore() {
    const f = document.getElementById('mainForm');
    let sum = 0;
    const parse = (n) => parseFloat(n) || 0;

    // Sum all inputs with class score-input
    const inputs = document.querySelectorAll('.score-input');
    inputs.forEach(i => sum += parse(i.value));

    document.getElementById('live-score').innerText = sum.toFixed(2);
}

function handleFormSubmit(e) {
    e.preventDefault();
    const form = e.target;
    const formData = new FormData(form);
    const obj = {};
    formData.forEach((value, key) => obj[key] = value);

    showLoader(true);
    closeModal();

    google.script.run
        .withSuccessHandler(res => {
            Swal.fire('Success', res.message, 'success');

            // Optimistic / Immediate Update
            if (res.savedRow) {
                showLoader(false);
                const newRowValues = res.savedRow;
                // Construct object like fetchData
                // Values Mapping same as onDataSuccess
                const newItem = {
                    id: parseInt(obj.rowId) || res.savedId, // If edit use existing ID, if add use new
                    no: (parseInt(obj.rowId) || res.savedId) - 1, // Dynamic No derived from Row ID (Row - 1)
                    year: newRowValues[1],
                    agency: newRowValues[2],
                    name: newRowValues[3],
                    code: newRowValues[4],
                    type: newRowValues[5],
                    total: parseFloat(newRowValues[31]) || 0,
                    grade: newRowValues[32],
                    trend: newRowValues[42],
                    advice: JSON.parse(newRowValues[43] || '{}'),
                    fullRow: newRowValues
                };

                if (obj.rowId) {
                    // Edit: Replace in list
                    const idx = dataList.findIndex(d => d.id == obj.rowId);
                    if (idx !== -1) dataList[idx] = newItem;
                } else {
                    // Add: Push to list (and maybe re-sort later? for now just push)
                    dataList.push(newItem);
                }


                // Toggle Add buttons specific logic
                const addBtn = document.querySelector('button[onclick="openModal()"]');
                const importBtn = document.querySelector('button[onclick*="excel-file"]');
                if (addBtn) addBtn.style.display = isAdmin ? 'flex' : 'none';
                if (importBtn) importBtn.style.display = isAdmin ? 'flex' : 'none';

                // Show Logout if admin
                if (isAdmin) {
                    document.getElementById('admin-logout-btn').classList.remove('hidden');
                } else {
                    document.getElementById('admin-logout-btn').classList.add('hidden');
                }

                renderTable(); // Instant render
            } else {
                fetchData(); // Fallback if no row returned
            }
        })
        .withFailureHandler(err => {
            showLoader(false);
            Swal.fire('Error', err.message, 'error');
        })
        .saveData(obj);
}

function deleteData(id) {
    Swal.fire({
        title: 'ยืนยันการลบ?',
        text: "ข้อมูลจะถูกลบถาวร!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#EF5350',
        confirmButtonText: 'ลบข้อมูล'
    }).then((result) => {
        if (result.isConfirmed) {
            showLoader(true);
            google.script.run
                .withSuccessHandler(() => {
                    Swal.fire('Deleted', 'ลบเรียบร้อย', 'success');
                    fetchData();
                })
                .withFailureHandler(err => {
                    showLoader(false);
                    Swal.fire('Error', err.message, 'error');
                })
                .deleteData(id);
        }
    })
}

// --- Excel Import ---
function handleExcelImport(input) {
    const file = input.files[0];
    if (!file) return;

    showLoader(true);
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

            // store raw array of arrays
            importRawData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            if (importRawData.length === 0) {
                showLoader(false);
                Swal.fire('Warning', 'ไม่พบข้อมูลในไฟล์', 'warning');
                return;
            }

            // Open Modal & Render Preview (Default start at row 2)
            showLoader(false);
            document.getElementById('import-start-row').value = 2; // Reset default
            openImportModal();
            renderImportPreview();

            // Reset input so same file can be selected again if cancelled
            input.value = '';

        } catch (err) {
            showLoader(false);
            console.error(err);
            Swal.fire('Error', 'เกิดข้อผิดพลาดในการอ่านไฟล์: ' + err.message, 'error');
            input.value = '';
        }
    };
    reader.readAsArrayBuffer(file);
}

// --- Import Modal logic ---
function openImportModal() {
    // Populate import year dropdown dynamically based on availableYears
    const importYearSelect = document.getElementById('import-eval-year');
    if (importYearSelect) {
        importYearSelect.innerHTML = '';
        availableYears.forEach(yr => {
            const opt = document.createElement('option');
            opt.value = yr;
            opt.innerText = `ปีงบประมาณ ${yr}`;
            importYearSelect.appendChild(opt);
        });
        importYearSelect.value = selectedYear;
    }

    document.getElementById('importModal').classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
}

function closeImportModal() {
    document.getElementById('importModal').classList.add('hidden');
    document.body.classList.remove('overflow-hidden');
    importRawData = []; // Clear memory
}

function renderImportPreview() {
    const startRow = parseInt(document.getElementById('import-start-row').value) || 1;
    // Excel Row 1 = Array Index 0.
    // If user says Start Row 7, we skip 0..5 (6 rows) => slice(6).
    // So slice index = startRow - 1.

    const sliceIndex = Math.max(0, startRow - 1);
    const previewData = importRawData.slice(sliceIndex);

    document.getElementById('import-total-rows').innerText = previewData.length;

    const tbody = document.getElementById('import-preview-body');
    tbody.innerHTML = '';

    // Show first 5 rows for preview
    previewData.slice(0, 5).forEach((row, idx) => {
        const tr = document.createElement('tr');
        tr.className = 'border-b hover:bg-gray-50 dark:hover:bg-gray-700 dark:border-gray-600';

        // Helper safely get value
        const v = (i) => row[i] !== undefined ? row[i] : '';

        tr.innerHTML = `
            <td class="px-2 py-1 bg-gray-50 dark:bg-gray-700 text-gray-400 dark:text-gray-300 text-xs text-center">${startRow + idx}</td>
            <td class="px-2 py-1 dark:text-gray-300">${v(1)}</td>
            <td class="px-2 py-1 text-primary dark:text-primary-400 font-medium">${v(2)}</td>
            <td class="px-2 py-1 dark:text-gray-300">${v(3)}</td>
            <td class="px-2 py-1 dark:text-gray-300">${v(4)}</td>
        `;
        tbody.appendChild(tr);
    });

    if (previewData.length > 5) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="5" class="text-center py-2 text-gray-400 dark:text-gray-500 italic">... และอีก ${previewData.length - 5} รายการ ...</td>`;
        tbody.appendChild(tr);
    }
}

function confirmImport() {
    if (importRawData.length === 0) return;

    const startRow = parseInt(document.getElementById('import-start-row').value) || 1;
    const sliceIndex = Math.max(0, startRow - 1);
    const rowsComponents = importRawData.slice(sliceIndex);

    if (rowsComponents.length === 0) {
        Swal.fire('Warning', 'ไม่พบข้อมูลในแถวที่เลือก', 'warning');
        return;
    }

    const importYear = document.getElementById('import-eval-year').value || selectedYear;

    showLoader(true);
    closeImportModal();

    // Transform to Payload
    const payload = rowsComponents.map(r => {
        const coopType = r[4] || "";
        const isFarmerChecked = document.getElementById('chk-farmer-group').checked;
        const isFarmerGroup = isFarmerChecked || coopType.includes("กลุ่มเกษตรกร");

        const v = (targetIndex) => {
            if (!isFarmerGroup) return r[targetIndex];
            if (targetIndex < 20) return r[targetIndex];
            if (targetIndex === 20) return 0;
            return r[targetIndex - 1];
        };

        return {
            eval_year: importYear,
            agency: v(1),
            coop_name: v(2),
            coop_code: v(3),
            coop_type: coopType,
            d1_1: v(5), d1_2: v(6), d1_3: v(7),
            d2_1: v(9), d2_2: v(10), d2_3: v(11), d2_4: v(12), d2_5: v(13), d2_6: v(14),
            d3_1: v(16), d3_2: v(17), d3_3: v(18), d3_4: v(19), d3_5: v(20),
            d3_fail_text: v(21),
            d4_1: v(23), d4_2: v(24), d4_3: v(25), d4_4: v(26),
            grade: v(31), // Column AF - ผลการจัดชั้นปี 68
            remark: v(32),
            no_balance_years: v(33),
            eval_year_round: v(34),
            agm_68: v(35),
            finance_68: v(36),
            acc_year_68: v(37),
            acc_year_67: v(38),
            change_year_note: v(39),
            grade_67: v(40)
        };
    });

    google.script.run
        .withSuccessHandler(res => {
            showLoader(false);
            Swal.fire('Success', res.message, 'success');
            fetchData();
        })
        .withFailureHandler(err => {
            showLoader(false);
            Swal.fire('Error', 'Import failed: ' + err.message, 'error');
        })
        .importBulkData(payload);
}

// --- Helper for Excel Import ---
function handleExcelImport(input) {
    const file = input.files[0];
    if (!file) return;

    input.value = '';

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        importRawData = jsonData;
        document.getElementById('import-total-rows').innerText = jsonData.length;
        openImportModal();
        renderImportPreview();
    };
    reader.readAsArrayBuffer(file);
}

// --- Admin Helpers ---
let currentFarmerGroupConfig = null; // For กลุ่มเกษตรกร

function updateIndicatorField(key, field, val) {
    // Get current item to check if it's farmer group
    const id = parseInt(document.getElementById('detail-modal-id').dataset.id);
    const item = dataList.find(d => d.id === id);
    const isFarmerGroup = item && item.name && item.name.includes('กลุ่มเกษตรกร');

    if (isFarmerGroup) {
        if (!currentFarmerGroupConfig) currentFarmerGroupConfig = { ...FARMER_GROUP_INDICATOR_INFO };
        if (!currentFarmerGroupConfig[key]) currentFarmerGroupConfig[key] = { ...FARMER_GROUP_INDICATOR_INFO[key] };
        currentFarmerGroupConfig[key][field] = val;
        FARMER_GROUP_INDICATOR_INFO[key][field] = val;
    } else {
        if (!currentIndicatorConfig) currentIndicatorConfig = { ...INDICATOR_INFO };
        if (!currentIndicatorConfig[key]) currentIndicatorConfig[key] = { ...INDICATOR_INFO[key] };
        currentIndicatorConfig[key][field] = val;
        INDICATOR_INFO[key][field] = val;
    }
}

function saveIndicatorConfig() {
    // Get current item to check if it's farmer group
    const id = parseInt(document.getElementById('detail-modal-id').dataset.id);
    const item = dataList.find(d => d.id === id);
    const isFarmerGroup = item && item.name && item.name.includes('กลุ่มเกษตรกร');

    const configToSave = isFarmerGroup ? currentFarmerGroupConfig : currentIndicatorConfig;
    const configType = isFarmerGroup ? 'farmer_group' : 'coop';

    if (!configToSave) {
        Swal.fire('Info', 'ไม่มีการเปลี่ยนแปลง', 'info');
        return;
    }

    showLoader(true);
    google.script.run
        .withSuccessHandler(res => {
            showLoader(false);
            Swal.fire('Success', res.message, 'success');
        })
        .withFailureHandler(err => {
            showLoader(false);
            Swal.fire('Error', err.message, 'error');
        })
        .saveIndicatorConfig(configToSave, configType);
}

function updateLocalAdvice(key, val) {
    // Local Update
    const id = parseInt(document.getElementById('detail-modal-id').dataset.id);
    const item = dataList.find(d => d.id === id);
    if (item) {
        if (!item.advice) item.advice = {};
        item.advice[key] = val;
    }
}

function saveLocalAdvice() {
    const id = parseInt(document.getElementById('detail-modal-id').dataset.id);
    const item = dataList.find(d => d.id === id);
    if (!item) return;

    showLoader(true);
    google.script.run
        .withSuccessHandler(res => {
            showLoader(false);
            Swal.fire('Success', res.message, 'success');
        })
        .withFailureHandler(err => {
            showLoader(false);
            Swal.fire('Error', err.message, 'error');
        })
        .saveCoopAdvice(item.id, JSON.stringify(item.advice));
}

// ===========================================
// INDICATOR CRITERIA FUNCTIONS
// ===========================================

/**
 * Fetch Criteria Data from Backend
 * @param {string} type - 'coop' or 'farmer_group'
 */
function fetchCriteriaData(type, callback) {
    google.script.run
        .withSuccessHandler((res) => {
            console.log('Criteria Fetch Result (' + type + '):', res);
            if (type === 'farmer_group') {
                criteriaDataFG = res.criteria || {};
            } else {
                criteriaDataCoop = res.criteria || {};
            }
            if (callback) callback();
        })
        .withFailureHandler((err) => {
            console.error('Failed to fetch criteria (' + type + '):', err);
            // Still callback to proceed chain even if failed
            if (callback) callback();
        })
        .getIndicatorCriteria(type);
}

/**
 * View Indicator Criteria Modal
 * @param {string} code - Indicator code (e.g., '1.1')
 * @param {string} type - 'coop' or 'farmer_group'
 */
function viewIndicatorCriteria(code, type) {
    currentCriteriaCode = code;
    currentCriteriaType = type || 'coop';

    const typeLabel = currentCriteriaType === 'farmer_group' ? '(กลุ่มเกษตรกร)' : '(สหกรณ์)';

    // Update modal header
    document.getElementById('criteria-indicator-code').innerText = code;
    document.getElementById('criteria-indicator-type').innerText = typeLabel;

    // Get criteria content
    const criteriaData = currentCriteriaType === 'farmer_group' ? criteriaDataFG : criteriaDataCoop;

    // Check if criteria data is fully loaded
    const isLoaded = Object.keys(criteriaData).length > 0;
    const html = criteriaData[code] || '';

    console.log('Viewing Criteria:', { code: code, type: type, found: !!html, isLoaded: isLoaded });

    const contentDiv = document.getElementById('criteria-content');

    if (isLoaded) {
        if (html && html.trim()) {
            contentDiv.innerHTML = html;
            // Trigger MathJax to render formulas
            if (window.MathJax && MathJax.typesetPromise) {
                MathJax.typesetPromise([contentDiv]).catch((err) => console.error('MathJax error:', err));
            }
        } else {
            contentDiv.innerHTML = '<p class="text-gray-500 dark:text-gray-400 italic text-center py-8">ยังไม่มีข้อมูลเกณฑ์สำหรับตัวชี้วัดนี้</p>';
        }
    } else {
        // Show Loading State
        contentDiv.innerHTML = `
            <div class="flex flex-col items-center justify-center py-8 text-blue-600 dark:text-blue-400">
                <i class="fas fa-spinner fa-spin text-2xl mb-2"></i>
                <p>กำลังโหลดข้อมูลเกณฑ์ล่าสุด...</p>
                <p class="text-sm text-gray-400 mt-1">กรุณารอสักครู่แล้วเปิดใหม่อีกครั้ง</p>
            </div>
        `;
    }

    // Show/hide edit button based on admin status
    const editBtn = document.getElementById('criteria-edit-btn');
    if (editBtn) {
        editBtn.classList.toggle('hidden', !isAdmin);
    }

    // Show modal
    document.getElementById('criteriaModal').classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
}

/**
 * Close Criteria Modal
 */
function closeCriteriaModal() {
    document.getElementById('criteriaModal').classList.add('hidden');
    document.body.classList.remove('overflow-hidden');
}

/**
 * Open Criteria Editor (Admin Only)
 */
function openCriteriaEditor() {
    if (!isAdmin) {
        Swal.fire('Error', 'คุณไม่มีสิทธิ์แก้ไข', 'error');
        return;
    }

    const code = currentCriteriaCode;
    const type = currentCriteriaType;
    const typeLabel = type === 'farmer_group' ? '(กลุ่มเกษตรกร)' : '(สหกรณ์)';

    // Update editor header
    document.getElementById('editor-indicator-code').innerText = code;
    document.getElementById('editor-indicator-type').innerText = typeLabel;

    // Get current content
    const criteriaData = type === 'farmer_group' ? criteriaDataFG : criteriaDataCoop;
    const html = criteriaData[code] || '';

    // Initialize Quill if NOT exists
    if (!quillEditor) {
        quillEditor = new Quill('#criteria-editor-container', {
            theme: 'snow',
            modules: {
                toolbar: {
                    container: [
                        [{ 'header': [1, 2, 3, false] }],
                        ['bold', 'italic', 'underline', 'strike'],
                        [{ 'color': [] }, { 'background': [] }],
                        [{ 'list': 'ordered' }, { 'list': 'bullet' }],
                        [{ 'align': [] }],
                        ['link', 'image'],
                        ['clean']
                    ],
                    handlers: {
                        image: selectLocalImage
                    }
                }
            },
            placeholder: 'พิมพ์รายละเอียดเกณฑ์ที่นี่...\n(รูปรองรับการอัพโหลดไม่เกิน 5MB)'
        });
    }

    /**
     * Handler for Quill Image Upload
     */
    function selectLocalImage() {
        const input = document.createElement('input');
        input.setAttribute('type', 'file');
        input.setAttribute('accept', 'image/*');
        input.click();

        input.onchange = () => {
            const file = input.files[0];
            if (file) {
                // Limit size 5MB
                if (file.size > 5 * 1024 * 1024) {
                    Swal.fire('Error', 'ไฟล์รูปภาพต้องมีขนาดไม่เกิน 5MB', 'error');
                    return;
                }

                // Show loading placeholder at cursor
                const range = quillEditor.getSelection();
                const loadingPlaceholder = ' กําลังอัพโหลดรูปภาพ... ';
                quillEditor.insertText(range.index, loadingPlaceholder, { 'color': '#0288D1', 'italic': true });

                const reader = new FileReader();
                reader.onload = (e) => {
                    const base64Data = e.target.result.split(',')[1];

                    google.script.run
                        .withSuccessHandler((res) => {
                            // Remove placeholder
                            quillEditor.deleteText(range.index, loadingPlaceholder.length);

                            if (res.success) {
                                quillEditor.insertEmbed(range.index, 'image', res.url);
                            } else {
                                Swal.fire('Error', 'อัพโหลดรูปภาพไม่สำเร็จ: ' + res.error, 'error');
                            }
                        })
                        .withFailureHandler((err) => {
                            quillEditor.deleteText(range.index, loadingPlaceholder.length);
                            Swal.fire('Error', 'เกิดข้อผิดพลาดในการอัพโหลด: ' + err, 'error');
                        })
                        .uploadImage(base64Data, file.type, file.name);
                };
                reader.readAsDataURL(file);
            }
        };
    }

    // Set content (Always overwrite with current data)
    quillEditor.root.innerHTML = html || '';

    // Show editor modal
    document.getElementById('criteriaEditorModal').classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
}

/**
 * Close Criteria Editor
 */
function closeCriteriaEditor() {
    // Just hide the modal, keep the editor instance alive
    document.getElementById('criteriaEditorModal').classList.add('hidden');
    document.body.classList.remove('overflow-hidden');
}

/**
 * Save Criteria Content (Admin Only)
 */
function saveCriteriaContent() {
    if (!isAdmin) {
        Swal.fire('Error', 'คุณไม่มีสิทธิ์บันทึก', 'error');
        return;
    }

    const code = currentCriteriaCode;
    const type = currentCriteriaType;

    // Get content from Quill
    let html = '';
    if (quillEditor) {
        html = quillEditor.root.innerHTML;
    }

    showLoader(true);

    google.script.run
        .withSuccessHandler((res) => {
            showLoader(false);

            // Update local cache
            if (type === 'farmer_group') {
                criteriaDataFG[code] = html;
            } else {
                criteriaDataCoop[code] = html;
            }

            closeCriteriaEditor();

            // Update the view modal content
            viewIndicatorCriteria(code, type);

            Swal.fire('สำเร็จ', res.message, 'success');
        })
        .withFailureHandler((err) => {
            showLoader(false);
            Swal.fire('Error', 'บันทึกล้มเหลว: ' + err.message, 'error');
        })
        .saveIndicatorCriteria(code, html, type);
}

/**
 * Load all criteria data on app init
 */
// Load criteria sequentially to prevent Apps Script concurrent limit issues
function loadAllCriteriaData() {
    console.log('Starting criteria load sequence...');
    fetchCriteriaData('coop', () => {
        console.log('Coop criteria loaded. Fetching Farmer Group...');
        fetchCriteriaData('farmer_group', () => {
            console.log('All criteria data loaded.');
        });
    });
}
