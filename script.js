const API_URL = 'https://script.google.com/macros/s/AKfycbwpanLa78SMKtxUcD26pUdjsZ8sdBQx6T5HZAEs58NVNnvw7zMG_UvioVBCQSwIym3e/exec';

// Global variables to store personnel data
let personnelData = [];
let personnelMap = {};
let savingsData = [];
let monthlyChartInstance = null;
let personChartInstance = null;

// Format currency for Thai Baht
function formatCurrency(amount, showUnit = true) {
    if (isNaN(amount) || amount === null) return '0' + (showUnit ? ' บาท' : '');
    return amount.toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + (showUnit ? ' บาท' : '');
}

// Tab navigation
document.querySelectorAll('.nav-tab').forEach(tab => {
    tab.addEventListener('click', function() {
        document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        this.classList.add('active');
        const tabId = this.getAttribute('data-tab');
        document.getElementById(tabId).classList.add('active');
        if (tabId === 'reports-tab') {
            initializeDashboard();
        }
    });
});

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    initializePersonnelSheet();
    initializeSavingsSheet();
    loadPersonnelData();
    loadPersonnelSequenceNumber();
    loadSavingsSequenceNumber();
    loadSavingsData();
    populateYearDropdowns();

    document.getElementById('person-selector').addEventListener('change', function() {
        const personNo = this.value;
        const personNameDisplay = document.getElementById('person-name-display');
        if (personNo && personnelMap[personNo]) {
            const person = personnelMap[personNo];
            personNameDisplay.textContent = `${person.title || ''} ${person.name || ''}`;
        } else {
            personNameDisplay.textContent = '';
        }
    });

    document.querySelector('input[name="personnel-birthday"]').addEventListener('blur', validateDateFormat);
    document.querySelector('input[name="personnel-workday"]').addEventListener('blur', validateDateFormat);

    document.getElementById('personnel-title-select').addEventListener('change', function() {
        const customTitleContainer = document.getElementById('custom-title-container');
        const customTitleInput = document.getElementById('personnel-title-custom');
        if (this.value === 'custom') {
            customTitleContainer.style.display = 'block';
            customTitleInput.required = true;
            customTitleInput.focus();
        } else {
            customTitleContainer.style.display = 'none';
            customTitleInput.required = false;
        }
        updateTitleField();
    });

    document.getElementById('personnel-title-custom').addEventListener('input', updateTitleField);
    document.getElementById('report-year-filter').addEventListener('change', filterReportData);
    document.getElementById('report-month-filter').addEventListener('change', filterReportData);
    document.getElementById('report-person-filter').addEventListener('change', filterReportData);
    document.getElementById('refresh-dashboard-btn').addEventListener('click', function() {
        Swal.fire({
            title: 'ยืนยันการรีเฟรชข้อมูล',
            text: 'การดำเนินการนี้จะโหลดข้อมูลใหม่ทั้งหมด ต้องการดำเนินการต่อหรือไม่?',
            icon: 'question',
            showCancelButton: true,
            confirmButtonText: 'ดำเนินการต่อ',
            cancelButtonText: 'ยกเลิก',
            confirmButtonColor: '#9c6ade',
            cancelButtonColor: '#6b7280'
        }).then((result) => {
            if (result.isConfirmed) {
                loadAllData();
            }
        });
    });
    document.getElementById('export-pdf-btn').addEventListener('click', exportToPDF);

    // Personnel form submission
    document.getElementById('personnel-form').addEventListener('submit', function(e) {
        e.preventDefault();

        const nameField = document.querySelector('input[name="personnel-name"]');
        if (!nameField.value.trim()) {
            Swal.fire({
                title: 'กรุณากรอกชื่อ-นามสกุล',
                icon: 'warning',
                confirmButtonText: 'ตกลง',
                confirmButtonColor: '#9c6ade'
            });
            nameField.focus();
            return;
        }

        updateTitleField();
        const birthdayField = document.querySelector('input[name="personnel-birthday"]');
        const workdayField = document.querySelector('input[name="personnel-workday"]');

        if (birthdayField.value && !isValidDateFormat(birthdayField.value)) {
            Swal.fire({
                title: 'รูปแบบวันที่ไม่ถูกต้อง',
                text: 'กรุณากรอกวันเกิดในรูปแบบ วัน/เดือน/ปี พ.ศ. เช่น 1/1/2539',
                icon: 'warning',
                confirmButtonText: 'ตกลง',
                confirmButtonColor: '#9c6ade'
            });
            birthdayField.focus();
            return;
        }

        if (workdayField.value && !isValidDateFormat(workdayField.value)) {
            Swal.fire({
                title: 'รูปแบบวันที่ไม่ถูกต้อง',
                text: 'กรุณากรอกวันที่เริ่มทำงานในรูปแบบ วัน/เดือน/ปี พ.ศ. เช่น 1/1/2539',
                icon: 'warning',
                confirmButtonText: 'ตกลง',
                confirmButtonColor: '#9c6ade'
            });
            workdayField.focus();
            return;
        }

        const titleField = document.getElementById('personnel-title');
        if (!titleField.value) {
            Swal.fire({
                title: 'กรุณาระบุคำนำหน้า',
                text: 'โปรดเลือกหรือระบุคำนำหน้า',
                icon: 'warning',
                confirmButtonText: 'ตกลง',
                confirmButtonColor: '#9c6ade'
            });
            document.getElementById('personnel-title-select').focus();
            return;
        }

        document.getElementById('personnel-loading').classList.remove('hidden');
        const formData = new FormData(this);
        const data = {
            no: document.getElementById('auto-personnel-no').value,
            title: document.getElementById('personnel-title').value,
            name: formData.get('personnel-name'),
            birthday: formData.get('personnel-birthday'),
            workday: formData.get('personnel-workday'),
            phone: formData.get('personnel-phone'),
            address: formData.get('personnel-address'),
            idline: formData.get('personnel-idline'),
            fb: formData.get('personnel-fb'),
            remake: formData.get('personnel-remake')
        };

        fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify(data),
            headers: {
                'Content-Type': 'application/json'
            }
        })
        .then(response => response.json())
        .then(result => {
            document.getElementById('personnel-loading').classList.add('hidden');
            if (result.status === 'success') {
                Swal.fire({
                    title: 'บันทึกสำเร็จ!',
                    text: 'ข้อมูลบุคลากรถูกบันทึกเรียบร้อยแล้ว',
                    icon: 'success',
                    confirmButtonText: 'ตกลง',
                    confirmButtonColor: '#9c6ade'
                });
                document.getElementById('personnel-form').reset();
                document.getElementById('custom-title-container').style.display = 'none';
                loadPersonnelData();
                loadPersonnelSequenceNumber();
            } else {
                throw new Error('Failed to save data');
            }
        })
        .catch(error => {
            document.getElementById('personnel-loading').classList.add('hidden');
            Swal.fire({
                title: 'เกิดข้อผิดพลาด!',
                text: 'ไม่สามารถบันทึกข้อมูลได้ กรุณาลองใหม่อีกครั้ง',
                icon: 'error',
                confirmButtonText: 'ตกลง',
                confirmButtonColor: '#9c6ade'
            });
            console.error('Error:', error);
        });
    });

    // Savings form submission
    document.getElementById('savings-form').addEventListener('submit', function(e) {
        e.preventDefault();
        document.getElementById('savings-loading').classList.remove('hidden');
        const formData = new FormData(this);
        const data = {
            no: document.getElementById('auto-savings-no').value,
            person_no: formData.get('person-no'),
            person_name: personnelMap[formData.get('person-no')]?.name || '',
            month: formData.get('savings-month'),
            year: formData.get('savings-year'),
            amount: formData.get('savings-amount'),
            remark: formData.get('savings-remark')
        };

        fetch(`${API_URL}?sheet=savings`, {
            method: 'POST',
            body: JSON.stringify(data),
            headers: {
                'Content-Type': 'application/json'
            }
        })
        .then(response => response.json())
        .then(result => {
            document.getElementById('savings-loading').classList.add('hidden');
            if (result.status === 'success') {
                Swal.fire({
                    title: 'บันทึกสำเร็จ!',
                    text: 'ข้อมูลการออมถูกบันทึกเรียบร้อยแล้ว',
                    icon: 'success',
                    confirmButtonText: 'ตกลง',
                    confirmButtonColor: '#9c6ade'
                });
                document.getElementById('savings-form').reset();
                document.getElementById('person-name-display').textContent = '';
                loadSavingsData();
                loadSavingsSequenceNumber();
            } else {
                throw new Error('Failed to save data');
            }
        })
        .catch(error => {
            document.getElementById('savings-loading').classList.add('hidden');
            Swal.fire({
                title: 'เกิดข้อผิดพลาด!',
                text: 'ไม่สามารถบันทึกข้อมูลได้ กรุณาลองใหม่อีกครั้ง',
                icon: 'error',
                confirmButtonText: 'ตกลง',
                confirmButtonColor: '#9c6ade'
            });
            console.error('Error:', error);
        });
    });
});

// Load all data for dashboard
function loadAllData() {
    document.getElementById('total-personnel').textContent = '...';
    document.getElementById('total-savings').textContent = '...';
    document.getElementById('avg-savings-per-person').textContent = '...';
    document.getElementById('total-savings-records').textContent = '...';
    loadPersonnelData();
    loadSavingsData();
}

// Initialize dashboard
function initializeDashboard() {
    if (savingsData.length === 0) {
        loadAllData();
    } else {
        updateDashboard(savingsData);
        updateReportData(savingsData);
    }
}

// Update dashboard with data
function updateDashboard(data) {
    const totalPersonnel = Object.keys(personnelMap).length;
    document.getElementById('total-personnel').textContent = totalPersonnel;
    let totalSavings = 0;
    data.forEach(item => {
        totalSavings += parseFloat(item.amount) || 0;
    });
    document.getElementById('total-savings').textContent = formatCurrency(totalSavings, false);
    const avgSavingsPerPerson = totalPersonnel > 0 ? totalSavings / totalPersonnel : 0;
    document.getElementById('avg-savings-per-person').textContent = formatCurrency(avgSavingsPerPerson, false);
    document.getElementById('total-savings-records').textContent = data.length;
    updateMonthlyChart(data);
    updatePersonChart(data);
}

// Update monthly savings chart
function updateMonthlyChart(data) {
    const thaiMonths = [
        'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
        'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];
    const currentYear = new Date().getFullYear() + 543;
    const monthlyTotals = {};
    thaiMonths.forEach(month => monthlyTotals[month] = 0);

    if (!data || !Array.isArray(data)) {
        console.warn('No valid data for monthly chart');
        return;
    }

    data.forEach(item => {
        if (parseInt(item.year) === currentYear && thaiMonths.includes(item.month)) {
            monthlyTotals[item.month] += parseFloat(item.amount) || 0;
        }
    });

    const chartLabels = thaiMonths;
    const chartData = thaiMonths.map(month => monthlyTotals[month]);
    const ctx = document.getElementById('monthly-savings-chart').getContext('2d');
    
    if (monthlyChartInstance) {
        monthlyChartInstance.destroy();
    }

    monthlyChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: chartLabels,
            datasets: [{
                label: `ยอดเงินออมรายเดือน ปี ${currentYear}`,
                data: chartData,
                backgroundColor: 'rgba(156, 106, 222, 0.6)',
                borderColor: 'rgba(156, 106, 222, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value.toLocaleString('th-TH') + ' บาท';
                        }
                    }
                }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + context.raw.toLocaleString('th-TH') + ' บาท';
                        }
                    }
                }
            }
        }
    });
}

// Update person savings chart
function updatePersonChart(data) {
    const personTotals = {};
    data.forEach(item => {
        const personNo = item.person_no;
        const personName = item.person_name;
        if (personNo && personName) {
            const key = `${personNo} - ${personName}`;
            if (!personTotals[key]) {
                personTotals[key] = 0;
            }
            personTotals[key] += parseFloat(item.amount) || 0;
        }
    });

    const sortedPersons = Object.entries(personTotals)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    const chartLabels = sortedPersons.map(item => {
        const name = item[0].split(' - ')[1];
        return name.length > 15 ? name.substring(0, 15) + '...' : name;
    });
    const chartData = sortedPersons.map(item => item[1]);
    const ctx = document.getElementById('person-savings-chart').getContext('2d');
    
    if (personChartInstance) {
        personChartInstance.destroy();
    }

    personChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: chartLabels,
            datasets: [{
                label: 'ยอดเงินออมรวม',
                data: chartData,
                backgroundColor: 'rgba(79, 209, 197, 0.6)',
                borderColor: 'rgba(79, 209, 197, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        callback: function(value) {
                            return value.toLocaleString('th-TH') + ' บาท';
                        }
                    }
                }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + context.raw.toLocaleString('th-TH') + ' บาท';
                        }
                    }
                }
            }
        }
    });
}

// Update report data
function updateReportData(data) {
    const yearFilter = document.getElementById('report-year-filter');
    const years = [...new Set(data.map(item => item.year))].filter(Boolean).sort((a, b) => b - a);
    
    while (yearFilter.options.length > 1) {
        yearFilter.remove(1);
    }
    
    years.forEach(year => {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        yearFilter.appendChild(option);
    });

    const personFilter = document.getElementById('report-person-filter');
    while (personFilter.options.length > 1) {
        personFilter.remove(1);
    }

    const uniquePersons = {};
    data.forEach(item => {
        if (item.person_no && item.person_name) {
            uniquePersons[item.person_no] = item.person_name;
        }
    });

    Object.entries(uniquePersons).forEach(([personNo, personName]) => {
        const option = document.createElement('option');
        option.value = personNo;
        option.textContent = `${personNo} - ${personName}`;
        personFilter.appendChild(option);
    });

    filterReportData();
}

// Filter report data based on selected filters
function filterReportData() {
    const yearFilter = document.getElementById('report-year-filter').value;
    const monthFilter = document.getElementById('report-month-filter').value;
    const personFilter = document.getElementById('report-person-filter').value;
    
    let filteredData = [...savingsData];
    if (yearFilter !== 'all') {
        filteredData = filteredData.filter(item => item.year === yearFilter);
    }
    if (monthFilter !== 'all') {
        filteredData = filteredData.filter(item => item.month === monthFilter);
    }
    if (personFilter !== 'all') {
        filteredData = filteredData.filter(item => item.person_no === personFilter);
    }
    
    filteredData.sort((a, b) => parseInt(b.no) - parseInt(a.no));
    updateReportTable(filteredData);
    updateReportSummary(filteredData);
}

// Update report table
function updateReportTable(data) {
    const tableBody = document.getElementById('report-table-body');
    tableBody.innerHTML = '';

    if (!data || !Array.isArray(data) || data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="6" class="text-center py-4">ไม่พบข้อมูล</td></tr>';
        return;
    }

    data.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.no || '-'}</td>
            <td>${row.person_name || '-'}</td>
            <td>${row.month || '-'}</td>
            <td>${row.year || '-'}</td>
            <td>${formatCurrency(row.amount)}</td>
            <td>${row.remark || '-'}</td>
        `;
        tableBody.appendChild(tr);
    });
}

// Update report summary
function updateReportSummary(data) {
    const totalRecords = data.length;
    document.getElementById('report-total-records').textContent = totalRecords;
    let totalAmount = 0;
    data.forEach(item => {
        totalAmount += parseFloat(item.amount) || 0;
    });
    document.getElementById('report-total-amount').textContent = formatCurrency(totalAmount);
    const avgAmount = totalRecords > 0 ? totalAmount / totalRecords : 0;
    document.getElementById('report-avg-amount').textContent = formatCurrency(avgAmount);
}

// Export report to PDF
function exportToPDF() {
    Swal.fire({
        title: 'กำลังสร้าง PDF',
        text: 'กรุณารอสักครู่...',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    const now = new Date();
    const dateStr = `วันที่ ${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear() + 543}`;
    document.getElementById('pdf-date').textContent = dateStr;

    document.getElementById('pdf-total-records').textContent = document.getElementById('report-total-records').textContent;
    document.getElementById('pdf-total-amount').textContent = document.getElementById('report-total-amount').textContent;
    document.getElementById('pdf-avg-amount').textContent = document.getElementById('report-avg-amount').textContent;

    const reportTableBody = document.getElementById('report-table-body');
    const pdfTableBody = document.getElementById('pdf-table-body');
    pdfTableBody.innerHTML = reportTableBody.innerHTML;

    const yearFilter = document.getElementById('report-year-filter');
    const monthFilter = document.getElementById('report-month-filter');
    const personFilter = document.getElementById('report-person-filter');

    const yearText = yearFilter.value !== 'all' ? yearFilter.options[yearFilter.selectedIndex].text : 'ทุกปี';
    const monthText = monthFilter.value !== 'all' ? monthFilter.options[monthFilter.selectedIndex].text : 'ทุกเดือน';
    const personText = personFilter.value !== 'all' ? personFilter.options[personFilter.selectedIndex].text.split(' - ')[1] : 'ทุกคน';
    const filename = `รายงานการออม_${yearText}_${monthText}_${personText}.pdf`;

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');
    const margin = 10;
    const pageWidth = 210 - 2 * margin;
    const pageHeight = 297 - 2 * margin;

    html2canvas(document.getElementById('pdf-template'), {
        scale: 2,
        useCORS: true,
        logging: false,
        windowWidth: 1200
    }).then(canvas => {
        const imgData = canvas.toDataURL('image/png');
        const imgHeight = canvas.height * pageWidth / canvas.width;
        let heightLeft = imgHeight;
        let position = margin;

        doc.addImage(imgData, 'PNG', margin, position, pageWidth, imgHeight);
        heightLeft -= pageHeight;

        while (heightLeft > 0) {
            doc.addPage();
            position = margin - (pageHeight * Math.floor(heightLeft / pageHeight));
            doc.addImage(imgData, 'PNG', margin, position, pageWidth, imgHeight);
            heightLeft -= pageHeight;
        }

        doc.save(filename);
        Swal.fire({
            title: 'สร้าง PDF สำเร็จ!',
            text: 'ไฟล์ PDF ถูกดาวน์โหลดแล้ว',
            icon: 'success',
            confirmButtonText: 'ตกลง',
            confirmButtonColor: '#9c6ade'
        });
    }).catch(error => {
        console.error('Error generating PDF:', error);
        Swal.fire({
            title: 'เกิดข้อผิดพลาด!',
            text: 'ไม่สามารถสร้างไฟล์ PDF ได้',
            icon: 'error',
            confirmButtonText: 'ตกลง',
            confirmButtonColor: '#9c6ade'
        });
    });
}

// Populate year dropdowns
function populateYearDropdowns() {
    const currentYear = new Date().getFullYear() + 543;
    const yearSelect = document.querySelector('select[name="savings-year"]');
    
    for (let i = 0; i < 6; i++) {
        const year = currentYear - i;
        const yearOption = document.createElement('option');
        yearOption.value = year;
        yearOption.textContent = year;
        yearSelect.appendChild(yearOption);
    }
}

// Update the hidden title field
function updateTitleField() {
    const titleSelect = document.getElementById('personnel-title-select');
    const customTitleInput = document.getElementById('personnel-title-custom');
    const titleField = document.getElementById('personnel-title');
    
    if (titleSelect.value === 'custom') {
        titleField.value = customTitleInput.value;
    } else {
        titleField.value = titleSelect.value;
    }
}

// Validate date format (dd/mm/yyyy)
function validateDateFormat(e) {
    const input = e.target;
    const value = input.value.trim();
    
    if (!value) return;
    
    const dateRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    const match = value.match(dateRegex);
    
    if (!match) {
        Swal.fire({
            title: 'รูปแบบวันที่ไม่ถูกต้อง',
            text: 'กรุณากรอกวันที่ในรูปแบบ วัน/เดือน/ปี พ.ศ. เช่น 1/1/2539',
            icon: 'warning',
            confirmButtonText: 'ตกลง',
            confirmButtonColor: '#9c6ade'
        });
        input.focus();
        return;
    }
    
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10);
    const year = parseInt(match[3], 10);
    
    if (month < 1 || month > 12 || day < 1 || year < 2400 || year > 2600) {
        Swal.fire({
            title: 'วันที่ไม่ถูกต้อง',
            text: 'กรุณาตรวจสอบวัน เดือน และปี พ.ศ. ให้ถูกต้อง',
            icon: 'warning',
            confirmButtonText: 'ตกลง',
            confirmButtonColor: '#9c6ade'
        });
        input.focus();
        return;
    }

    const daysInMonth = new Date(year - 543, month, 0).getDate();
    if (day > daysInMonth) {
        Swal.fire({
            title: 'วันที่ไม่ถูกต้อง',
            text: 'จำนวนวันเกินกว่าที่เดือนนี้มี',
            icon: 'warning',
            confirmButtonText: 'ตกลง',
            confirmButtonColor: '#9c6ade'
        });
        input.focus();
    }
}

// Check if date format is valid
function isValidDateFormat(dateString) {
    if (!dateString) return true;
    
    const dateRegex = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    const match = dateString.match(dateRegex);
    
    if (!match) return false;
    
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10);
    const year = parseInt(match[3], 10);
    
    if (month < 1 || month > 12 || day < 1 || year < 2400 || year > 2600) {
        return false;
    }
    
    const daysInMonth = new Date(year - 543, month, 0).getDate();
    if (day > daysInMonth) {
        return false;
    }
    
    return true;
}

// Initialize personnel sheet
function initializePersonnelSheet() {
    fetch(`${API_URL}?action=initialize&sheet=personnel`)
        .then(response => response.json())
        .then(data => {
            console.log('Personnel sheet initialization:', data);
        })
        .catch(error => {
            console.error('Error initializing personnel sheet:', error);
        });
}

// Initialize savings sheet
function initializeSavingsSheet() {
    fetch(`${API_URL}?action=initialize&sheet=savings`)
        .then(response => response.json())
        .then(data => {
            console.log('Savings sheet initialization:', data);
        })
        .catch(error => {
            console.error('Error initializing savings sheet:', error);
        });
}

// Load personnel data
function loadPersonnelData() {
    const cachedData = localStorage.getItem('personnelData');
    if (cachedData) {
        personnelData = JSON.parse(cachedData);
        updatePersonnelTable(personnelData);
        populatePersonnelDropdown(personnelData);
    }
    fetch(API_URL)
        .then(response => {
            if (!response.ok) throw new Error('Network response was not ok');
            return response.json();
        })
        .then(data => {
            if (!Array.isArray(data)) throw new Error('Invalid data format');
            personnelData = data;
            localStorage.setItem('personnelData', JSON.stringify(data));
            personnelMap = {};
            data.forEach(person => {
                if (person.no) {
                    personnelMap[person.no] = person;
                }
            });
            updatePersonnelTable(data);
            populatePersonnelDropdown(data);
        })
        .catch(error => {
            console.error('Error loading personnel data:', error);
            document.getElementById('personnel-body').innerHTML = '<tr><td colspan="10" class="text-center py-4 text-red-500">เกิดข้อผิดพลาดในการโหลดข้อมูล</td></tr>';
            Swal.fire({
                title: 'เกิดข้อผิดพลาด!',
                text: 'ไม่สามารถโหลดข้อมูลบุคลากรได้',
                icon: 'error',
                confirmButtonText: 'ตกลง',
                confirmButtonColor: '#9c6ade'
            });
        });
}

// Update personnel table
function updatePersonnelTable(data) {
    const tableBody = document.getElementById('personnel-body');
    tableBody.innerHTML = '';

    if (!data || !Array.isArray(data) || data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="10" class="text-center py-4">ไม่พบข้อมูล</td></tr>';
        return;
    }

    data.sort((a, b) => parseInt(b.no) - parseInt(a.no)).forEach(row => {
        const tr = document.createElement('tr');
        tr.className = 'border-b border-purple-100 hover:bg-purple-50';
        tr.innerHTML = `
            <td class="py-2 px-4">${row.no || '-'}</td>
            <td class="py-2 px-4">${row.title || '-'}</td>
            <td class="py-2 px-4">${row.name || '-'}</td>
            <td class="py-2 px-4">${row.birthday || '-'}</td>
            <td class="py-2 px-4">${row.workday || '-'}</td>
            <td class="py-2 px-4">${row.phone || '-'}</td>
            <td class="py-2 px-4">${row.address || '-'}</td>
            <td class="py-2 px-4">${row.idline || '-'}</td>
            <td class="py-2 px-4">${row.fb || '-'}</td>
            <td class="py-2 px-4">${row.remake || '-'}</td>
        `;
        tableBody.appendChild(tr);
    });
}

// Load savings data
function loadSavingsData() {
    const cachedData = localStorage.getItem('savingsData');
    if (cachedData) {
        savingsData = JSON.parse(cachedData);
        updateSavingsTable(savingsData);
        updateReportData(savingsData);
    }
    fetch(`${API_URL}?sheet=savings`)
        .then(response => {
            if (!response.ok) throw new Error('Network response was not ok');
            return response.json();
        })
        .then(data => {
            if (!Array.isArray(data)) throw new Error('Invalid data format');
            savingsData = data;
            localStorage.setItem('savingsData', JSON.stringify(data));
            updateSavingsTable(data);
            updateReportData(data);
        })
        .catch(error => {
            console.error('Error loading savings data:', error);
            document.getElementById('savings-body').innerHTML = '<tr><td colspan="6" class="text-center py-4 text-red-500">เกิดข้อผิดพลาดในการโหลดข้อมูล</td></tr>';
            Swal.fire({
                title: 'เกิดข้อผิดพลาด!',
                text: 'ไม่สามารถโหลดข้อมูลการออมได้',
                icon: 'error',
                confirmButtonText: 'ตกลง',
                confirmButtonColor: '#9c6ade'
            });
        });
}

// Update savings table
function updateSavingsTable(data) {
    const tableBody = document.getElementById('savings-body');
    tableBody.innerHTML = '';

    if (!data || !Array.isArray(data) || data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="6" class="text-center py-4">ไม่พบข้อมูล</td></tr>';
        return;
    }

    data.sort((a, b) => parseInt(b.no) - parseInt(a.no)).forEach(row => {
        const tr = document.createElement('tr');
        tr.className = 'border-b border-purple-100 hover:bg-purple-50';
        tr.innerHTML = `
            <td class="py-2 px-4">${row.no || '-'}</td>
            <td class="py-2 px-4">${row.person_name || '-'}</td>
            <td class="py-2 px-4">${row.month || '-'}</td>
            <td class="py-2 px-4">${row.year || '-'}</td>
            <td class="py-2 px-4">${formatCurrency(row.amount)}</td>
            <td class="py-2 px-4">${row.remark || '-'}</td>
        `;
        tableBody.appendChild(tr);
    });
}

// Load personnel sequence number
function loadPersonnelSequenceNumber() {
    fetch(API_URL)
        .then(response => {
            if (!response.ok) throw new Error('Network response was not ok');
            return response.json();
        })
        .then(data => {
            const nextNo = getNextSequenceNumber(data);
            document.querySelector('input[name="personnel-no"]').value = nextNo;
            document.getElementById('auto-personnel-no').value = nextNo;
        })
        .catch(error => {
            console.error('Error loading personnel sequence number:', error);
            document.querySelector('input[name="personnel-no"]').value = 1;
            document.getElementById('auto-personnel-no').value = 1;
        });
}

// Load savings sequence number
function loadSavingsSequenceNumber() {
    fetch(`${API_URL}?sheet=savings`)
        .then(response => {
            if (!response.ok) throw new Error('Network response was not ok');
            return response.json();
        })
        .then(data => {
            const nextNo = getNextSequenceNumber(data);
            document.querySelector('input[name="savings-no"]').value = nextNo;
            document.getElementById('auto-savings-no').value = nextNo;
        })
        .catch(error => {
            console.error('Error loading savings sequence number:', error);
            document.querySelector('input[name="savings-no"]').value = 1;
            document.getElementById('auto-savings-no').value = 1;
        });
}

// Get next sequence number
function getNextSequenceNumber(data) {
    if (!data || data.length === 0) {
        return 1;
    }
    
    let maxNo = 0;
    data.forEach(row => {
        const no = parseInt(row.no);
        if (!isNaN(no) && no > maxNo) {
            maxNo = no;
        }
    });
    
    return maxNo + 1;
}

// Populate personnel dropdown
function populatePersonnelDropdown(data) {
    const personSelector = document.getElementById('person-selector');
    while (personSelector.options.length > 1) {
        personSelector.remove(1);
    }
    
    if (!data || !Array.isArray(data)) return;
    
    data.forEach(person => {
        if (person.no && person.name) {
            const option = document.createElement('option');
            option.value = person.no;
            option.textContent = `${person.no} - ${person.title || ''} ${person.name}`;
            personSelector.appendChild(option);
        }
    });
}
