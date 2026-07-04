// =========================================
// 1. ANTI-DEVTOOLS / PENGHALANG INSPECT
// =========================================
document.addEventListener('contextmenu', function(e) {
    e.preventDefault(); // Halang Right-Click
});

document.addEventListener('keydown', function(e) {
    // Halang F12
    if (e.key === 'F12' || e.keyCode === 123) {
        e.preventDefault();
    }
    // Halang Ctrl + Shift + I (Inspect)
    if (e.ctrlKey && e.shiftKey && (e.key === 'I' || e.key === 'i')) {
        e.preventDefault();
    }
    // Halang Ctrl + Shift + C (Inspect Element)
    if (e.ctrlKey && e.shiftKey && (e.key === 'C' || e.key === 'c')) {
        e.preventDefault();
    }
    // Halang Ctrl + Shift + J (Console)
    if (e.ctrlKey && e.shiftKey && (e.key === 'J' || e.key === 'j')) {
        e.preventDefault();
    }
    // Halang Ctrl + U (View Source)
    if (e.ctrlKey && (e.key === 'U' || e.key === 'u')) {
        e.preventDefault();
    }
});

// =========================================
// 2. SISTEM UTAMA (MAIN LOGIC)
// =========================================
document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById('reportForm');
    const eng1Select = document.getElementById('eng1');
    const eng2Select = document.getElementById('eng2');
    const nextEng1Select = document.getElementById('nextEng1');
    const nextEng2Select = document.getElementById('nextEng2');
    const copyBtn = document.getElementById('copyBtn');
    const outputText = document.getElementById('outputText');
    const alertBox = document.getElementById('alertBox');

    // --- A. Load Engineers & Initialize UI ---
    fetch('engineers.json')
        .then(response => response.json())
        .then(data => {
            data.forEach(engineer => {
                eng1Select.add(new Option(engineer, engineer));
                eng2Select.add(new Option(engineer, engineer));
                nextEng1Select.add(new Option(engineer, engineer));
                nextEng2Select.add(new Option(engineer, engineer));
            });

            // Aktifkan Aesthetic Dropdown (Choices.js)
            const allSelects = document.querySelectorAll('select.glass-input');
            allSelects.forEach(select => {
                new Choices(select, {
                    searchEnabled: false,
                    itemSelectText: '',
                    shouldSort: false
                });
            });
        })
        .catch(err => showAlert('Error loading engineers.json. Make sure the file exists.', 'danger'));

    // --- B. UI Helpers ---
    function showAlert(message, type) {
        alertBox.className = `alert alert-${type} mb-3`;
        alertBox.innerHTML = `<i class="fa-solid fa-circle-exclamation me-2"></i>${message}`;
        alertBox.classList.remove('d-none');
        setTimeout(() => alertBox.classList.add('d-none'), 5000);
    }

    // --- C. Excel File Processing ---
    form.addEventListener('submit', (e) => {
        e.preventDefault();
        
        const fileInput = document.getElementById('excelFile').files[0];
        if (!fileInput) return showAlert('Please upload an Excel file.', 'danger');

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
                let headerRowIndex = 0;
                
                for (let i = 0; i < rawData.length; i++) {
                    if (rawData[i] && rawData[i].includes('Request ID')) {
                        headerRowIndex = i;
                        break;
                    }
                }

                const tickets = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex, raw: false, defval: "" });
                processTickets(tickets);
            } catch (error) {
                showAlert('Invalid Excel format. Please upload a valid .xlsx or .xls file.', 'danger');
                console.error(error);
            }
        };
        reader.readAsArrayBuffer(fileInput);
    });

    // --- D. Smart Date Engine & Ticketing Logic ---
    function processTickets(tickets) {
        if (tickets.length === 0) return showAlert('The uploaded Excel file is empty or formatted incorrectly.', 'danger');

        const firstTicket = tickets[0];
        const requiredColumns = ['Request ID', 'Subject', 'Request Status', 'Created Time'];
        
        const missingCols = requiredColumns.filter(col => !(col in firstTicket));
        
        if (missingCols.length > 0) {
            return showAlert(`Missing columns: ${missingCols.join(', ')}. Check export.`, 'danger');
        }

        const reportDateVal = document.getElementById('reportDate').value;
        const shiftVal = document.getElementById('shiftSelect').value;
        
        let shiftStart, shiftEnd;
        if (shiftVal === 'AM') {
            shiftStart = dayjs(`${reportDateVal} 07:00:00`);
            shiftEnd = dayjs(`${reportDateVal} 18:59:59`);
        } else {
            shiftStart = dayjs(`${reportDateVal} 19:00:00`);
            shiftEnd = dayjs(`${reportDateVal} 06:59:59`).add(1, 'day');
        }

        let total = 0, closed = 0;
        let inProgressTexts = [];
        let resolvedTexts = [];

        tickets.forEach(ticket => {
            const statusCol = ticket['Request Status'] ? ticket['Request Status'].toString().trim().toLowerCase() : '';
            let dateCol = ticket['Created Time'];
            
            if (!dateCol) return; 

            let ticketDate;

            if (typeof dateCol === 'number') {
                // Selesaikan masalah Excel Decimal 
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                const ms = Math.round(dateCol * 24 * 60 * 60 * 1000);
                const jsDate = new Date(excelEpoch.getTime() + ms);
                ticketDate = dayjs(`${jsDate.getUTCFullYear()}-${jsDate.getUTCMonth()+1}-${jsDate.getUTCDate()} ${jsDate.getUTCHours()}:${jsDate.getUTCMinutes()}`);
            } else {
                const dateStr = dateCol.toString().replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').trim();
                const match = dateStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s+(\d{1,2}):(\d{2})/);
                
                if (match) {
                    let part1 = parseInt(match[1]);
                    let part2 = parseInt(match[2]);
                    let year = parseInt(match[3]);
                    let hour = parseInt(match[4]);
                    let min = parseInt(match[5]);
                    
                    // SMART DATE LOGIC: Kesan format US (MM/DD) vs format Malaysia (DD/MM)
                    const expectedMonth = dayjs(reportDateVal).month() + 1;
                    const expectedMonthNext = dayjs(reportDateVal).add(1, 'day').month() + 1;
                    
                    let day, month;
                    if (part2 === expectedMonth || part2 === expectedMonthNext) {
                        day = part1; month = part2; 
                    } else if (part1 === expectedMonth || part1 === expectedMonthNext) {
                        month = part1; day = part2; 
                    } else {
                        day = part1; month = part2; 
                    }

                    const isPM = /PM/i.test(dateStr);
                    const isAM = /AM/i.test(dateStr);
                    const hasAmPm = isPM || isAM;
                    
                    if (isPM && hour < 12) hour += 12;
                    if (isAM && hour === 12) hour = 0;
                    
                    if (!hasAmPm) {
                        if (shiftVal === 'AM' && hour >= 1 && hour <= 6) hour += 12;
                        if (shiftVal === 'PM') {
                            if (hour >= 7 && hour <= 11) hour += 12;
                            if (hour === 12) hour = 0;
                        }
                    }
                    
                    ticketDate = dayjs(`${year}-${month.toString().padStart(2, '0')}-${day.toString().padStart(2, '0')} ${hour.toString().padStart(2, '0')}:${min.toString().padStart(2, '0')}`);
                } else {
                    ticketDate = dayjs(dateStr); 
                }
            }
            
            if (ticketDate.isValid() && ((ticketDate.isAfter(shiftStart) && ticketDate.isBefore(shiftEnd)) || ticketDate.isSame(shiftStart) || ticketDate.isSame(shiftEnd))) {
                total++;
                
                const reqId = ticket['Request ID'].toString().trim();
                const subject = ticket['Subject'].toString().replace(/[\n\r]+/g, ' ').trim();
                
                if (statusCol.includes('closed')) {
                    closed++;
                } else if (statusCol.includes('resolve')) {
                    resolvedTexts.push(`#${reqId} - ${subject}`);
                } else {
                    inProgressTexts.push(`#${reqId} - ${subject}`);
                }
            }
        });

        if (total === 0) {
            showAlert('No records found for the selected date and shift.', 'warning');
        }

        generateOutputString(reportDateVal, shiftVal, total, closed, inProgressTexts, resolvedTexts);
    }

    // --- E. Format Generator & Auto-Resize Textarea ---
    function generateOutputString(dateVal, shiftVal, total, closed, inProgressTexts, resolvedTexts) {
        const eng1 = document.getElementById('eng1').value;
        const eng2 = document.getElementById('eng2').value;
        const nextEng1 = document.getElementById('nextEng1').value;
        const nextEng2 = document.getElementById('nextEng2').value;
        
        const formattedDate = dayjs(dateVal).format('DD/MM/YYYY');
        const nextShift = shiftVal === 'AM' ? 'PM' : 'AM';
        const nextShiftDate = shiftVal === 'PM' ? dayjs(dateVal).add(1, 'day').format('DD/MM/YYYY') : formattedDate;
        
        let output = `Date: ${formattedDate}\n`;
        output += `Shift ${shiftVal} : ${eng1} & ${eng2}\n\n`;
        
        output += `==========================\n\n`;
        
        output += `Total : ${total}\n`;
        output += `Closed : ${closed}\n\n`;
        
        output += `In Progress :\n`;
        if (inProgressTexts.length > 0) {
            output += inProgressTexts.join('\n') + '\n\n';
        } else {
            output += '\n';
        }

        output += `Resolve :\n`;
        if (resolvedTexts.length > 0) {
            output += resolvedTexts.join('\n') + '\n\n';
        } else {
            output += '\n';
        }
        
        output += `==========================\n\n`;

        output += `Date: ${nextShiftDate}\n`;
        output += `Shift ${nextShift} : ${nextEng1} & ${nextEng2}`;

        // Set the text
        outputText.value = output;

        // BULLETPROOF AUTO-RESIZE TRICK
        outputText.style.height = '0px'; 
        outputText.style.height = (outputText.scrollHeight) + 'px';

        // Button states
        copyBtn.disabled = false;
        copyBtn.innerHTML = '<i class="fa-regular fa-copy me-1"></i> Copy Report';
        copyBtn.classList.remove('btn-outline-success');
        copyBtn.classList.add('btn-glass-success'); // Guna class butang custom kita
    }

    // --- F. Clipboard API (Butang Copy) ---
    copyBtn.addEventListener('click', () => {
        navigator.clipboard.writeText(outputText.value).then(() => {
            copyBtn.innerHTML = '<i class="fa-solid fa-check me-1"></i> Copied!';
        }).catch(err => {
            showAlert('Failed to copy to clipboard.', 'danger');
        });
    });
});
