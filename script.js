document.addEventListener('contextmenu', function(e) {
    e.preventDefault();
});

document.addEventListener('keydown', function(e) {
    if (e.key === 'F12' || e.keyCode === 123) e.preventDefault();
    if (e.ctrlKey && e.shiftKey && (e.key === 'I' || e.key === 'i')) e.preventDefault();
    if (e.ctrlKey && e.shiftKey && (e.key === 'C' || e.key === 'c')) e.preventDefault();
    if (e.ctrlKey && e.shiftKey && (e.key === 'J' || e.key === 'j')) e.preventDefault();
    if (e.ctrlKey && (e.key === 'U' || e.key === 'u')) e.preventDefault();
});

const waveParams = { speed: 1.0, opacity: 0.4 };

document.addEventListener("DOMContentLoaded", () => {
    const cards = document.querySelectorAll(".cyber-card");
    
    // Inisialisasi Vanilla-Tilt
    function initTilt() {
        VanillaTilt.init(cards, {
            max: 3,
            speed: 400,
            glare: true,
            "max-glare": 0.15,
            perspective: 1000
        });
    }
    initTilt();

    // KEMASKINI PENTING: Bunuh (destroy) terus Vanilla-Tilt sewaktu animasi supaya 0% penyek/cacat
    function destroyTilt() {
        cards.forEach(card => {
            if (card.vanillaTilt) {
                card.vanillaTilt.destroy();
                card.style.transform = ""; 
            }
        });
    }

    const form = document.getElementById('reportForm');
    const eng1Select = document.getElementById('eng1');
    const eng2Select = document.getElementById('eng2');
    const nextEng1Select = document.getElementById('nextEng1');
    const nextEng2Select = document.getElementById('nextEng2');
    const copyBtn = document.getElementById('copyBtn');
    const outputText = document.getElementById('outputText');
    const alertBox = document.getElementById('alertBox');
    const fileInput = document.getElementById('excelFile');
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const generateBtn = document.getElementById('generateBtn');
    const configCol = document.getElementById('configColumn');
    const terminalCol = document.getElementById('terminalColumn');
    const backBtn = document.getElementById('backBtn');
    const reportDateInput = document.getElementById('reportDate');

    reportDateInput.addEventListener('click', function() {
        try { this.showPicker(); } catch (err) {}
    });

    const cursorDot = document.querySelector('.cursor-dot');
    const cursorRing = document.querySelector('.cursor-ring');
    gsap.set(cursorDot, { xPercent: -50, yPercent: -50 });
    gsap.set(cursorRing, { xPercent: -50, yPercent: -50 });

    window.addEventListener('mousemove', (e) => {
        gsap.to(cursorDot, { x: e.clientX, y: e.clientY, duration: 0 });
        gsap.to(cursorRing, { x: e.clientX, y: e.clientY, duration: 0.15, ease: "power2.out" });
    });

    const bindCursorHover = () => {
        const interactiveElements = document.querySelectorAll('button, a, input, select, .file-drop-area, .choices, .mac-buttons-wrapper, .glass-textarea');
        interactiveElements.forEach(el => {
            el.addEventListener('mouseenter', () => cursorRing.classList.add('hovered'));
            el.addEventListener('mouseleave', () => cursorRing.classList.remove('hovered'));
        });
    };
    bindCursorHover();

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            fileNameDisplay.textContent = `✓ TARGET LOCKED: ${fileInput.files[0].name}`;
            fileNameDisplay.classList.remove('d-none');
        } else {
            fileNameDisplay.classList.add('d-none');
        }
    });

    fetch('engineers.json')
        .then(response => response.json())
        .then(data => {
            data.forEach(engineer => {
                eng1Select.add(new Option(engineer, engineer));
                eng2Select.add(new Option(engineer, engineer));
                nextEng1Select.add(new Option(engineer, engineer));
                nextEng2Select.add(new Option(engineer, engineer));
            });
            
            // KEMASKINI PENTING: Asingkan posisi dropdown. Yang bawah WAJIB buka ke atas (top) supaya tak terpotong!
            const topSelects = [document.getElementById('shiftSelect'), eng1Select, eng2Select];
            const bottomSelects = [nextEng1Select, nextEng2Select];

            topSelects.forEach(select => {
                new Choices(select, { searchEnabled: false, itemSelectText: '', shouldSort: false, position: 'bottom', shouldSortItems: false });
            });

            bottomSelects.forEach(select => {
                new Choices(select, { searchEnabled: false, itemSelectText: '', shouldSort: false, position: 'top', shouldSortItems: false });
            });

            document.querySelectorAll('.choices__inner, .choices__item').forEach(dd => {
                dd.addEventListener('mouseenter', () => cursorRing.classList.add('hovered'));
                dd.addEventListener('mouseleave', () => cursorRing.classList.remove('hovered'));
            });
        })
        .catch(err => showAlert('System Error: engineers.json missing.', 'danger'));

    function showAlert(message, type) {
        alertBox.className = `alert alert-${type} mb-4 fw-bold tracking-wider small shadow-lg`;
        alertBox.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-2 fa-bounce"></i>${message}`;
        alertBox.classList.remove('d-none');
        gsap.fromTo(alertBox, { opacity: 0, y: -20, scale: 0.95 }, { opacity: 1, y: 0, scale: 1, duration: 0.4, ease: "back.out(1.7)" });
        setTimeout(() => {
            gsap.to(alertBox, { opacity: 0, y: -20, scale: 0.95, duration: 0.3, onComplete: () => alertBox.classList.add('d-none') });
        }, 6000);
    }

    form.addEventListener('submit', (e) => {
        e.preventDefault();
        
        if (!reportDateInput.value || !eng1Select.value || !eng2Select.value || !nextEng1Select.value || !nextEng2Select.value) {
            return showAlert('INCOMPLETE MATRIX: Fill all required fields.', 'warning');
        }

        const file = fileInput.files[0];
        if (!file) return showAlert('MISSING DATA: Upload ManageEngine Log.', 'danger');

        const reader = new FileReader();
        reader.onload = (event) => {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
                let headerRowIndex = -1;
                
                for (let i = 0; i < rawData.length; i++) {
                    if (rawData[i] && rawData[i].includes('Request ID')) {
                        headerRowIndex = i;
                        break;
                    }
                }

                if (headerRowIndex === -1) {
                    return showAlert('COLUMN ERROR: "Request ID" header not found. Did ManageEngine export format change?', 'danger');
                }

                const tickets = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex, raw: false, defval: "" });
                processTickets(tickets);
            } catch (error) {
                showAlert(`CORRUPT DATA: ${error.message || 'Invalid Excel format.'}`, 'danger');
                console.error("Excel Reading Error:", error);
            }
        };
        reader.readAsArrayBuffer(file);
    });

    function parseExcelTimestamp(dateCol, reportDateVal, shiftVal) {
        if (!dateCol) return null;
        if (typeof dateCol === 'number') {
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            const ms = Math.round(dateCol * 24 * 60 * 60 * 1000);
            const jsDate = new Date(excelEpoch.getTime() + ms);
            return dayjs(`${jsDate.getUTCFullYear()}-${jsDate.getUTCMonth()+1}-${jsDate.getUTCDate()} ${jsDate.getUTCHours()}:${jsDate.getUTCMinutes()}`);
        } else {
            const dateStr = dateCol.toString().replace(/[\n\r]+/g, ' ').replace(/\s+/g, ' ').trim();
            const match = dateStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s+(\d{1,2}):(\d{2})/);
            if (match) {
                let part1 = parseInt(match[1]), part2 = parseInt(match[2]), year = parseInt(match[3]), hour = parseInt(match[4]), min = parseInt(match[5]);
                const expectedMonth = dayjs(reportDateVal).month() + 1;
                const expectedMonthNext = dayjs(reportDateVal).add(1, 'day').month() + 1;
                let day, month;
                if (part2 === expectedMonth || part2 === expectedMonthNext) { day = part1; month = part2; } 
                else if (part1 === expectedMonth || part1 === expectedMonthNext) { month = part1; day = part2; } 
                else { day = part1; month = part2; }

                const isPM = /PM/i.test(dateStr), isAM = /AM/i.test(dateStr), hasAmPm = isPM || isAM;
                if (isPM && hour < 12) hour += 12;
                if (isAM && hour === 12) hour = 0;
                if (!hasAmPm) {
                    if (shiftVal === 'AM' && hour >= 1 && hour <= 6) hour += 12;
                    if (shiftVal === 'PM') {
                        if (hour >= 1 && hour <= 6) hour += 12;
                        if (hour === 12) hour = 0;
                    }
                }
                return dayjs(`${year}-${month.toString().padStart(2, '0')}-${day.toString().padStart(2, '0')} ${hour.toString().padStart(2, '0')}:${min.toString().padStart(2, '0')}`);
            }
            return dayjs(dateStr);
        }
    }

    function processTickets(tickets) {
        try {
            if (tickets.length !== 0) {
                const firstTicket = tickets[0];
                const requiredColumns = ['Request ID', 'Subject', 'Request Status', 'Created Time'];
                const missingCols = requiredColumns.filter(col => !(col in firstTicket));
                if (missingCols.length > 0) {
                    return showAlert(`MISSING COLUMNS: [ ${missingCols.join(', ')} ]. Check your export settings.`, 'danger');
                }
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

            if (shiftStart.isAfter(dayjs())) {
                return showAlert('INVALID TIMELINE: Selected shift has not started yet! Please check your date.', 'warning');
            }

            let total = 0, closed = 0, inProgressTexts = [], resolvedTexts = [];

            tickets.forEach(ticket => {
                const statusCol = ticket['Request Status'] ? ticket['Request Status'].toString().trim().toLowerCase() : '';
                
                let assignedDate = parseExcelTimestamp(ticket['Assigned Time'], reportDateVal, shiftVal);
                let createdDate = parseExcelTimestamp(ticket['Created Time'], reportDateVal, shiftVal);
                
                let ticketDate = (assignedDate && assignedDate.isValid()) ? assignedDate : createdDate;
                if (!ticketDate || !ticketDate.isValid()) return;
                
                if ((ticketDate.isAfter(shiftStart) && ticketDate.isBefore(shiftEnd)) || ticketDate.isSame(shiftStart) || ticketDate.isSame(shiftEnd)) {
                    total++;
                    const reqId = ticket['Request ID'].toString().trim();
                    const subject = ticket['Subject'].toString().replace(/[\n\r]+/g, ' ').trim();
                    
                    if (statusCol.includes('closed')) closed++;
                    else if (statusCol.includes('resolve')) resolvedTexts.push(`#${reqId} - ${subject}`);
                    else inProgressTexts.push(`#${reqId} - ${subject}`);
                }
            });

            if (total === 0) {
                showAlert('QUIET SHIFT DETECTED: Zero incidents found. Generating clean handover report...', 'success');
            }

            generateBtn.innerHTML = '<i class="fa-solid fa-satellite-dish ms-2 fa-spin"></i> EXECUTING...';
            generateBtn.disabled = true;

            // KEMASKINI PENTING: Matikan (destroy) Tilt dan gunakan transisi Scale-Zoom (0% putaran = 0% penyek!)
            destroyTilt();

            const tl = gsap.timeline({
                onComplete: () => {
                    generateOutputString(reportDateVal, shiftVal, total, closed, inProgressTexts, resolvedTexts);
                    generateBtn.innerHTML = 'Execute Report <i class="fa-solid fa-arrow-right ms-2"></i>';
                    generateBtn.disabled = false;
                    initTilt(); // Hidupkan semula Tilt apabila transisi siap mendarat
                }
            });

            tl.to(waveParams, { speed: 30.0, opacity: 1, duration: 0.5, ease: "power4.in" }, 0);
            tl.to(configCol, { 
                scale: 0.92,
                y: -25,
                opacity: 0, 
                duration: 0.45, 
                ease: "power2.inOut",
                onComplete: () => {
                    configCol.classList.add('d-none');
                    terminalCol.classList.remove('d-none');
                    outputText.value = ""; 
                }
            }, 0.15);

            tl.fromTo(terminalCol, 
                { scale: 1.08, y: 25, opacity: 0 }, 
                { scale: 1, y: 0, opacity: 1, duration: 0.55, ease: "power2.out", clearProps: "transform" }, 
                ">"
            );

            tl.to(waveParams, { speed: 1.0, opacity: 0.4, duration: 1.5, ease: "power4.out" }, "-=0.2");
            
        } catch (err) {
            showAlert(`PROCESSING ERROR: ${err.message}. Check browser console for details.`, 'danger');
            console.error("Data Processing Error:", err);
        }
    }

    function generateOutputString(dateVal, shiftVal, total, closed, inProgressTexts, resolvedTexts) {
        const eng1 = document.getElementById('eng1').value;
        const eng2 = document.getElementById('eng2').value;
        const nextEng1 = document.getElementById('nextEng1').value;
        const nextEng2 = document.getElementById('nextEng2').value;
        const formattedDate = dayjs(dateVal).format('DD/MM/YYYY');
        const nextShift = shiftVal === 'AM' ? 'PM' : 'AM';
        const nextShiftDate = shiftVal === 'PM' ? dayjs(dateVal).add(1, 'day').format('DD/MM/YYYY') : formattedDate;
        
        let output = `Date: ${formattedDate}\n`;
        output += `Shift ${shiftVal}: ${eng1} & ${eng2}\n\n`;
        output += `==========================\n\n`;
        output += `Total : ${total}\n`;
        output += `Closed : ${closed}\n\n`;
        output += `In progress : ${inProgressTexts.length}\n`;
        if (inProgressTexts.length > 0) output += inProgressTexts.join('\n') + '\n\n';
        else output += '\n';

        output += `Resolve : ${resolvedTexts.length}\n`;
        if (resolvedTexts.length > 0) output += resolvedTexts.join('\n') + '\n\n';
        else output += '\n';
        
        output += `==========================\n\n`;
        output += `Date: ${nextShiftDate}\n`;
        output += `Shift ${nextShift} : ${nextEng1} & ${nextEng2}`;

        outputText.value = "";
        copyBtn.disabled = true;
        copyBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin-pulse me-1"></i> DECRYPTING...';
        copyBtn.classList.remove('btn-glass-success');
        copyBtn.classList.add('btn-outline-secondary');

        let i = 0;
        function typeWriter() {
            if (i < output.length) {
                outputText.value += output.charAt(i);
                i++;
                setTimeout(typeWriter, Math.random() * 5 + 1);
            } else {
                copyBtn.disabled = false;
                copyBtn.innerHTML = '<i class="fa-regular fa-copy me-1"></i> Copy Report';
                copyBtn.classList.remove('btn-outline-secondary');
                copyBtn.classList.add('btn-glass-success');
            }
        }
        
        setTimeout(typeWriter, 300);
    }

    backBtn.addEventListener('click', () => {
        destroyTilt();

        const tl = gsap.timeline({
            onComplete: () => {
                initTilt(); 
            }
        });

        tl.to(waveParams, { speed: -25.0, opacity: 0.8, duration: 0.4, ease: "power3.in" }, 0);
        
        tl.to(terminalCol, { 
            scale: 0.92,
            y: 25,
            opacity: 0, 
            duration: 0.4, 
            ease: "power2.inOut",
            onComplete: () => {
                terminalCol.classList.add('d-none');
                configCol.classList.remove('d-none');
                form.reset();
                fileNameDisplay.classList.add('d-none');
                fileNameDisplay.textContent = '';
                outputText.value = "";
            }
        }, 0.1);

        tl.fromTo(configCol, 
            { scale: 1.08, y: -25, opacity: 0 }, 
            { scale: 1, y: 0, opacity: 1, duration: 0.55, ease: "power2.out", clearProps: "transform" }, 
            ">"
        );

        tl.to(waveParams, { speed: 1.0, opacity: 0.4, duration: 1.2, ease: "power3.out" }, "-=0.2");
    });

    copyBtn.addEventListener('click', () => {
        navigator.clipboard.writeText(outputText.value).then(() => {
            copyBtn.innerHTML = '<i class="fa-solid fa-check me-1"></i> Copied!';
        }).catch(err => {
            showAlert('CLIPBOARD ERROR.', 'danger');
        });
    });
});

const canvas = document.getElementById('bg-canvas');
const scene = new THREE.Scene();
const camera = new THREE.PerspectiveCamera(60, window.innerWidth / window.innerHeight, 0.1, 1000);
const renderer = new THREE.WebGLRenderer({ canvas: canvas, alpha: true, antialias: true });

renderer.setSize(window.innerWidth, window.innerHeight);
renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

const particleCount = 1200;
const particleGeometry = new THREE.BufferGeometry();
const positions = new Float32Array(particleCount * 3);
const velocities = []; 

for (let i = 0; i < particleCount; i++) {
    positions[i * 3] = (Math.random() - 0.5) * 30; 
    positions[i * 3 + 1] = (Math.random() - 0.5) * 30; 
    positions[i * 3 + 2] = (Math.random() - 0.5) * 15 - 5; 

    velocities.push({
        x: (Math.random() - 0.5) * 0.002,
        y: (Math.random() * 0.005) + 0.002
    });
}

particleGeometry.setAttribute('position', new THREE.BufferAttribute(positions, 3));

const particleMaterial = new THREE.PointsMaterial({
    size: 0.04,
    color: 0x0ea5e9,
    transparent: true,
    opacity: 0.6,
    blending: THREE.AdditiveBlending
});

const particles = new THREE.Points(particleGeometry, particleMaterial);
scene.add(particles);

camera.position.z = 5;

let targetX = 0, targetY = 0;
document.addEventListener('mousemove', (event) => {
    targetX = (event.clientX / window.innerWidth - 0.5) * 0.8;
    targetY = (event.clientY / window.innerHeight - 0.5) * 0.8;
});

function animate() {
    requestAnimationFrame(animate);
    
    particleMaterial.opacity = waveParams.opacity;

    const positionAttribute = particleGeometry.getAttribute('position');
    for (let i = 0; i < particleCount; i++) {
        let y = positionAttribute.getY(i);
        let x = positionAttribute.getX(i);
        
        y += velocities[i].y * waveParams.speed;
        x += velocities[i].x * waveParams.speed;
        
        if (y > 15) y = -15;
        if (y < -15) y = 15; 
        if (x > 15) x = -15;
        if (x < -15) x = 15;
        
        positionAttribute.setY(i, y);
        positionAttribute.setX(i, x);
    }
    positionAttribute.needsUpdate = true;

    camera.position.x += (targetX - camera.position.x) * 0.03;
    camera.position.y += (-targetY - camera.position.y) * 0.03;
    camera.lookAt(scene.position);

    renderer.render(scene, camera);
}

animate();

window.addEventListener('resize', () => {
    camera.aspect = window.innerWidth / window.innerHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(window.innerWidth, window.innerHeight);
});