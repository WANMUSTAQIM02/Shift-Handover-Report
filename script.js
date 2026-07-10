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

// Parameter kelajuan dikurangkan untuk elak pening
const waveParams = { speed: 1.0, opacity: 0.4 };

document.addEventListener("DOMContentLoaded", () => {
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

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            fileNameDisplay.textContent = `✓ ${fileInput.files[0].name}`;
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
            const allSelects = document.querySelectorAll('select.glass-input');
            allSelects.forEach(select => {
                new Choices(select, { searchEnabled: false, itemSelectText: '', shouldSort: false });
            });
        })
        .catch(err => showAlert('Error loading engineers.json. Make sure the file exists.', 'danger'));

    function showAlert(message, type) {
        alertBox.className = `alert alert-${type} mb-3`;
        alertBox.innerHTML = `<i class="fa-solid fa-circle-exclamation me-2"></i>${message}`;
        alertBox.classList.remove('d-none');
        setTimeout(() => alertBox.classList.add('d-none'), 5000);
    }

    form.addEventListener('submit', (e) => {
        e.preventDefault();
        
        const file = fileInput.files[0];
        if (!file) return showAlert('Please upload an Excel file.', 'danger');

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
        reader.readAsArrayBuffer(file);
    });

    function processTickets(tickets) {
        if (tickets.length === 0) return showAlert('The uploaded Excel file is empty or formatted incorrectly.', 'danger');

        const firstTicket = tickets[0];
        const requiredColumns = ['Request ID', 'Subject', 'Request Status', 'Created Time'];
        const missingCols = requiredColumns.filter(col => !(col in firstTicket));
        
        if (missingCols.length > 0) return showAlert(`Missing columns: ${missingCols.join(', ')}. Check export.`, 'danger');

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

        let total = 0, closed = 0, inProgressTexts = [], resolvedTexts = [];

        tickets.forEach(ticket => {
            const statusCol = ticket['Request Status'] ? ticket['Request Status'].toString().trim().toLowerCase() : '';
            let dateCol = ticket['Created Time'];
            if (!dateCol) return; 

            let ticketDate;
            if (typeof dateCol === 'number') {
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                const ms = Math.round(dateCol * 24 * 60 * 60 * 1000);
                const jsDate = new Date(excelEpoch.getTime() + ms);
                ticketDate = dayjs(`${jsDate.getUTCFullYear()}-${jsDate.getUTCMonth()+1}-${jsDate.getUTCDate()} ${jsDate.getUTCHours()}:${jsDate.getUTCMinutes()}`);
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
                
                if (statusCol.includes('closed')) closed++;
                else if (statusCol.includes('resolve')) resolvedTexts.push(`#${reqId} - ${subject}`);
                else inProgressTexts.push(`#${reqId} - ${subject}`);
            }
        });

        if (total === 0) return showAlert('No records found for the selected date and shift.', 'warning');

        generateBtn.innerHTML = '<i class="fa-solid fa-atom fa-spin me-2"></i> Analyzing Log...';
        generateBtn.disabled = true;

        const tl = gsap.timeline({
            onComplete: () => {
                generateOutputString(reportDateVal, shiftVal, total, closed, inProgressTexts, resolvedTexts);
                generateBtn.innerHTML = 'Generate Report <i class="fa-solid fa-arrow-right ms-2"></i>';
                generateBtn.disabled = false;
            }
        });

        // Kesan Warp Speed memecut ke atas lurus
        tl.to(waveParams, { speed: 20.0, opacity: 0.8, duration: 0.8, ease: "power4.in" }, 0);
        tl.to(configCol, { 
            x: -100, 
            opacity: 0, 
            duration: 0.5, 
            ease: "power2.in",
            onComplete: () => {
                configCol.classList.add('d-none');
                terminalCol.classList.remove('d-none');
                outputText.value = ""; 
            }
        }, 0.4);

        tl.fromTo(terminalCol, 
            { x: 100, opacity: 0 }, 
            { x: 0, opacity: 1, duration: 0.8, ease: "back.out(1.4)" }, 
            ">"
        );

        // Kembali perlahan
        tl.to(waveParams, { speed: 1.0, opacity: 0.4, duration: 1.8, ease: "power4.out" }, "-=0.2");
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
        copyBtn.innerHTML = '<i class="fa-solid fa-terminal fa-fade me-1"></i> Compiling...';
        copyBtn.classList.remove('btn-glass-success');
        copyBtn.classList.add('btn-outline-secondary');
        copyBtn.style.color = '#94b899';

        let i = 0;
        function typeWriter() {
            if (i < output.length) {
                outputText.value += output.charAt(i);
                outputText.style.height = '0px'; 
                outputText.style.height = (outputText.scrollHeight) + 'px';
                i++;
                setTimeout(typeWriter, Math.random() * 8 + 2);
            } else {
                copyBtn.disabled = false;
                copyBtn.innerHTML = '<i class="fa-regular fa-copy me-1"></i> Copy Report';
                copyBtn.classList.remove('btn-outline-secondary');
                copyBtn.classList.add('btn-glass-success');
                copyBtn.style.color = '';
            }
        }
        
        setTimeout(typeWriter, 400);
    }

    backBtn.addEventListener('click', () => {
        const tl = gsap.timeline();
        // Zarah undur ke belakang perlahan-lahan
        tl.to(waveParams, { speed: -15.0, opacity: 0.7, duration: 0.5, ease: "power3.in" }, 0);
        tl.to(terminalCol, { 
            x: 100, 
            opacity: 0, 
            duration: 0.5, 
            ease: "power2.in",
            onComplete: () => {
                terminalCol.classList.add('d-none');
                configCol.classList.remove('d-none');
                form.reset();
                fileNameDisplay.classList.add('d-none');
                fileNameDisplay.textContent = '';
                outputText.value = "";
            }
        }, 0.2);

        tl.fromTo(configCol, 
            { x: -100, opacity: 0 }, 
            { x: 0, opacity: 1, duration: 0.8, ease: "back.out(1.4)" }, 
            ">"
        );

        tl.to(waveParams, { speed: 1.0, opacity: 0.4, duration: 1.5, ease: "power3.out" }, "-=0.2");
    });

    copyBtn.addEventListener('click', () => {
        navigator.clipboard.writeText(outputText.value).then(() => {
            copyBtn.innerHTML = '<i class="fa-solid fa-check me-1"></i> Copied!';
        }).catch(err => {
            showAlert('Failed to copy to clipboard.', 'danger');
        });
    });
});

// ==========================================
// BACKGROUND 3D BARU (TENANG & TAK MEMENINGKAN)
// ==========================================
const canvas = document.getElementById('bg-canvas');
const scene = new THREE.Scene();
const camera = new THREE.PerspectiveCamera(60, window.innerWidth / window.innerHeight, 0.1, 1000);
const renderer = new THREE.WebGLRenderer({ canvas: canvas, alpha: true, antialias: true });

renderer.setSize(window.innerWidth, window.innerHeight);
renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));

const particleCount = 1000;
const particleGeometry = new THREE.BufferGeometry();
const positions = new Float32Array(particleCount * 3);
const velocities = []; // Simpan kelajuan untuk setiap zarah

for (let i = 0; i < particleCount; i++) {
    // Taburan zarah yang rata di seluruh skrin
    positions[i * 3] = (Math.random() - 0.5) * 30; // x
    positions[i * 3 + 1] = (Math.random() - 0.5) * 30; // y
    positions[i * 3 + 2] = (Math.random() - 0.5) * 15 - 5; // z

    // Pergerakan hanya ke atas (y) dan sedikit ke tepi (x)
    velocities.push({
        x: (Math.random() - 0.5) * 0.001,
        y: (Math.random() * 0.003) + 0.001
    });
}

particleGeometry.setAttribute('position', new THREE.BufferAttribute(positions, 3));

const particleMaterial = new THREE.PointsMaterial({
    size: 0.05,
    color: 0x818cf8,
    transparent: true,
    opacity: 10.0,
    blending: THREE.AdditiveBlending
});

const particles = new THREE.Points(particleGeometry, particleMaterial);
scene.add(particles);

camera.position.z = 5;

// Mengurangkan sensitiviti pergerakan mouse supaya tak sakit mata
let targetX = 0, targetY = 0;
document.addEventListener('mousemove', (event) => {
    targetX = (event.clientX / window.innerWidth - 0.5) * 0.5;
    targetY = (event.clientY / window.innerHeight - 0.5) * 0.5;
});

function animate() {
    requestAnimationFrame(animate);
    
    particleMaterial.opacity = waveParams.opacity;

    // Gerakkan setiap zarah secara lurus
    const positionAttribute = particleGeometry.getAttribute('position');
    for (let i = 0; i < particleCount; i++) {
        let y = positionAttribute.getY(i);
        let x = positionAttribute.getX(i);
        
        y += velocities[i].y * waveParams.speed;
        x += velocities[i].x * waveParams.speed;
        
        // Kitar semula zarah bila terkeluar dari skrin
        if (y > 15) y = -15;
        if (y < -15) y = 15; // Untuk kesan butang Back (reverse)
        if (x > 15) x = -15;
        if (x < -15) x = 15;
        
        positionAttribute.setY(i, y);
        positionAttribute.setX(i, x);
    }
    positionAttribute.needsUpdate = true;

    // Kamera bergerak sangat perlahan mengikut mouse (Parallax ringan)
    camera.position.x += (targetX - camera.position.x) * 0.02;
    camera.position.y += (-targetY - camera.position.y) * 0.02;
    camera.lookAt(scene.position);

    renderer.render(scene, camera);
}

animate();

window.addEventListener('resize', () => {
    camera.aspect = window.innerWidth / window.innerHeight;
    camera.updateProjectionMatrix();
    renderer.setSize(window.innerWidth, window.innerHeight);
});