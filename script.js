// ==========================================
// VARIABLES GLOBALES
// ==========================================
let activeMethod = 'sturges'; 
let globalDatasets = []; 
let uploadedFilesMap = new Map(); 
const MAX_DATASETS = 10;
let currentSlide = 0;

// Utilidades Numéricas
const cleanNum = (num, decimals = 4) => {
    if (isNaN(num)) return 0;
    const fixed = parseFloat(num.toFixed(decimals));
    return Number.isInteger(fixed) ? fixed : fixed;
};

const getPercentile = (data, p) => {
    const n = data.length;
    const idx = (p / 100) * (n - 1);
    const l = Math.floor(idx);
    return l + 1 >= n ? data[l] : data[l] * (1 - (idx % 1)) + data[l + 1] * (idx % 1);
};

function createStatRow(label, value, formula) {
    return `
        <div class="stat-row">
            <span class="tooltip">${label}<span class="tooltiptext">${formula}</span></span>
            <b>${value}</b>
        </div>
    `;
}

// ==========================================
// CONTROLADORES DE UI (INPUTS Y COLA)
// ==========================================
document.querySelectorAll('input[name="kMethod"]').forEach(r => {
    r.addEventListener('change', (e) => {
        const inputManual = document.getElementById('kManualValue');
        inputManual.disabled = e.target.value !== 'manual';
        if (!inputManual.disabled) inputManual.focus();
    });
});

document.getElementById('fileInput').addEventListener('change', handleFileUpload);

function handleFileUpload(e) {
    const files = e.target.files;
    if (files.length > MAX_DATASETS) {
        alert(`Has subido ${files.length} archivos. Solo se procesarán los primeros ${MAX_DATASETS}.`);
    }

    uploadedFilesMap.clear();
    const queueList = document.getElementById('fileList');
    queueList.innerHTML = '';
    
    let limit = Math.min(files.length, MAX_DATASETS);
    for (let i = 0; i < limit; i++) {
        let file = files[i];
        let fileId = `file_${i}`;
        uploadedFilesMap.set(fileId, { file: file, customRanges: [] });

        let li = document.createElement('li');
        li.innerHTML = `
            <span>📄 ${file.name}</span>
            <button class="preview-btn" onclick="openExcelModal('${fileId}')">🔍 Previsualizar / Seleccionar</button>
        `;
        queueList.appendChild(li);
    }
    document.getElementById('fileQueue').classList.remove('hidden');
}

document.getElementById('processBtn').addEventListener('click', processAllData);
document.getElementById('exportBtn').addEventListener('click', exportAllToExcel);

// ==========================================
// PROCESAMIENTO PRINCIPAL
// ==========================================
async function processAllData() {
    if (uploadedFilesMap.size === 0) return alert("Sube al menos un archivo Excel.");
    
    activeMethod = document.querySelector('input[name="kMethod"]:checked').value;
    if (activeMethod === 'manual') {
        const manualK = parseInt(document.getElementById('kManualValue').value);
        if (isNaN(manualK) || manualK < 1) return alert("Ingresa un número de intervalos (k) válido.");
    }

    globalDatasets = [];
    document.getElementById('resultsArea').classList.add('hidden');
    
    for (let [fileId, fileData] of uploadedFilesMap) {
        if (fileData.customRanges.length > 0) {
            fileData.customRanges.forEach((rangeNums, idx) => {
                globalDatasets.push(calculateStatsForDataset(rangeNums, `${fileData.file.name} (Rango ${idx+1})`));
            });
        } else {
            const raw = await extractNumbersFromFile(fileData.file);
            if (raw.length > 0) {
                globalDatasets.push(calculateStatsForDataset(raw, fileData.file.name));
            }
        }
    }

    if(globalDatasets.length > MAX_DATASETS) {
        globalDatasets = globalDatasets.slice(0, MAX_DATASETS);
        alert(`Se han detectado demasiados conjuntos de datos. Se limitará a ${MAX_DATASETS} análisis.`);
    }

    renderCarousel();
}

function extractNumbersFromFile(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});
            
            let nums = [];
            json.forEach(row => {
                row.forEach(cell => {
                    let num = parseFloat(cell);
                    if (!isNaN(num)) nums.push(num);
                });
            });
            resolve(nums);
        };
        reader.readAsArrayBuffer(file);
    });
}

// ==========================================
// MODAL VISUAL INTERACTIVO (SHIFT + CTRL + AUTO-SCROLL)
// ==========================================
let preview2DArray = [];
let isDragging = false;
let startCell = null;
let lastClickedCell = null;
let autoScrollInterval = null;
let currentPreviewFileId = null;

function openExcelModal(fileId) {
    currentPreviewFileId = fileId;
    const fileObj = uploadedFilesMap.get(fileId);
    document.getElementById('modalTitle').innerText = `Previsualización: ${fileObj.file.name}`;
    updateRangeCount();
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        preview2DArray = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1, defval: ""});
        
        renderPreviewTable();
        document.getElementById('previewModal').classList.remove('hidden');
    };
    reader.readAsArrayBuffer(fileObj.file);
}

document.getElementById('closeModalBtn').onclick = () => document.getElementById('previewModal').classList.add('hidden');

function renderPreviewTable() {
    const container = document.getElementById('tableContainer');
    let html = '<table id="interactiveTable">';
    preview2DArray.forEach((row, rIdx) => {
        html += '<tr>';
        row.forEach((cell, cIdx) => {
            html += `<td data-r="${rIdx}" data-c="${cIdx}">${cell !== undefined ? cell : ''}</td>`;
        });
        html += '</tr>';
    });
    html += '</table>';
    container.innerHTML = html;

    const table = document.getElementById('interactiveTable');

    table.addEventListener('mousedown', (e) => {
        if(e.target.tagName !== 'TD') return;
        isDragging = true;
        const r = parseInt(e.target.dataset.r);
        const c = parseInt(e.target.dataset.c);

        if (e.ctrlKey || e.metaKey) {
            e.target.classList.toggle('cell-selected');
            lastClickedCell = {r, c};
            startCell = null; 
        } else if (e.shiftKey && lastClickedCell) {
            selectRange(lastClickedCell, {r, c}, true);
        } else {
            clearSelection();
            e.target.classList.add('cell-selected');
            startCell = {r, c};
            lastClickedCell = {r, c};
        }
    });

    table.addEventListener('mouseover', (e) => {
        if(!isDragging || !startCell || e.target.tagName !== 'TD') return;
        const r = parseInt(e.target.dataset.r);
        const c = parseInt(e.target.dataset.c);
        selectRange(startCell, {r, c}, true);
    });

    container.addEventListener('mousemove', (e) => {
        if(!isDragging) return;
        const rect = container.getBoundingClientRect();
        const buffer = 40; 
        let dx = 0, dy = 0;

        if (e.clientX < rect.left + buffer) dx = -15;
        else if (e.clientX > rect.right - buffer) dx = 15;
        
        if (e.clientY < rect.top + buffer) dy = -15;
        else if (e.clientY > rect.bottom - buffer) dy = 15;

        if (dx !== 0 || dy !== 0) {
            if (!autoScrollInterval) {
                autoScrollInterval = setInterval(() => {
                    container.scrollLeft += dx;
                    container.scrollTop += dy;
                }, 30);
            }
        } else {
            clearInterval(autoScrollInterval);
            autoScrollInterval = null;
        }
    });

    window.addEventListener('mouseup', () => { 
        isDragging = false; 
        clearInterval(autoScrollInterval);
        autoScrollInterval = null;
    });
}

function selectRange(start, end, clearFirst) {
    if(clearFirst) clearSelection();
    const minR = Math.min(start.r, end.r), maxR = Math.max(start.r, end.r);
    const minC = Math.min(start.c, end.c), maxC = Math.max(start.c, end.c);
    for(let r = minR; r <= maxR; r++) {
        for(let c = minC; c <= maxC; c++) {
            const td = document.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
            if(td && !td.classList.contains('cell-saved')) td.classList.add('cell-selected');
        }
    }
}

function clearSelection() {
    document.querySelectorAll('.cell-selected').forEach(td => td.classList.remove('cell-selected'));
}

document.getElementById('saveRangeBtn').addEventListener('click', () => {
    const fileObj = uploadedFilesMap.get(currentPreviewFileId);
    if(fileObj.customRanges.length >= MAX_DATASETS) return alert("Máximo de rangos alcanzado para este archivo.");
    
    const selected = document.querySelectorAll('.cell-selected');
    if(selected.length === 0) return alert("Selecciona datos primero.");

    let nums = [];
    selected.forEach(td => {
        let val = parseFloat(td.innerText);
        if(!isNaN(val)) nums.push(val);
        td.classList.remove('cell-selected');
        td.classList.add('cell-saved');
    });

    if(nums.length === 0) return alert("No hay números válidos en tu selección.");
    fileObj.customRanges.push(nums);
    updateRangeCount();
});

function updateRangeCount() {
    const fileObj = uploadedFilesMap.get(currentPreviewFileId);
    document.getElementById('rangeCount').innerText = `Rangos guardados: ${fileObj.customRanges.length}`;
}

document.getElementById('finishRangesBtn').addEventListener('click', () => {
    document.getElementById('previewModal').classList.add('hidden');
});

// ==========================================
// MOTOR ESTADÍSTICO
// ==========================================
function calculateStatsForDataset(raw, datasetName) {
    let data = [...raw].sort((a, b) => a - b);
    const n = data.length;
    const minVal = data[0];
    const maxVal = data[n - 1];
    const range = maxVal - minVal;
    
    let numClasses = activeMethod === 'manual' ? parseInt(document.getElementById('kManualValue').value) : Math.round(1 + 3.322 * Math.log10(n));
    if (numClasses < 1) numClasses = 1;
    
    const amplitude = range / numClasses;
    let classesData = [];
    let currentMin = minVal;
    let cumulativeFreq = 0;

    for (let i = 0; i < numClasses; i++) {
        let currentMax = currentMin + amplitude;
        let isLast = (i === numClasses - 1);
        if (isLast) currentMax = maxVal; 

        let count = data.filter(x => x >= currentMin && (isLast ? x <= currentMax : x < currentMax)).length;
        let xi = (currentMin + currentMax) / 2;
        cumulativeFreq += count;
        
        classesData.push({ min: currentMin, max: currentMax, isLast, xi, fi: count, Fi: cumulativeFreq, hi: count/n, Hi: cumulativeFreq/n });
        currentMin = currentMax;
    }

    const sum = data.reduce((a, b) => a + b, 0);
    const mean = sum / n;
    const geoMean = Math.exp(data.reduce((s, x) => s + Math.log(x), 0) / n);
    const harMean = n / data.reduce((s, x) => s + (1 / x), 0);
    const median = n % 2 !== 0 ? data[Math.floor(n/2)] : (data[Math.floor(n/2)-1] + data[Math.floor(n/2)]) / 2;

    let freqMap = {}; let maxFreq = 0; let mode = [];
    data.forEach(num => { freqMap[num] = (freqMap[num] || 0) + 1; if (freqMap[num] > maxFreq) maxFreq = freqMap[num]; });
    for (const key in freqMap) if (freqMap[key] === maxFreq) mode.push(Number(key));

    const variance = data.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (n - 1);
    const stdDev = Math.sqrt(variance);
    const cv = (stdDev / mean) * 100;

    let skewness = 0;
    if (n > 2 && stdDev > 0) skewness = (n / ((n - 1) * (n - 2))) * data.reduce((acc, val) => acc + Math.pow((val - mean) / stdDev, 3), 0);

    return { name: datasetName, data, n, minVal, maxVal, range, numClasses, amplitude, classesData, stats: { mean, geoMean, harMean, median, mode, variance, stdDev, cv, skewness, p10: getPercentile(data,10), q1: getPercentile(data,25), q2: getPercentile(data,50), q3: getPercentile(data,75), p90: getPercentile(data,90) }};
}

// ==========================================
// CARRUSEL DE RENDERIZADO Y SINCRONIZACIÓN DE UI
// ==========================================
function renderCarousel() {
    const carousel = document.getElementById('resultsCarousel');
    carousel.innerHTML = '';
    currentSlide = 0;
    
    let kFormula = activeMethod === 'sturges' ? 'k ≈ 1 + 3.322 · log₁₀(n)' : 'Manual';
    let methodLabel = activeMethod === 'sturges' ? ' (Sturges)' : ' (Manual)';

    globalDatasets.forEach((ds, index) => {
        let block = document.createElement('div');
        block.className = 'dataset-block';
        
        let freqHtml = `
            <h3>Análisis ${index + 1}: ${ds.name}</h3>
            <table>
                <thead>
                    <tr><th>Límite Inf. (Li)</th><th>Límite Sup. (Ls)</th><th>Marca de Clase (Xi)</th><th>Frec. Abs. (fi)</th><th>Frec. Acum. (Fi)</th><th>Frec. Rel. (hi)</th><th>Frec. Rel. Acum. (Hi)</th></tr>
                </thead>
                <tbody>
        `;
        
        ds.classesData.forEach(c => {
            freqHtml += `<tr><td>${cleanNum(c.min)}</td><td>${cleanNum(c.max)}</td><td>${cleanNum(c.xi)}</td><td>${c.fi}</td><td>${c.Fi}</td><td>${cleanNum(c.hi)}</td><td>${cleanNum(c.Hi)}</td></tr>`;
        });
        freqHtml += `</tbody></table>`;

        let statsHtml = `
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>Parámetros Base</h3>
                    ${createStatRow('Mínimo:', cleanNum(ds.minVal), 'min(xᵢ)')}
                    ${createStatRow('Máximo:', cleanNum(ds.maxVal), 'max(xᵢ)')}
                    ${createStatRow(`Intervalos (k)${methodLabel}:`, ds.numClasses, kFormula)}
                    ${createStatRow('Amplitud (A):', cleanNum(ds.amplitude), 'A = Rango / k')}
                </div>
                <div class="stat-card">
                    <h3>Tendencia Central</h3>
                    ${createStatRow('Media Arit.:', cleanNum(ds.stats.mean), 'x̄ = (Σxᵢ) / n')}
                    ${createStatRow('Media Geom.:', cleanNum(ds.stats.geoMean), 'MG = ⁿ√(x₁···xₙ)')}
                    ${createStatRow('Media Arm.:', cleanNum(ds.stats.harMean), 'MH = n / Σ(1/xᵢ)')}
                    ${createStatRow('Mediana:', cleanNum(ds.stats.median), 'Me = Lᵢ + A·[(n/2 - Fᵢ₋₁)/fᵢ]')}
                    ${createStatRow('Moda:', ds.stats.mode.map(m=>cleanNum(m)).join(','), 'Mo = Lᵢ + A·[(fᵢ - fᵢ₋₁)/(2fᵢ - fᵢ₋₁ - fᵢ₊₁)]')}
                </div>
                <div class="stat-card">
                    <h3>Dispersión y Forma</h3>
                    ${createStatRow('Rango:', cleanNum(ds.range), 'R = x_max - x_min')}
                    ${createStatRow('Varianza:', cleanNum(ds.stats.variance), 's² = Σ(xᵢ - x̄)² / (n - 1)')}
                    ${createStatRow('Desv. Est.:', cleanNum(ds.stats.stdDev), 's = √s²')}
                    ${createStatRow('C. Variación:', cleanNum(ds.stats.cv, 2) + '%', 'CV = (s/x̄)·100%')}
                    ${createStatRow('Asimetría:', cleanNum(ds.stats.skewness), 'As = [n/((n-1)(n-2))] · Σ[(xᵢ-x̄)/s]³')}
                </div>
                <div class="stat-card">
                    <h3>Posición (Percentiles)</h3>
                    ${createStatRow('P10 (10%):', cleanNum(ds.stats.p10), 'P₁₀ = Lᵢ + A·[(10n/100 - Fᵢ₋₁)/fᵢ]')}
                    ${createStatRow('Q1 (25%):', cleanNum(ds.stats.q1), 'Q₁')}
                    ${createStatRow('Q2 (50%):', cleanNum(ds.stats.q2), 'Q₂ = Mediana')}
                    ${createStatRow('Q3 (75%):', cleanNum(ds.stats.q3), 'Q₃')}
                    ${createStatRow('P90 (90%):', cleanNum(ds.stats.p90), 'P₉₀ = Lᵢ + A·[(90n/100 - Fᵢ₋₁)/fᵢ]')}
                </div>
            </div>
        `;

        block.innerHTML = freqHtml + statsHtml;
        carousel.appendChild(block);
    });

    document.getElementById('resultsArea').classList.remove('hidden');
    document.getElementById('exportBtn').classList.remove('hidden');
    
    // Resetear posición de scroll al inicio
    carousel.scrollLeft = 0;
    updateCarouselControls();
}

// Función exclusiva para actualizar textos y botones
function updateCarouselControls() {
    document.getElementById('carouselIndicator').innerText = `Tabla ${currentSlide + 1} de ${globalDatasets.length}`;
    document.getElementById('prevBtn').disabled = currentSlide === 0;
    document.getElementById('nextBtn').disabled = currentSlide === globalDatasets.length - 1;
}

// Control por Botones
document.getElementById('prevBtn').addEventListener('click', () => {
    if (currentSlide > 0) { 
        currentSlide--; 
        const carousel = document.getElementById('resultsCarousel');
        carousel.scrollTo({ left: carousel.clientWidth * currentSlide, behavior: 'smooth' });
        updateCarouselControls(); 
    }
});

document.getElementById('nextBtn').addEventListener('click', () => {
    if (currentSlide < globalDatasets.length - 1) { 
        currentSlide++; 
        const carousel = document.getElementById('resultsCarousel');
        carousel.scrollTo({ left: carousel.clientWidth * currentSlide, behavior: 'smooth' });
        updateCarouselControls(); 
    }
});

// Control por Scroll Manual (Sincronización)
document.getElementById('resultsCarousel').addEventListener('scroll', (e) => {
    const carousel = e.target;
    const width = carousel.clientWidth;
    // Calculamos matemáticamente en qué "slide" estamos basados en la posición del scroll
    const newSlide = Math.round(carousel.scrollLeft / width);
    
    // Si la diapositiva en pantalla cambió, actualizamos la UI
    if (newSlide !== currentSlide && newSlide >= 0 && newSlide < globalDatasets.length) {
        currentSlide = newSlide;
        updateCarouselControls();
    }
});

// ==========================================
// EXPORTACIÓN A EXCELJS
// ==========================================
async function exportAllToExcel() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Generador Estadístico Lotes';

    const headerStyle = {
        font: { bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } },
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }
    };

    globalDatasets.forEach((ds, idx) => {
        let shortName = ds.name.substring(0, 20).replace(/[:*?/"<>|]/g, ''); 
        
        const wsDatos = wb.addWorksheet(`D_${idx+1}_${shortName}`);
        wsDatos.getCell('A1').value = "DATOS ORDENADOS";
        wsDatos.getCell('A1').font = headerStyle.font; wsDatos.getCell('A1').fill = headerStyle.fill;
        ds.data.forEach((val, i) => { wsDatos.getCell(`A${i + 2}`).value = val; });
        wsDatos.getColumn('A').width = 20;
        
        const dataRange = `'D_${idx+1}_${shortName}'!A2:A${ds.n + 1}`;

        const ws = wb.addWorksheet(`A_${idx+1}_${shortName}`);
        const headers = ['LÍMITE INF. (LI)', 'LÍMITE SUP. (LS)', 'MARCA DE CLASE (XI)', 'FREC. ABS (FI)', 'FREC. ACUM (FI)', 'FREC. REL (HI)', 'FREC. REL ACUM (HI)'];
        ws.addRow(headers);
        ws.getRow(1).eachCell((cell) => { Object.assign(cell, headerStyle); });

        ds.classesData.forEach((cls, i) => {
            let rowNum = i + 2; 
            let cond = cls.isLast ? "<=" : "<";
            ws.addRow([
                cls.min, cls.max,
                { formula: `(A${rowNum}+B${rowNum})/2`, result: cls.xi },
                { formula: `COUNTIFS(${dataRange},">="&A${rowNum},${dataRange},"${cond}"&B${rowNum})`, result: cls.fi },
                i === 0 ? { formula: `D2`, result: cls.Fi } : { formula: `E${rowNum - 1}+D${rowNum}`, result: cls.Fi },
                { formula: `D${rowNum}/COUNT(${dataRange})`, result: cls.hi },
                i === 0 ? { formula: `F2`, result: cls.Hi } : { formula: `G${rowNum - 1}+F${rowNum}`, result: cls.Hi }
            ]).eachCell(cell => cell.alignment = { horizontal: 'center' });
        });

        for(let c = 1; c <= 8; c++) ws.getColumn(c).width = 20;

        let startRow = ds.numClasses + 4;
        ws.getCell(`A${startRow}`).value = "PARÁMETROS"; ws.getCell(`C${startRow}`).value = "TENDENCIA C."; 
        ws.getCell(`E${startRow}`).value = "DISPERSIÓN"; ws.getCell(`G${startRow}`).value = "POSICIÓN";
        [ws.getCell(`A${startRow}`), ws.getCell(`C${startRow}`), ws.getCell(`E${startRow}`), ws.getCell(`G${startRow}`)].forEach(c => { c.font = { bold: true }; c.border = { bottom: { style: 'medium' } }; });

        let formulaK = activeMethod === 'sturges' ? `ROUND(1+3.322*LOG10(COUNT(${dataRange})),0)` : ds.numClasses;
        let filaK = startRow + 4; 
        let formAmp = `(MAX(${dataRange})-MIN(${dataRange}))/B${filaK}`;

        const statsGrid = [
            { c1: 'A', l1: 'Mínimo:', f1: `MIN(${dataRange})`, c2: 'C', l2: 'Media Arit.:', f2: `AVERAGE(${dataRange})`, c3: 'E', l3: 'Rango:', f3: `MAX(${dataRange})-MIN(${dataRange})`, c4: 'G', l4: 'P10:', f4: `PERCENTILE(${dataRange}, 0.1)` },
            { c1: 'A', l1: 'Máximo:', f1: `MAX(${dataRange})`, c2: 'C', l2: 'Media Geom.:', f2: `GEOMEAN(${dataRange})`, c3: 'E', l3: 'Varianza:', f3: `VAR(${dataRange})`, c4: 'G', l4: 'Q1 (25%):', f4: `QUARTILE(${dataRange}, 1)` },
            { c1: 'A', l1: `Int. (k):`, f1: formulaK, c2: 'C', l2: 'Media Arm.:', f2: `HARMEAN(${dataRange})`, c3: 'E', l3: 'Desv. Est.:', f3: `STDEV(${dataRange})`, c4: 'G', l4: 'Q2 (50%):', f4: `MEDIAN(${dataRange})` },
            { c1: 'A', l1: 'Amplitud:', f1: formAmp, c2: 'C', l2: 'Mediana:', f2: `MEDIAN(${dataRange})`, c3: 'E', l3: 'CV:', f3: `STDEV(${dataRange})/AVERAGE(${dataRange})`, c4: 'G', l4: 'Q3 (75%):', f4: `QUARTILE(${dataRange}, 3)` },
            { c1: 'A', l1: '', f1: '', c2: 'C', l2: 'Moda:', f2: `MODE(${dataRange})`, c3: 'E', l3: 'Asimetría:', f3: `SKEW(${dataRange})`, c4: 'G', l4: 'P90:', f4: `PERCENTILE(${dataRange}, 0.9)` }
        ];

        statsGrid.forEach((st, i) => {
            let r = startRow + 2 + i;
            if(st.l1) { ws.getCell(`${st.c1}${r}`).value = st.l1; ws.getCell(`B${r}`).value = (st.c1 === 'A' && r === filaK && activeMethod === 'manual') ? st.f1 : { formula: st.f1 }; }
            if(st.l2) { ws.getCell(`${st.c2}${r}`).value = st.l2; ws.getCell(`D${r}`).value = { formula: st.f2 }; }
            if(st.l3) { ws.getCell(`${st.c3}${r}`).value = st.l3; ws.getCell(`F${r}`).value = { formula: st.f3 }; if(st.l3 === 'CV:') ws.getCell(`F${r}`).numFmt = '0.00%'; }
            if(st.l4) { ws.getCell(`${st.c4}${r}`).value = st.l4; ws.getCell(`H${r}`).value = { formula: st.f4 }; }
        });
    });

    const buffer = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), 'Analisis_Lotes_Avanzado.xlsx');
}