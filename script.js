// ==========================================
// VARIABLES GLOBALES
// ==========================================
let activeMethod = 'sturges'; 
let globalDatasets = []; 
let uploadedFilesMap = new Map(); 
const MAX_DATASETS = 10;
let currentSlide = 0;

if (typeof Chart !== 'undefined' && typeof ChartBoxPlot !== 'undefined') {
    Chart.register(ChartBoxPlot.BoxPlotController, ChartBoxPlot.BoxAndWiskers);
}

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
// CONTROLADORES DE UI
// ==========================================
document.querySelectorAll('input[name="kMethod"]').forEach(r => {
    r.addEventListener('change', (e) => {
        const inputManual = document.getElementById('kManualValue');
        inputManual.disabled = e.target.value !== 'manual';
        if (!inputManual.disabled) inputManual.focus();
    });
});

document.querySelectorAll('input[name="uploadMode"]').forEach(r => {
    r.addEventListener('change', (e) => {
        const fileInput = document.getElementById('fileInput');
        const fileQueue = document.getElementById('fileQueue');
        const pasteArea = document.getElementById('pasteArea');
        
        if (e.target.value === 'file') {
            fileInput.classList.remove('hidden');
            pasteArea.classList.add('hidden');
            fileInput.setAttribute('multiple', 'multiple');
            if (uploadedFilesMap.size > 0) fileQueue.classList.remove('hidden');
        } else if (e.target.value === 'paste') {
            fileInput.classList.add('hidden');
            fileQueue.classList.add('hidden');
            pasteArea.classList.remove('hidden');
        }
    });
});

document.getElementById('fileInput').addEventListener('change', handleFileUpload);

function handleFileUpload(e) {
    const files = e.target.files;
    if (files.length > MAX_DATASETS) alert(`Has subido ${files.length} archivos. Solo se procesarán los primeros ${MAX_DATASETS}.`);

    uploadedFilesMap.clear();
    const queueList = document.getElementById('fileList');
    queueList.innerHTML = '';
    
    let limit = Math.min(files.length, MAX_DATASETS);
    for (let i = 0; i < limit; i++) {
        let file = files[i];
        let fileId = `file_${i}`;
        uploadedFilesMap.set(fileId, { file: file, customRanges: [] });

        let li = document.createElement('li');
        li.innerHTML = `<span>${file.name}</span><button class="preview-btn" onclick="openExcelModal('${fileId}')">Previsualizar / Seleccionar</button>`;
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
    activeMethod = document.querySelector('input[name="kMethod"]:checked').value;
    if (activeMethod === 'manual') {
        const manualK = parseInt(document.getElementById('kManualValue').value);
        if (isNaN(manualK) || manualK < 1) return alert("Ingresa un número de intervalos válido.");
    }

    const uploadMode = document.querySelector('input[name="uploadMode"]:checked').value;
    globalDatasets = [];
    document.getElementById('resultsArea').classList.add('hidden');

    if (uploadMode === 'paste') {
        const text = document.getElementById('pasteInput').value;
        if (!text.trim()) return alert("Por favor, pega algunos datos en el cuadro de texto.");
        
        const rawStrings = text.split(/[;,\/\s\n]+/);
        const rawNums = rawStrings.map(s => parseFloat(s)).filter(n => !isNaN(n));
        if (rawNums.length === 0) return alert("No se encontraron números válidos. Asegúrate de usar punto (.) para los decimales.");
        
        globalDatasets.push(calculateStatsForDataset(rawNums, "Datos Pegados"));
        renderCarousel();
        return;
    }
    
    if (uploadedFilesMap.size === 0) return alert("Sube al menos un archivo Excel.");
    
    for (let [fileId, fileData] of uploadedFilesMap) {
        if (fileData.customRanges.length > 0) {
            fileData.customRanges.forEach((rangeNums, idx) => {
                globalDatasets.push(calculateStatsForDataset(rangeNums, `${fileData.file.name} (Rango ${idx+1})`));
            });
        } else {
            const raw = await extractNumbersFromFile(fileData.file);
            if (raw.length > 0) globalDatasets.push(calculateStatsForDataset(raw, fileData.file.name));
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
            json.forEach(row => row.forEach(cell => { let num = parseFloat(cell); if (!isNaN(num)) nums.push(num); }));
            resolve(nums);
        };
        reader.readAsArrayBuffer(file);
    });
}

// ==========================================
// MODAL DE EXCEL (INTERACTIVO)
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
        const savedRangesCount = fileObj.customRanges.length;
        if(savedRangesCount > 0) alert(`Este archivo ya tiene ${savedRangesCount} rango(s) guardado(s). Las celdas oscuras están bloqueadas.`);
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
        row.forEach((cell, cIdx) => html += `<td data-r="${rIdx}" data-c="${cIdx}">${cell !== undefined ? cell : ''}</td>`);
        html += '</tr>';
    });
    html += '</table>';
    container.innerHTML = html;

    const table = document.getElementById('interactiveTable');

    table.addEventListener('mousedown', (e) => {
        if(e.target.tagName !== 'TD' || e.target.classList.contains('cell-saved')) return; 
        isDragging = true;
        const r = parseInt(e.target.dataset.r), c = parseInt(e.target.dataset.c);

        if (e.ctrlKey || e.metaKey) {
            e.target.classList.toggle('cell-selected'); lastClickedCell = {r, c}; startCell = null; 
        } else if (e.shiftKey && lastClickedCell) {
            selectRange(lastClickedCell, {r, c}, true);
        } else {
            clearSelection(); e.target.classList.add('cell-selected'); startCell = {r, c}; lastClickedCell = {r, c};
        }
    });

    table.addEventListener('mouseover', (e) => {
        if(!isDragging || !startCell || e.target.tagName !== 'TD' || e.target.classList.contains('cell-saved')) return; 
        const r = parseInt(e.target.dataset.r), c = parseInt(e.target.dataset.c);
        selectRange(startCell, {r, c}, true);
    });

    container.addEventListener('mousemove', (e) => {
        if(!isDragging) return;
        const rect = container.getBoundingClientRect(); const buffer = 40; let dx = 0, dy = 0;
        if (e.clientX < rect.left + buffer) dx = -15; else if (e.clientX > rect.right - buffer) dx = 15;
        if (e.clientY < rect.top + buffer) dy = -15; else if (e.clientY > rect.bottom - buffer) dy = 15;

        if (dx !== 0 || dy !== 0) {
            if (!autoScrollInterval) autoScrollInterval = setInterval(() => { container.scrollLeft += dx; container.scrollTop += dy; }, 30);
        } else { clearInterval(autoScrollInterval); autoScrollInterval = null; }
    });

    window.addEventListener('mouseup', () => { isDragging = false; clearInterval(autoScrollInterval); autoScrollInterval = null; });
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

function clearSelection() { document.querySelectorAll('.cell-selected').forEach(td => td.classList.remove('cell-selected')); }
document.getElementById('clearSelectionBtn').addEventListener('click', clearSelection);

document.getElementById('resetRangesBtn').addEventListener('click', () => {
    if(!confirm("¿Estás seguro de que deseas borrar todos los rangos guardados para este archivo?")) return;
    const fileObj = uploadedFilesMap.get(currentPreviewFileId);
    fileObj.customRanges = []; document.querySelectorAll('.cell-saved').forEach(td => td.classList.remove('cell-saved')); updateRangeCount();
});

document.getElementById('saveRangeBtn').addEventListener('click', () => {
    const fileObj = uploadedFilesMap.get(currentPreviewFileId);
    if(fileObj.customRanges.length >= MAX_DATASETS) return alert("Máximo de rangos alcanzado para este archivo.");
    const selected = document.querySelectorAll('.cell-selected');
    if(selected.length === 0) return alert("Selecciona datos primero.");

    let nums = [];
    selected.forEach(td => { let val = parseFloat(td.innerText); if(!isNaN(val)) nums.push(val); td.classList.remove('cell-selected'); td.classList.add('cell-saved'); });

    if(nums.length === 0) return alert("No hay números válidos en tu selección.");
    fileObj.customRanges.push(nums); updateRangeCount();
});

function updateRangeCount() { document.getElementById('rangeCount').innerText = `Rangos guardados: ${uploadedFilesMap.get(currentPreviewFileId).customRanges.length}/10`; }
document.getElementById('finishRangesBtn').addEventListener('click', () => document.getElementById('previewModal').classList.add('hidden'));


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
// SISTEMA DE GRÁFICOS INTERACTIVOS (CHART.JS)
// ==========================================
let activeModalChart = null; // Variable para gestionar el gráfico ampliado en el modal

// Configuración general para gráficas pequeñas
const miniChartOptions = {
    responsive: true,
    maintainAspectRatio: false, // Permite que se adapte al CSS
    plugins: { legend: { display: false } },
    scales: { x: { display: true }, y: { display: true } }
};

function getChartConfig(ds, type) {
    const labels = ds.classesData.map(c => cleanNum(c.xi));
    
    if (type === 'hist') {
        return {
            data: {
                labels: labels,
                datasets: [
                    { type: 'bar', label: 'Frecuencia (fi)', data: ds.classesData.map(c => c.fi), backgroundColor: 'rgba(0, 0, 0, 0.1)', borderColor: '#000', borderWidth: 1, barPercentage: 1, categoryPercentage: 1 },
                    { type: 'line', label: 'Polígono', data: ds.classesData.map(c => c.fi), borderColor: '#000', borderWidth: 2, tension: 0.1, fill: false, pointBackgroundColor: '#000' }
                ]
            }
        };
    } else if (type === 'ojiva') {
        return {
            type: 'line',
            data: {
                labels: labels,
                datasets: [{ label: 'Ojiva (Hi %)', data: ds.classesData.map(c => c.Hi * 100), borderColor: '#000', borderWidth: 2, fill: true, backgroundColor: 'rgba(0, 0, 0, 0.05)', tension: 0.3 }]
            }
        };
    } else if (type === 'box') {
        return {
            type: 'boxplot',
            data: {
                labels: ['Distribución'],
                datasets: [{ label: ds.name, data: [ds.data], backgroundColor: 'rgba(0, 0, 0, 0.1)', borderColor: '#000', borderWidth: 2, itemRadius: 3, outlierBackgroundColor: '#000' }]
            }
        };
    }
}

function renderChartsForDataset(ds, index) {
    // 1. Histograma
    new Chart(document.getElementById(`chartHist-${index}`).getContext('2d'), {
        ...getChartConfig(ds, 'hist'),
        options: { ...miniChartOptions, scales: { y: { beginAtZero: true } } }
    });

    // 2. Ojiva
    new Chart(document.getElementById(`chartOjiva-${index}`).getContext('2d'), {
        ...getChartConfig(ds, 'ojiva'),
        options: { ...miniChartOptions, scales: { y: { beginAtZero: true, max: 100 } } }
    });

    // 3. Diagrama de Caja
    try {
        new Chart(document.getElementById(`chartBox-${index}`).getContext('2d'), {
            ...getChartConfig(ds, 'box'),
            options: { ...miniChartOptions, indexAxis: 'y' }
        });
    } catch (e) {
        document.getElementById(`chartBox-${index}`).parentElement.innerHTML = `<p style="color:#a00000; text-align:center; padding:20px;">Error al cargar BoxPlot.</p>`;
    }
}

// NUEVO: Función para abrir el modal con el gráfico ampliado
window.openChartModal = function(dsIndex, chartType) {
    const ds = globalDatasets[dsIndex];
    const modal = document.getElementById('chartModal');
    const titleEl = document.getElementById('chartModalTitle');
    const ctx = document.getElementById('modalCanvas').getContext('2d');
    
    // Destruir gráfico previo si existe para que no se superpongan
    if (activeModalChart) activeModalChart.destroy();
    
    modal.classList.remove('hidden');

    let titleText = "";
    let options = { responsive: true, maintainAspectRatio: false };

    if (chartType === 'hist') {
        titleText = "Histograma y Polígono de Frecuencias";
        options.scales = { y: { beginAtZero: true } };
    } else if (chartType === 'ojiva') {
        titleText = "Ojiva de Frecuencias (Menor que)";
        options.scales = { y: { beginAtZero: true, max: 100 } };
    } else if (chartType === 'box') {
        titleText = "Diagrama de Caja y Bigotes";
        options.indexAxis = 'y';
        options.plugins = { legend: { display: false } };
    }

    titleEl.innerText = `${titleText} - ${ds.name}`;

    activeModalChart = new Chart(ctx, {
        ...getChartConfig(ds, chartType),
        options: options
    });
};

document.getElementById('closeChartModalBtn').addEventListener('click', () => {
    document.getElementById('chartModal').classList.add('hidden');
});

window.toggleAccordion = function(header) {
    const item = header.parentElement;
    item.classList.toggle('active');
};


// ==========================================
// CARRUSEL DE RENDERIZADO
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
            <table><thead><tr><th>Límite Inf. (Li)</th><th>Límite Sup. (Ls)</th><th>Marca de Clase (Xi)</th><th>Frec. Abs. (fi)</th><th>Frec. Acum. (Fi)</th><th>Frec. Rel. (hi)</th><th>Frec. Rel. Acum. (Hi)</th></tr></thead><tbody>
        `;
        ds.classesData.forEach(c => { freqHtml += `<tr><td>${cleanNum(c.min)}</td><td>${cleanNum(c.max)}</td><td>${cleanNum(c.xi)}</td><td>${c.fi}</td><td>${c.Fi}</td><td>${cleanNum(c.hi)}</td><td>${cleanNum(c.Hi)}</td></tr>`; });
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

        // RESTRUCTURACIÓN A 1 SOLO ACORDEÓN CON 3 COLUMNAS
        let chartsHtml = `
            <div class="chart-accordion">
                <div class="accordion-item">
                    <div class="accordion-header" onclick="toggleAccordion(this)">
                        <span>Gráficos Estadísticos (Clic en un gráfico para ampliar)</span>
                        <span class="accordion-arrow">▼</span>
                    </div>
                    <div class="accordion-content">
                        <div class="chart-grid">
                            <div class="chart-item" onclick="openChartModal(${index}, 'hist')">
                                <h4>Histograma y Polígono</h4>
                                <div class="chart-canvas-wrapper"><canvas id="chartHist-${index}"></canvas></div>
                            </div>
                            <div class="chart-item" onclick="openChartModal(${index}, 'ojiva')">
                                <h4>Ojiva (Menor que)</h4>
                                <div class="chart-canvas-wrapper"><canvas id="chartOjiva-${index}"></canvas></div>
                            </div>
                            <div class="chart-item" onclick="openChartModal(${index}, 'box')">
                                <h4>Caja y Bigotes</h4>
                                <div class="chart-canvas-wrapper"><canvas id="chartBox-${index}"></canvas></div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;

        block.innerHTML = freqHtml + statsHtml + chartsHtml;
        carousel.appendChild(block);
        
        setTimeout(() => renderChartsForDataset(ds, index), 0);
    });

    document.getElementById('resultsArea').classList.remove('hidden');
    document.getElementById('exportBtn').classList.remove('hidden');
    
    carousel.scrollLeft = 0;
    updateCarouselControls();
}

function updateCarouselControls() {
    document.getElementById('carouselIndicator').innerText = `Tabla ${currentSlide + 1} de ${globalDatasets.length}`;
    document.getElementById('prevBtn').disabled = currentSlide === 0;
    document.getElementById('nextBtn').disabled = currentSlide === globalDatasets.length - 1;
}

document.getElementById('prevBtn').addEventListener('click', () => {
    if (currentSlide > 0) { currentSlide--; const carousel = document.getElementById('resultsCarousel'); carousel.scrollTo({ left: carousel.clientWidth * currentSlide, behavior: 'smooth' }); updateCarouselControls(); }
});

document.getElementById('nextBtn').addEventListener('click', () => {
    if (currentSlide < globalDatasets.length - 1) { currentSlide++; const carousel = document.getElementById('resultsCarousel'); carousel.scrollTo({ left: carousel.clientWidth * currentSlide, behavior: 'smooth' }); updateCarouselControls(); }
});

document.getElementById('resultsCarousel').addEventListener('scroll', (e) => {
    const newSlide = Math.round(e.target.scrollLeft / e.target.clientWidth);
    if (newSlide !== currentSlide && newSlide >= 0 && newSlide < globalDatasets.length) { currentSlide = newSlide; updateCarouselControls(); }
});

// ==========================================
// EXPORTACIÓN A EXCELJS
// ==========================================
async function exportAllToExcel() {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Generador Estadístico Lotes';

    const headerStyle = { font: { bold: true, color: { argb: 'FFFFFFFF' } }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } }, alignment: { horizontal: 'center', vertical: 'middle' }, border: { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} } };

    globalDatasets.forEach((ds, idx) => {
        let shortName = ds.name.substring(0, 20).replace(/[:*?/"<>|]/g, ''); 
        const wsDatos = wb.addWorksheet(`D_${idx+1}_${shortName}`);
        wsDatos.getCell('A1').value = "DATOS ORDENADOS"; wsDatos.getCell('A1').font = headerStyle.font; wsDatos.getCell('A1').fill = headerStyle.fill;
        ds.data.forEach((val, i) => wsDatos.getCell(`A${i + 2}`).value = val); wsDatos.getColumn('A').width = 20;
        
        const dataRange = `'D_${idx+1}_${shortName}'!A2:A${ds.n + 1}`;
        const ws = wb.addWorksheet(`A_${idx+1}_${shortName}`);
        ws.addRow(['LÍMITE INF. (LI)', 'LÍMITE SUP. (LS)', 'MARCA DE CLASE (XI)', 'FREC. ABS (FI)', 'FREC. ACUM (FI)', 'FREC. REL (HI)', 'FREC. REL ACUM (HI)']);
        ws.getRow(1).eachCell((cell) => Object.assign(cell, headerStyle));

        ds.classesData.forEach((cls, i) => {
            let rowNum = i + 2, cond = cls.isLast ? "<=" : "<";
            ws.addRow([
                cls.min, cls.max, { formula: `(A${rowNum}+B${rowNum})/2`, result: cls.xi },
                { formula: `COUNTIFS(${dataRange},">="&A${rowNum},${dataRange},"${cond}"&B${rowNum})`, result: cls.fi },
                i === 0 ? { formula: `D2`, result: cls.Fi } : { formula: `E${rowNum - 1}+D${rowNum}`, result: cls.Fi },
                { formula: `D${rowNum}/COUNT(${dataRange})`, result: cls.hi },
                i === 0 ? { formula: `F2`, result: cls.Hi } : { formula: `G${rowNum - 1}+F${rowNum}`, result: cls.Hi }
            ]).eachCell(cell => cell.alignment = { horizontal: 'center' });
        });

        for(let c = 1; c <= 8; c++) ws.getColumn(c).width = 20;

        let startRow = ds.numClasses + 4;
        ws.getCell(`A${startRow}`).value = "PARÁMETROS"; ws.getCell(`C${startRow}`).value = "TENDENCIA C."; ws.getCell(`E${startRow}`).value = "DISPERSIÓN"; ws.getCell(`G${startRow}`).value = "POSICIÓN";
        [ws.getCell(`A${startRow}`), ws.getCell(`C${startRow}`), ws.getCell(`E${startRow}`), ws.getCell(`G${startRow}`)].forEach(c => { c.font = { bold: true }; c.border = { bottom: { style: 'medium' } }; });

        let formulaK = activeMethod === 'sturges' ? `ROUND(1+3.322*LOG10(COUNT(${dataRange})),0)` : ds.numClasses;
        let filaK = startRow + 4, formAmp = `(MAX(${dataRange})-MIN(${dataRange}))/B${filaK}`;

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