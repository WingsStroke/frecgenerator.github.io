import { AppState } from './state.js';
import { cleanNum } from './math.js';

// --- CONFIGURACIÓN DE GRÁFICOS ---
if (typeof Chart !== 'undefined' && typeof ChartBoxPlot !== 'undefined') {
    Chart.register(ChartBoxPlot.BoxPlotController, ChartBoxPlot.BoxAndWiskers);
}

let activeModalChart = null; 
const miniChartOptions = { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { display: true }, y: { display: true } } };

function getChartConfig(ds, type) {
    const labels = ds.classesData.map(c => cleanNum(c.xi));
    if (type === 'hist') return { data: { labels, datasets: [ { type: 'bar', label: 'Frecuencia (fi)', data: ds.classesData.map(c => c.fi), backgroundColor: 'rgba(0, 0, 0, 0.1)', borderColor: '#000', borderWidth: 1, barPercentage: 1, categoryPercentage: 1 }, { type: 'line', label: 'Polígono', data: ds.classesData.map(c => c.fi), borderColor: '#000', borderWidth: 2, tension: 0.1, fill: false, pointBackgroundColor: '#000' } ] } };
    if (type === 'ojiva') return { type: 'line', data: { labels, datasets: [{ label: 'Ojiva (Hi %)', data: ds.classesData.map(c => c.Hi * 100), borderColor: '#000', borderWidth: 2, fill: true, backgroundColor: 'rgba(0, 0, 0, 0.05)', tension: 0.3 }] } };
    if (type === 'box') return { type: 'boxplot', data: { labels: ['Distribución'], datasets: [{ label: ds.name, data: [ds.data], backgroundColor: 'rgba(0, 0, 0, 0.1)', borderColor: '#000', borderWidth: 2, itemRadius: 3, outlierBackgroundColor: '#000' }] } };
}

function renderChartsForDataset(ds, index) {
    new Chart(document.getElementById(`chartHist-${index}`).getContext('2d'), { ...getChartConfig(ds, 'hist'), options: { ...miniChartOptions, scales: { y: { beginAtZero: true } } } });
    new Chart(document.getElementById(`chartOjiva-${index}`).getContext('2d'), { ...getChartConfig(ds, 'ojiva'), options: { ...miniChartOptions, scales: { y: { beginAtZero: true, max: 100 } } } });
    try { new Chart(document.getElementById(`chartBox-${index}`).getContext('2d'), { ...getChartConfig(ds, 'box'), options: { ...miniChartOptions, indexAxis: 'y' } });
    } catch (e) { document.getElementById(`chartBox-${index}`).parentElement.innerHTML = `<p style="color:#a00000; text-align:center; padding:20px;">Error al cargar BoxPlot.</p>`; }
}

window.openChartModal = function(dsIndex, chartType) {
    const ds = AppState.globalDatasets[dsIndex];
    if (activeModalChart) activeModalChart.destroy();
    document.getElementById('chartModal').classList.remove('hidden');

    let titleText = ""; let options = { responsive: true, maintainAspectRatio: false };
    if (chartType === 'hist') { titleText = "Histograma y Polígono"; options.scales = { y: { beginAtZero: true } }; } 
    else if (chartType === 'ojiva') { titleText = "Ojiva (Menor que)"; options.scales = { y: { beginAtZero: true, max: 100 } }; } 
    else if (chartType === 'box') { titleText = "Diagrama de Caja"; options.indexAxis = 'y'; options.plugins = { legend: { display: false } }; }

    document.getElementById('chartModalTitle').innerText = `${titleText} - ${ds.name}`;
    activeModalChart = new Chart(document.getElementById('modalCanvas').getContext('2d'), { ...getChartConfig(ds, chartType), options });
};
window.toggleAccordion = function(header) { header.parentElement.classList.toggle('active'); };

// --- RENDERIZADO HTML ---
function createStatRow(label, value, formula) {
    return `<div class="stat-row"><span class="tooltip">${label}<span class="tooltiptext">${formula}</span></span><b>${value}</b></div>`;
}

export function updateCarouselControls() {
    document.getElementById('carouselIndicator').innerText = `Tabla ${AppState.currentSlide + 1} de ${AppState.globalDatasets.length}`;
    document.getElementById('prevBtn').disabled = AppState.currentSlide === 0;
    document.getElementById('nextBtn').disabled = AppState.currentSlide === AppState.globalDatasets.length - 1;
}

export function renderCarousel() {
    const carousel = document.getElementById('resultsCarousel');
    carousel.innerHTML = ''; AppState.currentSlide = 0;
    let kFormula = AppState.activeMethod === 'sturges' ? 'k ≈ 1 + 3.322 · log₁₀(n)' : 'Manual';
    let methodLabel = AppState.activeMethod === 'sturges' ? ' (Sturges)' : ' (Manual)';

    AppState.globalDatasets.forEach((ds, index) => {
        let block = document.createElement('div'); block.className = 'dataset-block';
        let freqHtml = `<h3>Análisis ${index + 1}: ${ds.name}</h3><table><thead><tr><th>Límite Inf. (Li)</th><th>Límite Sup. (Ls)</th><th>Marca de Clase (Xi)</th><th>Frec. Abs. (fi)</th><th>Frec. Acum. (Fi)</th><th>Frec. Rel. (hi)</th><th>Frec. Rel. Acum. (Hi)</th></tr></thead><tbody>`;
        ds.classesData.forEach(c => { freqHtml += `<tr><td>${cleanNum(c.min)}</td><td>${cleanNum(c.max)}</td><td>${cleanNum(c.xi)}</td><td>${c.fi}</td><td>${c.Fi}</td><td>${cleanNum(c.hi)}</td><td>${cleanNum(c.Hi)}</td></tr>`; });
        freqHtml += `</tbody></table>`;

        let statsHtml = `<div class="stats-grid"><div class="stat-card"><h3>Parámetros Base</h3>${createStatRow('Mínimo:', cleanNum(ds.minVal), 'min(xᵢ)')}${createStatRow('Máximo:', cleanNum(ds.maxVal), 'max(xᵢ)')}${createStatRow(`Intervalos (k)${methodLabel}:`, ds.numClasses, kFormula)}${createStatRow('Amplitud (A):', cleanNum(ds.amplitude), 'A = Rango / k')}</div><div class="stat-card"><h3>Tendencia Central</h3>${createStatRow('Media Arit.:', cleanNum(ds.stats.mean), 'x̄ = (Σxᵢ) / n')}${createStatRow('Media Geom.:', cleanNum(ds.stats.geoMean), 'MG = ⁿ√(x₁···xₙ)')}${createStatRow('Media Arm.:', cleanNum(ds.stats.harMean), 'MH = n / Σ(1/xᵢ)')}${createStatRow('Mediana:', cleanNum(ds.stats.median), 'Me = Lᵢ + A·[(n/2 - Fᵢ₋₁)/fᵢ]')}${createStatRow('Moda:', ds.stats.mode.map(m=>cleanNum(m)).join(','), 'Mo = Lᵢ + A·[(fᵢ - fᵢ₋₁)/(2fᵢ - fᵢ₋₁ - fᵢ₊₁)]')}</div><div class="stat-card"><h3>Dispersión y Forma</h3>${createStatRow('Rango:', cleanNum(ds.range), 'R = x_max - x_min')}${createStatRow('Varianza:', cleanNum(ds.stats.variance), 's² = Σ(xᵢ - x̄)² / (n - 1)')}${createStatRow('Desv. Est.:', cleanNum(ds.stats.stdDev), 's = √s²')}${createStatRow('C. Variación:', cleanNum(ds.stats.cv, 2) + '%', 'CV = (s/x̄)·100%')}${createStatRow('Asimetría:', cleanNum(ds.stats.skewness), 'As = [n/((n-1)(n-2))] · Σ[(xᵢ-x̄)/s]³')}</div><div class="stat-card"><h3>Posición (Percentiles)</h3>${createStatRow('P10 (10%):', cleanNum(ds.stats.p10), 'P₁₀ = Lᵢ + A·[(10n/100 - Fᵢ₋₁)/fᵢ]')}${createStatRow('Q1 (25%):', cleanNum(ds.stats.q1), 'Q₁')}${createStatRow('Q2 (50%):', cleanNum(ds.stats.q2), 'Q₂ = Mediana')}${createStatRow('Q3 (75%):', cleanNum(ds.stats.q3), 'Q₃')}${createStatRow('P90 (90%):', cleanNum(ds.stats.p90), 'P₉₀ = Lᵢ + A·[(90n/100 - Fᵢ₋₁)/fᵢ]')}</div></div>`;
        let chartsHtml = `<div class="chart-accordion"><div class="accordion-item"><div class="accordion-header" onclick="toggleAccordion(this)"><span>Gráficos Estadísticos (Clic para ampliar)</span><span class="accordion-arrow">▼</span></div><div class="accordion-content"><div class="chart-grid"><div class="chart-item" onclick="openChartModal(${index}, 'hist')"><h4>Histograma</h4><div class="chart-canvas-wrapper"><canvas id="chartHist-${index}"></canvas></div></div><div class="chart-item" onclick="openChartModal(${index}, 'ojiva')"><h4>Ojiva</h4><div class="chart-canvas-wrapper"><canvas id="chartOjiva-${index}"></canvas></div></div><div class="chart-item" onclick="openChartModal(${index}, 'box')"><h4>Caja y Bigotes</h4><div class="chart-canvas-wrapper"><canvas id="chartBox-${index}"></canvas></div></div></div></div></div></div>`;

        block.innerHTML = freqHtml + statsHtml + chartsHtml; carousel.appendChild(block);
        setTimeout(() => renderChartsForDataset(ds, index), 0);
    });

    document.getElementById('resultsArea').classList.remove('hidden');
    document.getElementById('exportBtn').classList.remove('hidden');
    document.getElementById('floatingProcedureBtn').classList.remove('hidden');
    carousel.scrollLeft = 0; updateCarouselControls();
}

// --- MODAL DE EXCEL ---
let preview2DArray = []; let isDragging = false; let startCell = null; let lastClickedCell = null; let autoScrollInterval = null;

export function openExcelModal(fileId) {
    AppState.currentPreviewFileId = fileId;
    const fileObj = AppState.uploadedFilesMap.get(fileId);
    document.getElementById('modalTitle').innerText = `Previsualización: ${fileObj.file.name}`;
    updateRangeCount();
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const workbook = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
        preview2DArray = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1, defval: ""});
        renderPreviewTable();
        if(fileObj.customRanges.length > 0) alert(`Este archivo ya tiene ${fileObj.customRanges.length} rango(s) guardado(s).`);
        document.getElementById('previewModal').classList.remove('hidden');
    };
    reader.readAsArrayBuffer(fileObj.file);
}

function renderPreviewTable() {
    const container = document.getElementById('tableContainer');
    let html = '<table id="interactiveTable">';
    preview2DArray.forEach((row, rIdx) => {
        html += '<tr>'; row.forEach((cell, cIdx) => html += `<td data-r="${rIdx}" data-c="${cIdx}">${cell !== undefined ? cell : ''}</td>`); html += '</tr>';
    });
    container.innerHTML = html + '</table>';
    
    const table = document.getElementById('interactiveTable');
    table.addEventListener('mousedown', (e) => {
        if(e.target.tagName !== 'TD' || e.target.classList.contains('cell-saved')) return; 
        isDragging = true; const r = parseInt(e.target.dataset.r), c = parseInt(e.target.dataset.c);
        if (e.ctrlKey || e.metaKey) { e.target.classList.toggle('cell-selected'); lastClickedCell = {r, c}; startCell = null; } 
        else if (e.shiftKey && lastClickedCell) { selectRange(lastClickedCell, {r, c}, true); } 
        else { clearSelection(); e.target.classList.add('cell-selected'); startCell = {r, c}; lastClickedCell = {r, c}; }
    });

    table.addEventListener('mouseover', (e) => {
        if(!isDragging || !startCell || e.target.tagName !== 'TD' || e.target.classList.contains('cell-saved')) return; 
        selectRange(startCell, {r: parseInt(e.target.dataset.r), c: parseInt(e.target.dataset.c)}, true);
    });

    container.addEventListener('mousemove', (e) => {
        if(!isDragging) return;
        const rect = container.getBoundingClientRect(); const buffer = 40; let dx = 0, dy = 0;
        if (e.clientX < rect.left + buffer) dx = -15; else if (e.clientX > rect.right - buffer) dx = 15;
        if (e.clientY < rect.top + buffer) dy = -15; else if (e.clientY > rect.bottom - buffer) dy = 15;
        if (dx !== 0 || dy !== 0) { if (!autoScrollInterval) autoScrollInterval = setInterval(() => { container.scrollLeft += dx; container.scrollTop += dy; }, 30); } 
        else { clearInterval(autoScrollInterval); autoScrollInterval = null; }
    });
    window.addEventListener('mouseup', () => { isDragging = false; clearInterval(autoScrollInterval); autoScrollInterval = null; });
}

function selectRange(start, end, clearFirst) {
    if(clearFirst) clearSelection();
    const minR = Math.min(start.r, end.r), maxR = Math.max(start.r, end.r);
    const minC = Math.min(start.c, end.c), maxC = Math.max(start.c, end.c);
    for(let r = minR; r <= maxR; r++) for(let c = minC; c <= maxC; c++) {
        const td = document.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
        if(td && !td.classList.contains('cell-saved')) td.classList.add('cell-selected');
    }
}

function clearSelection() { document.querySelectorAll('.cell-selected').forEach(td => td.classList.remove('cell-selected')); }
export function updateRangeCount() { document.getElementById('rangeCount').innerText = `Rangos guardados: ${AppState.uploadedFilesMap.get(AppState.currentPreviewFileId).customRanges.length}/10`; }

// --- EVENTOS DE UI A INICIALIZAR ---
export function initUIListeners() {
    document.getElementById('prevBtn').addEventListener('click', () => { if (AppState.currentSlide > 0) { AppState.currentSlide--; document.getElementById('resultsCarousel').scrollTo({ left: document.getElementById('resultsCarousel').clientWidth * AppState.currentSlide, behavior: 'smooth' }); updateCarouselControls(); }});
    document.getElementById('nextBtn').addEventListener('click', () => { if (AppState.currentSlide < AppState.globalDatasets.length - 1) { AppState.currentSlide++; document.getElementById('resultsCarousel').scrollTo({ left: document.getElementById('resultsCarousel').clientWidth * AppState.currentSlide, behavior: 'smooth' }); updateCarouselControls(); }});
    document.getElementById('resultsCarousel').addEventListener('scroll', (e) => { const newSlide = Math.round(e.target.scrollLeft / e.target.clientWidth); if (newSlide !== AppState.currentSlide && newSlide >= 0 && newSlide < AppState.globalDatasets.length) { AppState.currentSlide = newSlide; updateCarouselControls(); }});
    
    document.getElementById('closeModalBtn').addEventListener('click', () => document.getElementById('previewModal').classList.add('hidden'));
    document.getElementById('closeChartModalBtn').addEventListener('click', () => document.getElementById('chartModal').classList.add('hidden'));
    document.getElementById('clearSelectionBtn').addEventListener('click', clearSelection);
    
    document.getElementById('resetRangesBtn').addEventListener('click', () => {
        if(!confirm("¿Borrar todos los rangos guardados para este archivo?")) return;
        AppState.uploadedFilesMap.get(AppState.currentPreviewFileId).customRanges = []; document.querySelectorAll('.cell-saved').forEach(td => td.classList.remove('cell-saved')); updateRangeCount();
    });

    document.getElementById('saveRangeBtn').addEventListener('click', () => {
        const fileObj = AppState.uploadedFilesMap.get(AppState.currentPreviewFileId);
        if(fileObj.customRanges.length >= AppState.MAX_DATASETS) return alert("Máximo de rangos alcanzado para este archivo.");
        const selected = document.querySelectorAll('.cell-selected'); if(selected.length === 0) return alert("Selecciona datos primero.");
        let nums = []; selected.forEach(td => { let val = parseFloat(td.innerText); if(!isNaN(val)) nums.push(val); td.classList.remove('cell-selected'); td.classList.add('cell-saved'); });
        if(nums.length === 0) return alert("No hay números válidos en tu selección.");
        fileObj.customRanges.push(nums); updateRangeCount();
    });

    document.getElementById('finishRangesBtn').addEventListener('click', () => document.getElementById('previewModal').classList.add('hidden'));

    // Procedimiento
    document.getElementById('floatingProcedureBtn').addEventListener('click', () => {
        const ds = AppState.globalDatasets[AppState.currentSlide]; 
        document.getElementById('procedureModalTitle').innerText = `Procedimiento: ${ds.name}`;
        let isSturges = AppState.activeMethod === 'sturges';
        
        document.getElementById('procedureContent').innerHTML = `
            <div class="procedure-step">
                <h3>1. Creación de la Tabla</h3>
                <div class="math-formula">Rango (R) = ${cleanNum(ds.maxVal)} - ${cleanNum(ds.minVal)} = <span class="math-highlight">${cleanNum(ds.range)}</span></div>
                <div class="math-formula">Intervalos (k) ${isSturges ? `≈ 1 + 3.322 * log₁₀(${ds.n}) = ` : 'Manual = '} <span class="math-highlight">${ds.numClasses}</span></div>
                <div class="math-formula">Amplitud (A) = ${cleanNum(ds.range)} / ${ds.numClasses} = <span class="math-highlight">${cleanNum(ds.amplitude)}</span></div>
            </div>
            <div class="procedure-step">
                <h3>2. Medidas Centrales</h3>
                <div class="math-formula">Media Aritmética (x̄) = ${cleanNum(ds.stats.sum)} / ${ds.n} = <span class="math-highlight">${cleanNum(ds.stats.mean)}</span></div>
                <div class="math-formula">Mediana (Me) = <span class="math-highlight">${cleanNum(ds.stats.median)}</span></div>
            </div>
            <div class="procedure-step">
                <h3>3. Dispersión</h3>
                <div class="math-formula">Varianza (s²) = ${cleanNum(ds.stats.varianceSum)} / (${ds.n} - 1) = <span class="math-highlight">${cleanNum(ds.stats.variance)}</span></div>
                <div class="math-formula">Desv. Estándar (s) = √${cleanNum(ds.stats.variance)} = <span class="math-highlight">${cleanNum(ds.stats.stdDev)}</span></div>
                <div class="math-formula">C. Variación (CV) = (${cleanNum(ds.stats.stdDev)} / ${cleanNum(ds.stats.mean)}) * 100% = <span class="math-highlight">${cleanNum(ds.stats.cv, 2)}%</span></div>
            </div>
        `;
        document.getElementById('procedureModal').classList.remove('hidden');
    });
    document.getElementById('closeProcedureModalBtn').addEventListener('click', () => document.getElementById('procedureModal').classList.add('hidden'));
}