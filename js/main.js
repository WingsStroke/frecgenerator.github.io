import { AppState } from './state.js';
import { calculateStatsForDataset } from './math.js';
import { extractNumbersFromFile, exportAllToExcel } from './excel.js';
import { renderCarousel, openExcelModal, initUIListeners } from './ui.js';
import { exportToPDF } from './pdf.js';

initUIListeners();

document.querySelectorAll('input[name="kMethod"]').forEach(r => r.addEventListener('change', (e) => {
    const inputManual = document.getElementById('kManualValue');
    inputManual.disabled = e.target.value !== 'manual';
    if (!inputManual.disabled) inputManual.focus();
}));

document.querySelectorAll('input[name="uploadMode"]').forEach(r => r.addEventListener('change', (e) => {
    const fileInput = document.getElementById('fileInput'), fileQueue = document.getElementById('fileQueue'), pasteArea = document.getElementById('pasteArea');
    if (e.target.value === 'file') {
        fileInput.classList.remove('hidden'); pasteArea.classList.add('hidden'); fileInput.setAttribute('multiple', 'multiple');
        if (AppState.uploadedFilesMap.size > 0) fileQueue.classList.remove('hidden');
    } else if (e.target.value === 'paste') {
        fileInput.classList.add('hidden'); fileQueue.classList.add('hidden'); pasteArea.classList.remove('hidden');
    }
}));

document.querySelectorAll('input[name="tableType"]').forEach(r => r.addEventListener('change', (e) => {
    const intervalSettings = document.getElementById('intervalSettings');
    if (e.target.value === 'grouped') {
        intervalSettings.classList.remove('hidden');
    } else {
        intervalSettings.classList.add('hidden');
    }
}));

document.getElementById('fileInput').addEventListener('change', (e) => {
    const files = e.target.files;
    if (files.length > AppState.MAX_DATASETS) alert(`Has subido ${files.length} archivos. Solo se procesarán ${AppState.MAX_DATASETS}.`);
    AppState.uploadedFilesMap.clear();
    const queueList = document.getElementById('fileList'); queueList.innerHTML = '';
    
    let limit = Math.min(files.length, AppState.MAX_DATASETS);
    for (let i = 0; i < limit; i++) {
        let file = files[i]; let fileId = `file_${i}`;
        AppState.uploadedFilesMap.set(fileId, { file: file, customRanges: [] });
        let li = document.createElement('li');
        li.innerHTML = `<span>${file.name}</span><button class="preview-btn" id="btn_${fileId}">Previsualizar / Seleccionar</button>`;
        queueList.appendChild(li);
        document.getElementById(`btn_${fileId}`).addEventListener('click', () => openExcelModal(fileId));
    }
    document.getElementById('fileQueue').classList.remove('hidden');
});

document.getElementById('processBtn').addEventListener('click', async () => {
    AppState.activeMethod = document.querySelector('input[name="kMethod"]:checked').value;
    const manualKValue = document.getElementById('kManualValue').value;
    const tableType = document.querySelector('input[name="tableType"]:checked').value;

    if (AppState.activeMethod === 'manual' && tableType === 'grouped') {
        const manualK = parseInt(manualKValue);
        if (isNaN(manualK) || manualK < 1) return alert("Ingresa un número de intervalos válido.");
    }

    const uploadMode = document.querySelector('input[name="uploadMode"]:checked').value;
    AppState.globalDatasets = [];
    document.getElementById('resultsArea').classList.add('hidden');
    document.getElementById('floatingProcedureBtn').classList.add('hidden');

    if (uploadMode === 'paste') {
        const text = document.getElementById('pasteInput').value;
        if (!text.trim()) return alert("Pega algunos datos.");
        const rawStrings = text.split(/[;,\/\s\n]+/);
        const rawNums = rawStrings.map(s => parseFloat(s)).filter(n => !isNaN(n));
        if (rawNums.length === 0) return alert("No se encontraron números válidos.");
        
        AppState.globalDatasets.push(calculateStatsForDataset(rawNums, "Datos Pegados", AppState.activeMethod, manualKValue, tableType));
        renderCarousel(); return;
    }
    
    if (AppState.uploadedFilesMap.size === 0) return alert("Sube al menos un archivo.");
    for (let [fileId, fileData] of AppState.uploadedFilesMap) {
        if (fileData.customRanges.length > 0) {
            fileData.customRanges.forEach((rangeNums, idx) => AppState.globalDatasets.push(calculateStatsForDataset(rangeNums, `${fileData.file.name} (Rango ${idx+1})`, AppState.activeMethod, manualKValue, tableType)));
        } else {
            const raw = await extractNumbersFromFile(fileData.file);
            if (raw.length > 0) AppState.globalDatasets.push(calculateStatsForDataset(raw, fileData.file.name, AppState.activeMethod, manualKValue, tableType));
        }
    }

    if(AppState.globalDatasets.length > AppState.MAX_DATASETS) {
        AppState.globalDatasets = AppState.globalDatasets.slice(0, AppState.MAX_DATASETS);
        alert(`Limitado a ${AppState.MAX_DATASETS} análisis.`);
    }
    renderCarousel();
});

document.getElementById('exportBtn').addEventListener('click', () => exportAllToExcel(AppState.globalDatasets, AppState.activeMethod));

document.getElementById('exportPdfBtn').addEventListener('click', () => {
    const currentDataset = AppState.globalDatasets[AppState.currentSlide];
    if(currentDataset) {
        exportToPDF(currentDataset);
    }
});