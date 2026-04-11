import { cleanNum } from './math.js';

// Función auxiliar para configurar los gráficos de la misma forma que en ui.js
function getChartConfig(ds, type) {
    const labels = ds.classesData.map(c => cleanNum(c.xi));
    if (type === 'hist') return { type: 'bar', data: { labels, datasets: [ { type: 'bar', label: 'Frecuencia fi', data: ds.classesData.map(c => c.fi), backgroundColor: 'rgba(0, 0, 0, 0.1)', borderColor: '#000', borderWidth: 1, barPercentage: 1, categoryPercentage: 1 }, { type: 'line', label: 'Polígono', data: ds.classesData.map(c => c.fi), borderColor: '#000', borderWidth: 2, tension: 0.1, fill: false, pointBackgroundColor: '#000' } ] } };
    if (type === 'ojiva') return { type: 'line', data: { labels, datasets: [{ label: 'Ojiva Hi %', data: ds.classesData.map(c => c.Hi * 100), borderColor: '#000', borderWidth: 2, fill: true, backgroundColor: 'rgba(0, 0, 0, 0.05)', tension: 0.3 }] } };
    if (type === 'box') return { type: 'boxplot', data: { labels: ['Distribución'], datasets: [{ label: ds.name, data: [ds.data], backgroundColor: 'rgba(0, 0, 0, 0.1)', borderColor: '#000', borderWidth: 2, itemRadius: 3, outlierBackgroundColor: '#000' }] } };
}

// Renderiza el gráfico en un canvas invisible para garantizar 100% de calidad
async function captureChartBase64(dataset, chartType) {
    return new Promise((resolve) => {
        const container = document.createElement('div');
        container.style.width = '1000px';
        container.style.height = '500px';
        container.style.position = 'absolute';
        container.style.left = '-9999px'; // Lo escondemos fuera de la pantalla
        document.body.appendChild(container);

        const canvas = document.createElement('canvas');
        container.appendChild(canvas);
        
        const config = getChartConfig(dataset, chartType);
        config.options = {
            responsive: false,
            animation: false, // Fundamental: desactiva la animación para capturar al instante
            plugins: { legend: { display: false } }
        };
        
        if(chartType === 'box') config.options.indexAxis = 'y';
        if(chartType === 'hist') config.options.scales = { y: { beginAtZero: true } };
        if(chartType === 'ojiva') config.options.scales = { y: { beginAtZero: true, max: 100 } };
        
        const chart = new Chart(canvas, config);
        
        // Damos un pequeñísimo margen para que el navegador pinte el canvas
        setTimeout(() => {
            const base64 = canvas.toDataURL('image/png', 1.0);
            chart.destroy();
            document.body.removeChild(container);
            resolve(base64);
        }, 150);
    });
}

export async function exportToPDF(dataset) {
    const btn = document.getElementById('exportPdfBtn');
    const originalText = btn.innerText;
    btn.innerText = 'Generando Reporte...';
    btn.disabled = true;

    try {
        // 1. Tomar capturas perfectas en segundo plano
        const imgHist = await captureChartBase64(dataset, 'hist');
        const imgOjiva = await captureChartBase64(dataset, 'ojiva');
        const imgBox = await captureChartBase64(dataset, 'box');

        // 2. Inicializar jsPDF
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        
        // TÍTULO
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(18);
        doc.text(`REPORTE ESTADÍSTICO: ${dataset.name.toUpperCase()}`, 105, 20, { align: 'center' });
        
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(10);
        doc.setTextColor(100);
        doc.text('Generado por el Generador de Tablas de Frecuencia', 105, 26, { align: 'center' });
        doc.setTextColor(0);

        // 3. TABLA DE FRECUENCIAS (Usando AutoTable)
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text(`1. Tabla de Frecuencias ${dataset.isGrouped ? '(Datos Agrupados)' : '(Frecuencias Simples)'}`, 14, 40);

        const tableHead = dataset.isGrouped 
            ? [['Límite Inf.', 'Límite Sup.', 'Marca Clase', 'Frec. Abs.', 'Frec. Acum.', 'Frec. Rel.', 'Frec. Rel. Acum.']]
            : [['Dato (Xi)', 'Frec. Abs. (fi)', 'Frec. Acum. (Fi)', 'Frec. Rel. (hi)', 'Frec. Rel. Acum. (Hi)']];
            
        const tableBody = dataset.classesData.map(c => {
            if (dataset.isGrouped) {
                return [cleanNum(c.min), cleanNum(c.max), cleanNum(c.xi), c.fi, c.Fi, cleanNum(c.hi), cleanNum(c.Hi)];
            } else {
                return [cleanNum(c.xi), c.fi, c.Fi, cleanNum(c.hi), cleanNum(c.Hi)];
            }
        });

        doc.autoTable({
            startY: 45,
            head: tableHead,
            body: tableBody,
            theme: 'grid',
            headStyles: { fillColor: [40, 40, 40], textColor: [255, 255, 255], halign: 'center' },
            bodyStyles: { halign: 'center' },
            margin: { left: 14, right: 14 }
        });

        // 4. MEDIDAS ESTADÍSTICAS (Usando AutoTable en formato columna)
        let finalY = doc.lastAutoTable.finalY + 15;
        if (finalY > 230) { doc.addPage(); finalY = 20; }

        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text('2. Medidas Estadísticas', 14, finalY);
        
        const statsData = [
            ['Total de datos (n):', dataset.n, 'Varianza:', cleanNum(dataset.stats.variance)],
            ['Mínimo:', cleanNum(dataset.minVal), 'Desviación Estándar:', cleanNum(dataset.stats.stdDev)],
            ['Máximo:', cleanNum(dataset.maxVal), 'Coeficiente Variación:', `${cleanNum(dataset.stats.cv, 2)}%`],
            ['Rango:', cleanNum(dataset.range), 'Asimetría:', cleanNum(dataset.stats.skewness)],
            ['Media Aritmética:', cleanNum(dataset.stats.mean), 'Cuartil 1 (Q1):', cleanNum(dataset.stats.q1)],
            ['Mediana:', cleanNum(dataset.stats.median), 'Cuartil 2 (Q2):', cleanNum(dataset.stats.q2)],
            ['Moda:', dataset.stats.mode.map(m=>cleanNum(m)).join(', '), 'Cuartil 3 (Q3):', cleanNum(dataset.stats.q3)]
        ];

        doc.autoTable({
            startY: finalY + 5,
            body: statsData,
            theme: 'plain',
            styles: { cellPadding: 2, fontSize: 10 },
            columnStyles: { 
                0: { fontStyle: 'bold', halign: 'right', cellWidth: 45 }, 
                1: { halign: 'left', cellWidth: 45 },
                2: { fontStyle: 'bold', halign: 'right', cellWidth: 45 },
                3: { halign: 'left', cellWidth: 45 }
            }
        });

        // 5. INYECCIÓN DE GRÁFICOS
        doc.addPage();
        let currentY = 20;
        doc.setFontSize(12);
        doc.setFont('helvetica', 'bold');
        doc.text('3. Gráficos Estadísticos', 14, currentY);
        currentY += 15;

        // Inyectar Histograma
        doc.setFontSize(10);
        doc.setFont('helvetica', 'normal');
        doc.text('Histograma y Polígono de Frecuencias', 105, currentY, { align: 'center' });
        doc.addImage(imgHist, 'PNG', 15, currentY + 5, 180, 90);
        currentY += 110;

        // Inyectar Ojiva
        if (currentY > 180) { doc.addPage(); currentY = 20; }
        doc.text('Ojiva (Menor que)', 105, currentY, { align: 'center' });
        doc.addImage(imgOjiva, 'PNG', 15, currentY + 5, 180, 90);
        currentY += 110;

        // Inyectar Boxplot
        if (currentY > 180) { doc.addPage(); currentY = 20; }
        doc.text('Diagrama de Caja y Bigotes', 105, currentY, { align: 'center' });
        doc.addImage(imgBox, 'PNG', 15, currentY + 5, 180, 90);

        // Guardar Archivo
        doc.save(`Reporte_Estadistico_${dataset.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.pdf`);

    } catch (error) {
        console.error("Error generando PDF:", error);
        alert("Hubo un error al generar el PDF. Revisa la consola.");
    } finally {
        btn.innerText = originalText;
        btn.disabled = false;
    }
}