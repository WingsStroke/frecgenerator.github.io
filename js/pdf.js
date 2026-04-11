import { cleanNum } from './math.js';

export async function exportToPDF(dataset, slideIndex) {
    // Creamos un contenedor virtual con un ancho fijo para evitar deformaciones
    const element = document.createElement('div');
    element.style.padding = '20px';
    element.style.fontFamily = 'Arial, Helvetica, sans-serif';
    element.style.color = '#000';
    element.style.backgroundColor = '#fff';
    element.style.width = '800px'; 
    element.style.margin = '0 auto';

    let html = `
        <div style="text-align: center; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px;">
            <h1 style="text-transform: uppercase; font-size: 22px; margin: 0; color: #000;">Reporte Estadístico: ${dataset.name}</h1>
            <p style="color: #555; margin: 5px 0 0 0; font-size: 12px;">Generado por el Generador de Tablas de Frecuencia</p>
        </div>
        
        <h3 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; font-size: 16px; margin-top: 0;">1. Tabla de Frecuencias ${dataset.isGrouped ? '(Agrupados)' : '(Simples)'}</h3>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; text-align: center; font-size: 12px;">
            <thead>
                <tr style="background-color: #f0f0f0; border: 1px solid #000; page-break-inside: avoid;">
    `;
    
    if (dataset.isGrouped) {
        html += `<th style="padding: 6px; border: 1px solid #000;">Límite Inf. (Li)</th><th style="padding: 6px; border: 1px solid #000;">Límite Sup. (Ls)</th><th style="padding: 6px; border: 1px solid #000;">Marca Clase (Xi)</th>`;
    } else {
        html += `<th style="padding: 6px; border: 1px solid #000;">Dato (Xi)</th>`;
    }
    
    html += `<th style="padding: 6px; border: 1px solid #000;">Frec. Abs. (fi)</th><th style="padding: 6px; border: 1px solid #000;">Frec. Acum. (Fi)</th><th style="padding: 6px; border: 1px solid #000;">Frec. Rel. (hi)</th><th style="padding: 6px; border: 1px solid #000;">Frec. Rel. Acum. (Hi)</th></tr></thead><tbody>`;
    
    dataset.classesData.forEach(c => {
        // Evitamos que una fila de la tabla se corte a la mitad entre dos páginas
        html += `<tr style="page-break-inside: avoid;">`;
        if (dataset.isGrouped) {
            html += `<td style="padding: 5px; border: 1px solid #000;">${cleanNum(c.min)}</td><td style="padding: 5px; border: 1px solid #000;">${cleanNum(c.max)}</td>`;
        }
        html += `<td style="padding: 5px; border: 1px solid #000;">${cleanNum(c.xi)}</td><td style="padding: 5px; border: 1px solid #000;">${c.fi}</td><td style="padding: 5px; border: 1px solid #000;">${c.Fi}</td><td style="padding: 5px; border: 1px solid #000;">${cleanNum(c.hi)}</td><td style="padding: 5px; border: 1px solid #000;">${cleanNum(c.Hi)}</td></tr>`;
    });
    html += `</tbody></table>`;

    // Usamos una estructura de <table> tradicional en lugar de flexbox para las estadísticas
    // Esto garantiza un renderizado perfecto y alineado en PDF
    html += `
        <div style="page-break-inside: avoid; margin-bottom: 25px;">
            <h3 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; font-size: 16px; margin-top: 0;">2. Medidas Estadísticas</h3>
            <table style="width: 100%; border: 1px solid #000; border-collapse: collapse; font-size: 13px; text-align: left;">
                <tr>
                    <td style="padding: 15px; width: 50%; vertical-align: top; border-right: 1px solid #000;">
                        <p style="margin:4px 0;"><b>Total de datos (n):</b> ${dataset.n}</p>
                        <p style="margin:4px 0;"><b>Mínimo:</b> ${cleanNum(dataset.minVal)}</p>
                        <p style="margin:4px 0;"><b>Máximo:</b> ${cleanNum(dataset.maxVal)}</p>
                        <p style="margin:4px 0;"><b>Rango:</b> ${cleanNum(dataset.range)}</p>
                        <hr style="border: 0; border-top: 1px dashed #ccc; margin: 12px 0;">
                        <p style="margin:4px 0;"><b>Media Aritmética:</b> ${cleanNum(dataset.stats.mean)}</p>
                        <p style="margin:4px 0;"><b>Mediana:</b> ${cleanNum(dataset.stats.median)}</p>
                        <p style="margin:4px 0;"><b>Moda:</b> ${dataset.stats.mode.map(m=>cleanNum(m)).join(', ')}</p>
                    </td>
                    <td style="padding: 15px; width: 50%; vertical-align: top;">
                        <p style="margin:4px 0;"><b>Varianza:</b> ${cleanNum(dataset.stats.variance)}</p>
                        <p style="margin:4px 0;"><b>Desviación Estándar:</b> ${cleanNum(dataset.stats.stdDev)}</p>
                        <p style="margin:4px 0;"><b>Coeficiente Variación:</b> ${cleanNum(dataset.stats.cv, 2)}%</p>
                        <p style="margin:4px 0;"><b>Asimetría:</b> ${cleanNum(dataset.stats.skewness)}</p>
                        <hr style="border: 0; border-top: 1px dashed #ccc; margin: 12px 0;">
                        <p style="margin:4px 0;"><b>Cuartil 1 (Q1):</b> ${cleanNum(dataset.stats.q1)}</p>
                        <p style="margin:4px 0;"><b>Cuartil 2 (Q2):</b> ${cleanNum(dataset.stats.q2)}</p>
                        <p style="margin:4px 0;"><b>Cuartil 3 (Q3):</b> ${cleanNum(dataset.stats.q3)}</p>
                    </td>
                </tr>
            </table>
        </div>
    `;

    const histCanvas = document.getElementById(`chartHist-${slideIndex}`);
    const ojivaCanvas = document.getElementById(`chartOjiva-${slideIndex}`);
    const boxCanvas = document.getElementById(`chartBox-${slideIndex}`);

    html += `<h3 style="border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-bottom: 15px; font-size: 16px;">3. Gráficos Estadísticos</h3>`;
    html += `<div style="text-align: center;">`;
    
    // Encapsulamos CADA gráfico en un div "avoid" para que fluyan naturalmente
    // Restringimos estrictamente la altura máxima a 240px
    if (histCanvas) {
        html += `<div style="page-break-inside: avoid; margin-bottom: 25px;">
                    <h4 style="margin: 0 0 8px 0; font-size: 14px; text-transform: uppercase; color: #333;">Histograma y Polígono de Frecuencias</h4>
                    <img src="${histCanvas.toDataURL('image/png', 1.0)}" style="width: 85%; max-height: 240px; object-fit: contain; border: 1px solid #ddd; padding: 10px;">
                 </div>`;
    }
    if (ojivaCanvas) {
        html += `<div style="page-break-inside: avoid; margin-bottom: 25px;">
                    <h4 style="margin: 0 0 8px 0; font-size: 14px; text-transform: uppercase; color: #333;">Ojiva (Menor que)</h4>
                    <img src="${ojivaCanvas.toDataURL('image/png', 1.0)}" style="width: 85%; max-height: 240px; object-fit: contain; border: 1px solid #ddd; padding: 10px;">
                 </div>`;
    }
    if (boxCanvas) {
        html += `<div style="page-break-inside: avoid; margin-bottom: 25px;">
                    <h4 style="margin: 0 0 8px 0; font-size: 14px; text-transform: uppercase; color: #333;">Diagrama de Caja y Bigotes</h4>
                    <img src="${boxCanvas.toDataURL('image/png', 1.0)}" style="width: 85%; max-height: 240px; object-fit: contain; border: 1px solid #ddd; padding: 10px;">
                 </div>`;
    }
    
    html += `</div>`;
    element.innerHTML = html;

    // Configuración robusta para html2pdf
    const opt = {
        margin:       [15, 10, 15, 10], // Margen: [arriba, derecha, abajo, izquierda]
        filename:     `Reporte_Estadistico_${dataset.name.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { 
            scale: 2, 
            useCORS: true, 
            windowWidth: 800 // Forzamos un renderizado de 800px de ancho real
        },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    const btn = document.getElementById('exportPdfBtn');
    const originalText = btn.innerText;
    btn.innerText = 'Generando Reporte...';
    btn.disabled = true;

    html2pdf().set(opt).from(element).save().then(() => {
        btn.innerText = originalText;
        btn.disabled = false;
    });
}