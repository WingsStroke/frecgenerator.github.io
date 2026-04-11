export function extractNumbersFromFile(file) {
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

export async function exportAllToExcel(globalDatasets, activeMethod) {
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