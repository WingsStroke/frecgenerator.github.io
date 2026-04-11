export const cleanNum = (num, decimals = 4) => {
    if (isNaN(num)) return 0;
    const fixed = parseFloat(num.toFixed(decimals));
    return Number.isInteger(fixed) ? fixed : fixed;
};

export const getPercentile = (data, p) => {
    const n = data.length;
    const idx = (p / 100) * (n - 1);
    const l = Math.floor(idx);
    return l + 1 >= n ? data[l] : data[l] * (1 - (idx % 1)) + data[l + 1] * (idx % 1);
};

export function calculateStatsForDataset(raw, datasetName, activeMethod, manualKValue) {
    let data = [...raw].sort((a, b) => a - b);
    const n = data.length;
    const minVal = data[0];
    const maxVal = data[n - 1];
    const range = maxVal - minVal;
    
    let numClasses = activeMethod === 'manual' ? parseInt(manualKValue) : Math.round(1 + 3.322 * Math.log10(n));
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

    const varianceSum = data.reduce((a, b) => a + Math.pow(b - mean, 2), 0); 
    const variance = varianceSum / (n - 1);
    const stdDev = Math.sqrt(variance);
    const cv = (stdDev / mean) * 100;

    let skewness = 0;
    if (n > 2 && stdDev > 0) skewness = (n / ((n - 1) * (n - 2))) * data.reduce((acc, val) => acc + Math.pow((val - mean) / stdDev, 3), 0);

    return { 
        name: datasetName, data, n, minVal, maxVal, range, numClasses, amplitude, classesData, 
        stats: { sum, mean, geoMean, harMean, median, mode, varianceSum, variance, stdDev, cv, skewness, p10: getPercentile(data,10), q1: getPercentile(data,25), q2: getPercentile(data,50), q3: getPercentile(data,75), p90: getPercentile(data,90) }
    };
}