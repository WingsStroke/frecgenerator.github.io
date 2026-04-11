export const AppState = {
    analysisMode: 'univariate', 
    activeMethod: 'sturges',
    globalDatasets: [],
    uploadedFilesMap: new Map(),
    MAX_DATASETS: 10,
    currentSlide: 0,
    currentPreviewFileId: null,
    tempBivariateX: [],
    tempBivariateY: []
};