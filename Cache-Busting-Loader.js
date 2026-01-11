// Updated by GitHub Actions
const LATEST_VERSION = '{{VERSION}}';

(function() {
    const storageKey = 'gsheets_table_version';
    const lastVersion = localStorage.getItem(storageKey);
    
    // Check if we need to bust cache
    const needsBust = LATEST_VERSION !== lastVersion;
    const versionSuffix = (LATEST_VERSION !== '{{VERSION}}' && LATEST_VERSION) ? '?v=' + LATEST_VERSION : '';
    
    if (needsBust && LATEST_VERSION !== '{{VERSION}}') {
        localStorage.setItem(storageKey, LATEST_VERSION);
    }

    // 1. Load Root CSS (Variables) first - Insert at the very beginning of <head>
    const rootCSS = document.createElement('link');
    rootCSS.rel = 'stylesheet';
    rootCSS.href = 'GSheets-To-HTML-Table-root.css' + versionSuffix;
    document.head.insertBefore(rootCSS, document.head.firstChild);

    // 2. Load Rules CSS (Uses Variables) - Insert at the end of <head> to follow overrides
    const rulesCSS = document.createElement('link');
    rulesCSS.rel = 'stylesheet';
    rulesCSS.href = 'GSheets-To-HTML-Table-rules.css' + versionSuffix;
    document.head.appendChild(rulesCSS);

    // 3. Load Main JS Logic
    const mainJS = document.createElement('script');
    mainJS.src = 'GSheets-To-HTML-Table.js' + versionSuffix;
    mainJS.defer = true;
    document.head.appendChild(mainJS);
})();
