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

    // Reference tag for "Insert After" logic
    const configTag = document.getElementById('GSheets-To-HTML-Table-Config');

    // 1. Load Root CSS (Variables)
    const rootCSS = document.createElement('link');
    rootCSS.rel = 'stylesheet';
    rootCSS.href = 'https://automybiz.github.io/GSheets-To-HTML-Table/latest/GSheets-To-HTML-Table-root.css' + versionSuffix;
    
    // 2. Load Rules CSS (Uses Variables)
    const rulesCSS = document.createElement('link');
    rulesCSS.rel = 'stylesheet';
    rulesCSS.href = 'https://automybiz.github.io/GSheets-To-HTML-Table/latest/GSheets-To-HTML-Table-rules.css' + versionSuffix;

    // 3. Load Main JS Logic
    const mainJS = document.createElement('script');
    mainJS.src = 'https://automybiz.github.io/GSheets-To-HTML-Table/latest/GSheets-To-HTML-Table.js' + versionSuffix;
    mainJS.defer = true;

    if (configTag) {
        // Use "Insert After" logic to guarantee execution order
        configTag.insertAdjacentElement('afterend', mainJS);
        configTag.insertAdjacentElement('afterend', rulesCSS);
        configTag.insertAdjacentElement('afterend', rootCSS);
        console.log('[GSheets-To-HTML-Table] Config tag found, injected scripts immediately after it.');
    } else {
        // Fallback to head injection if config tag not found (backwards compatibility)
        document.head.appendChild(rootCSS);
        document.head.appendChild(rulesCSS);
        document.head.appendChild(mainJS);
        console.warn('[GSheets-To-HTML-Table] WARNING: Script tag with ID "GSheets-To-HTML-Table-Config" not found. Falling back to <head> injection.');
    }
})();
