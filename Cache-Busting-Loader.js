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

    // Anchor tags for precise injection
    const configTag = document.getElementById('GSheets-To-HTML-Table-Config');
    const overridesTag = document.getElementById('GSheets-To-HTML-Table-Overrides');

    // Determine environment (Local vs Production)
    const isLocal = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
    const baseUrl = isLocal ? '' : 'https://automybiz.github.io/GSheets-To-HTML-Table/latest/';

    // 1. Root CSS (Variables)
    const rootCSS = document.createElement('link');
    rootCSS.rel = 'stylesheet';
    rootCSS.href = baseUrl + 'GSheets-To-HTML-Table-root.css' + versionSuffix;
    
    // 2. Rules CSS (Layout/Components)
    const rulesCSS = document.createElement('link');
    rulesCSS.rel = 'stylesheet';
    rulesCSS.href = baseUrl + 'GSheets-To-HTML-Table-rules.css' + versionSuffix;

    // 3. Main JS Logic
    const mainJS = document.createElement('script');
    mainJS.src = baseUrl + 'GSheets-To-HTML-Table.js' + versionSuffix;
    mainJS.defer = true;

    // ROGUE SANDWICH LOGIC 🥪
    
    // Handle CSS first (Sandwich your overrides)
    if (overridesTag) {
        // Guarantee: root.css < Overrides < rules.css
        overridesTag.insertAdjacentElement('beforebegin', rootCSS);
        overridesTag.insertAdjacentElement('afterend', rulesCSS);
        console.log('[GSheets-To-HTML-Table] CSS Sandwich assembled! root.css -> Overrides -> rules.css');
    } else if (configTag) {
        // Fallback for CSS
        configTag.insertAdjacentElement('afterend', rulesCSS);
        configTag.insertAdjacentElement('afterend', rootCSS);
    } else {
        document.head.appendChild(rootCSS);
        document.head.appendChild(rulesCSS);
    }

    // Handle JS separately (Anchor to config for logical DOM order)
    if (configTag) {
        // Guarantee: Config < Main JS
        configTag.insertAdjacentElement('afterend', mainJS);
        console.log('[GSheets-To-HTML-Table] Main JS anchored after Config tag.');
    } else if (overridesTag) {
        // Fallback to overrides tag for JS
        overridesTag.insertAdjacentElement('afterend', mainJS);
    } else {
        // Final fallback for JS
        document.head.appendChild(mainJS);
    }
})();
