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

    // 1. Root CSS (Variables)
    const rootCSS = document.createElement('link');
    rootCSS.rel = 'stylesheet';
    rootCSS.href = 'https://automybiz.github.io/GSheets-To-HTML-Table/latest/GSheets-To-HTML-Table-root.css' + versionSuffix;
    
    // 2. Rules CSS (Layout/Components)
    const rulesCSS = document.createElement('link');
    rulesCSS.rel = 'stylesheet';
    rulesCSS.href = 'https://automybiz.github.io/GSheets-To-HTML-Table/latest/GSheets-To-HTML-Table-rules.css' + versionSuffix;

    // 3. Main JS Logic
    const mainJS = document.createElement('script');
    mainJS.src = 'https://automybiz.github.io/GSheets-To-HTML-Table/latest/GSheets-To-HTML-Table.js' + versionSuffix;
    mainJS.defer = true;

    // ROGUE SANDWICH LOGIC 🥪
    if (overridesTag) {
        // Guarantee: root.css < Overrides < rules.css < main.js
        // 1. Insert root.css BEFORE overrides
        overridesTag.insertAdjacentElement('beforebegin', rootCSS);
        
        // 2. Insert rules.css AFTER overrides
        overridesTag.insertAdjacentElement('afterend', rulesCSS);
        
        // 3. Insert main.js AFTER rules.css
        rulesCSS.insertAdjacentElement('afterend', mainJS);
        
        console.log('[GSheets-To-HTML-Table] Rogue Sandwich assembled! root.css -> Overrides -> rules.css -> main.js');
    } else if (configTag) {
        // Fallback to stacking after config tag
        configTag.insertAdjacentElement('afterend', mainJS);
        configTag.insertAdjacentElement('afterend', rulesCSS);
        configTag.insertAdjacentElement('afterend', rootCSS);
        console.warn('[GSheets-To-HTML-Table] Overrides tag not found. Stacking after config tag.');
    } else {
        // Final fallback to head (backwards compatibility)
        document.head.appendChild(rootCSS);
        document.head.appendChild(rulesCSS);
        document.head.appendChild(mainJS);
        console.warn('[GSheets-To-HTML-Table] No anchor tags found. Falling back to <head> injection.');
    }
})();
