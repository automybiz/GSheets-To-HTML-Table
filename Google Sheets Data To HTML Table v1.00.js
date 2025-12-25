    (function() {
    
    // Generate unique instance ID
    const INSTANCE_ID = 'accordion_' + Math.random().toString(36).substr(2, 9);
    console.log('[Accordion] Instance ID:', INSTANCE_ID);
    
    // Find the wrapper (last one with class accordion-wrapper)
    const wrappers = document.querySelectorAll('.accordion-wrapper');
    const wrapper = wrappers[wrappers.length - 1];
    
    if (!wrapper || wrapper.dataset.initialized) {
        console.log('[Accordion] Wrapper not found or already initialized');
        return;
    }
    
    wrapper.dataset.initialized = 'true';
    wrapper.id = INSTANCE_ID + '-wrapper';
    
    // Create HTML structure
    const searchBoxHTML = CONFIG.SHOW_SEARCH_BOX ? `
        <input 
            type="text" 
            id="${INSTANCE_ID}-search" 
            placeholder="${CONFIG.SEARCH_PLACEHOLDER}" 
            class="accordion-search-input"
            autocomplete="off"
            data-scope="${CONFIG.SEARCH_SCOPE}"
        >
    ` : '';
    
    // Create common searches HTML if enabled
    const commonSearchesHTML = CONFIG.SEARCHES_COMMON && CONFIG.SEARCHES_COMMON.length > 0 ? `
        <span class="text-no-background"></span> <div class="accordion-common-searches">
            ${CONFIG.SEARCHES_COMMON.map(term => 
                `<span class="accordion-common-search-item" onclick="window.setSearch_${INSTANCE_ID}('${term}')">${term}</span>`
            ).join('')}
        </div>
    ` : '';
    
    wrapper.innerHTML = `
        ${searchBoxHTML}
        ${commonSearchesHTML}
        <div id="${INSTANCE_ID}-content">
            <div class="loading-message">
                <div class="spinner"></div>
                <p>Loading data...</p>
            </div>
        </div>
    `;
    
    // Store the instance ID globally for this specific accordion
    wrapper.accordionInstanceId = INSTANCE_ID;
    
    // ============================================
    // STATE
    // ============================================
    let allData = [];
    let headers = [];
    let searchTimeout = null;
    let isFirstDataRow = true;
    let retryTimeout = null;
    let countdownInterval = null;
    let isPaused = false;
    
    // Animation settings mapping
    const transitionEffects = {
        'smooth': 'ease',
        'ease-in': 'ease-in',
        'ease-out': 'ease-out',
        'bounce': 'cubic-bezier(0.68, -0.55, 0.265, 1.55)',
        'snap': 'cubic-bezier(0.95, 0.05, 0.795, 0.035)'
    };
    
    // ============================================
    // UTILITY FUNCTIONS
    // ============================================
    function columnLetterToIndex(letter) {
        return letter.toUpperCase().charCodeAt(0) - 65;
    }
    
    function getQuestionColumnIndices() {
        return CONFIG.QUESTION_COLUMNS.map(col => columnLetterToIndex(col));
    }
    
    function getAnswerColumnIndex() {
        return columnLetterToIndex(CONFIG.ANSWER_COLUMN);
    }

    function getUrlColumnIndices() {
        if (!CONFIG.URL_COLUMNS || !Array.isArray(CONFIG.URL_COLUMNS)) return [];
        return CONFIG.URL_COLUMNS.map(col => columnLetterToIndex(col));
    }
    
    function convertNewlinesToBR(text) {
        if (!text) return text;
        return text.replace(/\r\n|\r|\n/g, '<br>');
    }
    
    function preserveWhitespace(text) {
        if (!text) return text;
        text = text.replace(/ {2,}/g, match => '&nbsp;'.repeat(match.length));
        text = text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;');
        return text;
    }
    
    function isImageURL(url) {
        if (!url || typeof url !== 'string') return false;
        
        const trimmedUrl = url.trim();
        
        // Check if starts with http/https
        if (!trimmedUrl.startsWith('http://') && !trimmedUrl.startsWith('https://')) {
            return false;
        }
        
        // Check if ends with image extension
        const imageExtensions = /\.(jpg|jpeg|png|gif|bmp|webp|svg)(\?[^\s]*)?$/i;
        if (imageExtensions.test(trimmedUrl)) {
            return true;
        }
        
        // Check for known image hosting domains
        const googleStoragePattern = /https:\/\/storage\.googleapis\.com\//i;
        const googleDrivePattern = /https:\/\/drive\.google\.com\//i;
        const googleUserContentPattern = /https:\/\/lh[0-9]+\.googleusercontent\.com\//i;
        
        return googleStoragePattern.test(trimmedUrl) || 
               googleDrivePattern.test(trimmedUrl) ||
               googleUserContentPattern.test(trimmedUrl);
    }
    
    function isDirectImageURL(text) {
        if (!text || typeof text !== 'string') return false;
        
        const trimmedText = text.trim();
        
        // Check if the entire cell content is just a URL (no other text)
        const urlOnlyPattern = /^(https?:\/\/[^\s]+)$/;
        const match = trimmedText.match(urlOnlyPattern);
        
        if (match) {
            const url = match[1];
            // Check if starts with http and ends with image extension
            if ((url.startsWith('http://') || url.startsWith('https://')) && isImageURL(url)) {
                console.log('[Accordion] Direct image URL detected:', url);
                return url;
            }
        }
        
        return false;
    }
    
    function generateImageTag(imageUrl, isInCell = false) {
        let styles = [];
        
        if (CONFIG.IMAGE_MAX_WIDTH) {
            styles.push(`max-width: ${CONFIG.IMAGE_MAX_WIDTH}px`);
        }
        
        if (CONFIG.IMAGE_MAX_HEIGHT) {
            styles.push(`max-height: ${CONFIG.IMAGE_MAX_HEIGHT}px`);
        }
        
        if (CONFIG.IMAGE_MAINTAIN_ASPECT_RATIO) {
            styles.push('width: auto');
            styles.push('height: auto');
            styles.push('object-fit: contain');
        } else {
            if (CONFIG.IMAGE_MAX_WIDTH) styles.push('width: 100%');
            if (CONFIG.IMAGE_MAX_HEIGHT) styles.push(`height: ${CONFIG.IMAGE_MAX_HEIGHT}px`);
            styles.push('object-fit: fill');
        }
        
        // Center the image and remove extra space
        styles.push('display: block');
        styles.push('margin: 0 auto');
        
        // Add class to identify image cells for padding removal
        const className = isInCell ? 'class="accordion-image-content"' : '';
        
        return `<img src="${imageUrl}" ${className} style="${styles.join('; ')}" alt="Image" loading="lazy">`;
    }
    
    function extractImageFromCell(text) {
        if (!text) return null;
        
        const imageFormulaMatch = text.match(/=IMAGE\s*\(\s*["']([^"']+)["']/i);
        if (imageFormulaMatch) {
            console.log('[Accordion] Found IMAGE formula:', imageFormulaMatch[1]);
            return imageFormulaMatch[1];
        }
        
        const urlPattern = /(https?:\/\/[^\s<"']+)/;
        const match = text.match(urlPattern);
        if (match) {
            const url = match[1];
            if (isImageURL(url)) {
                console.log('[Accordion] Found image URL:', url);
                return url;
            }
        }
        
        return null;
    }
    
    function convertURLsToLinks(text, isInCell = false) {
        if (!text) return '';
        
        console.log('[Accordion] Processing cell content:', text.substring(0, 100));
        
        // First check: Is the entire cell just an image URL?
        const directImageUrl = isDirectImageURL(text);
        if (directImageUrl) {
            console.log('[Accordion] Rendering direct image URL:', directImageUrl);
            return generateImageTag(directImageUrl, isInCell);
        }
        
        // Second check: Is it an =IMAGE() formula?
        const imageUrl = extractImageFromCell(text);
        if (imageUrl) {
            console.log('[Accordion] Rendering IMAGE formula:', imageUrl);
            return generateImageTag(imageUrl, isInCell);
        }
        
        let result = preserveWhitespace(text);
        result = convertNewlinesToBR(result);
        
        const urlPattern = /(https?:\/\/[^\s<]+)/g;
        const urls = text.match(urlPattern) || [];
        
        urls.forEach(url => {
            let replacement = url;
            let isProcessed = false;
            
            if (url.includes('youtube.com') || url.includes('youtu.be')) {
                if (url === 'https://www.youtube.com' || url === 'https://www.youtube.com/' || 
                    url === 'https://youtube.com' || url === 'https://youtube.com/' ||
                    url === 'https://youtu.be' || url === 'https://youtu.be/') {
                    replacement = `<a href="${url}" target="_blank">${url}</a>`;
                    isProcessed = true;
                } else {
                    let videoId = null;
                    let timeParam = '';
                    let playlistParam = '';
                    
                    let match = url.match(/youtube\.com\/watch\?v=([a-zA-Z0-9_-]+)/);
                    if (match) {
                        videoId = match[1];
                        const tMatch = url.match(/[?&]t=(\d+)/);
                        if (tMatch) timeParam = `?start=${tMatch[1]}`;
                        const listMatch = url.match(/[?&]list=([a-zA-Z0-9_-]+)/);
                        if (listMatch) playlistParam = `&list=${listMatch[1]}`;
                    }
                    
                    if (!videoId) {
                        match = url.match(/youtu\.be\/([a-zA-Z0-9_-]+)/);
                        if (match) {
                            videoId = match[1];
                            const tMatch = url.match(/[?&]t=(\d+)/);
                            if (tMatch) timeParam = `?start=${tMatch[1]}`;
                        }
                    }
                    
                    if (videoId) {
                        replacement = `<iframe width="${CONFIG.YOUTUBE_EMBED_WIDTH}" height="${CONFIG.YOUTUBE_EMBED_HEIGHT}" src="https://www.youtube.com/embed/${videoId}${timeParam}${playlistParam}" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>`;
                        isProcessed = true;
                    }
                }
            }
            
            if (!isProcessed && isImageURL(url)) {
                replacement = generateImageTag(url, isInCell);
                isProcessed = true;
            }
            
            if (!isProcessed) {
                replacement = `<a href="${url}" target="_blank">${url}</a>`;
            }
            
            result = result.replace(url, replacement);
        });
        
        return result;
    }
    
    function animatePlaceholderText(inputElement, fullText) {
        let i = 0;
        let animationInterval;
    
        function typeLetter() {
            if (inputElement !== document.activeElement) { // Only animate if input is not focused
                if (i < fullText.length) {
                    inputElement.setAttribute('placeholder', fullText.substring(0, i + 1));
                    i++;
                } else {
                    // Loop the animation after a delay
                    clearInterval(animationInterval);
                    setTimeout(() => {
                        i = 0;
                        inputElement.setAttribute('placeholder', '');
                        animationInterval = setInterval(typeLetter, CONFIG.SEARCH_PLACEHOLDER_ANIMATION_SPEED);
                    }, CONFIG.SEARCH_PLACEHOLDER_ANIMATION_LOOP_DELAY);
                }
            }
        }
    
        // Start the animation
        animationInterval = setInterval(typeLetter, CONFIG.SEARCH_PLACEHOLDER_ANIMATION_SPEED);
    
        // Stop animation when input is focused
        inputElement.addEventListener('focus', () => {
            clearInterval(animationInterval);
            inputElement.setAttribute('placeholder', fullText);
        });
    
        // Restart animation when input is blurred if no text is entered
        inputElement.addEventListener('blur', () => {
            if (inputElement.value === '') {
                i = 0;
                inputElement.setAttribute('placeholder', '');
                animationInterval = setInterval(typeLetter, CONFIG.SEARCH_PLACEHOLDER_ANIMATION_SPEED);
            }
        });
    }
    
    // ============================================
    // SEARCH FUNCTIONALITY
    // ============================================
    function initializeSearch() {
        const searchInput = document.getElementById(INSTANCE_ID + '-search');
        
        if (searchInput) {
            const searchScope = searchInput.dataset.scope;
            
            searchInput.addEventListener('input', (e) => {
                clearTimeout(searchTimeout);
                
                searchTimeout = setTimeout(() => {
                    if (searchScope === 'all') {
                        performGlobalSearch(e.target.value);
                    } else {
                        performSearch(e.target.value);
                    }
                }, CONFIG.SEARCH_DELAY_MS);
            });
            // Animate placeholder on load
            animatePlaceholderText(searchInput, CONFIG.SEARCH_PLACEHOLDER);
            console.log('[Accordion] Search initialized with scope:', searchScope);
        }
    }
    
    window['setSearch_' + INSTANCE_ID] = function(searchTerm) {
        const searchInput = document.getElementById(INSTANCE_ID + '-search');
        if (searchInput) {
            searchInput.value = (searchTerm === 'All') ? '' : searchTerm;
            const searchScope = searchInput.dataset.scope;
            
            if (searchScope === 'all') {
                performGlobalSearch(searchInput.value);
            } else {
                performSearch(searchInput.value);
            }
        }
    };
    
    function performSearch(searchTerm) {
        const items = document.querySelectorAll('#' + INSTANCE_ID + '-content .accordion-item');
        const term = searchTerm.toLowerCase().trim();
        let visibleCount = 0;
        let lastVisibleItem = null;
    
        items.forEach((item, index) => {
            // Check if this item corresponds to the header row that should be kept visible
            if (CONFIG.HEADER_ROW_NUMBER !== 0) { // Changed from !== null to !== 0
                // Calculate the actual row number in the spreadsheet
                const actualRowNumber = CONFIG.STARTING_ROW + index;
                if (actualRowNumber === CONFIG.HEADER_ROW_NUMBER) {
                    item.classList.remove('hidden');
                    lastVisibleItem = item;
                    visibleCount++;
                    return;
                }
            }
            
            if (term === '') {
                item.classList.remove('hidden');
                // Remove highlights when search is cleared
                item.querySelectorAll('.accordion-search-highlight').forEach(el => {
                    el.outerHTML = el.textContent;
                });
                lastVisibleItem = item;
                visibleCount++;
            } else {
                const textContent = item.textContent.toLowerCase();
                let shouldBeVisible = false;
    
                // Custom search logic for "AI" to avoid matching substrings like in "main" or "campaigns"
                if (term === 'ai') {
                    // Match standalone "ai" word only
                    const aiSpecificRegex = /\b(ai)\b/i;
                    shouldBeVisible = aiSpecificRegex.test(textContent);
                } else {
                    // For all other terms, use general substring matching
                    shouldBeVisible = textContent.includes(term);
                }
    
                if (shouldBeVisible) {
                    item.classList.remove('hidden');
                    // Add highlighting to matched content
                    const cells = item.querySelectorAll('td');
                    cells.forEach(cell => {
                        if (!cell.classList.contains('accordion-auto-number-cell')) {
                            // Store original HTML if not already stored
                            if (!cell.dataset.originalHtml) {
                                cell.dataset.originalHtml = cell.innerHTML;
                            }
                            // Restore original HTML first, then highlight
                            cell.innerHTML = cell.dataset.originalHtml;
                            const highlightedHTML = highlightSearchTerms(cell.innerHTML, term);
                            cell.innerHTML = highlightedHTML;
                        }
                    });
                    lastVisibleItem = item;
                    visibleCount++;
                } else {
                    item.classList.add('hidden');
                }
            }
        });
    
        // Collapse all answers after search
        document.querySelectorAll('#' + INSTANCE_ID + '-content .accordion-item.active').forEach(item => {
            item.classList.remove('active');
        });
    
        // Remove last-visible-item class from all items
        items.forEach(item => item.classList.remove('last-visible-item'));
        
        // Add last-visible-item class to the last visible item
        if (lastVisibleItem) {
            lastVisibleItem.classList.add('last-visible-item');
        }
    
        // Re-apply alternating row colors
        applyAlternatingRowColors();
    
        const existingMessage = document.querySelector('#' + INSTANCE_ID + '-content .no-results-message');
        if (existingMessage) {
            existingMessage.remove();
        }
    
        if (visibleCount === 0 && term !== '') {
            const container = document.querySelector('#' + INSTANCE_ID + '-content .accordion-container');
            if (container) {
                const message = document.createElement('div');
                message.className = 'no-results-message';
                message.textContent = 'No results found for "' + searchTerm + '"';
                container.appendChild(message);
            }
        }
    }
    
    function escapeRegExp(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }
    
    function highlightSearchTerms(html, searchTerm) {
        if (!searchTerm || searchTerm.trim() === '') return html;
        
        const term = searchTerm.trim();
        let regex;
    
        if (term === 'ai') {
            regex = /(\bai\b)/gi; // Only highlight standalone "ai"
        } else {
            regex = new RegExp(`(${escapeRegExp(term)})`, 'gi');
        }
        
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;
        
        function highlightNode(node) {
            if (node.nodeType === 3) { // Text node
                const text = node.textContent;
                const replaced = text.replace(regex, '<span class="accordion-search-highlight">$1</span>');
                if (replaced !== text) {
                    const fragment = document.createRange().createContextualFragment(replaced);
                    node.parentNode.insertBefore(fragment, node);
                    node.parentNode.removeChild(node);
                }
            } else if (node.nodeType === 1 && node.tagName !== 'SCRIPT' && node.tagName !== 'STYLE') {
                // Recurse on a copy of childNodes to avoid live collection issues
                Array.from(node.childNodes).forEach(highlightNode);
            }
        }
        
        highlightNode(tempDiv);
        
            return tempDiv.innerHTML;
        }
    
        function applyAlternatingRowColors() {
            const visibleItems = document.querySelectorAll('#' + INSTANCE_ID + '-content .accordion-item:not(.hidden)');
            visibleItems.forEach((item, index) => {
                const questionRow = item.querySelector('.accordion-question-row');
                if (questionRow) {
                    questionRow.classList.remove('odd-row', 'even-row'); // Remove existing classes
                    if (index % 2 === 0) {
                        questionRow.classList.add('even-row');
                    } else {
                        questionRow.classList.add('odd-row');
                    }
                }
            });
        }
        
        function performGlobalSearch(searchTerm) {
        // Search across ALL accordion tables on the page
        const allItems = document.querySelectorAll('.accordion-item');
        const term = searchTerm.toLowerCase().trim();
        let totalVisibleCount = 0;
    
        // Track last visible item per container
        const containerLastVisible = new Map();
    
        allItems.forEach(item => {
            const container = item.closest('.accordion-container');
            
            if (term === '') {
                item.classList.remove('hidden');
                containerLastVisible.set(container, item);
                totalVisibleCount++;
            } else {
                const textContent = item.textContent.toLowerCase();
                if (textContent.includes(term)) {
                    item.classList.remove('hidden');
                    containerLastVisible.set(container, item);
                    totalVisibleCount++;
                } else {
                    item.classList.add('hidden');
                }
            }
        });
    
        // Collapse all answers after search
        document.querySelectorAll('.accordion-item.active').forEach(item => {
            item.classList.remove('active');
        });
    
        // Remove last-visible-item class from all items
        allItems.forEach(item => item.classList.remove('last-visible-item'));
        
        // Add last-visible-item class to last visible item in each container
        containerLastVisible.forEach((item) => {
            if (item) {
                item.classList.add('last-visible-item');
            }
        });
    
        // Remove all existing "no results" messages
        document.querySelectorAll('.no-results-message').forEach(msg => msg.remove());
    
        // Add "no results" message to each empty container
        if (totalVisibleCount === 0 && term !== '') {
            document.querySelectorAll('.accordion-container').forEach(container => {
                const visibleInContainer = Array.from(container.querySelectorAll('.accordion-item'))
                    .filter(item => !item.classList.contains('hidden')).length;
                
                if (visibleInContainer === 0) {
                    const message = document.createElement('div');
                    message.className = 'no-results-message';
                    message.textContent = 'No results found for "' + searchTerm + '"';
                    container.appendChild(message);
                }
            });
        }
        
        // Re-apply alternating row colors
        applyAlternatingRowColors();
        
        console.log('[Accordion] Global search found', totalVisibleCount, 'results');
    }
    
    // ============================================
    // DATA FETCHING
    // ============================================
    async function loadData() {
        const content = document.getElementById(INSTANCE_ID + '-content');
        if (!content) {
            console.error('[Accordion] Content element not found');
            return;
        }
        
        content.innerHTML = '<div class="loading-message"><div class="spinner"></div><p>Loading data...</p></div>';
        console.log('[Accordion] Starting data fetch');
        
        try {
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/values/${CONFIG.RANGE}?key=${CONFIG.API_KEY}`;
            const fetchUrl = CONFIG.USE_CORS_PROXY ? `https://cors-anywhere.herokuapp.com/${url}` : url;
            
            const response = await fetch(fetchUrl);
            
            if (!response.ok) {
                const errorMsg = `HTTP error! status: ${response.status}`;
                console.error('[Accordion] Fetch failed:', errorMsg);
                throw new Error(errorMsg);
            }
            
            const data = await response.json();
            console.log('[Accordion] Data received:', data.values.length, 'rows');
    
            if (!data.values || data.values.length === 0) {
                throw new Error('No data found in spreadsheet');
            }
    
            processData(data.values);
            
        } catch (error) {
            console.error('[Accordion] Error:', error);
            
            const hasMeaningfulError = error.message && 
                                       error.message !== '' && 
                                       !error.message.includes('Failed to fetch') &&
                                       !error.message.includes('NetworkError') &&
                                       !error.message.includes('Load failed');
            
            if (hasMeaningfulError) {
                showError(error.message, true);
            } else {
                showError(error.message || 'Connection issue', false);
            }
        }
    }
    
    function processData(values) {
        headers = values[0] || [];
        
        let actualStartIndex = Math.max(0, CONFIG.STARTING_ROW - 1); // 0-based index from 'values'
    
        // If HEADER_ROW_NUMBER is 0 (disabled), and STARTING_ROW is 1,
        // we assume the user intends to skip the first row (values[0])
        // as it's implicitly a header that should only be shown if HEADER_ROW_NUMBER is set.
        if (CONFIG.HEADER_ROW_NUMBER === 0 && CONFIG.STARTING_ROW === 1) {
            actualStartIndex = 1; // Effectively start data from the second row (values[1])
        }
    
        allData = values.slice(actualStartIndex);
    
        const filteredData = applyFilters(allData, CONFIG.FILTER_CONDITIONS);
        console.log('[Accordion] Filtered:', filteredData.length, 'rows');
    
        displayAccordion(filteredData);
        initializeSearch();
    }
    
    function applyFilters(data, filters) {
        if (filters.length === 0) return data;
    
        return data.filter((row, index) => {
            // Skip filtering for header row if HEADER_ROW_NUMBER is set
            if (CONFIG.HEADER_ROW_NUMBER !== 0) { // Changed from !== null to !== 0
                const actualRowNumber = CONFIG.STARTING_ROW + index;
                if (actualRowNumber === CONFIG.HEADER_ROW_NUMBER) {
                    return true; // Always include header row
                }
            }
    
            return filters.every(filter => {
                const columnIndex = filter.column.charCodeAt(0) - 65;
                const cellValue = (row[columnIndex] || '').toString().trim();
                const filterValue = (filter.value || '').toString().trim();
                
                switch (filter.condition) {
                    case 'equals':
                        return cellValue.toLowerCase() === filterValue.toLowerCase();
                    case 'does not equal':
                        return cellValue.toLowerCase() !== filterValue.toLowerCase();
                    case 'contains':
                        return cellValue.toLowerCase().includes(filterValue.toLowerCase());
                    case 'does not contain':
                        return !cellValue.toLowerCase().includes(filterValue.toLowerCase());
                    case 'starts with':
                        return cellValue.toLowerCase().startsWith(filterValue.toLowerCase());
                    case 'ends with':
                        return cellValue.toLowerCase().endsWith(filterValue.toLowerCase());
                    case 'is empty':
                        return cellValue === '';
                    case 'is not empty':
                        return cellValue !== '';
                    case 'greater than':
                        return parseFloat(cellValue) > parseFloat(filterValue);
                    case 'less than':
                        return parseFloat(cellValue) < parseFloat(filterValue);
                    case 'greater than or equal':
                        return parseFloat(cellValue) >= parseFloat(filterValue);
                    case 'less than or equal':
                        return parseFloat(cellValue) <= parseFloat(filterValue);
                    default:
                        return true;
                }
            });
        });
    }
    
    function displayAccordion(data) {
        const content = document.getElementById(INSTANCE_ID + '-content');
        
        if (data.length === 0) {
            content.innerHTML = '<div class="loading-message"><p>No items found matching the filter criteria.</p></div>';
            return;
        }
        
        const questionColumnIndices = getQuestionColumnIndices();
        const answerColumnIndex = getAnswerColumnIndex();
        const urlColumnIndices = getUrlColumnIndices();
        const hasAutoNumbering = CONFIG.AUTO_NUMBER_ROWS !== null;
        const numColumns = questionColumnIndices.length + (hasAutoNumbering ? 1 : 0);
        
        // Get transition settings (convert milliseconds to seconds)
        const transitionSpeed = (CONFIG.ANSWER_TRANSITION_SPEED / 1000) + 's';
        const transitionEffect = transitionEffects[CONFIG.ANSWER_TRANSITION_EFFECT] || 'ease';
        
        // Get text effect animation
        let textAnimation = 'none';
        if (CONFIG.ANSWER_TEXT_EFFECT !== 'none') {
            if (CONFIG.ANSWER_TEXT_EFFECT === 'slide') {
                const dir = CONFIG.ANSWER_TEXT_EFFECT_DIRECTION.charAt(0).toUpperCase() + CONFIG.ANSWER_TEXT_EFFECT_DIRECTION.slice(1);
                textAnimation = `slideFrom${dir}`;
            } else if (CONFIG.ANSWER_TEXT_EFFECT === 'bounce') {
                const dir = CONFIG.ANSWER_TEXT_EFFECT_DIRECTION.charAt(0).toUpperCase() + CONFIG.ANSWER_TEXT_EFFECT_DIRECTION.slice(1);
                textAnimation = `bounceFrom${dir}`;
            } else if (CONFIG.ANSWER_TEXT_EFFECT === 'zoom') {
                textAnimation = 'zoomIn';
            } else if (CONFIG.ANSWER_TEXT_EFFECT === 'fade') {
                textAnimation = 'fadeIn';
            } else if (CONFIG.ANSWER_TEXT_EFFECT === 'blur') {
                textAnimation = 'blurIn';
            } else if (CONFIG.ANSWER_TEXT_EFFECT === 'rotate') {
                textAnimation = 'rotateIn';
            } else if (CONFIG.ANSWER_TEXT_EFFECT === 'elastic') {
                textAnimation = 'elasticIn';
            }
        }
        
        const fadeDuration = CONFIG.ANSWER_TEXT_EFFECT_FADE / 1000;
        
        let html = '<div class="accordion-container"><table class="accordion-table">';
        
        isFirstDataRow = true;
        let rowNumber = CONFIG.AUTO_NUMBER_ROWS;
        
        console.log(`[Accordion] About to process ${data.length} rows in displayAccordion`);
        data.forEach((row, index) => {
            const answer = (row[answerColumnIndex] || '').trim();
            const hasQuestionContent = questionColumnIndices.some(colIndex => row[colIndex]);
    
            // Check if this is the header row that should always be displayed
            const actualRowNumber = CONFIG.STARTING_ROW + index;
            const isHeaderRow = CONFIG.HEADER_ROW_NUMBER !== 0 && actualRowNumber === CONFIG.HEADER_ROW_NUMBER;
    
            // Special handling: if HEADER_ROW_NUMBER is 1 and this is the first row, force display
            const forceDisplay = CONFIG.HEADER_ROW_NUMBER === 1 && index === 0;
    
            if (hasQuestionContent || isHeaderRow || forceDisplay) {
                // Use header row alignment only if HEADER_ROW_NUMBER is explicitly set and this is that row
                const alignmentToUse = (CONFIG.HEADER_ROW_NUMBER !== 0 && isHeaderRow) ? CONFIG.HEADER_ROW_COLUMNS_ALIGN : CONFIG.QUESTION_COLUMNS_ALIGN;
                
                html += `<tbody class="accordion-item" id="${INSTANCE_ID}-item-${index}" data-animation="${textAnimation}" data-duration="${fadeDuration}" data-transition-speed="${transitionSpeed}" data-transition-effect="${transitionEffect}">`;
    
                // Header rows are never expandable (no answers)
                if (answer && !isHeaderRow) {
                    html += `<tr class="accordion-question-row" onclick="window.toggleAccordion_${INSTANCE_ID}(event, ${index})">`;
                } else {
                    html += '<tr class="accordion-question-row no-answer">';
                }
                
                // Add auto-numbering column if enabled (FIRST - leftmost column)
                if (hasAutoNumbering) {
                    const autoNumberWidth = CONFIG.AUTO_NUMBER_COLUMN_WIDTH || 50;
                    html += `<td align="center" class="accordion-auto-number-cell" style="width: ${autoNumberWidth}px; max-width: ${autoNumberWidth}px;">${rowNumber}</td>`;
                }
                
				questionColumnIndices.forEach((colIndex, i) => {
					const cellValue = row[colIndex] || '';
					// Add question prefix to first question column only (but not for header rows)
					const questionPrefix = (i === 0 && CONFIG.QUESTION_PREFIX && !isHeaderRow && !forceDisplay) ? CONFIG.QUESTION_PREFIX : '';

					// Optional URL for this question column (from CONFIG.URL_COLUMNS)
					let linkedValue = '';
					const urlColIndex = urlColumnIndices[i];
					const rawUrl = (typeof urlColIndex === 'number') ? (row[urlColIndex] || '').toString().trim() : '';
					const hasUrl = rawUrl && /^https?:\/\//i.test(rawUrl);

					// Check if cell contains an image URL first
					const directImageUrl = isDirectImageURL(cellValue);
					
					if (directImageUrl && hasUrl) {
						// Cell has image URL AND a link URL - wrap image in link
						const imageTag = generateImageTag(directImageUrl, true);
						linkedValue = `<a href="${rawUrl}" target="_blank">${imageTag}</a>`;
					} else if (hasUrl) {
						// Cell has link URL but no image - wrap text in link
						let questionHtml = preserveWhitespace(cellValue);
						questionHtml = convertNewlinesToBR(questionHtml);
						linkedValue = `<a href="${rawUrl}" target="_blank">${questionPrefix}${questionHtml}</a>`;
					} else {
						// No link URL - use existing URL/image detection
						const processedValue = convertURLsToLinks(cellValue, true);
						linkedValue = questionPrefix + processedValue;
					}

                    const alignment = alignmentToUse[i] || 'left';
                    
                    // Check if cell contains only an image
                    const isImageCell = linkedValue.includes('class="accordion-image-content"');
                    const cellClass = isImageCell ? 'accordion-image-cell' : '';
                    
                    // Get column width (use array index directly, no offset for auto-numbering)
                    const columnWidth = CONFIG.COLUMN_WIDTHS && CONFIG.COLUMN_WIDTHS[i] ? CONFIG.COLUMN_WIDTHS[i] : '';
                    const widthStyle = columnWidth ? ` style="width: ${columnWidth}px; max-width: ${columnWidth}px;"` : '';
                    
                    html += `<td align="${alignment}" class="${cellClass}"${widthStyle}>${linkedValue}</td>`;
                });
                html += '</tr>';
                
				if (answer) {
					const answerPrefix = CONFIG.ANSWER_PREFIX || '';
					const processedAnswer = convertURLsToLinks(answer, false);
					
					// Get enlarged thumbnail if enabled
					let enlargedThumbnail = '';
					if (CONFIG.SHOW_ENLARGED_THUMBNAIL_IN_ANSWER_ROW) {
						// Find the leftmost column with an image
						for (let i = 0; i < questionColumnIndices.length; i++) {
							const colIndex = questionColumnIndices[i];
							const cellValue = row[colIndex] || '';
							const imageUrl = isDirectImageURL(cellValue);
							
							if (imageUrl) {
								enlargedThumbnail = `<div class="accordion-answer-thumbnail"><img src="${imageUrl}" alt="Enlarged thumbnail" loading="lazy"></div>`;
								break; // Use only the first (leftmost) image found
							}
						}
					}
					
					html += `
						<tr class="accordion-answer-row">
							<td colspan="${numColumns}" class="accordion-answer-cell">
								<div class="accordion-answer-wrapper">
									<div class="accordion-answer-content">${enlargedThumbnail}${answerPrefix}${processedAnswer}</div>
								</div>
							</td>
						</tr>
					`;
				}
                
                html += '</tbody>';
                isFirstDataRow = false;
                rowNumber++;
            }
        });
        
        html += '</table></div>';
        content.innerHTML = html;
        
        // Mark the last item as last-visible-item on initial load
        const allItems = content.querySelectorAll('.accordion-item');
        if (allItems.length > 0) {
            allItems[allItems.length - 1].classList.add('last-visible-item');
        }
    
        // Apply alternating row colors on initial display
        applyAlternatingRowColors();
        
        console.log('[Accordion] Display complete');
    }
    
    window['toggleAccordion_' + INSTANCE_ID] = function(event, index) {
        if (event.target.tagName === 'A' || event.target.closest('a')) {
            return;
        }
        
        const item = document.getElementById(`${INSTANCE_ID}-item-${index}`);
        if (item) {
            const wasActive = item.classList.contains('active');
            const answerWrapper = item.querySelector('.accordion-answer-wrapper');
            const answerContent = item.querySelector('.accordion-answer-content');
            
            if (!wasActive && answerContent && answerWrapper) {
                // Opening - apply transition and text animation
                const transitionSpeed = item.dataset.transitionSpeed;
                const transitionEffect = item.dataset.transitionEffect;
                const animation = item.dataset.animation;
                const duration = item.dataset.duration;
                
                // Apply wrapper transition
                answerWrapper.style.transition = `max-height ${transitionSpeed} ${transitionEffect}`;
                
                // Apply text animation
                if (animation && animation !== 'none') {
                    answerContent.style.animation = `${animation} ${duration}s ease-out`;
                    
                    // Reset animation after it completes
                    setTimeout(() => {
                        answerContent.style.animation = '';
                    }, parseFloat(duration) * 1000);
                }
                
                item.classList.add('active');
            } else {
                // Closing - remove instantly
                answerWrapper.style.transition = 'none';
                item.classList.remove('active');
                
                // Force reflow to apply transition: none immediately
                void answerWrapper.offsetHeight;
            }
        }
    };
    
    function showError(message, hasErrorCode) {
        const content = document.getElementById(INSTANCE_ID + '-content');
        
        if (hasErrorCode) {
            content.innerHTML = `
                <div class="error-message">
                    <h3>⚠️ Error Loading Data</h3>
                    <p><strong>Error:</strong> ${message}</p>
                    <p style="margin-top: 10px; font-size: 14px;">
                        Make sure your spreadsheet is publicly accessible and the API key is valid.
                    </p>
                </div>
            `;
        } else {
            let remainingMs = CONFIG.AUTO_REFETCH_DATA_ON_ERROR_DELAY;
            
            const errorDiv = document.createElement('div');
            errorDiv.className = 'error-message auto-retry';
            errorDiv.innerHTML = `
                <h3>⚠️ Connection Issue</h3>
                <p>${message ? 'Error: ' + message : 'No error code found.'}</p>
                <p><strong>Auto reloading data in <span id="${INSTANCE_ID}-countdown" style="font-weight: bold; color: #d39e00;">${(remainingMs / 1000).toFixed(1)}</span> seconds...</strong></p>
                <p style="margin-top: 10px; font-size: 12px; opacity: 0.7;">
                    Hover over this message to pause the countdown.
                </p>
            `;
            
            content.innerHTML = '';
            content.appendChild(errorDiv);
            
            const countdownSpan = document.getElementById(INSTANCE_ID + '-countdown');
            
            errorDiv.addEventListener('mouseenter', () => { isPaused = true; });
            errorDiv.addEventListener('mouseleave', () => { isPaused = false; });
            
            countdownInterval = setInterval(() => {
                if (!isPaused) {
                    remainingMs -= 100;
                    const remainingSeconds = remainingMs / 1000;
                    if (remainingSeconds <= 0) {
                        clearInterval(countdownInterval);
                        countdownSpan.textContent = '0.0';
                    } else {
                        countdownSpan.textContent = remainingSeconds.toFixed(1);
                    }
                }
            }, 100);
            
            retryTimeout = setTimeout(function checkAndRetry() {
                if (!isPaused) {
                    clearInterval(countdownInterval);
                    loadData();
                } else {
                    retryTimeout = setTimeout(checkAndRetry, 100);
                }
            }, CONFIG.AUTO_REFETCH_DATA_ON_ERROR_DELAY);
        }
    }
    
    // ============================================
    // AUTO-LOAD
    // ============================================
    if (CONFIG.FETCH_DATA_ON_PAGE_LOAD) {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', loadData);
        } else {
            setTimeout(loadData, 0);
        }
    }
    
    })();