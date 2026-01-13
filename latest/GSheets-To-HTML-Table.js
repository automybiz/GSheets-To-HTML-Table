(function() {
    const MAX_WAIT_MS = 5000;
    const POLL_INTERVAL_MS = 100;
    const startTime = Date.now();

    function initialize() {
        // 1. Resolve CONFIG
        // We look for CONFIG in the local scope first, then window.CONFIG
        // This handles cases where the script might be wrapped in a module or sandbox
        let activeConfig = null;
        try {
            activeConfig = window.CONFIG || (typeof CONFIG !== 'undefined' ? CONFIG : null);
        } catch (e) {
            // Accessing undefined variables might throw in strict mode
        }

        if (!activeConfig) {
            if (Date.now() - startTime < MAX_WAIT_MS) {
                setTimeout(initialize, POLL_INTERVAL_MS);
                return;
            }
            console.error('[GSheets-To-HTML-Table] ERROR: CONFIG object not found after ' + (MAX_WAIT_MS/1000) + 's. Ensure a script tag defines CONFIG before or shortly after this loader.');
            return;
        }

        // 2. Resolve Wrapper
        // We look for an uninitialized .accordion-wrapper
        const wrappers = document.querySelectorAll('.accordion-wrapper:not([data-initialized="true"])');
        const wrapper = wrappers[wrappers.length - 1]; // Pick the last one added to the DOM

        if (!wrapper) {
            if (Date.now() - startTime < MAX_WAIT_MS) {
                setTimeout(initialize, POLL_INTERVAL_MS);
                return;
            }
            console.error('[GSheets-To-HTML-Table] ERROR: No uninitialized .accordion-wrapper element found. Ensure <div class="accordion-wrapper"></div> exists in your HTML.');
            return;
        }

        // 3. Start Actual Initialization
        startAccordion(wrapper, activeConfig);
    }

    function startAccordion(wrapper, CONFIG) {
        // Generate unique instance ID
        const INSTANCE_ID = 'accordion_' + Math.random().toString(36).substr(2, 9);
        console.log('[GSheets-To-HTML-Table] Initializing instance:', INSTANCE_ID);
        
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
        const loadedFonts = new Set();
        let headers = [];
        let searchTimeout = null;
        let isFirstDataRow = true;
        let retryTimeout = null;
        let countdownInterval = null;
        let isPaused = false;
        let dateRange = { min: Infinity, max: -Infinity };
        const viewedTimers = new Map();
        
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
            if (letter === 'SHOW') return 'SHOW';
            return letter.toUpperCase().charCodeAt(0) - 65;
        }
        
        function getQuestionColumnIndices() {
            return CONFIG.QUESTION_COLUMNS.map(col => columnLetterToIndex(col));
        }
        
        function getAnswerColumnIndex() {
            return columnLetterToIndex(CONFIG.ANSWER_COLUMN);
        }

        // ============================================
        // VIEWED STATUS FUNCTIONS
        // ============================================
        const VIEWED_STORAGE_KEY = 'gsheets_viewed_' + CONFIG.SPREADSHEET_ID;
        
        function loadViewedData() {
            try {
                const data = localStorage.getItem(VIEWED_STORAGE_KEY);
                return data ? JSON.parse(data) : {};
            } catch (e) {
                console.warn('[Accordion] Failed to load viewed data', e);
                return {};
            }
        }
        
        function saveViewedData(data) {
            try {
                localStorage.setItem(VIEWED_STORAGE_KEY, JSON.stringify(data));
            } catch (e) {
                console.warn('[Accordion] Failed to save viewed data', e);
            }
        }

        function getViewedBadgeColor(timestamp, minTime, maxTime) {
            if (!timestamp || minTime === maxTime) return CONFIG.DATE_DOTS_HEAT_MAP.COLOR_MOST_RECENT;
            const factor = (timestamp - minTime) / (maxTime - minTime);
            return interpolateColor(CONFIG.DATE_DOTS_HEAT_MAP.COLOR_LEAST_RECENT, CONFIG.DATE_DOTS_HEAT_MAP.COLOR_MOST_RECENT, factor);
        }
        
        function createViewedBadge(timestamp, minTime, maxTime) {
            if (!timestamp) return '';
            
            const color = getViewedBadgeColor(timestamp, minTime, maxTime);
            const dateObj = new Date(timestamp);
            
            // Format: YYYY-MM-DD HH:MMam/pm
            const year = dateObj.getFullYear();
            const month = String(dateObj.getMonth() + 1).padStart(2, '0');
            const day = String(dateObj.getDate()).padStart(2, '0');
            let hours = dateObj.getHours();
            const minutes = String(dateObj.getMinutes()).padStart(2, '0');
            const ampm = hours >= 12 ? 'pm' : 'am';
            hours = hours % 12;
            hours = hours ? hours : 12; // the hour '0' should be '12'
            const strTime = `${hours}:${minutes}${ampm}`;
            const dateStr = `${year}-${month}-${day} ${strTime}`;
            
            // Tooltip content using CONFIG.VIEWED_TITLE template
            let title = CONFIG.VIEWED_TITLE || '';
            title = title.replace('[VIEWED_DATE]', dateStr);
            
            return `<span class="badges badge-viewed" style="background-color: ${color};" title="${title}">${CONFIG.VIEWED_TEXT}</span>`;
        }

        function createUnseenBadge() {
            if (!CONFIG.UNSEEN_TEXT) return '';
            return `<span class="badges badge-unseen" title="${CONFIG.UNSEEN_TITLE}">${CONFIG.UNSEEN_TEXT}</span>`;
        }

        function refreshAllViewedBadges() {
            // Reload data to get fresh min/max
            const viewedData = loadViewedData();
            const timestamps = Object.values(viewedData);
            if (timestamps.length === 0) return;
            
            const minTime = Math.min(...timestamps);
            const maxTime = Math.max(...timestamps);
            
            // Update all existing badges
            document.querySelectorAll('.badge-container').forEach(container => {
                const rowId = container.dataset.rowId;
                if (viewedData[rowId]) {
                    container.innerHTML = createViewedBadge(viewedData[rowId], minTime, maxTime);
                }
            });
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

        function isYouTubeURL(url) {
            if (!url || typeof url !== 'string') return false;
            return url.includes('youtube.com') || url.includes('youtu.be');
        }
        
        function isDirectImageURL(text) {
            if (!text || typeof text !== 'string') return false;
            
            const trimmedText = text.trim();
            
            // Split by newlines and filter out empty lines
            const lines = trimmedText.split(/\r?\n/).map(line => line.trim()).filter(line => line !== '');
            if (lines.length === 0) return false;

            // Check if every non-empty line is an image URL
            const allAreImages = lines.every(line => {
                const urlOnlyPattern = /^(https?:\/\/[^\s]+)$/;
                const match = line.match(urlOnlyPattern);
                return match && isImageURL(match[1]);
            });

            if (allAreImages) {
                console.log('[Accordion] Direct image URL(s) detected');
                return lines; // Return array of URLs
            }
            
            return false;
        }

        function interpolateColor(color1, color2, factor) {
            if (arguments.length < 3) factor = 0.5;
            
            const parseHex = (hex) => {
                let r = 0, g = 0, b = 0;
                hex = hex.replace('#', '');
                if (hex.length === 3) {
                    r = parseInt(hex[0] + hex[0], 16);
                    g = parseInt(hex[1] + hex[1], 16);
                    b = parseInt(hex[2] + hex[2], 16);
                } else if (hex.length === 6) {
                    r = parseInt(hex.substring(0, 2), 16);
                    g = parseInt(hex.substring(2, 4), 16);
                    b = parseInt(hex.substring(4, 6), 16);
                }
                return [r, g, b];
            };

            const c1 = parseHex(color1);
            const c2 = parseHex(color2);

            const r = Math.round(c1[0] + factor * (c2[0] - c1[0]));
            const g = Math.round(c1[1] + factor * (c2[1] - c1[1]));
            const b = Math.round(c1[2] + factor * (c2[2] - c1[2]));

            // Convert to shorthand hex by taking the first digit of each 2-digit hex value
            const rh = r.toString(16).padStart(2, '0')[0];
            const gh = g.toString(16).padStart(2, '0')[0];
            const bh = b.toString(16).padStart(2, '0')[0];

            return `#${rh}${gh}${bh}`;
        }

        function getHeatMapColor(dateStr, minTime, maxTime) {
            if (!dateStr || minTime === maxTime) return CONFIG.DATE_DOTS_HEAT_MAP.COLOR_MOST_RECENT;
            
            const time = new Date(dateStr).getTime();
            if (isNaN(time)) return null;

            const factor = (time - minTime) / (maxTime - minTime);
            return interpolateColor(CONFIG.DATE_DOTS_HEAT_MAP.COLOR_LEAST_RECENT, CONFIG.DATE_DOTS_HEAT_MAP.COLOR_MOST_RECENT, factor);
        }
        
        function generateImageTag(imageUrl, isInCell = false) {
            let styles = [];
            let maxWidth, maxHeight;

            if (isInCell) {
                // Thumbnail settings
                maxWidth = CONFIG.IMAGE_THUMB_MAX_WIDTH !== undefined ? CONFIG.IMAGE_THUMB_MAX_WIDTH : CONFIG.IMAGE_MAX_WIDTH;
                maxHeight = CONFIG.IMAGE_THUMB_MAX_HEIGHT !== undefined ? CONFIG.IMAGE_THUMB_MAX_HEIGHT : CONFIG.IMAGE_MAX_HEIGHT;
            } else {
                // Answer/Expanded settings
                maxWidth = CONFIG.IMAGE_ANSWER_MAX_WIDTH !== undefined ? CONFIG.IMAGE_ANSWER_MAX_WIDTH : 555;
                maxHeight = CONFIG.IMAGE_ANSWER_MAX_HEIGHT !== undefined ? CONFIG.IMAGE_ANSWER_MAX_HEIGHT : null;
            }
            
            if (maxWidth) {
                styles.push(`max-width: ${maxWidth}px`);
            }
            
            if (maxHeight) {
                styles.push(`max-height: ${maxHeight}px`);
            }
            
            // If both width and height are set for thumbnails, stretch to fit
            if (isInCell && maxWidth && maxHeight) {
                styles.push(`width: ${maxWidth}px`);
                styles.push(`height: ${maxHeight}px`);
                styles.push('object-fit: fill');
            } else {
                styles.push('width: auto');
                styles.push('height: auto');
                styles.push('object-fit: contain');
            }
            
            // Center the image and remove extra space
            styles.push('display: block');
            
            if (isInCell) {
                styles.push('margin: 0 auto');
            } else {
                const align = CONFIG.IMAGE_ANSWER_ALIGN || 'center';
                if (align === 'left') {
                    styles.push('margin: 0 auto 0 0');
                } else if (align === 'right') {
                    styles.push('margin: 0 0 0 auto');
                } else {
                    styles.push('margin: 0 auto');
                }
            }
            
            // Add class to identify image cells for padding removal
            const className = isInCell ? 'class="accordion-image-content"' : 'class="accordion-answer-image"';
            
            return `<img src="${imageUrl}" ${className} style="${styles.join('; ')}" alt="Image" loading="lazy">`;
        }
        
        function extractImagesFromCell(text) {
            if (!text) return [];
            
            // Handle IMAGE formula (usually only one per cell in GSheets)
            const imageFormulaMatch = text.match(/=IMAGE\s*\(\s*["']([^"']+)["']/i);
            if (imageFormulaMatch) {
                return [imageFormulaMatch[1]];
            }
            
            // Handle Rich Text from GSheets (where links might be wrapped in tags)
            // Use a robust way to extract all URLs that are image URLs from each line
            const urls = [];
            
            // Use a temporary element to clean up HTML and handle line breaks
            const temp = document.createElement('div');
            temp.innerHTML = text;
            
            // Replace various line-breaking tags with actual newlines
            const cleanedHtml = temp.innerHTML
                .replace(/<br\s*\/?>/gi, '\n')
                .replace(/<\/p>/gi, '\n')
                .replace(/<\/div>/gi, '\n')
                .replace(/<p[^>]*>/gi, '')
                .replace(/<div[^>]*>/gi, '');
                
            const temp2 = document.createElement('div');
            temp2.innerHTML = cleanedHtml;
            const plainText = temp2.textContent || temp2.innerText || '';

            // Split by newlines and filter out empty lines
            const lines = plainText.split(/\n/).map(line => line.trim()).filter(line => line !== '');
            
            lines.forEach(line => {
                // Match the first URL in the line
                const urlMatch = line.match(/(https?:\/\/[^\s<"']+)/);
                if (urlMatch) {
                    const url = urlMatch[1];
                    if (isImageURL(url)) {
                        urls.push(url);
                    }
                }
            });

            return urls;
        }

        function loadGoogleFont(fontName) {
            if (!fontName || loadedFonts.has(fontName)) return;

            // Skip default web-safe fonts
            const webSafeFonts = ['Arial', 'Verdana', 'Times New Roman', 'Courier New', 'Georgia', 'Trebuchet MS', 'Comic Sans MS', 'Impact'];
            if (webSafeFonts.includes(fontName)) return;

            console.log('[Accordion] Automatically loading Google Font:', fontName);
            
            const link = document.createElement('link');
            link.rel = 'stylesheet';
            // Normalize font name for Google Fonts API (spaces to +)
            const normalizedName = fontName.replace(/\s+/g, '+');
            link.href = `https://fonts.googleapis.com/css2?family=${normalizedName}&display=swap`;
            
            document.head.appendChild(link);
            loadedFonts.add(fontName);
        }
        
        // ============================================
        // HOVER EVENT MANAGEMENT
        // ============================================
        const hoverTimeouts = new Map();
        const HOVER_DELAY = 300; // milliseconds before loading content on hover
        
        function addHoverListeners() {
            const questionRows = document.querySelectorAll('#' + INSTANCE_ID + '-content .accordion-question-row');
            
            questionRows.forEach(row => {
                // Skip header rows and rows without answers
                if (row.classList.contains('no-answer')) return;
                
                row.addEventListener('mouseenter', function(e) {
                    const item = this.closest('.accordion-item');
                    if (!item) return;
                    
                    const itemId = item.id;
                    if (hoverTimeouts.has(itemId)) {
                        clearTimeout(hoverTimeouts.get(itemId));
                    }
                    
                    const timeoutId = setTimeout(() => {
                        processLazyContent(item);
                        hoverTimeouts.delete(itemId);
                    }, HOVER_DELAY);
                    
                    hoverTimeouts.set(itemId, timeoutId);
                });
                
                row.addEventListener('mouseleave', function(e) {
                    const item = this.closest('.accordion-item');
                    if (!item) return;
                    
                    const itemId = item.id;
                    if (hoverTimeouts.has(itemId)) {
                        clearTimeout(hoverTimeouts.get(itemId));
                        hoverTimeouts.delete(itemId);
                    }
                });
                
                // Also load content when clicking to expand
                row.addEventListener('click', function(e) {
                    // Don't interfere with link clicks
                    if (e.target.tagName === 'A' || e.target.closest('a')) return;
                    
                    const item = this.closest('.accordion-item');
                    if (item) {
                        // Process immediately on click (no delay)
                        processLazyContent(item);
                        
                        // Clear any pending hover timeouts
                        const itemId = item.id;
                        if (hoverTimeouts.has(itemId)) {
                            clearTimeout(hoverTimeouts.get(itemId));
                            hoverTimeouts.delete(itemId);
                        }
                    }
                });
            });
            
            console.log('[Accordion] Added hover listeners to', questionRows.length, 'question rows');
        }
        
        // ============================================
        // LAZY LOADING FUNCTIONS
        // ============================================
        function generateLazyImagePlaceholder(imageUrl, isInCell = false) {
            const className = isInCell ? 'class="accordion-image-content lazy-image-placeholder"' : 'class="lazy-image-placeholder"';
            const placeholderStyle = isInCell ? 'display: block; margin: 0 auto; background: #333; border: 2px dashed #666; border-radius: 4px; text-align: center; color: #999; font-size: 12px; padding: 20px 10px;' : 'background: #333; border: 2px dashed #666; border-radius: var(--answer-image-border-radius); text-align: center; color: #999; font-size: 14px; padding: 40px 20px; margin: 10px 0;';
            
            return `<div ${className} data-original-url="${imageUrl}" style="${placeholderStyle}" onclick="window.loadLazyImage(this)">${imageUrl}</div>`;
        }
        
        function generateLazyYouTubePlaceholder(url, videoId, videoTitle, timeParam, playlistParam) {
            const placeholderStyle = 'background: #222; border: 2px dashed #666; border-radius: var(--answer-video-border-radius); text-align: center; color: #999; font-size: 14px; padding: 40px 20px; cursor: pointer;';
            const videoData = encodeURIComponent(JSON.stringify({ url, videoId, timeParam, playlistParam }));
            
            return `<div class="lazy-youtube-placeholder" data-video-data="${videoData}" style="${placeholderStyle}" onclick="window.loadLazyYouTube(this)">${url}</div>`;
        }
        
        window.loadLazyImage = function(element) {
            if (element.classList.contains('loading')) return;
            
            element.classList.add('loading');
            element.innerHTML = '⏳ Loading image...';
            
            const imageUrl = element.dataset.originalUrl;
            const img = new Image();
            
            img.onload = function() {
                const isInCell = element.classList.contains('accordion-image-content');
                let styles = [];
                let maxWidth, maxHeight;

                if (isInCell) {
                    // Thumbnail settings
                    maxWidth = CONFIG.IMAGE_THUMB_MAX_WIDTH !== undefined ? CONFIG.IMAGE_THUMB_MAX_WIDTH : CONFIG.IMAGE_MAX_WIDTH;
                    maxHeight = CONFIG.IMAGE_THUMB_MAX_HEIGHT !== undefined ? CONFIG.IMAGE_THUMB_MAX_HEIGHT : CONFIG.IMAGE_MAX_HEIGHT;
                } else {
                    // Answer/Expanded settings
                    maxWidth = CONFIG.IMAGE_ANSWER_MAX_WIDTH !== undefined ? CONFIG.IMAGE_ANSWER_MAX_WIDTH : 555;
                    maxHeight = CONFIG.IMAGE_ANSWER_MAX_HEIGHT !== undefined ? CONFIG.IMAGE_ANSWER_MAX_HEIGHT : null;
                }
                
                if (maxWidth) {
                    styles.push(`max-width: ${maxWidth}px`);
                }
                
                if (maxHeight) {
                    styles.push(`max-height: ${maxHeight}px`);
                }
                
                if (CONFIG.IMAGE_MAINTAIN_ASPECT_RATIO) {
                    styles.push('width: auto');
                    styles.push('height: auto');
                    styles.push('object-fit: contain');
                } else {
                    if (maxWidth) styles.push('width: 100%');
                    if (maxHeight) styles.push(`height: ${maxHeight}px`);
                    styles.push('object-fit: fill');
                }
                
                styles.push('display: block');
                
                if (isInCell) {
                    styles.push('margin: 0 auto');
                } else {
                    const align = CONFIG.IMAGE_ANSWER_ALIGN || 'center';
                    if (align === 'left') {
                        styles.push('margin: 0 auto 0 0');
                    } else if (align === 'right') {
                        styles.push('margin: 0 0 0 auto');
                    } else {
                        styles.push('margin: 0 auto');
                    }
                }
                
                const className = isInCell ? 'class="accordion-image-content"' : 'class="accordion-answer-image"';
                const finalHtml = `<img src="${imageUrl}" ${className} style="${styles.join('; ')}" alt="Image" loading="lazy">`;
                
                element.outerHTML = finalHtml;
            };
            
            img.onerror = function() {
                element.innerHTML = '❌ Failed to load image<br><a href="' + imageUrl + '" target="_blank" style="color: #0ff;">Open in new tab</a>';
                element.classList.remove('loading');
            };
            
            img.src = imageUrl;
        };
        
        window.loadLazyYouTube = function(element) {
            if (element.classList.contains('loading')) return;
            
            element.classList.add('loading');
            element.innerHTML = '⏳ Loading video...';
            
            try {
                const videoData = JSON.parse(decodeURIComponent(element.dataset.videoData));
                const { url, videoId, timeParam, playlistParam } = videoData;
                
                const iframe = document.createElement('iframe');
                iframe.width = CONFIG.YOUTUBE_EMBED_WIDTH;
                iframe.height = CONFIG.YOUTUBE_EMBED_HEIGHT;
                
                // Build the same robust URL format for lazy loading
                let finalUrl;
                const playlistIdMatch = playlistParam ? playlistParam.match(/[&?]list=([a-zA-Z0-9_-]+)/) : null;
                const playlistId = playlistIdMatch ? playlistIdMatch[1] : null;
                
                const startMatch = timeParam ? timeParam.match(/[?&]start=(\d+)/) : null;
                const startTime = startMatch ? startMatch[1] : null;

                if (videoId && playlistId) {
                    const params = new URLSearchParams();
                    params.set('list', playlistId);
                    if (startTime) params.set('start', startTime);
                    finalUrl = `https://www.youtube.com/embed/${videoId}?${params.toString()}`;
                } else if (playlistId) {
                    const params = new URLSearchParams();
                    params.set('list', playlistId);
                    if (startTime) params.set('start', startTime);
                    finalUrl = `https://www.youtube.com/embed/videoseries?${params.toString()}`;
                } else {
                    finalUrl = `https://www.youtube.com/embed/${videoId}${timeParam}`;
                }

                iframe.src = finalUrl;
                iframe.title = 'YouTube video player';
                iframe.frameBorder = '0';
                iframe.setAttribute('allowtransparency', 'true');
                iframe.allow = 'accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture';
                iframe.allowFullscreen = true;
                iframe.className = 'accordion-answer-video';
                
                element.outerHTML = iframe.outerHTML;
            } catch (error) {
                console.error('[Accordion] Error loading YouTube video:', error);
                element.innerHTML = '❌ Failed to load video';
                element.classList.remove('loading');
            }
        };
        
        function processLazyContent(rowElement) {
            if (!rowElement || rowElement.dataset.lazyProcessed) return;
            
            const lazyImages = rowElement.querySelectorAll('.lazy-image-placeholder');
            const lazyYouTube = rowElement.querySelectorAll('.lazy-youtube-placeholder');
            
            // Process images
            lazyImages.forEach(placeholder => {
                window.loadLazyImage(placeholder);
            });
            
            // Process YouTube videos
            lazyYouTube.forEach(placeholder => {
                window.loadLazyYouTube(placeholder);
            });
            
            // Mark as processed
            rowElement.dataset.lazyProcessed = 'true';
        }
        
        function convertURLsToLinks(text, isInCell = false, lazyMode = false) {
            if (!text) return '';
            
            // Special Case: Is it an =IMAGE() formula? (Handled separately as it's a specific GSheets feature)
            const formulaImageUrls = extractImagesFromCell(text);
            // We only trigger the early-return for formulas if they are the only thing that would be returned,
            // or if it's a known =IMAGE formula which GSheets treats as the cell content.
            if (text.trim().startsWith('=IMAGE') && formulaImageUrls.length > 0) {
                if (isInCell) {
                    const firstImageUrl = formulaImageUrls[0];
                    return lazyMode ? generateLazyImagePlaceholder(firstImageUrl, isInCell) : generateImageTag(firstImageUrl, isInCell);
                } else {
                    return formulaImageUrls.map(url => {
                        return lazyMode ? generateLazyImagePlaceholder(url, isInCell) : generateImageTag(url, isInCell);
                    }).join('');
                }
            }

            let result = preserveWhitespace(text);
            
            // Pre-process: Unwrap <a> tags that point to YouTube or Images
            const anchorRegex = /<a\s+[^>]*href=["']([^"']+)["'][^>]*>(.*?)<\/a>/gi;
            result = result.replace(anchorRegex, (match, href, content) => {
                if (isYouTubeURL(href) || isImageURL(href)) {
                    return href; 
                }
                return match;
            });
            
            // Regex to match URLs and specifically capture a following newline (\n or \r\n)
            // Enhanced to handle &nbsp; (don't consume it) and strip trailing punctuation
            const urlRegex = /(href="|src="|href='|src=')((?:https?:\/\/)[^"']+)("|')|(https?:\/\/(?:[^\s<&]|&(?!nbsp;))+(?:[^\s<&.,?!:;]|\/))(\r\n|\n)?/g;
            
            result = result.replace(urlRegex, (match, attrPrefix, urlInAttr, quote, plainUrl, followingNewline) => {
                // If attrPrefix is defined, it means we matched an existing HTML attribute
                if (attrPrefix) return match;
                
                // Standalone URL
                let url = plainUrl;
                
                // Clean up common Rich Text wrapping tags if they've leaked into the URL string
                url = url.replace(/<\/?[aui]>|&nbsp;/gi, '');
                
                // YouTube Logic: Check for YouTube links to embed
                if (isYouTubeURL(url)) {
                    // Skip embedding for root domains
                    if (url === 'https://www.youtube.com' || url === 'https://www.youtube.com/' || 
                        url === 'https://youtube.com' || url === 'https://youtube.com/' ||
                        url === 'https://youtu.be' || url === 'https://youtu.be/') {
                        return `<a href="${url}" target="_blank">${url}</a>` + (followingNewline || '');
                    }
                    
                    let videoId = null;
                    let playlistId = null;
                    let startTime = null;
                    let videoTitle = 'YouTube Video';
                    
                    // Extract Video ID from watch?v= format
                    let match = url.match(/[?&]v=([a-zA-Z0-9_-]+)/);
                    if (match) {
                        videoId = match[1];
                    }
                    
                    // Extract Video ID from /shorts/ format
                    if (!videoId) {
                        match = url.match(/youtube\.com\/shorts\/([a-zA-Z0-9_-]+)/);
                        if (match) videoId = match[1];
                    }

                    // Extract Video ID from youtu.be/ format
                    if (!videoId) {
                        match = url.match(/youtu\.be\/([a-zA-Z0-9_-]+)/);
                        if (match) videoId = match[1];
                    }

                    // Extract Playlist ID
                    const listMatch = url.match(/[?&]list=([a-zA-Z0-9_-]+)/);
                    if (listMatch) {
                        playlistId = listMatch[1];
                    }

                    // Extract Timestamp (t=)
                    const tMatch = url.match(/[?&]t=([0-9hms]+)/);
                    if (tMatch) {
                        const tVal = tMatch[1];
                        if (/^\d+$/.test(tVal)) {
                            // Pure seconds
                            startTime = tVal;
                        } else {
                            // Handle formats like 1m30s, 1h2m, etc.
                            let totalSeconds = 0;
                            const hMatch = tVal.match(/(\d+)h/);
                            const mMatch = tVal.match(/(\d+)m/);
                            const sMatch = tVal.match(/(\d+)s/);
                            
                            if (hMatch) totalSeconds += parseInt(hMatch[1]) * 3600;
                            if (mMatch) totalSeconds += parseInt(mMatch[1]) * 60;
                            if (sMatch) totalSeconds += parseInt(sMatch[1]);
                            
                            if (totalSeconds > 0) startTime = totalSeconds.toString();
                        }
                    }
                    
                    if (videoId || playlistId) {
                        let finalUrl;
                        if (videoId && playlistId) {
                            // Priority format for starting at a specific video: embed/[VIDEO_ID]?list=[PLAYLIST_ID]
                            const params = new URLSearchParams();
                            params.set('list', playlistId);
                            if (startTime) params.set('start', startTime);
                            finalUrl = `https://www.youtube.com/embed/${videoId}?${params.toString()}`;
                        } else if (playlistId) {
                            // Fallback: embed/videoseries?list=[PLAYLIST_ID]
                            const params = new URLSearchParams();
                            params.set('list', playlistId);
                            if (startTime) params.set('start', startTime);
                            finalUrl = `https://www.youtube.com/embed/videoseries?${params.toString()}`;
                        } else {
                            // Standard: embed/[VIDEO_ID]?start=[TIME]
                            const params = new URLSearchParams();
                            if (startTime) params.set('start', startTime);
                            const query = params.toString();
                            finalUrl = `https://www.youtube.com/embed/${videoId}${query ? '?' + query : ''}`;
                        }

                        const videoHtml = lazyMode 
                            ? generateLazyYouTubePlaceholder(url, videoId, videoTitle, startTime ? `?start=${startTime}` : '', playlistId ? `&list=${playlistId}` : '')
                            : `<iframe width="${CONFIG.YOUTUBE_EMBED_WIDTH}" height="${CONFIG.YOUTUBE_EMBED_HEIGHT}" src="${finalUrl}" title="YouTube video player" frameborder="0" allowtransparency="true" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen class="accordion-answer-video"></iframe>`;
                        
                        if (!isInCell) {
                            const align = CONFIG.YOUTUBE_EMBED_ALIGN || 'right';
                            return `<div class="accordion-video-container" style="text-align: ${align} !important;">${videoHtml}</div>`;
                        }
                        return videoHtml;
                    }
                }
                
                // Image Logic: Check if the URL is a direct image link
                if (isImageURL(url)) {
                    let imgHtml = '';
                    if (lazyMode) {
                        imgHtml = generateLazyImagePlaceholder(url, isInCell);
                    } else {
                        imgHtml = generateImageTag(url, isInCell);
                    }
                    // Return media HTML but DO NOT restore the newline (pruning it as requested)
                    return imgHtml;
                }
                
                // Default: Convert plain URL to clickable link
                // Restore the newline for standard links so they stay on separate lines
                return `<a href="${url}" target="_blank">${url}</a>` + (followingNewline || '');
            });
            
            result = convertNewlinesToBR(result);
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
                        questionRow.classList.remove('odd-row', 'even-row', 'first-visible-row'); // Remove existing classes
                        if (index === 0) {
                            questionRow.classList.add('first-visible-row');
                        }
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
        function applyTextFormatting(cell) {
            if (!cell) return '';
            
            const formattedValue = cell.formattedValue || '';
            
            // If no format runs, check for cell-level formatting
            if (!cell.textFormatRuns || cell.textFormatRuns.length === 0) {
                const format = (cell.effectiveFormat && cell.effectiveFormat.textFormat) || {};
                let style = [];

                // Only apply font size if it's explicitly different from standard spreadsheet defaults (10 or 11pt)
                // This prevents overriding the --question-font-size defined in CSS for normal text.
                if (format.fontSize && format.fontSize !== 10 && format.fontSize !== 11) {
                    style.push(`font-size: ${format.fontSize}pt`);
                }
                
                if (format.fontFamily && format.fontFamily !== 'Arial') {
                    loadGoogleFont(format.fontFamily);
                    style.push(`font-family: '${format.fontFamily}'`);
                }
                
                // Only apply foreground color if it's not pure black (0,0,0) 
                // to avoid clashing with dark themes where black is often the default GSheets color
                if (format.foregroundColor) {
                    const color = format.foregroundColor;
                    const r = Math.round((color.red || 0) * 255);
                    const g = Math.round((color.green || 0) * 255);
                    const b = Math.round((color.blue || 0) * 255);
                    
                    // Only apply if NOT black AND NOT very dark (to handle various "almost black" defaults)
                    if (r > 30 || g > 30 || b > 30) {
                        style.push(`color: rgb(${r},${g},${b})`);
                    }
                }
                
                if (format.bold) style.push('font-weight: bold');
                if (format.italic) style.push('font-style: italic');
                if (format.underline) style.push('text-decoration: underline');
                if (format.strikethrough) style.push('text-decoration: line-through');

                if (style.length > 0) {
                    return `<span style="${style.join('; ')}">${formattedValue}</span>`;
                }
                return formattedValue;
            }

            const text = cell.userEnteredValue ? (cell.userEnteredValue.stringValue || String(cell.userEnteredValue.numberValue || cell.userEnteredValue.boolValue || '')) : formattedValue;
            if (!text) return '';

            let html = '';
            const runs = cell.textFormatRuns;
            
            for (let i = 0; i < runs.length; i++) {
                const run = runs[i];
                const nextRun = runs[i + 1];
                const start = run.startIndex || 0;
                const end = nextRun ? (nextRun.startIndex || text.length) : text.length;
                const segment = text.substring(start, end);
                
                if (!segment) continue;

                const format = run.format || {};
                
                // Split segment into content and trailing newlines to keep tags clean
                // This ensures <br> tags added later remain outside the formatted spans
                const newlineMatch = segment.match(/^([\s\S]*?)(\n+)$/);
                let content = segment;
                let trailing = '';
                if (newlineMatch) {
                    content = newlineMatch[1];
                    trailing = newlineMatch[2];
                }

                if (!content && !trailing) continue;

                let wrapperStart = '';
                let wrapperEnd = '';
                let style = [];

                if (content) {
                    // Formatting tags (Bold, Italic, etc.)
                    if (format.underline) {
                        wrapperStart += '<u>';
                        wrapperEnd = '</u>' + wrapperEnd;
                    }
                    if (format.strikethrough) {
                        wrapperStart += '<s>';
                        wrapperEnd = '</s>' + wrapperEnd;
                    }
                    if (format.italic) {
                        wrapperStart += '<i>';
                        wrapperEnd = '</i>' + wrapperEnd;
                    }
                    if (format.bold) {
                        wrapperStart += '<b>';
                        wrapperEnd = '</b>' + wrapperEnd;
                    }

                    // Hyperlink (Wraps the formatting tags)
                    if (format.link && format.link.uri) {
                        wrapperStart = `<a href="${format.link.uri}" target="_blank">` + wrapperStart;
                        wrapperEnd = wrapperEnd + '</a>';
                    }

                    // Style properties (Outermost wrapper)
                    if (format.fontSize) {
                        style.push(`font-size: ${format.fontSize}pt`);
                    }
                    if (format.fontFamily) {
                        // Automatically load the font if needed
                        loadGoogleFont(format.fontFamily);
                        // Use single quotes for font name to avoid double-quote collision in HTML attributes
                        style.push(`font-family: '${format.fontFamily}'`);
                    }
                    if (format.foregroundColor) {
                        const color = format.foregroundColor;
                        const r = Math.round((color.red || 0) * 255);
                        const g = Math.round((color.green || 0) * 255);
                        const b = Math.round((color.blue || 0) * 255);
                        style.push(`color: rgb(${r},${g},${b})`);
                    }

                    // Apply style span
                    if (style.length > 0) {
                        wrapperStart = `<span style="${style.join('; ')}">` + wrapperStart;
                        wrapperEnd = wrapperEnd + '</span>';
                    }
                }
                
                html += wrapperStart + content + wrapperEnd + trailing;
            }
            
            return html;
        }

        function processRichTextResponse(data) {
            if (!data.sheets || !data.sheets[0] || !data.sheets[0].data || !data.sheets[0].data[0].rowData) {
                console.error('[Accordion] Invalid data structure:', data);
                throw new Error('Invalid data structure from Google Sheets API');
            }

            const rows = data.sheets[0].data[0].rowData;
            const processedValues = [];

            rows.forEach(row => {
                if (!row.values) {
                    processedValues.push([]);
                    return;
                }

                const rowValues = row.values.map(cell => applyTextFormatting(cell));
                processedValues.push(rowValues);
            });

            return processedValues;
        }

        async function loadData() {
            const content = document.getElementById(INSTANCE_ID + '-content');
            if (!content) {
                console.error('[Accordion] Content element not found');
                return;
            }
            
            content.innerHTML = '<div class="loading-message"><div class="spinner"></div><p>Loading data...</p></div>';
            console.log('[Accordion] Starting data fetch');
            
            try {
                // Fetches rich text, formatted data, and effective formats
                const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}?ranges=${CONFIG.RANGE}&includeGridData=true&fields=sheets(data(rowData(values(userEnteredValue%2CtextFormatRuns%2CformattedValue%2CeffectiveFormat(textFormat)))))&key=${CONFIG.API_KEY}`;
                const fetchUrl = CONFIG.USE_CORS_PROXY ? `https://cors-anywhere.herokuapp.com/${url}` : url;
                
                const response = await fetch(fetchUrl);
                
                if (!response.ok) {
                    const errorMsg = `HTTP error! status: ${response.status}`;
                    console.error('[Accordion] Fetch failed:', errorMsg);
                    throw new Error(errorMsg);
                }
                
                const data = await response.json();
                console.log('[Accordion] Data received');
        
                if (!data.sheets || data.sheets.length === 0) {
                    throw new Error('No data found in spreadsheet');
                }

                const values = processRichTextResponse(data);
                console.log('[Accordion] Processed', values.length, 'rows from rich text');
        
                processData(values);
                
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
            
            // Calculate where data starts (excluding header row if enabled)
            let dataStartIndex = Math.max(0, CONFIG.STARTING_ROW - 1);
            
            let processedData = [];

            // If HEADER_ROW_NUMBER is enabled (e.g., 1), we must ensure that row is included 
            // in the values passed to displayAccordion, even if it's before the STARTING_ROW.
            if (CONFIG.HEADER_ROW_NUMBER > 0 && CONFIG.HEADER_ROW_NUMBER < CONFIG.STARTING_ROW) {
                const headerRowIndex = CONFIG.HEADER_ROW_NUMBER - 1;
                const headerRow = values[headerRowIndex];
                const dataRows = values.slice(dataStartIndex);
                processedData = [headerRow, ...dataRows];
            } else {
                // Standard behavior: start from STARTING_ROW
                // Special case: skip row 1 if it's an implicit header (HEADER_ROW_NUMBER: 0, STARTING_ROW: 1)
                if (CONFIG.HEADER_ROW_NUMBER === 0 && CONFIG.STARTING_ROW === 1) {
                    dataStartIndex = 1;
                }
                processedData = values.slice(dataStartIndex);
            }
            
            allData = processedData;

            // Calculate date range if heat map is enabled (supports true or 1)
            if (CONFIG.DATE_DOTS_HEAT_MAP && (CONFIG.DATE_DOTS_HEAT_MAP.ENABLED === true || CONFIG.DATE_DOTS_HEAT_MAP.ENABLED === 1)) {
                const dateCols = CONFIG.DATE_DOTS_HEAT_MAP.DATE_COLUMNS.map(col => columnLetterToIndex(col));
                dateRange = { min: Infinity, max: -Infinity };
                
                allData.forEach((row, index) => {
                    // Skip header row in calculation
                    let isHeaderRow = false;
                    if (CONFIG.HEADER_ROW_NUMBER !== 0) {
                        if (CONFIG.HEADER_ROW_NUMBER < CONFIG.STARTING_ROW) {
                            isHeaderRow = (index === 0);
                        } else {
                            const actualRowNumber = CONFIG.STARTING_ROW + index;
                            isHeaderRow = (actualRowNumber === CONFIG.HEADER_ROW_NUMBER);
                        }
                    }
                    if (isHeaderRow) return;

                    dateCols.forEach(colIdx => {
                        const val = row[colIdx];
                        if (val) {
                            // Strip HTML
                            const temp = document.createElement('div');
                            temp.innerHTML = val;
                            const text = temp.textContent || temp.innerText || '';
                            const time = new Date(text).getTime();
                            if (!isNaN(time)) {
                                dateRange.min = Math.min(dateRange.min, time);
                                dateRange.max = Math.max(dateRange.max, time);
                            }
                        }
                    });
                });
                console.log('[Accordion] Date range calculated:', new Date(dateRange.min), 'to', new Date(dateRange.max));
            }
        
            // When filtering, we must ALWAYS keep the header row if it exists
            const filteredData = applyFilters(allData, CONFIG.FILTER_CONDITIONS);
            console.log('[Accordion] Filtered:', filteredData.length, 'rows');
        
            displayAccordion(filteredData);
            initializeSearch();
        }
        
        function applyFilters(data, filters) {
            if (filters.length === 0) return data;
        
            return data.filter((row, index) => {
                // Always keep the header row if it's enabled
                if (CONFIG.HEADER_ROW_NUMBER !== 0) {
                    // If we prepended the header (HEADER_ROW < STARTING_ROW), it's always at index 0
                    if (CONFIG.HEADER_ROW_NUMBER < CONFIG.STARTING_ROW) {
                        if (index === 0) return true;
                    } else {
                        // Standard calculation
                        const actualRowNumber = CONFIG.STARTING_ROW + index;
                        if (actualRowNumber === CONFIG.HEADER_ROW_NUMBER) {
                            return true;
                        }
                    }
                }
        
                return filters.every(filter => {
                    const columnIndex = filter.column.charCodeAt(0) - 65;
                    // Strip HTML tags for filtering to handle formatted cells
                    let cellValue = (row[columnIndex] || '').toString();
                    if (cellValue.includes('<')) {
                        const temp = document.createElement('div');
                        temp.innerHTML = cellValue;
                        cellValue = temp.textContent || temp.innerText || '';
                    }
                    cellValue = cellValue.trim();
                    
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
            const hasAutoNumbering = CONFIG.AUTO_NUMBER_COLUMN !== null;
            const numColumns = questionColumnIndices.length + (hasAutoNumbering ? 1 : 0);

            // Viewed Status Setup
            const viewedColumnIndex = CONFIG.VIEWED_COLUMN ? columnLetterToIndex(CONFIG.VIEWED_COLUMN) : -1;
            const lastUpdatedColumnIndex = CONFIG.LAST_UPDATED_COLUMN ? columnLetterToIndex(CONFIG.LAST_UPDATED_COLUMN) : -1;
            const viewedData = loadViewedData();
            const viewedTimestamps = Object.values(viewedData);
            const viewedMinTime = viewedTimestamps.length ? Math.min(...viewedTimestamps) : 0;
            const viewedMaxTime = viewedTimestamps.length ? Math.max(...viewedTimestamps) : 0;
            
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
            let rowNumber = CONFIG.AUTO_NUMBER_COLUMN;
            
            console.log(`[Accordion] About to process ${data.length} rows in displayAccordion`);
            data.forEach((row, index) => {
                const answer = (row[answerColumnIndex] || '').trim();
                const hasQuestionContent = questionColumnIndices.some(colIndex => row[colIndex]);
        
                // Check if this is the header row that should always be displayed
                // We need to account for the fact that index 0 might be the header row 
                // if we manually inserted it in processData.
                let isHeaderRow = false;
                if (CONFIG.HEADER_ROW_NUMBER !== 0) {
                    // If the header row was prepended, index 0 is always the header
                    if (CONFIG.HEADER_ROW_NUMBER < CONFIG.STARTING_ROW) {
                        isHeaderRow = (index === 0);
                    } else {
                        const actualRowNumber = CONFIG.STARTING_ROW + index;
                        isHeaderRow = (actualRowNumber === CONFIG.HEADER_ROW_NUMBER);
                    }
                }

                if (hasQuestionContent || isHeaderRow) {
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
                        // If header row, show blank cell. Otherwise show row number.
                        const cellContent = isHeaderRow ? '' : rowNumber;
                        html += `<td align="center" class="accordion-auto-number-cell" style="width: ${autoNumberWidth}px; max-width: ${autoNumberWidth}px;">${cellContent}</td>`;
                    }
                    
                    let prefixApplied = false;
                    questionColumnIndices.forEach((colIndex, i) => {
                        // Check if this is the special SHOW_HIDE_ICON column
                        if (colIndex === 'SHOW') {
                            const alignment = alignmentToUse[i] || 'center';
                            const columnWidth = CONFIG.COLUMN_WIDTHS && CONFIG.COLUMN_WIDTHS[i] ? CONFIG.COLUMN_WIDTHS[i] : '';
                            const widthStyle = columnWidth ? ` width: ${columnWidth}px; max-width: ${columnWidth}px;` : '';
                            
                            // Only show icon if there is an answer AND it's not the header row
                            let iconHtml = '';
                            if (answer && !isHeaderRow) {
                                // Icon wrapper with initial rotation
                                const iconHiddenDir = (CONFIG.SHOW_HIDE_DIRECTION_HIDDEN || 'right').toLowerCase();
                                const iconShownDir = (CONFIG.SHOW_HIDE_DIRECTION_SHOWN || 'down').toLowerCase();
                                
                                // Use SVG from config or fallback to a default banner
                                let iconContent = CONFIG.SHOW_HIDE_ICON_SVG;
                                
                                // Fallback SVG if config variable is missing or empty
                                if (!iconContent) {
                                    // This SVG is a centered chevron pointing right by default (0 degrees)
                                    iconContent = `<svg viewBox="0 0 24 24" width="100%" height="100%" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"></polyline></svg>`;
                                }
                                
                                iconHtml = `<div class="accordion-toggle-container">
                                                <span class="accordion-toggle-icon">${iconContent}</span>
                                            </div>`;
                            } else {
                                // Empty cell for alignment, but attach data attributes just in case styles depend on them
                                const iconHiddenDir = (CONFIG.SHOW_HIDE_DIRECTION_HIDDEN || 'right').toLowerCase();
                                const iconShownDir = (CONFIG.SHOW_HIDE_DIRECTION_SHOWN || 'down').toLowerCase();
                                iconHtml = `<div class="accordion-toggle-container" style="visibility: hidden;"></div>`; 
                            }

                            // We still need the data-attributes on the cell for the toggle logic to not break, 
                            // even if invisible
                            const iconHiddenDir = (CONFIG.SHOW_HIDE_DIRECTION_HIDDEN || 'right').toLowerCase();
                            const iconShownDir = (CONFIG.SHOW_HIDE_DIRECTION_SHOWN || 'down').toLowerCase();

                            html += `<td align="${alignment}" class="accordion-toggle-cell" style="${widthStyle}" data-hidden-dir="${iconHiddenDir}" data-shown-dir="${iconShownDir}">
                                        ${iconHtml}
                                     </td>`;
                            return;
                        }

                        let cellValue = row[colIndex] || '';
                        
                        // Strip HTML tags for prefix detection to handle formatted cells
                        let cleanCellValue = cellValue.toString();
                        if (cleanCellValue.includes('<')) {
                            const temp = document.createElement('div');
                            temp.innerHTML = cleanCellValue;
                            cleanCellValue = temp.textContent || temp.innerText || '';
                        }
                        cleanCellValue = cleanCellValue.trim();

                        // Check if cell contains an image URL first
                        const directImageUrl = isDirectImageURL(cleanCellValue);

                        // Add question prefix to the first non-image question column (but not for header rows or image thumbnails)
                        let questionPrefix = '';
                        if (!prefixApplied && CONFIG.QUESTION_PREFIX && !isHeaderRow && !directImageUrl) {
                            questionPrefix = CONFIG.QUESTION_PREFIX;
                            prefixApplied = true;
                        }

                        // Optional URL for this question column (from CONFIG.URL_COLUMNS)
                        let linkedValue = '';
                        const urlColIndex = urlColumnIndices[i];
                        const rawUrl = (urlColIndex !== undefined && urlColIndex !== 'SHOW' && typeof urlColIndex === 'number') ? (row[urlColIndex] || '').toString() : '';
                        
                        // Strip HTML from rawUrl in case it's a formatted cell
                        let cleanUrl = rawUrl;
                        if (cleanUrl.includes('<')) {
                            const temp = document.createElement('div');
                            temp.innerHTML = cleanUrl;
                            cleanUrl = temp.textContent || temp.innerText || '';
                        }
                        cleanUrl = cleanUrl.trim();
                        const hasUrl = cleanUrl && /^https?:\/\//i.test(cleanUrl);
                        
                        if (directImageUrl && hasUrl) {
                            // Cell has image URL AND a link URL - wrap image in link
                            // Note: directImageUrl is an array here
                            const imageTag = generateImageTag(directImageUrl[0], true);
                            linkedValue = (isHeaderRow ? "" : questionPrefix) + `<a href="${cleanUrl}" target="_blank">${imageTag}</a>`;
                        } else if (hasUrl) {
                            // Cell has link URL but no image - wrap text in link
                            const processedValue = convertURLsToLinks(cellValue, true, false);
                            linkedValue = (isHeaderRow ? "" : questionPrefix) + `<a href="${cleanUrl}" target="_blank">${processedValue}</a>`;
                        } else {
                            // No link URL - use existing URL/image detection. No lazy loading for question cells.
                            const processedValue = convertURLsToLinks(cellValue, true, false);
                            linkedValue = (isHeaderRow ? "" : questionPrefix) + processedValue;
                        }

                        // Date Dot Heat Map logic (supports true or 1)
                        if (CONFIG.DATE_DOTS_HEAT_MAP && (CONFIG.DATE_DOTS_HEAT_MAP.ENABLED === true || CONFIG.DATE_DOTS_HEAT_MAP.ENABLED === 1) && !isHeaderRow) {
                            const columnLetter = CONFIG.QUESTION_COLUMNS[i];
                            if (CONFIG.DATE_DOTS_HEAT_MAP.DATE_COLUMNS.includes(columnLetter)) {
                                // Strip HTML to get raw date text
                                const temp = document.createElement('div');
                                temp.innerHTML = cellValue;
                                const text = temp.textContent || temp.innerText || '';
                                const dotColor = getHeatMapColor(text, dateRange.min, dateRange.max);
                                if (dotColor) {
                                    const title = CONFIG.DATE_DOTS_HEAT_MAP.TITLE || '';
                                    linkedValue = `<div class="date-cell-wrapper">${linkedValue}<span class="date-dot" style="background-color: ${dotColor}" title="${title}"></span></div>`;
                                }
                            }
                        }

                        // Check for Header Cell Suppression
                        if (isHeaderRow && CONFIG.HEADER_CELL_SUPRESSION && Array.isArray(CONFIG.HEADER_CELL_SUPRESSION)) {
                            const colID = CONFIG.QUESTION_COLUMNS[i];
                            if (CONFIG.HEADER_CELL_SUPRESSION.includes(colID)) {
                                linkedValue = ''; // Suppress text
                            }
                        }

                        const alignment = alignmentToUse[i] || 'left';
                        
                        // Add Viewed Badge if applicable
                        if (colIndex === viewedColumnIndex) {
                            const rawText = cellValue ? cellValue.replace(/<[^>]*>/g, '').trim() : '';
                            if (rawText && viewedData[rawText]) {
                                let badges = createViewedBadge(viewedData[rawText], viewedMinTime, viewedMaxTime);
                                
                                // Check for Unseen Changes
                                if (lastUpdatedColumnIndex !== -1) {
                                    const lastUpdatedRaw = (row[lastUpdatedColumnIndex] || '').replace(/<[^>]*>/g, '').trim();
                                    const lastUpdatedTime = new Date(lastUpdatedRaw).getTime();
                                    if (!isNaN(lastUpdatedTime) && lastUpdatedTime > viewedData[rawText]) {
                                        badges = createUnseenBadge() + badges;
                                    }
                                }
                                
                                const rowIdSafe = rawText.replace(/"/g, '"');
                                linkedValue = `<span class="badge-container" data-row-id="${rowIdSafe}">${badges}</span>` + linkedValue;
                            } else if (rawText) {
                                // Container for potential future badge
                                const rowIdSafe = rawText.replace(/"/g, '"');
                                linkedValue = `<span class="badge-container" data-row-id="${rowIdSafe}"></span>` + linkedValue;
                            }
                        }

                        // Check if cell contains only an image
                        const isImageCell = linkedValue.includes('class="accordion-image-content"');
                        const cellClass = isImageCell ? 'accordion-image-cell' : '';
                        
                        // Get column width (use array index directly, no offset for auto-numbering)
                        const columnWidth = CONFIG.COLUMN_WIDTHS && CONFIG.COLUMN_WIDTHS[i] ? CONFIG.COLUMN_WIDTHS[i] : '';
                        let styleParts = [];
                        if (columnWidth) {
                            styleParts.push(`width: ${columnWidth}px`);
                            styleParts.push(`max-width: ${columnWidth}px`);
                        }
                        
                        // Match header row styling to question rows
                        if (isHeaderRow) {
                            styleParts.push('padding-top: 10px');
                            styleParts.push('padding-bottom: 10px');
                            styleParts.push('font-size: var(--question-font-size)');
                            styleParts.push('color: var(--question-font-color)');
                        }

                        if (colIndex === viewedColumnIndex) {
                            styleParts.push('position: relative');
                        }

                        const styleAttribute = styleParts.length > 0 ? ` style="${styleParts.join('; ')}"` : '';
                        
                        html += `<td align="${alignment}" class="${cellClass}"${styleAttribute}>${linkedValue}</td>`;
                    });
                    html += '</tr>';
                    
                    if (answer) {
                        const answerPrefix = CONFIG.ANSWER_PREFIX || '';
                        let processedAnswer = convertURLsToLinks(answer, false, true);
                        
                        // FINAL CLEANUP: Ensure no lingering empty spans or tags remain from rich text runs
                        processedAnswer = processedAnswer.replace(/<span><\/span>/gi, '');
                        processedAnswer = processedAnswer.replace(/<u><\/u>/gi, '');
                        processedAnswer = processedAnswer.replace(/<b><\/b>/gi, '');

                        // Get enlarged thumbnail if enabled
                        let enlargedThumbnail = '';
                        if (CONFIG.SHOW_ENLARGED_THUMBNAIL_IN_ANSWER_ROW) {
                            // Find the leftmost column with an image
                            for (let i = 0; i < questionColumnIndices.length; i++) {
                                const colIndex = questionColumnIndices[i];
                                const cellValue = row[colIndex] || '';
                                
                                // Strip HTML tags for image detection to handle formatted cells
                                let cleanCellValue = cellValue.toString();
                                if (cleanCellValue.includes('<')) {
                                    const temp = document.createElement('div');
                                    temp.innerHTML = cleanCellValue;
                                    cleanCellValue = temp.textContent || temp.innerText || '';
                                }
                                cleanCellValue = cleanCellValue.trim();

                                const imageUrls = isDirectImageURL(cleanCellValue);
                                
                                if (imageUrls && imageUrls.length > 0) {
                                    // Check if this column has a mapped URL
                                    const urlColIndex = urlColumnIndices[i];
                                    const rawUrl = (urlColIndex !== undefined && urlColIndex !== 'SHOW' && typeof urlColIndex === 'number') ? (row[urlColIndex] || '').toString() : '';
                                    
                                    let cleanUrl = rawUrl;
                                    if (cleanUrl.includes('<')) {
                                        const temp = document.createElement('div');
                                        temp.innerHTML = cleanUrl;
                                        cleanUrl = temp.textContent || temp.innerText || '';
                                    }
                                    cleanUrl = cleanUrl.trim();
                                    const hasUrl = cleanUrl && /^https?:\/\//i.test(cleanUrl);

                                    // Stack all images from the first column that has images
                                    enlargedThumbnail = '<div class="accordion-answer-thumbnail">';
                                    imageUrls.forEach(url => {
                                        const imgTag = `<img src="${url}" alt="Enlarged thumbnail" loading="lazy" style="display: block; margin-bottom: 5px;">`;
                                        if (hasUrl) {
                                            enlargedThumbnail += `<a href="${cleanUrl}" target="_blank">${imgTag}</a>`;
                                        } else {
                                            enlargedThumbnail += imgTag;
                                        }
                                    });
                                    enlargedThumbnail += '</div>';
                                    break; // Use images from the first (leftmost) column found
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
                    
                    // Only increment row number if this wasn't a header row
                    if (!isHeaderRow) {
                        rowNumber++;
                    }
                }
            });
            
            html += '</table></div>';
            content.innerHTML = html;
            
            // Add hover event listeners for lazy loading
            addHoverListeners();
            
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
                    
                    // Start Viewed Timer
                    const badgeContainer = item.querySelector('.badge-container');
                    const rowId = badgeContainer ? badgeContainer.dataset.rowId : null;
                    
                    if (rowId) {
                        const cleanRowId = rowId;
                        
                        if (CONFIG.VIEWED_DELAY > 0) {
                            const timerId = setTimeout(() => {
                                const viewedData = loadViewedData();
                                viewedData[cleanRowId] = Date.now();
                                saveViewedData(viewedData);
                                refreshAllViewedBadges();
                                viewedTimers.delete(index);
                            }, CONFIG.VIEWED_DELAY * 1000);
                            viewedTimers.set(index, timerId);
                        } else {
                            // Immediate
                            const viewedData = loadViewedData();
                            viewedData[cleanRowId] = Date.now();
                            saveViewedData(viewedData);
                            refreshAllViewedBadges();
                        }
                    }

                    item.classList.add('active');
                } else {
                    // Closing - remove instantly
                    
                    // Cancel Viewed Timer
                    if (viewedTimers.has(index)) {
                        clearTimeout(viewedTimers.get(index));
                        viewedTimers.delete(index);
                    }

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
            loadData();
        }
    }

    // Start polling for requirements
    initialize();

})();
