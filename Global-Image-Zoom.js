(function() {
    /**
     * Global Image Zoom Overlay
     * A standalone script to add popup zoom functionality to images.
     * 
     * Usage: <script src="Global-Image-Zoom.js?classes=hl-optimized&mobile=false"></script>
     */

    // 1. Parse Script Parameters
    const scriptTag = document.currentScript;
    const scriptSrc = scriptTag ? scriptTag.src : '';
    const urlParams = new URLSearchParams(scriptSrc.split('?')[1] || '');
    
    const TARGET_CLASSES_PARAM = urlParams.get('classes') || '';
    const TARGET_IMAGE_CLASSES = TARGET_CLASSES_PARAM ? TARGET_CLASSES_PARAM.split(',').map(c => c.trim()).filter(c => c !== '') : [];
    
    const ENABLE_MOBILE_PARAM = urlParams.get('mobile');
    const ENABLE_POPUP_IMAGE_ZOOM_ON_MOBILE = ENABLE_MOBILE_PARAM === 'true'; // Default is false unless specified true

    // 2. Mobile Detection & Exit
    const isMobile = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) || (window.innerWidth <= 768);
    if (isMobile && !ENABLE_POPUP_IMAGE_ZOOM_ON_MOBILE) {
        console.log('[Global-Image-Zoom] Mobile device detected and feature is disabled via configuration.');
        return;
    }

    // 3. Inject CSS Styles
    const css = `
        :root {
            --global-overlay-bg-color: #000;
            --global-overlay-bg-opacity: 0.9;
            --global-overlay-z-index: 999999;
            --global-overlay-zoom-text-color: #FFF;
            --global-overlay-zoom-text-bg-color: #000;
            --global-overlay-zoom-text-bg-opacity: 0.5;
            --global-overlay-zoom-text-font-size: 18px;
            --global-overlay-image-glow-color: #0FF;
            --global-overlay-image-glow-size: 11px;
            --global-overlay-image-glow-opacity: 0.9;
            --global-overlay-image-glow-hover-color: #0FF;
            --global-overlay-image-glow-hover-size: 33px;
            --global-overlay-image-glow-hover-opacity: 0.9;
        }

        /* Target Image Hover Effects */
        ${TARGET_IMAGE_CLASSES.length > 0 
            ? TARGET_IMAGE_CLASSES.map(cls => `img.${cls}`).join(', ') 
            : 'img:not(.image-overlay-img):not(.image-overlay-thumb)'} {
            cursor: zoom-in !important;
            transition: filter 0.3s ease, box-shadow 0.3s ease;
        }

        ${TARGET_IMAGE_CLASSES.length > 0 
            ? TARGET_IMAGE_CLASSES.map(cls => `img.${cls}:hover`).join(', ') 
            : 'img:not(.image-overlay-img):not(.image-overlay-thumb):hover'} {
            filter: brightness(1.05);
            box-shadow: 0 0 15px rgba(0, 255, 255, 0.3);
        }

        /* Overlay Styles */
        .global-image-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background-color: color-mix(in srgb, var(--global-overlay-bg-color), transparent calc(100% * (1 - var(--global-overlay-bg-opacity))));
            z-index: var(--global-overlay-z-index);
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            opacity: 0;
            transition: opacity 0.3s ease;
            visibility: hidden;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        }

        .global-image-overlay.active {
            opacity: 1;
            visibility: visible;
        }

        .global-image-overlay-close {
            position: absolute;
            top: 20px;
            right: 30px;
            color: var(--global-overlay-zoom-text-color);
            font-size: 30px;
            cursor: pointer;
            z-index: 2;
            background: color-mix(in srgb, var(--global-overlay-zoom-text-bg-color), transparent calc(100% * (1 - var(--global-overlay-zoom-text-bg-opacity))));
            width: 40px;
            height: 40px;
            display: flex;
            justify-content: center;
            align-items: center;
            border-radius: 50%;
            user-select: none;
        }

        .global-image-overlay-zoom-text {
            position: absolute;
            top: 20px;
            left: 50%;
            transform: translateX(-50%);
            color: var(--global-overlay-zoom-text-color);
            background: color-mix(in srgb, var(--global-overlay-zoom-text-bg-color), transparent calc(100% * (1 - var(--global-overlay-zoom-text-bg-opacity))));
            padding: 5px 15px;
            border-radius: 20px;
            font-size: var(--global-overlay-zoom-text-font-size);
            z-index: 2;
            pointer-events: none;
            user-select: none;
        }

        .global-image-overlay-img-container {
            position: relative;
            width: 100%;
            height: 100%;
            display: flex;
            justify-content: center;
            align-items: center;
            overflow: hidden;
        }

        .global-image-overlay-img {
            max-width: 90%;
            max-height: 90%;
            object-fit: contain;
            transition: transform 0.1s ease-out, box-shadow 0.3s ease;
            transform-origin: center center;
            cursor: zoom-in;
            box-shadow: 0 0 var(--global-overlay-image-glow-size) color-mix(in srgb, var(--global-overlay-image-glow-color), transparent calc(100% * (1 - var(--global-overlay-image-glow-opacity))));
        }

        .global-image-overlay-img:hover {
            box-shadow: 0 0 var(--global-overlay-image-glow-hover-size) color-mix(in srgb, var(--global-overlay-image-glow-hover-color), transparent calc(100% * (1 - var(--global-overlay-image-glow-hover-opacity))));
        }
    `;

    const styleTag = document.createElement('style');
    styleTag.textContent = css;
    document.head.appendChild(styleTag);

    // 4. Overlay Logic
    let overlay = null;
    let currentScale = 1;

    function createOverlay() {
        if (overlay) return;

        overlay = document.createElement('div');
        overlay.id = 'global-image-zoom-overlay';
        overlay.className = 'global-image-overlay';
        
        overlay.innerHTML = `
            <div class="global-image-overlay-zoom-text" id="global-zoom-text">100%</div>
            <div class="global-image-overlay-close">&times;</div>
            <div class="global-image-overlay-img-container" id="global-img-container">
                <div id="global-img-wrapper" style="position: relative; display: inline-block; transition: transform 0.1s ease-out; transform-origin: center center;">
                    <img src="" class="global-image-overlay-img" id="global-main-img">
                </div>
            </div>
        `;
        document.body.appendChild(overlay);

        const mainImg = overlay.querySelector('#global-main-img');
        const imgWrapper = overlay.querySelector('#global-img-wrapper');
        const zoomText = overlay.querySelector('#global-zoom-text');
        const closeBtn = overlay.querySelector('.global-image-overlay-close');
        const container = overlay.querySelector('#global-img-container');

        function updateZoom() {
            imgWrapper.style.transform = `scale(${currentScale})`;
            zoomText.textContent = `${Math.round(currentScale * 100)}%`;
            imgWrapper.style.cursor = currentScale > 1 ? 'zoom-out' : 'zoom-in';
        }

        function resetZoom() {
            currentScale = 1;
            imgWrapper.style.transformOrigin = 'center center';
            updateZoom();
        }

        function closeOverlay() {
            overlay.classList.remove('active');
            setTimeout(() => {
                if (!overlay.classList.contains('active')) {
                    mainImg.src = ''; // Clear src when hidden
                }
            }, 300);
        }

        closeBtn.addEventListener('click', closeOverlay);
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay || e.target === container) {
                closeOverlay();
            }
        });

        // ESC Key to close
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && overlay.classList.contains('active')) {
                closeOverlay();
            }
        });

        // Double Click to Zoom
        imgWrapper.addEventListener('dblclick', (e) => {
            e.preventDefault();
            if (currentScale > 1) {
                resetZoom();
            } else {
                const rect = imgWrapper.getBoundingClientRect();
                const xPercent = ((e.clientX - rect.left) / rect.width) * 100;
                const yPercent = ((e.clientY - rect.top) / rect.height) * 100;
                imgWrapper.style.transformOrigin = `${xPercent}% ${yPercent}%`;
                currentScale = 2; // Default 200%
                updateZoom();
            }
        });

        // Mouse Wheel Zooming
        overlay.addEventListener('wheel', (e) => {
            e.preventDefault();
            const rect = imgWrapper.getBoundingClientRect();
            const mouseX = e.clientX - rect.left;
            const mouseY = e.clientY - rect.top;

            const isOverImage = mouseX >= 0 && mouseX <= rect.width && mouseY >= 0 && mouseY <= rect.height;
            const xPercent = isOverImage ? (mouseX / rect.width) * 100 : 50;
            const yPercent = isOverImage ? (mouseY / rect.height) * 100 : 50;

            imgWrapper.style.transformOrigin = `${xPercent}% ${yPercent}%`;

            const zoomStep = 0.2;
            if (e.deltaY < 0) {
                currentScale = Math.min(5, currentScale + zoomStep);
            } else {
                currentScale = Math.max(1, currentScale - zoomStep);
                if (currentScale === 1) imgWrapper.style.transformOrigin = 'center center';
            }
            updateZoom();
        }, { passive: false });

        overlay.open = function(src) {
            mainImg.src = src;
            resetZoom();
            overlay.classList.add('active');
        };
    }

    // 5. Global Click Listener with Conflict Prevention
    document.body.addEventListener('click', function(e) {
        const img = e.target.closest('img');
        if (!img) return;

        // Skip images inside an existing overlay or accordion
        if (img.closest('.image-overlay') || img.closest('.global-image-overlay') || img.closest('.accordion-wrapper')) {
            return;
        }

        // Skip images handled by GSheets To HTML Table
        if (img.dataset.popupEnabled === 'true' || 
            img.getAttribute('onclick')?.includes('openGSheetsImageOverlay') ||
            img.classList.contains('accordion-image-content') ||
            img.classList.contains('accordion-answer-image') ||
            img.classList.contains('image-overlay-thumb')) {
            return;
        }

        // Class Targeting Check
        if (TARGET_IMAGE_CLASSES.length > 0) {
            const hasTargetClass = TARGET_IMAGE_CLASSES.some(cls => img.classList.contains(cls));
            if (!hasTargetClass) return;
        }

        // If we reached here, handle the zoom!
        e.preventDefault();
        e.stopPropagation();
        
        if (!overlay) createOverlay();
        overlay.open(img.src);
    }, true); // Use capture phase to catch clicks early if needed

    console.log('[Global-Image-Zoom] Initialized. Target classes:', TARGET_IMAGE_CLASSES.length > 0 ? TARGET_IMAGE_CLASSES : 'All Images');

})();
