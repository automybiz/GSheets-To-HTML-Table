(function() {
    /**
     * Global Image Zoom Overlay
     * A standalone script to add popup zoom functionality to images.
     * 
     * Usage: <script src="Global-Image-Zoom.js?classes=hl-optimized&mobile=false&gallery_mode=enabled"></script>
     */

    // 1. Parse Script Parameters
    const scriptTag = document.currentScript;
    const scriptSrc = scriptTag ? scriptTag.src : '';
    const urlParams = new URLSearchParams(scriptSrc.split('?')[1] || '');
    
    const TARGET_CLASSES_PARAM = urlParams.get('classes') || '';
    const TARGET_IMAGE_CLASSES = TARGET_CLASSES_PARAM ? TARGET_CLASSES_PARAM.split(',').map(c => c.trim()).filter(c => c !== '') : [];
    
    const ENABLE_MOBILE_PARAM = urlParams.get('mobile');
    const ENABLE_POPUP_IMAGE_ZOOM_ON_MOBILE = ENABLE_MOBILE_PARAM === 'true'; // Default is false unless specified true

    const GALLERY_MODE_PARAM = urlParams.get('gallery_mode');
    const ENABLE_GALLERY_MODE = GALLERY_MODE_PARAM !== 'disabled'; // Default is true unless specified disabled

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
            --global-overlay-nav-button-bg: #000;
            --global-overlay-nav-button-bg-opacity: 0.5;
            --global-overlay-nav-button-hover-bg: #000;
            --global-overlay-nav-button-hover-bg-opacity: 0.8;
            --global-overlay-nav-button-color: #FFF;
            --global-overlay-nav-button-size: 80px;
            --global-overlay-thumbnail-size: 80px;
            --global-overlay-thumbnail-border: 2px solid #088;
            --global-overlay-thumbnail-active-border: 2px solid #0FF;
            --global-overlay-thumbnail-strip-bg: #000;
            --global-overlay-thumbnail-strip-bg-opacity: 0.7;
        }

        /* Target Image Hover Effects */
        ${TARGET_IMAGE_CLASSES.length > 0 
            ? TARGET_IMAGE_CLASSES.map(cls => `img.${cls}`).join(', ') 
            : 'img:not(.global-image-overlay-img):not(.global-image-overlay-thumb)'} {
            cursor: zoom-in !important;
            transition: filter 0.3s ease, box-shadow 0.3s ease;
        }

        ${TARGET_IMAGE_CLASSES.length > 0 
            ? TARGET_IMAGE_CLASSES.map(cls => `img.${cls}:hover`).join(', ') 
            : 'img:not(.global-image-overlay-img):not(.global-image-overlay-thumb):hover'} {
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
            z-index: 10;
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
            z-index: 10;
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
            flex: 1;
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

        /* Navigation Buttons */
        .global-image-overlay-nav {
            position: absolute;
            top: 50%;
            transform: translateY(-50%);
            width: var(--global-overlay-nav-button-size);
            height: var(--global-overlay-nav-button-size);
            background: color-mix(in srgb, var(--global-overlay-nav-button-bg), transparent calc(100% * (1 - var(--global-overlay-nav-button-bg-opacity))));
            color: var(--global-overlay-nav-button-color);
            display: none; /* Shown in gallery mode */
            justify-content: center;
            align-items: center;
            cursor: pointer;
            border-radius: 50%;
            z-index: 10;
            transition: background 0.2s ease, transform 0.2s ease;
            user-select: none;
        }

        .global-image-overlay-nav:hover {
            background: color-mix(in srgb, var(--global-overlay-nav-button-hover-bg), transparent calc(100% * (1 - var(--global-overlay-nav-button-hover-bg-opacity))));
            transform: translateY(-50%) scale(1.1);
        }

        .global-image-overlay-nav.prev {
            left: 20px;
        }

        .global-image-overlay-nav.next {
            right: 20px;
        }

        .global-image-overlay-nav svg {
            width: 60%;
            height: 60%;
        }

        /* Thumbnails Strip */
        .global-image-overlay-thumbnails {
            width: 100%;
            height: auto;
            min-height: calc(var(--global-overlay-thumbnail-size) + 20px);
            background: color-mix(in srgb, var(--global-overlay-thumbnail-strip-bg), transparent calc(100% * (1 - var(--global-overlay-thumbnail-strip-bg-opacity))));
            display: none; /* Shown in gallery mode */
            justify-content: center;
            align-items: center;
            gap: 10px;
            padding: 10px;
            box-sizing: border-box;
            overflow-x: auto;
            z-index: 10;
        }

        .global-image-overlay-thumb {
            width: var(--global-overlay-thumbnail-size);
            height: var(--global-overlay-thumbnail-size);
            object-fit: cover;
            cursor: pointer;
            border: var(--global-overlay-thumbnail-border);
            transition: transform 0.2s ease;
            border-radius: 4px;
            flex-shrink: 0;
        }

        .global-image-overlay-thumb:hover {
            transform: scale(1.1);
        }

        .global-image-overlay-thumb.active {
            border: var(--global-overlay-thumbnail-active-border);
        }
    `;

    const styleTag = document.createElement('style');
    styleTag.textContent = css;
    document.head.appendChild(styleTag);

    // 4. Overlay Logic
    let overlay = null;
    let currentScale = 1;
    let galleryImages = [];
    let currentIndex = 0;

    const navIconSvg = `<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M11.25 4.5l7.5 7.5-7.5 7.5m-6-15l7.5 7.5-7.5 7.5" /></svg>`;

    function createOverlay() {
        if (overlay) return;

        overlay = document.createElement('div');
        overlay.id = 'global-image-zoom-overlay';
        overlay.className = 'global-image-overlay';
        
        overlay.innerHTML = `
            <div class="global-image-overlay-zoom-text" id="global-zoom-text">100%</div>
            <div class="global-image-overlay-close">&times;</div>
            
            <div class="global-image-overlay-nav prev" id="global-overlay-prev" style="transform: translateY(-50%) rotate(180deg);">
                ${navIconSvg}
            </div>
            <div class="global-image-overlay-nav next" id="global-overlay-next">
                ${navIconSvg}
            </div>

            <div class="global-image-overlay-img-container" id="global-img-container">
                <div id="global-img-wrapper" style="position: relative; display: inline-block; transition: transform 0.1s ease-out; transform-origin: center center;">
                    <img src="" class="global-image-overlay-img" id="global-main-img">
                </div>
            </div>

            <div class="global-image-overlay-thumbnails" id="global-overlay-thumbnails"></div>
        `;
        document.body.appendChild(overlay);

        const mainImg = overlay.querySelector('#global-main-img');
        const imgWrapper = overlay.querySelector('#global-img-wrapper');
        const zoomText = overlay.querySelector('#global-zoom-text');
        const closeBtn = overlay.querySelector('.global-image-overlay-close');
        const container = overlay.querySelector('#global-img-container');
        const prevBtn = overlay.querySelector('#global-overlay-prev');
        const nextBtn = overlay.querySelector('#global-overlay-next');
        const thumbContainer = overlay.querySelector('#global-overlay-thumbnails');

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
                    mainImg.src = ''; 
                }
            }, 300);
        }

        function showImage(index) {
            if (index < 0) index = galleryImages.length - 1;
            if (index >= galleryImages.length) index = 0;
            currentIndex = index;

            const imgSrc = galleryImages[currentIndex].src || galleryImages[currentIndex];
            mainImg.src = imgSrc;
            resetZoom();

            // Update thumbnails
            const thumbs = thumbContainer.querySelectorAll('.global-image-overlay-thumb');
            thumbs.forEach((t, i) => {
                if (i === currentIndex) {
                    t.classList.add('active');
                    t.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
                } else {
                    t.classList.remove('active');
                }
            });

            // Update nav visibility
            if (galleryImages.length > 1 && ENABLE_GALLERY_MODE) {
                prevBtn.style.display = 'flex';
                nextBtn.style.display = 'flex';
                thumbContainer.style.display = 'flex';
            } else {
                prevBtn.style.display = 'none';
                nextBtn.style.display = 'none';
                thumbContainer.style.display = 'none';
            }
        }

        // Event Listeners
        closeBtn.addEventListener('click', closeOverlay);
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay || e.target === container) closeOverlay();
        });

        prevBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            showImage(currentIndex - 1);
        });

        nextBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            showImage(currentIndex + 1);
        });

        document.addEventListener('keydown', (e) => {
            if (!overlay.classList.contains('active')) return;
            if (e.key === 'Escape') closeOverlay();
            if (ENABLE_GALLERY_MODE) {
                if (e.key === 'ArrowLeft') showImage(currentIndex - 1);
                if (e.key === 'ArrowRight') showImage(currentIndex + 1);
            }
        });

        imgWrapper.addEventListener('dblclick', (e) => {
            e.preventDefault();
            if (currentScale > 1) {
                resetZoom();
            } else {
                const rect = imgWrapper.getBoundingClientRect();
                const xPercent = ((e.clientX - rect.left) / rect.width) * 100;
                const yPercent = ((e.clientY - rect.top) / rect.height) * 100;
                imgWrapper.style.transformOrigin = `${xPercent}% ${yPercent}%`;
                currentScale = 2;
                updateZoom();
            }
        });

        overlay.addEventListener('wheel', (e) => {
            e.preventDefault();
            const rect = imgWrapper.getBoundingClientRect();
            const mouseX = e.clientX - rect.left;
            const mouseY = e.clientY - rect.top;
            const xPercent = (mouseX >= 0 && mouseX <= rect.width) ? (mouseX / rect.width) * 100 : 50;
            const yPercent = (mouseY >= 0 && mouseY <= rect.height) ? (mouseY / rect.height) * 100 : 50;
            imgWrapper.style.transformOrigin = `${xPercent}% ${yPercent}%`;

            const zoomStep = 0.2;
            if (e.deltaY < 0) currentScale = Math.min(5, currentScale + zoomStep);
            else {
                currentScale = Math.max(1, currentScale - zoomStep);
                if (currentScale === 1) imgWrapper.style.transformOrigin = 'center center';
            }
            updateZoom();
        }, { passive: false });

        overlay.open = function(clickedSrc, imagesArray) {
            galleryImages = imagesArray || [clickedSrc];
            
            // Rebuild thumbnails
            thumbContainer.innerHTML = '';
            if (galleryImages.length > 1 && ENABLE_GALLERY_MODE) {
                galleryImages.forEach((imgData, idx) => {
                    const src = imgData.src || imgData;
                    const thumb = document.createElement('img');
                    thumb.src = src;
                    thumb.className = 'global-image-overlay-thumb';
                    thumb.onclick = (e) => {
                        e.stopPropagation();
                        showImage(idx);
                    };
                    thumbContainer.appendChild(thumb);
                });
            }

            const startIdx = galleryImages.findIndex(img => (img.src || img) === clickedSrc);
            showImage(startIdx === -1 ? 0 : startIdx);
            overlay.classList.add('active');
        };
    }

    // 5. Global Click Listener & Gallery Detection
    document.body.addEventListener('click', function(e) {
        const img = e.target.closest('img');
        if (!img) return;

        // Skip internal/accordion elements
        if (img.closest('.image-overlay') || img.closest('.global-image-overlay') || img.closest('.accordion-wrapper')) return;

        // Conflict Prevention with GSheets Table
        if (img.dataset.popupEnabled === 'true' || 
            img.getAttribute('onclick')?.includes('openGSheetsImageOverlay') ||
            img.classList.contains('accordion-image-content') ||
            img.classList.contains('accordion-answer-image') ||
            img.classList.contains('image-overlay-thumb')) {
            return;
        }

        // Class Targeting Check
        let isEligible = true;
        if (TARGET_IMAGE_CLASSES.length > 0) {
            isEligible = TARGET_IMAGE_CLASSES.some(cls => img.classList.contains(cls));
        }
        if (!isEligible) return;

        // Prevent default browser behavior
        e.preventDefault();
        e.stopPropagation();

        // Detect Gallery
        let imagesToGallery = [];
        if (ENABLE_GALLERY_MODE) {
            const allImgs = document.querySelectorAll('img');
            allImgs.forEach(item => {
                // Same eligibility check
                let itemEligible = true;
                if (TARGET_IMAGE_CLASSES.length > 0) {
                    itemEligible = TARGET_IMAGE_CLASSES.some(cls => item.classList.contains(cls));
                }
                // Skip if hidden or part of excluded systems
                const isHidden = item.offsetParent === null;
                const isExcluded = item.closest('.accordion-wrapper') || item.closest('.image-overlay') || item.closest('.global-image-overlay');
                
                if (itemEligible && !isHidden && !isExcluded) {
                    imagesToGallery.push(item.src);
                }
            });
        } else {
            imagesToGallery = [img.src];
        }

        if (!overlay) createOverlay();
        overlay.open(img.src, imagesToGallery);
    }, true);

    // 6. Public API for consolidation
    window.openGlobalImageOverlay = function(src, galleryArray) {
        if (!overlay) createOverlay();
        overlay.open(src, galleryArray);
    };

    console.log('[Global-Image-Zoom] Initialized. Mode:', ENABLE_GALLERY_MODE ? 'Gallery' : 'Single', 'Targets:', TARGET_IMAGE_CLASSES.length > 0 ? TARGET_IMAGE_CLASSES : 'All');
})();
