# Google Sheets to HTML Table (Accordion)

A lightweight, highly configurable JavaScript library that transforms Google Sheets data into an interactive, searchable HTML accordion table. Perfect for FAQs, knowledge bases, and data directories.

## 🚀 Features

### 📊 Data Integration
- **Direct Google Sheets Sync**: Fetches data in real-time using the Google Sheets API v4.
- **Rich Text Support**: Preserves formatting (bold, italic, underline, colors, and font families) directly from your spreadsheet.
- **Advanced Filtering**: Server-side style filtering based on custom conditions (equals, contains, greater than, etc.).

### 🔍 Search & Discovery
- **Instant Search**: Real-time filtering as you type with configurable delay.
- **Search Highlighting**: Automatically highlights matched terms within the table.
- **Common Search Tags**: Quick-click tags for frequently searched terms.
- **Animated Search Box**: Eye-catching typewriter animation for the search placeholder.
- **Global Search**: Support for searching across multiple table instances on the same page.
- **Viewed Status Tracking**: Map a specific column to track previously expanded answer rows by adding a 'Previously Viewed' Badge in the top right corner of the mapped cell which also has a configurable "viewed" delay. The background color also has a Heat Map feature which by default changes the color to blue for least recently viewed expanded answer rows and red for most recently viewed expanded answer rows.
- **Unseen Changes Badge (New Badge)**: Automatically displays a 'New' badge (or custom text) when an entry's 'Last Updated' date in the spreadsheet is more recent than the user's 'Previously Viewed' date stored in their local storage. This helps users quickly identify fresh or updated content since their last visit.

### 🖼️ Media & Links
- **Auto Image Detection**: Automatically converts image URLs or `=IMAGE()` formulas into thumbnails. Supports multiple image URLs per cell (one per line).
- **Enlarged Answer Thumbnails**: Optionally display all images from a row's leftmost image column as a vertically stacked gallery in the expanded answer row.
- **YouTube Embedding**: Detects YouTube URLs and converts them into embedded video players while also preserving time stamps in seconds. ie: ( ?t=123s ). Also works for Shorts and Playlists. Note: YouTube Playlist URL's don't need the ( &index= ) parameter. You can still use it if you want but the index param will be stripped out of the url to keep things neat and tidy. You can also load the default video you want to start playing in the playlist by defining that video using the ( ?v= ) param. 
- **Lazy Loading**: High-performance loading for images and videos (loads only when hovered or clicked).
- **Auto-Linking**: Converts plain text URLs into clickable, trackable links.

### 🎨 Design & Animation
- **Highly Configurable**: Control everything from column widths and alignment to custom prefixes.
- **Interactive Animations**: Smooth accordion transitions with multiple easing effects (Smooth, Snap, Bounce, etc.).
- **Text Entry Effects**: Stylish entry animations for expanded content (Fade, Slide, Zoom, Blur, Rotate, Elastic).
- **Date Dots Heat Mapping**: Automatically visualizes data recency with colored dots which by default changes the dot color to blue for least recent dates and red for most recent dates.
- **Responsive Layout**: Designed to work across different screen sizes with a clean, modern look.
- **Auto-Numbering**: Optional automatic row numbering toggle on/off JavaScript variable.

### 🛠️ Technical Highlights
- **No Dependencies**: Pure vanilla JavaScript - no jQuery or external libraries required.
- **Dynamic Cache-Busting Loader**: Features a tiny, uncacheable loader script that automatically detects new versions via GitHub Actions and forces a refresh of CSS and JS assets only when a new version (e.g., `v1.11`) is released.
- **Ordered DOM Injection (CSS Sandwich 🥪)**: The loader uses precise anchor points to ensure correct style cascading. It detects `<style id="GSheets-To-HTML-Table-Overrides">` to inject `root.css` **before** it and `rules.css` **after** it, guaranteeing your overrides always work.
- **Logical JS Anchoring**: Main logic is injected immediately after `<script id="GSheets-To-HTML-Table-Config">`, maintaining a clean and predictable DOM structure.
- **Multi-Instance Support**: Run multiple independent tables on a single page using unique instance IDs.
- **Header Management**: Keep headers visible during search or suppress specific header cells for a cleaner look.
- **Automatic Font Loading**: Automatically loads required Google Fonts specified in your spreadsheet.
- **CORS Proxy Support**: Built-in option for handling CORS issues when necessary.
- **Auto-Retry**: Robust error handling with automatic reconnection attempts on connection failure.

### 🚀 Automatic Expansion
- **Smart Page-Load Expansion**: Control the initial state of your table with the `SHOW_HIDE_ON_PAGE_LOAD` setting:
    - 'show': Expands all rows.
    - 'hide': All rows start collapsed (standard).
    - 'random': Automatically selects **exactly one** random data row to expand.
    - 'show#X': Expands the **X-th data row** (relative numbering, skips headers).
    - 'show>X': **Offset expansion**—skips the first X data rows and expands all remaining rows.
    - 'hide>X': **Threshold expansion**—hides everything if there are more than X rows, otherwise shows all (perfect for single-result queries).
    - 'show=[Substring]': Automatically expands any row where the question columns contain the specified text (**case-insensitive**).
- **URL Parameter Sharing**: Link directly to specific answers using URL parameters (e.g., '?show=Intro'). The script intelligently detects whether you are searching by index, keyword, or substring.
- **Dynamic Parameter Naming**: Customize the URL variable name using 'SHOW_GET_VAR_NAME' to avoid conflicts with other scripts on your site.
    - *Example*: Setting SHOW_GET_VAR_NAME: 'expand' or 'row' allows you to link to specific answers using 'yourpage.html?expand=Intro' or 'yourpage.html?row=Intro' instead of the default '?show=Intro'.
- **Immediate Media Loading**: Programmatically expanded rows automatically trigger lazy-loading for YouTube embeds and images, ensuring they appear immediately on page load.

## 📁 Project Structure

- `Cache-Busting-Loader.js`: The "smart" entry point. It manages versioning and dynamically builds the "CSS Sandwich" around your overrides.
- `GSheets-To-HTML-Table.js`: Core logic for fetching and rendering data.
- `GSheets-To-HTML-Table-root.css`: Global variables and base styling.
- `GSheets-To-HTML-Table-rules.css`: Layout and component-specific styling.
- `faq.html`: Example implementation for a Frequently Asked Questions page.
- `supplements.html`: Example implementation for supplemental data display.

## 🛠️ Quick Start

1. **Prepare your Google Sheet**: Ensure it is shared (Anyone with the link can view).
2. **Get an API Key**: Obtain a Google Cloud API Key with Google Sheets API enabled.
3. **Configuration & Wrapper**: Add the config block, optional override style tag, and the wrapper div to your HTML.
4. **Include the Loader**: Use the Dynamic Loader script to automatically pull in the latest CSS and JS.

```html
<!-- 1. Configuration Block -->
<script id="GSheets-To-HTML-Table-Config">
    const CONFIG = {
        SPREADSHEET_ID: 'YOUR_ID_HERE',
        API_KEY: 'YOUR_KEY_HERE',
        // ... see faq.html or supplements.html for full config options
    };
</script>

<!-- 2. CSS Overrides (Optional but Recommended if needed) -->
<!-- Identify this tag with id="GSheets-To-HTML-Table-Overrides" for correct load order -->
<style id="GSheets-To-HTML-Table-Overrides">
    :root {
        --question-padding-horizontal: 0px; 
        --question-padding-vertical: 0px;   
    }
</style>

<!-- 3. The Accordion Wrapper -->
<div class="accordion-wrapper"></div>

<!-- 4. The Dynamic Loader (Automated Cache Busting) -->
<script>
    (function() {
        const loader = document.createElement('script');
        loader.src = 'https://automybiz.github.io/GSheets-To-HTML-Table/latest/Cache-Busting-Loader.js?t=' + Date.now();
        document.head.appendChild(loader);
    })();
</script>
```

### ⚙️ Configuration Options

For more info on config options please see the HTML file examples : ))
