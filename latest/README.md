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

### 🖼️ Media & Links
- **Auto Image Detection**: Automatically converts image URLs or `=IMAGE()` formulas into thumbnails.
- **YouTube Embedding**: Detects YouTube URLs and converts them into embedded video players while also preserving time stamps in seconds. ie: ?t=123s
- **Lazy Loading**: High-performance loading for images and videos (loads only when hovered or clicked).
- **Auto-Linking**: Converts plain text URLs into clickable, trackable links.

### 🎨 Design & Animation
- **Highly Configurable**: Control everything from column widths and alignment to custom prefixes.
- **Interactive Animations**: Smooth accordion transitions with multiple easing effects (Smooth, Snap, Bounce, etc.).
- **Text Entry Effects**: Stylish entry animations for expanded content (Fade, Slide, Zoom, Blur, Rotate, Elastic).
- **Responsive Layout**: Designed to work across different screen sizes with a clean, modern look.
- **Auto-Numbering**: Optional automatic row numbering toggle on/off JavaScript variable.

### 🛠️ Technical Highlights
- **No Dependencies**: Pure vanilla JavaScript - no jQuery or external libraries required.
- **Multi-Instance Support**: Run multiple independent tables on a single page using unique instance IDs.
- **Automatic Font Loading**: Automatically loads required Google Fonts specified in your spreadsheet.
- **CORS Proxy Support**: Built-in option for handling CORS issues when necessary.
- **Auto-Retry**: Robust error handling with automatic reconnection attempts on connection failure.

## 📁 Project Structure

- `GSheets-To-HTML-Table.js`: Core logic for fetching and rendering data.
- `GSheets-To-HTML-Table-root.css`: Global variables and base styling.
- `GSheets-To-HTML-Table-rules.css`: Layout and component-specific styling.
- `faq.html`: Example implementation for a Frequently Asked Questions page.
- `supplements.html`: Example implementation for supplemental data display.

## 🛠️ Quick Start

1. **Prepare your Google Sheet**: Ensure it is shared (Anyone with the link can view).
2. **Get an API Key**: Obtain a Google Cloud API Key with Google Sheets API enabled.
3. **Configure**: Update the `CONFIG` object in your JS file with your `SPREADSHEET_ID` and `API_KEY`.
4. **Include Files**: Link the CSS and JS files in your HTML.

```html
    <link rel="stylesheet" href="GSheets-To-HTML-Table-root.css">
    <style>
        /* Overrides */
        :root {
            --question-padding-horizontal: 0px; /* No padding needed since thumbnail column is being used */
            --question-padding-vertical: 0px;   /* No padding wanted around the thumbnails */
        }
    </style>
    <link rel="stylesheet" href="GSheets-To-HTML-Table-rules.css">

    <div class="accordion-wrapper"></div>

    <script src="GSheets-To-HTML-Table.js"></script>
```
