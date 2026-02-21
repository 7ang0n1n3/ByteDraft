# ByteDraft

ByteDraft is a modern technical documentation tool for creating, editing, and exporting structured reports (DOCX) with advanced formatting, unlimited section depth, and a user-friendly interface. Built for offline use with editing capabilities.

## ✨ Features

### 📝 **Advanced Rich Text Editing**
- **Full TinyMCE Integration** with complete menu bar and toolbar
- **Professional Formatting**: Bold, italic, underline, strikethrough, text/background colors
- **Typography Options**: Multiple font families and sizes (8pt to 48pt)
- **Advanced Tools**: Search/replace, fullscreen mode, word count, character map
- **Media Support**: Drag-and-drop image insertion with base64 storage
- **Table Support**: Create and format tables with ease, including cell colors and borders
- **Code Blocks**: Syntax highlighting for code snippets with GitHub Dark theme

### 📚 **Document Structure**
- **Unlimited Nested Sections**: Create complex document hierarchies with automatic numbering
- **Drag & Drop Reordering**: Intuitive drag and drop interface to reorder sections and subsections
- **Professional Heading Hierarchy**: Automatic Heading 1, 2, 3 styles with proper numbering (1., 1.1., 1.1.1.)
- **Auto-updating Table of Contents**: Word-compatible TOC field that updates automatically
- **Citation Manager**: Insert numbered inline citations `[1]`, `[2]` via a toolbar button and manage the full reference list through a dedicated modal; a formatted References page is automatically appended to DOCX exports
- **Footnotes & Endnotes**: Insert footnote `[fn]` (blue) or endnote `[en]` (purple) superscript markers via a dedicated toolbar button; markers are numbered sequentially at DOCX export time and collected onto dedicated Footnotes and Endnotes pages (omitted if none exist)
- **Section Annotations**: Add, resolve, and delete review comments on any section or subsection; persisted in JSON exports, excluded from DOCX
- **Professional Templates**: Pre-built templates for various document types
- **Version History**: Track changes and document evolution
- **Custom Fields**: Add project-specific metadata

### 🖼️ **Image Support**
- **Drag & Drop**: Simply drag images into the editor
- **File Picker**: Browse and select images from your device
- **Paste Support**: Paste images directly from clipboard
- **Base64 Storage**: Images stored locally for offline use
- **DOCX Export**: Images properly included in exported documents with full formatting

### 📄 **Export Capabilities**
- **Modern DOCX Export**: Uses the latest docx library for better compatibility
- **Professional Document Structure**: 
  - Title page with document information (36pt title, 18pt description)
  - Document changelog page with custom fields
  - Auto-updating Table of Contents with "Click here to update" placeholder
  - Properly numbered sections and subsections
- **Advanced Table Export**: Preserves cell colors, borders, sizes, and formatting
- **Image Preservation**: All images included in exported documents
- **Formatting Retention**: Maintains all text formatting and structure (bold, italic, font sizes)
- **Professional Layout**: Proper headings, lists, tables, and styling
- **Document Metadata**: Includes title page, version history, and changelog
- **Custom Headers/Footers**: Support for `{{title}}` and `{{page}}` variables with table-based layout
- **Document Logo Support**: Upload and embed custom logos in DOCX title pages
- **Multi-section Export**: Title page, changelog, and TOC without headers/footers; main content with custom headers/footers

### 🎨 **User Interface**
- **"Slate Pro" Design System**: Refined, professional dark-first UI with indigo accent colors and DM Sans typography
- **Permanent-Dark Collapsible Sidebar**: Always-dark navigation panel that responds to light/dark theme; collapse to zero with a single click and re-expand via a floating edge tab
- **Frosted-Glass Top Bar**: Sticky top bar with `backdrop-filter` blur for a modern layered look
- **Pill-Shaped Buttons**: Primary actions use pill-shaped buttons with smooth hover transitions
- **Refined Modals**: Large border-radius modals with polished shadows and uppercase field labels
- **Responsive Layout**: Works on desktop and tablet devices
- **Project Management**: Organize multiple documentation projects
- **Status Tracking**: Draft, Working, and Publish statuses with luminous pill badges
- **Real-time Preview**: See changes reflected immediately
- **Dark/Light Theme**: Toggle between themes for comfortable editing with persistence; sidebar adapts to both modes
- **Centralized Notifications**: All status messages appear in the top center for consistent user experience
- **Smooth Animations**: Entrance animations for section cards and sidebar collapse/expand transitions

### 💾 **Data Management**
- **Local Storage**: All data stored in your browser (no server required)
- **Auto-save**: Automatic saving every 30 seconds — silently persists data without creating a revision history entry
- **Export/Import**: JSON export for backup and sharing with complete project data preservation
- **Offline Operation**: Works completely without internet connection
- **Document Information**: Custom document metadata fields
- **Changelog Management**: Professional changelog with version tracking
- **Complete Data Preservation**: JSON exports include document info, changelog, headers/footers, and version history

## 📋 Changelog

### v1.4.0
- **Feature**: Footnotes & Endnotes — a new `[N]` TinyMCE toolbar button opens a modal to insert footnote (`[fn]`, blue) or endnote (`[en]`, purple) superscript markers at the cursor; markers are stored as `<sup class="bd-fn">` with `data-note-type` and `data-note` attributes, styled with accent colours in the editor, and preserved across saves via `extended_valid_elements: 'sup[*]'`
- **Export**: At DOCX export time, a `resolveNotes()` pre-pass (DOMParser) numbers footnote markers `[1]`, `[2]`, … and endnote markers `[E1]`, `[E2]`, … sequentially, strips the custom attributes, and collects the note text; a "Footnotes" page (blue title) and an "Endnotes" page (purple title) are appended before the References page, and are omitted entirely if no notes of that type exist

### v1.3.6
- **Fix**: TOC sidebar links now scroll so the section card clears the sticky top bar — replaced `scrollIntoView({block: 'start'})` with a manual `window.scrollTo()` that subtracts the top bar's measured height plus 12 px of breathing room

### v1.3.5
- **Fix**: TinyMCE toolbar icons for Cite, XRef, and `{{}}` now display correctly — the previous inline SVGs used `xmlns`, `viewBox`, and `fill="currentColor"` which TinyMCE 7 ignores, producing invisible icons; replaced with locally stored SVG files (`libs/icons/`) loaded via `bytedraft-icons.js` in the correct TinyMCE 7 format (`width`/`height` only, no `xmlns`)

### v1.3.4
- **Fix**: `safeParseJSON()` helper — all 30+ raw `JSON.parse(localStorage.getItem(...))` calls replaced; corrupted localStorage data no longer crashes the app on load, falling back to empty state instead
- **Fix**: TinyMCE content sync before re-render — `renderSections()` now calls `updateAllSectionContents()` before `tinymce.remove()` so in-flight edits are never lost when a drag-drop or theme toggle triggers a re-render
- **Fix**: Drag listener stacking — `dragover` / `drop` handlers are removed before being added in `renderSections()` so N re-renders no longer produce N stacked handlers on the sections container
- **Fix**: `showConfirm()` callback stacking — the OK button is replaced with a `cloneNode(true)` copy on each call, atomically wiping prior listeners; a single `{ once: true }` handler is attached to the fresh clone
- **Fix**: Null guards — `dragHandle`, `subContainer`, and the `nameInput` / `badge` pair in `renderCustomFields` are each checked before use, preventing runtime crashes when a querySelector returns null
- **Fix**: `FileReader.onerror` handler added in `handleJSONImport()` — a failed file read now surfaces a toast instead of silently doing nothing
- **Fix**: `safeSetItem()` now catches all storage errors — non-quota exceptions (e.g. private browsing restrictions) show a generic "Storage error" toast instead of failing silently
- **Fix**: `updateProjectStatus()` validates the status value against `['draft', 'working', 'publish']` before mutating state; replaced the inline toast div with `showToast()`

### v1.3.3
- **Fix**: Proactive localStorage warning — `safeSetItem()` now measures current storage usage before each write; a yellow toast appears once per session when usage reaches 80% of the 5 MB quota (~4 MB), prompting the user to remove logos or export old projects before hitting the hard limit
- **Fix**: Bootstrap modal instance leak — every `new bootstrap.Modal(el)` replaced with `bootstrap.Modal.getOrCreateInstance(el)`, and a single delegated `hidden.bs.modal` listener calls `.dispose()` after every modal closes; prevents event-listener accumulation over long sessions

### v1.3.2
- **Fix**: Footer page number now always appears in the right cell using Word's native `PAGE` simple field (`<w:fldSimple w:instr=" PAGE ">`); no longer conditional on `{{page}}` being present in the footer template, and no longer relies on the complex field sequence that required Word to auto-update to display a value
- **Fix**: Removed previous `PageNumber.CURRENT` TextRun approach that produced an empty result when Word did not auto-refresh the field on open
- **Feature**: Keyboard shortcuts — Ctrl+S (Cmd+S on Mac) saves the project from anywhere in the UI; Ctrl+F (Cmd+F) opens the ByteDraft Find & Replace modal; both shortcuts work whether focus is in a TinyMCE editor or elsewhere on the page (TinyMCE's built-in Ctrl+F is overridden to keep the workflow consistent)
- **Feature**: Section duplication — each section toolbar now has a copy icon (<i class="fas fa-copy"></i>) that inserts a deep clone of the section (including all subsections and their content) immediately below the original; the clone is titled "… (Copy)", unlocked, and has fresh unique IDs throughout
- **Cleanup**: Removed 6 remaining debug `console.log` calls (DOCX export start/success messages, logo upload and logo preview functions)

### v1.3.1
- **Fix**: All 16 native `alert()` / `confirm()` dialogs replaced — validation errors and status messages now surface as centered toast notifications; the delete-project confirmation uses a Bootstrap modal (`#confirmModal`) with Cancel / Delete buttons so the browser never blocks the UI thread
- **Fix**: Cross-reference spans now export correctly to DOCX — previously `span.xref[data-path]` elements were stripped to plain text; they are now re-resolved before export (current section title and number), styled blue + underlined, and passed through `extractSpanProps()` so the resulting `TextRun` carries the correct color and underline in the Word document
- **Fix**: Drag-and-drop `handleDrop()` wrapped in try-catch-finally — a failed tree mutation no longer leaves `draggedElement` / `draggedPath` set, preventing a stale-drag lock; on error a toast is shown and `renderSections()` re-syncs the DOM to the data model
- **Cleanup**: Removed 11 debug `console.log` calls from drag-and-drop functions (`handleDragEnd`, `handleDrop`, `moveNodeByPath`)

### v1.3.0
- **Improvement**: Page Settings auto-save — all fields (paper size, chars per line, lines per page, header height, paragraph spacing) are written to localStorage the moment a value changes; settings persist across page refreshes without requiring the Save button

### v1.2.0
- **Feature**: Paper size selector in Page Settings — choose US Letter, A4, Legal, or A3; selection is saved per-project and applied to all sections of the exported DOCX

### v1.1.0
- **Fix**: Blank page between Document Changelog and TOC eliminated — replaced `pageBreakBefore: true` (OOXML paragraph property) on the TOC title with an explicit `<w:br w:type="page"/>` run (`PageBreak`) at the end of the changelog content; the paragraph property was causing a double page break whenever the changelog table filled the previous page to its boundary
- **Fix**: Blank page between Title page and Changelog eliminated — the assembly was emitting a standalone `pageBreakBefore` paragraph before the changelog *and* the empty-changelog branch emitted its own leading `pageBreakBefore` paragraph, producing two consecutive forced breaks; consolidated so only the changelog title paragraph carries `pageBreakBefore: true`

### v1.0.0
- **Redesign**: Complete UI overhaul — "Slate Pro" design system replacing the generic Bootstrap appearance
- **Feature**: Collapsible sidebar — collapse the left panel with the `‹` button; a floating accent tab on the left edge restores it; state persisted to localStorage
- **Feature**: Permanent-dark sidebar — sidebar retains a deep slate palette in light mode and shifts to near-black in dark mode, matching the theme toggle
- **Feature**: Frosted-glass top bar — sticky top bar now uses `backdrop-filter` blur for a layered glass effect
- **Feature**: Pill-shaped top-bar buttons — primary action buttons use a pill border-radius with smooth hover lift/color transitions
- **Feature**: Refined modals — all modals upgraded to larger corner radius, softer shadows, and uppercase tracked form labels
- **Feature**: DM Sans typography — distinctive humanist sans-serif replaces the generic system font stack
- **Feature**: Luminous status badges — Draft / Working / Publish now use semi-transparent pill badges with glow-coloured text
- **Feature**: Section entrance animation — section cards fade and slide up on render without creating a CSS stacking context (fixing TinyMCE fullscreen)
- **Fix**: TinyMCE fullscreen — `animation-fill-mode: both` was leaving a permanent `transform: translateY(0)` on `.section-item`, creating a stacking context that trapped the `position: fixed` fullscreen overlay; changed to `backwards` fill-mode with `transform: none` final state
- **Fix**: TinyMCE fullscreen double-editor — blanket `z-index: 9900` was applied to all editors when any one went fullscreen; now only the `.tox-tinymce.tox-fullscreen` container is promoted and other editors are hidden while fullscreen is active
- **Infra**: `index.html` is now the modernised UI; previous UI preserved as `index.bak`

### v0.0.54
- **Feature**: `{{Field Name}}` placeholder substitution — type `{{Field Name}}` anywhere in a section and the value is resolved at DOCX export time; unmatched placeholders are left as-is
- **Feature**: `{{}}` toolbar button in TinyMCE — opens a dropdown of all defined custom fields and inserts the selected placeholder at the cursor; shows "No fields defined" when the list is empty
- **Improvement**: Custom fields sidebar redesigned — each field now shows an editable name input with a live-updating `{{keyword}}` badge and a separate value input, so the placeholder syntax is always visible and field names can be renamed in place without deleting and recreating

### v0.0.53
- **Fix**: Applying text color or highlight inside a table cell no longer colors the entire cell — child-element background-color scanning in `processTable` was picking up text-highlight spans and applying them as cell shading; cell background is now read only from the `<td>` element's own style

### v0.0.52
- **Fix**: Font color and background highlight now correctly exported — a second style-cleaning pass in `cleanContentForDocx` was explicitly stripping `color`, `background-color`, and `font-family` from all non-table elements after the first pass had preserved them; second pass now only removes layout-only properties (margin, padding, display, float, position)

### v0.0.51
- **Fix**: Font color and font-family now correctly exported to DOCX — `cleanContentForDocx` was stripping `color` and `font-family` properties when rebuilding span styles, so they never reached the export pipeline

### v0.0.50
- **Fix**: Font color now exported correctly to DOCX — `extractSpanProps` was calling `this.parseColor()` which doesn't exist (correct method is `convertColorToHex`), causing a silent TypeError that dropped all formatting from colored paragraphs
- **Fix**: Text background color (highlighting) now exported to DOCX via TextRun shading

### v0.0.49
- **Fix**: Paragraph text alignment (left/center/right/justify) outside tables now correctly exported to DOCX — alignment was being set post-construction on the Paragraph object which the docx library ignores

### v0.0.48
- **Fix**: DOCX table export — tables were silently dropped due to `TableLayoutType` not existing in the loaded docx library version
- **Fix**: Images inside table cells now export correctly to DOCX
- **Fix**: Text formatting (bold, italic, color, font size, font family) inside table cells is now preserved in DOCX export — properties are passed via style inheritance so all run properties are set at `TextRun` construction time
- **Fix**: Bullet and numbered lists inside table cells now export with correct structure and indentation
- **Fix**: Per-paragraph text alignment (left/center/right/justify) inside table cells is now respected in DOCX export

### v0.0.47
- **Fix**: Auto-save no longer logs to revision history — only explicit manual saves create revision entries
- **Fix**: Project card action buttons (Export, Status, Delete) now render below the project title instead of beside it

### v0.0.46
- **Feature**: Section annotations — each section and subsection has a comment button (speech bubble icon); click to open a modal where you can add, resolve, unresolve, and delete review notes
- **Behaviour**: Comment button turns blue with an unresolved-count badge when comments are pending; grey when all resolved or none exist
- **Persistence**: Comments stored directly on section nodes (`node.comments[]`) — included automatically in JSON export/import, invisible to DOCX export
- **Dark mode**: Comment cards properly themed in dark mode

### v0.0.45
- **Feature**: Cross-references — `[XRef]` TinyMCE toolbar button opens a section picker and inserts a styled, non-editable `<span class="xref">` (blue, underlined, italic) at the cursor; resolves to "Section N — Title" plain text in DOCX export
- **Feature**: Collapsible sidebar sections — Projects, TOC, Templates, and Custom Fields panels can be collapsed/expanded with a chevron; state persisted to localStorage
- **Feature**: Find & Replace — modal with find/replace inputs, case-sensitive option, Find All (lists matches by section), Replace All (updates live editors and section titles, saves automatically)
- **Feature**: Word count / reading time — per-section word count badge updates live as you type; document total and estimated reading time (200 wpm) displayed below the TOC in the sidebar
- **Feature**: Section locking — lock any section read-only with a toolbar lock button; locked sections disable the title input, delete/add-subsection buttons, and drag handle; TinyMCE content is non-editable; fullscreen is blocked and visually greyed out; lock state persists with project data

### v0.0.44
- **Feature**: Citation Manager — insert numbered inline citations `[1]`, `[2]` via a new `[Cite]` TinyMCE toolbar button
- **Feature**: Citation Manager modal — add, edit, and delete references (Title, Authors, Year, Source, URL, Notes); one-click insertion of `<sup>[N]</sup>` at the cursor
- **Feature**: References page in DOCX export — a formatted "References" page is automatically appended when citations exist; omitted when the project has none
- **Storage**: New localStorage key `bytedraft_references` — per-project reference lists stored as JSON arrays

### v0.0.43
- **Refactor**: Extracted all inline JavaScript into `app.js` for cleaner project structure
- **Security**: Added `escapeHtml()` — all user data interpolated into `innerHTML` is now escaped (XSS prevention)
- **Fix**: Theme toggle no longer reloads the page; TinyMCE editors are cleanly recycled with the new skin
- **Fix**: Changelog JSON double-encoding bug — data is now stored as plain arrays, with a migration for existing saves
- **Fix**: Template preview now uses a Bootstrap modal instead of `alert()`
- **Fix**: TOC preview now shows estimated page numbers based on configurable chars/line and lines/page settings
- **Storage**: All `localStorage.setItem` calls wrapped in `safeSetItem()` — quota errors surface as a user-visible toast instead of silently failing
- **Removed**: Dead code — legacy JSZip/XML DOCX pipeline, dead PDF helpers, debug test functions, flat-index section management
- **Removed**: Unused script tags — `jszip.min.js`, `highlight.min.js`, `html2canvas.min.js` (~200 KB savings)
- **Cleanup**: Stripped debug `console.log` spam from editor init, theme toggle, and DOCX export
- **Refactor**: TinyMCE configuration centralised in `getTinyMCEBaseConfig()` / `buildTinyMCEContentStyle()`

### v0.0.42 and earlier
See git history.

---

## 🚀 Getting Started

### Quick Start
1. **Download or clone this repository**
2. **Open `index.html` in your web browser**
   - For best results, use a local web server (e.g., `python3 -m http.server 8000`)
   - File protocol (`file://`) works for most features
3. **Create a new project** or select an existing template
4. **Start editing** with the full-featured TinyMCE editor
5. **Export to DOCX** when ready

### Creating Your First Document
1. Click **"New Project"** in the sidebar
2. Enter a project name and description
3. Choose a template (optional) or start with a blank document
4. Add sections using the **"Add Section"** button
5. Edit content using the rich text editor
6. Add images by dragging them into the editor
7. Set document information using "Edit Document Info"
8. Add changelog entries using "Document Change Log"
9. Configure headers/footers using "Edit Header/Footer"
10. Export to DOCX when finished

## 🎯 Drag & Drop Usage

### **Reordering Sections:**
- **Drag any section** using the grip handle (⋮⋮)
- **Drop on another section** to reorder at the same level
- **Drop in empty space** to move to the end

### **Moving Subsections:**
- **🔵 Blue zone** (lower half of section): Makes it a child of that section
- **🔴 Red zone** (upper half of section): Promotes it to a top-level section
- **Drop between subsections**: Reorders within the same parent

### **Visual Feedback:**
- **Thick dashed borders**: 4px borders for clear visibility
- **Color-coded zones**: Blue for child, red for promotion
- **Hover effects**: Visual feedback during drag operations
- **Success messages**: Confirmation when reordering is complete

## 🛠️ Advanced Features

### Document Structure & Numbering
- **Automatic Section Numbering**: Sections are automatically numbered (1., 2., 3.)
- **Subsection Hierarchy**: Subsections get decimal numbering (1.1., 1.2., 1.1.1.)
- **Word Heading Styles**: Proper Heading 1, 2, 3 styles for TOC integration
- **Unlimited Depth**: Support for unlimited nested sections

### Table of Contents
- **Auto-updating TOC**: Word-compatible table of contents field
- **Professional Formatting**: Proper indentation and formatting
- **Update Instructions**: Clear guidance for users to update TOC in Word
- **Heading Integration**: Automatically picks up all numbered headings

### Advanced Table Export
- **Cell Color Preservation**: Background colors maintained in DOCX export
- **Border Styling**: Table borders, styles, and widths preserved
- **Cell Alignment**: Text alignment within cells maintained
- **Complex Formatting**: Nested content, images, and formatting preserved
- **Size Control**: Table dimensions and positioning maintained

### Drag & Drop Section Management
- **Intuitive Reordering**: Drag sections and subsections to reorder them easily
- **Color-Coded Drop Zones**: 
  - 🔵 **Blue zones**: Add as child/subsection
  - 🔴 **Red zones**: Promote to top-level section
  - **Thick borders**: 4px dashed borders for clear visibility
- **Visual Feedback**: Clear visual indicators during drag operations with hover effects
- **Hierarchical Movement**: Move subsections between different parent sections
- **Automatic Numbering**: Section numbers update automatically after reordering
- **Content Preservation**: All content, formatting, and nested structure is maintained
- **Cross-Section Movement**: Subsections can be moved to different parent sections
- **Drag Handles**: Dedicated grip handles (⋮⋮) for precise control over reordering
- **TinyMCE Integration**: Seamless editor handling during reordering operations

### Image Management
- **Supported Formats**: PNG, JPEG, GIF, WebP
- **Storage**: Images are converted to base64 and stored locally
- **Export**: All images are properly embedded in DOCX exports
- **Size Control**: Images maintain their aspect ratio
- **Multiple Insertion Methods**: Drag & drop, file picker, or paste

### Document Information
- **Custom Metadata**: Document title, author, owners, version, dates
- **Professional Layout**: Information displayed in formatted table
- **Export Integration**: All metadata included in DOCX exports
- **Flexible Fields**: Add custom fields as needed
- **Data Persistence**: All document info preserved in JSON exports and imports

### Changelog Management
- **Professional Changelog**: Version tracking with approval workflow
- **Custom Fields**: Version number, dates, author, reviewer, approver, description
- **Export Integration**: Changelog appears as dedicated page in exports
- **Data Persistence**: All changelog data stored locally and preserved in JSON exports
- **UI Integration**: Changelog data automatically loads on import without page refresh

### Document Templates
- **Built-in Templates**: Technical documentation, user guides, API docs
- **Custom Templates**: Edit `templates.js` to add your own templates
- **Template Structure**: Define sections, subsections, and default content
- **Professional Categories**: Security, compliance, technical, and business templates

### Version Control
- **Manual Saves Only**: Revision history entries are created only on explicit manual saves
- **Silent Auto-save**: Auto-save persists data every 30 seconds without polluting the revision log
- **Status Changes**: Document status updates are logged
- **Export History**: Track when documents were exported
- **Data Persistence**: Version history preserved in JSON exports and imports

### Document Logo Management
- **Logo Upload**: Upload custom logos (PNG, JPG, GIF) for DOCX title pages
- **Preview System**: Real-time logo preview in page settings
- **Automatic Sizing**: Logos are automatically sized and positioned on title pages
- **Data Persistence**: Logo data preserved in JSON exports and imports
- **Easy Management**: Remove logos with one click
- **Professional Integration**: Logos appear at the top of title pages in DOCX exports

### Page Settings & Layout
- **Configurable Parameters**: Adjust characters per line, lines per page, header height, and paragraph spacing
- **Accurate Estimation**: Better page number calculations for TOC and document planning
- **Professional Layout**: Optimized settings for different document types
- **User Control**: Fine-tune page layout parameters for specific requirements
- **Logo Management**: Upload and manage document logos directly from page settings
- **Integrated Interface**: All page-related settings in one convenient modal

### Enhanced JSON Import/Export
- **Complete Data Preservation**: All project data including document info, changelog, headers/footers, and version history
- **Automatic UI Refresh**: Imported data immediately available without page refresh
- **Conflict Resolution**: Automatic handling of duplicate project names during import
- **Data Validation**: Robust validation of imported JSON structure
- **Seamless Integration**: Imported projects work exactly like native projects

### Custom Fields & Placeholders
- **Project Metadata**: Add custom fields for project-specific information (e.g. Client Name, Review Date)
- **`{{Field}}` Placeholder Syntax**: Type `{{Field Name}}` anywhere in a section — values are substituted at DOCX export time
- **Insert from Toolbar**: Use the `{{}}` TinyMCE toolbar button to pick a field from a dropdown and insert the placeholder at the cursor
- **Editable Name & Value**: Each field in the sidebar shows an editable name input with a live `{{keyword}}` badge and a separate value input — rename fields without deleting and recreating
- **Flexible Structure**: Define field names and values as needed
- **Multiple Types**: Text, date, email, URL field types
- **Data Persistence**: Custom fields preserved in JSON exports and imports

### Theme System
- **Dark/Light Mode**: Toggle between themes with local storage persistence
- **No-Reload Switching**: Theme changes apply instantly — TinyMCE editors are cleanly destroyed and recreated with the correct skin without a page reload
- **Comprehensive Styling**: All UI components properly themed with CSS custom properties
- **Persistent Settings**: Theme preference saved and restored automatically
- **Modular CSS**: Well-organized styles with comprehensive documentation and logical sections

## 📋 Requirements

### Browser Compatibility
- **Chrome** 80+ (recommended)
- **Firefox** 75+
- **Edge** 80+
- **Safari** 13+

### System Requirements
- **Storage**: At least 50MB free space for image storage
- **Memory**: 2GB RAM recommended for large documents
- **Network**: No internet connection required (fully offline)

## 🔧 Technical Details

### Libraries Used
- **TinyMCE 6**: Professional rich text editor with GPL license
- **Bootstrap 5**: Modern UI framework
- **Font Awesome**: Icon library
- **docx**: Modern DOCX generation library
- **highlight.js**: Code syntax highlighting with GitHub Dark theme (CSS theme only; runtime highlighting not used)

### File Structure
```
ByteDraft/
├── index.html              # Main application shell (HTML + embedded "Slate Pro" styles + modals)
├── index.bak               # Previous UI (archived)
├── app.js                  # Application logic
├── templates.js            # Document templates
├── modernDocxExport.js     # Modern DOCX export module
├── libs/                   # External libraries
│   ├── css/               # CSS files
│   │   ├── styles.css     # Legacy styles (no longer loaded by main UI)
│   │   └── github-dark.min.css # Code syntax highlighting theme
│   ├── tinymce/           # TinyMCE editor
│   ├── bootstrap/         # Bootstrap JS + CSS
│   ├── docx/              # DOCX generation library
│   └── fonts/             # Font Awesome icons
└── README.md              # This file
```

### CSS Organization
- **Modular Structure**: Well-organized CSS with comprehensive comments
- **Theme Variables**: CSS custom properties for easy theming
- **Component-Based**: Logical sections for different UI components
- **Dark Mode Support**: Extensive dark theme overrides
- **Responsive Design**: Mobile-friendly adaptations

## 🎯 Use Cases

### Technical Documentation
- API documentation
- User manuals
- System specifications
- Process documentation
- Technical reports
- Security documentation
- Penetration test reports

### Business Documents
- Project proposals
- Business plans
- Standard operating procedures
- Policy documents
- Training materials
- Compliance reports
- Risk assessments

### Academic Writing
- Research papers
- Thesis documents
- Course materials
- Academic reports
- Literature reviews

## 🤝 Contributing

ByteDraft is designed to be easily customizable and extensible:

### Customizing Templates
Edit `templates.js` to add new document templates:
```javascript
templates['my-template'] = {
    name: 'My Template',
    description: 'Description of my template',
    sections: [
        { id: '1', title: 'Introduction', content: 'Default content...' }
    ]
};
```

### Customizing Styles
The UI styles are embedded in the `<style>` block inside `index.html`. The design system uses CSS custom properties (variables) defined at the top of that block — edit the `:root` and `[data-theme="dark"]` sections to restyle the entire application:
- `--accent` / `--accent-hi` — primary action colour
- `--sb-bg` / `--sb-text` — sidebar colours
- `--bd-bg` / `--bd-surface` — main content background and card surfaces
- `--r-*` — border-radius tokens (xs → pill)
- `--t-fast` / `--t-mid` — animation duration tokens

### Extending Functionality
The modular design allows easy addition of:
- New export formats
- Additional plugins
- Custom field types
- Enhanced templates


## 🙏 Credits

- **TinyMCE**: Professional rich text editing - GPL v2 or commercial license
- **Bootstrap**: Modern UI components - MIT License
- **Font Awesome**: Beautiful icons - MIT License
- **docx**: Modern DOCX generation - MIT License
- **JSZip**: File compression utilities - MIT License
- **highlight.js**: Code syntax highlighting - BSD License

---

**ByteDraft** - Professional documentation made simple
**Version** 1.4.0
© 2026 - Built for offline productivity