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
- **Modern Design**: Clean, professional interface with Bootstrap 5
- **Responsive Layout**: Works on desktop and tablet devices
- **Project Management**: Organize multiple documentation projects
- **Status Tracking**: Draft, Working, and Publish statuses
- **Real-time Preview**: See changes reflected immediately
- **Dark/Light Theme**: Toggle between themes for comfortable editing with persistence
- **Centralized Notifications**: All status messages appear in the top center for consistent user experience
- **Professional CSS Organization**: Well-structured styles with comprehensive documentation

### 💾 **Data Management**
- **Local Storage**: All data stored in your browser (no server required)
- **Auto-save**: Automatic saving every 30 seconds — silently persists data without creating a revision history entry
- **Export/Import**: JSON export for backup and sharing with complete project data preservation
- **Offline Operation**: Works completely without internet connection
- **Document Information**: Custom document metadata fields
- **Changelog Management**: Professional changelog with version tracking
- **Complete Data Preservation**: JSON exports include document info, changelog, headers/footers, and version history

## 📋 Changelog

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

### Custom Fields
- **Project Metadata**: Add custom fields for project-specific information
- **Flexible Structure**: Define field names and values as needed
- **Export Inclusion**: Custom fields appear in exported documents
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
├── index.html              # Main application shell (HTML + modals)
├── app.js                  # Application logic
├── templates.js            # Document templates
├── modernDocxExport.js     # Modern DOCX export module
├── libs/                   # External libraries
│   ├── css/               # CSS files
│   │   ├── styles.css     # Main application styles
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
Modify the CSS in `libs/css/styles.css` to change the appearance:
- Color schemes
- Layout adjustments
- Typography changes
- Component styling

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
**Version** 0.0.49
© 2025 - Built for offline productivity