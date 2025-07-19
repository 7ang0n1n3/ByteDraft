# ByteDraft

ByteDraft is a modern technical documentation tool for creating, editing, and exporting structured reports (DOCX) with advanced formatting, unlimited section depth, and a user-friendly interface. Built for offline use with professional-grade editing capabilities.

## ✨ Features

### 📝 **Advanced Rich Text Editing**
- **Full TinyMCE Integration** with complete menu bar and toolbar
- **Professional Formatting**: Bold, italic, underline, strikethrough, text/background colors
- **Typography Options**: Multiple font families and sizes (8pt to 48pt)
- **Advanced Tools**: Search/replace, fullscreen mode, word count, character map
- **Media Support**: Drag-and-drop image insertion with base64 storage
- **Table Support**: Create and format tables with ease
- **Code Blocks**: Syntax highlighting for code snippets with GitHub Dark theme

### 📚 **Document Structure**
- **Unlimited Nested Sections**: Create complex document hierarchies with automatic numbering
- **Professional Heading Hierarchy**: Automatic Heading 1, 2, 3 styles with proper numbering (1., 1.1., 1.1.1.)
- **Auto-updating Table of Contents**: Word-compatible TOC field that updates automatically
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
  - Title page with document information
  - Document changelog page with custom fields
  - Auto-updating Table of Contents
  - Properly numbered sections and subsections
- **Image Preservation**: All images included in exported documents
- **Formatting Retention**: Maintains all text formatting and structure
- **Professional Layout**: Proper headings, lists, tables, and styling
- **Document Metadata**: Includes title page, version history, and changelog

### 🎨 **User Interface**
- **Modern Design**: Clean, professional interface with Bootstrap 5
- **Responsive Layout**: Works on desktop and tablet devices
- **Project Management**: Organize multiple documentation projects
- **Status Tracking**: Draft, Working, and Publish statuses
- **Real-time Preview**: See changes reflected immediately
- **Dark/Light Theme**: Toggle between themes for comfortable editing

### 💾 **Data Management**
- **Local Storage**: All data stored in your browser (no server required)
- **Auto-save**: Automatic saving every 30 seconds
- **Export/Import**: JSON export for backup and sharing
- **Offline Operation**: Works completely without internet connection
- **Document Information**: Custom document metadata fields
- **Changelog Management**: Professional changelog with version tracking

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
9. Export to DOCX when finished

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

### Changelog Management
- **Professional Changelog**: Version tracking with approval workflow
- **Custom Fields**: Version number, dates, author, reviewer, approver, description
- **Export Integration**: Changelog appears as dedicated page in exports
- **Data Persistence**: All changelog data stored locally

### Document Templates
- **Built-in Templates**: Technical documentation, user guides, API docs
- **Custom Templates**: Edit `templates.js` to add your own templates
- **Template Structure**: Define sections, subsections, and default content
- **Professional Categories**: Security, compliance, technical, and business templates

### Version Control
- **Automatic Tracking**: Every save creates a version entry
- **Manual Saves**: Explicit saves are tracked separately
- **Status Changes**: Document status updates are logged
- **Export History**: Track when documents were exported

### Custom Fields
- **Project Metadata**: Add custom fields for project-specific information
- **Flexible Structure**: Define field names and values as needed
- **Export Inclusion**: Custom fields appear in exported documents
- **Multiple Types**: Text, date, email, URL field types

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
- **JSZip**: File compression for exports
- **highlight.js**: Code syntax highlighting with GitHub Dark theme

### File Structure
```
ByteDraft/
├── index.html              # Main application
├── templates.js            # Document templates
├── modernDocxExport.js     # Modern DOCX export module
├── libs/                   # External libraries
│   ├── tinymce/           # TinyMCE editor
│   ├── bootstrap/         # Bootstrap CSS/JS
│   ├── docx/              # DOCX generation library
│   ├── fonts/             # Font Awesome icons
│   └── github-dark.min.css # Code syntax highlighting theme
└── README.md              # This file
```

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

### Adding Templates
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
Modify the CSS in `index.html` to change the appearance:
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

## 📄 License

ByteDraft uses the following open-source licenses:
- **TinyMCE**: GPL v2 or commercial license
- **Bootstrap**: MIT License
- **Font Awesome**: MIT License
- **docx**: MIT License
- **JSZip**: MIT License
- **highlight.js**: BSD License

## 🙏 Credits

- **TinyMCE**: Professional rich text editing
- **Bootstrap**: Modern UI components
- **Font Awesome**: Beautiful icons
- **docx**: Modern DOCX generation
- **JSZip**: File compression utilities
- **highlight.js**: Code syntax highlighting

---

**ByteDraft** - Professional documentation made simple  
© 2025 - Built for offline productivity 