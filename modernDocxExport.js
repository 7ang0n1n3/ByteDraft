// Modern DOCX Export Module using docx library
// This replaces the custom JSZip-based implementation with better image support

class ModernDocxExporter {
    constructor() {
        this.docx = window.docx;
        if (!this.docx) {
            console.error('DOCX library not loaded');
            return;
        }
    }

    async exportProjectToDocx(project, customChangelog = '', versionHistory = []) {
        if (!this.docx) {
            throw new Error('DOCX library not available');
        }

        try {
            // Collect all content from sections and subsections
            const allContent = this.collectAllContent(project);
            
            // Clean up the content for better DOCX conversion
            const cleanedContent = this.cleanContentForDocx(allContent);
            
            // Convert HTML to DOCX elements
            const docxElements = this.htmlToDocxElements(cleanedContent);
            
            // Add title page (includes document info)
            const titlePage = this.createTitlePage(project);
            
            // Add changelog page as second page
            const changelogPage = this.createChangelogPage(project);
            
            // Add TOC page as third page
            const tocPage = this.createTOCPage(project);
            
            // Get custom header and footer from localStorage
            const headerFooter = JSON.parse(localStorage.getItem('bytedraft_header_footer') || '{}');
            const projectHeaderFooter = headerFooter[project.id] || { header: '', footer: '' };
            

            
            // Process header text to replace {{title}} with project name
            let headerText = projectHeaderFooter.header || project.name;
            headerText = headerText.replace(/\{\{title\}\}/g, project.name);
            

            
            // Process footer text to replace {{page}} and {{title}} variables
            let footerText = projectHeaderFooter.footer || 'ByteDraft Document | Generated on ' + new Date().toLocaleDateString();
            
            // Replace {{title}} with project name
            footerText = footerText.replace(/\{\{title\}\}/g, project.name);
            

            
            // Create headers and footers using custom data
            const header = new this.docx.Header({
                children: [
                    new this.docx.Paragraph({
                        children: [
                            new this.docx.TextRun({
                                text: headerText,
                                bold: true,
                                size: 20,
                                color: '2563EB'
                            })
                        ],
                        alignment: this.docx.AlignmentType.LEFT
                    })
                ]
            });

            // Create footer with left-aligned content and right-aligned page number
            const footer = new this.docx.Footer({
                children: [
                    new this.docx.Table({
                        rows: [
                            new this.docx.TableRow({
                                children: [
                                    // Left-aligned content (without page number)
                                    new this.docx.TableCell({
                                        children: [
                                            new this.docx.Paragraph({
                                                children: [
                                                    new this.docx.TextRun({
                                                        text: footerText.replace(/\{\{page\}\}/g, '').trim(),
                                                        size: 16,
                                                        color: '666666'
                                                    })
                                                ],
                                                alignment: this.docx.AlignmentType.LEFT
                                            })
                                        ],
                                        width: { size: 70, type: this.docx.WidthType.PERCENTAGE },
                                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                                    }),
                                    // Right-aligned page number
                                    new this.docx.TableCell({
                                        children: [
                                            new this.docx.Paragraph({
                                                children: [
                                                    new this.docx.TextRun({
                                                        text: footerText.includes('{{page}}') ? '4' : '',
                                                        size: 16,
                                                        color: '666666'
                                                    })
                                                ],
                                                alignment: this.docx.AlignmentType.RIGHT
                                            })
                                        ],
                                        width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                        margins: { top: 0, bottom: 0, left: 0, right: 0 }
                                    })
                                ]
                            })
                        ],
                        width: { size: 100, type: this.docx.WidthType.PERCENTAGE },
                        borders: {
                            top: { style: this.docx.BorderStyle.NONE },
                            bottom: { style: this.docx.BorderStyle.NONE },
                            left: { style: this.docx.BorderStyle.NONE },
                            right: { style: this.docx.BorderStyle.NONE },
                            insideHorizontal: { style: this.docx.BorderStyle.NONE },
                            insideVertical: { style: this.docx.BorderStyle.NONE }
                        }
                    })
                ]
            });

            // Create the document with multiple sections
            const doc = new this.docx.Document({
                sections: [
                    // Section 1: Title page, changelog, TOC (no headers/footers)
                    {
                        properties: {
                            page: {
                                margin: {
                                    top: 1440,    // 1 inch
                                    right: 1440,  // 1 inch
                                    bottom: 1440, // 1 inch
                                    left: 1440    // 1 inch
                                }
                            }
                        },
                        children: [
                            // Title page content
                            ...titlePage,
                            
                            // Page break after title page
                            new this.docx.Paragraph({
                                pageBreakBefore: true
                            }),
                            
                            // Changelog page content
                            ...changelogPage,
                            
                            
                            // TOC page content
                            ...tocPage
                        ]
                    },
                    // Section 2: Main content (with headers/footers)
                    {
                        properties: {
                            page: {
                                margin: {
                                    top: 1440,    // 1 inch
                                    right: 1440,  // 1 inch
                                    bottom: 1440, // 1 inch
                                    left: 1440    // 1 inch
                                },
                                pageNumbers: {
                                    start: 4  // Start page numbering at 4
                                }
                            }
                        },
                        headers: {
                            default: header
                        },
                        footers: {
                            default: footer
                        },
                        children: [
                            
                            // Main content
                            ...docxElements
                        ]
                    }
                ]
            });
            // Generate and download the file
            const blob = await this.docx.Packer.toBlob(doc);
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${project.name.replace(/[^a-z0-9]/gi, '_')}.docx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            
        } catch (error) {
            console.error('Error exporting to DOCX:', error);
            console.error('Error stack:', error.stack);
            throw error;
        }
    }

    collectAllContent(project) {
        let allContent = '';
        
        if (project.sections && Array.isArray(project.sections)) {
            project.sections.forEach((section, index) => {
                allContent += this.collectSectionContent(section, 1, [index + 1]);
            });
        }
        
        return allContent;
    }

    collectSectionContent(section, level, numberParts = []) {
        let content = '';
        
        // Generate section number
        const sectionNumber = numberParts.join('.');
        const numberedTitle = sectionNumber ? `${sectionNumber}. ${section.title}` : section.title;
        
        // Add section heading with proper level
        const headingTag = `h${Math.min(level, 6)}`;
        content += `<${headingTag}>${numberedTitle}</${headingTag}>`;
        
        // Add section content
        if (section.content) {
            content += section.content;
        }
        
        // Add subsections recursively with proper numbering
        if (section.subsections && Array.isArray(section.subsections)) {
            section.subsections.forEach((subsection, index) => {
                const newNumberParts = [...numberParts, index + 1];
                content += this.collectSectionContent(subsection, level + 1, newNumberParts);
            });
        }
        
        return content;
    }

    cleanContentForDocx(html) {
        // Create a temporary div to parse and clean the HTML
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;
        

        
        // Remove inline styles that might interfere with DOCX conversion
        const elementsWithStyles = tempDiv.querySelectorAll('[style]');
        elementsWithStyles.forEach(el => {
            el.removeAttribute('style');
        });
        
        // Convert divs with only text content to paragraphs
        const divs = tempDiv.querySelectorAll('div');
        divs.forEach(div => {
            if (div.children.length === 0 && div.textContent.trim()) {
                const p = document.createElement('p');
                p.textContent = div.textContent;
                div.parentNode.replaceChild(p, div);
            }
        });
        
        // Remove any changelog JSON content that might be embedded in the HTML
        const textContent = tempDiv.textContent || tempDiv.innerText || '';
        if (textContent.includes('"version"') && textContent.includes('"author"') && textContent.includes('"reviewer"')) {
            // This looks like changelog JSON, remove it
            const paragraphs = tempDiv.querySelectorAll('p');
            paragraphs.forEach(p => {
                const text = p.textContent || p.innerText || '';
                if (text.includes('"version"') && text.includes('"author"') && text.includes('"reviewer"')) {
                    p.remove();
                }
            });
        }
        
        return tempDiv.innerHTML;
    }

    htmlToDocxElements(html) {
        try {
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = html;
            
            const elements = [];
            
            // Process each child node recursively
            this.processNode(tempDiv, elements);
            
            return elements;
        } catch (error) {
            console.error('Error in htmlToDocxElements:', error);
            // Return a simple error paragraph if conversion fails
            return [
                new this.docx.Paragraph({
                    children: [new this.docx.TextRun({ text: 'Error converting content to DOCX format.' })]
                })
            ];
        }
    }

    processNode(node, elements) {
        try {
            if (node.nodeType === Node.ELEMENT_NODE) {
                const tagName = node.tagName.toLowerCase();
                

                
                switch (tagName) {
                case 'h1':
                    elements.push(new this.docx.Paragraph({
                        text: node.textContent,
                        heading: this.docx.HeadingLevel.HEADING_1,
                        spacing: { before: 400, after: 200 }
                    }));
                    break;
                case 'h2':
                    elements.push(new this.docx.Paragraph({
                        text: node.textContent,
                        heading: this.docx.HeadingLevel.HEADING_2,
                        spacing: { before: 300, after: 150 }
                    }));
                    break;
                case 'h3':
                    elements.push(new this.docx.Paragraph({
                        text: node.textContent,
                        heading: this.docx.HeadingLevel.HEADING_3,
                        spacing: { before: 200, after: 100 }
                    }));
                    break;
                case 'h4':
                    elements.push(new this.docx.Paragraph({
                        text: node.textContent,
                        heading: this.docx.HeadingLevel.HEADING_4,
                        spacing: { before: 150, after: 100 }
                    }));
                    break;
                case 'h5':
                    elements.push(new this.docx.Paragraph({
                        text: node.textContent,
                        heading: this.docx.HeadingLevel.HEADING_5,
                        spacing: { before: 100, after: 100 }
                    }));
                    break;
                case 'h6':
                    elements.push(new this.docx.Paragraph({
                        text: node.textContent,
                        heading: this.docx.HeadingLevel.HEADING_6,
                        spacing: { before: 100, after: 100 }
                    }));
                    break;
                case 'p':
                    // Check if paragraph contains images
                    const imgElements = node.querySelectorAll('img');
                    if (imgElements.length > 0) {

                        
                        // Process each image in the paragraph
                        imgElements.forEach(img => {
                            this.processImage(img, elements);
                        });
                        
                        // Also process any text content in the paragraph
                        const textContent = node.textContent.trim();
                        if (textContent) {
                            const children = this.processInlineElements(node);
                            if (children.length > 0) {
                                const paragraph = new this.docx.Paragraph({ 
                                    children: children,
                                    spacing: { after: 200 }
                                });
                                
                                // Handle text alignment
                                const textAlign = this.getTextAlignment(node);
                                if (textAlign) {
                                    paragraph.alignment = textAlign;
                                }
                                
                                elements.push(paragraph);
                            }
                        }
                    } else if (node.textContent.trim()) {
                        // Handle paragraphs with only text content and formatting
                        const children = this.processInlineElements(node);
                        if (children.length > 0) {
                            const paragraph = new this.docx.Paragraph({ 
                                children: children,
                                spacing: { after: 200 }
                            });
                            
                            // Handle text alignment
                            const textAlign = this.getTextAlignment(node);
                            if (textAlign) {
                                paragraph.alignment = textAlign;
                            }
                            
                            elements.push(paragraph);
                        }
                    }
                    break;
                case 'ul':
                case 'ol':
                    this.processList(node, elements);
                    break;
                case 'table':
                    this.processTable(node, elements);
                    break;
                case 'img':
    
                    this.processImage(node, elements);
                    break;
                case 'blockquote':
                    this.processBlockquote(node, elements);
                    break;
                default:
                    // Process child nodes recursively
                    Array.from(node.childNodes).forEach(child => {
                        this.processNode(child, elements);
                    });
                    break;
            }
        } else if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
            elements.push(new this.docx.Paragraph({
                children: [new this.docx.TextRun({ text: node.textContent.trim() })],
                spacing: { after: 200 }
            }));
        }
        } catch (error) {
            console.error('Error processing node:', error);
            // Add a simple text element as fallback
            if (node.textContent && node.textContent.trim()) {
                elements.push(new this.docx.Paragraph({
                    children: [new this.docx.TextRun({ text: node.textContent.trim() })],
                    spacing: { after: 200 }
                }));
            }
        }
    }

    processImage(imgElement, elements) {
        const src = imgElement.getAttribute('src');
        const alt = imgElement.getAttribute('alt') || 'Image';
        const width = imgElement.getAttribute('width') || 400;
        const height = imgElement.getAttribute('height') || 300;
        
        
        
        if (src) {
            try {
                // Convert base64 data URL to Uint8Array for browser environment
                let imageData;
                if (src.startsWith('data:image/')) {
                    // Handle base64 images
                    const base64Data = src.split(',')[1];
                    const binaryString = atob(base64Data);
                    const bytes = new Uint8Array(binaryString.length);
                    for (let i = 0; i < binaryString.length; i++) {
                        bytes[i] = binaryString.charCodeAt(i);
                    }
                    imageData = bytes;
                    

                } else {
                    // For external URLs, we would need to fetch them
                    // For now, we'll skip external images in DOCX export
                    elements.push(new this.docx.Paragraph({
                        children: [new this.docx.TextRun({ text: `[External Image: ${alt}]` })]
                    }));
                    return;
                }
                
                // Try different approaches for adding image to DOCX
                try {
                    // Method 1: Direct ImageRun
                    const imageRun = new this.docx.ImageRun({
                        data: imageData,
                        transformation: {
                            width: parseInt(width),
                            height: parseInt(height)
                        }
                    });
                    
                    elements.push(new this.docx.Paragraph({
                        children: [imageRun],
                        alignment: this.docx.AlignmentType.CENTER
                    }));
                    

                    
                } catch (imageError) {
                    console.error('ImageRun failed, trying alternative method:', imageError);
                    
                    // Method 2: Try with different parameters
                    try {
                        const imageRun = new this.docx.ImageRun({
                            data: imageData,
                            transformation: {
                                width: parseInt(width) * 9525, // Convert to EMUs
                                height: parseInt(height) * 9525
                            }
                        });
                        
                        elements.push(new this.docx.Paragraph({
                            children: [imageRun],
                            alignment: this.docx.AlignmentType.CENTER
                        }));
                        

                        
                    } catch (emuError) {
                        console.error('EMU conversion failed:', emuError);
                        
                        // Method 3: Try without transformation
                        try {
                            const imageRun = new this.docx.ImageRun({
                                data: imageData
                            });
                            
                            elements.push(new this.docx.Paragraph({
                                children: [imageRun],
                                alignment: this.docx.AlignmentType.CENTER
                            }));
                            

                            
                        } catch (simpleError) {
                            console.error('Simple image addition failed:', simpleError);
                            throw simpleError;
                        }
                    }
                }
                
            } catch (error) {
                console.error('Error processing image for DOCX:', error);
                // Fallback: add image description as text
                elements.push(new this.docx.Paragraph({
                    children: [new this.docx.TextRun({ text: `[Image: ${alt}]` })]
                }));
            }
        } else {
            elements.push(new this.docx.Paragraph({
                children: [new this.docx.TextRun({ text: `[Image: ${alt}]` })]
            }));
        }
    }

    processInlineElements(element) {
        const children = [];
        
        Array.from(element.childNodes).forEach(child => {
            if (child.nodeType === Node.TEXT_NODE) {
                if (child.textContent.trim()) {
                    children.push(new this.docx.TextRun({ text: child.textContent }));
                }
            } else if (child.nodeType === Node.ELEMENT_NODE) {
                const tagName = child.tagName.toLowerCase();
                
                switch (tagName) {
                    case 'br':
                        children.push(new this.docx.TextRun({ text: '\n' }));
                        break;
                    case 'strong':
                    case 'b':
                        children.push(new this.docx.TextRun({ 
                            text: child.textContent, 
                            bold: true 
                        }));
                        break;
                    case 'em':
                    case 'i':
                        children.push(new this.docx.TextRun({ 
                            text: child.textContent, 
                            italics: true 
                        }));
                        break;
                    case 'u':
                        children.push(new this.docx.TextRun({ 
                            text: child.textContent, 
                            underline: {} 
                        }));
                        break;
                    case 'a':
                        children.push(new this.docx.TextRun({ 
                            text: child.textContent, 
                            color: '0563C1',
                            underline: { type: 'single' }
                        }));
                        break;
                    case 'code':
                        children.push(new this.docx.TextRun({ 
                            text: child.textContent, 
                            font: 'Courier New',
                            size: 20
                        }));
                        break;
                    case 'mark':
                        children.push(new this.docx.TextRun({ 
                            text: child.textContent, 
                            highlight: 'yellow'
                        }));
                        break;
                    case 'sub':
                        children.push(new this.docx.TextRun({ 
                            text: child.textContent, 
                            subScript: true
                        }));
                        break;
                    case 'sup':
                        children.push(new this.docx.TextRun({ 
                            text: child.textContent, 
                            superScript: true
                        }));
                        break;
                    default:
                        // Recursively process other inline elements
                        const nestedChildren = this.processInlineElements(child);
                        children.push(...nestedChildren);
                        break;
                }
            }
        });
        
        return children;
    }

    getTextAlignment(element) {
        const style = element.style || {};
        const textAlign = style.textAlign || style['text-align'];
        
        switch (textAlign) {
            case 'center':
                return this.docx.AlignmentType.CENTER;
            case 'right':
                return this.docx.AlignmentType.RIGHT;
            case 'justify':
                return this.docx.AlignmentType.JUSTIFIED;
            case 'left':
            default:
                return this.docx.AlignmentType.LEFT;
        }
    }

    processList(listElement, elements) {
        const isOrdered = listElement.tagName.toLowerCase() === 'ol';
        const listItems = listElement.querySelectorAll('li');
        
        listItems.forEach((item, index) => {
            const children = this.processInlineElements(item);
            if (children.length > 0) {
                const listItem = new this.docx.Paragraph({
                    children: children,
                    spacing: { after: 100 },
                    numbering: {
                        type: isOrdered ? this.docx.NumberFormatType.DECIMAL : this.docx.NumberFormatType.BULLET,
                        level: 0
                    }
                });
                elements.push(listItem);
            }
        });
    }

    processTable(tableElement, elements) {
        const rows = tableElement.querySelectorAll('tr');
        const tableRows = [];
        
        rows.forEach(row => {
            const cells = row.querySelectorAll('td, th');
            const tableRowChildren = [];
            
            cells.forEach(cell => {
                // Process all content in the cell, including images
                const cellContent = [];
                this.processNode(cell, cellContent);
                
                // Get cell styling
                const cellStyle = this.getCellStyling(cell);
                
                // Create table cell with proper styling and content
                const tableCell = new this.docx.TableCell({
                    children: cellContent.length > 0 ? cellContent : [new this.docx.Paragraph({ text: ' ' })],
                    width: { size: 100, type: this.docx.WidthType.PERCENTAGE },
                    margins: {
                        top: 100,
                        bottom: 100,
                        left: 100,
                        right: 100
                    },
                    verticalAlign: cellStyle.verticalAlign
                });
                
                // Apply shading if available
                if (cellStyle.shading) {
                    tableCell.shading = cellStyle.shading;
                }
                tableRowChildren.push(tableCell);
            });
            
            const tableRow = new this.docx.TableRow({
                children: tableRowChildren
            });
            tableRows.push(tableRow);
        });
        
        if (tableRows.length > 0) {
            const table = new this.docx.Table({
                rows: tableRows,
                width: { size: 100, type: this.docx.WidthType.PERCENTAGE },
                borders: {
                    top: { style: this.docx.BorderStyle.SINGLE, size: 1 },
                    bottom: { style: this.docx.BorderStyle.SINGLE, size: 1 },
                    left: { style: this.docx.BorderStyle.SINGLE, size: 1 },
                    right: { style: this.docx.BorderStyle.SINGLE, size: 1 },
                    insideHorizontal: { style: this.docx.BorderStyle.SINGLE, size: 1 },
                    insideVertical: { style: this.docx.BorderStyle.SINGLE, size: 1 }
                }
            });
            elements.push(table);
        }
    }

    getCellStyling(cell) {
        const style = cell.getAttribute('style') || '';
        const className = cell.className || '';
        const isHeader = cell.tagName.toLowerCase() === 'th';
        
        let shading = undefined;
        let verticalAlign = this.docx.VerticalAlign.TOP;
        
        // Check if cell has any child elements with background colors
        const cellChildren = cell.querySelectorAll('*');
        
        // Look for background colors in child elements (TinyMCE might apply colors to content)
        for (let child of cellChildren) {
            const childStyle = child.getAttribute('style') || '';
            const childComputedStyle = window.getComputedStyle(child);
            const childBgColor = childComputedStyle.backgroundColor;
            
            if (childStyle.includes('background-color:')) {
                const bgMatch = childStyle.match(/background-color:\s*([^;]+)/);
                if (bgMatch) {
                    const color = this.convertColorToHex(bgMatch[1].trim());
                    if (color) {
                        shading = { fill: color };
                        break;
                    }
                }
            } else if (childBgColor && childBgColor !== 'rgba(0, 0, 0, 0)' && childBgColor !== 'transparent') {
                const color = this.convertColorToHex(childBgColor);
                if (color) {
                    shading = { fill: color };
                    break;
                }
            }
        }
        
        // Handle background colors from inline styles on the cell itself
        if (!shading && style.includes('background-color:')) {
            const bgMatch = style.match(/background-color:\s*([^;]+)/);
            if (bgMatch) {
                const color = this.convertColorToHex(bgMatch[1].trim());
                if (color) {
                    shading = { fill: color };
                }
            }
        }
        
        // Check computed styles as fallback
        if (!shading) {
            const computedStyle = window.getComputedStyle(cell);
            const bgColor = computedStyle.backgroundColor;
            
            if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent') {
                const color = this.convertColorToHex(bgColor);
                if (color) {
                    shading = { fill: color };
                }
            }
        }
        
        // Check for TinyMCE-specific background color attributes
        if (!shading) {
            const bgColorAttr = cell.getAttribute('bgcolor');
            if (bgColorAttr) {
                const color = this.convertColorToHex(bgColorAttr);
                if (color) {
                    shading = { fill: color };
                }
            }
        }
        
        // Check for data attributes that TinyMCE might use
        if (!shading) {
            const dataBgColor = cell.getAttribute('data-mce-bgcolor') || cell.getAttribute('data-bgcolor');
            if (dataBgColor) {
                const color = this.convertColorToHex(dataBgColor);
                if (color) {
                    shading = { fill: color };
                }
            }
        }
        
        // Handle header styling (yellow background for headers)
        if (isHeader && !shading) {
            shading = { fill: 'FFFF00' }; // Yellow for headers
        }
        
        // Handle vertical alignment
        if (style.includes('vertical-align: middle') || style.includes('vertical-align:middle')) {
            verticalAlign = this.docx.VerticalAlign.CENTER;
        } else if (style.includes('vertical-align: bottom') || style.includes('vertical-align:bottom')) {
            verticalAlign = this.docx.VerticalAlign.BOTTOM;
        }
        
        // Handle common CSS classes
        if (className.includes('header') || className.includes('th')) {
            if (!shading) {
                shading = { fill: 'FFFF00' }; // Yellow for headers
            }
        }
        
        return { shading, verticalAlign };
    }

    convertColorToHex(color) {
        // Handle named colors
        const colorMap = {
            'red': 'FF0000',
            'green': '00FF00',
            'blue': '0000FF',
            'yellow': 'FFFF00',
            'cyan': '00FFFF',
            'magenta': 'FF00FF',
            'black': '000000',
            'white': 'FFFFFF',
            'gray': '808080',
            'grey': '808080',
            'lightgray': 'D3D3D3',
            'lightgrey': 'D3D3D3',
            'darkgray': '404040',
            'darkgrey': '404040',
            'orange': 'FFA500',
            'purple': '800080',
            'brown': 'A52A2A',
            'pink': 'FFC0CB',
            'lime': '00FF00',
            'navy': '000080',
            'teal': '008080',
            'silver': 'C0C0C0',
            'gold': 'FFD700',
            'indigo': '4B0082',
            'violet': 'EE82EE',
            'coral': 'FF7F50',
            'salmon': 'FA8072',
            'khaki': 'F0E68C',
            'plum': 'DDA0DD',
            'turquoise': '40E0D0',
            'azure': 'F0FFFF',
            'ivory': 'FFFFF0',
            'wheat': 'F5DEB3',
            'beige': 'F5F5DC',
            'lavender': 'E6E6FA',
            'mint': 'F5FFFA',
            'peach': 'FFDAB9',
            'cream': 'FFFDD0',
            'rose': 'FFE4E1'
        };
        
        // Remove spaces and convert to lowercase
        color = color.toLowerCase().replace(/\s/g, '');
        
        // Check if it's a named color
        if (colorMap[color]) {
            return colorMap[color];
        }
        
        // Handle hex colors
        if (color.startsWith('#')) {
            return color.substring(1).toUpperCase();
        }
        
        // Handle rgb/rgba colors
        if (color.startsWith('rgb')) {
            const match = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
            if (match) {
                const r = parseInt(match[1]).toString(16).padStart(2, '0');
                const g = parseInt(match[2]).toString(16).padStart(2, '0');
                const b = parseInt(match[3]).toString(16).padStart(2, '0');
                return (r + g + b).toUpperCase();
            }
        }
        
        // Handle hsl/hsla colors
        if (color.startsWith('hsl')) {
            // Simple conversion for common HSL values
            if (color.includes('hsl(0, 0%, 0%)')) return '000000'; // black
            if (color.includes('hsl(0, 0%, 100%)')) return 'FFFFFF'; // white
            if (color.includes('hsl(0, 0%, 50%)')) return '808080'; // gray
            // Add more HSL conversions as needed
        }
        
        return undefined;
    }

    processBlockquote(element, elements) {
        const children = this.processInlineElements(element);
        if (children.length > 0) {
            const blockquote = new this.docx.Paragraph({
                children: children,
                spacing: { before: 200, after: 200 },
                indent: { left: 720, right: 720 },
                border: {
                    left: { space: 4, color: 'CCCCCC', style: this.docx.BorderStyle.SINGLE }
                }
            });
            elements.push(blockquote);
        }
    }

    createTitlePage(project) {
        // Get document info data from localStorage
        const allInfo = JSON.parse(localStorage.getItem('bytedraft_docinfo') || '{}');
        const docInfo = allInfo[project.id] || {
            title: '',
            author: '',
            docOwner: '',
            procOwner: '',
            version: '',
            effDate: '',
            lastRev: '',
            nextRev: '',
            link: ''
        };
        
        // Create title page with proper positioning
        return [
            // 5 empty lines to position title 5 lines down from top
            new this.docx.Paragraph({
                text: '',
                spacing: { before: 0, after: 1500 } // 5 lines worth of space (300 per line)
            }),
            
            // Project title - 36pt font size
            new this.docx.Paragraph({
                children: [
                    new this.docx.TextRun({
                        text: project.name,
                        size: 72, // 36pt = 72 half-points
                        font: 'Calibri',
                        color: '2563eb' // Blue color matching the image
                    })
                ],
                alignment: this.docx.AlignmentType.CENTER,
                spacing: { before: 0, after: 400 }
            }),
            
            // Project description - 18pt font size
            new this.docx.Paragraph({
                children: [
                    new this.docx.TextRun({
                        text: project.description || 'No description provided',
                        size: 36, // 18pt = 36 half-points
                        font: 'Calibri',
                        color: '000000' // Black color
                    })
                ],
                alignment: this.docx.AlignmentType.CENTER,
                spacing: { before: 0, after: 200 }
            }),
            
            // Add spacing to push document info to bottom (14 lines up from bottom - moved up 4 lines)
            new this.docx.Paragraph({
                text: '',
                spacing: { before: 0, after: 4200 } // Space to push info box to bottom (14 lines = 4200, reduced from 18 lines)
            }),
            
            // Add additional spacing to ensure title page fills properly
            new this.docx.Paragraph({
                text: '',
                spacing: { before: 0, after: 2000 } // Additional spacing to fill the page
            }),
            
            // Document information box positioned 5 lines up from bottom
            new this.docx.Paragraph({
                children: [
                    new this.docx.TextRun({
                        text: 'Document Information',
                        size: 32, // 16pt = 32 half-points
                        font: 'Calibri',
                        color: '2563eb', // Blue color matching the image
                        bold: true
                    })
                ],
                alignment: this.docx.AlignmentType.CENTER,
                spacing: { before: 0, after: 200 }
            }),
            
            // Document info table using data from Edit Document Info
            new this.docx.Table({
                rows: [
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Document Title:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' } // Dark blue background
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.title || project.name })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Author:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' }
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.author || '' })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Document Owner:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' }
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.docOwner || '' })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Process Owner:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' }
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.procOwner || '' })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Version No:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' }
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.version || '' })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Effective Date:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' }
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.effDate || '' })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Last Reviewed Date:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' }
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.lastRev || '' })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Next Reviewed Date:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' }
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.nextRev || '' })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new this.docx.TableRow({
                        children: [
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ 
                                    children: [new this.docx.TextRun({ 
                                        text: 'Document Link:', 
                                        bold: true,
                                        color: 'FFFFFF'
                                    })]
                                })],
                                width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                shading: { fill: '002060' }
                            }),
                            new this.docx.TableCell({
                                children: [new this.docx.Paragraph({ text: docInfo.link || '' })],
                                width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                            })
                        ]
                    })
                ],
                width: { size: 80, type: this.docx.WidthType.PERCENTAGE },
                alignment: this.docx.AlignmentType.CENTER
            })
        ];
    }

    createDocumentInfoTable(project) {
        const info = [
            ['Document Title:', project.name],
            ['Status:', project.status || 'Draft'],
            ['Created Date:', new Date(project.createdAt).toLocaleDateString()],
            ['Last Updated:', new Date(project.updatedAt).toLocaleDateString()],
            ['Description:', project.description || 'No description provided']
        ];
        
        const tableRows = info.map(([label, value]) => 
            new this.docx.TableRow({
                children: [
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ text: label })],
                        width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                        shading: { fill: 'F0F0F0' }
                    }),
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ text: value || '' })],
                        width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                    })
                ]
            })
        );
        
        return [
            new this.docx.Paragraph({
                text: 'Document Information',
                heading: this.docx.HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 }
            }),
            new this.docx.Table({
                rows: tableRows,
                width: { size: 100, type: this.docx.WidthType.PERCENTAGE }
            }),
            new this.docx.Paragraph({
                text: '',
                spacing: { after: 400 }
            })
        ];
    }

    createVersionHistorySection(versionHistory) {
        if (!versionHistory || versionHistory.length === 0) {
            return [];
        }
        
        const tableRows = versionHistory.slice(-10).map(version => 
            new this.docx.TableRow({
                children: [
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ text: new Date(version.timestamp).toLocaleDateString() })],
                        width: { size: 30, type: this.docx.WidthType.PERCENTAGE }
                    }),
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ text: version.description })],
                        width: { size: 70, type: this.docx.WidthType.PERCENTAGE }
                    })
                ]
            })
        );
        
        return [
            new this.docx.Paragraph({
                text: 'Version History',
                heading: this.docx.HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 }
            }),
            new this.docx.Table({
                rows: tableRows,
                width: { size: 100, type: this.docx.WidthType.PERCENTAGE }
            }),
            new this.docx.Paragraph({
                text: '',
                spacing: { after: 400 }
            })
        ];
    }

    createChangelogSection(customChangelog) {
        if (!customChangelog || customChangelog.trim() === '') {
            return [];
        }
        
        return [
            new this.docx.Paragraph({
                text: 'Changelog',
                heading: this.docx.HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 }
            }),
            new this.docx.Paragraph({
                text: customChangelog,
                spacing: { after: 400 }
            })
        ];
    }

    createChangelogPage(project) {
        // Get changelog data from localStorage
        const allChangelog = JSON.parse(localStorage.getItem('bytedraft_custom_changelog') || '{}');
        const changelogData = allChangelog[project.id] ? JSON.parse(allChangelog[project.id]) : [];
        
        if (changelogData.length === 0) {
            // If no changelog data, return empty page with just title
            return [
                new this.docx.Paragraph({
                    text: '',
                    pageBreakBefore: true
                }),
                new this.docx.Paragraph({
                    children: [
                        new this.docx.TextRun({
                            text: 'Document Changelog',
                            size: 48, // 24pt = 48 half-points
                            font: 'Calibri',
                            bold: true,
                            color: '2563eb' // Blue color
                        })
                    ],
                    alignment: this.docx.AlignmentType.CENTER,
                    spacing: { before: 400, after: 400 }
                }),
                new this.docx.Paragraph({
                    text: 'No changelog entries available.',
                    alignment: this.docx.AlignmentType.CENTER,
                    spacing: { before: 200, after: 200 }
                }),
                new this.docx.Paragraph({
                    text: '',
                    pageBreakBefore: true
                })
            ];
        }

        // Create changelog table
        const tableRows = [
            // Header row
            new this.docx.TableRow({
                children: [
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ 
                            children: [new this.docx.TextRun({ 
                                text: 'Version Number', 
                                bold: true,
                                color: 'FFFFFF'
                            })]
                        })],
                        width: { size: 15, type: this.docx.WidthType.PERCENTAGE },
                        shading: { fill: '002060' }
                    }),
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ 
                            children: [new this.docx.TextRun({ 
                                text: 'Approved Date', 
                                bold: true,
                                color: 'FFFFFF'
                            })]
                        })],
                        width: { size: 15, type: this.docx.WidthType.PERCENTAGE },
                        shading: { fill: '002060' }
                    }),
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ 
                            children: [new this.docx.TextRun({ 
                                text: 'Author', 
                                bold: true,
                                color: 'FFFFFF'
                            })]
                        })],
                        width: { size: 15, type: this.docx.WidthType.PERCENTAGE },
                        shading: { fill: '002060' }
                    }),
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ 
                            children: [new this.docx.TextRun({ 
                                text: 'Reviewer', 
                                bold: true,
                                color: 'FFFFFF'
                            })]
                        })],
                        width: { size: 15, type: this.docx.WidthType.PERCENTAGE },
                        shading: { fill: '002060' }
                    }),
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ 
                            children: [new this.docx.TextRun({ 
                                text: 'Approver', 
                                bold: true,
                                color: 'FFFFFF'
                            })]
                        })],
                        width: { size: 15, type: this.docx.WidthType.PERCENTAGE },
                        shading: { fill: '002060' }
                    }),
                    new this.docx.TableCell({
                        children: [new this.docx.Paragraph({ 
                            children: [new this.docx.TextRun({ 
                                text: 'Description', 
                                bold: true,
                                color: 'FFFFFF'
                            })]
                        })],
                        width: { size: 25, type: this.docx.WidthType.PERCENTAGE },
                        shading: { fill: '002060' }
                    })
                ]
            })
        ];

        // Add data rows
        changelogData.forEach(row => {
            tableRows.push(
                new this.docx.TableRow({
                    children: [
                        new this.docx.TableCell({
                            children: [new this.docx.Paragraph({ text: row.version || '' })],
                            width: { size: 15, type: this.docx.WidthType.PERCENTAGE }
                        }),
                        new this.docx.TableCell({
                            children: [new this.docx.Paragraph({ text: row.date || '' })],
                            width: { size: 15, type: this.docx.WidthType.PERCENTAGE }
                        }),
                        new this.docx.TableCell({
                            children: [new this.docx.Paragraph({ text: row.author || '' })],
                            width: { size: 15, type: this.docx.WidthType.PERCENTAGE }
                        }),
                        new this.docx.TableCell({
                            children: [new this.docx.Paragraph({ text: row.reviewer || '' })],
                            width: { size: 15, type: this.docx.WidthType.PERCENTAGE }
                        }),
                        new this.docx.TableCell({
                            children: [new this.docx.Paragraph({ text: row.approver || '' })],
                            width: { size: 15, type: this.docx.WidthType.PERCENTAGE }
                        }),
                        new this.docx.TableCell({
                            children: [new this.docx.Paragraph({ text: row.desc || '' })],
                            width: { size: 25, type: this.docx.WidthType.PERCENTAGE }
                        })
                    ]
                })
            );
        });

        return [
            // Changelog title
            new this.docx.Paragraph({
                children: [
                    new this.docx.TextRun({
                        text: 'Document Change Log',
                        size: 48, // 24pt = 48 half-points
                        font: 'Calibri',
                        bold: true,
                        color: '2563eb' // Blue color
                    })
                ],
                alignment: this.docx.AlignmentType.CENTER,
                spacing: { before: 400, after: 400 }
            }),
            
            // Changelog table
            new this.docx.Table({
                rows: tableRows,
                width: { size: 100, type: this.docx.WidthType.PERCENTAGE },
                alignment: this.docx.AlignmentType.CENTER
            }),
            


        ];
    }

    createTOCPage(project) {
        return [
            // Page break to start TOC page
            new this.docx.Paragraph({
                text: '',
                pageBreakBefore: true
            }),
            
            // TOC title
            new this.docx.Paragraph({
                children: [
                    new this.docx.TextRun({
                        text: 'Table of Contents',
                        size: 48, // 24pt = 48 half-points
                        font: 'Calibri',
                        bold: true,
                        color: '2563eb' // Blue color
                    })
                ],
                alignment: this.docx.AlignmentType.CENTER,
                spacing: { before: 400, after: 400 }
            }),
            
            // Auto-updating TOC field using the proper TableOfContents class
            new this.docx.TableOfContents("Click here to update the table of contents", {
                headingStyleRange: "1-3",
                hyperlink: true,
                useAppliedParagraphOutlineLevel: true,
                preserveTabInEntries: true,
                preserveNewLineInEntries: true,
                hideTabAndPageNumbersInWebView: true
            })
        ];
    }
}

// Global function for backward compatibility
async function exportProjectToDocxModern(project, customChangelog = '', versionHistory = []) {
    const exporter = new ModernDocxExporter();
    return await exporter.exportProjectToDocx(project, customChangelog, versionHistory);
}

// Make it available globally
window.exportProjectToDocxModern = exportProjectToDocxModern;
window.ModernDocxExporter = ModernDocxExporter; 