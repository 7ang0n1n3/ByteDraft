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
                            ...docxElements,

                            // References page (only if citations exist)
                            ...this.createReferencesPage(project)
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
        

        
        // Convert TinyMCE's complex styling to simpler HTML structure
        // TinyMCE often uses spans with complex styling, we need to convert these to semantic HTML
        
        // Note: Span conversion is now handled in the comprehensive span processing below
        
        // Handle TinyMCE's specific styling patterns - preserve spans and clean them
        const allSpans = tempDiv.querySelectorAll('span');
        allSpans.forEach(span => {
            const style = span.getAttribute('style') || '';
            const className = span.getAttribute('class') || '';
            

            
            // Skip spans that have background colors - preserve them for table processing
            if (style.includes('background-color') || style.includes('background:')) {
                return; // Don't convert spans with background colors
            }
            
            // Clean up the span style, preserving only essential formatting
            let cleanStyle = '';
            
            // Preserve font-size
            const fontSizeMatch = style.match(/font-size\s*:\s*[^;]+;?/);
            if (fontSizeMatch) {
                cleanStyle += fontSizeMatch[0].trim() + ' ';
            }
            
            // Preserve font-weight
            const fontWeightMatch = style.match(/font-weight\s*:\s*[^;]+;?/);
            if (fontWeightMatch) {
                cleanStyle += fontWeightMatch[0].trim() + ' ';
            }
            
            // Preserve font-style
            const fontStyleMatch = style.match(/font-style\s*:\s*[^;]+;?/);
            if (fontStyleMatch) {
                cleanStyle += fontStyleMatch[0].trim() + ' ';
            }
            
            // Preserve text-decoration
            const textDecorationMatch = style.match(/text-decoration\s*:\s*[^;]+;?/);
            if (textDecorationMatch) {
                cleanStyle += textDecorationMatch[0].trim() + ' ';
            }

            // Preserve color (use anchored match to avoid matching background-color)
            const colorMatch = style.match(/(?:^|;)\s*color\s*:\s*([^;]+)/);
            if (colorMatch) {
                cleanStyle += `color: ${colorMatch[1].trim()}; `;
            }

            // Preserve font-family
            const fontFamilyMatch = style.match(/font-family\s*:\s*[^;]+;?/);
            if (fontFamilyMatch) {
                cleanStyle += fontFamilyMatch[0].trim() + ' ';
            }

            // Update the span style
            if (cleanStyle.trim()) {
                span.setAttribute('style', cleanStyle.trim());
            } else {
                span.removeAttribute('style');
            }
            
            // Remove class attribute as it might interfere
            if (className) {
                span.removeAttribute('class');
            }
        });
        
        // Remove remaining problematic inline styles that might interfere
        const elementsWithStyles = tempDiv.querySelectorAll('[style]');
        elementsWithStyles.forEach(el => {
            // Keep only essential styles, remove others
            const style = el.getAttribute('style');
            if (style) {
                // Special handling for table cells - preserve background colors and text alignment
                if (el.tagName.toLowerCase() === 'td' || el.tagName.toLowerCase() === 'th') {
                    // For table cells, keep background-color and text-align, remove everything else
                    let cleanStyle = '';
                    const backgroundMatch = style.match(/background-color\s*:\s*[^;]+;?/);
                    const textAlignMatch = style.match(/text-align\s*:\s*[^;]+;?/);
                    
                    if (backgroundMatch) {
                        cleanStyle += backgroundMatch[0].trim() + ' ';
                    }
                    if (textAlignMatch) {
                        cleanStyle += textAlignMatch[0].trim() + ' ';
                    }
                    
                    if (cleanStyle.trim()) {
                        el.setAttribute('style', cleanStyle.trim());
                    } else {
                        el.removeAttribute('style');
                    }
                } else if (el.tagName.toLowerCase() === 'table') {
                    // For tables, preserve width, border-collapse, border-style, border-width, margin
                    let cleanStyle = '';
                    const widthMatch = style.match(/width\s*:\s*[^;]+;?/);
                    const borderCollapseMatch = style.match(/border-collapse\s*:\s*[^;]+;?/);
                    const borderStyleMatch = style.match(/border-style\s*:\s*[^;]+;?/);
                    const borderWidthMatch = style.match(/border-width\s*:\s*[^;]+;?/);
                    const marginMatch = style.match(/margin[^;]*;?/g);
                    
                    if (widthMatch) cleanStyle += widthMatch[0].trim() + ' ';
                    if (borderCollapseMatch) cleanStyle += borderCollapseMatch[0].trim() + ' ';
                    if (borderStyleMatch) cleanStyle += borderStyleMatch[0].trim() + ' ';
                    if (borderWidthMatch) cleanStyle += borderWidthMatch[0].trim() + ' ';
                    if (marginMatch) {
                        marginMatch.forEach(margin => {
                            cleanStyle += margin.trim() + ' ';
                        });
                    }
                    
                    if (cleanStyle.trim()) {
                        el.setAttribute('style', cleanStyle.trim());
                    } else {
                        el.removeAttribute('style');
                    }
                } else if (el.tagName.toLowerCase() === 'tr') {
                    // For table rows, preserve height
                    let cleanStyle = '';
                    const heightMatch = style.match(/height\s*:\s*[^;]+;?/);
                    
                    if (heightMatch) {
                        cleanStyle += heightMatch[0].trim() + ' ';
                    }
                    
                    if (cleanStyle.trim()) {
                        el.setAttribute('style', cleanStyle.trim());
                    } else {
                        el.removeAttribute('style');
                    }
                } else {
                    // For other elements: strip layout-only properties but preserve
                    // all text-formatting properties (color, background-color,
                    // font-size, font-weight, font-style, font-family,
                    // text-decoration, text-align, vertical-align, etc.)
                    const cleanStyle = style
                        .replace(/\bmargin[^;]*;?/g, '')
                        .replace(/\bpadding[^;]*;?/g, '')
                        .replace(/\bdisplay[^;]*;?/g, '')
                        .replace(/\bfloat[^;]*;?/g, '')
                        .replace(/\bposition[^;]*;?/g, '')
                        .replace(/\bwidth[^;]*;?/g, '')
                        .replace(/\bheight[^;]*;?/g, '')
                        .trim();

                    if (cleanStyle) {
                        el.setAttribute('style', cleanStyle);
                    } else {
                        el.removeAttribute('style');
                    }
                }
            }
        });
        
        // Remove class attributes that might interfere
        const elementsWithClasses = tempDiv.querySelectorAll('[class]');
        elementsWithClasses.forEach(el => {
            el.removeAttribute('class');
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
        
        // Ensure proper heading structure
        const headings = tempDiv.querySelectorAll('h1, h2, h3, h4, h5, h6');
        headings.forEach(heading => {
            // Remove any remaining inline styles from headings
            heading.removeAttribute('style');
            heading.removeAttribute('class');
        });
        
        // Clean up list formatting
        const lists = tempDiv.querySelectorAll('ul, ol');
        lists.forEach(list => {
            list.removeAttribute('style');
            list.removeAttribute('class');
            
            // Ensure list items are properly structured
            const listItems = list.querySelectorAll('li');
            listItems.forEach(item => {
                item.removeAttribute('style');
                item.removeAttribute('class');
            });
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
                case 'p': {
                    const textAlign = this.getTextAlignment(node);
                    const imgElements = node.querySelectorAll('img');
                    if (imgElements.length > 0) {
                        // Process each image in the paragraph
                        imgElements.forEach(img => {
                            this.processImage(img, elements);
                        });
                        // Also push any inline text from the paragraph
                        const children = this.processInlineElements(node);
                        if (children.length > 0) {
                            elements.push(new this.docx.Paragraph({
                                children,
                                alignment: textAlign,
                                spacing: { after: 200 }
                            }));
                        }
                    } else if (node.textContent.trim()) {
                        const children = this.processInlineElements(node);
                        if (children.length > 0) {
                            elements.push(new this.docx.Paragraph({
                                children,
                                alignment: textAlign,
                                spacing: { after: 200 }
                            }));
                        }
                    }
                    break;
                }
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

    // processInlineElements walks child nodes and returns an array of TextRun / hyperlink
    // objects ready to pass as `children` to a docx.Paragraph.
    //
    // Formatting is accumulated in `inheritedProps` and passed down the recursion tree
    // so that every property is present when the TextRun is *constructed* — the docx
    // library builds its internal XML at construction time and ignores post-construction
    // property assignments.
    processInlineElements(element, inheritedProps = {}) {
        const children = [];

        Array.from(element.childNodes).forEach(child => {
            if (child.nodeType === Node.TEXT_NODE) {
                const text = child.textContent;
                if (text) {
                    children.push(new this.docx.TextRun({ text, ...inheritedProps }));
                }
            } else if (child.nodeType === Node.ELEMENT_NODE) {
                const tagName = child.tagName.toLowerCase();

                switch (tagName) {
                    case 'br':
                        children.push(new this.docx.TextRun({ text: '\n' }));
                        break;
                    case 'strong':
                    case 'b':
                        children.push(...this.processInlineElements(child, { ...inheritedProps, bold: true }));
                        break;
                    case 'em':
                    case 'i':
                        children.push(...this.processInlineElements(child, { ...inheritedProps, italics: true }));
                        break;
                    case 'u':
                        children.push(...this.processInlineElements(child, {
                            ...inheritedProps,
                            underline: { type: this.docx.UnderlineType.SINGLE, color: '000000' }
                        }));
                        break;
                    case 's':
                    case 'strike':
                        children.push(...this.processInlineElements(child, { ...inheritedProps, strike: true }));
                        break;
                    case 'sup':
                        children.push(...this.processInlineElements(child, { ...inheritedProps, superScript: true }));
                        break;
                    case 'sub':
                        children.push(...this.processInlineElements(child, { ...inheritedProps, subScript: true }));
                        break;
                    case 'code':
                        children.push(...this.processInlineElements(child, {
                            ...inheritedProps,
                            font: 'Courier New',
                            size: inheritedProps.size || 20
                        }));
                        break;
                    case 'mark':
                        children.push(...this.processInlineElements(child, { ...inheritedProps, highlight: 'yellow' }));
                        break;
                    case 'small':
                        children.push(...this.processInlineElements(child, { ...inheritedProps, size: 16 }));
                        break;
                    case 'big':
                        children.push(...this.processInlineElements(child, { ...inheritedProps, size: 28 }));
                        break;
                    case 'a': {
                        const href = child.getAttribute('href') || '';
                        const linkText = child.textContent || '';
                        if (!linkText.trim()) break;
                        const linkProps = { ...inheritedProps, color: '0563C1', underline: { type: 'single' } };
                        const linkRuns = this.processInlineElements(child, linkProps);
                        if (href && linkRuns.length > 0) {
                            children.push(new this.docx.ExternalHyperlink({ children: linkRuns, link: href }));
                        } else {
                            children.push(...linkRuns);
                        }
                        break;
                    }
                    case 'img':
                        // Block-level images are handled by the caller; skip here
                        break;
                    case 'span': {
                        const spanProps = this.extractSpanProps(child);
                        children.push(...this.processInlineElements(child, { ...inheritedProps, ...spanProps }));
                        break;
                    }
                    default:
                        children.push(...this.processInlineElements(child, inheritedProps));
                        break;
                }
            }
        });

        return children;
    }

    // Extract docx-compatible run properties from a <span> element's inline style.
    extractSpanProps(span) {
        const style = span.style || {};
        const props = {};

        const fontWeight = style.fontWeight || '';
        if (fontWeight === 'bold' || fontWeight === 'bolder' || parseInt(fontWeight) >= 700) {
            props.bold = true;
        }

        const fontStyle = style.fontStyle || '';
        if (fontStyle === 'italic') {
            props.italics = true;
        }

        const textDecoration = style.textDecoration || '';
        if (textDecoration.includes('underline')) {
            props.underline = { type: this.docx.UnderlineType.SINGLE, color: '000000' };
        }
        if (textDecoration.includes('line-through')) {
            props.strike = true;
        }

        const verticalAlign = style.verticalAlign || '';
        if (verticalAlign === 'super') props.superScript = true;
        else if (verticalAlign === 'sub') props.subScript = true;

        const fontSize = style.fontSize || '';
        if (fontSize) {
            const size = this.parseFontSize(fontSize);
            if (size) props.size = size;
        }

        // Text color (TinyMCE outputs rgb() or #hex)
        const color = style.color || '';
        if (color) {
            const hex = this.convertColorToHex(color);
            if (hex) props.color = hex;
        }

        // Background color → TextRun shading (TinyMCE outputs background-color on spans)
        const bgColor = style.backgroundColor || '';
        if (bgColor) {
            const hex = this.convertColorToHex(bgColor);
            if (hex) props.shading = { fill: hex, val: 'clear' };
        }

        // Font family — take the first name in the stack
        const fontFamily = style.fontFamily || '';
        if (fontFamily) {
            const cleanFamily = fontFamily.replace(/['"]/g, '').split(',')[0].trim();
            if (cleanFamily) props.font = cleanFamily;
        }

        return props;
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
        
        Array.from(listElement.children).forEach((item, index) => {
            // Process list item content with formatting
            const children = this.processInlineElements(item);
            
            if (children.length > 0) {
                // Create bullet or number prefix
                const prefix = isOrdered ? `${index + 1}. ` : '• ';
                const prefixRun = new this.docx.TextRun({ text: prefix, bold: true });
                
                // Combine prefix with content
                const allChildren = [prefixRun, ...children];
                
                const paragraph = new this.docx.Paragraph({
                    children: allChildren,
                    spacing: { after: 100 },
                    indent: { left: 576 } // 576 twips = 0.4 inches (four spaces)
                });
                
                elements.push(paragraph);
            }
        });
    }

    // Process the block-level children of a table cell, returning an array of
    // Paragraph objects for use as TableCell.children.
    // `defaultAlignment` is the cell-level alignment used when a child paragraph
    // carries no explicit text-align of its own.
    processCellContent(cellElement, defaultAlignment) {
        const paragraphs = [];

        Array.from(cellElement.childNodes).forEach(node => {
            if (node.nodeType === Node.TEXT_NODE) {
                const text = node.textContent.trim();
                if (text) {
                    paragraphs.push(new this.docx.Paragraph({
                        children: [new this.docx.TextRun({ text })],
                        alignment: defaultAlignment
                    }));
                }
            } else if (node.nodeType === Node.ELEMENT_NODE) {
                const tagName = node.tagName.toLowerCase();

                switch (tagName) {
                    case 'p': {
                        const runs = this.processInlineElements(node);
                        if (runs.length > 0) {
                            // Honour paragraph-level alignment; fall back to cell default
                            const pStyle = node.style || {};
                            const pAlign = pStyle.textAlign || pStyle['text-align'] || '';
                            let alignment = defaultAlignment;
                            switch (pAlign) {
                                case 'center':  alignment = this.docx.AlignmentType.CENTER;    break;
                                case 'right':   alignment = this.docx.AlignmentType.RIGHT;     break;
                                case 'justify': alignment = this.docx.AlignmentType.JUSTIFIED; break;
                                case 'left':    alignment = this.docx.AlignmentType.LEFT;      break;
                            }
                            paragraphs.push(new this.docx.Paragraph({ children: runs, alignment }));
                        }
                        break;
                    }
                    case 'ul':
                    case 'ol': {
                        const isOrdered = tagName === 'ol';
                        Array.from(node.children).forEach((li, idx) => {
                            const runs = this.processInlineElements(li);
                            const prefix = isOrdered ? `${idx + 1}. ` : '\u2022 ';
                            const prefixRun = new this.docx.TextRun({ text: prefix });
                            paragraphs.push(new this.docx.Paragraph({
                                children: [prefixRun, ...(runs.length > 0 ? runs : [new this.docx.TextRun({ text: li.textContent })])],
                                indent: { left: 360 },
                                alignment: defaultAlignment
                            }));
                        });
                        break;
                    }
                    case 'h1': case 'h2': case 'h3':
                    case 'h4': case 'h5': case 'h6': {
                        const level = parseInt(tagName[1]);
                        const runs = this.processInlineElements(node);
                        paragraphs.push(new this.docx.Paragraph({
                            children: runs.length > 0 ? runs : [new this.docx.TextRun({ text: node.textContent, bold: true })],
                            heading: this.docx.HeadingLevel[`HEADING_${level}`]
                        }));
                        break;
                    }
                    case 'img':
                        this.processImage(node, paragraphs);
                        break;
                    case 'br':
                        paragraphs.push(new this.docx.Paragraph({ children: [new this.docx.TextRun({ text: '' })] }));
                        break;
                    default: {
                        // div, span wrappers, etc. — recurse block-style then fall back to inline
                        const nested = this.processCellContent(node, defaultAlignment);
                        if (nested.length > 0) {
                            paragraphs.push(...nested);
                        } else {
                            const runs = this.processInlineElements(node);
                            if (runs.length > 0) {
                                paragraphs.push(new this.docx.Paragraph({ children: runs, alignment: defaultAlignment }));
                            }
                        }
                        break;
                    }
                }
            }
        });

        return paragraphs;
    }

    processTable(tableElement, elements) {
        const rows = [];
        
        // Process table rows (tr elements)
        const tableRows = tableElement.querySelectorAll('tr');
        
        tableRows.forEach((rowElement, rowIndex) => {
            const cells = [];
            
            // Process table cells (td/th elements) in this row
            const tableCells = rowElement.querySelectorAll('td, th');
            
            tableCells.forEach((cellElement, cellIndex) => {
                // Check for cell merging (colspan and rowspan)
                const colspan = parseInt(cellElement.getAttribute('colspan')) || 1;
                const rowspan = parseInt(cellElement.getAttribute('rowspan')) || 1;
                
                // Skip cells that are part of a colspan (they'll be handled by the main cell)
                if (cellElement.hasAttribute('data-colspan-skip')) {
                    return;
                }
                
                // Extract background color and text alignment from style
                const style = cellElement.getAttribute('style') || '';
                
                // Try different color extraction patterns
                let bgColor = 'FFFFFF'; // default white
                
                // First check the cell's own style
                // Pattern 1: background-color: #RRGGBB
                let bgColorMatch = style.match(/background-color:\s*(#[0-9a-fA-F]{6})/i);
                if (bgColorMatch) {
                    bgColor = bgColorMatch[1];
                } else {
                    // Pattern 2: background-color: #RGB
                    bgColorMatch = style.match(/background-color:\s*(#[0-9a-fA-F]{3})/i);
                    if (bgColorMatch) {
                        bgColor = bgColorMatch[1];
                    } else {
                        // Pattern 3: background-color: rgb(r, g, b)
                        bgColorMatch = style.match(/background-color:\s*rgb\((\d+),\s*(\d+),\s*(\d+)\)/i);
                        if (bgColorMatch) {
                            const r = parseInt(bgColorMatch[1]).toString(16).padStart(2, '0');
                            const g = parseInt(bgColorMatch[2]).toString(16).padStart(2, '0');
                            const b = parseInt(bgColorMatch[3]).toString(16).padStart(2, '0');
                            bgColor = `#${r}${g}${b}`;
                        } else {
                            // Pattern 4: background: #RRGGBB
                            bgColorMatch = style.match(/background:\s*(#[0-9a-fA-F]{6})/i);
                            if (bgColorMatch) {
                                bgColor = bgColorMatch[1];
                            } else {
                                // Pattern 5: background: #RGB
                                bgColorMatch = style.match(/background:\s*(#[0-9a-fA-F]{3})/i);
                                if (bgColorMatch) {
                                    bgColor = bgColorMatch[1];
                                }
                            }
                        }
                    }
                }
                
                // If no background color found on the cell itself, check child elements (spans)
                if (bgColor === 'FFFFFF') {
                    const childElements = cellElement.querySelectorAll('*');
                    for (let child of childElements) {
                        const childStyle = child.getAttribute('style') || '';
                        if (childStyle.includes('background-color') || childStyle.includes('background:')) {
                            // Pattern 1: background-color: #RRGGBB
                            let childBgMatch = childStyle.match(/background-color:\s*(#[0-9a-fA-F]{6})/i);
                            if (childBgMatch) {
                                bgColor = childBgMatch[1];
                                break;
                            } else {
                                // Pattern 2: background-color: #RGB
                                childBgMatch = childStyle.match(/background-color:\s*(#[0-9a-fA-F]{3})/i);
                                if (childBgMatch) {
                                    bgColor = childBgMatch[1];
                                    break;
                                } else {
                                    // Pattern 3: background-color: rgb(r, g, b)
                                    childBgMatch = childStyle.match(/background-color:\s*rgb\((\d+),\s*(\d+),\s*(\d+)\)/i);
                                    if (childBgMatch) {
                                        const r = parseInt(childBgMatch[1]).toString(16).padStart(2, '0');
                                        const g = parseInt(childBgMatch[2]).toString(16).padStart(2, '0');
                                        const b = parseInt(childBgMatch[3]).toString(16).padStart(2, '0');
                                        bgColor = `#${r}${g}${b}`;
                                        break;
                                    } else {
                                        // Pattern 4: background: #RRGGBB
                                        childBgMatch = childStyle.match(/background:\s*(#[0-9a-fA-F]{6})/i);
                                        if (childBgMatch) {
                                            bgColor = childBgMatch[1];
                                            break;
                                        } else {
                                            // Pattern 5: background: #RGB
                                            childBgMatch = childStyle.match(/background:\s*(#[0-9a-fA-F]{3})/i);
                                            if (childBgMatch) {
                                                bgColor = childBgMatch[1];
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
                // Extract text alignment
                let textAlignment = this.docx.AlignmentType.LEFT; // default
                const textAlignMatch = style.match(/text-align:\s*([^;]+)/i);
                if (textAlignMatch) {
                    const align = textAlignMatch[1].trim();
                    switch (align) {
                        case 'center':
                            textAlignment = this.docx.AlignmentType.CENTER;
                            break;
                        case 'right':
                            textAlignment = this.docx.AlignmentType.RIGHT;
                            break;
                        case 'justify':
                            textAlignment = this.docx.AlignmentType.JUSTIFIED;
                            break;
                    }
                }
                
                // Build cell paragraphs - block-by-block to preserve lists, headings,
                // per-paragraph alignment, and embedded images
                const cellParagraphs = this.processCellContent(cellElement, textAlignment);
                if (cellParagraphs.length === 0) {
                    cellParagraphs.push(new this.docx.Paragraph({
                        children: [new this.docx.TextRun({ text: '' })]
                    }));
                }

                // Create table cell with shading on the CELL itself
                let tableCell;

                // Base cell properties
                const cellProperties = {
                    children: cellParagraphs
                };
                
                // Add colspan if greater than 1
                if (colspan > 1) {
                    cellProperties.columnSpan = colspan;
                }
                
                // Add rowspan if greater than 1
                if (rowspan > 1) {
                    cellProperties.rowSpan = rowspan;
                }
                
                // Add shading if color exists
                if (bgColor !== 'FFFFFF') {
                    cellProperties.shading = {
                        fill: bgColor.replace('#', ''),
                        val: 'solid'
                    };
                }
                
                tableCell = new this.docx.TableCell(cellProperties);
                cells.push(tableCell);
            });
            
            // Process row height if specified
            const rowStyle = rowElement.getAttribute('style') || '';
            let rowHeight = undefined;
            const heightMatch = rowStyle.match(/height:\s*([^;]+)/i);
            if (heightMatch) {
                const heightValue = heightMatch[1].trim();
                if (heightValue.includes('px')) {
                    // Convert px to twips (1 px = 20 twips)
                    const pxValue = parseFloat(heightValue.replace('px', ''));
                    rowHeight = Math.round(pxValue * 20);
                }
            }
            
            // Ensure all children are valid before creating TableRow
            const validCells = cells.filter(cell => cell !== null && cell !== undefined);
            
            const rowProperties = { children: validCells };
            
            // Only add height if it's a valid positive number
            if (rowHeight && rowHeight > 0 && !isNaN(rowHeight)) {
                // Use the proper format for row height
                rowProperties.height = {
                    value: rowHeight,
                    rule: this.docx.HeightRule.EXACT
                };
            }
            
            rows.push(new this.docx.TableRow(rowProperties));
        });
        
        // Extract table properties from HTML
        const tableStyle = tableElement.getAttribute('style') || '';
        let tableWidth = 100; // default percentage
        let widthType = this.docx.WidthType.PERCENTAGE;
        let tableAlignment = this.docx.AlignmentType.LEFT; // default
        
        // Try to extract width from style
        const widthMatch = tableStyle.match(/width:\s*([^;]+)/);
        if (widthMatch) {
            const widthValue = widthMatch[1].trim();
            
            if (widthValue.includes('%')) {
                tableWidth = parseFloat(widthValue.replace('%', ''));
                widthType = this.docx.WidthType.PERCENTAGE;
            } else if (widthValue.includes('px')) {
                tableWidth = parseFloat(widthValue.replace('px', ''));
                widthType = this.docx.WidthType.DXA; // Convert to twips later if needed
            }
        }
        
        // Extract table alignment from margins
        const marginLeftMatch = tableStyle.match(/margin-left:\s*([^;]+)/i);
        const marginRightMatch = tableStyle.match(/margin-right:\s*([^;]+)/i);
        if (marginLeftMatch && marginRightMatch) {
            const leftMargin = marginLeftMatch[1].trim();
            const rightMargin = marginRightMatch[1].trim();
            if (leftMargin === 'auto' && rightMargin === 'auto') {
                tableAlignment = this.docx.AlignmentType.CENTER;
            } else if (rightMargin === 'auto') {
                tableAlignment = this.docx.AlignmentType.LEFT;
            } else if (leftMargin === 'auto') {
                tableAlignment = this.docx.AlignmentType.RIGHT;
            }
        }
        
        // Don't force wider table - respect the original width
        // if (tableWidth < 80) {
        //     tableWidth = 100;
        //     widthType = this.docx.WidthType.PERCENTAGE;
        // }
        
        // Extract column widths from colgroup/col elements
        let columnWidths = [];
        const colgroup = tableElement.querySelector('colgroup');
        if (colgroup) {
            const cols = colgroup.querySelectorAll('col');
            
            cols.forEach((col, index) => {
                const colStyle = col.getAttribute('style') || '';
                const widthMatch = colStyle.match(/width:\s*([^;]+)/);
                if (widthMatch) {
                    const widthValue = widthMatch[1].trim();
                    if (widthValue.includes('%')) {
                        const width = parseFloat(widthValue.replace('%', ''));
                        columnWidths.push(width);
                    } else {
                        // Default to equal width if no percentage
                        columnWidths.push(100 / cols.length);
                    }
                } else {
                    // Default to equal width if no style
                    columnWidths.push(100 / cols.length);
                }
            });
        } else {
            // No colgroup found, calculate based on actual DOM cells
            let maxColumns = 1;
            tableRows.forEach(tr => {
                const cellCount = tr.querySelectorAll('td, th').length;
                if (cellCount > maxColumns) maxColumns = cellCount;
            });
            columnWidths = Array(maxColumns).fill(100 / maxColumns);
        }
        
        // Extract border properties
        let borderStyle = this.docx.BorderStyle.SINGLE; // default
        let borderSize = 1; // default
        
        const borderStyleMatch = tableStyle.match(/border-style:\s*([^;]+)/i);
        if (borderStyleMatch) {
            const style = borderStyleMatch[1].trim();
            switch (style) {
                case 'dotted':
                    borderStyle = this.docx.BorderStyle.DOTTED;
                    break;
                case 'dashed':
                    borderStyle = this.docx.BorderStyle.DASHED;
                    break;
                case 'double':
                    borderStyle = this.docx.BorderStyle.DOUBLE;
                    break;
                case 'thick':
                    borderStyle = this.docx.BorderStyle.THICK;
                    break;
                case 'none':
                    borderStyle = this.docx.BorderStyle.NONE;
                    break;
                default:
                    borderStyle = this.docx.BorderStyle.SINGLE;
            }
        }
        
        const borderWidthMatch = tableStyle.match(/border-width:\s*([^;]+)/i);
        if (borderWidthMatch) {
            const width = borderWidthMatch[1].trim();
            if (width.includes('px')) {
                borderSize = parseFloat(width.replace('px', ''));
            }
        }
        
        // Check for HTML border attribute
        const borderAttr = tableElement.getAttribute('border');
        if (borderAttr && !borderStyleMatch) {
            // If border attribute exists but no border-style in CSS, use single border
            borderStyle = this.docx.BorderStyle.SINGLE;
            borderSize = parseInt(borderAttr) || 1;
        }
        

        
        // Only create table if we have rows
        if (rows && rows.length > 0) {
            const tableProperties = {
                rows: rows,
                width: {
                    size: tableWidth,
                    type: widthType
                },
                alignment: tableAlignment
            };
            
            // Add borders if they exist
            if (borderStyle !== this.docx.BorderStyle.NONE) {
                tableProperties.borders = {
                    top: { style: borderStyle, size: borderSize },
                    bottom: { style: borderStyle, size: borderSize },
                    left: { style: borderStyle, size: borderSize },
                    right: { style: borderStyle, size: borderSize },
                    insideHorizontal: { style: borderStyle, size: borderSize },
                    insideVertical: { style: borderStyle, size: borderSize }
                };
            }
            
            elements.push(new this.docx.Table(tableProperties));
        }
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

    parseFontSize(fontSize) {
        // Remove spaces and convert to lowercase
        fontSize = fontSize.toLowerCase().replace(/\s/g, '');
        
        // Handle pixel values (px)
        if (fontSize.endsWith('px')) {
            const size = parseInt(fontSize);
            if (!isNaN(size)) {
                // Convert pixels to half-points (1px = 1pt = 2 half-points)
                const result = Math.round(size * 2);
                return result;
            }
        }
        
        // Handle point values (pt)
        if (fontSize.endsWith('pt')) {
            const size = parseInt(fontSize);
            if (!isNaN(size)) {
                // Convert points to half-points (1pt = 2 half-points)
                const result = size * 2;
                return result;
            }
        }
        
        // Handle em values (em)
        if (fontSize.endsWith('em')) {
            const size = parseFloat(fontSize);
            if (!isNaN(size)) {
                // Convert em to half-points (1em = 12pt = 24 half-points)
                return Math.round(size * 12 * 2);
            }
        }
        
        // Handle rem values (rem)
        if (fontSize.endsWith('rem')) {
            const size = parseFloat(fontSize);
            if (!isNaN(size)) {
                // Convert rem to half-points (1rem = 16pt = 32 half-points)
                return Math.round(size * 16 * 2);
            }
        }
        
        // Handle percentage values (%)
        if (fontSize.endsWith('%')) {
            const size = parseFloat(fontSize);
            if (!isNaN(size)) {
                // Convert percentage to half-points (100% = 12pt = 24 half-points)
                return Math.round((size / 100) * 12 * 2);
            }
        }
        
        // Handle numeric values (assume pixels)
        const size = parseInt(fontSize);
        if (!isNaN(size)) {
            return Math.round(size * 0.75 * 2);
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
        
        // Get logo data from localStorage
        const allLogos = JSON.parse(localStorage.getItem('bytedraft_logos') || '{}');
        const logoData = allLogos[project.id];
        
        if (logoData) {
        }
        
        // Create title page with proper positioning
        const titlePageElements = [];
        
        // Add logo if available
        if (logoData) {
            try {
                // Convert base64 data URL to Uint8Array for browser environment
                const base64Data = logoData.split(',')[1];
                const binaryString = atob(base64Data);
                const bytes = new Uint8Array(binaryString.length);
                for (let i = 0; i < binaryString.length; i++) {
                    bytes[i] = binaryString.charCodeAt(i);
                }
                
                // Add logo at the top using a table for left alignment
                titlePageElements.push(
                    new this.docx.Table({
                        rows: [
                            new this.docx.TableRow({
                                children: [
                                    new this.docx.TableCell({
                                        children: [
                                            new this.docx.Paragraph({
                                                children: [
                                                    new this.docx.ImageRun({
                                                        data: bytes,
                                                        transformation: {
                                                            width: 200,
                                                            height: 120
                                                        }
                                                    })
                                                ]
                                            })
                                        ],
                                        width: { size: 30, type: this.docx.WidthType.PERCENTAGE },
                                        borders: {
                                            top: { style: this.docx.BorderStyle.NONE },
                                            bottom: { style: this.docx.BorderStyle.NONE },
                                            left: { style: this.docx.BorderStyle.NONE },
                                            right: { style: this.docx.BorderStyle.NONE }
                                        }
                                    }),
                                    new this.docx.TableCell({
                                        children: [new this.docx.Paragraph({ text: '' })],
                                        width: { size: 70, type: this.docx.WidthType.PERCENTAGE },
                                        borders: {
                                            top: { style: this.docx.BorderStyle.NONE },
                                            bottom: { style: this.docx.BorderStyle.NONE },
                                            left: { style: this.docx.BorderStyle.NONE },
                                            right: { style: this.docx.BorderStyle.NONE }
                                        }
                                    })
                                ]
                            })
                        ],
                        width: { size: 100, type: this.docx.WidthType.PERCENTAGE },
                        borders: {
                            top: { style: this.docx.BorderStyle.NONE },
                            bottom: { style: this.docx.BorderStyle.NONE },
                            left: { style: this.docx.BorderStyle.NONE },
                            right: { style: this.docx.BorderStyle.NONE }
                        }
                    })
                );
                
                // Add spacing after the logo table
                titlePageElements.push(
                    new this.docx.Paragraph({
                        text: '',
                        spacing: { before: 0, after: 800 }
                    })
                );
                
            } catch (error) {
                console.warn('Error processing logo:', error);
                // Fallback to original spacing if logo fails
                titlePageElements.push(
                    new this.docx.Paragraph({
                        text: '',
                        spacing: { before: 0, after: 1500 }
                    })
                );
            }
        } else {
            // 5 empty lines to position title 5 lines down from top (original spacing)
            titlePageElements.push(
                new this.docx.Paragraph({
                    text: '',
                    spacing: { before: 0, after: 1500 } // 5 lines worth of space (300 per line)
                })
            );
        }
            
        // Project title - 36pt font size
        titlePageElements.push(
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
            })
        );
            
        // Project description - 18pt font size
        titlePageElements.push(
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
            })
        );
            
        // Add spacing to push document info to bottom (13 lines up from bottom - moved up 5 lines)
        titlePageElements.push(
            new this.docx.Paragraph({
                text: '',
                spacing: { before: 0, after: 3900 } // Space to push info box to bottom (13 lines = 3900, reduced from 14 lines)
            })
        );
        
        // Add additional spacing to ensure title page fills properly
        titlePageElements.push(
            new this.docx.Paragraph({
                text: '',
                spacing: { before: 0, after: 2000 } // Additional spacing to fill the page
            })
        );
        
        // Document information box positioned 5 lines up from bottom
        titlePageElements.push(
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
            })
        );
        
        // Document info table using data from Edit Document Info
        titlePageElements.push(
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
        );
        
        return titlePageElements;
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
        const changelogData = allChangelog[project.id] || [];
        
        
        // Ensure changelogData is an array
        const changelogArray = Array.isArray(changelogData) ? changelogData : [];
        
        if (changelogArray.length === 0) {
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
        changelogArray.forEach(row => {
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

    createReferencesPage(project) {
        const all = JSON.parse(localStorage.getItem('bytedraft_references') || '{}');
        const refs = all[project.id] || [];
        if (refs.length === 0) return [];

        const elements = [
            new this.docx.Paragraph({ text: '', pageBreakBefore: true }),
            new this.docx.Paragraph({
                children: [
                    new this.docx.TextRun({
                        text: 'References',
                        size: 48,
                        font: 'Calibri',
                        bold: true,
                        color: '2563eb'
                    })
                ],
                spacing: { before: 200, after: 400 }
            })
        ];

        refs.forEach((ref, idx) => {
            const parts = [];
            if (ref.authors) parts.push(ref.authors);
            if (ref.year) parts.push(`(${ref.year})`);
            if (ref.title) parts.push(`${ref.title}.`);
            if (ref.source) parts.push(`${ref.source}.`);
            if (ref.url) parts.push(ref.url);
            const text = `[${idx + 1}] ${parts.join(' ')}`;

            elements.push(new this.docx.Paragraph({
                children: [new this.docx.TextRun({ text, font: 'Calibri', size: 22 })],
                spacing: { after: 160 },
                indent: { left: 360, hanging: 360 }
            }));
        });

        return elements;
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