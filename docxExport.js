// DOCX Export Module (Classic Script Version)
// All functions are now global and can be used from a classic <script> tag.

function escapeXml(text) {
    if (typeof text !== 'string') {
        text = (text === undefined || text === null) ? '' : String(text);
    }
    text = text.replace(/&nbsp;/g, ' ').replace(/\u00A0/g, ' ');
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function createDocxHeader(headerContent) {
    const content = (headerContent || '').replace(/\{\{page\}\}/g, '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>${escapeXml(content)}</w:t></w:r></w:p></w:hdr>`;
}

function createDocxFooter(footerContent) {
    let parts = (footerContent || '').split(/(\{\{page\}\})/g);
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>';
    parts.forEach(part => {
        if (part === '{{page}}') {
            xml += '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';
        } else if (part) {
            xml += `<w:r><w:t>${escapeXml(part)}</w:t></w:r>`;
        }
    });
    xml += '</w:p></w:ftr>';
    return xml;
}

function createDocxStyles() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n    <w:docDefaults>\n        <w:rPrDefault>\n            <w:rPr>\n                <w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>\n                <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>\n            </w:rPr>\n        </w:rPrDefault>\n    </w:docDefaults>\n    <!-- Heading 1 Style -->\n    <w:style w:type="paragraph" w:styleId="Heading1">\n        <w:name w:val="Heading 1"/>\n        <w:qFormat/>\n        <w:pPr>\n            <w:spacing w:before="240"/>\n            <w:outlineLvl w:val="0"/>\n            <w:ind w:left="0"/>\n            <w:numPr>\n                <w:numId w:val="1"/>\n                <w:ilvl w:val="0"/>\n            </w:numPr>\n        </w:pPr>\n        <w:rPr>\n            <w:b/>\n            <w:sz w:val="32"/>\n            <w:szCs w:val="32"/>\n            <w:color w:val="2563EB"/>\n        </w:rPr>\n    </w:style>\n    <!-- Heading 2 Style -->\n    <w:style w:type="paragraph" w:styleId="Heading2">\n        <w:name w:val="Heading 2"/>\n        <w:qFormat/>\n        <w:pPr>\n            <w:spacing w:before="240"/>\n            <w:outlineLvl w:val="1"/>\n            <w:ind w:left="0"/>\n            <w:numPr>\n                <w:numId w:val="1"/>\n                <w:ilvl w:val="1"/>\n            </w:numPr>\n        </w:pPr>\n        <w:rPr>\n            <w:b/>\n            <w:sz w:val="28"/>\n            <w:szCs w:val="28"/>\n            <w:color w:val="2563EB"/>\n        </w:rPr>\n    </w:style>\n    <!-- Heading 3 Style -->\n    <w:style w:type="paragraph" w:styleId="Heading3">\n        <w:name w:val="Heading 3"/>\n        <w:qFormat/>\n        <w:pPr>\n            <w:spacing w:before="240"/>\n            <w:outlineLvl w:val="2"/>\n            <w:ind w:left="0"/>\n            <w:numPr>\n                <w:numId w:val="1"/>\n                <w:ilvl w:val="2"/>\n            </w:numPr>\n        </w:pPr>\n        <w:rPr>\n            <w:b/>\n            <w:sz w:val="24"/>\n            <w:szCs w:val="24"/>\n            <w:color w:val="2563EB"/>\n        </w:rPr>\n    </w:style>\n    <!-- TOCHeading Style -->\n    <w:style w:type="paragraph" w:styleId="TOCHeading">\n      <w:name w:val="TOC Heading"/>\n      <w:next w:val="Normal"/>\n      <w:uiPriority w:val="39"/>\n      <w:qFormat/>\n      <w:pPr>\n        <w:keepNext/>\n        <w:spacing w:before="240" w:after="240"/>\n        <w:jc w:val="center"/>\n      </w:pPr>\n      <w:rPr>\n        <w:b/>\n        <w:color w:val="2563EB"/>\n        <w:sz w:val="44"/>\n        <w:szCs w:val="44"/>\n      </w:rPr>\n    </w:style>\n</w:styles>`;
}

function createDocxNumbering() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n    <w:abstractNum w:abstractNumId="0">\n        <w:nsid w:val="00000001"/>\n        <w:multiLevelType w:val="hybridMultilevel"/>\n        <w:lvl w:ilvl="0" w:tplc="04090001">\n            <w:start w:val="1"/>\n            <w:numFmt w:val="decimal"/>\n            <w:lvlText w:val="%1."/>\n            <w:lvlJc w:val="left"/>\n            <w:pPr>\n                <w:ind w:left="720" w:hanging="360"/>\n            </w:pPr>\n            <w:rPr>\n                <w:rFonts w:hint="default"/>\n            </w:rPr>\n        </w:lvl>\n        <w:lvl w:ilvl="1" w:tplc="04090003">\n            <w:start w:val="1"/>\n            <w:numFmt w:val="decimal"/>\n            <w:lvlText w:val="%1.%2."/>\n            <w:lvlJc w:val="left"/>\n            <w:pPr>\n                <w:ind w:left="1440" w:hanging="360"/>\n            </w:pPr>\n            <w:rPr>\n                <w:rFonts w:hint="default"/>\n            </w:rPr>\n        </w:lvl>\n        <w:lvl w:ilvl="2" w:tplc="04090005">\n            <w:start w:val="1"/>\n            <w:numFmt w:val="decimal"/>\n            <w:lvlText w:val="%1.%2.%3."/>\n            <w:lvlJc w:val="left"/>\n            <w:pPr>\n                <w:ind w:left="2160" w:hanging="360"/>\n            </w:pPr>\n            <w:rPr>\n                <w:rFonts w:hint="default"/>\n            </w:rPr>\n        </w:lvl>\n    </w:abstractNum>\n    <w:num w:numId="1">\n        <w:abstractNumId w:val="0"/>\n    </w:num>\n</w:numbering>`;
}

function createDocxSettings() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n    <w:zoom w:percent="100"/>\n    <w:defaultTabStop w:val="720"/>\n    <w:characterSpacingControl w:val="doNotCompress"/>\n    <w:compat/>\n    <w:rsids>\n        <w:rsidRoot w:val="00000000"/>\n    </w:rsids>\n    <w:themeFontLang w:val="en-US" w:eastAsia="en-US"/>\n    <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>\n</w:settings>`;
}

function createDocxContentTypes() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n    <Default Extension="xml" ContentType="application/xml"/>\n    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>\n    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>\n    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>\n    <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>\n    <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>\n</Types>`;
}

function createDocxRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n    <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>\n    <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>\n</Relationships>`;
}

function createDocxDocumentRels() {
    let rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n`;
    rels += '    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
    rels += '\n    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>';
    rels += '\n    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';
    rels += '\n    <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>';
    rels += '\n    <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>';
    if (typeof window !== 'undefined' && window._docxHyperlinks) {
        window._docxHyperlinks.forEach(link => {
            rels += `\n    <Relationship Id="${link.relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${link.href}" TargetMode="External"/>`;
        });
    }
    rels += '\n</Relationships>';
    return rels;
}

function generateDocInfoTableDOCX(info) {
    const rows = [
        ['Document Title:', info.title],
        ['Author:', info.author],
        ['Document Owner:', info.docOwner],
        ['Process Owner:', info.procOwner],
        ['Version No:', info.version],
        ['Effective Date:', info.effDate],
        ['Last Reviewed Date:', info.lastRev],
        ['Next Reviewed Date:', info.nextRev],
        ['Document Link:', info.link],
    ];
    let xml = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:jc w:val="center"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid>';
    rows.forEach(([label, value]) => {
        xml += `\n<w:tr>\n    <w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="002060"/></w:tcPr><w:p><w:r><w:rPr><w:color w:val="FFFFFF"/><w:b/></w:rPr><w:t>${escapeXml(label)}</w:t></w:r></w:p></w:tc>\n    <w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${escapeXml(value || '')}</w:t></w:r></w:p></w:tc>\n</w:tr>`;
    });
    xml += '</w:tbl>';
    return xml;
}

function processSubsectionsRecursively(subsections, htmlToDocxXml) {
    let xml = '';
    subsections.forEach((subsection) => {
        xml += `\n            <w:p>\n                <w:pPr>\n                    <w:pStyle w:val="Heading3"/>\n                </w:pPr>\n                <w:r>\n                    <w:rPr>\n                        <w:b/>\n                        <w:color w:val="2563EB"/>\n                        <w:sz w:val="24"/>\n                    </w:rPr>\n                    <w:t>${escapeXml(subsection.title)}</w:t>\n                </w:r>\n            </w:p>`;
        if (subsection.content) {
            const subsectionXmlContent = htmlToDocxXml(subsection.content);
            if (subsectionXmlContent) {
                xml += subsectionXmlContent;
            }
        }
        if (subsection.subsections && subsection.subsections.length > 0) {
            xml += processSubsectionsRecursively(subsection.subsections, htmlToDocxXml);
        }
    });
    return xml;
}

function htmlToDocxXml(html, htmlInlineToDocx) {
    if (!html) return '';
    html = html.replace(/<br\s*\/?>(?!\|\|\|)/gi, '</br>');
    html = html.replace(/<p/gi, '|||<p');
    html = html.replace(/<li/gi, '|||<li');
    html = html.replace(/<\/p>/gi, '</p>|||');
    html = html.replace(/<\/li>/gi, '</li>|||');
    html = html.replace(/<ul>/gi, '');
    html = html.replace(/<\/ul>/gi, '');
    html = html.replace(/<ol>/gi, '');
    html = html.replace(/<\/ol>/gi, '');
    html = html.replace(/<br\s*\/?>(?!\|\|\|)/gi, '</br>|||');
    const blocks = html.split('|||').map(s => s.trim()).filter(Boolean);
    let xml = '';
    let hyperlinks = [];
    blocks.forEach(block => {
        if (block.startsWith('<li')) {
            let text = block.replace(/<li[^>]*>/gi, '').replace(/<\/li>/gi, '');
            xml += `<w:p><w:r><w:t>\u2022 </w:t></w:r>${htmlInlineToDocx(text, hyperlinks)}</w:p>`;
        } else if (block.startsWith('<p')) {
            let align = '';
            const alignMatch = block.match(/text-align\s*:\s*(center|right|left)/i);
            if (alignMatch) {
                align = alignMatch[1].toLowerCase();
            }
            let fontSize = '';
            let fontFamily = '';
            const fontSizeMatch = block.match(/font-size\s*:\s*([0-9.]+)pt/i);
            if (fontSizeMatch) fontSize = fontSizeMatch[1];
            const fontFamilyMatch = block.match(/font-family\s*:\s*([^;"']+)/i);
            if (fontFamilyMatch) fontFamily = fontFamilyMatch[1];
            let text = block.replace(/<p[^>]*>/gi, '').replace(/<\/p>/gi, '');
            xml += `<w:p>`;
            if (align || fontSize || fontFamily) {
                xml += `<w:pPr>`;
                if (align) xml += `<w:jc w:val=\"${align}\"/>`;
                xml += `</w:pPr>`;
            }
            xml += htmlInlineToDocx(text, hyperlinks, fontSize, fontFamily) + `</w:p>`;
        } else if (block.startsWith('</br>')) {
            xml += `<w:p></w:p>`;
        } else if (block) {
            xml += `<w:p>${htmlInlineToDocx(block, hyperlinks)}</w:p>`;
        }
    });
    if (typeof window !== 'undefined') {
        if (!window._docxHyperlinks) window._docxHyperlinks = [];
        window._docxHyperlinks.push(...hyperlinks);
    }
    return xml;
}

function htmlInlineToDocx(html, hyperlinks, parentFontSize, parentFontFamily) {
    if (!html) return '';
    let runs = [];
    let tagStack = [];
    let buffer = '';
    let i = 0;
    while (i < html.length) {
        if (html[i] === '<') {
            if (buffer) {
                runs.push({text: buffer, tags: [...tagStack]});
                buffer = '';
            }
            const closeIdx = html.indexOf('>', i);
            if (closeIdx === -1) break;
            const tag = html.substring(i, closeIdx + 1);
            const tagName = tag.match(/<\/?([a-zA-Z]+)/)?.[1]?.toLowerCase();
            const isClose = tag.startsWith('</');
            if (['b','strong','i','em','u'].includes(tagName)) {
                if (isClose) {
                    tagStack.pop();
                } else {
                    tagStack.push(tagName);
                }
            } else if (tagName === 'span') {
                if (isClose) {
                    tagStack.pop();
                } else {
                    const fontSizeMatch = tag.match(/font-size:\s*([0-9.]+)pt/i);
                    const fontFamilyMatch = tag.match(/font-family:\s*([^;"']+)/i);
                    tagStack.push({type:'span', fontSize: fontSizeMatch ? fontSizeMatch[1] : undefined, fontFamily: fontFamilyMatch ? fontFamilyMatch[1] : undefined});
                }
            } else if (tagName === 'a') {
                if (isClose) {
                    tagStack.pop();
                } else {
                    const hrefMatch = tag.match(/href=["']([^"']+)["']/i);
                    const href = hrefMatch ? hrefMatch[1] : '';
                    const relId = `rId${(hyperlinks.length + 1)}`;
                    tagStack.push({type:'a', href, relId});
                    hyperlinks.push({href, relId});
                }
            }
            i = closeIdx + 1;
        } else {
            buffer += html[i];
            i++;
        }
    }
    if (buffer) {
        runs.push({text: buffer, tags: [...tagStack]});
    }
    let xml = '';
    runs.forEach(run => {
        let rPr = '';
        let text = run.text;
        if (!text) return;
        if (run.tags.includes('b') || run.tags.includes('strong')) rPr += '<w:b/>';
        if (run.tags.includes('i') || run.tags.includes('em')) rPr += '<w:i/>';
        if (run.tags.includes('u')) rPr += '<w:u w:val="single"/>';
        let fontSize = parentFontSize;
        let fontFamily = parentFontFamily;
        const spanTag = run.tags.find(t => typeof t === 'object' && t.type === 'span');
        if (spanTag) {
            if (spanTag.fontSize) fontSize = spanTag.fontSize;
            if (spanTag.fontFamily) fontFamily = spanTag.fontFamily;
        }
        if (fontSize) rPr += `<w:sz w:val="${Math.round(parseFloat(fontSize)*2)}"/>`;
        if (fontFamily) rPr += `<w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}"/>`;
        rPr += '<w:color w:val="000000"/>';
        const linkTag = run.tags.find(t => typeof t === 'object' && t.type === 'a');
        if (linkTag) {
            xml += `<w:hyperlink r:id="${linkTag.relId}" w:history="1"><w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:hyperlink>`;
        } else {
            xml += `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
        }
    });
    return xml;
}

function createDocxDocument(currentProject, customChangelog, versionHistory, getDocumentInfo, htmlToDocxXml, processSubsectionsRecursively, generateDocInfoTableDOCX, escapeXml) {
    try {
        if (!currentProject) {
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>ERROR: No project data</w:t></w:r></w:p></w:body></w:document>`;
        }
        if (!currentProject.sections || !Array.isArray(currentProject.sections)) {
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>ERROR: No sections</w:t></w:r></w:p></w:body></w:document>`;
        }
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n    <w:body>`;
        // Title page
        xml += `\n            <w:p>\n                <w:pPr>\n                    <w:jc w:val="center"/>\n                    <w:spacing w:after="480" w:before="480"/>\n                </w:pPr>\n                <w:r>\n                    <w:rPr>\n                        <w:b/>\n                        <w:sz w:val="72"/>\n                    </w:rPr>\n                    <w:t>${escapeXml(currentProject.name)}</w:t>\n                </w:r>\n            </w:p>`;
        if (currentProject.description) {
            xml += `\n            <w:p>\n                <w:pPr>\n                    <w:jc w:val="center"/>\n                    <w:spacing w:after="240"/>\n                </w:pPr>\n                <w:r>\n                    <w:rPr>\n                        <w:sz w:val="36"/>\n                    </w:rPr>\n                    <w:t>${escapeXml(currentProject.description)}</w:t>\n                </w:r>\n            </w:p>`;
        }
        // Document Info Table (29 lines before)
        const info = getDocumentInfo(currentProject.id);
        for (let i = 0; i < 29; i++) xml += '<w:p></w:p>';
        xml += generateDocInfoTableDOCX(info);
        for (let i = 0; i < 10; i++) xml += '<w:p></w:p>';
        // Document Changelog Page
        xml += `\n            <w:p>\n                <w:pPr>\n                    <w:spacing w:after="840" w:before="840"/>\n                </w:pPr>\n                <w:r>\n                    <w:rPr>\n                        <w:b/>\n                        <w:sz w:val="32"/>\n                    </w:rPr>\n                    <w:t>Document Changelog</w:t>\n                </w:r>\n            </w:p>`;
        // Changelog Table or Version History
        const changelogData = customChangelog[currentProject.id];
        let changelogRows = [];
        if (changelogData) {
            try {
                changelogRows = JSON.parse(changelogData);
            } catch (e) {}
        }
        if (changelogRows && changelogRows.length > 0) {
            // Changelog Table
            xml += `\n            <w:tbl>\n                <w:tblPr>\n                    <w:tblStyle w:val="TableGrid"/>\n                    <w:tblW w:w="0" w:type="auto"/>\n                    <w:tblBorders>\n                        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                    </w:tblBorders>\n                </w:tblPr>\n                <w:tblGrid>\n                    <w:gridCol w:w="2000"/>\n                    <w:gridCol w:w="2000"/>\n                    <w:gridCol w:w="2000"/>\n                    <w:gridCol w:w="2000"/>\n                    <w:gridCol w:w="2400"/>\n                    <w:gridCol w:w="4000"/>\n                </w:tblGrid>`;
            const headers = ['Version Number', 'Approved Date', 'Author', 'Reviewer', 'Approver for Change', 'Description'];
            xml += `\n                <w:tr>\n                    <w:trPr>\n                        <w:trHeight w:val="400"/>\n                    </w:trPr>`;
            headers.forEach(header => {
                xml += `\n                    <w:tc>\n                        <w:tcPr>\n                            <w:tcW w:w="2000" w:type="dxa"/>\n                            <w:shd w:val="clear" w:color="auto" w:fill="2563EB"/>\n                        </w:tcPr>\n                        <w:p>\n                            <w:pPr>\n                                <w:jc w:val="center"/>\n                            </w:pPr>\n                            <w:r>\n                                <w:rPr>\n                                    <w:b/>\n                                    <w:color w:val="FFFFFF"/>\n                                    <w:sz w:val="20"/>\n                                </w:rPr>\n                                <w:t>${escapeXml(header)}</w:t>\n                            </w:r>\n                        </w:p>\n                    </w:tc>`;
            });
            xml += `\n                </w:tr>`;
            changelogRows.forEach(row => {
                const rowData = [
                    row.version || '',
                    row.date || '',
                    row.author || '',
                    row.reviewer || '',
                    row.approver || '',
                    row.desc || ''
                ];
                xml += `\n                <w:tr>\n                    <w:trPr>\n                        <w:trHeight w:val="400"/>\n                    </w:trPr>`;
                rowData.forEach((cell, index) => {
                    const colWidth = index === 4 ? 2400 : index === 5 ? 4000 : 2000;
                    xml += `\n                    <w:tc>\n                        <w:tcPr>\n                            <w:tcW w:w="${colWidth}" w:type="dxa"/>\n                        </w:tcPr>\n                        <w:p>\n                            <w:pPr>\n                                <w:spacing w:after="0"/>\n                            </w:pPr>\n                            <w:r>\n                                <w:rPr>\n                                    <w:sz w:val="20"/>\n                                </w:rPr>\n                                <w:t>${escapeXml(cell)}</w:t>\n                            </w:r>\n                        </w:p>\n                    </w:tc>`;
                });
                xml += `\n                </w:tr>`;
            });
            xml += `\n            </w:tbl>`;
        } else {
            // Fallback to version history if no custom changelog
            const projectVersions = versionHistory.filter(v => v.projectId === currentProject.id).slice(-5);
            if (projectVersions.length > 0) {
                xml += `\n            <w:tbl>\n                <w:tblPr>\n                    <w:tblStyle w:val="TableGrid"/>\n                    <w:tblW w:w="0" w:type="auto"/>\n                    <w:tblBorders>\n                        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n                    </w:tblBorders>\n                </w:tblPr>\n                <w:tblGrid>\n                    <w:gridCol w:w="2000"/>\n                    <w:gridCol w:w="6000"/>\n                    <w:gridCol w:w="2000"/>\n                </w:tblGrid>\n                <w:tr>\n                    <w:trPr>\n                        <w:trHeight w:val="400"/>\n                    </w:trPr>\n                    <w:tc>\n                        <w:tcPr>\n                            <w:tcW w:w="2000" w:type="dxa"/>\n                            <w:shd w:val="clear" w:color="auto" w:fill="2563EB"/>\n                        </w:tcPr>\n                        <w:p>\n                            <w:pPr>\n                                <w:jc w:val="center"/>\n                            </w:pPr>\n                            <w:r>\n                                <w:rPr>\n                                    <w:b/>\n                                    <w:color w:val="FFFFFF"/>\n                                    <w:sz w:val="20"/>\n                                </w:rPr>\n                                <w:t>Version</w:t>\n                            </w:r>\n                        </w:p>\n                    </w:tc>\n                    <w:tc>\n                        <w:tcPr>\n                            <w:tcW w:w="6000" w:type="dxa"/>\n                            <w:shd w:val="clear" w:color="auto" w:fill="2563EB"/>\n                        </w:tcPr>\n                        <w:p>\n                            <w:pPr>\n                                <w:jc w:val="center"/>\n                            </w:pPr>\n                            <w:r>\n                                <w:rPr>\n                                    <w:b/>\n                                    <w:color w:val="FFFFFF"/>\n                                    <w:sz w:val="20"/>\n                                </w:rPr>\n                                <w:t>Description</w:t>\n                            </w:r>\n                        </w:p>\n                    </w:tc>\n                    <w:tc>\n                        <w:tcPr>\n                            <w:tcW w:w="2000" w:type="dxa"/>\n                            <w:shd w:val="clear" w:color="auto" w:fill="2563EB"/>\n                        </w:tcPr>\n                        <w:p>\n                            <w:pPr>\n                                <w:jc w:val="center"/>\n                            </w:pPr>\n                            <w:r>\n                                <w:rPr>\n                                    <w:b/>\n                                    <w:color w:val="FFFFFF"/>\n                                    <w:sz w:val="20"/>\n                                </w:rPr>\n                                <w:t>Date</w:t>\n                            </w:r>\n                        </w:p>\n                    </w:tc>\n                </w:tr>`;
                projectVersions.forEach((version, index) => {
                    xml += `\n                <w:tr>\n                    <w:trPr>\n                        <w:trHeight w:val="400"/>\n                    </w:trPr>\n                    <w:tc>\n                        <w:tcPr>\n                            <w:tcW w:w="2000" w:type="dxa"/>\n                        </w:tcPr>\n                        <w:p>\n                            <w:pPr>\n                                <w:spacing w:after="0"/>\n                            </w:pPr>\n                            <w:r>\n                                <w:rPr>\n                                    <w:sz w:val="20"/>\n                                </w:rPr>\n                                <w:t>${escapeXml((projectVersions.length - index).toString())}</w:t>\n                            </w:r>\n                        </w:p>\n                    </w:tc>\n                    <w:tc>\n                        <w:tcPr>\n                            <w:tcW w:w="6000" w:type="dxa"/>\n                        </w:tcPr>\n                        <w:p>\n                            <w:pPr>\n                                <w:spacing w:after="0"/>\n                            </w:pPr>\n                            <w:r>\n                                <w:rPr>\n                                    <w:sz w:val="20"/>\n                                </w:rPr>\n                                <w:t>${escapeXml(version.description)}</w:t>\n                            </w:r>\n                        </w:p>\n                    </w:tc>\n                    <w:tc>\n                        <w:tcPr>\n                            <w:tcW w:w="2000" w:type="dxa"/>\n                        </w:tcPr>\n                        <w:p>\n                            <w:pPr>\n                                <w:spacing w:after="0"/>\n                            </w:pPr>\n                            <w:r>\n                                <w:rPr>\n                                    <w:sz w:val="20"/>\n                                </w:rPr>\n                                <w:t>${escapeXml(new Date(version.timestamp).toLocaleDateString())}</w:t>\n                            </w:r>\n                        </w:p>\n                    </w:tc>\n                </w:tr>`;
                });
                xml += `\n            </w:tbl>`;
            } else {
                xml += `\n            <w:p>\n                <w:r>\n                    <w:t>No version history available.</w:t>\n                </w:r>\n            </w:p>`;
            }
        }
        // Page break
        xml += `\n            <w:p>\n                <w:r>\n                    <w:br w:type="page"/>\n                </w:r>\n            </w:p>`;
        // Table of Contents
        xml += `\n            <w:p>\n                <w:pPr>\n                    <w:pStyle w:val="TOCHeading"/>\n                </w:pPr>\n                <w:r>\n                    <w:rPr>\n                        <w:b/>\n                        <w:color w:val="2563EB"/>\n                        <w:sz w:val="44"/>\n                    </w:rPr>\n                    <w:t>Table of Contents</w:t>\n                </w:r>\n            </w:p>\n            <w:p>\n                <w:r>\n                    <w:fldChar w:fldCharType="begin"/>\n                </w:r>\n                <w:r>\n                    <w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>\n                </w:r>\n                <w:r>\n                    <w:fldChar w:fldCharType="separate"/>\n                </w:r>\n                <w:r>\n                    <w:t>Click here to update the table of contents</w:t>\n                </w:r>\n                <w:r>\n                    <w:fldChar w:fldCharType="end"/>\n                </w:r>\n            </w:p>`;
        // Page break
        xml += `\n            <w:p>\n                <w:r>\n                    <w:br w:type="page"/>\n                </w:r>\n            </w:p>`;
        // Main content
        currentProject.sections.forEach((section, index) => {
            xml += `\n            <w:p>\n                <w:pPr>\n                    <w:pStyle w:val="Heading1"/>\n                </w:pPr>\n                <w:r>\n                    <w:rPr>\n                        <w:b/>\n                        <w:color w:val="2563EB"/>\n                        <w:sz w:val="32"/>\n                    </w:rPr>\n                    <w:t>${escapeXml(section.title)}</w:t>\n                </w:r>\n            </w:p>`;
            if (section.content) {
                const xmlContent = htmlToDocxXml(section.content);
                if (xmlContent) {
                    xml += xmlContent;
                }
            }
            if (section.subsections && section.subsections.length > 0) {
                section.subsections.forEach((sub, subIndex) => {
                    xml += `\n            <w:p>\n                <w:pPr>\n                    <w:pStyle w:val="Heading2"/>\n                </w:pPr>\n                <w:r>\n                    <w:rPr>\n                        <w:b/>\n                        <w:color w:val="2563EB"/>\n                        <w:sz w:val="28"/>\n                    </w:rPr>\n                    <w:t>${escapeXml(sub.title)}</w:t>\n                </w:r>\n            </w:p>`;
                    if (sub.content) {
                        const subXmlContent = htmlToDocxXml(sub.content);
                        if (subXmlContent) {
                            xml += subXmlContent;
                        }
                    }
                    if (sub.subsections && sub.subsections.length > 0) {
                        xml += processSubsectionsRecursively(sub.subsections, htmlToDocxXml);
                    }
                });
            }
            if (index < currentProject.sections.length - 1) {
                xml += `\n            <w:p>\n                <w:pPr>\n                    <w:spacing w:after="240"/>\n                </w:pPr>\n            </w:p>`;
            }
        });
        // Final sectPr and closing tags
        xml += `\n        <w:sectPr>\n            <w:pgSz w:w="12240" w:h="15840"/>\n            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>\n            <w:headerReference w:type="default" r:id="rIdHeader1"/>\n            <w:footerReference w:type="default" r:id="rIdFooter1"/>\n        </w:sectPr>\n    </w:body>\n</w:document>`;
        return xml;
    } catch (error) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>ERROR: Failed to create DOCX document</w:t></w:r></w:p></w:body></w:document>`;
    }
} 