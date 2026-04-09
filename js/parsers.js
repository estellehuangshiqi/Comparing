/**
 * Document parsers for DOCX, PDF, and Excel files.
 * Each parser converts a file into a unified DocContent structure.
 */

/**
 * @typedef {Object} RunInfo
 * @property {string} text
 * @property {Object} style - Run style (bold, italic, font, size, color, underline)
 * @property {string} [xmlFragment] - Original XML fragment for DOCX runs
 */

/**
 * @typedef {Object} ParagraphInfo
 * @property {string} text
 * @property {RunInfo[]} runs
 * @property {Object} style - Paragraph style info
 * @property {string} [xmlFragment] - Original XML for DOCX paragraphs
 */

/**
 * @typedef {Object} DocContent
 * @property {ParagraphInfo[]} paragraphs
 * @property {Array} tables
 * @property {Object} metadata
 */

const Parsers = {

    /**
     * Detect file type from extension
     * @param {File} file
     * @returns {string} 'docx' | 'pdf' | 'xlsx' | 'unknown'
     */
    detectType(file) {
        const name = file.name.toLowerCase();
        if (name.endsWith('.docx')) return 'docx';
        if (name.endsWith('.pdf')) return 'pdf';
        if (name.endsWith('.xlsx') || name.endsWith('.xls')) return 'xlsx';
        return 'unknown';
    },

    /**
     * Parse any supported document
     * @param {File} file
     * @param {Function} onProgress
     * @returns {Promise<DocContent>}
     */
    async parse(file, onProgress) {
        const type = this.detectType(file);
        if (type === 'unknown') {
            throw new Error(`不支持的文件格式: ${file.name}\n仅支持 .docx / .pdf / .xlsx / .xls 格式`);
        }
        switch (type) {
            case 'docx': return this.parseDocx(file, onProgress);
            case 'pdf': return this.parsePdf(file, onProgress);
            case 'xlsx': return this.parseXlsx(file, onProgress);
        }
    },

    // ==================== DOCX Parser ====================

    /**
     * Parse a DOCX file, preserving XML structure for later modification
     * @param {File} file
     * @param {Function} onProgress
     * @returns {Promise<DocContent>}
     */
    async parseDocx(file, onProgress) {
        onProgress && onProgress('正在解析 Word 文档...');
        const arrayBuffer = await file.arrayBuffer();
        let zip;
        try {
            zip = await JSZip.loadAsync(arrayBuffer);
        } catch (e) {
            throw new Error(`文件损坏或不是有效的 .docx 文件: ${e.message}`);
        }

        const docXml = zip.file('word/document.xml');
        if (!docXml) {
            throw new Error('无效的 .docx 文件：找不到 document.xml');
        }

        const xmlStr = await docXml.async('string');
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');

        // Check for parse errors
        const parseError = xmlDoc.querySelector('parsererror');
        if (parseError) {
            throw new Error('解析 .docx XML 结构时出错');
        }

        const nsResolver = (prefix) => {
            const ns = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            };
            return ns[prefix] || null;
        };

        // Get all w:body children - paragraphs and tables
        const body = xmlDoc.getElementsByTagNameNS(
            'http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'body'
        )[0];

        if (!body) {
            throw new Error('无效的 .docx 文件：找不到文档正文');
        }

        const paragraphs = [];
        const tables = [];
        const wNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

        for (const child of body.children) {
            if (child.localName === 'p') {
                paragraphs.push(this._parseDocxParagraph(child, wNS));
            } else if (child.localName === 'tbl') {
                tables.push(this._parseDocxTable(child, wNS));
            }
        }

        // Count words
        const totalChars = paragraphs.reduce((sum, p) => sum + p.text.length, 0);

        return {
            paragraphs,
            tables,
            metadata: {
                sourceFormat: 'docx',
                charCount: totalChars,
                pageCount: null,
                ocrUsed: false
            },
            _zip: zip,
            _xmlStr: xmlStr,
            _xmlDoc: xmlDoc
        };
    },

    /**
     * Parse a single DOCX paragraph element
     */
    _parseDocxParagraph(pElem, wNS) {
        const runs = [];
        let fullText = '';

        // Extract paragraph style
        const style = {};
        const pPr = pElem.getElementsByTagNameNS(wNS, 'pPr')[0];
        if (pPr) {
            const pStyle = pPr.getElementsByTagNameNS(wNS, 'pStyle')[0];
            if (pStyle) style.styleId = pStyle.getAttribute('w:val');

            const jc = pPr.getElementsByTagNameNS(wNS, 'jc')[0];
            if (jc) style.alignment = jc.getAttribute('w:val');

            const numPr = pPr.getElementsByTagNameNS(wNS, 'numPr')[0];
            if (numPr) style.isList = true;

            const outlineLvl = pPr.getElementsByTagNameNS(wNS, 'outlineLvl')[0];
            if (outlineLvl) style.outlineLevel = parseInt(outlineLvl.getAttribute('w:val'));
        }

        // Extract runs
        for (const child of pElem.children) {
            if (child.localName === 'r' && child.namespaceURI === wNS) {
                const run = this._parseDocxRun(child, wNS);
                if (run.text) {
                    runs.push(run);
                    fullText += run.text;
                }
            }
            // Also handle text in hyperlinks
            if (child.localName === 'hyperlink') {
                for (const hChild of child.children) {
                    if (hChild.localName === 'r' && hChild.namespaceURI === wNS) {
                        const run = this._parseDocxRun(hChild, wNS);
                        if (run.text) {
                            runs.push(run);
                            fullText += run.text;
                        }
                    }
                }
            }
        }

        // Serialize the paragraph element for later use
        const serializer = new XMLSerializer();
        const xmlFragment = serializer.serializeToString(pElem);

        return { text: fullText, runs, style, xmlFragment, _element: pElem };
    },

    /**
     * Parse a single DOCX run element
     */
    _parseDocxRun(rElem, wNS) {
        let text = '';
        const style = {};

        // Run properties
        const rPr = rElem.getElementsByTagNameNS(wNS, 'rPr')[0];
        if (rPr) {
            if (rPr.getElementsByTagNameNS(wNS, 'b')[0]) style.bold = true;
            if (rPr.getElementsByTagNameNS(wNS, 'i')[0]) style.italic = true;
            if (rPr.getElementsByTagNameNS(wNS, 'u')[0]) style.underline = true;

            const sz = rPr.getElementsByTagNameNS(wNS, 'sz')[0];
            if (sz) style.fontSize = parseInt(sz.getAttribute('w:val'));

            const rFonts = rPr.getElementsByTagNameNS(wNS, 'rFonts')[0];
            if (rFonts) {
                style.font = rFonts.getAttribute('w:ascii') ||
                             rFonts.getAttribute('w:eastAsia') ||
                             rFonts.getAttribute('w:hAnsi');
            }

            const color = rPr.getElementsByTagNameNS(wNS, 'color')[0];
            if (color) style.color = color.getAttribute('w:val');
        }

        // Text content - handle w:t, w:tab, w:br
        for (const child of rElem.children) {
            if (child.localName === 't' && child.namespaceURI === wNS) {
                text += child.textContent;
            } else if (child.localName === 'tab' && child.namespaceURI === wNS) {
                text += '\t';
            } else if (child.localName === 'br' && child.namespaceURI === wNS) {
                text += '\n';
            }
        }

        const serializer = new XMLSerializer();
        const rPrXml = rPr ? serializer.serializeToString(rPr) : '';

        return { text, style, _rPrXml: rPrXml, _element: rElem };
    },

    /**
     * Parse a DOCX table element
     */
    _parseDocxTable(tblElem, wNS) {
        const rows = [];
        const trElems = tblElem.getElementsByTagNameNS(wNS, 'tr');
        for (const tr of trElems) {
            const cells = [];
            const tcElems = tr.getElementsByTagNameNS(wNS, 'tc');
            for (const tc of tcElems) {
                let cellText = '';
                const pElems = tc.getElementsByTagNameNS(wNS, 'p');
                for (const p of pElems) {
                    const rElems = p.getElementsByTagNameNS(wNS, 'r');
                    for (const r of rElems) {
                        const tElems = r.getElementsByTagNameNS(wNS, 't');
                        for (const t of tElems) {
                            cellText += t.textContent;
                        }
                    }
                    cellText += '\n';
                }
                // Check for gridSpan (colspan)
                let colspan = 1;
                const tcPr = tc.getElementsByTagNameNS(wNS, 'tcPr')[0];
                if (tcPr) {
                    const gridSpan = tcPr.getElementsByTagNameNS(wNS, 'gridSpan')[0];
                    if (gridSpan) colspan = parseInt(gridSpan.getAttribute('w:val')) || 1;
                }
                cells.push({ text: cellText.trim(), colspan, rowspan: 1 });
            }
            rows.push({ cells });
        }
        return { rows };
    },

    // ==================== PDF Parser ====================

    /**
     * Parse a PDF file, with OCR fallback for scanned documents.
     * Extracts positional and font information so that the generated
     * revision document can preserve the original visual formatting.
     * @param {File} file
     * @param {Function} onProgress
     * @returns {Promise<DocContent>}
     */
    async parsePdf(file, onProgress) {
        onProgress && onProgress('正在解析 PDF 文档...');

        // Set pdf.js worker
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc =
                'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }

        const arrayBuffer = await file.arrayBuffer();
        let pdfDoc;
        try {
            pdfDoc = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        } catch (e) {
            throw new Error(`无法解析 PDF 文件: ${e.message}`);
        }

        const pageCount = pdfDoc.numPages;
        const paragraphs = [];
        let ocrUsed = false;
        let totalChars = 0;
        let pageWidth = 612;

        for (let i = 1; i <= pageCount; i++) {
            onProgress && onProgress(`正在提取第 ${i}/${pageCount} 页文字...`);
            const page = await pdfDoc.getPage(i);
            const viewport = page.getViewport({ scale: 1.0 });
            if (i === 1) pageWidth = viewport.width;

            const textContent = await page.getTextContent();

            // Quick scanned-page check
            const rawText = textContent.items.map(it => it.str).join('');
            if (rawText.replace(/\s/g, '').length < 50) {
                onProgress && onProgress(`正在识别第 ${i}/${pageCount} 页（OCR）...`);
                ocrUsed = true;

                let ocrText = '';
                try {
                    ocrText = await this._ocrPage(page, i, pageCount, onProgress);
                } catch (e) {
                    ocrText = `[第${i}页：OCR识别失败 - ${e.message}]`;
                }
                if (!ocrText || !ocrText.trim()) {
                    ocrText = `[第${i}页：该页图像质量过低，无法识别]`;
                }
                const ocrParas = ocrText.split(/\n\s*\n|\n/).map(s => s.trim()).filter(Boolean);
                for (const pText of ocrParas) {
                    paragraphs.push({
                        text: pText,
                        runs: [{ text: pText, style: {} }],
                        style: {}
                    });
                    totalChars += pText.length;
                }
                continue;
            }

            // Build styled lines from the text items, then group into paragraphs
            const lines = this._extractPdfLines(textContent, viewport);
            const pageParas = this._groupPdfLinesIntoParagraphs(lines, viewport);
            for (const para of pageParas) {
                paragraphs.push(para);
                totalChars += para.text.length;
            }
        }

        if (paragraphs.length === 0) {
            throw new Error('PDF 文档内容为空，无法提取任何文字内容');
        }

        return {
            paragraphs,
            tables: [],
            metadata: {
                sourceFormat: 'pdf',
                charCount: totalChars,
                pageCount,
                ocrUsed,
                pageWidth
            }
        };
    },

    /**
     * Extract lines with font info from a PDF page's text content.
     * Each returned line has: {y, x, maxFontSize, text, runs, centered}
     */
    _extractPdfLines(textContent, viewport) {
        const items = textContent.items || [];
        const styles = textContent.styles || {};
        if (items.length === 0) return [];

        // Normalize items with position, size, font, bold, italic
        const normalized = [];
        for (const item of items) {
            if (!item.str) continue;
            const tx = item.transform || [1, 0, 0, 1, 0, 0];
            const fontSize = Math.hypot(tx[2], tx[3]) || Math.hypot(tx[0], tx[1]) || 12;
            const x = tx[4];
            const y = tx[5];
            const styleInfo = styles[item.fontName] || {};
            const fontFamily = (styleInfo.fontFamily || '').replace(/^"|"$/g, '') || '';
            const combinedName = (fontFamily + ' ' + (item.fontName || '')).toLowerCase();
            const bold = /bold|heavy|black|semibold|demibold/.test(combinedName);
            const italic = /italic|oblique/.test(combinedName);
            normalized.push({
                str: item.str,
                x, y,
                width: item.width || 0,
                height: item.height || fontSize,
                fontSize,
                fontFamily: this._cleanFontFamily(fontFamily),
                bold, italic
            });
        }

        if (normalized.length === 0) return [];

        // Group items into lines by Y coordinate (tolerance depends on font size)
        const sorted = normalized.slice().sort((a, b) => {
            if (Math.abs(a.y - b.y) > 1) return b.y - a.y;
            return a.x - b.x;
        });

        const lineBuckets = [];
        for (const it of sorted) {
            const tol = Math.max(2, it.fontSize * 0.35);
            let attached = false;
            for (const bucket of lineBuckets) {
                if (Math.abs(bucket.y - it.y) <= tol) {
                    bucket.items.push(it);
                    // Weighted average y
                    bucket.y = (bucket.y * (bucket.items.length - 1) + it.y) / bucket.items.length;
                    attached = true;
                    break;
                }
            }
            if (!attached) lineBuckets.push({ y: it.y, items: [it] });
        }

        lineBuckets.sort((a, b) => b.y - a.y);

        const lines = [];
        for (const bucket of lineBuckets) {
            bucket.items.sort((a, b) => a.x - b.x);

            const runs = [];
            let lineText = '';
            let prev = null;
            let minX = Infinity, maxX = -Infinity, maxFontSize = 0;

            for (const it of bucket.items) {
                if (!it.str) continue;

                // Insert a space between items when there is a visible horizontal gap
                if (prev) {
                    const prevEnd = prev.x + prev.width;
                    const gap = it.x - prevEnd;
                    const spaceW = Math.max(prev.fontSize, it.fontSize) * 0.25;
                    const endsWithSpace = /\s$/.test(lineText);
                    const startsWithSpace = /^\s/.test(it.str);
                    if (gap > spaceW && !endsWithSpace && !startsWithSpace) {
                        this._appendToRuns(runs, ' ', prev);
                        lineText += ' ';
                    }
                }

                this._appendToRuns(runs, it.str, it);
                lineText += it.str;
                minX = Math.min(minX, it.x);
                maxX = Math.max(maxX, it.x + it.width);
                maxFontSize = Math.max(maxFontSize, it.fontSize);
                prev = it;
            }

            if (!lineText.trim()) continue;

            // Decide alignment by looking at left/right margins
            const leftMargin = minX;
            const rightMargin = viewport.width - maxX;
            const marginDiff = Math.abs(leftMargin - rightMargin);
            const centered = marginDiff < 20 && leftMargin > 40;

            lines.push({
                y: bucket.y,
                x: minX,
                right: maxX,
                maxFontSize,
                text: lineText,
                runs,
                centered
            });
        }

        return lines;
    },

    _cleanFontFamily(name) {
        if (!name) return '';
        // Strip subset prefixes like "ABCDEF+FontName"
        const m = name.match(/^[A-Z]{6}\+(.+)$/);
        if (m) name = m[1];
        // Strip style suffixes commonly included in font family names
        return name.replace(/[,\-](Bold|Italic|Oblique|Regular|Light|Medium|Semibold|Demibold|Heavy|Black)(Italic|Oblique)?$/i, '').trim();
    },

    _appendToRuns(runs, text, item) {
        const style = {
            fontSize: Math.max(1, Math.round(item.fontSize * 2)), // half-points
            font: item.fontFamily || '',
            bold: !!item.bold,
            italic: !!item.italic
        };
        const last = runs[runs.length - 1];
        if (last &&
            last.style.fontSize === style.fontSize &&
            last.style.font === style.font &&
            last.style.bold === style.bold &&
            last.style.italic === style.italic) {
            last.text += text;
        } else {
            runs.push({ text, style });
        }
    },

    /**
     * Group consecutive lines into paragraphs based on vertical spacing
     * and indentation changes.
     */
    _groupPdfLinesIntoParagraphs(lines, viewport) {
        const paragraphs = [];
        if (lines.length === 0) return paragraphs;

        let current = null;
        let prevLine = null;

        const isCJK = (ch) => {
            if (!ch) return false;
            const c = ch.charCodeAt(0);
            return (c >= 0x4E00 && c <= 0x9FFF) ||
                   (c >= 0x3400 && c <= 0x4DBF) ||
                   (c >= 0x3000 && c <= 0x303F) ||
                   (c >= 0xFF00 && c <= 0xFFEF);
        };

        const flush = () => {
            if (current && current.text.trim()) {
                paragraphs.push(current);
            }
            current = null;
        };

        for (const line of lines) {
            let newPara = false;
            if (!current) {
                newPara = true;
            } else if (prevLine) {
                const gap = prevLine.y - line.y;
                const expected = Math.max(prevLine.maxFontSize, line.maxFontSize) * 1.15;
                if (gap > expected * 1.55) newPara = true;
                // Indentation change suggests a new paragraph
                if (Math.abs(line.x - prevLine.x) > prevLine.maxFontSize * 1.2) newPara = true;
                // Font size change (heading boundary)
                if (Math.abs(line.maxFontSize - prevLine.maxFontSize) > 1.5) newPara = true;
                // Previous line ends with a sentence terminator
                const lastCh = prevLine.text.trim().slice(-1);
                if (/[。？！\.\?!]/.test(lastCh) && gap > expected * 0.9) newPara = true;
            }

            if (newPara) {
                flush();
                current = {
                    text: '',
                    runs: [],
                    style: {
                        alignment: line.centered ? 'center' : null,
                        fontSize: Math.max(1, Math.round(line.maxFontSize * 2))
                    }
                };
            }

            // Join line text to current paragraph.
            // For CJK-to-CJK boundaries we do not insert a space.
            if (current.text) {
                const lastCh = current.text.slice(-1);
                const firstCh = line.text.charAt(0);
                const joinWithSpace = !(isCJK(lastCh) || isCJK(firstCh));
                if (joinWithSpace && !/\s$/.test(current.text) && !/^\s/.test(line.text)) {
                    // Append a space to the last run
                    if (current.runs.length > 0) {
                        current.runs[current.runs.length - 1].text += ' ';
                    }
                    current.text += ' ';
                }
            }

            current.text += line.text;
            for (const run of line.runs) {
                const last = current.runs[current.runs.length - 1];
                if (last &&
                    last.style.fontSize === run.style.fontSize &&
                    last.style.font === run.style.font &&
                    last.style.bold === run.style.bold &&
                    last.style.italic === run.style.italic) {
                    last.text += run.text;
                } else {
                    current.runs.push({ text: run.text, style: { ...run.style } });
                }
            }
            prevLine = line;
        }

        flush();
        return paragraphs;
    },

    /**
     * OCR a single PDF page using tesseract.js
     */
    async _ocrPage(page, pageNum, totalPages, onProgress) {
        // Render page to canvas
        const viewport = page.getViewport({ scale: 2.0 }); // Higher scale for better OCR
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');

        await page.render({ canvasContext: ctx, viewport }).promise;

        // Use tesseract.js for OCR
        const worker = await Tesseract.createWorker('chi_sim+eng', 1, {
            logger: (m) => {
                if (m.status === 'recognizing text' && onProgress) {
                    const pageProgress = Math.round(m.progress * 100);
                    onProgress(`正在识别第 ${pageNum}/${totalPages} 页... ${pageProgress}%`);
                }
            }
        });

        const { data } = await worker.recognize(canvas);
        await worker.terminate();

        return data.text;
    },

    // ==================== Excel Parser ====================

    /**
     * Parse an Excel file
     * @param {File} file
     * @param {Function} onProgress
     * @returns {Promise<DocContent>}
     */
    async parseXlsx(file, onProgress) {
        onProgress && onProgress('正在解析 Excel 文档...');

        const arrayBuffer = await file.arrayBuffer();
        let workbook;
        try {
            workbook = XLSX.read(arrayBuffer, { type: 'array' });
        } catch (e) {
            throw new Error(`无法解析 Excel 文件: ${e.message}`);
        }

        const paragraphs = [];
        const tables = [];
        let totalChars = 0;

        for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];

            // Add sheet name as heading
            paragraphs.push({
                text: `[工作表: ${sheetName}]`,
                runs: [{ text: `[工作表: ${sheetName}]`, style: { bold: true } }],
                style: { styleId: 'Heading2' }
            });

            // Convert sheet to array of arrays
            const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            if (data.length === 0) continue;

            // Build table structure
            const tableRows = [];
            for (const row of data) {
                const cells = [];
                for (const cell of row) {
                    const cellText = String(cell);
                    cells.push({ text: cellText, colspan: 1, rowspan: 1 });
                    totalChars += cellText.length;
                }
                tableRows.push({ cells });
            }
            tables.push({ rows: tableRows, sheetName });

            // Also add as paragraphs for text comparison
            for (const row of data) {
                const rowText = row.map(c => String(c)).filter(c => c).join('\t');
                if (rowText.trim()) {
                    paragraphs.push({
                        text: rowText,
                        runs: [{ text: rowText, style: {} }],
                        style: {}
                    });
                }
            }
        }

        return {
            paragraphs,
            tables,
            metadata: {
                sourceFormat: 'xlsx',
                charCount: totalChars,
                pageCount: workbook.SheetNames.length,
                ocrUsed: false
            }
        };
    }
};
