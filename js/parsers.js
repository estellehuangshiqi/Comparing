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
     *
     * Layout-aware extraction: groups text items into lines using y-coordinates,
     * then merges lines into paragraphs based on vertical gaps. Soft line wraps
     * (including hyphenation) are healed so the same content extracted from two
     * PDFs produces matching paragraphs even if line breaks differ.
     *
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

        for (let i = 1; i <= pageCount; i++) {
            onProgress && onProgress(`正在提取第 ${i}/${pageCount} 页文字...`);
            const page = await pdfDoc.getPage(i);
            const textContent = await page.getTextContent();

            // Layout-aware paragraph extraction
            let pageParas = this._extractPdfParagraphs(textContent);

            // Detect scanned page (too few extractable characters)
            const totalText = pageParas.map(p => p.text).join('');
            if (totalText.replace(/\s/g, '').length < 50) {
                onProgress && onProgress(`正在识别第 ${i}/${pageCount} 页（OCR）...`);
                ocrUsed = true;

                try {
                    const ocrText = await this._ocrPage(page, i, pageCount, onProgress);
                    if (ocrText && ocrText.trim().length > 0) {
                        pageParas = this._textToParagraphs(ocrText);
                    } else {
                        pageParas = [{
                            text: `[第${i}页：该页图像质量过低，无法识别]`,
                            runs: [{ text: `[第${i}页：该页图像质量过低，无法识别]`, style: {} }],
                            style: {}
                        }];
                    }
                } catch (e) {
                    pageParas = [{
                        text: `[第${i}页：OCR识别失败 - ${e.message}]`,
                        runs: [{ text: `[第${i}页：OCR识别失败 - ${e.message}]`, style: {} }],
                        style: {}
                    }];
                }
            }

            for (const p of pageParas) {
                if (p.text && p.text.trim()) {
                    if (!p.runs || p.runs.length === 0) {
                        p.runs = [{ text: p.text, style: p.style || {} }];
                    }
                    p.style = p.style || {};
                    p.style.pageNumber = i;
                    paragraphs.push(p);
                    totalChars += p.text.length;
                }
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
                ocrUsed
            }
        };
    },

    /**
     * Extract paragraphs from a PDF page's text content using layout info.
     * Items are first grouped into lines by y-coordinate, then lines into
     * paragraphs by larger vertical gaps. Soft wraps and hyphenation are
     * healed when merging lines into a paragraph.
     */
    _extractPdfParagraphs(textContent) {
        const items = (textContent.items || []).filter(it => it && typeof it.str === 'string');
        if (items.length === 0) return [];

        // Step 1: group items into lines by y-coordinate
        const rawLines = [];
        let currentLine = null;

        for (const item of items) {
            const tr = item.transform || [1, 0, 0, 1, 0, 0];
            const y = tr[5];
            const x = tr[4];
            const fontH = item.height || Math.abs(tr[3]) || 12;

            if (!currentLine) {
                currentLine = { y, x, height: fontH, items: [item] };
            } else {
                const yTol = Math.max(2, currentLine.height * 0.5);
                if (Math.abs(y - currentLine.y) <= yTol) {
                    currentLine.items.push(item);
                    if (fontH > currentLine.height) currentLine.height = fontH;
                    if (x < currentLine.x) currentLine.x = x;
                } else {
                    rawLines.push(currentLine);
                    currentLine = { y, x, height: fontH, items: [item] };
                }
            }

            // hasEOL flag from pdf.js: forces line break
            if (item.hasEOL) {
                rawLines.push(currentLine);
                currentLine = null;
            }
        }
        if (currentLine) rawLines.push(currentLine);

        // Step 2: turn each raw line into a structured line object with text + style
        const lines = rawLines.map(rl => this._buildPdfLine(rl)).filter(l => l !== null);
        if (lines.length === 0) return [];

        // Step 3: estimate typical inter-line gap to detect paragraph breaks
        const gaps = [];
        for (let k = 1; k < lines.length; k++) {
            const g = Math.abs(lines[k - 1].y - lines[k].y);
            if (g > 0) gaps.push(g);
        }
        gaps.sort((a, b) => a - b);
        const medianGap = gaps.length > 0 ? gaps[Math.floor(gaps.length / 2)] : 14;
        const paraGapThreshold = medianGap * 1.55;

        // Step 4: group lines into paragraphs by vertical gap
        const paragraphs = [];
        let buffer = [];

        for (let k = 0; k < lines.length; k++) {
            const line = lines[k];
            if (!line.text.trim()) {
                if (buffer.length > 0) {
                    paragraphs.push(this._mergePdfLines(buffer));
                    buffer = [];
                }
                continue;
            }

            if (buffer.length > 0) {
                const prev = buffer[buffer.length - 1];
                const gap = Math.abs(prev.y - line.y);
                if (gap > paraGapThreshold) {
                    paragraphs.push(this._mergePdfLines(buffer));
                    buffer = [];
                }
            }
            buffer.push(line);
        }
        if (buffer.length > 0) paragraphs.push(this._mergePdfLines(buffer));

        return paragraphs;
    },

    /**
     * Build a structured line from a group of pdf.js text items.
     */
    _buildPdfLine(rawLine) {
        // Sort items left-to-right
        const sorted = [...rawLine.items].sort((a, b) => {
            const ax = (a.transform || [0, 0, 0, 0, 0, 0])[4];
            const bx = (b.transform || [0, 0, 0, 0, 0, 0])[4];
            return ax - bx;
        });

        let text = '';
        let lastEnd = null;
        let lastWasCJK = false;

        for (const item of sorted) {
            if (!item.str) continue;
            const tr = item.transform || [1, 0, 0, 1, 0, 0];
            const x = tr[4];
            const w = item.width || 0;
            const fontH = item.height || Math.abs(tr[3]) || 12;
            const startsCJK = this._isCJKChar(item.str.charAt(0));

            if (lastEnd !== null && text && !text.endsWith(' ') && !item.str.startsWith(' ')) {
                const gap = x - lastEnd;
                // Insert space if there's a visible gap, but not between two CJK chars
                if (gap > fontH * 0.25 && !(lastWasCJK && startsCJK)) {
                    text += ' ';
                }
            }

            text += item.str;
            lastEnd = x + w;
            const lastChar = item.str.charAt(item.str.length - 1);
            lastWasCJK = this._isCJKChar(lastChar);
        }

        if (!text) return null;

        // Capture dominant font info from the first non-empty item
        const styleSrc = sorted.find(it => it.str && it.str.trim());
        const style = {};
        if (styleSrc) {
            if (styleSrc.fontName) style.fontName = styleSrc.fontName;
            const h = styleSrc.height || Math.abs((styleSrc.transform || [0, 0, 0, 12])[3]);
            if (h) style.fontSize = Math.max(8, Math.round(h * 2));
        }

        return {
            y: rawLine.y,
            x: rawLine.x,
            height: rawLine.height,
            text,
            style
        };
    },

    /**
     * Merge a sequence of lines belonging to the same paragraph, healing
     * soft wraps and hyphenation so the same content always normalizes
     * to the same paragraph string.
     */
    _mergePdfLines(lines) {
        let text = '';
        for (let i = 0; i < lines.length; i++) {
            const lineText = lines[i].text;
            if (i === 0) {
                text = lineText;
                continue;
            }

            const prevEnds = text.slice(-1);
            const nextStarts = lineText.charAt(0);

            if (prevEnds === '-' && /[a-zA-Z]/.test(nextStarts)) {
                // Hyphenated word break: drop the hyphen and join
                text = text.slice(0, -1) + lineText;
            } else if (this._isCJKChar(prevEnds) && this._isCJKChar(nextStarts)) {
                // CJK soft wrap: no inserted space
                text += lineText;
            } else {
                if (!text.endsWith(' ') && !lineText.startsWith(' ')) {
                    text += ' ';
                }
                text += lineText;
            }
        }

        text = text.trim();
        const style = lines[0] && lines[0].style ? { ...lines[0].style } : {};
        return {
            text,
            runs: [{ text, style }],
            style
        };
    },

    /**
     * Convert plain text (e.g. OCR output) into paragraph objects.
     */
    _textToParagraphs(text) {
        return text
            .split(/\n\s*\n|\n/)
            .map(t => t.trim())
            .filter(t => t.length > 0)
            .map(t => ({
                text: t,
                runs: [{ text: t, style: {} }],
                style: {}
            }));
    },

    _isCJKChar(ch) {
        if (!ch) return false;
        const code = ch.charCodeAt(0);
        return (code >= 0x4E00 && code <= 0x9FFF) ||
               (code >= 0x3400 && code <= 0x4DBF) ||
               (code >= 0xF900 && code <= 0xFAFF) ||
               (code >= 0x3040 && code <= 0x30FF);
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
