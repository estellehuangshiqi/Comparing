/**
 * DOCX Generator with Track Changes (revision marks).
 * For DOCX-to-DOCX: modifies original XML to insert <w:ins>/<w:del>.
 * For other formats: generates new DOCX with track changes.
 */
const Generator = {

    /** Namespace URIs */
    W_NS: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    R_NS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    WP_NS: 'http://schemas.openxmlformats.org/package/2006/relationships',

    /** Running revision ID counter */
    _revId: 1,

    /**
     * Generate a DOCX with track changes from diff results.
     * @param {DocContent} docA - Original document
     * @param {DocContent} docB - Modified document
     * @param {Object} diffResult - From Differ.compare()
     * @param {string} author - Revision author name
     * @returns {Promise<Blob>} The generated .docx file as a Blob
     */
    async generate(docA, docB, diffResult, author) {
        this._revId = 1;
        const revAuthor = author || '作者';
        const revDate = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');

        // If original is DOCX, modify its XML directly to preserve formatting
        if (docA.metadata.sourceFormat === 'docx' && docA._zip) {
            return this._generateFromDocx(docA, docB, diffResult, revAuthor, revDate);
        }

        // Otherwise, generate a new DOCX
        return this._generateNewDocx(docA, docB, diffResult, revAuthor, revDate);
    },

    /**
     * Modify original DOCX XML to add track changes.
     */
    async _generateFromDocx(docA, docB, diffResult, author, date) {
        const zip = docA._zip;
        const xmlStr = docA._xmlStr;

        // Parse the XML fresh for modification
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlStr, 'application/xml');
        const body = xmlDoc.getElementsByTagNameNS(this.W_NS, 'body')[0];

        // Get all paragraph elements from body (direct children only)
        const bodyParas = [];
        for (const child of body.children) {
            if (child.localName === 'p' && child.namespaceURI === this.W_NS) {
                bodyParas.push(child);
            }
        }

        // Capture a default rPr/pPr template from the first non-empty body paragraph.
        // Used as fallback styling for inserted-from-docB paragraphs so they blend
        // visually with the surrounding original content.
        const { rPr: defaultRPr, pPr: defaultPPr } = this._captureDefaultParaTemplate(bodyParas);

        // Process diffs in reverse order to maintain index validity
        const sortedDiffs = [...diffResult.diffs].sort((a, b) => {
            const idxA = a.indexA >= 0 ? a.indexA : (a.indexB >= 0 ? a.indexB + 0.5 : 0);
            const idxB = b.indexA >= 0 ? b.indexA : (b.indexB >= 0 ? b.indexB + 0.5 : 0);
            return idxB - idxA;
        });

        // Track which original paragraphs need modification
        for (const diff of sortedDiffs) {
            if (diff.type === 'equal') continue;

            if (diff.type === 'modified' && diff.indexA >= 0 && diff.indexA < bodyParas.length) {
                // Modify existing paragraph with track changes
                this._applyModifiedDiff(xmlDoc, bodyParas[diff.indexA], diff.changes, author, date);
            } else if (diff.type === 'deleted' && diff.indexA >= 0 && diff.indexA < bodyParas.length) {
                // Mark entire paragraph as deleted
                this._applyDeletedParagraph(xmlDoc, bodyParas[diff.indexA], author, date);
            } else if (diff.type === 'inserted') {
                // Insert new paragraph with insertion marks; borrow nearby rPr/pPr
                // for visual consistency with surrounding original content.
                const refIndex = this._findInsertionPoint(diff, diffResult.diffs, bodyParas.length);
                const neighborTemplate = this._neighborTemplate(bodyParas, refIndex) || { rPr: defaultRPr, pPr: defaultPPr };
                const newPara = this._createInsertedParagraph(
                    xmlDoc, diff.paraB, author, date,
                    neighborTemplate.rPr, neighborTemplate.pPr
                );
                if (refIndex < bodyParas.length) {
                    body.insertBefore(newPara, bodyParas[refIndex]);
                } else {
                    // Append before sectPr if exists, otherwise at end of body
                    const sectPr = body.getElementsByTagNameNS(this.W_NS, 'sectPr')[0];
                    if (sectPr) {
                        body.insertBefore(newPara, sectPr);
                    } else {
                        body.appendChild(newPara);
                    }
                }
            }
        }

        // Serialize back to string
        const serializer = new XMLSerializer();
        const newXmlStr = serializer.serializeToString(xmlDoc);

        // Replace in zip
        zip.file('word/document.xml', newXmlStr);

        // Generate the output file
        const blob = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            compression: 'DEFLATE'
        });

        return blob;
    },

    /**
     * Apply character-level modifications to a paragraph element.
     */
    _applyModifiedDiff(xmlDoc, paraElem, changes, author, date) {
        if (!changes || changes.length === 0) return;

        const wNS = this.W_NS;

        // Collect all existing runs and their text
        const existingRuns = [];
        for (const child of Array.from(paraElem.children)) {
            if (child.localName === 'r' && child.namespaceURI === wNS) {
                existingRuns.push(child);
            }
        }

        // Get the run properties from the first run (to preserve formatting)
        let templateRPr = null;
        if (existingRuns.length > 0) {
            const rPr = existingRuns[0].getElementsByTagNameNS(wNS, 'rPr')[0];
            if (rPr) {
                templateRPr = rPr.cloneNode(true);
            }
        }

        // Remove all existing runs from paragraph
        for (const run of existingRuns) {
            paraElem.removeChild(run);
        }

        // Build new content from changes
        const pPr = paraElem.getElementsByTagNameNS(wNS, 'pPr')[0];
        let insertAfter = pPr || null;

        for (const change of changes) {
            if (!change.value) continue;

            if (change.added) {
                // Create insertion mark
                const insElem = xmlDoc.createElementNS(wNS, 'w:ins');
                insElem.setAttribute('w:id', String(this._revId++));
                insElem.setAttribute('w:author', author);
                insElem.setAttribute('w:date', date);

                const run = this._createRun(xmlDoc, change.value, templateRPr);
                insElem.appendChild(run);

                this._insertAfterNode(paraElem, insElem, insertAfter);
                insertAfter = insElem;
            } else if (change.removed) {
                // Create deletion mark
                const delElem = xmlDoc.createElementNS(wNS, 'w:del');
                delElem.setAttribute('w:id', String(this._revId++));
                delElem.setAttribute('w:author', author);
                delElem.setAttribute('w:date', date);

                const run = this._createDeleteRun(xmlDoc, change.value, templateRPr);
                delElem.appendChild(run);

                this._insertAfterNode(paraElem, delElem, insertAfter);
                insertAfter = delElem;
            } else {
                // Unchanged text
                const run = this._createRun(xmlDoc, change.value, templateRPr);
                this._insertAfterNode(paraElem, run, insertAfter);
                insertAfter = run;
            }
        }
    },

    /**
     * Mark an entire paragraph as deleted.
     */
    _applyDeletedParagraph(xmlDoc, paraElem, author, date) {
        const wNS = this.W_NS;

        // Wrap all runs in <w:del>
        const runs = Array.from(paraElem.children).filter(
            c => c.localName === 'r' && c.namespaceURI === wNS
        );

        for (const run of runs) {
            const delElem = xmlDoc.createElementNS(wNS, 'w:del');
            delElem.setAttribute('w:id', String(this._revId++));
            delElem.setAttribute('w:author', author);
            delElem.setAttribute('w:date', date);

            // Convert w:t to w:delText in the run
            const tElems = run.getElementsByTagNameNS(wNS, 't');
            for (const t of Array.from(tElems)) {
                const delText = xmlDoc.createElementNS(wNS, 'w:delText');
                delText.setAttribute('xml:space', 'preserve');
                delText.textContent = t.textContent;
                t.parentNode.replaceChild(delText, t);
            }

            paraElem.insertBefore(delElem, run);
            delElem.appendChild(run);
        }

        // Also mark paragraph mark as deleted
        let pPr = paraElem.getElementsByTagNameNS(wNS, 'pPr')[0];
        if (!pPr) {
            pPr = xmlDoc.createElementNS(wNS, 'w:pPr');
            paraElem.insertBefore(pPr, paraElem.firstChild);
        }
        let rPr = pPr.getElementsByTagNameNS(wNS, 'rPr')[0];
        if (!rPr) {
            rPr = xmlDoc.createElementNS(wNS, 'w:rPr');
            pPr.appendChild(rPr);
        }
        const delMark = xmlDoc.createElementNS(wNS, 'w:del');
        delMark.setAttribute('w:id', String(this._revId++));
        delMark.setAttribute('w:author', author);
        delMark.setAttribute('w:date', date);
        rPr.appendChild(delMark);
    },

    /**
     * Create a new paragraph element marked as inserted.
     * Optional templateRPr / templatePPr provide styling so the inserted
     * paragraph visually matches surrounding original content.
     */
    _createInsertedParagraph(xmlDoc, paraInfo, author, date, templateRPr, templatePPr) {
        const wNS = this.W_NS;
        const para = xmlDoc.createElementNS(wNS, 'w:p');

        // Build paragraph properties: clone template pPr (without any existing
        // del/ins marks), then attach an rPr containing the ins mark for the
        // paragraph mark itself.
        const pPr = templatePPr
            ? this._clonePPrWithoutRevisionMarks(xmlDoc, templatePPr)
            : xmlDoc.createElementNS(wNS, 'w:pPr');

        // Ensure rPr exists in pPr and contains the ins mark for paragraph mark
        let pPrRPr = pPr.getElementsByTagNameNS(wNS, 'rPr')[0];
        if (!pPrRPr) {
            pPrRPr = xmlDoc.createElementNS(wNS, 'w:rPr');
            pPr.appendChild(pPrRPr);
        }
        const insMark = xmlDoc.createElementNS(wNS, 'w:ins');
        insMark.setAttribute('w:id', String(this._revId++));
        insMark.setAttribute('w:author', author);
        insMark.setAttribute('w:date', date);
        pPrRPr.appendChild(insMark);
        para.appendChild(pPr);

        // Wrap content run in <w:ins>
        const insElem = xmlDoc.createElementNS(wNS, 'w:ins');
        insElem.setAttribute('w:id', String(this._revId++));
        insElem.setAttribute('w:author', author);
        insElem.setAttribute('w:date', date);

        const text = paraInfo.text || '';
        const run = this._createRun(xmlDoc, text, templateRPr || null);
        insElem.appendChild(run);
        para.appendChild(insElem);

        return para;
    },

    /**
     * Capture a default rPr / pPr template from the first non-empty paragraph.
     * Used as fallback styling for inserted paragraphs in cross-format diffs.
     */
    _captureDefaultParaTemplate(bodyParas) {
        const wNS = this.W_NS;
        for (const para of bodyParas) {
            const runs = Array.from(para.children).filter(
                c => c.localName === 'r' && c.namespaceURI === wNS
            );
            if (runs.length === 0) continue;

            const firstRun = runs.find(r => r.textContent && r.textContent.trim()) || runs[0];
            const rPrSrc = firstRun.getElementsByTagNameNS(wNS, 'rPr')[0];
            const pPrSrc = para.getElementsByTagNameNS(wNS, 'pPr')[0];

            return {
                rPr: rPrSrc ? rPrSrc.cloneNode(true) : null,
                pPr: pPrSrc ? pPrSrc.cloneNode(true) : null
            };
        }
        return { rPr: null, pPr: null };
    },

    /**
     * Pick an rPr / pPr template from a paragraph adjacent to the insertion
     * point so the new paragraph blends in with its neighbours.
     */
    _neighborTemplate(bodyParas, refIndex) {
        const wNS = this.W_NS;
        const candidates = [];
        if (refIndex > 0 && refIndex - 1 < bodyParas.length) candidates.push(bodyParas[refIndex - 1]);
        if (refIndex >= 0 && refIndex < bodyParas.length) candidates.push(bodyParas[refIndex]);

        for (const para of candidates) {
            if (!para) continue;
            const runs = Array.from(para.children).filter(
                c => c.localName === 'r' && c.namespaceURI === wNS
            );
            if (runs.length === 0) continue;

            const firstRun = runs.find(r => r.textContent && r.textContent.trim()) || runs[0];
            const rPrSrc = firstRun.getElementsByTagNameNS(wNS, 'rPr')[0];
            const pPrSrc = para.getElementsByTagNameNS(wNS, 'pPr')[0];

            return {
                rPr: rPrSrc ? rPrSrc.cloneNode(true) : null,
                pPr: pPrSrc ? pPrSrc.cloneNode(true) : null
            };
        }
        return null;
    },

    /**
     * Clone a pPr element while stripping any existing revision marks
     * (w:ins / w:del inside its rPr) — we'll add fresh ones for the new para.
     */
    _clonePPrWithoutRevisionMarks(xmlDoc, sourcePPr) {
        const wNS = this.W_NS;
        const cloned = sourcePPr.cloneNode(true);
        const rPr = cloned.getElementsByTagNameNS(wNS, 'rPr')[0];
        if (rPr) {
            const oldMarks = Array.from(rPr.children).filter(
                c => (c.localName === 'ins' || c.localName === 'del') && c.namespaceURI === wNS
            );
            for (const m of oldMarks) rPr.removeChild(m);
        }
        return cloned;
    },

    /**
     * Find the insertion point index for a new paragraph.
     */
    _findInsertionPoint(insertDiff, allDiffs, totalParas) {
        // Find the nearest preceding diff that has an indexA
        let bestIndex = totalParas;
        for (const d of allDiffs) {
            if (d.indexB < insertDiff.indexB && d.indexA >= 0) {
                bestIndex = d.indexA + 1;
            }
            if (d.indexB >= insertDiff.indexB && d.indexA >= 0) {
                bestIndex = d.indexA;
                break;
            }
        }
        return bestIndex;
    },

    /**
     * Create a w:r element with text.
     */
    _createRun(xmlDoc, text, templateRPr) {
        const wNS = this.W_NS;
        const run = xmlDoc.createElementNS(wNS, 'w:r');

        if (templateRPr) {
            run.appendChild(templateRPr.cloneNode(true));
        }

        const tElem = xmlDoc.createElementNS(wNS, 'w:t');
        tElem.setAttribute('xml:space', 'preserve');
        tElem.textContent = text;
        run.appendChild(tElem);

        return run;
    },

    /**
     * Create a w:r element with w:delText (for deletion runs).
     */
    _createDeleteRun(xmlDoc, text, templateRPr) {
        const wNS = this.W_NS;
        const run = xmlDoc.createElementNS(wNS, 'w:r');

        if (templateRPr) {
            run.appendChild(templateRPr.cloneNode(true));
        }

        const delText = xmlDoc.createElementNS(wNS, 'w:delText');
        delText.setAttribute('xml:space', 'preserve');
        delText.textContent = text;
        run.appendChild(delText);

        return run;
    },

    /**
     * Insert a node after a reference node in parent.
     */
    _insertAfterNode(parent, newNode, refNode) {
        if (refNode && refNode.nextSibling) {
            parent.insertBefore(newNode, refNode.nextSibling);
        } else if (refNode) {
            parent.appendChild(newNode);
        } else {
            // Insert at beginning (after any existing pPr)
            const firstChild = parent.firstChild;
            if (firstChild) {
                parent.insertBefore(newNode, firstChild.nextSibling || null);
            } else {
                parent.appendChild(newNode);
            }
        }
    },

    // ==================== New DOCX Generation (for non-DOCX sources) ====================

    /**
     * Generate a completely new DOCX file with track changes.
     */
    async _generateNewDocx(docA, docB, diffResult, author, date) {
        const zip = new JSZip();

        // Build document.xml content
        const bodyContent = this._buildBodyXml(diffResult, author, date);

        // [Content_Types].xml
        zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`);

        // _rels/.rels
        zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

        // word/_rels/document.xml.rels
        zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);

        // word/styles.xml
        zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="等线" w:eastAsia="等线" w:hAnsi="等线"/><w:sz w:val="21"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="480"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="200"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
</w:styles>`);

        // word/document.xml
        zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:mv="urn:schemas-microsoft-com:mac:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
  <w:body>
${bodyContent}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`);

        const blob = await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            compression: 'DEFLATE'
        });

        return blob;
    },

    /**
     * Build the body XML content from diff results (for new DOCX generation).
     *
     * Carries through paragraph style info captured by parsers (font, size,
     * alignment) so a PDF or XLSX source produces an output that visually
     * resembles the original. Inserts page breaks when paragraphs cross
     * page boundaries (PDF source).
     */
    _buildBodyXml(diffResult, author, date) {
        const lines = [];
        let lastPage = 0;

        for (const diff of diffResult.diffs) {
            const srcPara = diff.paraA || diff.paraB;
            const style = (srcPara && srcPara.style) || {};
            const pageNum = style.pageNumber || 0;

            // Insert page break when crossing pages (PDF source)
            if (pageNum && lastPage > 0 && pageNum > lastPage) {
                lines.push('    <w:p><w:r><w:br w:type="page"/></w:r></w:p>');
            }
            if (pageNum) lastPage = pageNum;

            const rPrInner = this._buildRPrInnerFromStyle(style);
            const rPrXml = rPrInner ? `<w:rPr>${rPrInner}</w:rPr>` : '';
            const pPrInner = this._buildPPrInnerFromStyle(style);
            const pPrXml = pPrInner ? `<w:pPr>${pPrInner}</w:pPr>` : '';

            if (diff.type === 'equal') {
                const text = this._escapeXml(diff.paraA.text);
                lines.push(`    <w:p>${pPrXml}<w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`);
            } else if (diff.type === 'modified') {
                lines.push(this._buildModifiedParaXml(diff.changes, author, date, rPrXml, pPrXml));
            } else if (diff.type === 'deleted') {
                const text = this._escapeXml(diff.paraA.text);
                const id1 = this._revId++;
                const id2 = this._revId++;
                const a = this._escapeXml(author);
                lines.push(`    <w:p>
      <w:pPr>${pPrInner}<w:rPr>${rPrInner}<w:del w:id="${id1}" w:author="${a}" w:date="${date}"/></w:rPr></w:pPr>
      <w:del w:id="${id2}" w:author="${a}" w:date="${date}">
        <w:r>${rPrXml}<w:delText xml:space="preserve">${text}</w:delText></w:r>
      </w:del>
    </w:p>`);
            } else if (diff.type === 'inserted') {
                const text = this._escapeXml(diff.paraB.text);
                const id1 = this._revId++;
                const id2 = this._revId++;
                const a = this._escapeXml(author);
                lines.push(`    <w:p>
      <w:pPr>${pPrInner}<w:rPr>${rPrInner}<w:ins w:id="${id1}" w:author="${a}" w:date="${date}"/></w:rPr></w:pPr>
      <w:ins w:id="${id2}" w:author="${a}" w:date="${date}">
        <w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r>
      </w:ins>
    </w:p>`);
            }
        }

        return lines.join('\n');
    },

    /**
     * Build the inner content of a w:rPr element from a parsed paragraph style.
     */
    _buildRPrInnerFromStyle(style) {
        if (!style) return '';
        const parts = [];
        if (style.bold) parts.push('<w:b/>');
        if (style.italic) parts.push('<w:i/>');
        if (style.underline) parts.push('<w:u w:val="single"/>');
        if (style.fontName) {
            const f = this._escapeXml(style.fontName);
            parts.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>`);
        }
        if (style.fontSize) {
            const sz = String(style.fontSize);
            parts.push(`<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`);
        }
        if (style.color) {
            parts.push(`<w:color w:val="${this._escapeXml(style.color)}"/>`);
        }
        return parts.join('');
    },

    /**
     * Build the inner content of a w:pPr element from a parsed paragraph style.
     */
    _buildPPrInnerFromStyle(style) {
        if (!style) return '';
        const parts = [];
        if (style.styleId) {
            parts.push(`<w:pStyle w:val="${this._escapeXml(style.styleId)}"/>`);
        }
        if (style.alignment) {
            parts.push(`<w:jc w:val="${this._escapeXml(style.alignment)}"/>`);
        }
        return parts.join('');
    },

    /**
     * Build XML for a modified paragraph with inline track changes.
     * Optional rPrXml/pPrXml carry through paragraph-level styling.
     */
    _buildModifiedParaXml(changes, author, date, rPrXml = '', pPrXml = '') {
        const parts = [`    <w:p>${pPrXml}`];
        const a = this._escapeXml(author);

        for (const change of changes) {
            const text = this._escapeXml(change.value);
            if (change.added) {
                const id = this._revId++;
                parts.push(`      <w:ins w:id="${id}" w:author="${a}" w:date="${date}">
        <w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r>
      </w:ins>`);
            } else if (change.removed) {
                const id = this._revId++;
                parts.push(`      <w:del w:id="${id}" w:author="${a}" w:date="${date}">
        <w:r>${rPrXml}<w:delText xml:space="preserve">${text}</w:delText></w:r>
      </w:del>`);
            } else {
                parts.push(`      <w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r>`);
            }
        }

        parts.push('    </w:p>');
        return parts.join('\n');
    },

    /**
     * Escape XML special characters.
     */
    _escapeXml(str) {
        if (!str) return '';
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    },

    /**
     * Validate the generated DOCX blob by checking ZIP integrity.
     * @param {Blob} blob
     * @returns {Promise<boolean>}
     */
    async validate(blob) {
        try {
            const zip = await JSZip.loadAsync(blob);
            const docXml = zip.file('word/document.xml');
            if (!docXml) return false;
            const content = await docXml.async('string');
            // Basic XML validity check
            const parser = new DOMParser();
            const doc = parser.parseFromString(content, 'application/xml');
            return !doc.querySelector('parsererror');
        } catch {
            return false;
        }
    }
};
