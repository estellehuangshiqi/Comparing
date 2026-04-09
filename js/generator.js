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
                // Insert new paragraph with insertion marks
                const refIndex = this._findInsertionPoint(diff, diffResult.diffs, bodyParas.length);
                const newPara = this._createInsertedParagraph(xmlDoc, diff.paraB, author, date);
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
     */
    _createInsertedParagraph(xmlDoc, paraInfo, author, date) {
        const wNS = this.W_NS;
        const para = xmlDoc.createElementNS(wNS, 'w:p');

        // Paragraph properties with insertion mark
        const pPr = xmlDoc.createElementNS(wNS, 'w:pPr');
        const rPr = xmlDoc.createElementNS(wNS, 'w:rPr');
        const insMark = xmlDoc.createElementNS(wNS, 'w:ins');
        insMark.setAttribute('w:id', String(this._revId++));
        insMark.setAttribute('w:author', author);
        insMark.setAttribute('w:date', date);
        rPr.appendChild(insMark);
        pPr.appendChild(rPr);
        para.appendChild(pPr);

        // Wrap content in <w:ins>
        const insElem = xmlDoc.createElementNS(wNS, 'w:ins');
        insElem.setAttribute('w:id', String(this._revId++));
        insElem.setAttribute('w:author', author);
        insElem.setAttribute('w:date', date);

        const text = paraInfo.text || '';
        const run = this._createRun(xmlDoc, text, null);
        insElem.appendChild(run);
        para.appendChild(insElem);

        return para;
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
     * Preserves per-run style information (font, size, bold, italic) when
     * available so non-DOCX sources (e.g. PDF) render with formatting close
     * to the original rather than a plain-text dump.
     */
    _buildBodyXml(diffResult, author, date) {
        const lines = [];
        const authorEsc = this._escapeXml(author);

        for (const diff of diffResult.diffs) {
            if (diff.type === 'equal') {
                lines.push(this._buildPlainParaXml(diff.paraB || diff.paraA));
            } else if (diff.type === 'modified') {
                lines.push(this._buildModifiedParaXml(diff.changes, diff.paraA, diff.paraB, authorEsc, date));
            } else if (diff.type === 'deleted') {
                lines.push(this._buildWholeParaDeletedXml(diff.paraA, authorEsc, date));
            } else if (diff.type === 'inserted') {
                lines.push(this._buildWholeParaInsertedXml(diff.paraB, authorEsc, date));
            }
        }

        return lines.join('\n');
    },

    /**
     * Build an unchanged paragraph preserving its runs/styles.
     */
    _buildPlainParaXml(para) {
        const parts = ['    <w:p>'];
        const pPr = this._buildPPrXml(para && para.style);
        if (pPr) parts.push('      ' + pPr);
        for (const run of this._runsFor(para)) {
            parts.push('      ' + this._buildRunXml(run.text, run.style, false));
        }
        parts.push('    </w:p>');
        return parts.join('\n');
    },

    /**
     * Build a paragraph that is entirely deleted (wrap all runs in w:del).
     */
    _buildWholeParaDeletedXml(para, authorEsc, date) {
        const parts = ['    <w:p>'];
        const delIdMark = this._revId++;
        const pPrXml = this._buildPPrXml(para && para.style, `<w:rPr><w:del w:id="${delIdMark}" w:author="${authorEsc}" w:date="${date}"/></w:rPr>`);
        if (pPrXml) parts.push('      ' + pPrXml);
        const delId = this._revId++;
        parts.push(`      <w:del w:id="${delId}" w:author="${authorEsc}" w:date="${date}">`);
        for (const run of this._runsFor(para)) {
            parts.push('        ' + this._buildRunXml(run.text, run.style, true));
        }
        parts.push('      </w:del>');
        parts.push('    </w:p>');
        return parts.join('\n');
    },

    /**
     * Build a paragraph that is entirely inserted (wrap all runs in w:ins).
     */
    _buildWholeParaInsertedXml(para, authorEsc, date) {
        const parts = ['    <w:p>'];
        const insIdMark = this._revId++;
        const pPrXml = this._buildPPrXml(para && para.style, `<w:rPr><w:ins w:id="${insIdMark}" w:author="${authorEsc}" w:date="${date}"/></w:rPr>`);
        if (pPrXml) parts.push('      ' + pPrXml);
        const insId = this._revId++;
        parts.push(`      <w:ins w:id="${insId}" w:author="${authorEsc}" w:date="${date}">`);
        for (const run of this._runsFor(para)) {
            parts.push('        ' + this._buildRunXml(run.text, run.style, false));
        }
        parts.push('      </w:ins>');
        parts.push('    </w:p>');
        return parts.join('\n');
    },

    /**
     * Build XML for a modified paragraph with inline track changes.
     * Splits each diff segment across the original runs so that the
     * surviving formatting (font/size/bold/italic) is carried into the
     * revision document.
     */
    _buildModifiedParaXml(changes, paraA, paraB, authorEsc, date) {
        const parts = ['    <w:p>'];
        const pPr = this._buildPPrXml((paraA && paraA.style) || (paraB && paraB.style));
        if (pPr) parts.push('      ' + pPr);

        const runsA = this._runsFor(paraA);
        const runsB = this._runsFor(paraB);

        let offsetA = 0;
        let offsetB = 0;

        for (const change of changes) {
            if (!change.value) continue;

            if (change.added) {
                const id = this._revId++;
                const chunks = this._sliceRuns(runsB, offsetB, change.value.length, change.value);
                parts.push(`      <w:ins w:id="${id}" w:author="${authorEsc}" w:date="${date}">`);
                for (const chunk of chunks) {
                    parts.push('        ' + this._buildRunXml(chunk.text, chunk.style, false));
                }
                parts.push('      </w:ins>');
                offsetB += change.value.length;
            } else if (change.removed) {
                const id = this._revId++;
                const chunks = this._sliceRuns(runsA, offsetA, change.value.length, change.value);
                parts.push(`      <w:del w:id="${id}" w:author="${authorEsc}" w:date="${date}">`);
                for (const chunk of chunks) {
                    parts.push('        ' + this._buildRunXml(chunk.text, chunk.style, true));
                }
                parts.push('      </w:del>');
                offsetA += change.value.length;
            } else {
                const chunks = this._sliceRuns(runsA, offsetA, change.value.length, change.value);
                for (const chunk of chunks) {
                    parts.push('      ' + this._buildRunXml(chunk.text, chunk.style, false));
                }
                offsetA += change.value.length;
                offsetB += change.value.length;
            }
        }

        parts.push('    </w:p>');
        return parts.join('\n');
    },

    /**
     * Return a usable runs array for a paragraph. Falls back to a synthetic
     * run when the paragraph has no run data.
     */
    _runsFor(para) {
        if (!para) return [{ text: '', style: {} }];
        if (para.runs && para.runs.length > 0) {
            return para.runs.filter(r => r.text && r.text.length > 0);
        }
        return [{ text: para.text || '', style: {} }];
    },

    /**
     * Slice a [startOffset, startOffset + length) window out of the runs
     * array. If the text actually living at that window does not match the
     * expected segment (e.g. because diff was computed on a slightly
     * different representation) we fall back to a single chunk using the
     * style of the run at startOffset.
     */
    _sliceRuns(runs, startOffset, length, expectedText) {
        const out = [];
        if (!runs || runs.length === 0 || length <= 0) {
            if (expectedText) out.push({ text: expectedText, style: {} });
            return out;
        }

        let pos = 0;
        let remaining = length;
        let cursor = startOffset;
        let produced = '';

        for (const run of runs) {
            if (remaining <= 0) break;
            const runLen = run.text.length;
            const runStart = pos;
            const runEnd = pos + runLen;
            pos = runEnd;
            if (runEnd <= cursor) continue;

            const localStart = Math.max(0, cursor - runStart);
            const takeLen = Math.min(runLen - localStart, remaining);
            if (takeLen <= 0) continue;

            const chunkText = run.text.substr(localStart, takeLen);
            produced += chunkText;
            out.push({ text: chunkText, style: run.style || {} });
            cursor += takeLen;
            remaining -= takeLen;
        }

        // Fallback: if the produced text does not line up with the expected
        // segment (which can happen when runs and the diff source strings
        // diverged) emit the expected text verbatim with a best-guess style.
        if (expectedText && produced !== expectedText) {
            const style = (out[0] && out[0].style) ||
                          (runs[0] && runs[0].style) || {};
            return [{ text: expectedText, style }];
        }

        if (out.length === 0 && expectedText) {
            out.push({ text: expectedText, style: {} });
        }
        return out;
    },

    /**
     * Serialize paragraph properties (pPr) for the generated body.
     * Accepts an optional extra XML fragment that is appended inside pPr.
     */
    _buildPPrXml(style, extra) {
        const fragments = [];
        if (style) {
            if (style.alignment) {
                const val = this._escapeXml(style.alignment);
                fragments.push(`<w:jc w:val="${val}"/>`);
            }
        }
        if (extra) fragments.push(extra);
        if (fragments.length === 0) return '';
        return `<w:pPr>${fragments.join('')}</w:pPr>`;
    },

    /**
     * Build an inline <w:r> element with style-aware <w:rPr>.
     */
    _buildRunXml(text, style, asDelete) {
        const rPr = this._buildRPrXml(style);
        const body = this._escapeXml(text || '');
        const textTag = asDelete ? 'w:delText' : 'w:t';
        return `<w:r>${rPr}<${textTag} xml:space="preserve">${body}</${textTag}></w:r>`;
    },

    /**
     * Build <w:rPr> for a given run style object.
     */
    _buildRPrXml(style) {
        if (!style) return '';
        const parts = [];
        if (style.font) {
            const f = this._escapeXml(style.font);
            parts.push(`<w:rFonts w:ascii="${f}" w:eastAsia="${f}" w:hAnsi="${f}" w:cs="${f}"/>`);
        }
        if (style.bold) parts.push('<w:b/><w:bCs/>');
        if (style.italic) parts.push('<w:i/><w:iCs/>');
        if (style.underline) parts.push('<w:u w:val="single"/>');
        if (style.fontSize) {
            const sz = Math.max(2, parseInt(style.fontSize, 10));
            parts.push(`<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`);
        }
        if (style.color) {
            const c = this._escapeXml(String(style.color).replace(/^#/, ''));
            parts.push(`<w:color w:val="${c}"/>`);
        }
        if (parts.length === 0) return '';
        return `<w:rPr>${parts.join('')}</w:rPr>`;
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
