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

        // Collect all content-bearing children: direct runs plus runs nested
        // inside hyperlinks/smartTag/etc. We must strip every container that
        // carries the original text, otherwise the output duplicates content.
        const removables = [];
        const firstRuns = [];
        for (const child of Array.from(paraElem.children)) {
            if (child.namespaceURI !== wNS) continue;
            const local = child.localName;
            if (local === 'r') {
                removables.push(child);
                firstRuns.push(child);
            } else if (local === 'hyperlink' || local === 'smartTag' || local === 'sdt' ||
                       local === 'ins' || local === 'del') {
                // These wrap runs — remove them entirely along with their content.
                removables.push(child);
                const nested = child.getElementsByTagNameNS(wNS, 'r');
                if (nested.length > 0) firstRuns.push(nested[0]);
            }
        }

        // Get the run properties from the first run (to preserve formatting)
        let templateRPr = null;
        if (firstRuns.length > 0) {
            const rPr = firstRuns[0].getElementsByTagNameNS(wNS, 'rPr')[0];
            if (rPr) {
                templateRPr = rPr.cloneNode(true);
            }
        }

        // Remove all existing content-bearing children from paragraph
        for (const node of removables) {
            if (node.parentNode === paraElem) paraElem.removeChild(node);
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
     */
    _buildBodyXml(diffResult, author, date) {
        const lines = [];

        for (const diff of diffResult.diffs) {
            if (diff.type === 'equal') {
                // Unchanged paragraph
                const text = this._escapeXml(diff.paraA.text);
                lines.push(`    <w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`);
            } else if (diff.type === 'modified') {
                // Modified paragraph - inline track changes
                lines.push(this._buildModifiedParaXml(diff.changes, author, date));
            } else if (diff.type === 'deleted') {
                // Entire paragraph deleted
                const text = this._escapeXml(diff.paraA.text);
                const id1 = this._revId++;
                const id2 = this._revId++;
                lines.push(`    <w:p>
      <w:pPr><w:rPr><w:del w:id="${id1}" w:author="${this._escapeXml(author)}" w:date="${date}"/></w:rPr></w:pPr>
      <w:del w:id="${id2}" w:author="${this._escapeXml(author)}" w:date="${date}">
        <w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>
      </w:del>
    </w:p>`);
            } else if (diff.type === 'inserted') {
                // Entire paragraph inserted
                const text = this._escapeXml(diff.paraB.text);
                const id1 = this._revId++;
                const id2 = this._revId++;
                lines.push(`    <w:p>
      <w:pPr><w:rPr><w:ins w:id="${id1}" w:author="${this._escapeXml(author)}" w:date="${date}"/></w:rPr></w:pPr>
      <w:ins w:id="${id2}" w:author="${this._escapeXml(author)}" w:date="${date}">
        <w:r><w:t xml:space="preserve">${text}</w:t></w:r>
      </w:ins>
    </w:p>`);
            }
        }

        return lines.join('\n');
    },

    /**
     * Build XML for a modified paragraph with inline track changes.
     */
    _buildModifiedParaXml(changes, author, date) {
        const parts = ['    <w:p>'];

        for (const change of changes) {
            const text = this._escapeXml(change.value);
            if (change.added) {
                const id = this._revId++;
                parts.push(`      <w:ins w:id="${id}" w:author="${this._escapeXml(author)}" w:date="${date}">
        <w:r><w:t xml:space="preserve">${text}</w:t></w:r>
      </w:ins>`);
            } else if (change.removed) {
                const id = this._revId++;
                parts.push(`      <w:del w:id="${id}" w:author="${this._escapeXml(author)}" w:date="${date}">
        <w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>
      </w:del>`);
            } else {
                parts.push(`      <w:r><w:t xml:space="preserve">${text}</w:t></w:r>`);
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
