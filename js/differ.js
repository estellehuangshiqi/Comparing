/**
 * Diff engine: paragraph alignment + character/word level diff.
 */
const Differ = {

    /**
     * Compare two DocContent structures.
     * @param {DocContent} docA - Original document
     * @param {DocContent} docB - Modified document
     * @param {boolean} strictMode - If true, skip normalization
     * @returns {Object} Diff result with paired paragraphs and their diffs
     */
    compare(docA, docB, strictMode) {
        const parasA = docA.paragraphs;
        const parasB = docB.paragraphs;

        // Step 1: Align paragraphs using LCS-based similarity matching
        const alignment = this._alignParagraphs(parasA, parasB, strictMode);

        // Step 2: For each aligned pair, compute character-level diffs
        const diffs = [];
        let insCount = 0, delCount = 0, modCount = 0;
        let insChars = 0, delChars = 0;
        const totalCharsA = parasA.reduce((s, p) => s + p.text.length, 0);

        for (const item of alignment) {
            if (item.type === 'equal') {
                diffs.push({
                    type: 'equal',
                    paraA: item.paraA,
                    paraB: item.paraB,
                    indexA: item.indexA,
                    indexB: item.indexB,
                    changes: null
                });
            } else if (item.type === 'modified') {
                const changes = this._diffText(item.paraA.text, item.paraB.text, strictMode);
                // Check if there are actual changes after normalization
                const hasRealChanges = changes.some(c => c.added || c.removed);
                if (!hasRealChanges) {
                    diffs.push({
                        type: 'equal',
                        paraA: item.paraA,
                        paraB: item.paraB,
                        indexA: item.indexA,
                        indexB: item.indexB,
                        changes: null
                    });
                } else {
                    diffs.push({
                        type: 'modified',
                        paraA: item.paraA,
                        paraB: item.paraB,
                        indexA: item.indexA,
                        indexB: item.indexB,
                        changes
                    });
                    modCount++;
                    for (const c of changes) {
                        if (c.added) insChars += c.value.length;
                        if (c.removed) delChars += c.value.length;
                    }
                }
            } else if (item.type === 'deleted') {
                diffs.push({
                    type: 'deleted',
                    paraA: item.paraA,
                    paraB: null,
                    indexA: item.indexA,
                    indexB: -1,
                    changes: null
                });
                delCount++;
                delChars += item.paraA.text.length;
            } else if (item.type === 'inserted') {
                diffs.push({
                    type: 'inserted',
                    paraA: null,
                    paraB: item.paraB,
                    indexA: -1,
                    indexB: item.indexB,
                    changes: null
                });
                insCount++;
                insChars += item.paraB.text.length;
            }
        }

        // Granularity warning check
        const modifiedChars = insChars + delChars;
        const similarity = totalCharsA > 0 ? 1 - (modifiedChars / (totalCharsA * 2)) : 1;
        let warning = null;
        if (modifiedChars > totalCharsA * 0.7 && similarity > 0.3) {
            warning = '修订粒度可能过粗，建议检查对比结果';
        }

        return {
            diffs,
            stats: {
                totalParagraphs: Math.max(parasA.length, parasB.length),
                insertedParagraphs: insCount,
                deletedParagraphs: delCount,
                modifiedParagraphs: modCount,
                unchangedParagraphs: diffs.filter(d => d.type === 'equal').length,
                insertedChars: insChars,
                deletedChars: delChars,
            },
            warning
        };
    },

    /**
     * Align paragraphs from two documents using LCS-based approach.
     * This ensures that matching paragraphs are paired, and unmatched ones are marked as inserted/deleted.
     */
    _alignParagraphs(parasA, parasB, strictMode) {
        const n = parasA.length;
        const m = parasB.length;

        if (n === 0 && m === 0) return [];
        if (n === 0) return parasB.map((p, i) => ({ type: 'inserted', paraB: p, indexB: i }));
        if (m === 0) return parasA.map((p, i) => ({ type: 'deleted', paraA: p, indexA: i }));

        // Compute similarity matrix
        const simMatrix = [];
        for (let i = 0; i < n; i++) {
            simMatrix[i] = [];
            for (let j = 0; j < m; j++) {
                simMatrix[i][j] = this._similarity(parasA[i].text, parasB[j].text, strictMode);
            }
        }

        // Use dynamic programming to find optimal alignment (similar to LCS but with similarity scores)
        const MATCH_THRESHOLD = 0.3; // Minimum similarity to consider as a match
        const dp = Array.from({ length: n + 1 }, () => new Float64Array(m + 1));
        const trace = Array.from({ length: n + 1 }, () => new Int8Array(m + 1));

        // Fill DP table
        for (let i = 1; i <= n; i++) {
            for (let j = 1; j <= m; j++) {
                const sim = simMatrix[i - 1][j - 1];
                const matchScore = sim >= MATCH_THRESHOLD ? dp[i - 1][j - 1] + sim : -Infinity;
                const skipA = dp[i - 1][j];
                const skipB = dp[i][j - 1];

                if (matchScore >= skipA && matchScore >= skipB && matchScore > -Infinity) {
                    dp[i][j] = matchScore;
                    trace[i][j] = 0; // match
                } else if (skipA >= skipB) {
                    dp[i][j] = skipA;
                    trace[i][j] = 1; // skip A (delete)
                } else {
                    dp[i][j] = skipB;
                    trace[i][j] = 2; // skip B (insert)
                }
            }
        }

        // Trace back to get alignment
        const result = [];
        let i = n, j = m;
        while (i > 0 || j > 0) {
            if (i > 0 && j > 0 && trace[i][j] === 0) {
                const sim = simMatrix[i - 1][j - 1];
                if (sim >= 0.99) {
                    result.unshift({
                        type: 'equal',
                        paraA: parasA[i - 1],
                        paraB: parasB[j - 1],
                        indexA: i - 1,
                        indexB: j - 1
                    });
                } else {
                    result.unshift({
                        type: 'modified',
                        paraA: parasA[i - 1],
                        paraB: parasB[j - 1],
                        indexA: i - 1,
                        indexB: j - 1
                    });
                }
                i--; j--;
            } else if (i > 0 && (j === 0 || trace[i][j] === 1)) {
                result.unshift({
                    type: 'deleted',
                    paraA: parasA[i - 1],
                    indexA: i - 1
                });
                i--;
            } else {
                result.unshift({
                    type: 'inserted',
                    paraB: parasB[j - 1],
                    indexB: j - 1
                });
                j--;
            }
        }

        return result;
    },

    /**
     * Calculate similarity between two strings (0-1)
     */
    _similarity(a, b, strictMode) {
        const textA = strictMode ? a : Normalizer.normalize(a);
        const textB = strictMode ? b : Normalizer.normalize(b);

        if (textA === textB) return 1.0;
        if (!textA && !textB) return 1.0;
        if (!textA || !textB) return 0.0;

        // Quick length-based filter
        const lenRatio = Math.min(textA.length, textB.length) / Math.max(textA.length, textB.length);
        if (lenRatio < 0.3) return 0.0;

        // Use Levenshtein-based similarity for short strings, LCS ratio for longer
        if (textA.length < 200 && textB.length < 200) {
            const dist = this._levenshtein(textA, textB);
            return 1 - dist / Math.max(textA.length, textB.length);
        }

        // For longer strings, use a simpler approach: common character ratio
        return this._lcsRatio(textA, textB);
    },

    /**
     * Levenshtein distance
     */
    _levenshtein(a, b) {
        const n = a.length, m = b.length;
        if (n === 0) return m;
        if (m === 0) return n;

        // Use two rows for space optimization
        let prev = new Uint32Array(m + 1);
        let curr = new Uint32Array(m + 1);
        for (let j = 0; j <= m; j++) prev[j] = j;

        for (let i = 1; i <= n; i++) {
            curr[0] = i;
            for (let j = 1; j <= m; j++) {
                if (a[i - 1] === b[j - 1]) {
                    curr[j] = prev[j - 1];
                } else {
                    curr[j] = 1 + Math.min(prev[j - 1], prev[j], curr[j - 1]);
                }
            }
            [prev, curr] = [curr, prev];
        }
        return prev[m];
    },

    /**
     * LCS ratio for longer texts (approximate using chunks)
     */
    _lcsRatio(a, b) {
        // Sample-based approach for performance
        const sampleSize = 100;
        const chunkSize = Math.max(2, Math.floor(a.length / sampleSize));
        const chunks = new Set();
        for (let i = 0; i <= a.length - chunkSize; i += chunkSize) {
            chunks.add(a.substring(i, i + chunkSize));
        }
        let matches = 0;
        const totalChunksB = Math.ceil(b.length / chunkSize);
        for (let i = 0; i <= b.length - chunkSize; i += chunkSize) {
            if (chunks.has(b.substring(i, i + chunkSize))) matches++;
        }
        return totalChunksB > 0 ? matches / totalChunksB : 0;
    },

    /**
     * Character/word level diff between two texts.
     * Chinese: character-level; English: word-level.
     * Returns array of {value, added?, removed?} objects.
     */
    _diffText(textA, textB, strictMode) {
        let a = textA;
        let b = textB;

        if (!strictMode) {
            a = Normalizer.normalize(a);
            b = Normalizer.normalize(b);
        }

        if (a === b) return [{ value: textA }];

        // Tokenize: split into characters for CJK, words for Latin
        const tokensA = this._tokenize(a);
        const tokensB = this._tokenize(b);

        // Use jsdiff on tokens
        const changes = Diff.diffArrays(tokensA, tokensB);

        // Merge tokens back into strings and consolidate adjacent same-type operations
        const result = [];
        for (const change of changes) {
            const text = change.value.join('');
            if (text.length === 0) continue;

            const entry = { value: text };
            if (change.added) entry.added = true;
            if (change.removed) entry.removed = true;

            // Merge with previous if same type
            if (result.length > 0) {
                const prev = result[result.length - 1];
                if (prev.added === entry.added && prev.removed === entry.removed) {
                    prev.value += entry.value;
                    continue;
                }
            }
            result.push(entry);
        }

        return result;
    },

    /**
     * Tokenize text: CJK characters individually, English words as units.
     * @param {string} text
     * @returns {string[]}
     */
    _tokenize(text) {
        const tokens = [];
        let i = 0;
        while (i < text.length) {
            const ch = text.charCodeAt(i);
            // CJK Unified Ideographs and common ranges
            if (
                (ch >= 0x4E00 && ch <= 0x9FFF) ||  // CJK Unified
                (ch >= 0x3400 && ch <= 0x4DBF) ||  // CJK Extension A
                (ch >= 0xF900 && ch <= 0xFAFF) ||  // CJK Compatibility
                (ch >= 0x3000 && ch <= 0x303F) ||  // CJK Punctuation
                (ch >= 0xFF00 && ch <= 0xFFEF)     // Fullwidth
            ) {
                tokens.push(text[i]);
                i++;
            } else if (/\s/.test(text[i])) {
                // Whitespace
                let ws = '';
                while (i < text.length && /\s/.test(text[i])) {
                    ws += text[i];
                    i++;
                }
                tokens.push(ws);
            } else if (/[a-zA-Z0-9]/.test(text[i])) {
                // Latin word
                let word = '';
                while (i < text.length && /[a-zA-Z0-9''-]/.test(text[i])) {
                    word += text[i];
                    i++;
                }
                tokens.push(word);
            } else {
                // Other characters (punctuation, etc.)
                tokens.push(text[i]);
                i++;
            }
        }
        return tokens;
    }
};
