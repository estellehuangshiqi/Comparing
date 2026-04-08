/**
 * Diff engine: paragraph alignment + character/word level diff.
 *
 * Key design decisions:
 * - Paragraph alignment uses LCS-based DP with similarity scoring
 * - Match threshold 0.5: only pair paragraphs that are genuinely related
 * - Post-processing: if >60% of a "modified" pair is changes, split into delete+insert
 * - Character-level diff preserves original text (normalization only for comparison)
 * - N-gram Jaccard similarity for long texts (replaces broken chunk-based approach)
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
                const hasRealChanges = changes.some(c => c.added || c.removed);

                if (!hasRealChanges) {
                    // Normalization made them equal - treat as unchanged
                    diffs.push({
                        type: 'equal',
                        paraA: item.paraA,
                        paraB: item.paraB,
                        indexA: item.indexA,
                        indexB: item.indexB,
                        changes: null
                    });
                } else {
                    // Calculate change ratio to decide display strategy
                    // If most of the text is changed, it's cleaner to show as delete+insert
                    let changedLen = 0, totalLen = 0;
                    for (const c of changes) {
                        totalLen += c.value.length;
                        if (c.added || c.removed) changedLen += c.value.length;
                    }
                    const changeRatio = totalLen > 0 ? changedLen / totalLen : 0;

                    if (changeRatio > 0.5) {
                        // Too many changes - split into delete + insert for minimal display
                        diffs.push({
                            type: 'deleted',
                            paraA: item.paraA,
                            paraB: null,
                            indexA: item.indexA,
                            indexB: -1,
                            changes: null
                        });
                        diffs.push({
                            type: 'inserted',
                            paraA: null,
                            paraB: item.paraB,
                            indexA: -1,
                            indexB: item.indexB,
                            changes: null
                        });
                        delCount++;
                        insCount++;
                        delChars += item.paraA.text.length;
                        insChars += item.paraB.text.length;
                    } else {
                        // Genuine inline modification - show character-level diff
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
            warning: null
        };
    },

    /**
     * Align paragraphs from two documents using LCS-based approach.
     * Pre-normalizes all texts to avoid redundant normalization in O(n*m) comparisons.
     */
    _alignParagraphs(parasA, parasB, strictMode) {
        const n = parasA.length;
        const m = parasB.length;

        if (n === 0 && m === 0) return [];
        if (n === 0) return parasB.map((p, i) => ({ type: 'inserted', paraB: p, indexB: i }));
        if (m === 0) return parasA.map((p, i) => ({ type: 'deleted', paraA: p, indexA: i }));

        // Pre-normalize texts once for all similarity computations
        const normsA = parasA.map(p => strictMode ? p.text : Normalizer.normalize(p.text));
        const normsB = parasB.map(p => strictMode ? p.text : Normalizer.normalize(p.text));

        // Compute similarity matrix
        const simMatrix = [];
        for (let i = 0; i < n; i++) {
            simMatrix[i] = new Float64Array(m);
            for (let j = 0; j < m; j++) {
                if (normsA[i] === normsB[j]) {
                    simMatrix[i][j] = 1.0;
                } else {
                    simMatrix[i][j] = this._similarityPreNorm(normsA[i], normsB[j]);
                }
            }
        }

        // Dynamic programming: find optimal alignment maximizing total similarity
        // Only pair paragraphs with similarity >= MATCH_THRESHOLD
        const MATCH_THRESHOLD = 0.5;
        const dp = Array.from({ length: n + 1 }, () => new Float64Array(m + 1));
        const trace = Array.from({ length: n + 1 }, () => new Int8Array(m + 1));

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
     * Calculate similarity between two pre-normalized strings (0-1).
     * Uses Levenshtein for texts up to 2000 chars, n-gram Jaccard for longer.
     */
    _similarityPreNorm(textA, textB) {
        if (textA === textB) return 1.0;
        if (!textA && !textB) return 1.0;
        if (!textA || !textB) return 0.0;

        const lenA = textA.length;
        const lenB = textB.length;
        const maxLen = Math.max(lenA, lenB);
        const lenRatio = Math.min(lenA, lenB) / maxLen;

        // Quick reject for very different lengths
        if (lenRatio < 0.2) return 0.0;

        // Use Levenshtein for texts up to 2000 chars (accurate, O(n*m))
        if (maxLen <= 2000) {
            const dist = this._levenshtein(textA, textB);
            return 1 - dist / maxLen;
        }

        // For longer texts, use n-gram Jaccard similarity (fast, O(n+m))
        return this._ngramSimilarity(textA, textB, 3);
    },

    /**
     * Backward-compatible similarity with optional normalization.
     */
    _similarity(a, b, strictMode) {
        const textA = strictMode ? a : Normalizer.normalize(a);
        const textB = strictMode ? b : Normalizer.normalize(b);
        return this._similarityPreNorm(textA, textB);
    },

    /**
     * N-gram based Jaccard similarity for longer texts.
     * Uses character-level n-grams with multiset Jaccard index.
     * Much more accurate than the previous chunk-based approach.
     */
    _ngramSimilarity(a, b, n) {
        if (a.length < n && b.length < n) {
            return a === b ? 1.0 : 0.0;
        }

        const ngramsA = new Map();
        for (let i = 0; i <= a.length - n; i++) {
            const ng = a.substring(i, i + n);
            ngramsA.set(ng, (ngramsA.get(ng) || 0) + 1);
        }

        const ngramsB = new Map();
        for (let i = 0; i <= b.length - n; i++) {
            const ng = b.substring(i, i + n);
            ngramsB.set(ng, (ngramsB.get(ng) || 0) + 1);
        }

        let intersection = 0;
        let union = 0;

        const allKeys = new Set([...ngramsA.keys(), ...ngramsB.keys()]);
        for (const key of allKeys) {
            const countA = ngramsA.get(key) || 0;
            const countB = ngramsB.get(key) || 0;
            intersection += Math.min(countA, countB);
            union += Math.max(countA, countB);
        }

        return union > 0 ? intersection / union : 0;
    },

    /**
     * Levenshtein distance (space-optimized two-row approach).
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
     * Character/word level diff between two texts.
     * Chinese: character-level; English: word-level.
     * Returns array of {value, added?, removed?} objects.
     *
     * In non-strict mode, uses original text for output but normalization
     * for comparison via a custom comparator. This preserves the original
     * character forms (e.g., full-width punctuation) in the result while
     * ignoring normalization-level differences.
     */
    _diffText(textA, textB, strictMode) {
        // Exact match - no changes
        if (textA === textB) return [{ value: textA }];

        // Normalized match (non-strict only) - treat as equal
        if (!strictMode && Normalizer.normalize(textA) === Normalizer.normalize(textB)) {
            return [{ value: textA }];
        }

        // Tokenize ORIGINAL text (preserves original characters in output)
        const tokensA = this._tokenize(textA);
        const tokensB = this._tokenize(textB);

        // For non-strict mode, use a comparator that normalizes tokens before comparing
        // This way the diff output contains original text but comparison ignores
        // normalization differences (full-width/half-width, whitespace, etc.)
        const options = {};
        if (!strictMode) {
            const cache = new Map();
            const norm = (t) => {
                let v = cache.get(t);
                if (v === undefined) {
                    v = Normalizer.normalize(t);
                    cache.set(t, v);
                }
                return v;
            };
            options.comparator = (a, b) => norm(a) === norm(b);
        }

        const changes = Diff.diffArrays(tokensA, tokensB, options);

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

        // Post-process: collapse "noisy bursts" of alternating tiny changes into a
        // single before/after pair so the rendered revision is minimal and readable.
        return this._collapseChangeBursts(result);
    },

    /**
     * Collapse bursts of small alternating changes into a single delete+insert
     * pair. A "burst" is a region with several add/remove segments separated
     * only by short equal runs. Replacing the whole region with one removed
     * block followed by one added block produces a much cleaner revision.
     */
    _collapseChangeBursts(changes) {
        if (changes.length < 3) return changes;

        const SMALL_GAP = 4;          // chars in equal-runs that count as gaps
        const MIN_CHANGES_TO_COLLAPSE = 3;

        const result = [];
        let i = 0;

        while (i < changes.length) {
            const cur = changes[i];

            // Pure equal segments pass through
            if (!cur.added && !cur.removed) {
                result.push(cur);
                i++;
                continue;
            }

            // Find the extent of a burst starting at i.
            // A burst extends as long as we see change segments, possibly
            // separated by small equal runs.
            let j = i;
            let lastChangeEnd = i;
            while (j < changes.length) {
                const c = changes[j];
                if (c.added || c.removed) {
                    lastChangeEnd = j;
                    j++;
                    continue;
                }
                // equal run: only continue burst if short AND followed by another change
                if (c.value.length <= SMALL_GAP && j + 1 < changes.length &&
                    (changes[j + 1].added || changes[j + 1].removed)) {
                    j++;
                    continue;
                }
                break;
            }

            // Burst spans [i, lastChangeEnd]
            const burst = changes.slice(i, lastChangeEnd + 1);
            const changeCount = burst.filter(c => c.added || c.removed).length;

            if (changeCount >= MIN_CHANGES_TO_COLLAPSE) {
                // Collapse: emit one removed (old text) + one added (new text)
                let oldText = '';
                let newText = '';
                for (const c of burst) {
                    if (c.removed) {
                        oldText += c.value;
                    } else if (c.added) {
                        newText += c.value;
                    } else {
                        // Equal run inside the burst belongs to both versions
                        oldText += c.value;
                        newText += c.value;
                    }
                }
                if (oldText) result.push({ value: oldText, removed: true });
                if (newText) result.push({ value: newText, added: true });
            } else {
                for (const c of burst) result.push(c);
            }

            i = lastChangeEnd + 1;
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
                while (i < text.length && /[a-zA-Z0-9''\-]/.test(text[i])) {
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
