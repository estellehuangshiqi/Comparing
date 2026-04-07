/**
 * Text normalizer for semantic equivalence checking.
 * Normalizes text to reduce false-positive diffs.
 */
const Normalizer = {
    /** Full-width to half-width mapping */
    fullToHalf: {
        '（': '(', '）': ')', '【': '[', '】': ']',
        '｛': '{', '｝': '}', '，': ',', '。': '.',
        '；': ';', '：': ':', '？': '?', '！': '!',
        '＂': '"', '＇': "'", '～': '~', '＠': '@',
        '＃': '#', '＄': '$', '％': '%', '＆': '&',
        '＊': '*', '＋': '+', '－': '-', '／': '/',
        '＝': '=', '＜': '<', '＞': '>', '＾': '^',
        '＿': '_', '｜': '|', '｀': '`', '＼': '\\',
        '０': '0', '１': '1', '２': '2', '３': '3', '４': '4',
        '５': '5', '６': '6', '７': '7', '８': '8', '９': '9',
        'Ａ': 'A', 'Ｂ': 'B', 'Ｃ': 'C', 'Ｄ': 'D', 'Ｅ': 'E',
        'Ｆ': 'F', 'Ｇ': 'G', 'Ｈ': 'H', 'Ｉ': 'I', 'Ｊ': 'J',
        'Ｋ': 'K', 'Ｌ': 'L', 'Ｍ': 'M', 'Ｎ': 'N', 'Ｏ': 'O',
        'Ｐ': 'P', 'Ｑ': 'Q', 'Ｒ': 'R', 'Ｓ': 'S', 'Ｔ': 'T',
        'Ｕ': 'U', 'Ｖ': 'V', 'Ｗ': 'W', 'Ｘ': 'X', 'Ｙ': 'Y', 'Ｚ': 'Z',
        'ａ': 'a', 'ｂ': 'b', 'ｃ': 'c', 'ｄ': 'd', 'ｅ': 'e',
        'ｆ': 'f', 'ｇ': 'g', 'ｈ': 'h', 'ｉ': 'i', 'ｊ': 'j',
        'ｋ': 'k', 'ｌ': 'l', 'ｍ': 'm', 'ｎ': 'n', 'ｏ': 'o',
        'ｐ': 'p', 'ｑ': 'q', 'ｒ': 'r', 'ｓ': 's', 'ｔ': 't',
        'ｕ': 'u', 'ｖ': 'v', 'ｗ': 'w', 'ｘ': 'x', 'ｙ': 'y', 'ｚ': 'z',
    },

    /** Invisible characters to remove */
    invisibleChars: /[\u200B\u200C\u200D\uFEFF\u00AD\u200E\u200F\u2028\u2029\u202A-\u202E\u2060-\u2064\u2066-\u2069]/g,

    /**
     * Normalize text for comparison (non-strict mode).
     * @param {string} text
     * @returns {string}
     */
    normalize(text) {
        if (!text) return '';
        let result = text;

        // Remove invisible characters
        result = result.replace(this.invisibleChars, '');

        // Unify line breaks
        result = result.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

        // Full-width to half-width
        result = result.split('').map(ch => this.fullToHalf[ch] || ch).join('');

        // Normalize whitespace: collapse multiple spaces/tabs into one
        result = result.replace(/[ \t]+/g, ' ');

        // Trim spaces around CJK characters
        result = result.replace(/ ?([\u4e00-\u9fff\u3000-\u303f]) ?/g, '$1');

        // Normalize consecutive punctuation
        result = result.replace(/。{2,}/g, '……');
        result = result.replace(/\.{3,}/g, '…');
        result = result.replace(/…{2,}/g, '……');

        return result;
    },

    /**
     * Check if two strings are semantically equivalent after normalization.
     * @param {string} a
     * @param {string} b
     * @returns {boolean}
     */
    areEquivalent(a, b) {
        return this.normalize(a) === this.normalize(b);
    }
};
