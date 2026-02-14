/**
 * OOXML Reconciliation Pipeline - Numbering Service
 * 
 * Manages list numbering to ensure continuation and consistent formatting.
 */

import { NS_W, NumberFormat, NumberSuffix } from '../core/types.js';

export class NumberingService {
    constructor() {
        this.contextMap = new Map(); // Cache for numIds found in the document
        this.nextNumId = 1000;
        this.customConfigs = []; // Track custom configs needed for current run
    }

    /**
     * Preserves an existing numId for a given format signature.
     * 
     * @param {string} signature - Format signature (e.g., 'bullet' or 'decimal')
     * @param {string} numId - Existing numId from Word
     */
    registerExistingNumId(signature, numId) {
        this.contextMap.set(signature, numId);
    }

    /**
     * Resolves the best numId to use for a requested list format.
     * 
     * @param {Object} formatConfig - Requested format (type, depth)
     * @param {Object} existingContext - Context from adjacent paragraph
     * @returns {string} The numId to use
     */
    getOrCreateNumId(formatConfig, existingContext = null) {
        const requestedType = formatConfig.type || NumberFormat.BULLET;

        // Priority 1: Use existing context if it matches the requested type
        if (existingContext && existingContext.numId) {
            if (existingContext.type === requestedType || existingContext.type === 'unknown') {
                return existingContext.numId;
            }
        }

        // Priority 2: Use cached numId for this format
        if (this.contextMap.has(requestedType)) {
            return this.contextMap.get(requestedType);
        }

        // Priority 3: Special Handling for Outline (recursive 1.1.1)
        if (requestedType === NumberFormat.OUTLINE) {
            // Use our new predefined Outline scheme (numId 3)
            return '3';
        }

        // Priority 4: Fallback based on type
        if (requestedType === NumberFormat.DECIMAL) return '2';
        if (requestedType === NumberFormat.BULLET) return '1';

        // Priority 5: Handle specialized Alpha/Roman at Level 0 (or if context doesn't match)
        // If we are at Level 0 and want something other than Decimal, we need a custom config
        const ilvl = existingContext ? parseInt(existingContext.ilvl || '0') : 0;
        if (ilvl === 0 && requestedType !== NumberFormat.DECIMAL && requestedType !== NumberFormat.BULLET) {
            // Find or create a custom config for this specific format
            const signature = `custom_${requestedType}`;
            if (this.contextMap.has(signature)) {
                return this.contextMap.get(signature);
            }

            const newNumId = String(this.nextNumId++);
            this.customConfigs.push({
                numId: newNumId,
                levels: [
                    { format: requestedType, suffix: formatConfig.suffix || NumberSuffix.PERIOD }
                ]
            });
            this.contextMap.set(signature, newNumId);
            return newNumId;
        }

        // Default fallback to NumId 2 (Legal/Nested)
        return '2';
    }

    /**
     * Detects numbering format from a string marker (e.g. "1.", "(a)", "i.", "1.1.")
     * 
     * @param {string} marker - The marker text
     * @returns {Object} { format, suffix, depth }
     */
    detectNumberingFormat(marker) {
        const m = (marker || '').trim();
        if (!m) return { format: NumberFormat.BULLET, suffix: NumberSuffix.NONE, depth: 0 };

        // Bullet
        if (/^[-*•]$/.test(m)) {
            return { format: NumberFormat.BULLET, suffix: NumberSuffix.NONE, depth: 0 };
        }

        // Hierarchical outline: 1.1.2 or 4.1.2.3
        const outlineMatch = m.match(/^(\d+(?:\.\d+)+)\.?$/);
        if (outlineMatch) {
            const depth = outlineMatch[1].split('.').length - 1;
            return { format: NumberFormat.OUTLINE, suffix: NumberSuffix.PERIOD, depth };
        }

        // Parenthesized formats: (a), (i), (1)
        if (/^\([a-z]\)$/.test(m)) {
            return { format: NumberFormat.LOWER_ALPHA, suffix: NumberSuffix.PAREN_BOTH, depth: 0 };
        }
        if (/^\([ivxlc]+\)$/i.test(m)) {
            const isLower = m === m.toLowerCase();
            return { format: isLower ? NumberFormat.LOWER_ROMAN : NumberFormat.UPPER_ROMAN, suffix: NumberSuffix.PAREN_BOTH, depth: 0 };
        }
        if (/^\(\d+\)$/.test(m)) {
            return { format: NumberFormat.DECIMAL, suffix: NumberSuffix.PAREN_BOTH, depth: 0 };
        }

        // Standard formats with period: 1., a., A., i., I.
        if (/^\d+\.$/.test(m)) {
            return { format: NumberFormat.DECIMAL, suffix: NumberSuffix.PERIOD, depth: 0 };
        }
        if (/^[a-z]\.$/.test(m)) {
            return { format: NumberFormat.LOWER_ALPHA, suffix: NumberSuffix.PERIOD, depth: 0 };
        }
        if (/^[A-Z]\.$/.test(m)) {
            return { format: NumberFormat.UPPER_ALPHA, suffix: NumberSuffix.PERIOD, depth: 0 };
        }
        if (/^[ivxlc]+\.$/i.test(m)) {
            const isLower = m === m.toLowerCase();
            return { format: isLower ? NumberFormat.LOWER_ROMAN : NumberFormat.UPPER_ROMAN, suffix: NumberSuffix.PERIOD, depth: 0 };
        }

        // Default to decimal if it looks like a number
        if (/^\d+/.test(m)) return { format: NumberFormat.DECIMAL, suffix: NumberSuffix.PERIOD, depth: 0 };

        return { format: NumberFormat.BULLET, suffix: NumberSuffix.NONE, depth: 0 };
    }

    /**
     * Maps internal format to OOXML numFmt string
     */
    formatToOoxmlNumFmt(format) {
        const map = {
            [NumberFormat.DECIMAL]: 'decimal',
            [NumberFormat.LOWER_ALPHA]: 'lowerLetter',
            [NumberFormat.UPPER_ALPHA]: 'upperLetter',
            [NumberFormat.LOWER_ROMAN]: 'lowerRoman',
            [NumberFormat.UPPER_ROMAN]: 'upperRoman',
            [NumberFormat.BULLET]: 'bullet',
            [NumberFormat.OUTLINE]: 'decimal'
        };
        return map[format] || 'decimal';
    }

    /**
     * Maps levels and suffix to OOXML lvlText
     */
    suffixToOoxmlLevelText(format, suffix, ilvl = 0) {
        if (format === NumberFormat.BULLET) return '•';

        // Placeholder for current level
        const num = `%${ilvl + 1}`;

        if (format === NumberFormat.OUTLINE) {
            // Outline %1.%2.%3.
            return Array(ilvl + 1).fill(0).map((_, i) => `%${i + 1}`).join('.') + '.';
        }

        switch (suffix) {
            case NumberSuffix.PERIOD: return `${num}.`;
            case NumberSuffix.PAREN_RIGHT: return `${num})`;
            case NumberSuffix.PAREN_BOTH: return `(${num})`;
            default: return num;
        }
    }

    /**
     * Builds paragraph properties for a list item.
     * 
     * @param {string} numId - The numId
     * @param {number} ilvl - Indentation level
     * @returns {string} Serialized w:pPr XML
     */
    buildListPPr(numId, ilvl, options = {}) {
        const includeListParagraphStyle = options.includeListParagraphStyle === true;
        const styleXml = includeListParagraphStyle ? '\n                <w:pStyle w:val="ListParagraph"/>' : '';
        return `
            <w:pPr>${styleXml}
                <w:numPr>
                    <w:ilvl w:val="${ilvl}"/>
                    <w:numId w:val="${numId}"/>
                </w:numPr>
            </w:pPr>
        `;
    }

    /**
     * Generates a full w:numbering XML block including custom legal schemes.
     * 
     * @param {Array} externalConfigs - Optional array of { numId, levels: [{ format, suffix }] }
     * @returns {string} w:numbering XML
     */
    generateNumberingXml(externalConfigs = []) {
        // Merge internal customConfigs with external ones
        const allCustomConfigs = [...this.customConfigs, ...externalConfigs];

        // Default Bullet (numId 1)
        let abstractNum0 = `
        <w:abstractNum w:abstractNumId="0">
            <w:multiLevelType w:val="hybridMultilevel"/>
            ${[0, 1, 2, 3, 4, 5, 6, 7, 8].map(lvl => `
            <w:lvl w:ilvl="${lvl}">
                <w:start w:val="1"/>
                <w:numFmt w:val="${lvl % 3 === 0 ? 'bullet' : lvl % 3 === 1 ? 'circle' : 'square'}"/>
                <w:lvlText w:val="${lvl % 3 === 0 ? '•' : lvl % 3 === 1 ? '○' : '■'}"/>
                <w:lvlJc w:val="left"/>
                <w:pPr><w:ind w:left="${720 * (lvl + 1)}" w:hanging="360"/></w:pPr>
            </w:lvl>`).join('')}
        </w:abstractNum>`;

        // Default Numbered (numId 2) - US/Legal Style
        // Level 0: 1.
        // Level 1: (a)
        // Level 2: (i)
        // Level 3: (1)
        // Level 4: (a) - repeating with different indent...
        let abstractNum1 = `
        <w:abstractNum w:abstractNumId="1">
            <w:multiLevelType w:val="multilevel"/>
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="(%3)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="(%4)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%5."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="3600" w:hanging="360"/></w:pPr></w:lvl>
        </w:abstractNum>`;

        // Outline Numbered (numId 3) - 1 / 1.1 / 1.1.1
        let abstractNum2 = `
        <w:abstractNum w:abstractNumId="2">
            <w:multiLevelType w:val="multilevel"/>
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3.%4"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3.%4.%5"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="3600" w:hanging="360"/></w:pPr></w:lvl>
        </w:abstractNum>`;

        // Handle custom configurations (e.g. outline 1.1.1)
        let customAbstractNums = '';
        let customNums = '';

        allCustomConfigs.forEach((config, idx) => {
            const absId = 10 + idx;
            customAbstractNums += `
            <w:abstractNum w:abstractNumId="${absId}">
                <w:multiLevelType w:val="multilevel"/>
                ${config.levels.map((l, ilvl) => `
                <w:lvl w:ilvl="${ilvl}">
                    <w:start w:val="1"/>
                    <w:numFmt w:val="${this.formatToOoxmlNumFmt(l.format)}"/>
                    <w:lvlText w:val="${this.suffixToOoxmlLevelText(l.format, l.suffix, ilvl)}"/>
                    <w:lvlJc w:val="left"/>
                    <w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/></w:pPr>
                </w:lvl>`).join('')}
            </w:abstractNum>`;

            customNums += `
            <w:num w:numId="${config.numId}">
                <w:abstractNumId w:val="${absId}"/>
            </w:num>`;
        });

        return `
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            ${abstractNum0}
            ${abstractNum1}
            ${abstractNum2}
            ${customAbstractNums}
            <w:num w:numId="1">
                <w:abstractNumId w:val="0"/>
            </w:num>
            <w:num w:numId="2">
                <w:abstractNumId w:val="1"/>
            </w:num>
            <w:num w:numId="3">
                <w:abstractNumId w:val="2"/>
            </w:num>
            ${customNums}
        </w:numbering>`;
    }
}
