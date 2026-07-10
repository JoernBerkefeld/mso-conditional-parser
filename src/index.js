/**
 * mso-conditional-parser
 *
 * Parses and translates Microsoft Outlook MSO conditional comments
 * (downlevel-hidden and downlevel-revealed variants) into structured objects
 * with human-readable English translations.
 *
 * Supports all real-world comment patterns used in HTML email development.
 */

/**
 * Maps MSO version numbers to Outlook product names.
 * Note: there is no MSO version 13 — Outlook jumps from 12 (2007) to 14 (2010).
 *
 * @type {Record<string, string>}
 */
const VERSION_MAP = {
    9: 'Outlook 2000',
    10: 'Outlook 2002',
    11: 'Outlook 2003',
    12: 'Outlook 2007',
    14: 'Outlook 2010',
    15: 'Outlook 2013',
    16: 'Outlook 2016, 2019, and 365',
};

/** Set of valid MSO version numbers as strings. */
const VALID_VERSIONS = new Set(Object.keys(VERSION_MAP));

/** Set of valid comparison operators. */
const VALID_OPERATORS = new Set(['gte', 'gt', 'lte', 'lt', 'eq']);

/**
 * Regex to match a downlevel-hidden opener:
 *   <!--[if condition]>
 *   <!--[if condition]><!--   (downlevel-revealed standard form)
 */
const HIDDEN_OPENER_RE = /^<!--\[if\s+([\s\S]+?)\]>(<!--)?$/;

/**
 * Regex to match a downlevel-revealed (non-standard) opener:
 *   <![if condition]>
 */
const REVEALED_OPENER_RE = /^<!\[if\s+([\s\S]+?)\]>$/;

/**
 * Regex to match closers:
 *   <![endif]-->   (closes a downlevel-hidden opener)
 *   <!--<![endif]--> (closes <!--[if !mso]><!--> opener)
 *   <![endif]>     (closes a <![if …]> opener)
 */
const HIDDEN_CLOSER_RE = /^(?:<!--)?<!\[endif\]-->$/;
const REVEALED_CLOSER_RE = /^<!\[endif\]>$/;

/**
 * Translates a raw MSO condition string into human-readable English.
 *
 * @param {string} condition - The raw condition string (e.g. "gte mso 16").
 * @returns {string} Human-readable translation.
 */
export function translateCondition(condition) {
    const trimmed = condition.trim();

    if (trimmed === 'mso') {
        return 'all Outlook versions';
    }

    if (trimmed === '!mso') {
        return 'non-Outlook email clients';
    }

    if (trimmed === 'IE') {
        return 'Internet Explorer';
    }

    if (trimmed === '!IE') {
        return 'non-Internet Explorer clients';
    }

    // Compound conditions with & or |
    if (trimmed.includes('&') || trimmed.includes('|')) {
        const separator = trimmed.includes('&') ? ' AND ' : ' OR ';
        const parts = trimmed
            .split(/[&|]/)
            .map((part) => translateCondition(part.replaceAll(/[()]/g, '').trim()));
        return parts.join(separator);
    }

    // operator mso version — e.g. "gte mso 16"
    const operatorVersionMatch = trimmed.match(/^(gte|gt|lte|lt|eq)\s+mso\s+(\d+)$/i);
    if (operatorVersionMatch) {
        const op = operatorVersionMatch[1].toLowerCase();
        const version = operatorVersionMatch[2];
        const name = VERSION_MAP[version];
        const versionLabel = name ?? `version ${version}`;
        switch (op) {
            case 'gte': {
                return `${versionLabel} and newer`;
            }
            case 'gt': {
                return `newer than ${versionLabel}`;
            }
            case 'lte': {
                return `${versionLabel} and older`;
            }
            case 'lt': {
                return `older than ${versionLabel}`;
            }
            case 'eq': {
                return versionLabel;
            }
        }
    }

    // exact version — e.g. "mso 16"
    const exactVersionMatch = trimmed.match(/^mso\s+(\d+)$/i);
    if (exactVersionMatch) {
        const version = exactVersionMatch[1];
        return VERSION_MAP[version] ?? `Outlook version ${version}`;
    }

    // negation — e.g. "!mso 16"
    const negationMatch = trimmed.match(/^!\s*mso(?:\s+(\d+))?$/i);
    if (negationMatch) {
        if (negationMatch[1]) {
            const name = VERSION_MAP[negationMatch[1]] ?? `Outlook version ${negationMatch[1]}`;
            return `not ${name}`;
        }

        return 'non-Outlook email clients';
    }

    return condition.trim();
}

/**
 * Assesses a conditional comment expression.
 *
 * @param {string} condition - Raw condition text from an opener.
 * @returns {{ error: string, fix?: string } | null} Assessment result, or null when valid.
 */
function assessCondition(condition) {
    const trimmed = condition.trim();

    if (/\bmos\b/i.test(trimmed)) {
        return {
            error: "Use the keyword 'mso', not 'mos'",
            fix: trimmed.replaceAll(/\bmos\b/gi, 'mso'),
        };
    }

    // Compound: validate each part individually
    if (trimmed.includes('&') || trimmed.includes('|')) {
        const parts = trimmed.split(/[&|]/).map((p) => p.replaceAll(/[()]/g, '').trim());
        for (const part of parts) {
            const partResult = assessCondition(part);
            if (partResult) {
                return partResult;
            }
        }

        return null;
    }

    if (['mso', '!mso', 'IE', '!IE'].includes(trimmed)) {
        return null;
    }

    const ieOperatorMatch = trimmed.match(/^(gte|gt|lte|lt|eq)\s+IE\s+\d+$/i);
    if (ieOperatorMatch) {
        return null;
    }

    if (/^IE\s+\d+$/i.test(trimmed)) {
        return null;
    }

    // operator mso version
    const operatorVersionMatch = trimmed.match(/^(gte|gt|lte|lt|eq)\s+mso\s+(\d+)$/i);
    if (operatorVersionMatch) {
        const version = operatorVersionMatch[2];
        if (!VALID_VERSIONS.has(version)) {
            return {
                error: `Unknown MSO version ${version}. Use one of: ${[...VALID_VERSIONS].join(', ')}`,
            };
        }

        return null;
    }

    const operatorNoVersionMatch = trimmed.match(/^(gte|gt|lte|lt|eq)\s+mso\s*$/i);
    if (operatorNoVersionMatch) {
        return {
            error: `Operator '${operatorNoVersionMatch[1]}' needs a version number (for example '${operatorNoVersionMatch[1]} mso 16')`,
        };
    }

    const unknownOpMatch = trimmed.match(/^(\w+)\s+mso(?:\s+(\d+))?$/i);
    if (unknownOpMatch && !VALID_OPERATORS.has(unknownOpMatch[1].toLowerCase())) {
        return {
            error: `Unknown operator '${unknownOpMatch[1]}'. Use one of: ${[...VALID_OPERATORS].join(', ')}`,
        };
    }

    const exactVersionMatch = trimmed.match(/^mso\s+(\d+)$/i);
    if (exactVersionMatch) {
        const version = exactVersionMatch[1];
        if (!VALID_VERSIONS.has(version)) {
            return {
                error: `Unknown MSO version ${version}. Use one of: ${[...VALID_VERSIONS].join(', ')}`,
            };
        }

        return null;
    }

    const negationMatch = trimmed.match(/^!\s*mso\s+(\d+)$/i);
    if (negationMatch) {
        const version = negationMatch[1];
        if (!VALID_VERSIONS.has(version)) {
            return {
                error: `Unknown MSO version ${version}. Use one of: ${[...VALID_VERSIONS].join(', ')}`,
            };
        }

        return null;
    }

    return { error: `Unrecognized conditional expression: '${trimmed}'` };
}

/**
 * Validates an MSO condition string and returns an error message if invalid.
 *
 * @param {string} condition - The raw condition string extracted from the comment.
 * @returns {string|null} Error message or null if valid.
 */
export function validateCondition(condition) {
    const result = assessCondition(condition);
    return result?.error ?? null;
}

/**
 * Returns a replacement condition string when a deterministic fix exists.
 *
 * @param {string} condition - The raw condition string extracted from the comment.
 * @returns {string|null} Fixed condition text, or null when no safe fix is available.
 */
export function getConditionFix(condition) {
    return assessCondition(condition)?.fix ?? null;
}

/**
 * Parses an MSO conditional comment opener into a structured object.
 *
 * Handles:
 *   - `<!--[if mso]>`                     downlevel-hidden, any Outlook
 *   - `<!--[if gte mso 16]>`              downlevel-hidden, Outlook 2016+
 *   - `<!--[if !mso]><!-->               downlevel-revealed, standard form
 *   - `<![if !mso]>`                      downlevel-revealed, non-standard
 *
 * Returns `null` if the input is not a recognizable MSO opener.
 *
 * @param {string} string_ - Raw HTML comment string (trimmed).
 * @returns {{ type: string, condition: string, translation: string, isValid: boolean, error?: string } | null} Parsed opener object, or null if not an MSO opener.
 */
export function parseMsoComment(string_) {
    const trimmed = string_.trim();

    // Downlevel-revealed non-standard: <![if condition]>
    const revealedMatch = REVEALED_OPENER_RE.exec(trimmed);
    if (revealedMatch) {
        const condition = revealedMatch[1].trim();
        const assessment = assessCondition(condition);
        return {
            type: 'downlevel-revealed',
            condition,
            translation: translateCondition(condition),
            isValid: !assessment,
            ...(assessment && {
                error: assessment.error,
                ...(assessment.fix && { conditionFix: assessment.fix }),
            }),
        };
    }

    // Downlevel-hidden or downlevel-revealed standard: <!--[if condition]> or <!--[if condition]><!--
    const hiddenMatch = HIDDEN_OPENER_RE.exec(trimmed);
    if (hiddenMatch) {
        const condition = hiddenMatch[1].trim();
        const isRevealed = !!hiddenMatch[2]; // has trailing <!--
        const assessment = assessCondition(condition);
        return {
            type: isRevealed ? 'downlevel-revealed' : 'downlevel-hidden',
            condition,
            translation: translateCondition(condition),
            isValid: !assessment,
            ...(assessment && {
                error: assessment.error,
                ...(assessment.fix && { conditionFix: assessment.fix }),
            }),
        };
    }

    return null;
}

/**
 * Parses an MSO conditional comment closer into a structured object.
 *
 * Handles:
 *   - `<![endif]-->`      closes a downlevel-hidden opener
 *   - `<!--<![endif]-->` closes a `<!--[if !mso]><!-->`  opener
 *   - `<![endif]>`        closes a `<![if …]>` opener
 *
 * Returns `null` if the input is not a recognizable MSO closer.
 *
 * @param {string} string_ - Raw HTML comment string (trimmed).
 * @returns {{ type: string, isClosing: true } | null} Parsed closer object, or null if not an MSO closer.
 */
export function parseMsoEndComment(string_) {
    const trimmed = string_.trim();

    if (HIDDEN_CLOSER_RE.test(trimmed)) {
        return { type: 'downlevel-hidden-end', isClosing: true };
    }

    if (REVEALED_CLOSER_RE.test(trimmed)) {
        return { type: 'downlevel-revealed-end', isClosing: true };
    }

    return null;
}

/**
 * Returns true if the string is any MSO comment — opener or closer.
 *
 * @param {string} string_ - Raw HTML comment string.
 * @returns {boolean} True if the string is a recognized MSO comment opener or closer.
 */
export function isMsoComment(string_) {
    return parseMsoComment(string_) !== null || parseMsoEndComment(string_) !== null;
}
