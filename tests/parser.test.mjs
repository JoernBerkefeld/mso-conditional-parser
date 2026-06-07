import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
    getConditionFix,
    parseMsoComment,
    parseMsoEndComment,
    isMsoComment,
    translateCondition,
    validateCondition,
} from '../src/index.js';

// ── parseMsoComment — downlevel-hidden openers ──────────────────────────────

describe('parseMsoComment — downlevel-hidden openers', () => {
    it('parses bare mso', () => {
        const result = parseMsoComment('<!--[if mso]>');
        assert.equal(result.type, 'downlevel-hidden');
        assert.equal(result.condition, 'mso');
        assert.equal(result.isValid, true);
    });

    it('parses gte mso 16', () => {
        const result = parseMsoComment('<!--[if gte mso 16]>');
        assert.equal(result.type, 'downlevel-hidden');
        assert.equal(result.condition, 'gte mso 16');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('2016'));
    });

    it('parses gt mso 15', () => {
        const result = parseMsoComment('<!--[if gt mso 15]>');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('2013'));
    });

    it('parses lte mso 15', () => {
        const result = parseMsoComment('<!--[if lte mso 15]>');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('2013'));
    });

    it('parses lt mso 16', () => {
        const result = parseMsoComment('<!--[if lt mso 16]>');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('2016'));
    });

    it('parses exact version mso 12', () => {
        const result = parseMsoComment('<!--[if mso 12]>');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('2007'));
    });

    it('parses !mso', () => {
        const result = parseMsoComment('<!--[if !mso]>');
        assert.equal(result.type, 'downlevel-hidden');
        assert.equal(result.condition, '!mso');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('non-Outlook'));
    });

    it('handles whitespace around input', () => {
        const result = parseMsoComment('  <!--[if mso]>  ');
        assert.equal(result.type, 'downlevel-hidden');
        assert.equal(result.isValid, true);
    });
});

// ── parseMsoComment — downlevel-revealed openers ────────────────────────────

describe('parseMsoComment — downlevel-revealed openers', () => {
    it('parses <!--[if !mso]><!--> (standard revealed)', () => {
        const result = parseMsoComment('<!--[if !mso]><!--');
        assert.equal(result.type, 'downlevel-revealed');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('non-Outlook'));
    });

    it('parses <![if !mso]> (non-standard revealed)', () => {
        const result = parseMsoComment('<![if !mso]>');
        assert.equal(result.type, 'downlevel-revealed');
        assert.equal(result.isValid, true);
    });

    it('parses <![if mso]> (non-standard any Outlook)', () => {
        const result = parseMsoComment('<![if mso]>');
        assert.equal(result.type, 'downlevel-revealed');
        assert.equal(result.isValid, true);
    });
});

// ── parseMsoComment — version map completeness ──────────────────────────────

describe('parseMsoComment — all valid versions', () => {
    const validVersions = [9, 10, 11, 12, 14, 15, 16];
    for (const v of validVersions) {
        it(`accepts mso ${v}`, () => {
            const result = parseMsoComment(`<!--[if mso ${v}]>`);
            assert.equal(result.isValid, true, `mso ${v} should be valid`);
        });
    }

    it('rejects mso 13 (does not exist)', () => {
        const result = parseMsoComment('<!--[if mso 13]>');
        assert.equal(result.isValid, false);
        assert.ok(result.error.includes('13'));
    });

    it('rejects mso 17 (does not exist)', () => {
        const result = parseMsoComment('<!--[if mso 17]>');
        assert.equal(result.isValid, false);
    });
});

// ── parseMsoComment — validation / error detection ──────────────────────────

describe('parseMsoComment — validation', () => {
    it('flags "mos" typo', () => {
        const result = parseMsoComment('<!--[if mos]>');
        assert.equal(result.isValid, false);
        assert.ok(result.error.toLowerCase().includes('mso'));
        assert.equal(result.conditionFix, 'mso');
    });

    it('flags operator without version number', () => {
        const result = parseMsoComment('<!--[if gte mso]>');
        assert.equal(result.isValid, false);
        assert.ok(result.error.includes('version number'));
    });

    it('flags unknown operator', () => {
        const result = parseMsoComment('<!--[if foobar mso 16]>');
        assert.equal(result.isValid, false);
    });

    it('returns null for non-MSO input', () => {
        assert.equal(parseMsoComment('<!-- regular comment -->'), null);
        assert.equal(parseMsoComment('<div class="foo">'), null);
        assert.equal(parseMsoComment(''), null);
    });
});

// ── parseMsoComment — compound conditions ───────────────────────────────────

describe('parseMsoComment — compound conditions', () => {
    it('parses AND compound condition', () => {
        const result = parseMsoComment('<!--[if (gte mso 9)&(lte mso 15)]>');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('AND'));
    });

    it('parses OR compound condition', () => {
        const result = parseMsoComment('<!--[if (mso 12)|(mso 14)]>');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('OR'));
    });
});

// ── parseMsoEndComment ───────────────────────────────────────────────────────

describe('parseMsoEndComment', () => {
    it('parses downlevel-hidden closer <![endif]-->', () => {
        const result = parseMsoEndComment('<![endif]-->');
        assert.equal(result.type, 'downlevel-hidden-end');
        assert.equal(result.isClosing, true);
    });

    it('parses downlevel-revealed closer <!--<![endif]-->', () => {
        const result = parseMsoEndComment('<!--<![endif]-->');
        assert.equal(result.type, 'downlevel-hidden-end');
        assert.equal(result.isClosing, true);
    });

    it('parses downlevel-revealed non-standard closer <![endif]>', () => {
        const result = parseMsoEndComment('<![endif]>');
        assert.equal(result.type, 'downlevel-revealed-end');
        assert.equal(result.isClosing, true);
    });

    it('returns null for non-closer input', () => {
        assert.equal(parseMsoEndComment('<!--[if mso]>'), null);
        assert.equal(parseMsoEndComment('<!-- comment -->'), null);
        assert.equal(parseMsoEndComment(''), null);
    });
});

// ── isMsoComment ────────────────────────────────────────────────────────────

describe('isMsoComment', () => {
    it('returns true for an opener', () => {
        assert.equal(isMsoComment('<!--[if mso]>'), true);
    });

    it('returns true for a closer', () => {
        assert.equal(isMsoComment('<![endif]-->'), true);
    });

    it('returns false for a regular HTML comment', () => {
        assert.equal(isMsoComment('<!-- hello -->'), false);
    });

    it('returns false for empty string', () => {
        assert.equal(isMsoComment(''), false);
    });
});

// ── translateCondition ───────────────────────────────────────────────────────

describe('translateCondition', () => {
    it('translates mso', () => {
        assert.equal(translateCondition('mso'), 'all Outlook versions');
    });

    it('translates !mso', () => {
        assert.ok(translateCondition('!mso').includes('non-Outlook'));
    });

    it('translates gte mso 16', () => {
        const result = translateCondition('gte mso 16');
        assert.ok(result.includes('2016'));
        assert.ok(result.includes('newer'));
    });

    it('translates exact version mso 12', () => {
        assert.ok(translateCondition('mso 12').includes('2007'));
    });

    it('translates compound AND condition', () => {
        const result = translateCondition('(gte mso 9)&(lte mso 15)');
        assert.ok(result.includes('AND'));
    });
});

// ── translateCondition — additional operator coverage ────────────────────────

describe('translateCondition — additional operators and edge cases', () => {
    it('translates gt mso 15 (newer than Outlook 2013)', () => {
        const result = translateCondition('gt mso 15');
        assert.ok(result.includes('newer than'));
        assert.ok(result.includes('2013'));
    });

    it('translates lte mso 16 (Outlook 2016 and older)', () => {
        const result = translateCondition('lte mso 16');
        assert.ok(result.includes('2016'));
        assert.ok(result.includes('older'));
    });

    it('translates lt mso 16 (older than Outlook 2016)', () => {
        const result = translateCondition('lt mso 16');
        assert.ok(result.includes('older than'));
        assert.ok(result.includes('2016'));
    });

    it('translates eq mso 12 (returns exact name)', () => {
        const result = translateCondition('eq mso 12');
        assert.ok(result.includes('2007'));
    });

    it('translates !mso 12 (negated exact version)', () => {
        const result = translateCondition('!mso 12');
        assert.ok(result.includes('not'));
        assert.ok(result.includes('2007'));
    });

    it('falls back to "version N and newer" for unmapped version with gte', () => {
        const result = translateCondition('gte mso 13');
        assert.ok(result.includes('version 13'));
        assert.ok(result.includes('newer'));
    });

    it('translates compound OR condition', () => {
        const result = translateCondition('(mso 12)|(mso 14)');
        assert.ok(result.includes('OR'));
        assert.ok(result.includes('2007'));
        assert.ok(result.includes('2010'));
    });

    it('returns the raw string for completely unrecognized conditions (fallback)', () => {
        const result = translateCondition('garbage condition');
        assert.equal(result, 'garbage condition');
    });
});

// ── parseMsoComment — additional coverage ────────────────────────────────────

describe('parseMsoComment — additional coverage', () => {
    it('parses eq operator', () => {
        const result = parseMsoComment('<!--[if eq mso 16]>');
        assert.equal(result.isValid, true);
        assert.equal(result.condition, 'eq mso 16');
    });

    it('parses negated exact version !mso 12', () => {
        const result = parseMsoComment('<!--[if !mso 12]>');
        assert.equal(result.isValid, true);
        assert.equal(result.condition, '!mso 12');
        assert.ok(result.translation.includes('2007'));
    });

    it('rejects negated invalid version !mso 13', () => {
        const result = parseMsoComment('<!--[if !mso 13]>');
        assert.equal(result.isValid, false);
        assert.ok(result.error.includes('13'));
    });

    it('rejects compound condition with invalid sub-part', () => {
        const result = parseMsoComment('<!--[if (gte mso 9)&(mso 13)]>');
        assert.equal(result.isValid, false);
        assert.ok(result.error.includes('13'));
    });

    it('handles case-insensitive operator and keyword', () => {
        const result = parseMsoComment('<!--[if GTE MSO 16]>');
        assert.equal(result.isValid, true);
        assert.ok(result.translation.includes('2016'));
    });

    it('<!--[if !mso]><!-- reports type as downlevel-revealed', () => {
        const result = parseMsoComment('<!--[if !mso]><!--');
        assert.equal(result.type, 'downlevel-revealed');
    });
});

// ── parseMsoEndComment — type field completeness ─────────────────────────────

describe('parseMsoEndComment — type field', () => {
    it('<![endif]> has type downlevel-revealed-end', () => {
        const result = parseMsoEndComment('<![endif]>');
        assert.equal(result.type, 'downlevel-revealed-end');
    });
});

// ── isMsoComment — non-standard revealed form ─────────────────────────────────

describe('validateCondition and IE support', () => {
    it('accepts legacy IE conditions', () => {
        assert.equal(validateCondition('IE'), null);
        assert.equal(validateCondition('!IE'), null);
        assert.equal(validateCondition('gte IE 7'), null);
    });

    it('returns a deterministic fix for mos typo', () => {
        assert.equal(getConditionFix('mos'), 'mso');
        assert.equal(getConditionFix('gte mos 16'), 'gte mso 16');
    });

    it('parses <!--[if IE]> as valid', () => {
        const result = parseMsoComment('<!--[if IE]>');
        assert.equal(result.isValid, true);
        assert.equal(result.condition, 'IE');
    });
});

describe('isMsoComment — non-standard forms', () => {
    it('returns true for non-standard revealed opener <![if mso]>', () => {
        assert.equal(isMsoComment('<![if mso]>'), true);
    });
});
