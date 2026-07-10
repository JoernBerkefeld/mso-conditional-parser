# mso-conditional-parser

Parser and translator for **MSO (Outlook) conditional comments** in HTML email.

Parses all real-world MSO comment patterns, validates conditions against known Outlook versions and operators, and translates them into human-readable English.

Used as the core engine by [`eslint-plugin-mso-email`](https://github.com/JoernBerkefeld/eslint-plugin-mso-email) and the [`MSO Conditional Comments`](https://marketplace.visualstudio.com/items?itemName=joernberkefeld.mso-conditionals) VS Code extension.

## Installation

```bash
npm install mso-conditional-parser
```

Requires Node.js 18+. ES module only (`"type": "module"`).

## API

### `parseMsoComment(str)`

Parses an MSO conditional comment opener.

Returns a result object or `null` if the input is not a recognised MSO opener.

```js
import { parseMsoComment } from 'mso-conditional-parser';

parseMsoComment('<!--[if gte mso 16]>');
// {
//   type: 'downlevel-hidden',
//   condition: 'gte mso 16',
//   translation: 'Outlook 2016, 2019, and 365 and newer',
//   isValid: true
// }

parseMsoComment('<!--[if mos]>');
// {
//   type: 'downlevel-hidden',
//   condition: 'mos',
//   translation: 'mos',
//   isValid: false,
//   error: "Use the keyword 'mso', not 'mos'",
//   conditionFix: 'mso'
// }
```

**Supported opener patterns:**

| Pattern | Type |
|---|---|
| `<!--[if mso]>` | downlevel-hidden |
| `<!--[if gte mso 16]>` | downlevel-hidden |
| `<!--[if !mso]><!--` | downlevel-revealed |
| `<![if !mso]>` | downlevel-revealed (non-standard) |

**Result shape:**

| Field | Type | Description |
|---|---|---|
| `type` | `string` | `'downlevel-hidden'` or `'downlevel-revealed'` |
| `condition` | `string` | Raw condition string extracted from the comment |
| `translation` | `string` | Human-readable English translation |
| `isValid` | `boolean` | `true` if the condition is syntactically valid |
| `error` | `string` | Present when `isValid` is `false` |
| `conditionFix` | `string` | Optional replacement condition when a safe autofix exists |

---

### `validateCondition(condition)`

Returns an error message for an invalid condition string, or `null` when valid.

```js
import { validateCondition } from 'mso-conditional-parser';

validateCondition('mos'); // "Use the keyword 'mso', not 'mos'"
validateCondition('gte mso 16'); // null
```

---

### `getConditionFix(condition)`

Returns a replacement condition string when a deterministic fix exists (for example `mos` → `mso`), or `null` otherwise.

```js
import { getConditionFix } from 'mso-conditional-parser';

getConditionFix('mos'); // 'mso'
getConditionFix('gte mso 16'); // null
```

---

### `parseMsoEndComment(str)`

Parses an MSO conditional comment closer.

Returns `{ type, isClosing: true }` or `null` if not recognised.

```js
import { parseMsoEndComment } from 'mso-conditional-parser';

parseMsoEndComment('<![endif]-->');
// { type: 'downlevel-hidden-end', isClosing: true }

parseMsoEndComment('<![endif]>');
// { type: 'downlevel-revealed-end', isClosing: true }
```

---

### `translateCondition(condition)`

Translates a raw condition string into English without full parsing.

```js
import { translateCondition } from 'mso-conditional-parser';

translateCondition('gte mso 16');  // 'Outlook 2016, 2019, and 365 and newer'
translateCondition('!mso');        // 'non-Outlook email clients'
translateCondition('mso 12');      // 'Outlook 2007'
```

---

### `isMsoComment(str)`

Returns `true` if the string is any MSO comment — opener or closer.

```js
import { isMsoComment } from 'mso-conditional-parser';

isMsoComment('<!--[if mso]>');   // true
isMsoComment('<![endif]-->');    // true
isMsoComment('<p>hello</p>');    // false
```

## Supported conditions

| Syntax | Example | Meaning |
|---|---|---|
| `mso` | `<!--[if mso]>` | Any Outlook version |
| `!mso` | `<!--[if !mso]><!--` | Non-Outlook clients |
| `mso <version>` | `<!--[if mso 16]>` | Exact Outlook version |
| `<op> mso <version>` | `<!--[if gte mso 14]>` | Version comparison |
| `!mso <version>` | `<!--[if !mso 16]>` | Not this version |
| `IE` / `!IE` | `<!--[if IE]>` | Legacy Internet Explorer conditions |
| `<op> IE <version>` | `<!--[if gte IE 8]>` | IE version comparison |
| Compound `&` / `\|` | `<!--[if (gt mso 9)&(lte mso 11)]>` | AND / OR |

**Valid operators:** `gte`, `gt`, `lte`, `lt`, `eq`

**Valid Outlook versions:** 9 (2000), 10 (2002), 11 (2003), 12 (2007), 14 (2010), 15 (2013), 16 (2016/2019/365)

> Note: there is no MSO version 13 — Outlook skips from 12 (2007) to 14 (2010).

## License

MIT
