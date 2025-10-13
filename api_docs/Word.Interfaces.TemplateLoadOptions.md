# Word.Interfaces.TemplateLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a document template.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- farEastLineBreakLanguage — Specifies the East Asian language to use when breaking lines of text in the document or template.
- farEastLineBreakLevel — Specifies the line break control level for the document.
- fullName — Returns the name of the template, including the drive or Web path.
- hasNoProofing — Specifies whether the spelling and grammar checker ignores documents based on this template.
- justificationMode — Specifies the character spacing adjustment for the template.
- kerningByAlgorithm — Specifies if Microsoft Word kerns half-width Latin characters and punctuation marks in the document.
- languageId — Specifies a LanguageId value that represents the language in the template.
- languageIdFarEast — Specifies an East Asian language for the language in the template.
- name — Returns only the name of the document template (excluding any path or other location information).
- noLineBreakAfter — Specifies the kinsoku characters after which Microsoft Word will not break a line.
- noLineBreakBefore — Specifies the kinsoku characters before which Microsoft Word will not break a line.
- path — Returns the path to the document template.
- saved — Specifies true if the template has not changed since it was last saved, false if Microsoft Word displays a prompt to save changes when the document is closed.
- type — Returns the template type.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property value: boolean

---

### farEastLineBreakLanguage

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the East Asian language to use when breaking lines of text in the document or template.

```typescript
farEastLineBreakLanguage?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### farEastLineBreakLevel

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line break control level for the document.

```typescript
farEastLineBreakLevel?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fullName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the name of the template, including the drive or Web path.

```typescript
fullName?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hasNoProofing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the spelling and grammar checker ignores documents based on this template.

```typescript
hasNoProofing?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### justificationMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the character spacing adjustment for the template.

```typescript
justificationMode?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### kerningByAlgorithm

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if Microsoft Word kerns half-width Latin characters and punctuation marks in the document.

```typescript
kerningByAlgorithm?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### languageId

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LanguageId value that represents the language in the template.

```typescript
languageId?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### languageIdFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an East Asian language for the language in the template.

```typescript
languageIdFarEast?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns only the name of the document template (excluding any path or other location information).

```typescript
name?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### noLineBreakAfter

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the kinsoku characters after which Microsoft Word will not break a line.

```typescript
noLineBreakAfter?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### noLineBreakBefore

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the kinsoku characters before which Microsoft Word will not break a line.

```typescript
noLineBreakBefore?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### path

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the path to the document template.

```typescript
path?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### saved

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies true if the template has not changed since it was last saved, false if Microsoft Word displays a prompt to save changes when the document is closed.

```typescript
saved?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the template type.

```typescript
type?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)