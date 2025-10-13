# Word.Interfaces.TemplateCollectionLoadOptions interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Contains a collection of Word.Template objects that represent all the templates that are currently available. This collection includes open templates, templates attached to open documents, and global templates loaded in the Templates and Add-ins dialog box. To learn how to access this dialog in the Word UI, see https://support.microsoft.com/office/2479fe53-f849-4394-88bb-2a6e2a39479d.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- farEastLineBreakLanguage — For EACH ITEM in the collection: Specifies the East Asian language to use when breaking lines of text in the document or template.
- farEastLineBreakLevel — For EACH ITEM in the collection: Specifies the line break control level for the document.
- fullName — For EACH ITEM in the collection: Returns the name of the template, including the drive or Web path.
- hasNoProofing — For EACH ITEM in the collection: Specifies whether the spelling and grammar checker ignores documents based on this template.
- justificationMode — For EACH ITEM in the collection: Specifies the character spacing adjustment for the template.
- kerningByAlgorithm — For EACH ITEM in the collection: Specifies if Microsoft Word kerns half-width Latin characters and punctuation marks in the document.
- languageId — For EACH ITEM in the collection: Specifies a LanguageId value that represents the language in the template.
- languageIdFarEast — For EACH ITEM in the collection: Specifies an East Asian language for the language in the template.
- name — For EACH ITEM in the collection: Returns only the name of the document template (excluding any path or other location information).
- noLineBreakAfter — For EACH ITEM in the collection: Specifies the kinsoku characters after which Microsoft Word will not break a line.
- noLineBreakBefore — For EACH ITEM in the collection: Specifies the kinsoku characters before which Microsoft Word will not break a line.
- path — For EACH ITEM in the collection: Returns the path to the document template.
- saved — For EACH ITEM in the collection: Specifies true if the template has not changed since it was last saved, false if Microsoft Word displays a prompt to save changes when the document is closed.
- type — For EACH ITEM in the collection: Returns the template type.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value
- boolean

### farEastLineBreakLanguage

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the East Asian language to use when breaking lines of text in the document or template.

```typescript
farEastLineBreakLanguage?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### farEastLineBreakLevel

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the line break control level for the document.

```typescript
farEastLineBreakLevel?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### fullName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the name of the template, including the drive or Web path.

```typescript
fullName?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### hasNoProofing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the spelling and grammar checker ignores documents based on this template.

```typescript
hasNoProofing?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### justificationMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the character spacing adjustment for the template.

```typescript
justificationMode?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### kerningByAlgorithm

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if Microsoft Word kerns half-width Latin characters and punctuation marks in the document.

```typescript
kerningByAlgorithm?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### languageId

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies a LanguageId value that represents the language in the template.

```typescript
languageId?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### languageIdFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies an East Asian language for the language in the template.

```typescript
languageIdFarEast?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns only the name of the document template (excluding any path or other location information).

```typescript
name?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### noLineBreakAfter

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the kinsoku characters after which Microsoft Word will not break a line.

```typescript
noLineBreakAfter?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### noLineBreakBefore

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the kinsoku characters before which Microsoft Word will not break a line.

```typescript
noLineBreakBefore?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### path

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the path to the document template.

```typescript
path?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### saved

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies true if the template has not changed since it was last saved, false if Microsoft Word displays a prompt to save changes when the document is closed.

```typescript
saved?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the template type.

```typescript
type?: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]