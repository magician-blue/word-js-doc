# Word.Interfaces.StyleCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Style](/en-us/javascript/api/word/word.style) objects.

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [$all](#all): Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- [automaticallyUpdate](#automaticallyupdate): For EACH ITEM in the collection: Specifies whether the style is automatically redefined based on the selection.
- [baseStyle](#basestyle): For EACH ITEM in the collection: Specifies the name of an existing style to use as the base formatting of another style.
- [borders](#borders): For EACH ITEM in the collection: Specifies a BorderCollection object that represents all the borders for the specified style.
- [builtIn](#builtin): For EACH ITEM in the collection: Gets whether the specified style is a built-in style.
- [description](#description): For EACH ITEM in the collection: Gets the description of the specified style.
- [font](#font): For EACH ITEM in the collection: Gets a font object that represents the character formatting of the specified style.
- [frame](#frame): For EACH ITEM in the collection: Returns a `Frame` object that represents the frame formatting for the style.
- [hasProofing](#hasproofing): For EACH ITEM in the collection: Specifies whether the spelling and grammar checker ignores text formatted with this style.
- [inUse](#inuse): For EACH ITEM in the collection: Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.
- [languageId](#languageid): For EACH ITEM in the collection: Specifies a `LanguageId` value that represents the language for the style.
- [languageIdFarEast](#languageidfareast): For EACH ITEM in the collection: Specifies an East Asian language for the style.
- [linked](#linked): For EACH ITEM in the collection: Gets whether a style is a linked style that can be used for both paragraph and character formatting.
- [linkStyle](#linkstyle): For EACH ITEM in the collection: Specifies a link between a paragraph and a character style.
- [listLevelNumber](#listlevelnumber): For EACH ITEM in the collection: Returns the list level for the style.
- [listTemplate](#listtemplate): For EACH ITEM in the collection: Gets a ListTemplate object that represents the list formatting for the specified Style object.
- [locked](#locked): For EACH ITEM in the collection: Specifies whether the style cannot be changed or edited.
- [nameLocal](#namelocal): For EACH ITEM in the collection: Gets the name of a style in the language of the user.
- [nextParagraphStyle](#nextparagraphstyle): For EACH ITEM in the collection: Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.
- [noSpaceBetweenParagraphsOfSameStyle](#nospacebetweenparagraphsofsamestyle): For EACH ITEM in the collection: Specifies whether to remove spacing between paragraphs that are formatted using the same style.
- [paragraphFormat](#paragraphformat): For EACH ITEM in the collection: Gets a ParagraphFormat object that represents the paragraph settings for the specified style.
- [priority](#priority): For EACH ITEM in the collection: Specifies the priority.
- [quickStyle](#quickstyle): For EACH ITEM in the collection: Specifies whether the style corresponds to an available quick style.
- [shading](#shading): For EACH ITEM in the collection: Gets a Shading object that represents the shading for the specified style. Not applicable to List style.
- [tableStyle](#tablestyle): For EACH ITEM in the collection: Gets a TableStyle object representing Style properties that can be applied to a table.
- [type](#type): For EACH ITEM in the collection: Gets the style type.
- [unhideWhenUsed](#unhidewhenused): For EACH ITEM in the collection: Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.
- [visibility](#visibility): For EACH ITEM in the collection: Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### automaticallyUpdate

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the style is automatically redefined based on the selection.

```typescript
automaticallyUpdate?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### baseStyle

For EACH ITEM in the collection: Specifies the name of an existing style to use as the base formatting of another style.

```typescript
baseStyle?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)  
Note: The ability to set `baseStyle` was introduced in WordApi 1.6.

---

### borders

For EACH ITEM in the collection: Specifies a BorderCollection object that represents all the borders for the specified style.

```typescript
borders?: Word.Interfaces.BorderCollectionLoadOptions;
```

Property Value: [Word.Interfaces.BorderCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.bordercollectionloadoptions)

Remarks: [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### builtIn

For EACH ITEM in the collection: Gets whether the specified style is a built-in style.

```typescript
builtIn?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### description

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the description of the specified style.

```typescript
description?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### font

For EACH ITEM in the collection: Gets a font object that represents the character formatting of the specified style.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### frame

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns a `Frame` object that represents the frame formatting for the style.

```typescript
frame?: Word.Interfaces.FrameLoadOptions;
```

Property Value: [Word.Interfaces.FrameLoadOptions](/en-us/javascript/api/word/word.interfaces.frameloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hasProofing

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the spelling and grammar checker ignores text formatted with this style.

```typescript
hasProofing?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### inUse

For EACH ITEM in the collection: Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.

```typescript
inUse?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### languageId

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies a `LanguageId` value that represents the language for the style.

```typescript
languageId?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### languageIdFarEast

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies an East Asian language for the style.

```typescript
languageIdFarEast?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### linked

For EACH ITEM in the collection: Gets whether a style is a linked style that can be used for both paragraph and character formatting.

```typescript
linked?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### linkStyle

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies a link between a paragraph and a character style.

```typescript
linkStyle?: Word.Interfaces.StyleLoadOptions;
```

Property Value: [Word.Interfaces.StyleLoadOptions](/en-us/javascript/api/word/word.interfaces.styleloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listLevelNumber

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns the list level for the style.

```typescript
listLevelNumber?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### listTemplate

For EACH ITEM in the collection: Gets a ListTemplate object that represents the list formatting for the specified Style object.

```typescript
listTemplate?: Word.Interfaces.ListTemplateLoadOptions;
```

Property Value: [Word.Interfaces.ListTemplateLoadOptions](/en-us/javascript/api/word/word.interfaces.listtemplateloadoptions)

Remarks: [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### locked

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the style cannot be changed or edited.

```typescript
locked?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### nameLocal

For EACH ITEM in the collection: Gets the name of a style in the language of the user.

```typescript
nameLocal?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### nextParagraphStyle

For EACH ITEM in the collection: Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.

```typescript
nextParagraphStyle?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)  
Note: The ability to set `nextParagraphStyle` was introduced in WordApi 1.6.

---

### noSpaceBetweenParagraphsOfSameStyle

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether to remove spacing between paragraphs that are formatted using the same style.

```typescript
noSpaceBetweenParagraphsOfSameStyle?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### paragraphFormat

For EACH ITEM in the collection: Gets a ParagraphFormat object that represents the paragraph settings for the specified style.

```typescript
paragraphFormat?: Word.Interfaces.ParagraphFormatLoadOptions;
```

Property Value: [Word.Interfaces.ParagraphFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.paragraphformatloadoptions)

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### priority

For EACH ITEM in the collection: Specifies the priority.

```typescript
priority?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### quickStyle

For EACH ITEM in the collection: Specifies whether the style corresponds to an available quick style.

```typescript
quickStyle?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shading

For EACH ITEM in the collection: Gets a Shading object that represents the shading for the specified style. Not applicable to List style.

```typescript
shading?: Word.Interfaces.ShadingLoadOptions;
```

Property Value: [Word.Interfaces.ShadingLoadOptions](/en-us/javascript/api/word/word.interfaces.shadingloadoptions)

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tableStyle

For EACH ITEM in the collection: Gets a TableStyle object representing Style properties that can be applied to a table.

```typescript
tableStyle?: Word.Interfaces.TableStyleLoadOptions;
```

Property Value: [Word.Interfaces.TableStyleLoadOptions](/en-us/javascript/api/word/word.interfaces.tablestyleloadoptions)

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

For EACH ITEM in the collection: Gets the style type.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### unhideWhenUsed

For EACH ITEM in the collection: Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.

```typescript
unhideWhenUsed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### visibility

For EACH ITEM in the collection: Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.

```typescript
visibility?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)