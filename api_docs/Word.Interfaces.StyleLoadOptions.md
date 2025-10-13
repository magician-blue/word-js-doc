# Word.Interfaces.StyleLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a style in a Word document.

## Remarks
[API set: WordApi 1.3]

## Properties
- $all: Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- automaticallyUpdate: Specifies whether the style is automatically redefined based on the selection.
- baseStyle: Specifies the name of an existing style to use as the base formatting of another style.
- borders: Specifies a BorderCollection object that represents all the borders for the specified style.
- builtIn: Gets whether the specified style is a built-in style.
- description: Gets the description of the specified style.
- font: Gets a font object that represents the character formatting of the specified style.
- frame: Returns a Frame object that represents the frame formatting for the style.
- hasProofing: Specifies whether the spelling and grammar checker ignores text formatted with this style.
- inUse: Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.
- languageId: Specifies a LanguageId value that represents the language for the style.
- languageIdFarEast: Specifies an East Asian language for the style.
- linked: Gets whether a style is a linked style that can be used for both paragraph and character formatting.
- linkStyle: Specifies a link between a paragraph and a character style.
- listLevelNumber: Returns the list level for the style.
- listTemplate: Gets a ListTemplate object that represents the list formatting for the specified Style object.
- locked: Specifies whether the style cannot be changed or edited.
- nameLocal: Gets the name of a style in the language of the user.
- nextParagraphStyle: Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.
- noSpaceBetweenParagraphsOfSameStyle: Specifies whether to remove spacing between paragraphs that are formatted using the same style.
- paragraphFormat: Gets a ParagraphFormat object that represents the paragraph settings for the specified style.
- priority: Specifies the priority.
- quickStyle: Specifies whether the style corresponds to an available quick style.
- shading: Gets a Shading object that represents the shading for the specified style. Not applicable to List style.
- tableStyle: Gets a TableStyle object representing Style properties that can be applied to a table.
- type: Gets the style type.
- unhideWhenUsed: Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.
- visibility: Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.

## Property Details

### $all
Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### automaticallyUpdate
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the style is automatically redefined based on the selection.

```typescript
automaticallyUpdate?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### baseStyle
Specifies the name of an existing style to use as the base formatting of another style.

```typescript
baseStyle?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]
- Note: The ability to set baseStyle was introduced in WordApi 1.6.

---

### borders
Specifies a BorderCollection object that represents all the borders for the specified style.

```typescript
borders?: Word.Interfaces.BorderCollectionLoadOptions;
```

Property Value: [Word.Interfaces.BorderCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.bordercollectionloadoptions)

Remarks
- [API set: WordApiDesktop 1.1]

---

### builtIn
Gets whether the specified style is a built-in style.

```typescript
builtIn?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]

---

### description
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the description of the specified style.

```typescript
description?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### font
Gets a font object that represents the character formatting of the specified style.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks
- [API set: WordApi 1.5]

---

### frame
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Frame object that represents the frame formatting for the style.

```typescript
frame?: Word.Interfaces.FrameLoadOptions;
```

Property Value: [Word.Interfaces.FrameLoadOptions](/en-us/javascript/api/word/word.interfaces.frameloadoptions)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### hasProofing
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the spelling and grammar checker ignores text formatted with this style.

```typescript
hasProofing?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### inUse
Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.

```typescript
inUse?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]

---

### languageId
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LanguageId value that represents the language for the style.

```typescript
languageId?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### languageIdFarEast
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an East Asian language for the style.

```typescript
languageIdFarEast?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### linked
Gets whether a style is a linked style that can be used for both paragraph and character formatting.

```typescript
linked?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]

---

### linkStyle
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a link between a paragraph and a character style.

```typescript
linkStyle?: Word.Interfaces.StyleLoadOptions;
```

Property Value: [Word.Interfaces.StyleLoadOptions](/en-us/javascript/api/word/word.interfaces.styleloadoptions)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### listLevelNumber
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the list level for the style.

```typescript
listLevelNumber?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### listTemplate
Gets a ListTemplate object that represents the list formatting for the specified Style object.

```typescript
listTemplate?: Word.Interfaces.ListTemplateLoadOptions;
```

Property Value: [Word.Interfaces.ListTemplateLoadOptions](/en-us/javascript/api/word/word.interfaces.listtemplateloadoptions)

Remarks
- [API set: WordApiDesktop 1.1]

---

### locked
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the style cannot be changed or edited.

```typescript
locked?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### nameLocal
Gets the name of a style in the language of the user.

```typescript
nameLocal?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]

---

### nextParagraphStyle
Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.

```typescript
nextParagraphStyle?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]
- Note: The ability to set nextParagraphStyle was introduced in WordApi 1.6.

---

### noSpaceBetweenParagraphsOfSameStyle
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove spacing between paragraphs that are formatted using the same style.

```typescript
noSpaceBetweenParagraphsOfSameStyle?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

---

### paragraphFormat
Gets a ParagraphFormat object that represents the paragraph settings for the specified style.

```typescript
paragraphFormat?: Word.Interfaces.ParagraphFormatLoadOptions;
```

Property Value: [Word.Interfaces.ParagraphFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.paragraphformatloadoptions)

Remarks
- [API set: WordApi 1.5]

---

### priority
Specifies the priority.

```typescript
priority?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]

---

### quickStyle
Specifies whether the style corresponds to an available quick style.

```typescript
quickStyle?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]

---

### shading
Gets a Shading object that represents the shading for the specified style. Not applicable to List style.

```typescript
shading?: Word.Interfaces.ShadingLoadOptions;
```

Property Value: [Word.Interfaces.ShadingLoadOptions](/en-us/javascript/api/word/word.interfaces.shadingloadoptions)

Remarks
- [API set: WordApi 1.6]

---

### tableStyle
Gets a TableStyle object representing Style properties that can be applied to a table.

```typescript
tableStyle?: Word.Interfaces.TableStyleLoadOptions;
```

Property Value: [Word.Interfaces.TableStyleLoadOptions](/en-us/javascript/api/word/word.interfaces.tablestyleloadoptions)

Remarks
- [API set: WordApi 1.6]

---

### type
Gets the style type.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]

---

### unhideWhenUsed
Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.

```typescript
unhideWhenUsed?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]

---

### visibility
Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.

```typescript
visibility?: boolean;
```

Property Value: boolean

Remarks
- [API set: WordApi 1.5]