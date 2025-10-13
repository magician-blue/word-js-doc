# Word.Interfaces.RangeCollectionLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Range](/en-us/javascript/api/word/word.range) objects.

## Remarks

[API set: WordApi 1.1]

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- bold: For EACH ITEM in the collection: Specifies whether the range is formatted as bold.
- boldBidirectional: For EACH ITEM in the collection: Specifies whether the range is formatted as bold in a right-to-left language document.
- case: For EACH ITEM in the collection: Specifies a `CharacterCase` value that represents the case of the text in the range.
- characterWidth: For EACH ITEM in the collection: Specifies the character width of the range.
- combineCharacters: For EACH ITEM in the collection: Specifies if the range contains combined characters.
- disableCharacterSpaceGrid: For EACH ITEM in the collection: Specifies if Microsoft Word ignores the number of characters per line for the corresponding `Range` object.
- emphasisMark: For EACH ITEM in the collection: Specifies the emphasis mark for a character or designated character string.
- end: For EACH ITEM in the collection: Specifies the ending character position of the range.
- fitTextWidth: For EACH ITEM in the collection: Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.
- font: For EACH ITEM in the collection: Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
- grammarChecked: For EACH ITEM in the collection: Specifies if a grammar check has been run on the range or document.
- hasNoProofing: For EACH ITEM in the collection: Specifies the proofing status (spelling and grammar checking) of the range.
- highlightColorIndex: For EACH ITEM in the collection: Specifies the highlight color for the range.
- horizontalInVertical: For EACH ITEM in the collection: Specifies the formatting for horizontal text set within vertical text.
- hyperlink: For EACH ITEM in the collection: Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
- id: For EACH ITEM in the collection: Specifies the ID for the range.
- isEmpty: For EACH ITEM in the collection: Checks whether the range length is zero.
- isEndOfRowMark: For EACH ITEM in the collection: Gets if the range is collapsed and is located at the end-of-row mark in a table.
- isTextVisibleOnScreen: For EACH ITEM in the collection: Gets whether the text in the range is visible on the screen.
- italic: For EACH ITEM in the collection: Specifies if the font or range is formatted as italic.
- italicBidirectional: For EACH ITEM in the collection: Specifies if the font or range is formatted as italic (right-to-left languages).
- kana: For EACH ITEM in the collection: Specifies whether the range of Japanese language text is hiragana or katakana.
- languageDetected: For EACH ITEM in the collection: Specifies whether Microsoft Word has detected the language of the text in the range.
- languageId: For EACH ITEM in the collection: Specifies a `LanguageId` value that represents the language for the range.
- languageIdFarEast: For EACH ITEM in the collection: Specifies an East Asian language for the range.
- languageIdOther: For EACH ITEM in the collection: Specifies a language for the range that isn't classified as an East Asian language.
- listFormat: For EACH ITEM in the collection: Returns a `ListFormat` object that represents all the list formatting characteristics of the range.
- parentBody: For EACH ITEM in the collection: Gets the parent body of the range.
- parentContentControl: For EACH ITEM in the collection: Gets the currently supported content control that contains the range. Throws an `ItemNotFound` error if there isn't a parent content control.
- parentContentControlOrNullObject: For EACH ITEM in the collection: Gets the currently supported content control that contains the range. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTable: For EACH ITEM in the collection: Gets the table that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table.
- parentTableCell: For EACH ITEM in the collection: Gets the table cell that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table cell.
- parentTableCellOrNullObject: For EACH ITEM in the collection: Gets the table cell that contains the range. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTableOrNullObject: For EACH ITEM in the collection: Gets the table that contains the range. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- shading: For EACH ITEM in the collection: Returns a `ShadingUniversal` object that refers to the shading formatting for the range.
- showAll: For EACH ITEM in the collection: Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.
- spellingChecked: For EACH ITEM in the collection: Specifies if spelling has been checked throughout the range or document.
- start: For EACH ITEM in the collection: Specifies the starting character position of the range.
- storyLength: For EACH ITEM in the collection: Gets the number of characters in the story that contains the range.
- storyType: For EACH ITEM in the collection: Gets the story type for the range.
- style: For EACH ITEM in the collection: Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBuiltIn: For EACH ITEM in the collection: Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- text: For EACH ITEM in the collection: Gets the text of the range.
- twoLinesInOne: For EACH ITEM in the collection: Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.
- underline: For EACH ITEM in the collection: Specifies the type of underline applied to the range.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### bold

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the range is formatted as bold.

```typescript
bold?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### boldBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the range is formatted as bold in a right-to-left language document.

```typescript
boldBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### case

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies a `CharacterCase` value that represents the case of the text in the range.

```typescript
case?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### characterWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the character width of the range.

```typescript
characterWidth?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### combineCharacters

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if the range contains combined characters.

```typescript
combineCharacters?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### disableCharacterSpaceGrid

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if Microsoft Word ignores the number of characters per line for the corresponding `Range` object.

```typescript
disableCharacterSpaceGrid?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### emphasisMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the emphasis mark for a character or designated character string.

```typescript
emphasisMark?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### end

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the ending character position of the range.

```typescript
end?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### fitTextWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.

```typescript
fitTextWidth?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### font

For EACH ITEM in the collection: Gets the text format of the range. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks: [API set: WordApi 1.1]

---

### grammarChecked

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if a grammar check has been run on the range or document.

```typescript
grammarChecked?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### hasNoProofing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the proofing status (spelling and grammar checking) of the range.

```typescript
hasNoProofing?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### highlightColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the highlight color for the range.

```typescript
highlightColorIndex?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### horizontalInVertical

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the formatting for horizontal text set within vertical text.

```typescript
horizontalInVertical?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### hyperlink

For EACH ITEM in the collection: Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.

```typescript
hyperlink?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3]

---

### id

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the ID for the range.

```typescript
id?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### isEmpty

For EACH ITEM in the collection: Checks whether the range length is zero.

```typescript
isEmpty?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3]

---

### isEndOfRowMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets if the range is collapsed and is located at the end-of-row mark in a table.

```typescript
isEndOfRowMark?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### isTextVisibleOnScreen

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets whether the text in the range is visible on the screen.

```typescript
isTextVisibleOnScreen?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### italic

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if the font or range is formatted as italic.

```typescript
italic?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### italicBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if the font or range is formatted as italic (right-to-left languages).

```typescript
italicBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### kana

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the range of Japanese language text is hiragana or katakana.

```typescript
kana?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### languageDetected

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether Microsoft Word has detected the language of the text in the range.

```typescript
languageDetected?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### languageId

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies a `LanguageId` value that represents the language for the range.

```typescript
languageId?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### languageIdFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies an East Asian language for the range.

```typescript
languageIdFarEast?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### languageIdOther

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies a language for the range that isn't classified as an East Asian language.

```typescript
languageIdOther?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### listFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns a `ListFormat` object that represents all the list formatting characteristics of the range.

```typescript
listFormat?: Word.Interfaces.ListFormatLoadOptions;
```

Property Value: [Word.Interfaces.ListFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.listformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### parentBody

For EACH ITEM in the collection: Gets the parent body of the range.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

Property Value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentContentControl

For EACH ITEM in the collection: Gets the currently supported content control that contains the range. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.1]

---

### parentContentControlOrNullObject

For EACH ITEM in the collection: Gets the currently supported content control that contains the range. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentTable

For EACH ITEM in the collection: Gets the table that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentTableCell

For EACH ITEM in the collection: Gets the table cell that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentTableCellOrNullObject

For EACH ITEM in the collection: Gets the table cell that contains the range. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3]

---

### parentTableOrNullObject

For EACH ITEM in the collection: Gets the table that contains the range. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3]

---

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns a `ShadingUniversal` object that refers to the shading formatting for the range.

```typescript
shading?: Word.Interfaces.ShadingUniversalLoadOptions;
```

Property Value: [Word.Interfaces.ShadingUniversalLoadOptions](/en-us/javascript/api/word/word.interfaces.shadinguniversalloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### showAll

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.

```typescript
showAll?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### spellingChecked

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if spelling has been checked throughout the range or document.

```typescript
spellingChecked?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### start

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the starting character position of the range.

```typescript
start?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### storyLength

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the number of characters in the story that contains the range.

```typescript
storyLength?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### storyType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the story type for the range.

```typescript
storyType?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### style

For EACH ITEM in the collection: Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1]

---

### styleBuiltIn

For EACH ITEM in the collection: Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3]

---

### text

For EACH ITEM in the collection: Gets the text of the range.

```typescript
text?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1]

---

### twoLinesInOne

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.

```typescript
twoLinesInOne?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### underline

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the type of underline applied to the range.

```typescript
underline?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]