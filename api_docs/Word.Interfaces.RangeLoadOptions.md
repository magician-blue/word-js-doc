# Word.Interfaces.RangeLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a contiguous area in a document.

## Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- [$all](#word-word-interfaces-rangeloadoptions-all-member): Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- [bold](#word-word-interfaces-rangeloadoptions-bold-member): Specifies whether the range is formatted as bold.
- [boldBidirectional](#word-word-interfaces-rangeloadoptions-boldbidirectional-member): Specifies whether the range is formatted as bold in a right-to-left language document.
- [case](#word-word-interfaces-rangeloadoptions-case-member): Specifies a CharacterCase value that represents the case of the text in the range.
- [characterWidth](#word-word-interfaces-rangeloadoptions-characterwidth-member): Specifies the character width of the range.
- [combineCharacters](#word-word-interfaces-rangeloadoptions-combinecharacters-member): Specifies if the range contains combined characters.
- [disableCharacterSpaceGrid](#word-word-interfaces-rangeloadoptions-disablecharacterspacegrid-member): Specifies if Microsoft Word ignores the number of characters per line for the corresponding Range object.
- [emphasisMark](#word-word-interfaces-rangeloadoptions-emphasismark-member): Specifies the emphasis mark for a character or designated character string.
- [end](#word-word-interfaces-rangeloadoptions-end-member): Specifies the ending character position of the range.
- [fitTextWidth](#word-word-interfaces-rangeloadoptions-fittextwidth-member): Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.
- [font](#word-word-interfaces-rangeloadoptions-font-member): Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
- [grammarChecked](#word-word-interfaces-rangeloadoptions-grammarchecked-member): Specifies if a grammar check has been run on the range or document.
- [hasNoProofing](#word-word-interfaces-rangeloadoptions-hasnoproofing-member): Specifies the proofing status (spelling and grammar checking) of the range.
- [highlightColorIndex](#word-word-interfaces-rangeloadoptions-highlightcolorindex-member): Specifies the highlight color for the range.
- [horizontalInVertical](#word-word-interfaces-rangeloadoptions-horizontalinvertical-member): Specifies the formatting for horizontal text set within vertical text.
- [hyperlink](#word-word-interfaces-rangeloadoptions-hyperlink-member): Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
- [id](#word-word-interfaces-rangeloadoptions-id-member): Specifies the ID for the range.
- [isEmpty](#word-word-interfaces-rangeloadoptions-isempty-member): Checks whether the range length is zero.
- [isEndOfRowMark](#word-word-interfaces-rangeloadoptions-isendofrowmark-member): Gets if the range is collapsed and is located at the end-of-row mark in a table.
- [isTextVisibleOnScreen](#word-word-interfaces-rangeloadoptions-istextvisibleonscreen-member): Gets whether the text in the range is visible on the screen.
- [italic](#word-word-interfaces-rangeloadoptions-italic-member): Specifies if the font or range is formatted as italic.
- [italicBidirectional](#word-word-interfaces-rangeloadoptions-italicbidirectional-member): Specifies if the font or range is formatted as italic (right-to-left languages).
- [kana](#word-word-interfaces-rangeloadoptions-kana-member): Specifies whether the range of Japanese language text is hiragana or katakana.
- [languageDetected](#word-word-interfaces-rangeloadoptions-languagedetected-member): Specifies whether Microsoft Word has detected the language of the text in the range.
- [languageId](#word-word-interfaces-rangeloadoptions-languageid-member): Specifies a LanguageId value that represents the language for the range.
- [languageIdFarEast](#word-word-interfaces-rangeloadoptions-languageidfareast-member): Specifies an East Asian language for the range.
- [languageIdOther](#word-word-interfaces-rangeloadoptions-languageidother-member): Specifies a language for the range that isn't classified as an East Asian language.
- [listFormat](#word-word-interfaces-rangeloadoptions-listformat-member): Returns a ListFormat object that represents all the list formatting characteristics of the range.
- [parentBody](#word-word-interfaces-rangeloadoptions-parentbody-member): Gets the parent body of the range.
- [parentContentControl](#word-word-interfaces-rangeloadoptions-parentcontentcontrol-member): Gets the currently supported content control that contains the range. Throws an ItemNotFound error if there isn't a parent content control.
- [parentContentControlOrNullObject](#word-word-interfaces-rangeloadoptions-parentcontentcontrolornullobject-member): Gets the currently supported content control that contains the range. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- [parentTable](#word-word-interfaces-rangeloadoptions-parenttable-member): Gets the table that contains the range. Throws an ItemNotFound error if it isn't contained in a table.
- [parentTableCell](#word-word-interfaces-rangeloadoptions-parenttablecell-member): Gets the table cell that contains the range. Throws an ItemNotFound error if it isn't contained in a table cell.
- [parentTableCellOrNullObject](#word-word-interfaces-rangeloadoptions-parenttablecellornullobject-member): Gets the table cell that contains the range. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- [parentTableOrNullObject](#word-word-interfaces-rangeloadoptions-parenttableornullobject-member): Gets the table that contains the range. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- [shading](#word-word-interfaces-rangeloadoptions-shading-member): Returns a ShadingUniversal object that refers to the shading formatting for the range.
- [showAll](#word-word-interfaces-rangeloadoptions-showall-member): Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.
- [spellingChecked](#word-word-interfaces-rangeloadoptions-spellingchecked-member): Specifies if spelling has been checked throughout the range or document.
- [start](#word-word-interfaces-rangeloadoptions-start-member): Specifies the starting character position of the range.
- [storyLength](#word-word-interfaces-rangeloadoptions-storylength-member): Gets the number of characters in the story that contains the range.
- [storyType](#word-word-interfaces-rangeloadoptions-storytype-member): Gets the story type for the range.
- [style](#word-word-interfaces-rangeloadoptions-style-member): Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- [styleBuiltIn](#word-word-interfaces-rangeloadoptions-stylebuiltin-member): Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- [text](#word-word-interfaces-rangeloadoptions-text-member): Gets the text of the range.
- [twoLinesInOne](#word-word-interfaces-rangeloadoptions-twolinesinone-member): Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.
- [underline](#word-word-interfaces-rangeloadoptions-underline-member): Specifies the type of underline applied to the range.

## Property Details

<a id="word-word-interfaces-rangeloadoptions-all-member"></a>
### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

<a id="word-word-interfaces-rangeloadoptions-bold-member"></a>
### bold

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range is formatted as bold.

```typescript
bold?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-boldbidirectional-member"></a>
### boldBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range is formatted as bold in a right-to-left language document.

```typescript
boldBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-case-member"></a>
### case

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a CharacterCase value that represents the case of the text in the range.

```typescript
case?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-characterwidth-member"></a>
### characterWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the character width of the range.

```typescript
characterWidth?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-combinecharacters-member"></a>
### combineCharacters

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the range contains combined characters.

```typescript
combineCharacters?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-disablecharacterspacegrid-member"></a>
### disableCharacterSpaceGrid

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if Microsoft Word ignores the number of characters per line for the corresponding Range object.

```typescript
disableCharacterSpaceGrid?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-emphasismark-member"></a>
### emphasisMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the emphasis mark for a character or designated character string.

```typescript
emphasisMark?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-end-member"></a>
### end

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ending character position of the range.

```typescript
end?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-fittextwidth-member"></a>
### fitTextWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.

```typescript
fitTextWidth?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-font-member"></a>
### font

Gets the text format of the range. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontLoadOptions;
```

Property Value: [Word.Interfaces.FontLoadOptions](/en-us/javascript/api/word/word.interfaces.fontloadoptions)

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-grammarchecked-member"></a>
### grammarChecked

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if a grammar check has been run on the range or document.

```typescript
grammarChecked?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-hasnoproofing-member"></a>
### hasNoProofing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the proofing status (spelling and grammar checking) of the range.

```typescript
hasNoProofing?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-highlightcolorindex-member"></a>
### highlightColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the highlight color for the range.

```typescript
highlightColorIndex?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-horizontalinvertical-member"></a>
### horizontalInVertical

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the formatting for horizontal text set within vertical text.

```typescript
horizontalInVertical?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-hyperlink-member"></a>
### hyperlink

Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.

```typescript
hyperlink?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-id-member"></a>
### id

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ID for the range.

```typescript
id?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-isempty-member"></a>
### isEmpty

Checks whether the range length is zero.

```typescript
isEmpty?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-isendofrowmark-member"></a>
### isEndOfRowMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets if the range is collapsed and is located at the end-of-row mark in a table.

```typescript
isEndOfRowMark?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-istextvisibleonscreen-member"></a>
### isTextVisibleOnScreen

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the text in the range is visible on the screen.

```typescript
isTextVisibleOnScreen?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-italic-member"></a>
### italic

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font or range is formatted as italic.

```typescript
italic?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-italicbidirectional-member"></a>
### italicBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font or range is formatted as italic (right-to-left languages).

```typescript
italicBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-kana-member"></a>
### kana

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range of Japanese language text is hiragana or katakana.

```typescript
kana?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-languagedetected-member"></a>
### languageDetected

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word has detected the language of the text in the range.

```typescript
languageDetected?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-languageid-member"></a>
### languageId

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LanguageId value that represents the language for the range.

```typescript
languageId?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-languageidfareast-member"></a>
### languageIdFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an East Asian language for the range.

```typescript
languageIdFarEast?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-languageidother-member"></a>
### languageIdOther

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a language for the range that isn't classified as an East Asian language.

```typescript
languageIdOther?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-listformat-member"></a>
### listFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ListFormat object that represents all the list formatting characteristics of the range.

```typescript
listFormat?: Word.Interfaces.ListFormatLoadOptions;
```

Property Value: [Word.Interfaces.ListFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.listformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-parentbody-member"></a>
### parentBody

Gets the parent body of the range.

```typescript
parentBody?: Word.Interfaces.BodyLoadOptions;
```

Property Value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-parentcontentcontrol-member"></a>
### parentContentControl

Gets the currently supported content control that contains the range. Throws an ItemNotFound error if there isn't a parent content control.

```typescript
parentContentControl?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-parentcontentcontrolornullobject-member"></a>
### parentContentControlOrNullObject

Gets the currently supported content control that contains the range. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentContentControlOrNullObject?: Word.Interfaces.ContentControlLoadOptions;
```

Property Value: [Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-parenttable-member"></a>
### parentTable

Gets the table that contains the range. Throws an ItemNotFound error if it isn't contained in a table.

```typescript
parentTable?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-parenttablecell-member"></a>
### parentTableCell

Gets the table cell that contains the range. Throws an ItemNotFound error if it isn't contained in a table cell.

```typescript
parentTableCell?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-parenttablecellornullobject-member"></a>
### parentTableCellOrNullObject

Gets the table cell that contains the range. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableCellOrNullObject?: Word.Interfaces.TableCellLoadOptions;
```

Property Value: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-parenttableornullobject-member"></a>
### parentTableOrNullObject

Gets the table that contains the range. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
parentTableOrNullObject?: Word.Interfaces.TableLoadOptions;
```

Property Value: [Word.Interfaces.TableLoadOptions](/en-us/javascript/api/word/word.interfaces.tableloadoptions)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-shading-member"></a>
### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the range.

```typescript
shading?: Word.Interfaces.ShadingUniversalLoadOptions;
```

Property Value: [Word.Interfaces.ShadingUniversalLoadOptions](/en-us/javascript/api/word/word.interfaces.shadinguniversalloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-showall-member"></a>
### showAll

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.

```typescript
showAll?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-spellingchecked-member"></a>
### spellingChecked

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if spelling has been checked throughout the range or document.

```typescript
spellingChecked?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-start-member"></a>
### start

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting character position of the range.

```typescript
start?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-storylength-member"></a>
### storyLength

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the number of characters in the story that contains the range.

```typescript
storyLength?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-storytype-member"></a>
### storyType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the story type for the range.

```typescript
storyType?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-style-member"></a>
### style

Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-stylebuiltin-member"></a>
### styleBuiltIn

Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-text-member"></a>
### text

Gets the text of the range.

```typescript
text?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-twolinesinone-member"></a>
### twoLinesInOne

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.

```typescript
twoLinesInOne?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-interfaces-rangeloadoptions-underline-member"></a>
### underline

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the type of underline applied to the range.

```typescript
underline?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)