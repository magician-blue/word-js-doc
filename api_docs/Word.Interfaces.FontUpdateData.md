# Word.Interfaces.FontUpdateData interface

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface for updating data on the Font object, for use in font.set({ ... }).

## Properties

- allCaps  
  Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters. The possible values are as follows:
  - true: All the text has the All Caps attribute.
  - false: None of the text has the All Caps attribute.
  - null: Returned if some, but not all, of the text has the All Caps attribute.

- bold  
  Specifies a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.

- boldBidirectional  
  Specifies whether the font is formatted as bold in a right-to-left language document. The possible values are as follows:
  - true: All the text is bold.
  - false: None of the text is bold.
  - null: Returned if some, but not all, of the text is bold.

- color  
  Specifies the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.

- colorIndex  
  Specifies a ColorIndex value that represents the color for the font.

- colorIndexBidirectional  
  Specifies the color for the Font object in a right-to-left language document.

- contextualAlternates  
  Specifies whether contextual alternates are enabled for the font.

- diacriticColor  
  Specifies the color to be used for diacritics for the Font object. You can provide the value in the '#RRGGBB' format.

- disableCharacterSpaceGrid  
  Specifies whether Microsoft Word ignores the number of characters per line for the corresponding Font object.

- doubleStrikeThrough  
  Specifies a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.

- emboss  
  Specifies whether the font is formatted as embossed. The possible values are as follows:
  - true: All the text is embossed.
  - false: None of the text is embossed.
  - null: Returned if some, but not all, of the text is embossed.

- emphasisMark  
  Specifies an EmphasisMark value that represents the emphasis mark for a character or designated character string.

- engrave  
  Specifies whether the font is formatted as engraved. The possible values are as follows:
  - true: All the text is engraved.
  - false: None of the text is engraved.
  - null: Returned if some, but not all, of the text is engraved.

- fill  
  Returns a FillFormat object that contains fill formatting properties for the font used by the range of text.

- glow  
  Returns a GlowFormat object that represents the glow formatting for the font used by the range of text.

- hidden  
  Specifies a value that indicates whether the font is tagged as hidden. True if the font is formatted as hidden text, otherwise, false.

- highlightColor  
  Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.

- italic  
  Specifies a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.

- italicBidirectional  
  Specifies whether the font is italicized in a right-to-left language document. The possible values are as follows:
  - true: All the text is italicized.
  - false: None of the text is italicized.
  - null: Returned if some, but not all, of the text is italicized.

- kerning  
  Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.

- ligature  
  Specifies the ligature setting for the Font object.

- line  
  Returns a LineFormat object that specifies the formatting for a line.

- name  
  Specifies a value that represents the name of the font.

- nameAscii  
  Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).

- nameBidirectional  
  Specifies the font name in a right-to-left language document.

- nameFarEast  
  Specifies the East Asian font name.

- nameOther  
  Specifies the font used for characters with codes from 128 through 255.

- numberForm  
  Specifies the number form setting for an OpenType font.

- numberSpacing  
  Specifies the number spacing setting for the font.

- outline  
  Specifies if the font is formatted as outlined. The possible values are as follows:
  - true: All the text is outlined.
  - false: None of the text is outlined.
  - null: Returned if some, but not all, of the text is outlined.

- position  
  Specifies the position of text (in points) relative to the base line.

- reflection  
  Returns a ReflectionFormat object that represents the reflection formatting for a shape.

- scaling  
  Specifies the scaling percentage applied to the font.

- shadow  
  Specifies if the font is formatted as shadowed. The possible values are as follows:
  - true: All the text is shadowed.
  - false: None of the text is shadowed.
  - null: Returned if some, but not all, of the text is shadowed.

- size  
  Specifies a value that represents the font size in points.

- sizeBidirectional  
  Specifies the font size in points for right-to-left text.

- smallCaps  
  Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters. The possible values are as follows:
  - true: All the text has the Small Caps attribute.
  - false: None of the text has the Small Caps attribute.
  - null: Returned if some, but not all, of the text has the Small Caps attribute.

- spacing  
  Specifies the spacing between characters.

- strikeThrough  
  Specifies a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.

- stylisticSet  
  Specifies the stylistic set for the font.

- subscript  
  Specifies a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.

- superscript  
  Specifies a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.

- textColor  
  Returns a ColorFormat object that represents the color for the font.

- textShadow  
  Returns a ShadowFormat object that specifies the shadow formatting for the font.

- threeDimensionalFormat  
  Returns a ThreeDimensionalFormat object that contains 3-dimensional (3D) effect formatting properties for the font.

- underline  
  Specifies a value that indicates the font's underline type. 'None' if the font isn't underlined.

- underlineColor  
  Specifies the color of the underline for the Font object. You can provide the value in the '#RRGGBB' format.

## Property Details

### allCaps

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters. The possible values are as follows:
- true: All the text has the All Caps attribute.
- false: None of the text has the All Caps attribute.
- null: Returned if some, but not all, of the text has the All Caps attribute.

```typescript
allCaps?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bold

Specifies a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.

```typescript
bold?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### boldBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as bold in a right-to-left language document. The possible values are as follows:
- true: All the text is bold.
- false: None of the text is bold.
- null: Returned if some, but not all, of the text is bold.

```typescript
boldBidirectional?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### color

Specifies the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
color?: string;
```

- Property value: string
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### colorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a ColorIndex value that represents the color for the font.

```typescript
colorIndex?: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

- Property value: [Word.ColorIndex](https://learn.microsoft.com/en-us/javascript/api/word/word.colorindex) | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### colorIndexBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the Font object in a right-to-left language document.

```typescript
colorIndexBidirectional?: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

- Property value: [Word.ColorIndex](https://learn.microsoft.com/en-us/javascript/api/word/word.colorindex) | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### contextualAlternates

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether contextual alternates are enabled for the font.

```typescript
contextualAlternates?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### diacriticColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color to be used for diacritics for the Font object. You can provide the value in the '#RRGGBB' format.

```typescript
diacriticColor?: string;
```

- Property value: string
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### disableCharacterSpaceGrid

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word ignores the number of characters per line for the corresponding Font object.

```typescript
disableCharacterSpaceGrid?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### doubleStrikeThrough

Specifies a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.

```typescript
doubleStrikeThrough?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### emboss

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as embossed. The possible values are as follows:
- true: All the text is embossed.
- false: None of the text is embossed.
- null: Returned if some, but not all, of the text is embossed.

```typescript
emboss?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### emphasisMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an EmphasisMark value that represents the emphasis mark for a character or designated character string.

```typescript
emphasisMark?: Word.EmphasisMark | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle";
```

- Property value: [Word.EmphasisMark](https://learn.microsoft.com/en-us/javascript/api/word/word.emphasismark) | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### engrave

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as engraved. The possible values are as follows:
- true: All the text is engraved.
- false: None of the text is engraved.
- null: Returned if some, but not all, of the text is engraved.

```typescript
engrave?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### fill

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a FillFormat object that contains fill formatting properties for the font used by the range of text.

```typescript
fill?: Word.Interfaces.FillFormatUpdateData;
```

- Property value: [Word.Interfaces.FillFormatUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.fillformatupdatedata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### glow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a GlowFormat object that represents the glow formatting for the font used by the range of text.

```typescript
glow?: Word.Interfaces.GlowFormatUpdateData;
```

- Property value: [Word.Interfaces.GlowFormatUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.glowformatupdatedata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### hidden

Specifies a value that indicates whether the font is tagged as hidden. True if the font is formatted as hidden text, otherwise, false.

```typescript
hidden?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### highlightColor

Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.

```typescript
highlightColor?: string;
```

- Property value: string
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### italic

Specifies a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.

```typescript
italic?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### italicBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is italicized in a right-to-left language document. The possible values are as follows:
- true: All the text is italicized.
- false: None of the text is italicized.
- null: Returned if some, but not all, of the text is italicized.

```typescript
italicBidirectional?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### kerning

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.

```typescript
kerning?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### ligature

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ligature setting for the Font object.

```typescript
ligature?: Word.Ligature | "None" | "Standard" | "Contextual" | "StandardContextual" | "Historical" | "StandardHistorical" | "ContextualHistorical" | "StandardContextualHistorical" | "Discretional" | "StandardDiscretional" | "ContextualDiscretional" | "StandardContextualDiscretional" | "HistoricalDiscretional" | "StandardHistoricalDiscretional" | "ContextualHistoricalDiscretional" | "All";
```

- Property value: [Word.Ligature](https://learn.microsoft.com/en-us/javascript/api/word/word.ligature) | "None" | "Standard" | "Contextual" | "StandardContextual" | "Historical" | "StandardHistorical" | "ContextualHistorical" | "StandardContextualHistorical" | "Discretional" | "StandardDiscretional" | "ContextualDiscretional" | "StandardContextualDiscretional" | "HistoricalDiscretional" | "StandardHistoricalDiscretional" | "ContextualHistoricalDiscretional" | "All"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### line

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a LineFormat object that specifies the formatting for a line.

```typescript
line?: Word.Interfaces.LineFormatUpdateData;
```

- Property value: [Word.Interfaces.LineFormatUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.lineformatupdatedata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name

Specifies a value that represents the name of the font.

```typescript
name?: string;
```

- Property value: string
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nameAscii

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).

```typescript
nameAscii?: string;
```

- Property value: string
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nameBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font name in a right-to-left language document.

```typescript
nameBidirectional?: string;
```

- Property value: string
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nameFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the East Asian font name.

```typescript
nameFarEast?: string;
```

- Property value: string
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nameOther

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font used for characters with codes from 128 through 255.

```typescript
nameOther?: string;
```

- Property value: string
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### numberForm

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number form setting for an OpenType font.

```typescript
numberForm?: Word.NumberForm | "Default" | "Lining" | "OldStyle";
```

- Property value: [Word.NumberForm](https://learn.microsoft.com/en-us/javascript/api/word/word.numberform) | "Default" | "Lining" | "OldStyle"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### numberSpacing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number spacing setting for the font.

```typescript
numberSpacing?: Word.NumberSpacing | "Default" | "Proportional" | "Tabular";
```

- Property value: [Word.NumberSpacing](https://learn.microsoft.com/en-us/javascript/api/word/word.numberspacing) | "Default" | "Proportional" | "Tabular"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### outline

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font is formatted as outlined. The possible values are as follows:
- true: All the text is outlined.
- false: None of the text is outlined.
- null: Returned if some, but not all, of the text is outlined.

```typescript
outline?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### position

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position of text (in points) relative to the base line.

```typescript
position?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### reflection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ReflectionFormat object that represents the reflection formatting for a shape.

```typescript
reflection?: Word.Interfaces.ReflectionFormatUpdateData;
```

- Property value: [Word.Interfaces.ReflectionFormatUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.reflectionformatupdatedata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### scaling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the scaling percentage applied to the font.

```typescript
scaling?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shadow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font is formatted as shadowed. The possible values are as follows:
- true: All the text is shadowed.
- false: None of the text is shadowed.
- null: Returned if some, but not all, of the text is shadowed.

```typescript
shadow?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### size

Specifies a value that represents the font size in points.

```typescript
size?: number;
```

- Property value: number
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sizeBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font size in points for right-to-left text.

```typescript
sizeBidirectional?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### smallCaps

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters. The possible values are as follows:
- true: All the text has the Small Caps attribute.
- false: None of the text has the Small Caps attribute.
- null: Returned if some, but not all, of the text has the Small Caps attribute.

```typescript
smallCaps?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### spacing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the spacing between characters.

```typescript
spacing?: number;
```

- Property value: number
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### strikeThrough

Specifies a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.

```typescript
strikeThrough?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### stylisticSet

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the stylistic set for the font.

```typescript
stylisticSet?: Word.StylisticSet | "Default" | "Set01" | "Set02" | "Set03" | "Set04" | "Set05" | "Set06" | "Set07" | "Set08" | "Set09" | "Set10" | "Set11" | "Set12" | "Set13" | "Set14" | "Set15" | "Set16" | "Set17" | "Set18" | "Set19" | "Set20";
```

- Property value: [Word.StylisticSet](https://learn.microsoft.com/en-us/javascript/api/word/word.stylisticset) | "Default" | "Set01" | "Set02" | "Set03" | "Set04" | "Set05" | "Set06" | "Set07" | "Set08" | "Set09" | "Set10" | "Set11" | "Set12" | "Set13" | "Set14" | "Set15" | "Set16" | "Set17" | "Set18" | "Set19" | "Set20"
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### subscript

Specifies a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.

```typescript
subscript?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### superscript

Specifies a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.

```typescript
superscript?: boolean;
```

- Property value: boolean
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ColorFormat object that represents the color for the font.

```typescript
textColor?: Word.Interfaces.ColorFormatUpdateData;
```

- Property value: [Word.Interfaces.ColorFormatUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.colorformatupdatedata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textShadow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadowFormat object that specifies the shadow formatting for the font.

```typescript
textShadow?: Word.Interfaces.ShadowFormatUpdateData;
```

- Property value: [Word.Interfaces.ShadowFormatUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shadowformatupdatedata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### threeDimensionalFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ThreeDimensionalFormat object that contains 3-dimensional (3D) effect formatting properties for the font.

```typescript
threeDimensionalFormat?: Word.Interfaces.ThreeDimensionalFormatUpdateData;
```

- Property value: [Word.Interfaces.ThreeDimensionalFormatUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.threedimensionalformatupdatedata)
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### underline

Specifies a value that indicates the font's underline type. 'None' if the font isn't underlined.

```typescript
underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
```

- Property value: [Word.UnderlineType](https://learn.microsoft.com/en-us/javascript/api/word/word.underlinetype) | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble"
- Remarks: [API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### underlineColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color of the underline for the Font object. You can provide the value in the '#RRGGBB' format.

```typescript
underlineColor?: string;
```

- Property value: string
- Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)