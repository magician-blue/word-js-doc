# Word.Interfaces.FontLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a font.

## Remarks

[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- `$all`  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- `allCaps`  
  Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters. The possible values are as follows:
  - `true`: All the text has the All Caps attribute.
  - `false`: None of the text has the All Caps attribute.
  - `null`: Returned if some, but not all, of the text has the All Caps attribute.

- `bold`  
  Specifies a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.

- `boldBidirectional`  
  Specifies whether the font is formatted as bold in a right-to-left language document. The possible values are as follows:
  - `true`: All the text is bold.
  - `false`: None of the text is bold.
  - `null`: Returned if some, but not all, of the text is bold.

- `color`  
  Specifies the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.

- `colorIndex`  
  Specifies a `ColorIndex` value that represents the color for the font.

- `colorIndexBidirectional`  
  Specifies the color for the `Font` object in a right-to-left language document.

- `contextualAlternates`  
  Specifies whether contextual alternates are enabled for the font.

- `diacriticColor`  
  Specifies the color to be used for diacritics for the `Font` object. You can provide the value in the '#RRGGBB' format.

- `disableCharacterSpaceGrid`  
  Specifies whether Microsoft Word ignores the number of characters per line for the corresponding `Font` object.

- `doubleStrikeThrough`  
  Specifies a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.

- `emboss`  
  Specifies whether the font is formatted as embossed. The possible values are as follows:
  - `true`: All the text is embossed.
  - `false`: None of the text is embossed.
  - `null`: Returned if some, but not all, of the text is embossed.

- `emphasisMark`  
  Specifies an `EmphasisMark` value that represents the emphasis mark for a character or designated character string.

- `engrave`  
  Specifies whether the font is formatted as engraved. The possible values are as follows:
  - `true`: All the text is engraved.
  - `false`: None of the text is engraved.
  - `null`: Returned if some, but not all, of the text is engraved.

- `fill`  
  Returns a `FillFormat` object that contains fill formatting properties for the font used by the range of text.

- `glow`  
  Returns a `GlowFormat` object that represents the glow formatting for the font used by the range of text.

- `hidden`  
  Specifies a value that indicates whether the font is tagged as hidden. True if the font is formatted as hidden text, otherwise, false.

- `highlightColor`  
  Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or `null` for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.

- `italic`  
  Specifies a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.

- `italicBidirectional`  
  Specifies whether the font is italicized in a right-to-left language document. The possible values are as follows:
  - `true`: All the text is italicized.
  - `false`: None of the text is italicized.
  - `null`: Returned if some, but not all, of the text is italicized.

- `kerning`  
  Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.

- `ligature`  
  Specifies the ligature setting for the `Font` object.

- `line`  
  Returns a `LineFormat` object that specifies the formatting for a line.

- `name`  
  Specifies a value that represents the name of the font.

- `nameAscii`  
  Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).

- `nameBidirectional`  
  Specifies the font name in a right-to-left language document.

- `nameFarEast`  
  Specifies the East Asian font name.

- `nameOther`  
  Specifies the font used for characters with codes from 128 through 255.

- `numberForm`  
  Specifies the number form setting for an OpenType font.

- `numberSpacing`  
  Specifies the number spacing setting for the font.

- `outline`  
  Specifies if the font is formatted as outlined. The possible values are as follows:
  - `true`: All the text is outlined.
  - `false`: None of the text is outlined.
  - `null`: Returned if some, but not all, of the text is outlined.

- `position`  
  Specifies the position of text (in points) relative to the base line.

- `reflection`  
  Returns a `ReflectionFormat` object that represents the reflection formatting for a shape.

- `scaling`  
  Specifies the scaling percentage applied to the font.

- `shadow`  
  Specifies if the font is formatted as shadowed. The possible values are as follows:
  - `true`: All the text is shadowed.
  - `false`: None of the text is shadowed.
  - `null`: Returned if some, but not all, of the text is shadowed.

- `size`  
  Specifies a value that represents the font size in points.

- `sizeBidirectional`  
  Specifies the font size in points for right-to-left text.

- `smallCaps`  
  Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters. The possible values are as follows:
  - `true`: All the text has the Small Caps attribute.
  - `false`: None of the text has the Small Caps attribute.
  - `null`: Returned if some, but not all, of the text has the Small Caps attribute.

- `spacing`  
  Specifies the spacing between characters.

- `strikeThrough`  
  Specifies a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.

- `stylisticSet`  
  Specifies the stylistic set for the font.

- `subscript`  
  Specifies a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.

- `superscript`  
  Specifies a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.

- `textColor`  
  Returns a `ColorFormat` object that represents the color for the font.

- `textShadow`  
  Returns a `ShadowFormat` object that specifies the shadow formatting for the font.

- `threeDimensionalFormat`  
  Returns a `ThreeDimensionalFormat` object that contains 3-dimensional (3D) effect formatting properties for the font.

- `underline`  
  Specifies a value that indicates the font's underline type. 'None' if the font isn't underlined.

- `underlineColor`  
  Specifies the color of the underline for the `Font` object. You can provide the value in the '#RRGGBB' format.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### allCaps

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters. The possible values are as follows:
- `true`: All the text has the All Caps attribute.
- `false`: None of the text has the All Caps attribute.
- `null`: Returned if some, but not all, of the text has the All Caps attribute.

```typescript
allCaps?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bold

Specifies a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.

```typescript
bold?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### boldBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as bold in a right-to-left language document. The possible values are as follows:
- `true`: All the text is bold.
- `false`: None of the text is bold.
- `null`: Returned if some, but not all, of the text is bold.

```typescript
boldBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### color

Specifies the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
color?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### colorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `ColorIndex` value that represents the color for the font.

```typescript
colorIndex?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### colorIndexBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the `Font` object in a right-to-left language document.

```typescript
colorIndexBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contextualAlternates

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether contextual alternates are enabled for the font.

```typescript
contextualAlternates?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### diacriticColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color to be used for diacritics for the `Font` object. You can provide the value in the '#RRGGBB' format.

```typescript
diacriticColor?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### disableCharacterSpaceGrid

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word ignores the number of characters per line for the corresponding `Font` object.

```typescript
disableCharacterSpaceGrid?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### doubleStrikeThrough

Specifies a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.

```typescript
doubleStrikeThrough?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### emboss

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as embossed. The possible values are as follows:
- `true`: All the text is embossed.
- `false`: None of the text is embossed.
- `null`: Returned if some, but not all, of the text is embossed.

```typescript
emboss?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### emphasisMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an `EmphasisMark` value that represents the emphasis mark for a character or designated character string.

```typescript
emphasisMark?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### engrave

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as engraved. The possible values are as follows:
- `true`: All the text is engraved.
- `false`: None of the text is engraved.
- `null`: Returned if some, but not all, of the text is engraved.

```typescript
engrave?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fill

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `FillFormat` object that contains fill formatting properties for the font used by the range of text.

```typescript
fill?: Word.Interfaces.FillFormatLoadOptions;
```

Property Value: [Word.Interfaces.FillFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.fillformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### glow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `GlowFormat` object that represents the glow formatting for the font used by the range of text.

```typescript
glow?: Word.Interfaces.GlowFormatLoadOptions;
```

Property Value: [Word.Interfaces.GlowFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.glowformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hidden

Specifies a value that indicates whether the font is tagged as hidden. True if the font is formatted as hidden text, otherwise, false.

```typescript
hidden?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### highlightColor

Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or `null` for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.

```typescript
highlightColor?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### italic

Specifies a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.

```typescript
italic?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### italicBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is italicized in a right-to-left language document. The possible values are as follows:
- `true`: All the text is italicized.
- `false`: None of the text is italicized.
- `null`: Returned if some, but not all, of the text is italicized.

```typescript
italicBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### kerning

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.

```typescript
kerning?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### ligature

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ligature setting for the `Font` object.

```typescript
ligature?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### line

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `LineFormat` object that specifies the formatting for a line.

```typescript
line?: Word.Interfaces.LineFormatLoadOptions;
```

Property Value: [Word.Interfaces.LineFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.lineformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### name

Specifies a value that represents the name of the font.

```typescript
name?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### nameAscii

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).

```typescript
nameAscii?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### nameBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font name in a right-to-left language document.

```typescript
nameBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### nameFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the East Asian font name.

```typescript
nameFarEast?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### nameOther

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font used for characters with codes from 128 through 255.

```typescript
nameOther?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### numberForm

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number form setting for an OpenType font.

```typescript
numberForm?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### numberSpacing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number spacing setting for the font.

```typescript
numberSpacing?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### outline

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font is formatted as outlined. The possible values are as follows:
- `true`: All the text is outlined.
- `false`: None of the text is outlined.
- `null`: Returned if some, but not all, of the text is outlined.

```typescript
outline?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### position

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position of text (in points) relative to the base line.

```typescript
position?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### reflection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ReflectionFormat` object that represents the reflection formatting for a shape.

```typescript
reflection?: Word.Interfaces.ReflectionFormatLoadOptions;
```

Property Value: [Word.Interfaces.ReflectionFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.reflectionformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### scaling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the scaling percentage applied to the font.

```typescript
scaling?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### shadow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font is formatted as shadowed. The possible values are as follows:
- `true`: All the text is shadowed.
- `false`: None of the text is shadowed.
- `null`: Returned if some, but not all, of the text is shadowed.

```typescript
shadow?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### size

Specifies a value that represents the font size in points.

```typescript
size?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### sizeBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font size in points for right-to-left text.

```typescript
sizeBidirectional?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### smallCaps

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters. The possible values are as follows:
- `true`: All the text has the Small Caps attribute.
- `false`: None of the text has the Small Caps attribute.
- `null`: Returned if some, but not all, of the text has the Small Caps attribute.

```typescript
smallCaps?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### spacing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the spacing between characters.

```typescript
spacing?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### strikeThrough

Specifies a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.

```typescript
strikeThrough?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### stylisticSet

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the stylistic set for the font.

```typescript
stylisticSet?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### subscript

Specifies a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.

```typescript
subscript?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### superscript

Specifies a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.

```typescript
superscript?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the color for the font.

```typescript
textColor?: Word.Interfaces.ColorFormatLoadOptions;
```

Property Value: [Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### textShadow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ShadowFormat` object that specifies the shadow formatting for the font.

```typescript
textShadow?: Word.Interfaces.ShadowFormatLoadOptions;
```

Property Value: [Word.Interfaces.ShadowFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.shadowformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### threeDimensionalFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ThreeDimensionalFormat` object that contains 3-dimensional (3D) effect formatting properties for the font.

```typescript
threeDimensionalFormat?: Word.Interfaces.ThreeDimensionalFormatLoadOptions;
```

Property Value: [Word.Interfaces.ThreeDimensionalFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.threedimensionalformatloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### underline

Specifies a value that indicates the font's underline type. 'None' if the font isn't underlined.

```typescript
underline?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### underlineColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color of the underline for the `Font` object. You can provide the value in the '#RRGGBB' format.

```typescript
underlineColor?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)