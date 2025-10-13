# Word.Interfaces.FontData interface

Package: word

An interface describing the data returned by calling font.toJSON().

## Properties

- allCaps
  - Specifies whether the font is formatted as all capital letters, which makes lowercase letters appear as uppercase letters. The possible values are as follows:
    - true: All the text has the All Caps attribute.
    - false: None of the text has the All Caps attribute.
    - null: Returned if some, but not all, of the text has the All Caps attribute.

- bold
  - Specifies a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.

- boldBidirectional
  - Specifies whether the font is formatted as bold in a right-to-left language document. The possible values are as follows:
    - true: All the text is bold.
    - false: None of the text is bold.
    - null: Returned if some, but not all, of the text is bold.

- borders
  - Returns a BorderUniversalCollection object that represents all the borders for the font.

- color
  - Specifies the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.

- colorIndex
  - Specifies a ColorIndex value that represents the color for the font.

- colorIndexBidirectional
  - Specifies the color for the Font object in a right-to-left language document.

- contextualAlternates
  - Specifies whether contextual alternates are enabled for the font.

- diacriticColor
  - Specifies the color to be used for diacritics for the Font object. You can provide the value in the '#RRGGBB' format.

- disableCharacterSpaceGrid
  - Specifies whether Microsoft Word ignores the number of characters per line for the corresponding Font object.

- doubleStrikeThrough
  - Specifies a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.

- emboss
  - Specifies whether the font is formatted as embossed. The possible values are as follows:
    - true: All the text is embossed.
    - false: None of the text is embossed.
    - null: Returned if some, but not all, of the text is embossed.

- emphasisMark
  - Specifies an EmphasisMark value that represents the emphasis mark for a character or designated character string.

- engrave
  - Specifies whether the font is formatted as engraved. The possible values are as follows:
    - true: All the text is engraved.
    - false: None of the text is engraved.
    - null: Returned if some, but not all, of the text is engraved.

- fill
  - Returns a FillFormat object that contains fill formatting properties for the font used by the range of text.

- glow
  - Returns a GlowFormat object that represents the glow formatting for the font used by the range of text.

- hidden
  - Specifies a value that indicates whether the font is tagged as hidden. True if the font is formatted as hidden text, otherwise, false.

- highlightColor
  - Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.

- italic
  - Specifies a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.

- italicBidirectional
  - Specifies whether the font is italicized in a right-to-left language document. The possible values are as follows:
    - true: All the text is italicized.
    - false: None of the text is italicized.
    - null: Returned if some, but not all, of the text is italicized.

- kerning
  - Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.

- ligature
  - Specifies the ligature setting for the Font object.

- line
  - Returns a LineFormat object that specifies the formatting for a line.

- name
  - Specifies a value that represents the name of the font.

- nameAscii
  - Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).

- nameBidirectional
  - Specifies the font name in a right-to-left language document.

- nameFarEast
  - Specifies the East Asian font name.

- nameOther
  - Specifies the font used for characters with codes from 128 through 255.

- numberForm
  - Specifies the number form setting for an OpenType font.

- numberSpacing
  - Specifies the number spacing setting for the font.

- outline
  - Specifies if the font is formatted as outlined. The possible values are as follows:
    - true: All the text is outlined.
    - false: None of the text is outlined.
    - null: Returned if some, but not all, of the text is outlined.

- position
  - Specifies the position of text (in points) relative to the base line.

- reflection
  - Returns a ReflectionFormat object that represents the reflection formatting for a shape.

- scaling
  - Specifies the scaling percentage applied to the font.

- shadow
  - Specifies if the font is formatted as shadowed. The possible values are as follows:
    - true: All the text is shadowed.
    - false: None of the text is shadowed.
    - null: Returned if some, but not all, of the text is shadowed.

- size
  - Specifies a value that represents the font size in points.

- sizeBidirectional
  - Specifies the font size in points for right-to-left text.

- smallCaps
  - Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters. The possible values are as follows:
    - true: All the text has the Small Caps attribute.
    - false: None of the text has the Small Caps attribute.
    - null: Returned if some, but not all, of the text has the Small Caps attribute.

- spacing
  - Specifies the spacing between characters.

- strikeThrough
  - Specifies a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.

- stylisticSet
  - Specifies the stylistic set for the font.

- subscript
  - Specifies a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.

- superscript
  - Specifies a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.

- textColor
  - Returns a ColorFormat object that represents the color for the font.

- textShadow
  - Returns a ShadowFormat object that specifies the shadow formatting for the font.

- threeDimensionalFormat
  - Returns a ThreeDimensionalFormat object that contains 3-dimensional (3D) effect formatting properties for the font.

- underline
  - Specifies a value that indicates the font's underline type. 'None' if the font isn't underlined.

- underlineColor
  - Specifies the color of the underline for the Font object. You can provide the value in the '#RRGGBB' format.

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

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### bold

Specifies a value that indicates whether the font is bold. True if the font is formatted as bold, otherwise, false.

```typescript
bold?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi 1.1 ]

---

### boldBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as bold in a right-to-left language document. The possible values are as follows:
- true: All the text is bold.
- false: None of the text is bold.
- null: Returned if some, but not all, of the text is bold.

```typescript
boldBidirectional?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### borders

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders for the font.

```typescript
borders?: Word.Interfaces.BorderUniversalData[];
```

#### Property Value
Word.Interfaces.BorderUniversalData[]

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### color

Specifies the color for the specified font. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
color?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi 1.1 ]

---

### colorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a ColorIndex value that represents the color for the font.

```typescript
colorIndex?: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

#### Property Value
Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### colorIndexBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color for the Font object in a right-to-left language document.

```typescript
colorIndexBidirectional?: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

#### Property Value
Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### contextualAlternates

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether contextual alternates are enabled for the font.

```typescript
contextualAlternates?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### diacriticColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color to be used for diacritics for the Font object. You can provide the value in the '#RRGGBB' format.

```typescript
diacriticColor?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### disableCharacterSpaceGrid

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word ignores the number of characters per line for the corresponding Font object.

```typescript
disableCharacterSpaceGrid?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### doubleStrikeThrough

Specifies a value that indicates whether the font has a double strikethrough. True if the font is formatted as double strikethrough text, otherwise, false.

```typescript
doubleStrikeThrough?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi 1.1 ]

---

### emboss

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as embossed. The possible values are as follows:
- true: All the text is embossed.
- false: None of the text is embossed.
- null: Returned if some, but not all, of the text is embossed.

```typescript
emboss?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### emphasisMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an EmphasisMark value that represents the emphasis mark for a character or designated character string.

```typescript
emphasisMark?: Word.EmphasisMark | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle";
```

#### Property Value
Word.EmphasisMark | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### engrave

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as engraved. The possible values are as follows:
- true: All the text is engraved.
- false: None of the text is engraved.
- null: Returned if some, but not all, of the text is engraved.

```typescript
engrave?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### fill

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a FillFormat object that contains fill formatting properties for the font used by the range of text.

```typescript
fill?: Word.Interfaces.FillFormatData;
```

#### Property Value
Word.Interfaces.FillFormatData

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### glow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a GlowFormat object that represents the glow formatting for the font used by the range of text.

```typescript
glow?: Word.Interfaces.GlowFormatData;
```

#### Property Value
Word.Interfaces.GlowFormatData

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### hidden

Specifies a value that indicates whether the font is tagged as hidden. True if the font is formatted as hidden text, otherwise, false.

```typescript
hidden?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApiDesktop 1.2 ]

---

### highlightColor

Specifies the highlight color. To set it, use a value either in the '#RRGGBB' format or the color name. To remove highlight color, set it to null. The returned highlight color can be in the '#RRGGBB' format, an empty string for mixed highlight colors, or null for no highlight color. Note: Only the default highlight colors are available in Office for Windows Desktop. These are "Yellow", "Lime", "Turquoise", "Pink", "Blue", "Red", "DarkBlue", "Teal", "Green", "Purple", "DarkRed", "Olive", "Gray", "LightGray", and "Black". When the add-in runs in Office for Windows Desktop, any other color is converted to the closest color when applied to the font.

```typescript
highlightColor?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi 1.1 ]

---

### italic

Specifies a value that indicates whether the font is italicized. True if the font is italicized, otherwise, false.

```typescript
italic?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi 1.1 ]

---

### italicBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is italicized in a right-to-left language document. The possible values are as follows:
- true: All the text is italicized.
- false: None of the text is italicized.
- null: Returned if some, but not all, of the text is italicized.

```typescript
italicBidirectional?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### kerning

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the minimum font size for which Microsoft Word will adjust kerning automatically.

```typescript
kerning?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### ligature

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ligature setting for the Font object.

```typescript
ligature?: Word.Ligature | "None" | "Standard" | "Contextual" | "StandardContextual" | "Historical" | "StandardHistorical" | "ContextualHistorical" | "StandardContextualHistorical" | "Discretional" | "StandardDiscretional" | "ContextualDiscretional" | "StandardContextualDiscretional" | "HistoricalDiscretional" | "StandardHistoricalDiscretional" | "ContextualHistoricalDiscretional" | "All";
```

#### Property Value
Word.Ligature | "None" | "Standard" | "Contextual" | "StandardContextual" | "Historical" | "StandardHistorical" | "ContextualHistorical" | "StandardContextualHistorical" | "Discretional" | "StandardDiscretional" | "ContextualDiscretional" | "StandardContextualDiscretional" | "HistoricalDiscretional" | "StandardHistoricalDiscretional" | "ContextualHistoricalDiscretional" | "All"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### line

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a LineFormat object that specifies the formatting for a line.

```typescript
line?: Word.Interfaces.LineFormatData;
```

#### Property Value
Word.Interfaces.LineFormatData

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### name

Specifies a value that represents the name of the font.

```typescript
name?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi 1.1 ]

---

### nameAscii

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font used for Latin text (characters with character codes from 0 (zero) through 127).

```typescript
nameAscii?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### nameBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font name in a right-to-left language document.

```typescript
nameBidirectional?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### nameFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the East Asian font name.

```typescript
nameFarEast?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### nameOther

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font used for characters with codes from 128 through 255.

```typescript
nameOther?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### numberForm

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number form setting for an OpenType font.

```typescript
numberForm?: Word.NumberForm | "Default" | "Lining" | "OldStyle";
```

#### Property Value
Word.NumberForm | "Default" | "Lining" | "OldStyle"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### numberSpacing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the number spacing setting for the font.

```typescript
numberSpacing?: Word.NumberSpacing | "Default" | "Proportional" | "Tabular";
```

#### Property Value
Word.NumberSpacing | "Default" | "Proportional" | "Tabular"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### outline

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font is formatted as outlined. The possible values are as follows:
- true: All the text is outlined.
- false: None of the text is outlined.
- null: Returned if some, but not all, of the text is outlined.

```typescript
outline?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### position

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the position of text (in points) relative to the base line.

```typescript
position?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### reflection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ReflectionFormat object that represents the reflection formatting for a shape.

```typescript
reflection?: Word.Interfaces.ReflectionFormatData;
```

#### Property Value
Word.Interfaces.ReflectionFormatData

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### scaling

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the scaling percentage applied to the font.

```typescript
scaling?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### shadow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font is formatted as shadowed. The possible values are as follows:
- true: All the text is shadowed.
- false: None of the text is shadowed.
- null: Returned if some, but not all, of the text is shadowed.

```typescript
shadow?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### size

Specifies a value that represents the font size in points.

```typescript
size?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi 1.1 ]

---

### sizeBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the font size in points for right-to-left text.

```typescript
sizeBidirectional?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### smallCaps

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the font is formatted as small caps, which makes lowercase letters appear as small uppercase letters. The possible values are as follows:
- true: All the text has the Small Caps attribute.
- false: None of the text has the Small Caps attribute.
- null: Returned if some, but not all, of the text has the Small Caps attribute.

```typescript
smallCaps?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### spacing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the spacing between characters.

```typescript
spacing?: number;
```

#### Property Value
number

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### strikeThrough

Specifies a value that indicates whether the font has a strikethrough. True if the font is formatted as strikethrough text, otherwise, false.

```typescript
strikeThrough?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi 1.1 ]

---

### stylisticSet

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the stylistic set for the font.

```typescript
stylisticSet?: Word.StylisticSet | "Default" | "Set01" | "Set02" | "Set03" | "Set04" | "Set05" | "Set06" | "Set07" | "Set08" | "Set09" | "Set10" | "Set11" | "Set12" | "Set13" | "Set14" | "Set15" | "Set16" | "Set17" | "Set18" | "Set19" | "Set20";
```

#### Property Value
Word.StylisticSet | "Default" | "Set01" | "Set02" | "Set03" | "Set04" | "Set05" | "Set06" | "Set07" | "Set08" | "Set09" | "Set10" | "Set11" | "Set12" | "Set13" | "Set14" | "Set15" | "Set16" | "Set17" | "Set18" | "Set19" | "Set20"

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### subscript

Specifies a value that indicates whether the font is a subscript. True if the font is formatted as subscript, otherwise, false.

```typescript
subscript?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi 1.1 ]

---

### superscript

Specifies a value that indicates whether the font is a superscript. True if the font is formatted as superscript, otherwise, false.

```typescript
superscript?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi 1.1 ]

---

### textColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ColorFormat object that represents the color for the font.

```typescript
textColor?: Word.Interfaces.ColorFormatData;
```

#### Property Value
Word.Interfaces.ColorFormatData

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### textShadow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadowFormat object that specifies the shadow formatting for the font.

```typescript
textShadow?: Word.Interfaces.ShadowFormatData;
```

#### Property Value
Word.Interfaces.ShadowFormatData

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### threeDimensionalFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ThreeDimensionalFormat object that contains 3-dimensional (3D) effect formatting properties for the font.

```typescript
threeDimensionalFormat?: Word.Interfaces.ThreeDimensionalFormatData;
```

#### Property Value
Word.Interfaces.ThreeDimensionalFormatData

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### underline

Specifies a value that indicates the font's underline type. 'None' if the font isn't underlined.

```typescript
underline?: Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble";
```

#### Property Value
Word.UnderlineType | "Mixed" | "None" | "Hidden" | "DotLine" | "Single" | "Word" | "Double" | "Thick" | "Dotted" | "DottedHeavy" | "DashLine" | "DashLineHeavy" | "DashLineLong" | "DashLineLongHeavy" | "DotDashLine" | "DotDashLineHeavy" | "TwoDotDashLine" | "TwoDotDashLineHeavy" | "Wave" | "WaveHeavy" | "WaveDouble"

#### Remarks
[ API set: WordApi 1.1 ]

---

### underlineColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color of the underline for the Font object. You can provide the value in the '#RRGGBB' format.

```typescript
underlineColor?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]