# Word.Interfaces.RangeData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface describing the data returned by calling range.toJSON().

## Properties

- bold — Specifies whether the range is formatted as bold.
- boldBidirectional — Specifies whether the range is formatted as bold in a right-to-left language document.
- borders — Returns a BorderUniversalCollection object that represents all the borders for the range.
- case — Specifies a CharacterCase value that represents the case of the text in the range.
- characterWidth — Specifies the character width of the range.
- combineCharacters — Specifies if the range contains combined characters.
- disableCharacterSpaceGrid — Specifies if Microsoft Word ignores the number of characters per line for the corresponding Range object.
- emphasisMark — Specifies the emphasis mark for a character or designated character string.
- end — Specifies the ending character position of the range.
- fields — Gets the collection of field objects in the range.
- fitTextWidth — Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.
- font — Gets the text format of the range. Use this to get and set font name, size, color, and other properties.
- frames — Gets a FrameCollection object that represents all the frames in the range.
- grammarChecked — Specifies if a grammar check has been run on the range or document.
- hasNoProofing — Specifies the proofing status (spelling and grammar checking) of the range.
- highlightColorIndex — Specifies the highlight color for the range.
- horizontalInVertical — Specifies the formatting for horizontal text set within vertical text.
- hyperlink — Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.
- hyperlinks — Returns a HyperlinkCollection object that represents all the hyperlinks in the range.
- id — Specifies the ID for the range.
- inlinePictures — Gets the collection of inline picture objects in the range.
- isEmpty — Checks whether the range length is zero.
- isEndOfRowMark — Gets if the range is collapsed and is located at the end-of-row mark in a table.
- isTextVisibleOnScreen — Gets whether the text in the range is visible on the screen.
- italic — Specifies if the font or range is formatted as italic.
- italicBidirectional — Specifies if the font or range is formatted as italic (right-to-left languages).
- kana — Specifies whether the range of Japanese language text is hiragana or katakana.
- languageDetected — Specifies whether Microsoft Word has detected the language of the text in the range.
- languageId — Specifies a LanguageId value that represents the language for the range.
- languageIdFarEast — Specifies an East Asian language for the range.
- languageIdOther — Specifies a language for the range that isn't classified as an East Asian language.
- listFormat — Returns a ListFormat object that represents all the list formatting characteristics of the range.
- pages — Gets the collection of pages in the range.
- sections — Gets the collection of sections in the range.
- shading — Returns a ShadingUniversal object that refers to the shading formatting for the range.
- shapes — Gets the collection of shape objects anchored in the range, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.
- showAll — Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.
- spellingChecked — Specifies if spelling has been checked throughout the range or document.
- start — Specifies the starting character position of the range.
- storyLength — Gets the number of characters in the story that contains the range.
- storyType — Gets the story type for the range.
- style — Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBuiltIn — Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- tableColumns — Gets a TableColumnCollection object that represents all the table columns in the range.
- text — Gets the text of the range.
- twoLinesInOne — Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.
- underline — Specifies the type of underline applied to the range.

## Property Details

### bold

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range is formatted as bold.

```typescript
bold?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### boldBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range is formatted as bold in a right-to-left language document.

```typescript
boldBidirectional?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### borders

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders for the range.

```typescript
borders?: Word.Interfaces.BorderUniversalData[];
```

#### Property Value
[Word.Interfaces.BorderUniversalData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.borderuniversaldata)[]

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### case

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a CharacterCase value that represents the case of the text in the range.

```typescript
case?: Word.CharacterCase | "Next" | "Lower" | "Upper" | "TitleWord" | "TitleSentence" | "Toggle" | "HalfWidth" | "FullWidth" | "Katakana" | "Hiragana";
```

#### Property Value
[Word.CharacterCase](https://learn.microsoft.com/en-us/javascript/api/word/word.charactercase) | "Next" | "Lower" | "Upper" | "TitleWord" | "TitleSentence" | "Toggle" | "HalfWidth" | "FullWidth" | "Katakana" | "Hiragana"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### characterWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the character width of the range.

```typescript
characterWidth?: Word.CharacterWidth | "Half" | "Full";
```

#### Property Value
[Word.CharacterWidth](https://learn.microsoft.com/en-us/javascript/api/word/word.characterwidth) | "Half" | "Full"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### combineCharacters

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the range contains combined characters.

```typescript
combineCharacters?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### disableCharacterSpaceGrid

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if Microsoft Word ignores the number of characters per line for the corresponding Range object.

```typescript
disableCharacterSpaceGrid?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### emphasisMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the emphasis mark for a character or designated character string.

```typescript
emphasisMark?: Word.EmphasisMark | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle";
```

#### Property Value
[Word.EmphasisMark](https://learn.microsoft.com/en-us/javascript/api/word/word.emphasismark) | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### end

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ending character position of the range.

```typescript
end?: number;
```

#### Property Value
number

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### fields

Gets the collection of field objects in the range.

```typescript
fields?: Word.Interfaces.FieldData[];
```

#### Property Value
[Word.Interfaces.FieldData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.fielddata)[]

Remarks  
[ API set: WordApi 1.4 ]

---

### fitTextWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.

```typescript
fitTextWidth?: number;
```

#### Property Value
number

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### font

Gets the text format of the range. Use this to get and set font name, size, color, and other properties.

```typescript
font?: Word.Interfaces.FontData;
```

#### Property Value
[Word.Interfaces.FontData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.fontdata)

Remarks  
[ API set: WordApi 1.1 ]

---

### frames

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a FrameCollection object that represents all the frames in the range.

```typescript
frames?: Word.Interfaces.FrameData[];
```

#### Property Value
[Word.Interfaces.FrameData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.framedata)[]

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### grammarChecked

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if a grammar check has been run on the range or document.

```typescript
grammarChecked?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### hasNoProofing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the proofing status (spelling and grammar checking) of the range.

```typescript
hasNoProofing?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### highlightColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the highlight color for the range.

```typescript
highlightColorIndex?: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

#### Property Value
[Word.ColorIndex](https://learn.microsoft.com/en-us/javascript/api/word/word.colorindex) | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### horizontalInVertical

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the formatting for horizontal text set within vertical text.

```typescript
horizontalInVertical?: Word.HorizontalInVerticalType | "None" | "FitInLine" | "ResizeLine";
```

#### Property Value
[Word.HorizontalInVerticalType](https://learn.microsoft.com/en-us/javascript/api/word/word.horizontalinverticaltype) | "None" | "FitInLine" | "ResizeLine"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### hyperlink

Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.

```typescript
hyperlink?: string;
```

#### Property Value
string

Remarks  
[ API set: WordApi 1.3 ]

---

### hyperlinks

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a HyperlinkCollection object that represents all the hyperlinks in the range.

```typescript
hyperlinks?: Word.Interfaces.HyperlinkData[];
```

#### Property Value
[Word.Interfaces.HyperlinkData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkdata)[]

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### id

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ID for the range.

```typescript
id?: string;
```

#### Property Value
string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### inlinePictures

Gets the collection of inline picture objects in the range.

```typescript
inlinePictures?: Word.Interfaces.InlinePictureData[];
```

#### Property Value
[Word.Interfaces.InlinePictureData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.inlinepicturedata)[]

Remarks  
[ API set: WordApi 1.2 ]

---

### isEmpty

Checks whether the range length is zero.

```typescript
isEmpty?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi 1.3 ]

---

### isEndOfRowMark

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets if the range is collapsed and is located at the end-of-row mark in a table.

```typescript
isEndOfRowMark?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isTextVisibleOnScreen

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the text in the range is visible on the screen.

```typescript
isTextVisibleOnScreen?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### italic

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font or range is formatted as italic.

```typescript
italic?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### italicBidirectional

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font or range is formatted as italic (right-to-left languages).

```typescript
italicBidirectional?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### kana

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range of Japanese language text is hiragana or katakana.

```typescript
kana?: Word.Kana | "Katakana" | "Hiragana";
```

#### Property Value
[Word.Kana](https://learn.microsoft.com/en-us/javascript/api/word/word.kana) | "Katakana" | "Hiragana"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### languageDetected

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word has detected the language of the text in the range.

```typescript
languageDetected?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### languageId

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LanguageId value that represents the language for the range.

```typescript
languageId?: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

#### Property Value
[Word.LanguageId](https://learn.microsoft.com/en-us/javascript/api/word/word.languageid) | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### languageIdFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an East Asian language for the range.

```typescript
languageIdFarEast?: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

#### Property Value
[Word.LanguageId](https://learn.microsoft.com/en-us/javascript/api/word/word.languageid) | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### languageIdOther

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a language for the range that isn't classified as an East Asian language.

```typescript
languageIdOther?: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

#### Property Value
[Word.LanguageId](https://learn.microsoft.com/en-us/javascript/api/word/word.languageid) | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### listFormat

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ListFormat object that represents all the list formatting characteristics of the range.

```typescript
listFormat?: Word.Interfaces.ListFormatData;
```

#### Property Value
[Word.Interfaces.ListFormatData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.listformatdata)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### pages

Gets the collection of pages in the range.

```typescript
pages?: Word.Interfaces.PageData[];
```

#### Property Value
[Word.Interfaces.PageData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.pagedata)[]

Remarks  
[ API set: WordApiDesktop 1.2 ]

---

### sections

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the collection of sections in the range.

```typescript
sections?: Word.Interfaces.SectionData[];
```

#### Property Value
[Word.Interfaces.SectionData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.sectiondata)[]

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### shading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the range.

```typescript
shading?: Word.Interfaces.ShadingUniversalData;
```

#### Property Value
[Word.Interfaces.ShadingUniversalData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shadinguniversaldata)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### shapes

Gets the collection of shape objects anchored in the range, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

```typescript
shapes?: Word.Interfaces.ShapeData[];
```

#### Property Value
[Word.Interfaces.ShapeData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.shapedata)[]

Remarks  
[ API set: WordApiDesktop 1.2 ]

---

### showAll

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.

```typescript
showAll?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### spellingChecked

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if spelling has been checked throughout the range or document.

```typescript
spellingChecked?: boolean;
```

#### Property Value
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### start

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting character position of the range.

```typescript
start?: number;
```

#### Property Value
number

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### storyLength

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the number of characters in the story that contains the range.

```typescript
storyLength?: number;
```

#### Property Value
number

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### storyType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the story type for the range.

```typescript
storyType?: Word.StoryType | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice";
```

#### Property Value
[Word.StoryType](https://learn.microsoft.com/en-us/javascript/api/word/word.storytype) | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### style

Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style?: string;
```

#### Property Value
string

Remarks  
[ API set: WordApi 1.1 ]

---

### styleBuiltIn

Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn?: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

#### Property Value
[Word.BuiltInStyleName](https://learn.microsoft.com/en-us/javascript/api/word/word.builtinstylename) | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"

Remarks  
[ API set: WordApi 1.3 ]

---

### tableColumns

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a TableColumnCollection object that represents all the table columns in the range.

```typescript
tableColumns?: Word.Interfaces.TableColumnData[];
```

#### Property Value
[Word.Interfaces.TableColumnData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tablecolumndata)[]

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### text

Gets the text of the range.

```typescript
text?: string;
```

#### Property Value
string

Remarks  
[ API set: WordApi 1.1 ]

---

### twoLinesInOne

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.

```typescript
twoLinesInOne?: Word.TwoLinesInOneType | "None" | "NoBrackets" | "Parentheses" | "SquareBrackets" | "AngleBrackets" | "CurlyBrackets";
```

#### Property Value
[Word.TwoLinesInOneType](https://learn.microsoft.com/en-us/javascript/api/word/word.twolinesinonetype) | "None" | "NoBrackets" | "Parentheses" | "SquareBrackets" | "AngleBrackets" | "CurlyBrackets"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### underline

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the type of underline applied to the range.

```typescript
underline?: Word.Underline | "None" | "Single" | "Words" | "Double" | "Dotted" | "Thick" | "Dash" | "DotDash" | "DotDotDash" | "Wavy" | "WavyHeavy" | "DottedHeavy" | "DashHeavy" | "DotDashHeavy" | "DotDotDashHeavy" | "DashLong" | "DashLongHeavy" | "WavyDouble";
```

#### Property Value
[Word.Underline](https://learn.microsoft.com/en-us/javascript/api/word/word.underline) | "None" | "Single" | "Words" | "Double" | "Dotted" | "Thick" | "Dash" | "DotDash" | "DotDotDash" | "Wavy" | "WavyHeavy" | "DottedHeavy" | "DashHeavy" | "DotDashHeavy" | "DotDotDashHeavy" | "DashLong" | "DashLongHeavy" | "WavyDouble"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]