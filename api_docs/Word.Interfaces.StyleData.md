# Word.Interfaces.StyleData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling style.toJSON().

## Properties

- automaticallyUpdate
  - Specifies whether the style is automatically redefined based on the selection.
- baseStyle
  - Specifies the name of an existing style to use as the base formatting of another style.
- borders
  - Specifies a BorderCollection object that represents all the borders for the specified style.
- builtIn
  - Gets whether the specified style is a built-in style.
- description
  - Gets the description of the specified style.
- font
  - Gets a font object that represents the character formatting of the specified style.
- frame
  - Returns a Frame object that represents the frame formatting for the style.
- hasProofing
  - Specifies whether the spelling and grammar checker ignores text formatted with this style.
- inUse
  - Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.
- languageId
  - Specifies a LanguageId value that represents the language for the style.
- languageIdFarEast
  - Specifies an East Asian language for the style.
- linked
  - Gets whether a style is a linked style that can be used for both paragraph and character formatting.
- linkStyle
  - Specifies a link between a paragraph and a character style.
- listLevelNumber
  - Returns the list level for the style.
- listTemplate
  - Gets a ListTemplate object that represents the list formatting for the specified Style object.
- locked
  - Specifies whether the style cannot be changed or edited.
- nameLocal
  - Gets the name of a style in the language of the user.
- nextParagraphStyle
  - Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.
- noSpaceBetweenParagraphsOfSameStyle
  - Specifies whether to remove spacing between paragraphs that are formatted using the same style.
- paragraphFormat
  - Gets a ParagraphFormat object that represents the paragraph settings for the specified style.
- priority
  - Specifies the priority.
- quickStyle
  - Specifies whether the style corresponds to an available quick style.
- shading
  - Gets a Shading object that represents the shading for the specified style. Not applicable to List style.
- tableStyle
  - Gets a TableStyle object representing Style properties that can be applied to a table.
- type
  - Gets the style type.
- unhideWhenUsed
  - Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.
- visibility
  - Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.

## Property Details

### automaticallyUpdate

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the style is automatically redefined based on the selection.

```typescript
automaticallyUpdate?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### baseStyle

Specifies the name of an existing style to use as the base formatting of another style.

```typescript
baseStyle?: string;
```

Property Value
- string

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- Note: The ability to set baseStyle was introduced in WordApi 1.6.

### borders

Specifies a BorderCollection object that represents all the borders for the specified style.

```typescript
borders?: Word.Interfaces.BorderData[];
```

Property Value
- [Word.Interfaces.BorderData](/en-us/javascript/api/word/word.interfaces.borderdata)[]

Remarks
- [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### builtIn

Gets whether the specified style is a built-in style.

```typescript
builtIn?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### description

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the description of the specified style.

```typescript
description?: string;
```

Property Value
- string

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### font

Gets a font object that represents the character formatting of the specified style.

```typescript
font?: Word.Interfaces.FontData;
```

Property Value
- [Word.Interfaces.FontData](/en-us/javascript/api/word/word.interfaces.fontdata)

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### frame

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Frame object that represents the frame formatting for the style.

```typescript
frame?: Word.Interfaces.FrameData;
```

Property Value
- [Word.Interfaces.FrameData](/en-us/javascript/api/word/word.interfaces.framedata)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### hasProofing

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the spelling and grammar checker ignores text formatted with this style.

```typescript
hasProofing?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### inUse

Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.

```typescript
inUse?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### languageId

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LanguageId value that represents the language for the style.

```typescript
languageId?: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

Property Value
- [Word.LanguageId](/en-us/javascript/api/word/word.languageid) | all listed string literals above

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### languageIdFarEast

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an East Asian language for the style.

```typescript
languageIdFarEast?: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

Property Value
- [Word.LanguageId](/en-us/javascript/api/word/word.languageid) | all listed string literals above

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### linked

Gets whether a style is a linked style that can be used for both paragraph and character formatting.

```typescript
linked?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### linkStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a link between a paragraph and a character style.

```typescript
linkStyle?: Word.Interfaces.StyleData;
```

Property Value
- [Word.Interfaces.StyleData](/en-us/javascript/api/word/word.interfaces.styledata)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listLevelNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the list level for the style.

```typescript
listLevelNumber?: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listTemplate

Gets a ListTemplate object that represents the list formatting for the specified Style object.

```typescript
listTemplate?: Word.Interfaces.ListTemplateData;
```

Property Value
- [Word.Interfaces.ListTemplateData](/en-us/javascript/api/word/word.interfaces.listtemplatedata)

Remarks
- [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### locked

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the style cannot be changed or edited.

```typescript
locked?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nameLocal

Gets the name of a style in the language of the user.

```typescript
nameLocal?: string;
```

Property Value
- string

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nextParagraphStyle

Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.

```typescript
nextParagraphStyle?: string;
```

Property Value
- string

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- Note: The ability to set nextParagraphStyle was introduced in WordApi 1.6.

### noSpaceBetweenParagraphsOfSameStyle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove spacing between paragraphs that are formatted using the same style.

```typescript
noSpaceBetweenParagraphsOfSameStyle?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### paragraphFormat

Gets a ParagraphFormat object that represents the paragraph settings for the specified style.

```typescript
paragraphFormat?: Word.Interfaces.ParagraphFormatData;
```

Property Value
- [Word.Interfaces.ParagraphFormatData](/en-us/javascript/api/word/word.interfaces.paragraphformatdata)

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### priority

Specifies the priority.

```typescript
priority?: number;
```

Property Value
- number

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### quickStyle

Specifies whether the style corresponds to an available quick style.

```typescript
quickStyle?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shading

Gets a Shading object that represents the shading for the specified style. Not applicable to List style.

```typescript
shading?: Word.Interfaces.ShadingData;
```

Property Value
- [Word.Interfaces.ShadingData](/en-us/javascript/api/word/word.interfaces.shadingdata)

Remarks
- [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tableStyle

Gets a TableStyle object representing Style properties that can be applied to a table.

```typescript
tableStyle?: Word.Interfaces.TableStyleData;
```

Property Value
- [Word.Interfaces.TableStyleData](/en-us/javascript/api/word/word.interfaces.tablestyledata)

Remarks
- [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Gets the style type.

```typescript
type?: Word.StyleType | "Character" | "List" | "Paragraph" | "Table";
```

Property Value
- [Word.StyleType](/en-us/javascript/api/word/word.styletype) | "Character" | "List" | "Paragraph" | "Table"

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### unhideWhenUsed

Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.

```typescript
unhideWhenUsed?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### visibility

Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.

```typescript
visibility?: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)