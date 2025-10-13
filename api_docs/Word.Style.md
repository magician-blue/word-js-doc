# Word.Style class

Package: [word](/en-us/javascript/api/word)

Represents a style in a Word document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Applies the specified style to a paragraph.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to apply.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else if (style.type != Word.StyleType.paragraph) {
    console.log(`The '${styleName}' style isn't a paragraph style.`);
  } else {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    const paragraph: Word.Paragraph = body.paragraphs.getFirst();
    paragraph.style = style.nameLocal;
    console.log(`'${styleName}' style applied to first paragraph.`);
  }
});
```

## Properties

| Property | Description |
| --- | --- |
| [automaticallyUpdate](#automaticallyupdate) | Specifies whether the style is automatically redefined based on the selection. |
| [baseStyle](#basestyle) | Specifies the name of an existing style to use as the base formatting of another style. |
| [borders](#borders) | Specifies a BorderCollection object that represents all the borders for the specified style. |
| [builtIn](#builtin) | Gets whether the specified style is a built-in style. |
| [context](#context) | The request context associated with the object. This connects the add-in's process to the Office host application's process. |
| [description](#description) | Gets the description of the specified style. |
| [font](#font) | Gets a font object that represents the character formatting of the specified style. |
| [frame](#frame) | Returns a Frame object that represents the frame formatting for the style. |
| [hasProofing](#hasproofing) | Specifies whether the spelling and grammar checker ignores text formatted with this style. |
| [inUse](#inuse) | Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document. |
| [languageId](#languageid) | Specifies a LanguageId value that represents the language for the style. |
| [languageIdFarEast](#languageidfareast) | Specifies an East Asian language for the style. |
| [linked](#linked) | Gets whether a style is a linked style that can be used for both paragraph and character formatting. |
| [linkStyle](#linkstyle) | Specifies a link between a paragraph and a character style. |
| [listLevelNumber](#listlevelnumber) | Returns the list level for the style. |
| [listTemplate](#listtemplate) | Gets a ListTemplate object that represents the list formatting for the specified Style object. |
| [locked](#locked) | Specifies whether the style cannot be changed or edited. |
| [nameLocal](#namelocal) | Gets the name of a style in the language of the user. |
| [nextParagraphStyle](#nextparagraphstyle) | Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style. |
| [noSpaceBetweenParagraphsOfSameStyle](#nospacebetweenparagraphsofsamestyle) | Specifies whether to remove spacing between paragraphs that are formatted using the same style. |
| [paragraphFormat](#paragraphformat) | Gets a ParagraphFormat object that represents the paragraph settings for the specified style. |
| [priority](#priority) | Specifies the priority. |
| [quickStyle](#quickstyle) | Specifies whether the style corresponds to an available quick style. |
| [shading](#shading) | Gets a Shading object that represents the shading for the specified style. Not applicable to List style. |
| [tableStyle](#tablestyle) | Gets a TableStyle object representing Style properties that can be applied to a table. |
| [type](#type) | Gets the style type. |
| [unhideWhenUsed](#unhidewhenused) | Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document. |
| [visibility](#visibility) | Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane. |

## Methods

| Method | Description |
| --- | --- |
| [delete()](#delete) | Deletes the style. |
| [linkToListTemplate(listTemplate)](#linktolisttemplatelisttemplate) | Links this style to a list template so that the style's formatting can be applied to lists. |
| [load(options)](#loadoptions) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [load(propertyNames)](#loadpropertynames) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [load(propertyNamesAndPaths)](#loadpropertynamesandpaths) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [set(properties, options)](#setproperties-options) | Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type. |
| [set(properties)](#setproperties) | Sets multiple properties on the object at the same time, based on an existing loaded object. |
| [toJSON()](#tojson) | Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify(). |
| [track()](#track) | Track the object for automatic adjustment based on surrounding changes in the document. |
| [untrack()](#untrack) | Release the memory associated with this object, if it has previously been tracked. |

## Property Details

### automaticallyUpdate
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the style is automatically redefined based on the selection.

```typescript
automaticallyUpdate: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### baseStyle
Specifies the name of an existing style to use as the base formatting of another style.

```typescript
baseStyle: string;
```

Property value: string

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)  
Note: The ability to set baseStyle was introduced in WordApi 1.6.

### borders
Specifies a BorderCollection object that represents all the borders for the specified style.

```typescript
readonly borders: Word.BorderCollection;
```

Property value: [Word.BorderCollection](/en-us/javascript/api/word/word.bordercollection)

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Updates border properties (e.g., type, width, color) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update border properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const borders: Word.BorderCollection = style.borders;
    borders.load("items");
    await context.sync();

    borders.outsideBorderType = Word.BorderType.dashed;
    borders.outsideBorderWidth = Word.BorderWidth.pt025;
    borders.outsideBorderColor = "green";
    console.log("Updated outside borders.");
  }
});
```

### builtIn
Gets whether the specified style is a built-in style.

```typescript
readonly builtIn: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### description
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the description of the specified style.

```typescript
readonly description: string;
```

Property value: string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### font
Gets a font object that represents the character formatting of the specified style.

```typescript
readonly font: Word.Font;
```

Property value: [Word.Font](/en-us/javascript/api/word/word.font)

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Updates font properties (e.g., color, size) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update font properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const font: Word.Font = style.font;
    font.color = "#FF0000";
    font.size = 20;
    console.log(`Successfully updated font properties of the '${styleName}' style.`);
  }
});
```

### frame
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Frame object that represents the frame formatting for the style.

```typescript
readonly frame: Word.Frame;
```

Property value: [Word.Frame](/en-us/javascript/api/word/word.frame)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### hasProofing
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the spelling and grammar checker ignores text formatted with this style.

```typescript
hasProofing: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### inUse
Gets whether the specified style is a built-in style that has been modified or applied in the document or a new style that has been created in the document.

```typescript
readonly inUse: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### languageId
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a LanguageId value that represents the language for the style.

```typescript
languageId: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

Property value: [Word.LanguageId](/en-us/javascript/api/word/word.languageid) | full string union as above

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### languageIdFarEast
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an East Asian language for the style.

```typescript
languageIdFarEast: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

Property value: [Word.LanguageId](/en-us/javascript/api/word/word.languageid) | full string union as above

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### linked
Gets whether a style is a linked style that can be used for both paragraph and character formatting.

```typescript
readonly linked: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### linkStyle
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a link between a paragraph and a character style.

```typescript
linkStyle: Word.Style;
```

Property value: [Word.Style](/en-us/javascript/api/word/word.style)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listLevelNumber
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the list level for the style.

```typescript
readonly listLevelNumber: number;
```

Property value: number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listTemplate
Gets a ListTemplate object that represents the list formatting for the specified Style object.

```typescript
readonly listTemplate: Word.ListTemplate;
```

Property value: [Word.ListTemplate](/en-us/javascript/api/word/word.listtemplate)

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/manage-list-styles.yaml

// Gets the properties of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to get properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load("type");
  await context.sync();

  if (style.isNullObject || style.type != Word.StyleType.list) {
    console.warn(`There's no existing style with the name '${styleName}'. Or this isn't a list style.`);
  } else {
    // Load objects to log properties and their values in the console.
    style.load();
    style.listTemplate.load();
    await context.sync();

    console.log(`Properties of the '${styleName}' style:`, style);

    const listLevels = style.listTemplate.listLevels;
    listLevels.load("items");
    await context.sync();

    console.log(`List levels of the '${styleName}' style:`, listLevels);
  }
});
```

### locked
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the style cannot be changed or edited.

```typescript
locked: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### nameLocal
Gets the name of a style in the language of the user.

```typescript
readonly nameLocal: string;
```

Property value: string

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Applies the specified style to a paragraph.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to apply.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else if (style.type != Word.StyleType.paragraph) {
    console.log(`The '${styleName}' style isn't a paragraph style.`);
  } else {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    const paragraph: Word.Paragraph = body.paragraphs.getFirst();
    paragraph.style = style.nameLocal;
    console.log(`'${styleName}' style applied to first paragraph.`);
  }
});
```

### nextParagraphStyle
Specifies the name of the style to be applied automatically to a new paragraph that is inserted after a paragraph formatted with the specified style.

```typescript
nextParagraphStyle: string;
```

Property value: string

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)  
Note: The ability to set nextParagraphStyle was introduced in WordApi 1.6.

### noSpaceBetweenParagraphsOfSameStyle
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove spacing between paragraphs that are formatted using the same style.

```typescript
noSpaceBetweenParagraphsOfSameStyle: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### paragraphFormat
Gets a ParagraphFormat object that represents the paragraph settings for the specified style.

```typescript
readonly paragraphFormat: Word.ParagraphFormat;
```

Property value: [Word.ParagraphFormat](/en-us/javascript/api/word/word.paragraphformat)

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Sets certain aspects of the specified style's paragraph format e.g., the left indent size and the alignment.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update its paragraph format.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.paragraphFormat.leftIndent = 30;
    style.paragraphFormat.alignment = Word.Alignment.centered;
    console.log(`Successfully the paragraph format of the '${styleName}' style.`);
  }
});
```

### priority
Specifies the priority.

```typescript
priority: number;
```

Property value: number

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### quickStyle
Specifies whether the style corresponds to an available quick style.

```typescript
quickStyle: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shading
Gets a Shading object that represents the shading for the specified style. Not applicable to List style.

```typescript
readonly shading: Word.Shading;
```

Property value: [Word.Shading](/en-us/javascript/api/word/word.shading)

Remarks  
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Updates shading properties (e.g., texture, pattern colors) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update shading properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const shading: Word.Shading = style.shading;
    shading.load();
    await context.sync();

    shading.backgroundPatternColor = "blue";
    shading.foregroundPatternColor = "yellow";
    shading.texture = Word.ShadingTextureType.darkTrellis;

    console.log("Updated shading.");
  }
});
```

### tableStyle
Gets a TableStyle object representing Style properties that can be applied to a table.

```typescript
readonly tableStyle: Word.TableStyle;
```

Property value: [Word.TableStyle](/en-us/javascript/api/word/word.tablestyle)

Remarks  
[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type
Gets the style type.

```typescript
readonly type: Word.StyleType | "Character" | "List" | "Paragraph" | "Table";
```

Property value: [Word.StyleType](/en-us/javascript/api/word/word.styletype) | "Character" | "List" | "Paragraph" | "Table"

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### unhideWhenUsed
Specifies whether the specified style is made visible as a recommended style in the Styles and in the Styles task pane in Microsoft Word after it's used in the document.

```typescript
unhideWhenUsed: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### visibility
Specifies whether the specified style is visible as a recommended style in the Styles gallery and in the Styles task pane.

```typescript
visibility: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### delete()
Deletes the style.

```typescript
delete(): void;
```

Returns: void

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Deletes the custom style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to delete.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    style.delete();
    console.log(`Successfully deleted custom style '${styleName}'.`);
  }
});
```

### linkToListTemplate(listTemplate)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Links this style to a list template so that the style's formatting can be applied to lists.

```typescript
linkToListTemplate(listTemplate: Word.ListTemplate): void;
```

Parameters
- listTemplate: [Word.ListTemplate](/en-us/javascript/api/word/word.listtemplate)  
  A ListTemplate to link to the style.

Returns: void

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.StyleLoadOptions): Word.Style;
```

Parameters
- options: [Word.Interfaces.StyleLoadOptions](/en-us/javascript/api/word/word.interfaces.styleloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.Style](/en-us/javascript/api/word/word.style)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Style;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.Style](/en-us/javascript/api/word/word.style)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Style;
```

Parameters
- propertyNamesAndPaths:  
  - select?: string  
  - expand?: string

propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.Style](/en-us/javascript/api/word/word.style)

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.StyleUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.StyleUpdateData](/en-us/javascript/api/word/word.interfaces.styleupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Style): void;
```

Parameters
- properties: [Word.Style](/en-us/javascript/api/word/word.style)

Returns: void

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). JSON.stringify, in turn, calls the toJSON method of the object that's passed to it. Whereas the original Word.Style object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.StyleData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.StyleData;
```

Returns: [Word.Interfaces.StyleData](/en-us/javascript/api/word/word.interfaces.styledata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Style;
```

Returns: [Word.Style](/en-us/javascript/api/word/word.style)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Style;
```

Returns: [Word.Style](/en-us/javascript/api/word/word.style)