# Word.Application class

Package: [word](/en-us/javascript/api/word)

Represents the application object.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-external-document.yaml

// Updates the text of the current document with the text from another document passed in as a Base64-encoded string.
await Word.run(async (context) => {
  // Use the Base64-encoded string representation of the selected .docx file.
  const externalDoc: Word.DocumentCreated = context.application.createDocument(externalDocument);
  await context.sync();

  if (!Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
    console.warn("The WordApiHiddenDocument 1.3 requirement set isn't supported on this client so can't proceed. Try this action on a platform that supports this requirement set.");
    return;
  }

  const externalDocBody: Word.Body = externalDoc.body;
  externalDocBody.load("text");
  await context.sync();

  // Insert the external document's text at the beginning of the current document's body.
  const externalDocBodyText = externalDocBody.text;
  const currentDocBody: Word.Body = context.document.body;
  currentDocBody.insertText(externalDocBodyText, Word.InsertLocation.start);
  await context.sync();
});
```

## Properties

- [bibliography](#word-word-application-bibliography-member)  
  Returns a `Bibliography` object that represents the bibliography reference sources stored in Microsoft Word.
- [checkLanguage](#word-word-application-checklanguage-member)  
  Specifies if Microsoft Word automatically detects the language you are using as you type.
- [context](#word-word-application-context-member)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [language](#word-word-application-language-member)  
  Gets a `LanguageId` value that represents the language selected for the Microsoft Word user interface.
- [templates](#word-word-application-templates-member)  
  Returns a `TemplateCollection` object that represents all the available templates: global templates and those attached to open documents.

## Methods

- [createDocument(base64File)](#word-word-application-createdocument-member1)  
  Creates a new document by using an optional Base64-encoded .docx file.
- [load(options)](#word-word-application-load-member1)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#word-word-application-load-member2)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-application-load-member3)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [newObject(context)](#word-word-application-newobject-member1)  
  Create a new instance of the `Word.Application` object.
- [openDocument(filePath)](#word-word-application-opendocument-member1)  
  Opens a document and displays it in a new tab or window. Examples for supported clients and platforms:
  - Remote or cloud location example: `https://microsoft.sharepoint.com/some/path/Document.docx`
  - Local location examples for Windows: `C:\\Users\\Someone\\Documents\\Document.docx` (includes required escaped backslashes), `file://mycomputer/myfolder/Document.docx`
  - Local location example for Mac and iOS: `/User/someone/document.docx`
- [retrieveStylesFromBase64(base64File)](#word-word-application-retrievestylesfrombase64-member1)  
  Parse styles from template Base64 file and return JSON format of retrieved styles as a string.
- [set(properties, options)](#word-word-application-set-member1)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#word-word-application-set-member2)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- [toJSON()](#word-word-application-tojson-member1)  
  Overrides the JavaScript `toJSON()` method to provide more useful output when an API object is passed to `JSON.stringify()`.

## Property Details

### bibliography
Returns a `Bibliography` object that represents the bibliography reference sources stored in Microsoft Word.

```typescript
readonly bibliography: Word.Bibliography;
```

Property Value: [Word.Bibliography](/en-us/javascript/api/word/word.bibliography)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### checkLanguage
Specifies if Microsoft Word automatically detects the language you are using as you type.

```typescript
checkLanguage: boolean;
```

Property Value: boolean

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### language
Gets a `LanguageId` value that represents the language selected for the Microsoft Word user interface.

```typescript
readonly language: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

Property Value: [Word.LanguageId](/en-us/javascript/api/word/word.languageid) | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### templates
Returns a `TemplateCollection` object that represents all the available templates: global templates and those attached to open documents.

```typescript
readonly templates: Word.TemplateCollection;
```

Property Value: [Word.TemplateCollection](/en-us/javascript/api/word/word.templatecollection)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

## Method Details

### createDocument(base64File)
Creates a new document by using an optional Base64-encoded .docx file.

```typescript
createDocument(base64File?: string): Word.DocumentCreated;
```

Parameters:
- base64File: string  
  Optional. The Base64-encoded .docx file. The default value is null.

Returns: [Word.DocumentCreated](/en-us/javascript/api/word/word.documentcreated)

Remarks: [API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-external-document.yaml

// Updates the text of the current document with the text from another document passed in as a Base64-encoded string.
await Word.run(async (context) => {
  // Use the Base64-encoded string representation of the selected .docx file.
  const externalDoc: Word.DocumentCreated = context.application.createDocument(externalDocument);
  await context.sync();

  if (!Office.context.requirements.isSetSupported("WordApiHiddenDocument", "1.3")) {
    console.warn("The WordApiHiddenDocument 1.3 requirement set isn't supported on this client so can't proceed. Try this action on a platform that supports this requirement set.");
    return;
  }

  const externalDocBody: Word.Body = externalDoc.body;
  externalDocBody.load("text");
  await context.sync();

  // Insert the external document's text at the beginning of the current document's body.
  const externalDocBodyText = externalDocBody.text;
  const currentDocBody: Word.Body = context.document.body;
  currentDocBody.insertText(externalDocBodyText, Word.InsertLocation.start);
  await context.sync();
});
```

---

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ApplicationLoadOptions): Word.Application;
```

Parameters:
- options: [Word.Interfaces.ApplicationLoadOptions](/en-us/javascript/api/word/word.interfaces.applicationloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.Application](/en-us/javascript/api/word/word.application)

---

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Application;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.Application](/en-us/javascript/api/word/word.application)

---

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Application;
```

Parameters:
- propertyNamesAndPaths:  
  ```
  {
    select?: string;
    expand?: string;
  }
  ```
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.Application](/en-us/javascript/api/word/word.application)

---

### newObject(context)
Create a new instance of the `Word.Application` object.

```typescript
static newObject(context: OfficeExtension.ClientRequestContext): Word.Application;
```

Parameters:
- context: [OfficeExtension.ClientRequestContext](/en-us/javascript/api/office/officeextension.clientrequestcontext)

Returns: [Word.Application](/en-us/javascript/api/word/word.application)

---

### openDocument(filePath)
Opens a document and displays it in a new tab or window. The following are examples for the various supported clients and platforms.
- Remote or cloud location example: `https://microsoft.sharepoint.com/some/path/Document.docx`
- Local location examples for Windows: `C:\\Users\\Someone\\Documents\\Document.docx` (includes required escaped backslashes), `file://mycomputer/myfolder/Document.docx`
- Local location example for Mac and iOS: `/User/someone/document.docx`

```typescript
openDocument(filePath: string): void;
```

Parameters:
- filePath: string  
  Required. The absolute path of the .docx file. Word on the web only supports remote (cloud) locations, while Word on Windows, on Mac, and on iOS support local and remote locations.

Returns: void

Remarks: [API set: WordApi 1.6]

---

### retrieveStylesFromBase64(base64File)
Parse styles from template Base64 file and return JSON format of retrieved styles as a string.

```typescript
retrieveStylesFromBase64(base64File: string): OfficeExtension.ClientResult<string>;
```

Parameters:
- base64File: string  
  Required. The template file.

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks: [API set: WordApi 1.5]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/get-external-styles.yaml

// Gets style info from another document passed in as a Base64-encoded string.
await Word.run(async (context) => {
  const retrievedStyles = context.application.retrieveStylesFromBase64(externalDocument);
  await context.sync();

  console.log("Styles from the other document:", retrievedStyles.value);
});
```

---

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ApplicationUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.ApplicationUpdateData](/en-us/javascript/api/word/word.interfaces.applicationupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

---

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Application): void;
```

Parameters:
- properties: [Word.Application](/en-us/javascript/api/word/word.application)

Returns: void

---

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Application` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ApplicationData;
```

Returns: [Word.Interfaces.ApplicationData](/en-us/javascript/api/word/word.interfaces.applicationdata)