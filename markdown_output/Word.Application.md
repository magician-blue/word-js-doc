# Word.Application

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the application object.

## Class Examples

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

### bibliography

**Type:** `Word.Bibliography`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a `Bibliography` object that represents the bibliography reference sources stored in Microsoft Word.

#### Examples

**Example**: Get the count of bibliography sources currently stored in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    const bibliography = context.application.bibliography;
    const sources = bibliography.sources;
    sources.load("items");
    
    await context.sync();
    
    console.log(`Total bibliography sources: ${sources.items.length}`);
});
```

---

### checkLanguage

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if Microsoft Word automatically detects the language you are using as you type.

#### Examples

**Example**: Disable automatic language detection while typing in the Word document

```typescript
await Word.run(async (context) => {
    // Disable automatic language detection
    context.application.checkLanguage = false;
    
    await context.sync();
    console.log("Automatic language detection has been disabled.");
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context to check the Office host application's connection status and sync changes to the document

```typescript
await Word.run(async (context) => {
    // Access the application object
    const app = context.application;
    
    // The context property connects the add-in to the Office host
    const requestContext = app.context;
    
    // Use the context to load application properties
    app.load("name");
    
    // Sync the context to execute queued commands and retrieve data
    await requestContext.sync();
    
    console.log("Connected to: " + app.name);
    
    // Make changes to the document
    const body = context.document.body;
    body.insertText("Context synced successfully!", Word.InsertLocation.end);
    
    // Sync again to apply changes
    await requestContext.sync();
});
```

---

### language

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a `LanguageId` value that represents the language selected for the Microsoft Word user interface.

#### Examples

**Example**: Display the current language of the Word user interface in a message to the user

```typescript
await Word.run(async (context) => {
    const application = context.application;
    application.load("language");
    
    await context.sync();
    
    console.log(`Current Word UI language: ${application.language}`);
    
    // You can also use the language value for conditional logic
    if (application.language === "EnglishUS" || application.language === "EnglishUK") {
        console.log("English language detected");
    }
});
```

---

### templates

**Type:** `Word.TemplateCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a `TemplateCollection` object that represents all the available templates: global templates and those attached to open documents.

#### Examples

**Example**: Get the count of all available templates (global templates and those attached to open documents) and display it in the console.

```typescript
await Word.run(async (context) => {
    const templates = context.application.templates;
    templates.load("count");
    
    await context.sync();
    
    console.log(`Total number of available templates: ${templates.count}`);
});
```

---

## Methods

### createDocument

**Kind:** `create`

Creates a new document by using an optional Base64-encoded .docx file.

#### Signature

**Parameters:**
- `base64File`: `string` (optional)
  Optional. The Base64-encoded .docx file. The default value is null.

**Returns:** `Word.DocumentCreated`

#### Examples

**Example**: Insert the text content from an external Base64-encoded Word document at the beginning of the current document's body.

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

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.ApplicationLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Application`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Application`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Application`

#### Examples

**Example**: Load and display the application's name and version properties

```typescript
await Word.run(async (context) => {
    const app = context.application;
    
    // Load specific properties of the application
    app.load("name, version");
    
    await context.sync();
    
    console.log(`Application: ${app.name}`);
    console.log(`Version: ${app.version}`);
});
```

---

### newObject

**Kind:** `create`

Create a new instance of the `Word.Application` object.

#### Signature

**Parameters:**
- `context`: `OfficeExtension.ClientRequestContext` (required)

**Returns:** `Word.Application`

#### Examples

**Example**: Access the application object to check the Word application's name property

```typescript
await Word.run(async (context) => {
    const app = context.application;
    app.load("name");
    
    await context.sync();
    
    console.log(`Application name: ${app.name}`);
});
```

---

### openDocument

Opens a document and displays it in a new tab or window. The following are examples for the various supported clients and platforms.
- Remote or cloud location example: `https://microsoft.sharepoint.com/some/path/Document.docx`
- Local location examples for Windows: `C:\Users\Someone\Documents\Document.docx` (includes required escaped backslashes), `file://mycomputer/myfolder/Document.docx`
- Local location example for Mac and iOS: `/User/someone/document.docx`

#### Signature

**Parameters:**
- `filePath`: `string` (required)
  Required. The absolute path of the .docx file. Word on the web only supports remote (cloud) locations, while Word on Windows, on Mac, and on iOS support local and remote locations.

**Returns:** `void`

#### Examples

**Example**: Open a document from a SharePoint location and display it in a new tab

```typescript
await Word.run(async (context) => {
    const app = context.application;
    
    // Open a document from SharePoint
    app.openDocument("https://microsoft.sharepoint.com/sites/team/Documents/Report.docx");
    
    await context.sync();
});
```

---

### retrieveStylesFromBase64

Parse styles from template Base64 file and return JSON format of retrieved styles as a string.

#### Signature

**Parameters:**
- `base64File`: `string` (required)
  Required. The template file.

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Retrieve style definitions from an external Word document provided as a Base64-encoded string and log them to the console.

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

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ApplicationUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Application` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple application-level settings at once, including showing the task pane and enabling screen updating

```typescript
await Word.run(async (context) => {
    const app = context.application;
    
    // Set multiple application properties at once
    app.set({
        showTaskpane: true,
        screenUpdating: true
    });
    
    await context.sync();
    console.log("Application settings updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Application` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ApplicationData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ApplicationData`

#### Examples

**Example**: Serialize the Word application object to JSON format to inspect its properties and log them to the console.

```typescript
await Word.run(async (context) => {
    // Get the application object
    const application = context.application;
    
    // Load properties you want to serialize
    application.load("name");
    
    await context.sync();
    
    // Convert the application object to a plain JavaScript object
    const applicationJSON = application.toJSON();
    
    // Log the serialized object
    console.log("Application as JSON:", JSON.stringify(applicationJSON, null, 2));
    console.log("Application name:", applicationJSON.name);
});
```

---

## Source

- /en-us/javascript/api/word
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-external-document.yaml
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/get-external-styles.yaml
