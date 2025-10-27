# Word.Template

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a document template.

## Properties

### buildingBlockEntries

**Type:** `Word.BuildingBlockEntryCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BuildingBlockEntryCollection object that represents the collection of building block entries in the template.

#### Examples

**Example**: Get all building block entries from the current document's template and log their names to the console.

```typescript
await Word.run(async (context) => {
    // Get the template of the current document
    const template = context.document.getTemplate();
    
    // Get the building block entries collection from the template
    const buildingBlockEntries = template.buildingBlockEntries;
    
    // Load the name property for each building block entry
    buildingBlockEntries.load("items/name");
    
    await context.sync();
    
    // Log the names of all building block entries
    console.log("Building Block Entries in Template:");
    buildingBlockEntries.items.forEach((entry) => {
        console.log(`- ${entry.name}`);
    });
});
```

---

### buildingBlockTypes

**Type:** `Word.BuildingBlockTypeItemCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BuildingBlockTypeItemCollection object that represents the collection of building block types that are contained in the template.

#### Examples

**Example**: Get the names of all building block types available in the current document's template

```typescript
await Word.run(async (context) => {
    const template = context.document.getActiveTemplate();
    const buildingBlockTypes = template.buildingBlockTypes;
    buildingBlockTypes.load("items/name");
    
    await context.sync();
    
    console.log("Available building block types:");
    buildingBlockTypes.items.forEach(type => {
        console.log(`- ${type.name}`);
    });
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Template object to verify the connection between the add-in and Word application before performing template operations.

```typescript
await Word.run(async (context) => {
    // Get the current document's template
    const template = context.document.properties.template;
    
    // Access the request context associated with the template
    const templateContext = template.context;
    
    // Use the context to load template properties
    template.load("name");
    
    await templateContext.sync();
    
    console.log("Template context is connected to Word");
    console.log("Template name: " + template.name);
});
```

---

### farEastLineBreakLanguage

**Type:** `Word.FarEastLineBreakLanguageId | "TraditionalChinese" | "Japanese" | "Korean" | "SimplifiedChinese"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the East Asian language to use when breaking lines of text in the document or template.

#### Examples

**Example**: Set the document template's line break language to Japanese for proper East Asian text formatting

```typescript
await Word.run(async (context) => {
    const template = context.document.body.styleTemplates;
    
    // Set the Far East line break language to Japanese
    context.document.body.style = "Normal";
    const docTemplate = context.document.properties.template;
    
    // Access template and set Far East line break language
    const settings = context.document;
    settings.properties.farEastLineBreakLanguage = Word.FarEastLineBreakLanguageId.japanese;
    // Or use string literal:
    // settings.properties.farEastLineBreakLanguage = "Japanese";
    
    await context.sync();
    
    console.log("Far East line break language set to Japanese");
});
```

---

### farEastLineBreakLevel

**Type:** `Word.FarEastLineBreakLevel | "Normal" | "Strict" | "Custom"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the line break control level for the document.

#### Examples

**Example**: Set the document's Far East line break control level to "Strict" to enforce stricter line breaking rules for Asian languages

```typescript
await Word.run(async (context) => {
    // Get the document template
    const template = context.document.body.style.template;
    
    // Set the Far East line break level to Strict
    template.farEastLineBreakLevel = Word.FarEastLineBreakLevel.strict;
    
    await context.sync();
    
    console.log("Far East line break level set to Strict");
});
```

---

### fullName

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the name of the template, including the drive or Web path.

#### Examples

**Example**: Get and display the full path and name of the current document's template

```typescript
await Word.run(async (context) => {
    // Get the template of the current document
    const template = context.document.template;
    
    // Load the fullName property
    template.load("fullName");
    
    await context.sync();
    
    // Display the full template path and name
    console.log("Template full path: " + template.fullName);
});
```

---

### hasNoProofing

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the spelling and grammar checker ignores documents based on this template.

#### Examples

**Example**: Check if spell-checking is disabled for the current document's template and display the result in the console

```typescript
await Word.run(async (context) => {
    const template = context.document.properties.template;
    template.load("hasNoProofing");
    
    await context.sync();
    
    console.log(`Spell-checking disabled: ${template.hasNoProofing}`);
});
```

---

### justificationMode

**Type:** `Word.JustificationMode | "Expand" | "Compress" | "CompressKana"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the character spacing adjustment for the template.

#### Examples

**Example**: Set the template's character spacing adjustment to compress text for a more compact document layout

```typescript
await Word.run(async (context) => {
    const template = context.document.body.styleTemplates.getByNameOrNullObject("Normal");
    template.load("justificationMode");
    await context.sync();
    
    // Set the justification mode to compress characters
    template.justificationMode = Word.JustificationMode.compress;
    // Or use the string literal: template.justificationMode = "Compress";
    
    await context.sync();
    console.log("Template character spacing set to compress mode");
});
```

---

### kerningByAlgorithm

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if Microsoft Word kerns half-width Latin characters and punctuation marks in the document.

#### Examples

**Example**: Enable kerning for half-width Latin characters and punctuation marks in the document template

```typescript
await Word.run(async (context) => {
    // Get the document template
    const template = context.document.body.template;
    
    // Enable kerning for half-width Latin characters and punctuation
    template.kerningByAlgorithm = true;
    
    // Sync to apply the changes
    await context.sync();
    
    console.log("Kerning by algorithm has been enabled for the document.");
});
```

---

### languageId

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a LanguageId value that represents the language in the template.

#### Examples

**Example**: Set the template's language to French Canadian for a document being prepared for Quebec-based users

```typescript
await Word.run(async (context) => {
    const template = context.document.body.insertTemplate();
    template.languageId = "FrenchCanadian";
    
    await context.sync();
    console.log("Template language set to French Canadian");
});
```

---

### languageIdFarEast

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies an East Asian language for the language in the template.

#### Examples

**Example**: Set the template's East Asian language to Japanese for proper text formatting and spell-checking of Japanese content.

```typescript
await Word.run(async (context) => {
    const template = context.document.body.styleOrNullObject.template;
    template.languageIdFarEast = "Japanese";
    
    await context.sync();
    console.log("Template East Asian language set to Japanese");
});
```

---

### name

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns only the name of the document template (excluding any path or other location information).

#### Examples

**Example**: Display the current document template name in a message to the user

```typescript
await Word.run(async (context) => {
    const template = context.document.properties.template;
    template.load("name");
    
    await context.sync();
    
    console.log(`Current template name: ${template.name}`);
});
```

---

### noLineBreakAfter

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the kinsoku characters after which Microsoft Word will not break a line.

#### Examples

**Example**: Set the kinsoku characters after which Word will not break a line to include common Japanese punctuation marks

```typescript
await Word.run(async (context) => {
    // Get the document template
    const template = context.document.body.style.template;
    
    // Set characters that should not appear at the end of a line
    template.noLineBreakAfter = "([{「『（［｛〈《【〔〖〘〚";
    
    // Sync to apply the changes
    await context.sync();
    
    console.log("No line break after characters set successfully");
});
```

---

### noLineBreakBefore

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the kinsoku characters before which Microsoft Word will not break a line.

#### Examples

**Example**: Set the kinsoku characters before which Word will not break a line to include common Japanese punctuation marks

```typescript
await Word.run(async (context) => {
    const template = context.document.body.style.template;
    
    // Set characters that should not appear at the beginning of a line
    template.noLineBreakBefore = "、。，．）」』】";
    
    await context.sync();
    
    console.log("No line break before characters set successfully");
});
```

---

### path

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the path to the document template.

#### Examples

**Example**: Display the file path of the current document's template in the console

```typescript
await Word.run(async (context) => {
    const template = context.document.properties.template;
    template.load("path");
    
    await context.sync();
    
    console.log("Template path: " + template.path);
});
```

---

### saved

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies true if the template has not changed since it was last saved, false if Microsoft Word displays a prompt to save changes when the document is closed.

#### Examples

**Example**: Check if the template has unsaved changes and display the save status to the user

```typescript
await Word.run(async (context) => {
    const template = context.document.properties.template;
    template.load("saved");
    
    await context.sync();
    
    if (template.saved) {
        console.log("Template has no unsaved changes");
    } else {
        console.log("Template has unsaved changes");
    }
});
```

---

### type

**Type:** `Word.TemplateType | "Normal" | "Global" | "Attached"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the template type.

#### Examples

**Example**: Check if the current document's template is a "Normal" template and display the template type in the console.

```typescript
await Word.run(async (context) => {
    const template = context.document.properties.template;
    template.load("type");
    
    await context.sync();
    
    console.log(`Template type: ${template.type}`);
    
    if (template.type === "Normal" || template.type === Word.TemplateType.normal) {
        console.log("This document uses the Normal template.");
    }
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TemplateLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Template`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Template`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Template`

#### Examples

**Example**: Load and display the template's name and type properties from a Word document template

```typescript
await Word.run(async (context) => {
    // Get the template object
    const template = context.document.properties.template;
    
    // Load specific properties of the template
    template.load("name, type");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now the properties are available to read
    console.log("Template Name: " + template.name);
    console.log("Template Type: " + template.type);
});
```

---

### save

**Kind:** `write`

Saves the template.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Save changes made to a document template after modifying its properties

```typescript
await Word.run(async (context) => {
    // Get the template
    const template = context.document.template;
    
    // Load template properties
    template.load("name");
    await context.sync();
    
    // Make changes to the template (example: modify content)
    context.document.body.insertParagraph("Template content updated", Word.InsertLocation.end);
    
    // Save the template
    template.save();
    
    await context.sync();
    console.log("Template saved successfully");
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.TemplateUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Template` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple template properties at once, including its name and description

```typescript
await Word.run(async (context) => {
    const template = context.document.template;
    
    // Set multiple properties of the template at once
    template.set({
        name: "Company Report Template",
        description: "Standard template for quarterly reports"
    });
    
    await context.sync();
    console.log("Template properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Template object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TemplateData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.TemplateData`

#### Examples

**Example**: Serialize a template object to JSON format for logging or data transfer purposes

```typescript
await Word.run(async (context) => {
    // Get the template associated with the current document
    const template = context.document.properties.template;
    
    // Load the template properties
    template.load("name");
    
    await context.sync();
    
    // Convert the template object to a plain JavaScript object
    const templateData = template.toJSON();
    
    // Log the serialized template data
    console.log("Template data:", JSON.stringify(templateData, null, 2));
    
    // You can now use this plain object for storage, transmission, or comparison
    return templateData;
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Template`

#### Examples

**Example**: Track a template object across multiple sync calls to prevent InvalidObjectPath errors when accessing its properties after document changes

```typescript
await Word.run(async (context) => {
    const template = context.document.body.insertTemplate("MyTemplate", Word.InsertLocation.start);
    
    // Track the template object to maintain reference across sync calls
    template.track();
    
    // First sync to load initial data
    await context.sync();
    
    // Perform some operations that might change the document
    context.document.body.insertParagraph("New content", Word.InsertLocation.end);
    
    // Second sync - without track(), accessing template here might throw InvalidObjectPath
    await context.sync();
    
    // Safe to access template properties because it's being tracked
    template.load("name");
    await context.sync();
    
    console.log("Template name: " + template.name);
    
    // Clean up: untrack when done
    template.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Template`

#### Examples

**Example**: Get a template object, track it for performance optimization, use it to retrieve template information, then untrack it to free memory when done.

```typescript
await Word.run(async (context) => {
    // Get the template and track it
    const template = context.document.properties.template;
    template.track();
    template.load("name");
    
    await context.sync();
    
    // Use the template object
    console.log("Template name: " + template.name);
    
    // Untrack the template to release memory
    template.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.buildingblockentrycollection
- /en-us/javascript/api/word/word.buildingblocktypeitemcollection
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.fareastlinebreaklanguageid
- /en-us/javascript/api/word/word.fareastlinebreaklevel
- /en-us/javascript/api/word/word.justificationmode
- /en-us/javascript/api/word/word.languageid
- /en-us/javascript/api/word/word.templatetype
- /en-us/javascript/api/word/word.interfaces.templateloadoptions
- /en-us/javascript/api/word/word.template
- /en-us/javascript/api/office/officeextension.updateoptions
- /en-us/javascript/api/word/word.interfaces.templateupdatedata
- /en-us/javascript/api/word/word.interfaces.templatedata
- /en-us/javascript/api/office/officeextension.clientrequestcontext
