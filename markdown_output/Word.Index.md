# Word.Index

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a single index. The Index object is a member of the Word.IndexCollection. The IndexCollection includes all the indexes in the document.

## Properties

### context

**Type:** `RequestContext`

**Since:** WordApi BETA (PREVIEW ONLY)

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from an Index object to synchronize changes and load properties for a document index.

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.getIndexes();
    const firstIndex = indexes.getFirst();
    
    // Load the heading property using the index's context
    firstIndex.load('heading');
    
    // Sync using the context from the index object
    await firstIndex.context.sync();
    
    // Access the loaded property
    console.log("Index heading: " + firstIndex.heading);
    
    // The index.context property connects to the same context as the Word.run context
    console.log("Contexts match: " + (firstIndex.context === context));
});
```

---

### filter

**Type:** `Word.IndexFilter | "None" | "Aiueo" | "Akasatana" | "Chosung" | "Low" | "Medium" | "Full"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a value that represents how Microsoft Word classifies the first character of entries in the index. See IndexFilter for available values.

#### Examples

**Example**: Get the filter classification of the first index in the document and display it in the console

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    const firstIndex = indexes.getFirst();
    
    // Load the filter property
    firstIndex.load("filter");
    
    await context.sync();
    
    // Display the filter classification
    console.log("Index filter classification: " + firstIndex.filter);
});
```

---

### headingSeparator

**Type:** `Word.HeadingSeparator | "None" | "BlankLine" | "Letter" | "LetterLow" | "LetterFull"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the text between alphabetical groups (entries that start with the same letter) in the index. Corresponds to the \h switch for an INDEX field.

#### Examples

**Example**: Get the current heading separator style used between alphabetical groups in the first index of the document and display it in the console.

```typescript
await Word.run(async (context) => {
    const indexes = context.document.body.indexes;
    indexes.load("items");
    await context.sync();
    
    if (indexes.items.length > 0) {
        const firstIndex = indexes.items[0];
        firstIndex.load("headingSeparator");
        await context.sync();
        
        console.log("Heading separator: " + firstIndex.headingSeparator);
    }
});
```

---

### indexLanguage

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a LanguageId value that represents the sorting language to use for the index.

#### Examples

**Example**: Get the sorting language of the first index in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.getIndexes();
    const firstIndex = indexes.getFirst();
    
    // Load the indexLanguage property
    firstIndex.load("indexLanguage");
    
    await context.sync();
    
    // Display the index language
    console.log("Index sorting language: " + firstIndex.indexLanguage);
});
```

---

### numberOfColumns

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the number of columns for each page of the index.

#### Examples

**Example**: Get the number of columns configured for the first index in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    const firstIndex = indexes.getFirst();
    
    // Load the numberOfColumns property
    firstIndex.load("numberOfColumns");
    
    await context.sync();
    
    // Display the number of columns
    console.log(`The index has ${firstIndex.numberOfColumns} column(s)`);
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a Range object that represents the portion of the document that is contained within the index.

#### Examples

**Example**: Get the text content from the first index in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    const firstIndex = indexes.getFirst();
    
    // Get the range of the index
    const indexRange = firstIndex.range;
    
    // Load the text property of the range
    indexRange.load("text");
    
    await context.sync();
    
    // Display the index content
    console.log("Index content: " + indexRange.text);
});
```

---

### rightAlignPageNumbers

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if page numbers are aligned with the right margin in the index.

#### Examples

**Example**: Set the first index in the document to right-align its page numbers

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    const firstIndex = indexes.getFirst();
    
    // Set page numbers to be right-aligned
    firstIndex.rightAlignPageNumbers = true;
    
    await context.sync();
});
```

---

### separateAccentedLetterHeadings

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets if the index contains separate headings for accented letters (for example, words that begin with "À" are under one heading and words that begin with "A" are under another).

#### Examples

**Example**: Check if the first index in the document uses separate headings for accented letters and display the result in the console.

```typescript
await Word.run(async (context) => {
    const indexes = context.document.body.indexes;
    indexes.load("items");
    
    await context.sync();
    
    if (indexes.items.length > 0) {
        const firstIndex = indexes.items[0];
        firstIndex.load("separateAccentedLetterHeadings");
        
        await context.sync();
        
        console.log(`Separate accented letter headings: ${firstIndex.separateAccentedLetterHeadings}`);
    } else {
        console.log("No indexes found in the document.");
    }
});
```

---

### sortBy

**Type:** `Word.IndexSortBy | "Stroke" | "Syllable"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the sorting criteria for the index.

#### Examples

**Example**: Set the index sorting criteria to sort by syllable instead of the default stroke order

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    indexes.load("items");
    await context.sync();
    
    if (indexes.items.length > 0) {
        const firstIndex = indexes.items[0];
        
        // Set the sorting criteria to syllable
        firstIndex.sortBy = Word.IndexSortBy.syllable;
        
        await context.sync();
        console.log("Index sorting set to syllable");
    }
});
```

---

### tabLeader

**Type:** `Word.TabLeader | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the leader character between entries in the index and their associated page numbers.

#### Examples

**Example**: Set the index tab leader to dots so that page numbers in the index are connected to their entries with dotted lines

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    indexes.load("items");
    await context.sync();
    
    if (indexes.items.length > 0) {
        const firstIndex = indexes.items[0];
        
        // Set the tab leader to dots
        firstIndex.tabLeader = Word.TabLeader.dots;
        
        await context.sync();
        console.log("Index tab leader set to dots");
    }
});
```

---

### type

**Type:** `Word.IndexType | "Indent" | "Runin"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the index type.

#### Examples

**Example**: Check if the first index in the document is formatted as an indented index or a run-in index and display the result in the console.

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    const firstIndex = indexes.getFirst();
    
    // Load the type property
    firstIndex.load("type");
    
    await context.sync();
    
    // Display the index type
    console.log(`Index type: ${firstIndex.type}`);
    // Output will be either "Indent" or "Runin"
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes this index.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete the first index from the document

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    const firstIndex = indexes.getFirst();
    
    // Delete the index
    firstIndex.delete();
    
    await context.sync();
    console.log("Index deleted successfully");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.IndexLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Index`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Index`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Index`

#### Examples

**Example**: Load and display the heading type property of the first index in the document

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    const firstIndex = indexes.getFirst();
    
    // Load the headingType property
    firstIndex.load("headingType");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the loaded property
    console.log("Index heading type: " + firstIndex.headingType);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.IndexUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Index` (required)

  **Returns:** `void`

#### Examples

**Example**: Update an existing index in the document by setting multiple properties at once, including its heading and letter navigation settings

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    indexes.load("items");
    await context.sync();
    
    if (indexes.items.length > 0) {
        const firstIndex = indexes.items[0];
        
        // Set multiple properties at once
        firstIndex.set({
            heading: "A",
            letterNavigation: true
        });
        
        await context.sync();
        console.log("Index properties updated successfully");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Index object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.IndexData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.IndexData`

#### Examples

**Example**: Serialize an index object to JSON format to log or store its properties outside of the Word context.

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    indexes.load("items");
    await context.sync();
    
    if (indexes.items.length > 0) {
        const firstIndex = indexes.items[0];
        
        // Load properties you want to serialize
        firstIndex.load("type");
        await context.sync();
        
        // Convert the Index object to a plain JavaScript object
        const indexData = firstIndex.toJSON();
        
        // Now you can use the plain object outside Word context
        console.log("Index data:", JSON.stringify(indexData, null, 2));
        
        // The plain object can be stored, transmitted, or processed
        // without maintaining the Word API context
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Index`

#### Examples

**Example**: Track an index object across multiple sync calls to maintain its reference while modifying its properties and reading updated values from the document.

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    indexes.load("items");
    await context.sync();
    
    if (indexes.items.length > 0) {
        const firstIndex = indexes.items[0];
        
        // Track the index object to use it across multiple sync calls
        firstIndex.track();
        
        // Load properties
        firstIndex.load("type");
        await context.sync();
        
        // Use the tracked object after sync - it remains valid
        console.log("Index type: " + firstIndex.type);
        
        // Can continue to use the tracked object in subsequent operations
        firstIndex.load("headingSeparator");
        await context.sync();
        
        console.log("Heading separator: " + firstIndex.headingSeparator);
        
        // Untrack when done to free up memory
        firstIndex.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Index`

#### Examples

**Example**: Get the first index in the document, use it to perform operations, then untrack it to release memory

```typescript
await Word.run(async (context) => {
    // Get the first index in the document
    const indexes = context.document.body.indexes;
    indexes.load("items");
    await context.sync();
    
    if (indexes.items.length > 0) {
        const firstIndex = indexes.items[0];
        
        // Track the object for use
        firstIndex.track();
        
        // Load properties to work with
        firstIndex.load("type");
        await context.sync();
        
        // Perform operations with the index
        console.log("Index type: " + firstIndex.type);
        
        // Untrack the object to release memory
        firstIndex.untrack();
        await context.sync();
    }
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.index
