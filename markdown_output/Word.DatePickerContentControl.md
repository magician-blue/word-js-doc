# DatePickerContentControl

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the DatePickerContentControl object.

## Properties

### appearance

**Type:** `Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the appearance of the content control.

#### Examples

**Example**: Set a date picker content control's appearance to show only bounding box borders without displaying tags

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject();
    
    datePicker.load("appearance");
    await context.sync();
    
    if (!datePicker.isNullObject) {
        // Set the appearance to BoundingBox
        datePicker.appearance = Word.ContentControlAppearance.boundingBox;
        
        await context.sync();
        console.log("Date picker appearance set to BoundingBox");
    }
});
```

---

### color

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the color of a date picker content control to blue (#0000FF)

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerContentControl = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]).getFirst();
    
    // Set the color to blue
    datePickerContentControl.color = "#0000FF";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a date picker content control to verify the connection to the Word host application and log its properties.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject() as Word.DatePickerContentControl;
    
    // Load the date picker
    datePicker.load("title");
    await context.sync();
    
    // Access the request context from the date picker content control
    const requestContext = datePicker.context;
    
    // Verify the context is connected and log information
    console.log("Request context is available:", requestContext !== null);
    console.log("Context type:", typeof requestContext);
    console.log("Date picker title:", datePicker.title);
    
    await context.sync();
});
```

---

### dateCalendarType

**Type:** `Word.CalendarType | "Western" | "Arabic" | "Hebrew" | "Taiwan" | "Japan" | "Thai" | "Korean" | "SakaEra" | "TranslitEnglish" | "TranslitFrench" | "Umalqura"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a CalendarType value that represents the calendar type for the date picker content control.

#### Examples

**Example**: Set a date picker content control to use the Hebrew calendar type

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject() as Word.DatePickerContentControl;
    
    // Load the content control
    datePicker.load("dateCalendarType");
    await context.sync();
    
    // Set the calendar type to Hebrew
    if (!datePicker.isNullObject) {
        datePicker.dateCalendarType = Word.CalendarType.hebrew;
        await context.sync();
    }
});
```

---

### dateDisplayFormat

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the format in which dates are displayed.

#### Examples

**Example**: Set the date display format to "MMMM d, yyyy" (e.g., "January 15, 2024") for a date picker content control

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControl = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]).getFirst();
    
    // Set the date display format
    datePickerControl.dateDisplayFormat = "MMMM d, yyyy";
    
    await context.sync();
});
```

---

### dateDisplayLocale

**Type:** `Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a LanguageId that represents the language format for the date displayed in the date picker content control.

#### Examples

**Example**: Set the date display format in a date picker content control to use French (France) locale

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject() as Word.DatePickerContentControl;
    
    // Load the content control to check if it exists
    datePicker.load("id");
    await context.sync();
    
    if (!datePicker.isNullObject) {
        // Set the date display locale to French
        datePicker.dateDisplayLocale = "French";
        
        await context.sync();
        console.log("Date picker locale set to French");
    } else {
        console.log("No date picker content control found");
    }
});
```

---

### dateStorageFormat

**Type:** `Word.ContentControlDateStorageFormat | "Text" | "Date" | "DateTime"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a ContentControlDateStorageFormat value that represents the format for storage and retrieval of dates when the date picker content control is bound to the XML data store of the active document.

#### Examples

**Example**: Set a date picker content control to store dates in DateTime format when bound to XML data

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject();
    
    datePicker.load("dateStorageFormat");
    await context.sync();
    
    if (!datePicker.isNullObject) {
        // Set the storage format to DateTime for XML data binding
        datePicker.dateStorageFormat = Word.ContentControlDateStorageFormat.dateTime;
        
        await context.sync();
        console.log("Date storage format set to DateTime");
    }
});
```

---

### id

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the identification for the content control.

#### Examples

**Example**: Get and display the ID of a date picker content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerContentControl = context.document.contentControls
        .getByTypes([Word.ContentControlType.datePicker])
        .getFirst();
    
    // Load the id property
    datePickerContentControl.load("id");
    
    await context.sync();
    
    // Display the content control ID
    console.log("Date Picker Content Control ID: " + datePickerContentControl.id);
});
```

---

### isTemporary

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

#### Examples

**Example**: Set a date picker content control to be automatically removed from the document when the user edits its content

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject() as Word.DatePickerContentControl;
    
    // Load the control to check if it exists
    datePicker.load("isTemporary");
    await context.sync();
    
    if (!datePicker.isNullObject) {
        // Set the control to be temporary (removed when user edits it)
        datePicker.isTemporary = true;
        await context.sync();
        
        console.log("Date picker will be removed when user edits its content");
    }
});
```

---

### level

**Type:** `Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the level of the content controlâwhether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

#### Examples

**Example**: Check if a date picker content control is inline or at paragraph level and display the level information to the user.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject();
    
    datePicker.load("level");
    await context.sync();
    
    if (!datePicker.isNullObject) {
        // Get the level of the date picker content control
        const level = datePicker.level;
        console.log(`Date picker content control level: ${level}`);
        // Output will be one of: "Inline", "Paragraph", "Row", or "Cell"
    } else {
        console.log("No date picker content control found in the document.");
    }
});
```

---

### lockContentControl

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

#### Examples

**Example**: Lock a date picker content control to prevent users from deleting it from the document

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerContentControl = context.document.contentControls.getByTag("myDatePicker").getFirst();
    
    // Lock the content control so it cannot be deleted
    datePickerContentControl.lockContentControl = true;
    
    await context.sync();
    
    console.log("Date picker content control is now locked and cannot be deleted.");
});
```

---

### lockContents

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

#### Examples

**Example**: Lock the contents of a date picker content control to prevent users from editing the selected date

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControl = context.document.contentControls.getByTag("dateOfBirth").getFirst();
    
    // Lock the contents so users cannot edit the date
    datePickerControl.lockContents = true;
    
    await context.sync();
    
    console.log("Date picker contents are now locked");
});
```

---

### placeholderText

**Type:** `Word.BuildingBlock`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BuildingBlock object that represents the placeholder text for the content control.

#### Examples

**Example**: Get and display the placeholder text content from a date picker content control by accessing its BuildingBlock object.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject() as Word.DatePickerContentControl;
    
    // Load the placeholder text BuildingBlock
    datePicker.load("placeholderText");
    const placeholderBlock = datePicker.placeholderText;
    placeholderBlock.load("value");
    
    await context.sync();
    
    if (!datePicker.isNullObject) {
        console.log("Placeholder text: " + placeholderBlock.value);
    } else {
        console.log("No date picker content control found.");
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a Range object that represents the contents of the content control in the active document.

#### Examples

**Example**: Get the text content from a date picker content control and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirst();
    
    // Get the range of the date picker content control
    const range = datePicker.range;
    range.load("text");
    
    await context.sync();
    
    // Display the text content
    console.log("Date picker content: " + range.text);
});
```

---

### showingPlaceholderText

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether the placeholder text for the content control is being displayed.

#### Examples

**Example**: Check if a date picker content control is showing placeholder text and display an alert with the result.

```typescript
await Word.run(async (context) => {
    const datePickerControl = context.document.contentControls.getByTag("myDatePicker").getFirst();
    datePickerControl.load("showingPlaceholderText");
    
    await context.sync();
    
    if (datePickerControl.showingPlaceholderText) {
        console.log("The date picker is showing placeholder text - no date has been selected yet.");
    } else {
        console.log("The date picker has a value - a date has been selected.");
    }
});
```

---

### tag

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a tag to identify the content control.

#### Examples

**Example**: Set a tag "birthdate-field" on a date picker content control to identify it for later retrieval or processing

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject();
    
    datePicker.load("tag");
    await context.sync();
    
    if (!datePicker.isNullObject) {
        // Set a tag to identify this date picker
        datePicker.tag = "birthdate-field";
        await context.sync();
        
        console.log("Tag set successfully");
    }
});
```

---

### title

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the title for the content control.

#### Examples

**Example**: Set the title of a date picker content control to "Project Deadline"

```typescript
await Word.run(async (context) => {
    const datePickerContentControl = context.document.contentControls.getByTag("datePickerTag").getFirst();
    datePickerContentControl.title = "Project Deadline";
    
    await context.sync();
});
```

---

### xmlMapping

**Type:** `Word.XmlMapping`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

#### Examples

**Example**: Check if a date picker content control has an XML mapping and display the mapping's XPath if it exists.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject() as Word.DatePickerContentControl;
    
    // Load the XML mapping property
    datePicker.load("xmlMapping");
    
    await context.sync();
    
    if (!datePicker.isNullObject) {
        const xmlMapping = datePicker.xmlMapping;
        xmlMapping.load("isMapped, xpath");
        
        await context.sync();
        
        if (xmlMapping.isMapped) {
            console.log("Date picker is mapped to XML path: " + xmlMapping.xpath);
        } else {
            console.log("Date picker is not mapped to XML data");
        }
    }
});
```

---

## Methods

### copy

Copies the content control from the active document to the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Copy a date picker content control to the clipboard so it can be pasted elsewhere in the document

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    datePickerControls.load("items");
    
    await context.sync();
    
    if (datePickerControls.items.length > 0) {
        const datePickerControl = datePickerControls.items[0] as Word.DatePickerContentControl;
        
        // Copy the date picker content control to the clipboard
        datePickerControl.copy();
        
        await context.sync();
        
        console.log("Date picker content control copied to clipboard");
    } else {
        console.log("No date picker content control found in the document");
    }
});
```

---

### cut

Removes the content control from the active document and moves the content control to the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove a date picker content control from the document and move it to the clipboard so it can be pasted elsewhere

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTag("dateControl");
    datePickerControls.load("items");
    
    await context.sync();
    
    if (datePickerControls.items.length > 0) {
        const datePicker = datePickerControls.items[0];
        
        // Cut the date picker content control to clipboard
        datePicker.cut();
        
        await context.sync();
        console.log("Date picker content control has been cut to clipboard");
    }
});
```

---

### delete

**Kind:** `delete`

Deletes this content control and the contents of the content control.

#### Signature

**Parameters:**
- `deleteContents`: `boolean` (optional)
  Optional. If true, deletes the contents as well.

**Returns:** `void`

#### Examples

**Example**: Delete a date picker content control and its contents from the document

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerContentControl = context.document.contentControls
        .getByTag("dateControl")
        .getFirst();
    
    // Delete the date picker content control and its contents
    datePickerContentControl.delete(true);
    
    await context.sync();
    console.log("Date picker content control and its contents deleted.");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.DatePickerContentControlLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.DatePickerContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.DatePickerContentControl`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.DatePickerContentControl`

#### Examples

**Example**: Load and display the date value and title properties of the first date picker content control in the document.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControl = context.document.contentControls
        .getByTypes([Word.ContentControlType.datePicker])
        .getFirst() as Word.DatePickerContentControl;
    
    // Load the properties we want to read
    datePickerControl.load("title, date");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Title: " + datePickerControl.title);
    console.log("Date: " + datePickerControl.date);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.DatePickerContentControlUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.DatePickerContentControl` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple properties of a date picker content control at once, including its title, placeholder text, and date display format.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControl = context.document.contentControls.getByTag("myDatePicker").getFirst();
    
    // Set multiple properties at once
    datePickerControl.set({
        title: "Project Deadline",
        placeholderText: "Select a deadline date",
        dateDisplayFormat: "MMMM dd, yyyy"
    });
    
    await context.sync();
    console.log("Date picker properties updated successfully");
});
```

---

### setPlaceholderText

**Kind:** `write`

Sets the placeholder text that displays in the content control until a user enters their own text.

#### Signature

**Parameters:**
- `options`: `Word.ContentControlPlaceholderOptions` (optional)
  Optional. The options for configuring the content control's placeholder text.

**Returns:** `void`

#### Examples

**Example**: Set placeholder text "Select a date..." for a date picker content control at the beginning of the document.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject();
    
    datePicker.load("id");
    await context.sync();
    
    if (!datePicker.isNullObject) {
        // Set the placeholder text
        datePicker.setPlaceholderText("Select a date...");
        await context.sync();
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DatePickerContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DatePickerContentControlData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.DatePickerContentControlData`

#### Examples

**Example**: Serialize a date picker content control to JSON format for logging or data transfer purposes.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject() as Word.DatePickerContentControl;
    
    // Load properties to serialize
    datePicker.load("title,tag,dateDisplayFormat,placeholderText");
    
    await context.sync();
    
    // Convert to JSON
    const datePickerJSON = datePicker.toJSON();
    
    // Log the JSON representation
    console.log("Date Picker Control as JSON:", JSON.stringify(datePickerJSON, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.DatePickerContentControl`

#### Examples

**Example**: Track a date picker content control across multiple sync calls to prevent InvalidObjectPath errors when accessing its properties after document changes

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControls = context.document.contentControls.getByTypes([Word.ContentControlType.datePicker]);
    const datePicker = datePickerControls.getFirstOrNullObject() as Word.DatePickerContentControl;
    
    // Track the object to use it across multiple sync calls
    datePicker.track();
    
    await context.sync();
    
    // Now we can safely access and modify the date picker across syncs
    if (!datePicker.isNullObject) {
        datePicker.title = "Updated Date";
        await context.sync();
        
        // Access properties again after sync without errors
        datePicker.placeholderText = "Select a date";
        await context.sync();
        
        // Untrack when done to free up memory
        datePicker.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.DatePickerContentControl`

#### Examples

**Example**: Track a date picker content control, modify its properties, then untrack it to release memory after the operations are complete.

```typescript
await Word.run(async (context) => {
    // Get the first date picker content control in the document
    const datePickerControl = context.document.contentControls.getByTag("myDatePicker").getFirst();
    
    // Track the object to monitor changes
    datePickerControl.track();
    
    // Load and modify properties
    datePickerControl.load("title");
    await context.sync();
    
    datePickerControl.title = "Updated Date";
    await context.sync();
    
    // Untrack the object to release memory after we're done using it
    datePickerControl.untrack();
    await context.sync();
    
    console.log("Date picker control updated and untracked");
});
```

---
