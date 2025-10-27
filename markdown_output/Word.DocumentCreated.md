# DocumentCreated

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

The DocumentCreated object is the top level object created by Application.CreateDocument. A DocumentCreated object is a special Document object.

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

### body

**Type:** `Word.Body`

**Since:** WordApiHiddenDocument 1.3

Gets the body object of the document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

#### Examples

**Example**: Add a paragraph of text to the document body after creating a new document

```typescript
await Word.run(async (context) => {
    const documentCreated = context.application.createDocument();
    const body = documentCreated.body;
    
    body.insertParagraph(
        "This is the first paragraph in the newly created document.",
        Word.InsertLocation.start
    );
    
    await context.sync();
    documentCreated.open();
});
```

---

### contentControls

**Type:** `Word.ContentControlCollection`

**Since:** WordApiHiddenDocument 1.3

Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.

#### Examples

**Example**: Find and highlight all content controls in a newly created document by setting their appearance to tags visible

```typescript
await Word.run(async (context) => {
    // Create a new document
    const myDocument = context.application.createDocument();
    
    // Access the content controls collection
    const contentControls = myDocument.contentControls;
    contentControls.load("items");
    
    await context.sync();
    
    // Make all content controls visible by showing their tags
    for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].appearance = Word.ContentControlAppearance.tags;
    }
    
    await context.sync();
    
    console.log(`Found ${contentControls.items.length} content controls in the document`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a newly created document to synchronize operations with the Office host application

```typescript
await Word.run(async (context) => {
    // Create a new document
    const newDoc = context.application.createDocument();
    
    // Access the request context of the newly created document
    const newDocContext = newDoc.context;
    
    // Use the context to insert text into the new document
    const body = newDoc.body;
    body.insertText("This text is inserted using the document's context.", Word.InsertLocation.start);
    
    // Synchronize the context with the Office host
    await newDocContext.sync();
    
    console.log("New document created and synchronized using its context.");
});
```

---

### customXmlParts

**Type:** `Word.CustomXmlPartCollection`

**Since:** WordApiHiddenDocument 1.4

Gets the custom XML parts in the document.

#### Examples

**Example**: Add a custom XML part to a newly created document and then retrieve all custom XML parts to verify it was added.

```typescript
await Word.run(async (context) => {
    // Create a new document
    const documentCreated = context.application.createDocument();
    
    // Add a custom XML part with sample data
    const xmlString = '<?xml version="1.0"?><employees><employee><name>John Doe</name><id>12345</id></employee></employees>';
    documentCreated.customXmlParts.add(xmlString);
    
    // Get all custom XML parts from the document
    const customXmlParts = documentCreated.customXmlParts;
    customXmlParts.load("items");
    
    await context.sync();
    
    // Log the count of custom XML parts
    console.log(`Total custom XML parts: ${customXmlParts.items.length}`);
    
    // Open the created document
    documentCreated.open();
    
    await context.sync();
});
```

---

### properties

**Type:** `Word.DocumentProperties`

**Since:** WordApiHiddenDocument 1.3

Gets the properties of the document.

#### Examples

**Example**: Read and display the document's title and author from the document properties.

```typescript
await Word.run(async (context) => {
    const documentCreated = context.application.createDocument();
    const properties = documentCreated.properties;
    
    properties.load(["title", "author"]);
    await context.sync();
    
    console.log("Document Title: " + properties.title);
    console.log("Document Author: " + properties.author);
});
```

---

### saved

**Type:** `boolean`

**Since:** WordApiHiddenDocument 1.3

Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

#### Examples

**Example**: Check if a newly created document has unsaved changes and display an appropriate message to the user

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Load the saved property
    document.load("saved");
    await context.sync();
    
    // Check if document has been saved
    if (document.saved) {
        console.log("Document has no unsaved changes");
    } else {
        console.log("Document has unsaved changes - please save your work");
    }
});
```

---

### sections

**Type:** `Word.SectionCollection`

**Since:** WordApiHiddenDocument 1.3

Gets the collection of section objects in the document.

#### Examples

**Example**: Get all sections in a newly created document and display the count of sections in the console.

```typescript
await Word.run(async (context) => {
    const docCreated = context.application.createDocument();
    const sections = docCreated.sections;
    sections.load("items");
    
    await context.sync();
    
    console.log(`The document has ${sections.items.length} section(s)`);
});
```

---

### settings

**Type:** `Word.SettingCollection`

**Since:** WordApiHiddenDocument 1.4

Gets the add-in's settings in the document.

#### Examples

**Example**: Store and retrieve a custom user preference setting in the document, such as saving a "lastViewedSection" value that persists with the document.

```typescript
await Word.run(async (context) => {
    // Get the settings collection from the document
    const settings = context.document.settings;
    
    // Add or update a setting
    settings.add("lastViewedSection", "Chapter3");
    settings.add("userPreference", "darkMode");
    
    // Load the settings to verify
    settings.load("items");
    await context.sync();
    
    // Retrieve a specific setting
    const lastSection = settings.getItemOrNullObject("lastViewedSection");
    lastSection.load("value");
    await context.sync();
    
    if (!lastSection.isNullObject) {
        console.log("Last viewed section:", lastSection.value);
    }
});
```

---

## Methods

### addStyle

**Kind:** `create`

Adds a style into the document by name and type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `name`: `string` (required)
    A string representing the style name.
  - `type`: `Word.StyleType` (required)
    The style type, including character, list, paragraph, or table.

  **Returns:** `Word.Style`

**Overload 2:**

  **Parameters:**
  - `name`: `string` (required)
    A string representing the style name.
  - `type`: `"Character" | "List" | "Paragraph" | "Table"` (required)
    The style type, including character, list, paragraph, or table.

  **Returns:** `Word.Style`

#### Examples

**Example**: Add a new character style named "HighlightText" and a paragraph style named "CustomHeading" to the document

```typescript
await Word.run(async (context) => {
    const documentCreated = context.document;
    
    // Add a character style for highlighting text
    documentCreated.addStyle("HighlightText", "Character");
    
    // Add a paragraph style for custom headings
    documentCreated.addStyle("CustomHeading", "Paragraph");
    
    await context.sync();
    
    console.log("Styles added successfully");
});
```

---

### deleteBookmark

**Kind:** `delete`

Deletes a bookmark, if it exists, from the document.

#### Signature

**Parameters:**
- `name`: `string` (required)
  The case-insensitive bookmark name.

**Returns:** `void`

#### Examples

**Example**: Delete a bookmark named "Section1Start" from the document

```typescript
await Word.run(async (context) => {
    const doc = context.document;
    
    // Delete the bookmark named "Section1Start"
    doc.deleteBookmark("Section1Start");
    
    await context.sync();
    console.log("Bookmark 'Section1Start' has been deleted.");
});
```

---

### getBookmarkRange

**Kind:** `read`

Gets a bookmark's range. Throws an ItemNotFound error if the bookmark doesn't exist.

#### Signature

**Parameters:**
- `name`: `string` (required)
  The case-insensitive bookmark name.

**Returns:** `Word.Range`

#### Examples

**Example**: Retrieve the text content from a bookmark named "CompanyName" in the document

```typescript
await Word.run(async (context) => {
    const doc = context.document;
    
    // Get the range of the bookmark named "CompanyName"
    const bookmarkRange = doc.getBookmarkRange("CompanyName");
    bookmarkRange.load("text");
    
    await context.sync();
    
    console.log("Bookmark text: " + bookmarkRange.text);
});
```

---

### getBookmarkRangeOrNullObject

**Kind:** `read`

Gets a bookmark's range. If the bookmark doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Parameters:**
- `name`: `string` (required)
  The case-insensitive bookmark name. Only alphanumeric and underscore characters are supported. It must begin with a letter but if you want to tag the bookmark as hidden, then start the name with an underscore character. Names can't be longer than 40 characters.

**Returns:** `Word.Range`

#### Examples

**Example**: Check if a bookmark named "Introduction" exists in the document and highlight its text, or show a message if the bookmark is not found.

```typescript
await Word.run(async (context) => {
    const doc = context.document;
    const bookmarkRange = doc.getBookmarkRangeOrNullObject("Introduction");
    
    bookmarkRange.load("isNullObject, text");
    await context.sync();
    
    if (bookmarkRange.isNullObject) {
        console.log("Bookmark 'Introduction' does not exist.");
    } else {
        bookmarkRange.font.highlightColor = "yellow";
        console.log("Bookmark found and highlighted: " + bookmarkRange.text);
    }
    
    await context.sync();
});
```

---

### getContentControls

**Kind:** `read`

Gets the currently supported content controls in the document.

#### Signature

**Parameters:**
- `options`: `Word.ContentControlOptions` (optional)
  Options that define which content controls are returned.

**Returns:** `Word.ContentControlCollection`

#### Examples

**Example**: Get all content controls in a newly created document and log their count to the console

```typescript
await Word.run(async (context) => {
    // Create a new document
    const documentCreated = context.application.createDocument();
    
    // Get all content controls in the document
    const contentControls = documentCreated.getContentControls();
    
    // Load the count property
    contentControls.load("items");
    
    await context.sync();
    
    // Log the number of content controls found
    console.log(`Found ${contentControls.items.length} content controls in the document`);
});
```

---

### getStyles

**Kind:** `read`

Gets a StyleCollection object that represents the whole style set of the document.

#### Signature

**Returns:** `Word.StyleCollection`

#### Examples

**Example**: Get all styles from a newly created document and log the name of each style to the console.

```typescript
await Word.run(async (context) => {
    // Create a new document
    const myDocument = context.application.createDocument();
    
    // Get the styles collection from the document
    const styles = myDocument.getStyles();
    
    // Load the name property for each style
    styles.load("items/name");
    
    await context.sync();
    
    // Log each style name
    console.log("Document styles:");
    for (let i = 0; i < styles.items.length; i++) {
        console.log(styles.items[i].name);
    }
});
```

---

### insertFileFromBase64

**Kind:** `create`

Inserts a document into the target document at a specific location with additional properties. Headers, footers, watermarks, and other section properties are copied by default.

#### Signature

**Parameters:**
- `base64File`: `string` (required)
  The Base64-encoded content of a .docx file.
- `insertLocation`: `Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"` (required)
  The value must be 'Replace', 'Start', or 'End'.
- `insertFileOptions`: `Word.InsertFileOptions` (optional)
  The additional properties that should be imported to the destination document.

**Returns:** `Word.SectionCollection`

#### Examples

**Example**: Insert a template document from a base64-encoded file into the current document at the end, including all headers and footers

```typescript
await Word.run(async (context) => {
    const docCreated = context.application.createDocument() as Word.DocumentCreated;
    
    // Base64-encoded .docx file content
    const base64Template = "UEsDBBQABgAIAAAAIQDd..."; // truncated for brevity
    
    // Insert the template at the end of the document
    docCreated.insertFileFromBase64(
        base64Template,
        Word.InsertLocation.end,
        {
            importHeaders: true,
            importFooters: true,
            importPageSetup: true,
            importStyles: true
        }
    );
    
    await context.sync();
    
    // Open the newly created document
    docCreated.open();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.DocumentCreatedLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.DocumentCreated`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.DocumentCreated`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.DocumentCreated`

#### Examples

**Example**: Create a new document and load its properties to verify the document was created successfully before performing operations on it.

```typescript
await Word.run(async (context) => {
    // Create a new document
    const newDoc = context.application.createDocument();
    
    // Load properties of the newly created document
    newDoc.load("body,properties");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can safely access the loaded properties
    console.log("Document created successfully");
    
    // Perform operations on the new document
    newDoc.body.insertParagraph("This is a new document", Word.InsertLocation.start);
    
    await context.sync();
});
```

---

### open

Opens the document.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Create a new Word document and open it in a separate tab or window.

```typescript
// Create and open a new document in a new tab or window.
await Word.run(async (context) => {
  const externalDoc = context.application.createDocument();
  await context.sync();

  externalDoc.open();
  await context.sync();
});
```

---

### save

Saves the document.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `saveBehavior`: `Word.SaveBehavior` (optional)
    DocumentCreated only supports 'Save'.
  - `fileName`: `string` (optional)
    The file name (exclude file extension). Only takes effect for a new document.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `saveBehavior`: `"Save" | "Prompt"` (optional)
    DocumentCreated only supports 'Save'.
  - `fileName`: `string` (optional)
    The file name (exclude file extension). Only takes effect for a new document.

  **Returns:** `void`

#### Examples

**Example**: Create a new document, add content to it, and save it to a specific file location

```typescript
// Create a new document
const doc = await Word.createDocument();

await Word.run(doc, async (context) => {
    // Add some content to the document
    const body = context.document.body;
    body.insertParagraph("This is a new document with sample content.", Word.InsertLocation.start);
    
    await context.sync();
    
    // Save the document with a specific filename
    context.document.save(Word.SaveBehavior.save, "MyNewDocument.docx");
    
    await context.sync();
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.DocumentCreatedUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.DocumentCreated` (required)

  **Returns:** `void`

#### Examples

**Example**: Set multiple properties on a newly created document, including title and subject metadata

```typescript
await Word.run(async (context) => {
    const documentCreated = context.application.createDocument();
    
    documentCreated.set({
        properties: {
            title: "Q4 Sales Report",
            subject: "Quarterly Financial Analysis",
            author: "Sales Team"
        }
    });
    
    await context.sync();
    
    // Open the newly created document
    documentCreated.open();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DocumentCreated object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DocumentCreatedData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.DocumentCreatedData`

#### Examples

**Example**: Create a new document and serialize its properties to JSON format for logging or data transfer purposes.

```typescript
await Word.run(async (context) => {
    // Create a new document
    const newDoc = context.application.createDocument();
    
    // Load properties of the created document
    newDoc.load("body,properties");
    
    await context.sync();
    
    // Convert the DocumentCreated object to a plain JavaScript object
    const docJSON = newDoc.toJSON();
    
    // Use the JSON representation (e.g., for logging or sending to a server)
    console.log("Document data:", JSON.stringify(docJSON, null, 2));
    
    // The JSON object contains shallow copies of loaded properties
    console.log("Document properties:", docJSON.properties);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.DocumentCreated`

#### Examples

**Example**: Create a new document and track it to safely access its properties across multiple sync calls without getting an "InvalidObjectPath" error

```typescript
await Word.run(async (context) => {
    // Create a new document
    const newDoc = context.application.createDocument();
    
    // Track the document object for use across sync calls
    newDoc.track();
    
    // First sync to load the document
    await context.sync();
    
    // Now we can safely work with the document across multiple syncs
    const body = newDoc.body;
    body.insertParagraph("This is content in the tracked document.", "Start");
    
    await context.sync();
    
    // Can continue to use the document object safely
    body.insertParagraph("Adding more content after another sync.", "End");
    
    await context.sync();
    
    // Untrack when done to free up memory
    newDoc.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.DocumentCreated`

#### Examples

**Example**: Create a new document, add content to it, and properly release the tracked DocumentCreated object from memory after use

```typescript
await Word.run(async (context) => {
    // Create a new document
    const newDoc = context.application.createDocument();
    context.load(newDoc);
    await context.sync();
    
    // Add some content to the new document
    const body = newDoc.body;
    body.insertParagraph("This is content in the new document.", Word.InsertLocation.start);
    await context.sync();
    
    // Release the memory associated with the tracked DocumentCreated object
    newDoc.untrack();
    await context.sync();
    
    console.log("New document created and untracked successfully.");
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-external-document.yaml
