# Word.SectionCollection

**Package:** `word`

**API Set:** WordApi 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains the collection of the document's [Word.Section](https://learn.microsoft.com/en-us/javascript/api/word/word.section) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-section-breaks.yaml

// Inserts a section break on the next even page.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.insertBreak(Word.BreakType.sectionEven, Word.InsertLocation.end);

  await context.sync();

  console.log("Inserted section break on next even page.");
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a SectionCollection to verify the connection between the add-in and Word application before performing section operations.

```typescript
await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    
    // Access the request context associated with the SectionCollection
    const requestContext = sections.context;
    
    // Verify the context is valid and connected
    if (requestContext) {
        console.log("SectionCollection is connected to Word application");
        console.log(`Number of sections: ${sections.items.length}`);
    }
});
```

---

### items

**Type:** `Word.Section[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all sections in the document and log the count of sections to the console.

```typescript
await Word.run(async (context) => {
    // Get the section collection from the document
    const sections = context.document.sections;
    
    // Load the items property to access the array of sections
    sections.load("items");
    
    await context.sync();
    
    // Access the loaded sections array and log the count
    console.log(`Total sections in document: ${sections.items.length}`);
    
    // Optionally, iterate through each section
    sections.items.forEach((section, index) => {
        console.log(`Section ${index + 1} found`);
    });
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first section in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.Section`

#### Examples

**Example**: Get the first section of the document and change its page margins to 1 inch on all sides.

```typescript
await Word.run(async (context) => {
    // Get the first section in the document
    const firstSection = context.document.sections.getFirst();
    
    // Set page margins to 1 inch (72 points = 1 inch)
    firstSection.pageSetup.topMargin = 72;
    firstSection.pageSetup.bottomMargin = 72;
    firstSection.pageSetup.leftMargin = 72;
    firstSection.pageSetup.rightMargin = 72;
    
    await context.sync();
    
    console.log("First section margins updated successfully");
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first section in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Section`

#### Examples

**Example**: Check if the document has any sections and display the first section's header text, or show a message if no sections exist

```typescript
await Word.run(async (context) => {
    const sections = context.document.sections;
    const firstSection = sections.getFirstOrNullObject();
    firstSection.load("isNullObject");
    
    await context.sync();
    
    if (firstSection.isNullObject) {
        console.log("No sections found in the document.");
    } else {
        const header = firstSection.getHeader(Word.HeaderFooterType.primary);
        header.load("text");
        await context.sync();
        
        console.log("First section header text: " + header.text);
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.SectionCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.SectionCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.SectionCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.SectionCollection`

#### Examples

**Example**: Load and display the body text of all sections in the document

```typescript
await Word.run(async (context) => {
    // Get all sections in the document
    const sections = context.document.sections;
    
    // Load the body text property for all sections
    sections.load("items/body/text");
    
    // Synchronize to execute the load command
    await context.sync();
    
    // Display the text from each section
    sections.items.forEach((section, index) => {
        console.log(`Section ${index + 1}: ${section.body.text}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.SectionCollectionData`

#### Examples

**Example**: Export section information to JSON format for logging or external processing

```typescript
await Word.run(async (context) => {
    // Get all sections in the document
    const sections = context.document.sections;
    
    // Load properties we want to export
    sections.load("items/body/text");
    
    await context.sync();
    
    // Convert the section collection to a plain JavaScript object
    const sectionsJSON = sections.toJSON();
    
    // Now you can use the plain object (e.g., log it, send to server, etc.)
    console.log("Section data:", JSON.stringify(sectionsJSON, null, 2));
    console.log(`Total sections: ${sectionsJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.SectionCollection`

#### Examples

**Example**: Access and modify section properties across multiple sync calls by tracking the section collection to avoid "InvalidObjectPath" errors.

```typescript
await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    sections.track();
    
    // First sync - get section count
    console.log(`Document has ${sections.items.length} section(s)`);
    await context.sync();
    
    // Second sync - modify sections (tracking prevents InvalidObjectPath error)
    for (let i = 0; i < sections.items.length; i++) {
        const section = sections.items[i];
        section.body.insertParagraph(`Modified section ${i + 1}`, Word.InsertLocation.start);
    }
    await context.sync();
    
    // Clean up tracking when done
    sections.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.SectionCollection`

#### Examples

**Example**: Process all document sections to get their headers, then untrack the section collection to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the section collection and track it
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    
    // Process the sections (e.g., get header count)
    console.log(`Document has ${sections.items.length} sections`);
    
    // Untrack the section collection to release memory
    sections.untrack();
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.sectioncollection
