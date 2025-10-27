# Word.NoteItemCollection

**Package:** `word`

**API Set:** WordApi 1.5

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.NoteItem](/en-us/javascript/api/word/word.noteitem) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the first footnote in the document body and select its reference mark.
await Word.run(async (context) => {
  const reference: Word.Range = context.document.body.footnotes.getFirst().reference;
  reference.select();
  console.log("The first footnote is selected.");
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a NoteItemCollection to synchronize footnote data with the Office host application.

```typescript
await Word.run(async (context) => {
    // Get the footnotes collection from the document body
    const footnotes = context.document.body.footnotes;
    
    // Access the request context from the collection
    const requestContext = footnotes.context;
    
    // Use the context to load properties and sync with the host
    footnotes.load("items");
    await requestContext.sync();
    
    // Now we can work with the loaded footnote items
    console.log(`Number of footnotes: ${footnotes.items.length}`);
    
    for (let i = 0; i < footnotes.items.length; i++) {
        console.log(`Footnote ${i + 1} reference: ${footnotes.items[i].reference}`);
    }
});
```

---

### items

**Type:** `Word.NoteItem[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Iterate through all footnote items in the document and log their reference text to the console.

```typescript
await Word.run(async (context) => {
    // Get the footnotes collection from the document body
    const footnotes = context.document.body.footnotes;
    
    // Load the items property to access the array of footnote items
    footnotes.load("items");
    
    await context.sync();
    
    // Access the items array and iterate through each footnote
    const footnoteItems = footnotes.items;
    
    for (let i = 0; i < footnoteItems.length; i++) {
        footnoteItems[i].load("reference");
        await context.sync();
        
        console.log(`Footnote ${i + 1}: ${footnoteItems[i].reference}`);
    }
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first note item in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.NoteItem`

#### Examples

**Example**: Select the reference mark of the first footnote in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the first footnote in the document body and select its reference mark.
await Word.run(async (context) => {
  const reference: Word.Range = context.document.body.footnotes.getFirst().reference;
  reference.select();
  console.log("The first footnote is selected.");
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first note item in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.NoteItem`

#### Examples

**Example**: Check if a document has any footnotes and display the text of the first footnote if it exists

```typescript
await Word.run(async (context) => {
    const footnotes = context.document.body.footnotes;
    const firstFootnote = footnotes.getFirstOrNullObject();
    firstFootnote.load("isNullObject, body/text");
    
    await context.sync();
    
    if (firstFootnote.isNullObject) {
        console.log("No footnotes found in the document.");
    } else {
        console.log("First footnote text: " + firstFootnote.body.text);
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
  - `options`: `Word.Interfaces.NoteItemCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.NoteItemCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.NoteItemCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.NoteItemCollection`

#### Examples

**Example**: Load and display the text content of all footnote items in the active document

```typescript
await Word.run(async (context) => {
    // Get the footnotes collection from the document body
    const footnotes = context.document.body.footnotes;
    
    // Load the items collection with their body text property
    footnotes.load("items");
    await context.sync();
    
    // Access the note items and load their text
    const noteItems = footnotes.items;
    noteItems.load("body/text");
    await context.sync();
    
    // Display the footnote text
    noteItems.items.forEach((noteItem, index) => {
        console.log(`Footnote ${index + 1}: ${noteItem.body.text}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.NoteItemCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.NoteItemCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.NoteItemCollectionData`

#### Examples

**Example**: Export footnote items to JSON format for logging or external processing

```typescript
await Word.run(async (context) => {
    // Get the footnotes collection from the document body
    const footnotes = context.document.body.footnotes;
    
    // Load the footnote items
    footnotes.load("items");
    await context.sync();
    
    // Get the note items collection
    const noteItems = footnotes.items[0]?.body.noteItems;
    
    if (noteItems) {
        // Load properties of the note items
        noteItems.load("type");
        await context.sync();
        
        // Convert the collection to a plain JavaScript object
        const noteItemsData = noteItems.toJSON();
        
        // Log or process the JSON data
        console.log("Note Items Data:", JSON.stringify(noteItemsData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.NoteItemCollection`

#### Examples

**Example**: Track a collection of footnote items to maintain references across multiple sync calls when processing document notes

```typescript
await Word.run(async (context) => {
    // Get the body of the document
    const body = context.document.body;
    
    // Get all footnotes in the document
    const footnotes = body.footnotes;
    context.load(footnotes);
    await context.sync();
    
    // Track the footnote collection to use it across multiple sync calls
    footnotes.track();
    
    // First sync - load footnote properties
    footnotes.load("items");
    await context.sync();
    
    // Process footnotes (e.g., modify their content)
    for (let i = 0; i < footnotes.items.length; i++) {
        const footnote = footnotes.items[i];
        footnote.body.insertText(`[Note ${i + 1}] `, "Start");
    }
    
    // Second sync - apply changes
    await context.sync();
    
    // Untrack when done to free up memory
    footnotes.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.NoteItemCollection`

#### Examples

**Example**: Load footnote items from a document, process them, then untrack the collection to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the footnotes collection from the document body
    const footnotes = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();
    
    // Get the note items collection from the first footnote
    const noteItems = footnotes.items[0].body.noteItems;
    noteItems.load("type");
    
    // Track the collection for processing
    noteItems.track();
    await context.sync();
    
    // Process the note items (e.g., log their types)
    console.log(`Found ${noteItems.items.length} note items`);
    
    // Untrack the collection to release memory after we're done
    noteItems.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
- https://docs.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml
