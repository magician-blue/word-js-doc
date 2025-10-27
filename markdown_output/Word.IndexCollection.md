# Word.IndexCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

A collection of [Word.Index](/en-us/javascript/api/word/word.index) objects that represents all the indexes in the document.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from an IndexCollection to verify the connection to the Word document before performing operations on indexes.

```typescript
await Word.run(async (context) => {
    // Get the index collection from the document
    const indexCollection = context.document.body.indexes;
    
    // Access the request context associated with the collection
    const requestContext = indexCollection.context;
    
    // Verify the context is valid by loading a property
    indexCollection.load("items");
    await requestContext.sync();
    
    // Log confirmation that context is properly connected
    console.log(`Connected to Word document with ${indexCollection.items.length} indexes`);
    console.log("Request context is active and synchronized");
});
```

---

### items

**Type:** `Word.Index[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all indexes in the document and display their titles in the console.

```typescript
await Word.run(async (context) => {
    // Get the index collection from the document
    const indexes = context.document.body.indexes;
    
    // Load the items property to access the array of Index objects
    indexes.load("items");
    
    await context.sync();
    
    // Access the loaded items array and display each index title
    console.log(`Found ${indexes.items.length} index(es) in the document`);
    
    indexes.items.forEach((index, i) => {
        index.load("type");
        console.log(`Index ${i + 1}: ${index.type}`);
    });
    
    await context.sync();
});
```

---

## Methods

### add

**Kind:** `create`

Returns an Index object that represents a new index added to the document.

#### Signature

**Parameters:**
- `range`: `Word.Range` (required)
  The range where you want the index to appear. The index replaces the range, if the range is not collapsed.
- `indexAddOptions`: `Word.IndexAddOptions` (optional)
  Optional. The options for adding the index.

**Returns:** `Word.Index`

#### Examples

**Example**: Add a new index to the document at the current selection with specific formatting options

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const range = context.document.getSelection();
    
    // Define index options
    const indexOptions = {
        type: Word.IndexType.runIn,
        numberOfColumns: 2,
        letterHeadingFormat: Word.IndexLetterHeadingFormat.letter
    };
    
    // Add a new index at the selection
    const newIndex = context.document.indexes.add(range, indexOptions);
    
    // Load properties to verify
    newIndex.load("type");
    
    await context.sync();
    
    console.log("Index added successfully");
});
```

---

### getFormat

**Kind:** `read`

Gets the IndexFormat value that represents the formatting for the indexes in the document.

#### Signature

**Returns:** `OfficeExtension.ClientResult<Word.IndexFormat>`

#### Examples

**Example**: Get and display the formatting type of all indexes in the document (e.g., whether they use indented or run-in format)

```typescript
await Word.run(async (context) => {
    // Get the index collection from the document
    const indexes = context.document.body.indexes;
    
    // Get the format for the indexes
    const indexFormat = indexes.getFormat();
    indexFormat.load("type");
    
    await context.sync();
    
    // Display the index format type
    console.log("Index format type: " + indexFormat.type);
});
```

---

### getItem

**Kind:** `read`

Gets an Index object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of an Index object.

**Returns:** `Word.Index`

#### Examples

**Example**: Get the first index in the document and load its title property to display it

```typescript
await Word.run(async (context) => {
    // Get the collection of indexes in the document
    const indexes = context.document.body.indexes;
    
    // Get the first index (at position 0)
    const firstIndex = indexes.getItem(0);
    
    // Load the title property
    firstIndex.load("title");
    
    await context.sync();
    
    // Display the title of the first index
    console.log("First index title: " + firstIndex.title);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.IndexCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.IndexCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.IndexCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.IndexCollection`

#### Examples

**Example**: Load and display the count of all indexes in the document

```typescript
await Word.run(async (context) => {
    // Get the index collection from the document
    const indexes = context.document.body.indexes;
    
    // Load the count property of the index collection
    indexes.load("count");
    
    // Synchronize the document state
    await context.sync();
    
    // Display the number of indexes in the document
    console.log(`Total indexes in document: ${indexes.count}`);
});
```

---

### markAllEntries

**Kind:** `configure`

Inserts an [XE (Index Entry) field](https://support.microsoft.com/office/abaf7c78-6e21-418d-bf8b-f8186d2e4d08) after all instances of the text in the range.

#### Signature

**Parameters:**
- `range`: `Word.Range` (required)
  The range whose text is marked with an XE field throughout the document.
- `markAllEntriesOptions`: `Word.IndexMarkAllEntriesOptions` (optional)
  Optional. The options for marking all entries.

**Returns:** `void`

#### Examples

**Example**: Mark all instances of the word "JavaScript" in the document as index entries with the main entry text "Programming Languages"

```typescript
await Word.run(async (context) => {
    // Search for all instances of "JavaScript" in the document
    const searchResults = context.document.body.search("JavaScript", { matchCase: true });
    searchResults.load("items");
    
    await context.sync();
    
    if (searchResults.items.length > 0) {
        // Get the first search result to use as the range
        const range = searchResults.items[0];
        
        // Get the index collection
        const indexes = context.document.indexes;
        
        // Mark all instances of "JavaScript" as index entries
        indexes.markAllEntries(range, {
            entry: "Programming Languages",
            subEntry: "JavaScript"
        });
        
        await context.sync();
        console.log(`Marked ${searchResults.items.length} instances of "JavaScript" as index entries.`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.IndexCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.IndexCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.IndexCollectionData`

#### Examples

**Example**: Export all indexes in the document to JSON format for logging or external processing

```typescript
await Word.run(async (context) => {
    // Get the collection of all indexes in the document
    const indexes = context.document.body.indexes;
    
    // Load properties needed for the indexes
    indexes.load("items");
    
    await context.sync();
    
    // Convert the index collection to a plain JavaScript object
    const indexData = indexes.toJSON();
    
    // Log the JSON representation
    console.log("Index collection as JSON:", JSON.stringify(indexData, null, 2));
    
    // Access the items array from the JSON object
    console.log(`Number of indexes: ${indexData.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.IndexCollection`

#### Examples

**Example**: Track all indexes in a document to monitor and update their properties across multiple sync calls without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get the index collection from the document
    const indexes = context.document.body.indexes;
    
    // Load the indexes
    indexes.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    indexes.track();
    
    // First sync - get initial count
    console.log(`Found ${indexes.items.length} indexes`);
    await context.sync();
    
    // Second sync - can still access the collection safely
    // because it's being tracked
    for (let i = 0; i < indexes.items.length; i++) {
        indexes.items[i].load("type");
    }
    await context.sync();
    
    // Log index types
    indexes.items.forEach((index, i) => {
        console.log(`Index ${i + 1} type: ${index.type}`);
    });
    
    // Untrack when done to free up memory
    indexes.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.IndexCollection`

#### Examples

**Example**: Load all indexes in a document, process them, then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    // Load the indexes collection
    const indexCollection = context.document.body.indexes;
    indexCollection.load("items");
    
    await context.sync();
    
    // Process the indexes (e.g., log count)
    console.log(`Found ${indexCollection.items.length} indexes in the document`);
    
    // Untrack the collection to release memory
    indexCollection.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
