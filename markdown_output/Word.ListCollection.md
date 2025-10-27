# Word.ListCollection

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.List](/en-us/javascript/api/word/word.list) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Gets information about the first list in the document.
await Word.run(async (context) => {
  const lists: Word.ListCollection = context.document.body.lists;
  lists.load("items");

  await context.sync();

  if (lists.items.length === 0) {
    console.warn("There are no lists in this document.");
    return;
  }
  
  // Get the first list.
  const list: Word.List = lists.getFirst();
  list.load("levelTypes,levelExistences");

  await context.sync();

  const levelTypes  = list.levelTypes;
  console.log("Level types of the first list:");
  for (let i = 0; i < levelTypes.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelTypes[i]}`);
  }

  const levelExistences = list.levelExistences;
  console.log("Level existences of the first list:");
  for (let i = 0; i < levelExistences.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelExistences[i]}`);
  }
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ListCollection to verify the connection to the Word document and log context information for debugging purposes.

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    // Access the request context associated with the ListCollection
    const requestContext = lists.context;
    
    // Verify the context is valid and connected
    if (requestContext) {
        console.log("ListCollection is connected to Word context");
        console.log("Number of lists found:", lists.items.length);
        
        // The context can be used to perform operations
        await requestContext.sync();
    }
});
```

---

### items

**Type:** `Word.List[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all lists in the document and display the count and ID of each list.

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    
    await context.sync();
    
    console.log(`Total lists found: ${lists.items.length}`);
    
    lists.items.forEach((list, index) => {
        list.load("id");
    });
    
    await context.sync();
    
    lists.items.forEach((list, index) => {
        console.log(`List ${index + 1} ID: ${list.id}`);
    });
});
```

---

## Methods

### getById

**Kind:** `read`

Gets a list by its identifier. Throws an `ItemNotFound` error if there isn't a list with the identifier in this collection.

#### Signature

**Parameters:**
- `id`: `number` (required)
  A list identifier.

**Returns:** `Word.List`

#### Examples

**Example**: Get a list by its ID and change the font color of all items in that list to blue

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    const listId = 1; // Assuming we know the list ID
    
    const list = lists.getById(listId);
    list.load("paragraphs");
    
    await context.sync();
    
    // Change font color of all paragraphs in the list
    for (let i = 0; i < list.paragraphs.items.length; i++) {
        list.paragraphs.items[i].font.color = "blue";
    }
    
    await context.sync();
});
```

---

### getByIdOrNullObject

**Kind:** `read`

Gets a list by its identifier. If there isn't a list with the identifier in this collection, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Parameters:**
- `id`: `number` (required)
  A list identifier.

**Returns:** `Word.List`

#### Examples

**Example**: Check if a list with a specific ID exists in the document and display its level count, or show a message if the list is not found.

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    const listId = 1; // The ID of the list to find
    
    const list = lists.getByIdOrNullObject(listId);
    list.load("isNullObject, levelTypes");
    
    await context.sync();
    
    if (list.isNullObject) {
        console.log(`List with ID ${listId} does not exist in the document.`);
    } else {
        console.log(`List with ID ${listId} found. It has ${list.levelTypes.length} levels.`);
    }
});
```

---

### getFirst

**Kind:** `read`

Gets the first list in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.List`

#### Examples

**Example**: Get the first list in the document and highlight it in yellow

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const firstList = context.document.body.lists.getFirst();
    
    // Load the list's paragraphs to apply formatting
    firstList.load("paragraphs");
    await context.sync();
    
    // Highlight all paragraphs in the first list
    for (let i = 0; i < firstList.paragraphs.items.length; i++) {
        firstList.paragraphs.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first list in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.List`

#### Examples

**Example**: Check if the document contains any lists and display the ID of the first list, or show a message if no lists exist.

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    const firstList = lists.getFirstOrNullObject();
    firstList.load("id, isNullObject");
    
    await context.sync();
    
    if (firstList.isNullObject) {
        console.log("No lists found in the document.");
    } else {
        console.log(`First list ID: ${firstList.id}`);
    }
});
```

---

### getItem

**Kind:** `read`

Gets a list object by its ID.

#### Signature

**Parameters:**
- `id`: `number` (required)
  The list's ID.

**Returns:** `Word.List`

#### Examples

**Example**: Get a specific list by its ID and change the font color of all items in that list to blue

```typescript
await Word.run(async (context) => {
    // Get the list collection from the document body
    const lists = context.document.body.lists;
    
    // Assume we know the list ID (e.g., from a previous operation)
    const listId = 1;
    
    // Get the specific list by its ID
    const list = lists.getItem(listId);
    
    // Get all paragraphs in this list
    const listParagraphs = list.paragraphs;
    listParagraphs.load("font");
    
    await context.sync();
    
    // Change the font color of all items in the list
    for (let i = 0; i < listParagraphs.items.length; i++) {
        listParagraphs.items[i].font.color = "blue";
    }
    
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
  - `options`: `Word.Interfaces.ListCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ListCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ListCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ListCollection`

#### Examples

**Example**: Load and display the number of lists in the document along with the first list's ID

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    
    await context.sync();
    
    console.log(`Number of lists in document: ${lists.items.length}`);
    if (lists.items.length > 0) {
        console.log(`First list ID: ${lists.items[0].id}`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ListCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.ListCollectionData`

#### Examples

**Example**: Export all lists in the document to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get all lists in the document
    const lists = context.document.body.lists;
    
    // Load properties we want to export
    lists.load("id,levelTypes,levelExistences");
    
    await context.sync();
    
    // Convert the ListCollection to a plain JavaScript object
    const listsJSON = lists.toJSON();
    
    // Now you can stringify and use the data
    console.log(JSON.stringify(listsJSON, null, 2));
    
    // The listsJSON object contains an items array with the loaded properties
    console.log(`Found ${listsJSON.items.length} lists in the document`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ListCollection`

#### Examples

**Example**: Track all lists in the document and display their IDs across multiple sync calls to avoid InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    
    // Track the collection to use it across multiple sync calls
    lists.track();
    
    await context.sync();
    
    console.log(`Found ${lists.items.length} lists`);
    
    // Can safely use the lists collection after another sync
    await context.sync();
    
    // Display list IDs
    for (let i = 0; i < lists.items.length; i++) {
        lists.items[i].load("id");
    }
    
    await context.sync();
    
    lists.items.forEach((list, index) => {
        console.log(`List ${index + 1} ID: ${list.id}`);
    });
    
    // Untrack when done
    lists.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ListCollection`

#### Examples

**Example**: Load all lists in a document, process them, then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    // Get the list collection and track it
    const lists = context.document.body.lists;
    lists.load("items");
    
    await context.sync();
    
    // Process the lists (e.g., log count)
    console.log(`Found ${lists.items.length} lists in the document`);
    
    // Untrack the collection to release memory
    lists.untrack();
    
    await context.sync();
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml
