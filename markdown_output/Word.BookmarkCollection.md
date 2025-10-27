# Word.BookmarkCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

A collection of Word.Bookmark objects that represent the bookmarks in the specified selection, range, or document.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a BookmarkCollection to verify the connection to the Word host application and log its properties.

```typescript
await Word.run(async (context) => {
    // Get the bookmark collection from the document
    const bookmarks = context.document.body.getRange().bookmarks;
    
    // Access the request context associated with the bookmark collection
    const requestContext = bookmarks.context;
    
    // Use the context to verify connection and perform operations
    console.log("Request context retrieved:", requestContext);
    console.log("Context is connected to host application");
    
    // Load bookmark properties using the same context
    bookmarks.load("items");
    await context.sync();
    
    console.log(`Found ${bookmarks.items.length} bookmarks using the request context`);
});
```

---

### items

**Type:** `Word.Bookmark[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all bookmarks in the document and log their names to the console

```typescript
await Word.run(async (context) => {
    // Get the bookmark collection from the document
    const bookmarks = context.document.body.bookmarks;
    
    // Load the items property to access the array of bookmarks
    bookmarks.load("items");
    
    await context.sync();
    
    // Access the loaded bookmark items and log their names
    console.log(`Total bookmarks: ${bookmarks.items.length}`);
    
    bookmarks.items.forEach((bookmark, index) => {
        bookmark.load("name");
    });
    
    await context.sync();
    
    bookmarks.items.forEach((bookmark, index) => {
        console.log(`Bookmark ${index + 1}: ${bookmark.name}`);
    });
});
```

---

## Methods

### exists

**Kind:** `read`

Determines whether the specified bookmark exists.

#### Signature

**Parameters:**
- `name`: `string` (required)
  A bookmark name than cannot include more than 40 characters or more than one word.

**Returns:** `OfficeExtension.ClientResult<boolean>`
true if the bookmark exists.

#### Examples

**Example**: Check if a bookmark named "Introduction" exists in the document and display an alert with the result

```typescript
await Word.run(async (context) => {
    const bookmarks = context.document.body.bookmarks;
    const bookmarkExists = bookmarks.exists("Introduction");
    
    await context.sync();
    
    if (bookmarkExists.value) {
        console.log("The bookmark 'Introduction' exists in the document.");
    } else {
        console.log("The bookmark 'Introduction' does not exist in the document.");
    }
});
```

---

### getItem

**Kind:** `read`

Gets a Bookmark object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a Bookmark object.

**Returns:** `Word.Bookmark`

#### Examples

**Example**: Get the first bookmark in the document and change its name to "UpdatedBookmark"

```typescript
await Word.run(async (context) => {
    // Get the bookmark collection from the document
    const bookmarks = context.document.body.bookmarks;
    
    // Get the first bookmark by index (0)
    const firstBookmark = bookmarks.getItem(0);
    
    // Load the name property to read it
    firstBookmark.load("name");
    
    await context.sync();
    
    console.log("Original bookmark name: " + firstBookmark.name);
    
    // Update the bookmark name
    firstBookmark.name = "UpdatedBookmark";
    
    await context.sync();
    
    console.log("Bookmark name updated successfully");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.BookmarkCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.BookmarkCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BookmarkCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BookmarkCollection`

#### Examples

**Example**: Load and display the names of all bookmarks in the document

```typescript
await Word.run(async (context) => {
    // Get the bookmark collection from the document
    const bookmarks = context.document.body.bookmarks;
    
    // Load the bookmark names
    bookmarks.load("items/name");
    
    await context.sync();
    
    // Display the bookmark names
    console.log("Bookmarks in document:");
    bookmarks.items.forEach(bookmark => {
        console.log(`- ${bookmark.name}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BookmarkCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BookmarkCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.BookmarkCollectionData`

#### Examples

**Example**: Export all bookmarks in the document to a JSON string for logging or external storage purposes.

```typescript
await Word.run(async (context) => {
    // Get all bookmarks in the document
    const bookmarks = context.document.body.bookmarks;
    
    // Load the bookmark properties
    bookmarks.load("name, isHidden");
    
    await context.sync();
    
    // Convert the bookmark collection to a plain JavaScript object
    const bookmarksData = bookmarks.toJSON();
    
    // Convert to JSON string for logging or storage
    const jsonString = JSON.stringify(bookmarksData, null, 2);
    
    console.log("Bookmarks as JSON:", jsonString);
    
    // Example: You could also access the items array directly
    console.log(`Total bookmarks: ${bookmarksData.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BookmarkCollection`

#### Examples

**Example**: Track a bookmark collection across multiple sync calls to monitor and access bookmarks even after document changes occur

```typescript
await Word.run(async (context) => {
    // Get the bookmark collection from the document
    const bookmarks = context.document.body.bookmarks;
    
    // Track the collection to maintain reference across sync calls
    bookmarks.track();
    
    // Load bookmark properties
    bookmarks.load("items");
    await context.sync();
    
    // First sync - log current bookmark count
    console.log(`Initial bookmark count: ${bookmarks.items.length}`);
    
    // Make changes to the document (e.g., insert text)
    context.document.body.insertText("New content added to document", Word.InsertLocation.end);
    await context.sync();
    
    // Second sync - the tracked collection still works
    bookmarks.load("items");
    await context.sync();
    console.log(`Bookmark count after changes: ${bookmarks.items.length}`);
    
    // Untrack when done to free up memory
    bookmarks.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BookmarkCollection`

#### Examples

**Example**: Get all bookmarks in the document, process them to log their names, then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    // Get the bookmark collection from the document
    const bookmarks = context.document.body.bookmarks;
    
    // Track the collection for processing
    context.trackedObjects.add(bookmarks);
    
    // Load bookmark properties
    bookmarks.load("items");
    await context.sync();
    
    // Process the bookmarks
    for (let i = 0; i < bookmarks.items.length; i++) {
        console.log(`Bookmark ${i + 1}: ${bookmarks.items[i].name}`);
    }
    
    // Untrack the collection to release memory
    bookmarks.untrack();
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.bookmarkcollection
