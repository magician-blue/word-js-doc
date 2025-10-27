# Word.BreakCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Break](/en-us/javascript/api/word/word.break) objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a BreakCollection to verify the connection to the Word host application and log its properties.

```typescript
await Word.run(async (context) => {
    // Get the breaks collection from the document body
    const body = context.document.body;
    const breaks = body.getBreaks();
    
    // Access the request context from the collection
    const requestContext = breaks.context;
    
    // Use the context to verify connection and perform operations
    console.log("Request context retrieved:", requestContext);
    console.log("Context is connected to Word application");
    
    // Load and sync using the context
    breaks.load("items");
    await context.sync();
    
    console.log(`Found ${breaks.items.length} breaks in the document`);
});
```

---

### items

**Type:** `Word.Break[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all page breaks in the document and display their count and types in the console.

```typescript
await Word.run(async (context) => {
    // Get the body of the document
    const body = context.document.body;
    
    // Get all breaks in the document
    const breaks = body.getBreaks();
    breaks.load("items");
    
    await context.sync();
    
    // Access the items property to get the array of Break objects
    const breakItems = breaks.items;
    
    console.log(`Total breaks found: ${breakItems.length}`);
    
    // Load type property for each break
    for (let i = 0; i < breakItems.length; i++) {
        breakItems[i].load("type");
    }
    
    await context.sync();
    
    // Display information about each break
    breakItems.forEach((breakItem, index) => {
        console.log(`Break ${index + 1}: ${breakItem.type}`);
    });
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.BreakCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.BreakCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BreakCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BreakCollection`

#### Examples

**Example**: Load and display the types of all breaks in the active document

```typescript
await Word.run(async (context) => {
    // Get all breaks in the document
    const breaks = context.document.body.getBreaks();
    
    // Load the type property for all breaks in the collection
    breaks.load("type");
    
    await context.sync();
    
    // Display the break types
    console.log(`Found ${breaks.items.length} breaks in the document`);
    breaks.items.forEach((breakItem, index) => {
        console.log(`Break ${index + 1}: ${breakItem.type}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.BreakCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BreakCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.BreakCollectionData`

#### Examples

**Example**: Get a JSON representation of all breaks in the document body for logging or data export purposes

```typescript
await Word.run(async (context) => {
    // Get all breaks in the document body
    const breaks = context.document.body.getBreaks();
    
    // Load properties needed for the breaks
    breaks.load("type");
    
    await context.sync();
    
    // Convert the BreakCollection to a plain JavaScript object
    const breaksJSON = breaks.toJSON();
    
    // Log the JSON representation
    console.log("Breaks in document:", JSON.stringify(breaksJSON, null, 2));
    
    // The breaksJSON object contains an "items" array with break data
    console.log(`Total breaks found: ${breaksJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BreakCollection`

#### Examples

**Example**: Track a collection of breaks across multiple sync calls to monitor and adjust them as the document changes, preventing "InvalidObjectPath" errors when accessing the collection after document modifications.

```typescript
await Word.run(async (context) => {
    // Get the breaks collection from the document body
    const breaks = context.document.body.getBreaks();
    
    // Track the collection for automatic adjustment across sync calls
    breaks.track();
    
    // Load the breaks collection
    breaks.load("items");
    await context.sync();
    
    console.log(`Found ${breaks.items.length} breaks in the document`);
    
    // Make some changes to the document
    context.document.body.insertText("New content added", Word.InsertLocation.start);
    await context.sync();
    
    // Access the tracked collection again after document changes
    // Without track(), this might throw "InvalidObjectPath" error
    breaks.load("items");
    await context.sync();
    
    console.log(`After changes: ${breaks.items.length} breaks`);
    
    // Untrack when done to free up memory
    breaks.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.BreakCollection`

#### Examples

**Example**: Get all breaks in the document, process them, and then untrack the collection to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the break collection from the document body
    const breaks = context.document.body.getBreaks();
    
    // Track the collection for memory management
    breaks.track();
    
    // Load the breaks to work with them
    breaks.load("items");
    await context.sync();
    
    // Process the breaks (e.g., log the count)
    console.log(`Found ${breaks.items.length} breaks in the document`);
    
    // Untrack the collection to release memory
    breaks.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
