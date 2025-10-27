# Word.HyperlinkCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of Word.Hyperlink objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a hyperlink collection to verify the connection between the add-in and Word before performing operations on hyperlinks in the document.

```typescript
await Word.run(async (context) => {
    // Get the hyperlink collection from the document body
    const hyperlinkCollection = context.document.body.hyperlinks;
    
    // Access the request context associated with the hyperlink collection
    const requestContext = hyperlinkCollection.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (requestContext === context) {
        console.log("Request context is properly connected to Word");
        
        // Now safe to perform operations using this context
        hyperlinkCollection.load("items");
        await context.sync();
        
        console.log(`Found ${hyperlinkCollection.items.length} hyperlinks`);
    }
});
```

---

### items

**Type:** `Word.Hyperlink[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all hyperlinks in the document and log their display text and addresses to the console.

```typescript
await Word.run(async (context) => {
    // Get the hyperlinks collection from the document body
    const hyperlinks = context.document.body.hyperlinks;
    
    // Load the items property to access the array of hyperlinks
    hyperlinks.load("items");
    
    await context.sync();
    
    // Access the loaded hyperlinks through the items property
    const hyperlinkItems = hyperlinks.items;
    
    // Log information about each hyperlink
    for (let i = 0; i < hyperlinkItems.length; i++) {
        hyperlinkItems[i].load("textToDisplay, address");
    }
    
    await context.sync();
    
    hyperlinkItems.forEach((hyperlink, index) => {
        console.log(`Hyperlink ${index + 1}: "${hyperlink.textToDisplay}" -> ${hyperlink.address}`);
    });
});
```

---

## Methods

### add

**Kind:** `create`

Returns a Hyperlink object that represents a new hyperlink added to a range, selection, or document.

#### Signature

**Parameters:**
- `anchor`: `Word.Range` (required)
  The range to which the hyperlink is added.
- `options`: `Word.HyperlinkAddOptions` (optional)
  The options to further configure the new hyperlink.

**Returns:** `Word.Hyperlink`

#### Examples

**Example**: Add a hyperlink to the selected text that links to a website

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const selection = context.document.getSelection();
    
    // Add a hyperlink to the selection
    const hyperlink = selection.hyperlinks.add(
        selection,
        {
            address: "https://www.example.com",
            screenTip: "Visit Example Website"
        }
    );
    
    await context.sync();
    
    console.log("Hyperlink added successfully");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.HyperlinkCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.HyperlinkCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.HyperlinkCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.HyperlinkCollection`

#### Examples

**Example**: Load and display the text and address of all hyperlinks in the document

```typescript
await Word.run(async (context) => {
    // Get all hyperlinks in the document body
    const hyperlinks = context.document.body.hyperlinks;
    
    // Load the text and address properties of all hyperlinks
    hyperlinks.load("text, address");
    
    await context.sync();
    
    // Display the hyperlink information
    console.log(`Found ${hyperlinks.items.length} hyperlinks:`);
    hyperlinks.items.forEach((hyperlink, index) => {
        console.log(`${index + 1}. Text: "${hyperlink.text}", Address: ${hyperlink.address}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.HyperlinkCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.HyperlinkCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.HyperlinkCollectionData`

#### Examples

**Example**: Export all hyperlinks in the document to a JSON string for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get all hyperlinks in the document
    const hyperlinks = context.document.body.getHyperlinks();
    
    // Load the properties we want to export
    hyperlinks.load("items/address, items/screenTip, items/textToDisplay");
    
    await context.sync();
    
    // Convert the hyperlink collection to a plain JavaScript object
    const hyperlinkData = hyperlinks.toJSON();
    
    // Convert to JSON string for logging or storage
    const jsonString = JSON.stringify(hyperlinkData, null, 2);
    
    console.log("Hyperlinks as JSON:", jsonString);
    
    // Example: You could also access the items array directly
    console.log(`Found ${hyperlinkData.items.length} hyperlinks`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.HyperlinkCollection`

#### Examples

**Example**: Track all hyperlinks in the document across multiple sync calls to monitor and update their properties without encountering InvalidObjectPath errors.

```typescript
await Word.run(async (context) => {
    const hyperlinks = context.document.body.hyperlinks;
    hyperlinks.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    hyperlinks.track();
    
    // First sync - get initial count
    await context.sync();
    console.log(`Found ${hyperlinks.items.length} hyperlinks`);
    
    // Second sync - can still access the collection safely
    await context.sync();
    
    // Load properties of each hyperlink
    for (let i = 0; i < hyperlinks.items.length; i++) {
        hyperlinks.items[i].load("address, text");
    }
    await context.sync();
    
    // Display hyperlink information
    hyperlinks.items.forEach((hyperlink, index) => {
        console.log(`Hyperlink ${index + 1}: ${hyperlink.text} -> ${hyperlink.address}`);
    });
    
    // Untrack when done
    hyperlinks.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them.

#### Signature

**Returns:** `Word.HyperlinkCollection`

#### Examples

**Example**: Load all hyperlinks in the document, process them to get their URLs, then untrack the collection to free memory.

```typescript
await Word.run(async (context) => {
    // Get the hyperlink collection from the document body
    const hyperlinks = context.document.body.hyperlinks;
    
    // Track the collection for memory management
    hyperlinks.track();
    
    // Load the hyperlink properties
    hyperlinks.load("items");
    
    await context.sync();
    
    // Process the hyperlinks (e.g., log their URLs)
    for (let i = 0; i < hyperlinks.items.length; i++) {
        console.log(`Hyperlink ${i + 1}: ${hyperlinks.items[i].address}`);
    }
    
    // Release the memory associated with the tracked collection
    hyperlinks.untrack();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlinkcollection
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
