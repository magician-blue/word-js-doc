# Word.TrackedChangeCollection

**Package:** `word`

**API Set:** WordApi 1.6

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.TrackedChange](/en-us/javascript/api/word/word.trackedchange) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the range of the first tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const range: Word.Range = trackedChange.getRange();
  range.load();
  await context.sync();

  console.log("range.text: " + range.text);
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TrackedChangeCollection to verify the connection between the add-in and Word before processing tracked changes.

```typescript
await Word.run(async (context) => {
    const trackedChanges = context.document.body.trackedChanges;
    trackedChanges.load("items");
    await context.sync();
    
    // Access the request context from the collection
    const requestContext = trackedChanges.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (requestContext === context) {
        console.log("TrackedChangeCollection is properly connected to the Word context");
        console.log(`Found ${trackedChanges.items.length} tracked changes`);
    }
});
```

---

### items

**Type:** `Word.TrackedChange[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all tracked changes in the document and log their text content to the console

```typescript
await Word.run(async (context) => {
    // Get the tracked changes collection from the document
    const trackedChanges = context.document.getTrackedChanges();
    
    // Load the items property to access the array of tracked changes
    trackedChanges.load("items");
    
    await context.sync();
    
    // Access the items array and iterate through each tracked change
    const trackedChangeItems = trackedChanges.items;
    
    for (let i = 0; i < trackedChangeItems.length; i++) {
        const change = trackedChangeItems[i];
        change.load("text,type");
    }
    
    await context.sync();
    
    // Log information about each tracked change
    trackedChangeItems.forEach((change, index) => {
        console.log(`Tracked Change ${index + 1}: ${change.type} - "${change.text}"`);
    });
});
```

---

## Methods

### acceptAll

**Kind:** `write`

Accepts all the tracked changes in the collection.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Accept all tracked changes in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Accepts all tracked changes.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  trackedChanges.acceptAll();
  console.log("Accepted all tracked changes.");
});
```

---

### getFirst

**Kind:** `read`

Gets the first TrackedChange in this collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.TrackedChange`

#### Examples

**Example**: Retrieve and display the text content of the first tracked change in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the range of the first tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const range: Word.Range = trackedChange.getRange();
  range.load();
  await context.sync();

  console.log("range.text: " + range.text);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first TrackedChange in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.TrackedChange`

#### Examples

**Example**: Check if there are any tracked changes in the document and display the text of the first tracked change, or show a message if no tracked changes exist.

```typescript
await Word.run(async (context) => {
    const trackedChanges = context.document.getTrackedChanges();
    const firstChange = trackedChanges.getFirstOrNullObject();
    
    firstChange.load("isNullObject, text, type");
    await context.sync();
    
    if (firstChange.isNullObject) {
        console.log("No tracked changes found in the document.");
    } else {
        console.log(`First tracked change: "${firstChange.text}" (Type: ${firstChange.type})`);
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TrackedChangeCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TrackedChangeCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TrackedChangeCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TrackedChangeCollection`

#### Examples

**Example**: Load and display the author and text of all tracked changes in the document

```typescript
await Word.run(async (context) => {
    // Get the collection of tracked changes
    const trackedChanges = context.document.body.trackedChanges;
    
    // Load specific properties of all tracked changes
    trackedChanges.load("author, text, type");
    
    // Synchronize to execute the load command
    await context.sync();
    
    // Display information about each tracked change
    console.log(`Found ${trackedChanges.items.length} tracked changes`);
    trackedChanges.items.forEach((change, index) => {
        console.log(`Change ${index + 1}:`);
        console.log(`  Author: ${change.author}`);
        console.log(`  Type: ${change.type}`);
        console.log(`  Text: ${change.text}`);
    });
});
```

---

### rejectAll

**Kind:** `write`

Rejects all the tracked changes in the collection.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Reject all tracked changes in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Rejects all tracked changes.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  trackedChanges.rejectAll();
  console.log("Rejected all tracked changes.");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TrackedChangeCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TrackedChangeCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.TrackedChangeCollectionData`

#### Examples

**Example**: Export tracked changes from the document to JSON format for logging or external processing

```typescript
await Word.run(async (context) => {
    // Get all tracked changes in the document
    const trackedChanges = context.document.body.trackedChanges;
    
    // Load properties we want to export
    trackedChanges.load("text, type, author, date");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const trackedChangesData = trackedChanges.toJSON();
    
    // Log or process the JSON data
    console.log("Tracked Changes JSON:", JSON.stringify(trackedChangesData, null, 2));
    
    // Access the items array
    console.log(`Found ${trackedChangesData.items.length} tracked changes`);
    trackedChangesData.items.forEach((change, index) => {
        console.log(`Change ${index + 1}: ${change.type} by ${change.author}`);
    });
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TrackedChangeCollection`

#### Examples

**Example**: Track all tracked changes in the document across multiple sync calls to monitor their properties and ensure they remain valid references throughout the batch operation.

```typescript
await Word.run(async (context) => {
    // Get all tracked changes in the document
    const trackedChanges = context.document.body.trackedChanges;
    trackedChanges.load("items");
    await context.sync();

    // Track the collection to maintain valid references across sync calls
    trackedChanges.track();

    // Perform operations across multiple sync calls
    console.log(`Found ${trackedChanges.items.length} tracked changes`);
    await context.sync();

    // The tracked collection remains valid for further operations
    for (let i = 0; i < trackedChanges.items.length; i++) {
        trackedChanges.items[i].load("text,type");
    }
    await context.sync();

    // Access properties after multiple syncs (tracking prevents InvalidObjectPath errors)
    trackedChanges.items.forEach((change) => {
        console.log(`Change type: ${change.type}, Text: ${change.text}`);
    });

    // Untrack when done to free resources
    trackedChanges.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TrackedChangeCollection`

#### Examples

**Example**: Load tracked changes from a document, process them, then untrack the collection to free memory and improve performance.

```typescript
await Word.run(async (context) => {
    // Get the tracked changes collection from the document
    const trackedChanges = context.document.body.trackedChanges;
    
    // Load properties to work with the collection
    trackedChanges.load("items");
    await context.sync();
    
    // Process the tracked changes (e.g., log count)
    console.log(`Found ${trackedChanges.items.length} tracked changes`);
    
    // Untrack the collection to release memory
    trackedChanges.untrack();
    
    // Sync to apply the memory release
    await context.sync();
    
    console.log("TrackedChangeCollection memory released");
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.trackedchangecollection
