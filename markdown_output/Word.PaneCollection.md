# PaneCollection

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the collection of pane.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a PaneCollection to verify the connection between the add-in and Word application before performing pane operations.

```typescript
await Word.run(async (context) => {
    const panes = context.document.panes;
    
    // Access the request context associated with the PaneCollection
    const requestContext = panes.context;
    
    // Verify the context is properly connected
    if (requestContext) {
        console.log("PaneCollection is connected to Word application");
        
        // Load pane properties using the context
        panes.load("items");
        await context.sync();
        
        console.log(`Number of panes: ${panes.items.length}`);
    }
});
```

---

### items

**Type:** `Word.Pane[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all loaded panes in the document and log the count of available panes to the console.

```typescript
await Word.run(async (context) => {
    const panes = context.document.panes;
    panes.load("items");
    
    await context.sync();
    
    console.log(`Number of panes: ${panes.items.length}`);
    
    // Access individual panes from the items array
    panes.items.forEach((pane, index) => {
        console.log(`Pane ${index + 1} found`);
    });
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first pane in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.Pane`

#### Examples

**Example**: Get the first pane in the document and activate it to bring it into focus.

```typescript
await Word.run(async (context) => {
    const panes = context.document.panes;
    const firstPane = panes.getFirst();
    
    firstPane.activate();
    
    await context.sync();
    console.log("First pane has been activated");
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first pane in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Pane`

#### Examples

**Example**: Check if any panes exist in the document and display the first pane's view type, or handle the case when no panes are available.

```typescript
await Word.run(async (context) => {
    const panes = context.document.panes;
    const firstPane = panes.getFirstOrNullObject();
    firstPane.load("isNullObject, view/type");
    
    await context.sync();
    
    if (firstPane.isNullObject) {
        console.log("No panes available in the document.");
    } else {
        console.log("First pane view type: " + firstPane.view.type);
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
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.PaneCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.PaneCollection`

#### Examples

**Example**: Load and display the count of panes in the active Word document

```typescript
await Word.run(async (context) => {
    const panes = context.document.panes;
    
    // Load the count property of the pane collection
    panes.load("count");
    
    await context.sync();
    
    console.log(`Number of panes: ${panes.count}`);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.PaneCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.PaneCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.PaneCollectionData`

#### Examples

**Example**: Serialize the panes collection to a JSON object for logging or debugging purposes

```typescript
await Word.run(async (context) => {
    // Get the panes collection from the document
    const panes = context.document.panes;
    
    // Load the panes collection with their properties
    panes.load("items");
    
    await context.sync();
    
    // Convert the panes collection to a plain JavaScript object
    const panesJSON = panes.toJSON();
    
    // Log the JSON representation
    console.log("Panes collection as JSON:", JSON.stringify(panesJSON, null, 2));
    
    // You can now work with the plain JavaScript object
    console.log(`Number of panes: ${panesJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.PaneCollection`

#### Examples

**Example**: Track a pane object to maintain its reference across multiple sync calls when working with document panes in a Word add-in.

```typescript
await Word.run(async (context) => {
    // Get the pane collection
    const panes = context.document.panes;
    
    // Load the first pane
    const firstPane = panes.getFirst();
    firstPane.load("id");
    await context.sync();
    
    // Track the pane object to use it across multiple sync calls
    firstPane.track();
    
    // Perform additional operations that require sync
    firstPane.load("view/zoom");
    await context.sync();
    
    console.log("Pane ID: " + firstPane.id);
    console.log("Pane zoom: " + firstPane.view.zoom);
    
    // Untrack when done to free up memory
    firstPane.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.PaneCollection`

#### Examples

**Example**: Access panes in a document, track them for performance optimization, then release the tracked objects when done to free memory.

```typescript
await Word.run(async (context) => {
    // Get the panes collection and load its properties
    const panes = context.document.panes;
    panes.load("items");
    
    await context.sync();
    
    // Track the panes collection for performance
    panes.track();
    
    // Perform operations with the panes
    console.log(`Number of panes: ${panes.items.length}`);
    
    // When done, untrack to release memory
    panes.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
