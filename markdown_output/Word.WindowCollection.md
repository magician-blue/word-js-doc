# Word.WindowCollection

**Package:** `word`

**API Set:** WordApiDesktop 1.2 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the collection of window objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from the WindowCollection to verify the connection between the add-in and Word application before performing window operations.

```typescript
await Word.run(async (context) => {
    const windows = context.application.windows;
    
    // Access the request context associated with the WindowCollection
    const requestContext = windows.context;
    
    // Use the context to load window properties
    windows.load("items");
    await requestContext.sync();
    
    console.log(`Connected to Word with ${windows.items.length} window(s) open`);
    console.log(`Request context is valid: ${requestContext !== null}`);
});
```

---

### items

**Type:** `Word.Window[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Log the titles of all open Word document windows to the console

```typescript
await Word.run(async (context) => {
    // Get the collection of all open windows
    const windows = context.application.windows;
    
    // Load the items property to access the array of windows
    windows.load("items");
    
    await context.sync();
    
    // Access the loaded windows through the items property
    const windowItems = windows.items;
    
    // Log the title of each window
    for (let i = 0; i < windowItems.length; i++) {
        windowItems[i].load("document/properties/title");
    }
    
    await context.sync();
    
    windowItems.forEach((window, index) => {
        console.log(`Window ${index + 1}: ${window.document.properties.title}`);
    });
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first window in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.Window`

#### Examples

**Example**: Get and activate the first open Word window to bring it into focus

```typescript
await Word.run(async (context) => {
    // Get the collection of all windows
    const windows = context.application.windows;
    
    // Get the first window in the collection
    const firstWindow = windows.getFirst();
    
    // Activate the first window to bring it into focus
    firstWindow.activate();
    
    await context.sync();
    
    console.log("First window has been activated");
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first window in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Window`

#### Examples

**Example**: Check if any document windows are open and display the first window's document name, or show a message if no windows are available.

```typescript
await Word.run(async (context) => {
    const firstWindow = context.application.windows.getFirstOrNullObject();
    firstWindow.load("isNullObject");
    
    await context.sync();
    
    if (firstWindow.isNullObject) {
        console.log("No windows are currently open.");
    } else {
        firstWindow.load("document/name");
        await context.sync();
        console.log(`First window document: ${firstWindow.document.name}`);
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
  - `options`: `Word.Interfaces.WindowCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (required)
    Provides options for which properties of the object to load.

  **Returns:** `Word.WindowCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (required)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.WindowCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (required)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.WindowCollection`

#### Examples

**Example**: Load and display the count of all open Word document windows

```typescript
await Word.run(async (context) => {
    // Get the collection of windows
    const windows = context.application.windows;
    
    // Load the count property of the windows collection
    windows.load("count");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the number of open windows
    console.log(`Number of open windows: ${windows.count}`);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.WindowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.WindowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.WindowCollectionData`

#### Examples

**Example**: Serialize the window collection to a plain JavaScript object and log it to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the collection of windows
    const windows = context.application.windows;
    
    // Load properties we want to serialize
    windows.load("items");
    
    await context.sync();
    
    // Convert the WindowCollection to a plain JavaScript object
    const windowsData = windows.toJSON();
    
    // Log the serialized data
    console.log("Windows data:", JSON.stringify(windowsData, null, 2));
    console.log("Number of windows:", windowsData.items.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.WindowCollection`

#### Examples

**Example**: Track the first window in the collection to maintain a reference across multiple sync calls when monitoring window properties

```typescript
await Word.run(async (context) => {
    const windows = context.document.application.windows;
    windows.load("items");
    await context.sync();
    
    const firstWindow = windows.items[0];
    
    // Track the window to use it across multiple sync calls
    firstWindow.track();
    
    // First sync - get initial state
    firstWindow.load("width");
    await context.sync();
    console.log("Initial width:", firstWindow.width);
    
    // Perform other operations...
    await context.sync();
    
    // Second sync - the tracked object remains valid
    firstWindow.load("width");
    await context.sync();
    console.log("Current width:", firstWindow.width);
    
    // Clean up tracking when done
    firstWindow.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.WindowCollection`

#### Examples

**Example**: Get all open Word windows, iterate through them to log their names, then untrack the collection to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the collection of windows
    const windows = context.application.windows;
    windows.load("items");
    
    await context.sync();
    
    // Use the windows collection
    console.log(`Number of windows: ${windows.items.length}`);
    windows.items.forEach(window => {
        console.log(`Window ID: ${window.id}`);
    });
    
    // Untrack the collection to release memory
    windows.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word/word.windowcollection
