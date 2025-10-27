# FrameCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the collection of [Word.Frame](/en-us/javascript/api/word/word.frame) objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a FrameCollection to verify the connection between the add-in and Word before performing operations on frames.

```typescript
await Word.run(async (context) => {
    const frames = context.document.body.getFrames();
    
    // Access the request context associated with the FrameCollection
    const frameContext = frames.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (frameContext === context) {
        console.log("FrameCollection is properly connected to the Word context");
        
        // Load and sync using the context
        frames.load("items");
        await context.sync();
        
        console.log(`Found ${frames.items.length} frames in the document`);
    }
});
```

---

### items

**Type:** `Word.Frame[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all frames in the document and log the count and text content of each frame to the console.

```typescript
await Word.run(async (context) => {
    // Get the frame collection from the document body
    const frames = context.document.body.framesets.getFirst().frames;
    
    // Load the items property to access the array of frames
    frames.load("items");
    
    await context.sync();
    
    // Access the loaded frames using the items property
    console.log(`Total frames: ${frames.items.length}`);
    
    // Iterate through each frame in the items array
    for (let i = 0; i < frames.items.length; i++) {
        const frame = frames.items[i];
        frame.load("text");
        await context.sync();
        
        console.log(`Frame ${i + 1}: ${frame.text}`);
    }
});
```

---

## Methods

### add

**Kind:** `create`

Returns a `Frame` object that represents a new frame added to a range, selection, or document.

#### Signature

**Parameters:**
- `range`: `Word.Range` (required)
  The range where the frame will be added.

**Returns:** `Word.Frame`
A `Frame` object that represents the new frame.

#### Examples

**Example**: Add a frame around the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.load("text");
    
    await context.sync();
    
    // Add a frame around the first paragraph
    const frame = context.document.frames.add(firstParagraph.getRange());
    
    await context.sync();
    
    console.log("Frame added around the first paragraph");
});
```

---

### delete

**Kind:** `delete`

Deletes the `FrameCollection` object.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete all frames from the active Word document

```typescript
await Word.run(async (context) => {
    // Get all frames in the document
    const frames = context.document.body.frameSet.frames;
    
    // Load the frame collection
    frames.load("items");
    await context.sync();
    
    // Delete all frames
    frames.delete();
    
    await context.sync();
});
```

---

### getItem

**Kind:** `read`

Gets a `Frame` object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  The location of a `Frame` object.

**Returns:** `Word.Frame`

#### Examples

**Example**: Get the first frame in the document and change its border color to red.

```typescript
await Word.run(async (context) => {
    // Get the collection of frames in the document
    const frames = context.document.body.frames;
    
    // Get the first frame by index
    const firstFrame = frames.getItem(0);
    
    // Load the frame's properties
    firstFrame.load("borderColor");
    
    await context.sync();
    
    // Change the border color to red
    firstFrame.borderColor = "#FF0000";
    
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
  - `options`: `Word.Interfaces.FrameCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.FrameCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.FrameCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.FrameCollection`

#### Examples

**Example**: Load and display the text content of all frames in the document

```typescript
await Word.run(async (context) => {
    // Get all frames in the document
    const frames = context.document.body.frameCollection;
    
    // Load the text content property for all frames
    frames.load("items/textContent");
    
    await context.sync();
    
    // Display the text content of each frame
    frames.items.forEach((frame, index) => {
        console.log(`Frame ${index + 1}: ${frame.textContent}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.FrameCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FrameCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.FrameCollectionData`

#### Examples

**Example**: Export frame collection data to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get all frames in the document
    const frames = context.document.body.framesets.getFirst().frames;
    
    // Load properties needed for the frames
    frames.load("items");
    
    await context.sync();
    
    // Convert the frame collection to a plain JavaScript object
    const framesJSON = frames.toJSON();
    
    // Log the JSON representation (can be used for debugging or data export)
    console.log(JSON.stringify(framesJSON, null, 2));
    
    // The framesJSON object contains an "items" array with frame data
    console.log(`Number of frames: ${framesJSON.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.FrameCollection`

#### Examples

**Example**: Track all frames in the document to monitor and work with them across multiple sync calls without getting InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get all frames in the document
    const frames = context.document.body.frames;
    
    // Track the frames collection for use across sync calls
    frames.track();
    
    // Load frame properties
    frames.load("items");
    await context.sync();
    
    // Now we can safely work with frames across multiple sync calls
    console.log(`Found ${frames.items.length} frames`);
    
    // Perform operations across sync boundaries
    for (let i = 0; i < frames.items.length; i++) {
        frames.items[i].load("width,height");
    }
    await context.sync();
    
    // Access properties after another sync
    frames.items.forEach((frame, index) => {
        console.log(`Frame ${index}: ${frame.width}x${frame.height}`);
    });
    
    // Untrack when done to release memory
    frames.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.FrameCollection`

#### Examples

**Example**: Load frame collection, process the frames, then untrack the collection to free memory after use

```typescript
await Word.run(async (context) => {
    // Get the frames collection from the document body
    const frames = context.document.body.framesets.getFirst().frames;
    
    // Track the collection for processing
    frames.track();
    
    // Load properties we need
    frames.load("items");
    await context.sync();
    
    // Process the frames (e.g., log count)
    console.log(`Found ${frames.items.length} frames`);
    
    // Untrack the collection to release memory
    frames.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
