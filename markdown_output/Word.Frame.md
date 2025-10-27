# Frame

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a frame. The Frame object is a member of the Word.FrameCollection object.

## Properties

### borders

**Type:** `Word.BorderUniversalCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BorderUniversalCollection object that represents all the borders for the frame.

#### Examples

**Example**: Set the frame's top border to a double line style with blue color and 2.25pt width

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const frame = frames.getFirst();
    
    // Access the borders collection
    const borders = frame.borders;
    const topBorder = borders.getItem(Word.BorderLocation.top);
    
    // Configure the top border
    topBorder.type = Word.BorderType.double;
    topBorder.color = "#0000FF"; // Blue
    topBorder.width = 2.25;
    
    await context.sync();
    console.log("Frame top border configured successfully");
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a frame object to verify the connection between the add-in and Word application before performing operations on the frame.

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    frames.load("items");
    await context.sync();
    
    if (frames.items.length > 0) {
        const frame = frames.items[0];
        
        // Access the request context associated with the frame
        const frameContext = frame.context;
        
        // Use the context to load properties and sync
        frame.load("width,height");
        await frameContext.sync();
        
        console.log(`Frame dimensions: ${frame.width} x ${frame.height}`);
    }
});
```

---

### height

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the height (in points) of the frame.

#### Examples

**Example**: Set the height of the first frame in the document to 150 points

```typescript
await Word.run(async (context) => {
    const frames = context.document.body.frameCollection;
    frames.load("items");
    await context.sync();
    
    if (frames.items.length > 0) {
        const firstFrame = frames.items[0];
        firstFrame.height = 150;
        await context.sync();
    }
});
```

---

### heightRule

**Type:** `Word.FrameSizeRule | "Auto" | "AtLeast" | "Exact"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a FrameSizeRule value that represents the rule for determining the height of the frame.

#### Examples

**Example**: Set a frame's height rule to "AtLeast" to ensure the frame is at least a minimum height but can expand if needed

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const frame = frames.getFirst();
    
    // Set the height rule to "AtLeast"
    frame.heightRule = "AtLeast";
    
    await context.sync();
    console.log("Frame height rule set to AtLeast");
});
```

---

### horizontalDistanceFromText

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the horizontal distance between the frame and the surrounding text, in points.

#### Examples

**Example**: Set the horizontal distance between a frame and its surrounding text to 20 points

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frames;
    frames.load("items");
    await context.sync();
    
    if (frames.items.length > 0) {
        const frame = frames.items[0];
        
        // Set horizontal distance from text to 20 points
        frame.horizontalDistanceFromText = 20;
        
        await context.sync();
        console.log("Horizontal distance from text set to 20 points");
    }
});
```

---

### horizontalPosition

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the horizontal distance between the edge of the frame and the item specified by the relativeHorizontalPosition property.

#### Examples

**Example**: Set a frame's horizontal position to 72 points (1 inch) from its relative horizontal anchor point

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const frame = frames.getFirst();
    
    // Set the horizontal position to 72 points from the anchor
    frame.horizontalPosition = 72;
    
    await context.sync();
});
```

---

### lockAnchor

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the frame is locked.

#### Examples

**Example**: Lock a frame's anchor position to prevent it from moving when content is added or removed from the document.

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frames;
    frames.load("items");
    await context.sync();
    
    if (frames.items.length > 0) {
        const frame = frames.items[0];
        
        // Lock the frame's anchor
        frame.lockAnchor = true;
        
        await context.sync();
        console.log("Frame anchor has been locked");
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a Range object that represents the portion of the document that's contained within the frame.

#### Examples

**Example**: Get the text content from a frame and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const firstFrame = frames.getFirst();
    
    // Get the range contained within the frame
    const frameRange = firstFrame.range;
    frameRange.load("text");
    
    await context.sync();
    
    // Display the text content from the frame
    console.log("Frame content: " + frameRange.text);
});
```

---

### relativeHorizontalPosition

**Type:** `Word.RelativeHorizontalPosition | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the relative horizontal position of the frame.

#### Examples

**Example**: Set a frame's horizontal position to be relative to the page margins

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameCollection;
    frames.load("items");
    await context.sync();
    
    if (frames.items.length > 0) {
        const frame = frames.items[0];
        
        // Set the frame's horizontal position to be relative to the margin
        frame.relativeHorizontalPosition = Word.RelativeHorizontalPosition.margin;
        
        await context.sync();
        console.log("Frame horizontal position set to margin");
    }
});
```

---

### relativeVerticalPosition

**Type:** `Word.RelativeVerticalPosition | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the relative vertical position of the frame.

#### Examples

**Example**: Set a frame's vertical position to be relative to the page margins

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameCollection;
    frames.load("items");
    await context.sync();
    
    if (frames.items.length > 0) {
        const frame = frames.items[0];
        
        // Set the frame's vertical position to be relative to the page margin
        frame.relativeVerticalPosition = Word.RelativeVerticalPosition.margin;
        
        await context.sync();
        console.log("Frame vertical position set to margin");
    }
});
```

---

### shading

**Type:** `Word.ShadingUniversal`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a ShadingUniversal object that refers to the shading formatting for the frame.

#### Examples

**Example**: Apply yellow background shading to a frame in the document

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const frame = frames.getFirst();
    
    // Access the shading property and set background color to yellow
    frame.shading.backgroundPatternColor = "yellow";
    
    await context.sync();
});
```

---

### textWrap

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if document text wraps around the frame.

#### Examples

**Example**: Enable text wrapping around a frame so that document text flows around it instead of being displaced by it.

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.framesets.getFirst().frames;
    const frame = frames.getFirst();
    
    // Enable text wrapping around the frame
    frame.textWrap = true;
    
    await context.sync();
    
    console.log("Text wrapping enabled for the frame");
});
```

---

### verticalDistanceFromText

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical distance (in points) between the frame and the surrounding text.

#### Examples

**Example**: Set the vertical distance between a frame and its surrounding text to 12 points

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frames;
    frames.load("items");
    await context.sync();
    
    if (frames.items.length > 0) {
        const frame = frames.items[0];
        
        // Set the vertical distance from text to 12 points
        frame.verticalDistanceFromText = 12;
        
        await context.sync();
        console.log("Vertical distance from text set to 12 points");
    }
});
```

---

### verticalPosition

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical distance between the edge of the frame and the item specified by the relativeVerticalPosition property.

#### Examples

**Example**: Set a frame's vertical position to 50 points from the top margin

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const frame = frames.getFirst();
    
    // Set the vertical position to 50 points
    frame.verticalPosition = 50;
    
    await context.sync();
});
```

---

### width

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width (in points) of the frame.

#### Examples

**Example**: Set the width of the first frame in the document to 300 points

```typescript
await Word.run(async (context) => {
    const frames = context.document.body.frameCollection;
    frames.load("items");
    await context.sync();
    
    if (frames.items.length > 0) {
        const firstFrame = frames.items[0];
        firstFrame.width = 300;
        await context.sync();
    }
});
```

---

### widthRule

**Type:** `Word.FrameSizeRule | "Auto" | "AtLeast" | "Exact"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the rule used to determine the width of the frame.

#### Examples

**Example**: Set a frame's width rule to "Exact" to enforce a fixed width of 200 points

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const frame = frames.getFirst();
    
    // Set the width rule to "Exact" for fixed width
    frame.widthRule = "Exact";
    frame.width = 200; // Set width to 200 points
    
    await context.sync();
    
    console.log("Frame width rule set to Exact with 200 points width");
});
```

---

## Methods

### copy

Copies the frame to the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Copy the first frame in the document to the Clipboard so it can be pasted elsewhere

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frames;
    const firstFrame = frames.getFirst();
    
    // Copy the frame to the Clipboard
    firstFrame.copy();
    
    await context.sync();
    
    console.log("Frame copied to Clipboard");
});
```

---

### cut

Removes the frame from the document and places it on the Clipboard.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove the first frame from the document and place it on the Clipboard

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameCollection;
    frames.load("items");
    
    await context.sync();
    
    if (frames.items.length > 0) {
        const firstFrame = frames.items[0];
        
        // Cut the frame to the Clipboard
        firstFrame.cut();
        
        await context.sync();
        console.log("Frame has been cut to the Clipboard");
    } else {
        console.log("No frames found in the document");
    }
});
```

---

### delete

**Kind:** `delete`

Deletes the frame.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete the first frame in the document

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameCollection;
    const firstFrame = frames.getFirst();
    
    // Delete the frame
    firstFrame.delete();
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.FrameLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Frame`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Frame`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Frame`

#### Examples

**Example**: Load and display the width and height properties of the first frame in the document.

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const firstFrame = frames.getFirst();
    
    // Load the width and height properties of the frame
    firstFrame.load("width, height");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the loaded properties
    console.log(`Frame width: ${firstFrame.width}`);
    console.log(`Frame height: ${firstFrame.height}`);
});
```

---

### select

Selects the frame.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Select the first frame in the document to highlight it for the user

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frames;
    const firstFrame = frames.getFirst();
    
    // Select the frame
    firstFrame.select();
    
    await context.sync();
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.FrameUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Frame` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple frame properties at once to set the frame's width to 200 points and height to 150 points

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frames;
    frames.load("items");
    await context.sync();
    
    const frame = frames.items[0];
    
    // Set multiple properties at once using the set() method
    frame.set({
        width: 200,
        height: 150
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Frame object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FrameData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.FrameData`

#### Examples

**Example**: Serialize a frame's properties to a plain JavaScript object and log it to the console for debugging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.framesets.getFirst().frames;
    const frame = frames.getFirst();
    
    // Load properties we want to serialize
    frame.load("width,height,left,top");
    
    await context.sync();
    
    // Convert the frame to a plain JavaScript object
    const frameData = frame.toJSON();
    
    // Log the serialized data
    console.log("Frame data:", JSON.stringify(frameData, null, 2));
    
    // The frameData object can now be used for storage, comparison, or transmission
    // It contains only the loaded properties as plain JavaScript values
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Frame`

#### Examples

**Example**: Track a frame object to maintain its reference across multiple sync calls while modifying its properties and content in separate operations.

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    frames.load("items");
    await context.sync();
    
    const frame = frames.items[0];
    
    // Track the frame object for use across multiple sync calls
    frame.track();
    
    // Load and modify frame properties in first operation
    frame.load("width,height");
    await context.sync();
    
    console.log(`Original size: ${frame.width}pt x ${frame.height}pt`);
    
    // Modify the frame in a second operation
    frame.width = 300;
    frame.height = 200;
    await context.sync();
    
    console.log("Frame resized successfully");
    
    // Untrack when done to release memory
    frame.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Frame`

#### Examples

**Example**: Get a frame from the document, perform operations on it, then untrack it to release memory after you're done using it.

```typescript
await Word.run(async (context) => {
    // Get the first frame in the document
    const frames = context.document.body.frameSet.frames;
    const firstFrame = frames.getFirst();
    
    // Track the frame object for changes
    firstFrame.track();
    
    // Load properties to work with
    firstFrame.load("width,height");
    await context.sync();
    
    // Perform operations with the frame
    console.log(`Frame dimensions: ${firstFrame.width} x ${firstFrame.height}`);
    
    // Untrack the frame to release memory when done
    firstFrame.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
