# Canvas

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a canvas in the document. To get the corresponding Shape object, use Canvas.shape.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the Canvas's request context to verify the add-in is properly connected to the Word host application before performing operations on the canvas.

```typescript
await Word.run(async (context) => {
    // Get the first canvas in the document
    const canvases = context.document.body.inlinePictures;
    const canvas = canvases.getFirst() as Word.Canvas;
    
    // Access the request context associated with the canvas
    const canvasContext = canvas.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (canvasContext) {
        console.log("Canvas is properly connected to the Word host application");
        
        // Use the context to load and sync canvas properties
        canvas.load("width,height");
        await canvasContext.sync();
        
        console.log(`Canvas dimensions: ${canvas.width}x${canvas.height}`);
    }
});
```

---

### id

**Type:** `number`

**Since:** WordApiDesktop 1.2

Gets an integer that represents the canvas identifier.

#### Examples

**Example**: Get the canvas identifier and display it in the console to track which canvas is being processed in the document.

```typescript
await Word.run(async (context) => {
    const canvas = context.document.body.getCanvases().getFirst();
    canvas.load("id");
    
    await context.sync();
    
    console.log(`Canvas ID: ${canvas.id}`);
});
```

---

### shape

**Type:** `Word.Shape`

**Since:** WordApiDesktop 1.2

Gets the Shape object associated with the canvas.

#### Examples

**Example**: Get the canvas's associated shape and set its width to 200 pixels and height to 150 pixels.

```typescript
await Word.run(async (context) => {
    const canvas = context.document.body.insertCanvas(100, 100, Word.InsertLocation.end);
    
    // Get the Shape object associated with the canvas
    const canvasShape = canvas.shape;
    
    // Set dimensions using the shape object
    canvasShape.width = 200;
    canvasShape.height = 150;
    
    await context.sync();
});
```

---

### shapes

**Type:** `Word.ShapeCollection`

**Since:** WordApiDesktop 1.2

Gets the collection of Shape objects. Currently, only text boxes, pictures, and geometric shapes are supported.

#### Examples

**Example**: Get all shapes within a canvas and log their count and types to the console

```typescript
await Word.run(async (context) => {
    // Get the first canvas in the document
    const canvases = context.document.body.inlinePictures.getFirst().getAsCanvasOrNullObject();
    const canvas = context.document.body.canvases.getFirst();
    
    // Get the shapes collection from the canvas
    const shapes = canvas.shapes;
    shapes.load("items");
    
    await context.sync();
    
    // Log the count and types of shapes
    console.log(`Total shapes in canvas: ${shapes.items.length}`);
    
    shapes.items.forEach((shape, index) => {
        shape.load("type");
    });
    
    await context.sync();
    
    shapes.items.forEach((shape, index) => {
        console.log(`Shape ${index + 1}: ${shape.type}`);
    });
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.CanvasLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Canvas`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Canvas`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Canvas`

#### Examples

**Example**: Load and read the width and height properties of the first canvas in the document

```typescript
await Word.run(async (context) => {
    // Get the first canvas in the document
    const canvas = context.document.body.inlinePictures.getFirst().canvas;
    
    // Load the width and height properties
    canvas.load("width, height");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log(`Canvas width: ${canvas.width}`);
    console.log(`Canvas height: ${canvas.height}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.CanvasUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Canvas` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a canvas element at once, setting its width, height, and left position

```typescript
await Word.run(async (context) => {
    // Get the first canvas in the document
    const canvas = context.document.body.inlinePictures.getFirst().getAsCanvasOrNullObject();
    
    // Set multiple properties at once
    canvas.set({
        width: 400,
        height: 300,
        left: 50
    });
    
    await context.sync();
    console.log("Canvas properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Canvas object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CanvasData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CanvasData`

#### Examples

**Example**: Serialize a canvas object to JSON format to log or store its properties

```typescript
await Word.run(async (context) => {
    // Get the first canvas in the document
    const canvases = context.document.body.inlinePictures.getFirst().getOrNullObject();
    const canvas = context.document.body.canvases.getFirst();
    
    // Load properties of the canvas
    canvas.load("id,height,width");
    
    await context.sync();
    
    // Convert the canvas object to a plain JavaScript object
    const canvasJSON = canvas.toJSON();
    
    // Log the serialized canvas data
    console.log("Canvas data:", JSON.stringify(canvasJSON, null, 2));
    
    // The JSON object can now be stored or transmitted
    return canvasJSON;
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Canvas`

#### Examples

**Example**: Track a canvas object to maintain its reference across multiple sync calls while modifying its properties

```typescript
await Word.run(async (context) => {
    // Get the first canvas in the document
    const canvases = context.document.body.inlinePictures;
    canvases.load("items");
    await context.sync();
    
    // Assume the first inline picture is a canvas
    const canvas = canvases.items[0].getAsCanvasOrNullObject() as Word.Canvas;
    canvas.load("width,height");
    
    // Track the canvas object for use across multiple sync calls
    canvas.track();
    
    await context.sync();
    
    // Now we can safely use the canvas across multiple operations
    console.log(`Canvas dimensions: ${canvas.width} x ${canvas.height}`);
    
    // Perform additional operations after sync
    const shape = canvas.shape;
    shape.load("width");
    await context.sync();
    
    console.log(`Shape width: ${shape.width}`);
    
    // Untrack when done
    canvas.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Canvas`

#### Examples

**Example**: Track a canvas object to work with it across multiple sync() calls, then untrack it to free memory when done processing

```typescript
await Word.run(async (context) => {
    // Get the first canvas in the document
    const canvas = context.document.body.inlinePictures.getFirst().getAsCanvasOrNullObject();
    
    // Track the canvas object to maintain reference across sync calls
    canvas.track();
    await context.sync();
    
    // Work with the canvas across multiple operations
    if (!canvas.isNullObject) {
        // Perform operations with the canvas
        const shape = canvas.shape;
        shape.load("width,height");
        await context.sync();
        
        console.log(`Canvas dimensions: ${shape.width} x ${shape.height}`);
        
        // Untrack the canvas to release memory when done
        canvas.untrack();
        await context.sync();
    }
});
```

---

## Source

- /en-us/javascript/api/word
