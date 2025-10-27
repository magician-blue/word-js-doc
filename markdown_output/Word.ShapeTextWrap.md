# Word.ShapeTextWrap

**Package:** `word`

**API Set:** WordApiDesktop 1.2 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents all the properties for wrapping text around a shape.

## Properties

### bottomDistance

**Type:** `number`

**Since:** WordApiDesktop 1.2

Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.

#### Examples

**Example**: Set the bottom distance to 20 points between the document text and a shape's text-free area

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textWrap = shape.textWrap;
        textWrap.bottomDistance = 20;
        
        await context.sync();
        console.log("Bottom distance set to 20 points");
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a shape's text wrap object to verify the connection to the Word host application before modifying wrap properties.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textWrap = shape.textWrap;
        
        // Access the request context from the textWrap object
        const wrapContext = textWrap.context;
        
        // Verify the context is valid and connected
        console.log("Context is connected:", wrapContext !== null);
        
        // Use the context to load and sync properties
        textWrap.load("type");
        await wrapContext.sync();
        
        console.log("Text wrap type:", textWrap.type);
    }
});
```

---

### leftDistance

**Type:** `number`

**Since:** WordApiDesktop 1.2

Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.

#### Examples

**Example**: Set the left distance to 20 points between the document text and a shape's text-free area

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.textWrap.leftDistance = 20;
        await context.sync();
    }
});
```

---

### rightDistance

**Type:** `number`

**Since:** WordApiDesktop 1.2

Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.

#### Examples

**Example**: Set the right distance to 20 points between the document text and a shape's right edge to create proper spacing

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the right distance to 20 points
        shape.textWrap.rightDistance = 20;
        
        await context.sync();
        console.log("Right distance set to 20 points");
    }
});
```

---

### side

**Type:** `Word.ShapeTextWrapSide | "None" | "Both" | "Left" | "Right" | "Largest"`

**Since:** WordApiDesktop 1.2

Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.

#### Examples

**Example**: Set a shape's text wrapping to wrap on both sides of the shape

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textWrap = shape.textWrap;
        
        // Set text to wrap on both sides of the shape
        textWrap.side = Word.ShapeTextWrapSide.both;
        
        await context.sync();
    }
});
```

---

### topDistance

**Type:** `number`

**Since:** WordApiDesktop 1.2

Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.

#### Examples

**Example**: Set the top distance to 20 points between the document text and a shape's text-free area

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textWrap = shape.textWrap;
        textWrap.topDistance = 20;
        
        await context.sync();
    }
});
```

---

### type

**Type:** `Word.ShapeTextWrapType | "Inline" | "Square" | "Tight" | "Through" | "TopBottom" | "Behind" | "Front"`

**Since:** WordApiDesktop 1.2

Specifies the text wrap type around the shape. See `Word.ShapeTextWrapType` for details.

#### Examples

**Example**: Set the text wrap type of a shape to "Square" so that text flows around the shape in a square pattern.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    
    // Set the text wrap type to Square
    shape.textWrap.type = "Square";
    
    await context.sync();
    
    console.log("Text wrap type set to Square");
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
  - `options`: `Word.Interfaces.ShapeTextWrapLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ShapeTextWrap`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ShapeTextWrap`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ShapeTextWrap`

#### Examples

**Example**: Load and read the text wrapping type of the first shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const firstShape = shapes.getFirst();
    const textWrap = firstShape.textWrap;
    
    // Load the text wrap properties
    textWrap.load("type");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded property
    console.log("Text wrap type: " + textWrap.type);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ShapeTextWrapUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ShapeTextWrap` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure text wrapping settings for a shape to wrap text on both sides with specific distance margins

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textWrap = shape.textWrap;
    
    // Set multiple text wrap properties at once
    textWrap.set({
        type: Word.WrapType.square,
        side: Word.WrapSide.bothSides,
        distanceTop: 10,
        distanceBottom: 10,
        distanceLeft: 15,
        distanceRight: 15
    });
    
    await context.sync();
    console.log("Text wrap settings applied to shape");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeTextWrap` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeTextWrapData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ShapeTextWrapData`

#### Examples

**Example**: Get the text wrapping properties of a shape as a plain JavaScript object and log it to the console for debugging or serialization purposes.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textWrap = shape.textWrap;
        
        // Load the text wrap properties
        textWrap.load("type, side, distanceTop, distanceBottom, distanceLeft, distanceRight");
        await context.sync();
        
        // Convert to plain JavaScript object
        const textWrapData = textWrap.toJSON();
        
        // Log the plain object (useful for debugging or serialization)
        console.log("Text Wrap Properties:", textWrapData);
        console.log("Wrap Type:", textWrapData.type);
        console.log("Wrap Side:", textWrapData.side);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ShapeTextWrap`

#### Examples

**Example**: Track a shape's text wrap object across multiple sync calls to monitor and adjust wrap properties without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const textWrap = shape.textWrap;
    
    // Track the textWrap object for use across multiple sync calls
    textWrap.track();
    textWrap.load("type");
    await context.sync();
    
    console.log("Current wrap type:", textWrap.type);
    
    // Modify the wrap type in a subsequent sync call
    textWrap.type = Word.WrapType.square;
    await context.sync();
    
    console.log("Updated wrap type:", textWrap.type);
    
    // Untrack when done
    textWrap.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ShapeTextWrap`

#### Examples

**Example**: Access a shape's text wrapping properties, use them to check the wrapping type, then untrack the object to free memory

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const textWrap = shape.textWrap;
        
        // Load and track the text wrap properties
        textWrap.load("type");
        await context.sync();
        
        // Use the text wrap information
        console.log("Text wrap type: " + textWrap.type);
        
        // Release memory by untracking the object when done
        textWrap.untrack();
        await context.sync();
    }
});
```

---

## Source

- /en-us/javascript/api/word/word.shapetextwrap
