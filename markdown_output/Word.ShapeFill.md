# ShapeFill

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the fill formatting of a shape object.

## Properties

### backgroundColor

**Type:** `string`

**Since:** WordApiDesktop 1.2

Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.

#### Examples

**Example**: Set the background color of a shape's fill to light blue using a hex color code

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    // Get the first shape in the document
    const shape = shapes.items[0];
    
    // Set the fill background color to light blue
    shape.fill.backgroundColor = "#ADD8E6";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ShapeFill object to verify the connection to the Word host application before applying fill formatting.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shapeFill = shape.fill;
        
        // Access the request context from the ShapeFill object
        const fillContext = shapeFill.context;
        
        // Verify the context is valid and connected
        console.log("ShapeFill context is connected:", fillContext !== null);
        
        // Use the context to perform operations
        shapeFill.setSolidColor("blue");
        await context.sync();
        
        console.log("Fill formatting applied successfully");
    }
});
```

---

### foregroundColor

**Type:** `string`

**Since:** WordApiDesktop 1.2

Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.

#### Examples

**Example**: Set the foreground color of a shape's fill to blue

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shape = shapes.getItemAt(0);
    
    shape.fill.foregroundColor = "#0000FF";
    
    await context.sync();
});
```

---

### transparency

**Type:** `number`

**Since:** WordApiDesktop 1.2

Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.

#### Examples

**Example**: Set a rectangle shape's fill transparency to 50% to make it semi-transparent

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    const shape = shapes.getItem(0);
    const shapeFill = shape.fill;
    
    // Set transparency to 50% (0.5)
    shapeFill.transparency = 0.5;
    
    await context.sync();
});
```

---

### type

**Type:** `Word.ShapeFillType | "NoFill" | "Solid" | "Gradient" | "Pattern" | "Picture" | "Texture" | "Mixed"`

**Since:** WordApiDesktop 1.2

Returns the fill type of the shape. See Word.ShapeFillType for details.

#### Examples

**Example**: Check the fill type of a shape and display it to the user

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();

    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shapeFill = shape.fill;
        shapeFill.load("type");
        await context.sync();

        console.log(`Shape fill type: ${shapeFill.type}`);
        // Output examples: "Solid", "Gradient", "NoFill", "Picture", etc.
    }
});
```

---

## Methods

### clear

**Kind:** `write`

Clears the fill formatting of this shape and set it to Word.ShapeFillType.NoFill;

#### Signature

**Returns:** `void`

#### Examples

**Example**: Clear the fill formatting from the first shape in the document and set it to no fill

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shapeFill = shape.fill;
        shapeFill.clear();
        await context.sync();
        
        console.log("Shape fill formatting cleared");
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
  - `options`: `Word.Interfaces.ShapeFillLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ShapeFill`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ShapeFill`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ShapeFill`

#### Examples

**Example**: Load and read the fill color of the first shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const firstShape = shapes.getFirst();
    const shapeFill = firstShape.fill;
    
    // Load the fill properties
    shapeFill.load("type, foregroundColor");
    
    // Sync to execute the load command
    await context.sync();
    
    // Read the loaded properties
    console.log("Fill type: " + shapeFill.type);
    console.log("Fill color: " + shapeFill.foregroundColor);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ShapeFillUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ShapeFill` (required)

  **Returns:** `void`

#### Examples

**Example**: Set the fill color and transparency of a shape to create a semi-transparent blue background

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const shapeFill = shape.fill;
    
    // Set multiple fill properties at once
    shapeFill.set({
        foreColor: "#4472C4",
        transparency: 0.5,
        visible: true
    });
    
    await context.sync();
});
```

---

### setSolidColor

**Kind:** `write`

Sets the fill formatting of the shape to a uniform color. This changes the fill type to Word.ShapeFillType.Solid.

#### Signature

**Parameters:**
- `color`: `string` (required)
  A string that represents the fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.

**Returns:** `void`

#### Examples

**Example**: Set the fill color of the first shape in the document to solid red

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shapeFill = shape.fill;
        shapeFill.setSolidColor("red");
        await context.sync();
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ShapeFill object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShapeFillData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ShapeFillData`

#### Examples

**Example**: Get the fill properties of a shape as a plain JavaScript object and log it to the console for inspection or serialization.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shapeFill = shape.fill;
        
        // Load fill properties
        shapeFill.load("type,foregroundColor,transparency");
        await context.sync();
        
        // Convert to plain JavaScript object
        const fillData = shapeFill.toJSON();
        
        // Log the plain object (useful for debugging or serialization)
        console.log("Shape fill data:", fillData);
        console.log("Fill type:", fillData.type);
        console.log("Foreground color:", fillData.foregroundColor);
        console.log("Transparency:", fillData.transparency);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ShapeFill`

#### Examples

**Example**: Track a shape's fill object to maintain its reference across multiple sync calls while changing its color properties in separate batches

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    context.load(shapes);
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0].convertToShape();
        const shapeFill = shape.fill;
        
        // Track the fill object to use it across multiple sync calls
        shapeFill.track();
        
        await context.sync();
        
        // First batch: set fill to solid
        shapeFill.setSolidColor("blue");
        await context.sync();
        
        // Second batch: change the color (object remains valid because it's tracked)
        shapeFill.setSolidColor("red");
        await context.sync();
        
        // Untrack when done
        shapeFill.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ShapeFill`

#### Examples

**Example**: Get a shape's fill properties, use them, then untrack the fill object to free memory resources.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const shapeFill = shape.fill;
    
    // Track the fill object to work with it
    shapeFill.load("type");
    await context.sync();
    
    // Use the fill properties
    console.log("Fill type: " + shapeFill.type);
    
    // Untrack the fill object to release memory
    shapeFill.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word/word.shapefill
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
