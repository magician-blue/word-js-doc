# ShadowFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the shadow formatting for a shape or text in Word.

## Properties

### blur

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the blur level for a shadow format as a value between 0.0 and 100.0.

#### Examples

**Example**: Set the shadow blur level to 50 for a selected shape to create a medium soft shadow effect

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shadow blur level to 50
        shape.shadow.blur = 50;
        
        await context.sync();
        console.log("Shadow blur set to 50");
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the shadow format's request context to verify the add-in is properly connected to the Word host application before applying shadow settings.

```typescript
await Word.run(async (context) => {
    const shape = context.document.body.insertShape("Rectangle", "Inline", 100, 100);
    const shadowFormat = shape.shadow;
    
    // Access the request context associated with the shadow format
    const shadowContext = shadowFormat.context;
    
    // Verify the context is connected to the Word host application
    if (shadowContext) {
        // Now safely apply shadow properties
        shadowFormat.blur = 10;
        shadowFormat.color = "#FF0000";
        shadowFormat.transparency = 0.5;
        
        await context.sync();
        console.log("Shadow format applied successfully through connected context");
    }
});
```

---

### foregroundColor

**Type:** `Word.ColorFormat`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a ColorFormat object that represents the foreground color for the fill, line, or shadow.

#### Examples

**Example**: Set the shadow's foreground color to red for a selected shape in the document.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Access the shadow format and set the foreground color to red
        const shadowFormat = shape.shadow;
        shadowFormat.foregroundColor.set("#FF0000");
        
        await context.sync();
        console.log("Shadow foreground color set to red");
    }
});
```

---

### isVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the object or the formatting applied to it is visible.

#### Examples

**Example**: Make a shape's shadow visible by setting the isVisible property to true

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shadowFormat = shape.shadowFormat;
        shadowFormat.isVisible = true;
        
        await context.sync();
    }
});
```

---

### obscured

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies true if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill, false if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.

#### Examples

**Example**: Make a shape's shadow appear filled in and obscured by the shape itself

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shadow to be obscured (filled in) by the shape
        shape.shadow.obscured = true;
        
        await context.sync();
    }
});
```

---

### offsetX

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the horizontal offset (in points) of the shadow from the shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.

#### Examples

**Example**: Set a shadow on a selected shape with a 10-point offset to the right

```typescript
await Word.run(async (context) => {
    const shapes = context.document.getSelection().inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shadowFormat = shape.shadowFormat;
        shadowFormat.offsetX = 10;
        
        await context.sync();
    }
});
```

---

### offsetY

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical offset (in points) of the shadow from the shape. A positive value offsets the shadow to the top of the shape; a negative value offsets it to the bottom.

#### Examples

**Example**: Set a shadow on a selected shape with a vertical offset of -5 points to position the shadow below the shape

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Access the shadow format and set vertical offset
        const shadowFormat = shape.shadowFormat;
        shadowFormat.offsetY = -5;
        
        await context.sync();
        console.log("Shadow vertical offset set to -5 points (below the shape)");
    }
});
```

---

### rotateWithShape

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether to rotate the shadow when rotating the shape.

#### Examples

**Example**: Configure a shape's shadow to rotate along with the shape when the shape is rotated

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.getShapes();
    const shape = shapes.getFirst();
    
    // Enable shadow rotation with shape
    shape.shadowFormat.rotateWithShape = true;
    
    await context.sync();
    console.log("Shadow will now rotate with the shape");
});
```

---

### size

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width of the shadow.

#### Examples

**Example**: Set the shadow width to 10 points for a selected shape in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Set the shadow size (width) to 10 points
        shape.shadow.size = 10;
        
        await context.sync();
    }
});
```

---

### style

**Type:** `Word.ShadowStyle | "Mixed" | "OuterShadow" | "InnerShadow"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the type of shadow formatting to apply to a shape.

#### Examples

**Example**: Apply an outer shadow effect to the first shape in the document

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        shape.shadow.style = Word.ShadowStyle.outerShadow;
        await context.sync();
    }
});
```

---

### transparency

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).

#### Examples

**Example**: Set a shape's shadow transparency to 50% (semi-transparent)

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    shape.shadow.transparency = 0.5;
    
    await context.sync();
});
```

---

### type

**Type:** `Word.ShadowType | "Mixed" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9" | "Type10" | "Type11" | "Type12" | "Type13" | "Type14" | "Type15" | "Type16" | "Type17" | "Type18" | "Type19" | "Type20" | "Type21" | "Type22" | "Type23" | "Type24" | "Type25" | "Type26" | "Type27" | "Type28" | "Type29" | "Type30" | "Type31" | "Type32" | "Type33" | "Type34" | "Type35" | "Type36" | "Type37" | "Type38" | "Type39" | "Type40" | "Type41" | "Type42" | "Type43"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the shape shadow type.

#### Examples

**Example**: Set the shadow type of the first shape in the document to "Type2" (outer shadow)

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const firstShape = shapes.getFirst();
    
    // Set the shadow type to Type2
    firstShape.shadowFormat.type = "Type2";
    
    await context.sync();
    
    console.log("Shadow type set to Type2");
});
```

---

## Methods

### incrementOffsetX

**Kind:** `write`

Changes the horizontal offset of the shadow by the number of points.

#### Signature

**Parameters:**
- `increment`: `number` (required)
  The number of points to adjust.

**Returns:** `void`

#### Examples

**Example**: Increase the horizontal shadow offset of the first shape in the document by 5 points to shift the shadow further to the right

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shadowFormat = shape.shadowFormat;
        shadowFormat.incrementOffsetX(5);
        await context.sync();
    }
});
```

---

### incrementOffsetY

**Kind:** `write`

Changes the vertical offset of the shadow by the specified number of points.

#### Signature

**Parameters:**
- `increment`: `number` (required)
  The number of points to adjust.

**Returns:** `void`

#### Examples

**Example**: Increase the vertical offset of a shadow on the first shape in the document by 5 points to make it appear further below the shape.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shadowFormat = shape.shadow;
        shadowFormat.incrementOffsetY(5);
        await context.sync();
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
  - `options`: `Word.Interfaces.ShadowFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ShadowFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ShadowFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ShadowFormat`

#### Examples

**Example**: Load and read the shadow properties of the first shape in the document to check if it has a shadow enabled.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    const firstShape = shapes.getFirstOrNullObject();
    
    // Get the shadow format
    const shadowFormat = firstShape.shadowFormat;
    
    // Load shadow properties
    shadowFormat.load("enabled, color, transparency, blur");
    
    await context.sync();
    
    // Read the loaded properties
    if (!firstShape.isNullObject) {
        console.log("Shadow enabled:", shadowFormat.enabled);
        console.log("Shadow color:", shadowFormat.color);
        console.log("Shadow transparency:", shadowFormat.transparency);
        console.log("Shadow blur:", shadowFormat.blur);
    }
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ShadowFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ShadowFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply shadow formatting to a shape by setting multiple shadow properties at once, including blur, color, and offset values.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    const shape = shapes.items[0];
    const shadowFormat = shape.shadow;
    
    // Set multiple shadow properties at once
    shadowFormat.set({
        blur: 10,
        color: "#FF0000",
        offsetX: 5,
        offsetY: 5,
        transparency: 0.5
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ShadowFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShadowFormatData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ShadowFormatData`

#### Examples

**Example**: Get the shadow formatting properties of the first shape in the document as a plain JavaScript object and log it to the console.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shadowFormat = shape.shadow;
        
        // Load shadow properties
        shadowFormat.load("blur,color,offsetX,offsetY,transparency,type");
        await context.sync();
        
        // Convert to plain JavaScript object
        const shadowData = shadowFormat.toJSON();
        
        // Log the shadow properties
        console.log("Shadow Format Data:", shadowData);
        console.log("Shadow Type:", shadowData.type);
        console.log("Shadow Color:", shadowData.color);
        console.log("Shadow Blur:", shadowData.blur);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ShadowFormat`

#### Examples

**Example**: Apply shadow formatting to a shape and track it across multiple sync calls to ensure the shadow properties persist when making subsequent changes to the shape.

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shadowFormat = shape.shadowFormat;
        
        // Track the shadow format object for use across multiple sync calls
        shadowFormat.track();
        
        // First sync: Load current properties
        shadowFormat.load("blur,color");
        await context.sync();
        
        console.log("Current shadow blur: " + shadowFormat.blur);
        
        // Second sync: Modify shadow properties
        shadowFormat.blur = 10;
        shadowFormat.color = "#FF0000";
        await context.sync();
        
        // Untrack when done
        shadowFormat.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ShadowFormat`

#### Examples

**Example**: Apply shadow formatting to a shape, then untrack the shadow format object to free up memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const shadowFormat = shape.shadow;
        
        // Track the shadow format object for modifications
        shadowFormat.track();
        
        // Configure shadow properties
        shadowFormat.blur = 10;
        shadowFormat.color = "#FF0000";
        shadowFormat.transparency = 0.5;
        
        await context.sync();
        
        // Untrack the shadow format object to release memory
        shadowFormat.untrack();
        await context.sync();
    }
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
