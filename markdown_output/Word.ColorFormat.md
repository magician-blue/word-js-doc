# Word.ColorFormat

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the color formatting of a shape or text in Word.

## Properties

### brightness

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the brightness of a specified shape color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral.

#### Examples

**Example**: Increase the brightness of the first shape's fill color to make it 50% lighter

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Access the shape's fill color format and set brightness to 0.5 (50% lighter)
        const colorFormat = shape.imageFormat.colorFormat;
        colorFormat.brightness = 0.5;
        
        await context.sync();
    }
});
```

---

### context

**Type:** `RequestContext`

**Since:** WordApi BETA (PREVIEW ONLY)

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ColorFormat object to verify the connection between the add-in and Word before applying color changes to a shape.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const colorFormat = shape.fill.foregroundColor;
        
        // Access the request context from the ColorFormat object
        const requestContext = colorFormat.context;
        
        // Verify the context is valid and connected
        console.log("Context is connected:", requestContext !== null);
        
        // Use the context to perform operations
        colorFormat.load("value");
        await requestContext.sync();
        
        console.log("Current color:", colorFormat.value);
    }
});
```

---

### objectThemeColor

**Type:** `Word.ThemeColorIndex | "NotThemeColor" | "MainDark1" | "MainLight1" | "MainDark2" | "MainLight2" | "Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6" | "Hyperlink" | "HyperlinkFollowed" | "Background1" | "Text1" | "Background2" | "Text2"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the theme color for a color format.

#### Examples

**Example**: Set the fill color of the first shape in the document to use the Accent1 theme color

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.load("foreColor");
        await context.sync();
        
        // Set the theme color to Accent1
        fill.foreColor.objectThemeColor = Word.ThemeColorIndex.accent1;
        
        await context.sync();
    }
});
```

---

### rgb

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the red-green-blue (RGB) value of the specified color. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the fill color of the first shape in the document to bright orange using RGB format

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fill = shape.fill;
        fill.setSolidColor("#FF6600");
        
        await context.sync();
    }
});
```

---

### tintAndShade

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the lightening or darkening of a specified shape's color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral.

#### Examples

**Example**: Lighten a shape's fill color by 40% to create a softer appearance

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Access the shape's fill color format and lighten it by 40%
        const colorFormat = shape.fill.foregroundColor;
        colorFormat.tintAndShade = 0.4;
        
        await context.sync();
        console.log("Shape color lightened by 40%");
    }
});
```

---

### type

**Type:** `Word.ColorType | "rgb" | "scheme"`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the shape color type.

#### Examples

**Example**: Check if a shape's fill color is defined using RGB values or a theme color scheme

```typescript
await Word.run(async (context) => {
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const fillColor = shape.fill.foregroundColor;
        fillColor.load("type");
        await context.sync();
        
        console.log(`Color type: ${fillColor.type}`);
        
        if (fillColor.type === Word.ColorType.rgb) {
            console.log("The shape uses RGB color values");
        } else if (fillColor.type === Word.ColorType.scheme) {
            console.log("The shape uses a theme color scheme");
        }
    }
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
  - `options`: `Word.Interfaces.ColorFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ColorFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ColorFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ColorFormat`

#### Examples

**Example**: Get and display the RGB color value of the first shape's fill color in the document

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.shapes;
    const firstShape = shapes.getFirst();
    
    // Get the fill color format
    const fillColor = firstShape.fill.foregroundColor;
    
    // Load the RGB property of the color format
    fillColor.load("rgb");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the RGB color value
    console.log("Shape fill color (RGB):", fillColor.rgb);
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ColorFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ColorFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Set the fill color of a shape to red using the set() method to configure multiple color properties at once

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Get the color format of the shape's fill
        const colorFormat = shape.fill.foregroundColor;
        
        // Use set() to configure color properties
        colorFormat.set({
            rgb: "#FF0000" // Set to red
        });
        
        await context.sync();
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ColorFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ColorFormatData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ColorFormatData`

#### Examples

**Example**: Serialize a shape's color format to JSON and log it to the console for debugging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        
        // Access the shape's fill color format
        const colorFormat = shape.imageFormat.colorFormat;
        colorFormat.load("*");
        await context.sync();
        
        // Convert the ColorFormat object to a plain JavaScript object
        const colorData = colorFormat.toJSON();
        
        // Log the serialized color data
        console.log("Color Format Data:", JSON.stringify(colorData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ColorFormat`

#### Examples

**Example**: Track a shape's color format object across multiple sync calls to safely modify its color properties without getting an "InvalidObjectPath" error.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const shape = shapes.items[0];
        const colorFormat = shape.fill.foregroundColor;
        
        // Track the color format object for use across sync calls
        colorFormat.track();
        
        await context.sync();
        
        // Now we can safely use the color format object after sync
        colorFormat.load("rgb");
        await context.sync();
        
        console.log("Current color:", colorFormat.rgb);
        
        // Modify the color
        colorFormat.rgb = "#FF0000"; // Set to red
        await context.sync();
        
        // Untrack when done
        colorFormat.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ColorFormat`

#### Examples

**Example**: Get the fill color of the first shape, use it for processing, then untrack the color format object to free memory.

```typescript
await Word.run(async (context) => {
    // Get the first shape in the document
    const shapes = context.document.body.inlinePictures;
    const shape = shapes.getFirstOrNullObject();
    
    // Load the shape and its fill color format
    shape.load("fill");
    await context.sync();
    
    if (!shape.isNullObject) {
        const colorFormat = shape.fill.foregroundColor;
        
        // Track the color format object for use
        colorFormat.track();
        colorFormat.load("rgb");
        await context.sync();
        
        // Use the color information
        console.log("Shape color RGB:", colorFormat.rgb);
        
        // Untrack the object to release memory after we're done
        colorFormat.untrack();
        await context.sync();
    }
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
