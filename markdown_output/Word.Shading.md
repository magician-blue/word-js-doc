# Shading

**Package:** `word`

**API Set:** WordApi 1.6

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the shading object.

## Class Examples

```typescript
// Link to full sample: // Updates shading properties (e.g., texture, pattern colors) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update shading properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const shading: Word.Shading = style.shading;
    shading.load();
    await context.sync();

    shading.backgroundPatternColor = "blue";
    shading.foregroundPatternColor = "yellow";
    shading.texture = Word.ShadingTextureType.darkTrellis;

    console.log("Updated shading.");
  }
});
```

## Properties

### backgroundPatternColor

**Type:** `string`

**Since:** WordApi 1.6

Specifies the color for the background of the object. You can provide the value in the '#RRGGBB' format or the color name.

#### Examples

**Example**: Set the background color of the first paragraph's shading to light blue using a hex color code

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.shading;
    shading.backgroundPatternColor = "#ADD8E6";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Shading object to verify the connection between the add-in and Word before applying shading properties.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.getOrCreateShading();
    
    // Load the shading object
    shading.load("backgroundPatternColor");
    await context.sync();
    
    // Access the request context associated with the shading object
    const shadingContext = shading.context;
    
    // Verify the context is valid and connected
    if (shadingContext) {
        console.log("Shading context is connected to Word application");
        
        // Use the same context for additional operations
        shading.backgroundPatternColor = "yellow";
        await shadingContext.sync();
    }
});
```

---

### foregroundPatternColor

**Type:** `string`

**Since:** WordApiDesktop 1.1

Specifies the color for the foreground of the object. You can provide the value in the '#RRGGBB' format or the color name.

#### Examples

**Example**: Set the foreground pattern color of the selected paragraph's shading to light blue

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    paragraph.shading.foregroundPatternColor = "#87CEEB";
    
    await context.sync();
});
```

---

### texture

**Type:** `Word.ShadingTextureType | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid"`

**Since:** WordApiDesktop 1.1

Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see Add, change, or delete the background color in Word.

#### Examples

**Example**: Apply a light grid texture pattern to the shading of the first paragraph in the document.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.shadingOrNullObject;
    
    shading.texture = "LightGrid";
    
    await context.sync();
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
  - `options`: `Word.Interfaces.ShadingLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Shading`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Shading`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Shading`

#### Examples

**Example**: Get and display the background color of the shading applied to the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the shading object
    const shading = paragraph.shadingOrNullObject;
    
    // Load the backgroundPatternColor property
    shading.load("backgroundPatternColor");
    
    // Sync to execute the load command
    await context.sync();
    
    // Check if shading exists and display the color
    if (!shading.isNullObject) {
        console.log("Shading background color: " + shading.backgroundPatternColor);
    } else {
        console.log("No shading applied to the first paragraph");
    }
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ShadingUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Shading` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply yellow background shading and set the foreground color to red for the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const shading = firstParagraph.font.shading;
    
    shading.set({
        backgroundPatternColor: "yellow",
        foregroundPatternColor: "red"
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Shading object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShadingData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ShadingData`

#### Examples

**Example**: Get the shading properties of a paragraph as a JSON object and log it to the console

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.shading;
    
    // Load the shading properties
    shading.load("backgroundPatternColor");
    
    await context.sync();
    
    // Convert the shading object to a plain JavaScript object
    const shadingJSON = shading.toJSON();
    
    // Log the JSON representation
    console.log("Shading properties:", JSON.stringify(shadingJSON, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Shading`

#### Examples

**Example**: Apply shading to a paragraph, track the shading object to persist across multiple sync calls, and then modify its properties in a subsequent operation without getting an InvalidObjectPath error.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.shading;
    
    // Track the shading object for use across multiple sync calls
    shading.track();
    
    // First sync - set initial background color
    shading.backgroundPatternColor = "yellow";
    await context.sync();
    
    // Second sync - modify the shading again (tracking prevents InvalidObjectPath error)
    shading.backgroundPatternColor = "lightblue";
    await context.sync();
    
    // Untrack when done to free up memory
    shading.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Shading`

#### Examples

**Example**: Apply shading to a paragraph, then untrack the shading object to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.getShading();
    
    // Track the shading object to work with it
    shading.track();
    
    // Apply shading properties
    shading.backgroundPatternColor = "#FFFF00"; // Yellow background
    
    await context.sync();
    
    // Untrack the shading object to release memory
    shading.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.shading
