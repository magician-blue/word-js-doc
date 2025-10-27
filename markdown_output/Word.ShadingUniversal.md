# ShadingUniversal

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the ShadingUniversal object, which manages shading for a range, paragraph, frame, or table.

## Properties

### backgroundPatternColor

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the color that's applied to the background of the ShadingUniversal object. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the background pattern color of the first paragraph to light blue (#ADD8E6)

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.shading.backgroundPatternColor = "#ADD8E6";
    
    await context.sync();
});
```

---

### backgroundPatternColorIndex

**Type:** `Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the color that's applied to the background of the ShadingUniversal object.

#### Examples

**Example**: Set the background pattern color of the first paragraph to bright green using the color index

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.shading.backgroundPatternColorIndex = "BrightGreen";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ShadingUniversal object to synchronize changes with the Office host application

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the shading object for the paragraph
    const shading = paragraph.shading;
    
    // Access the request context from the shading object
    const shadingContext = shading.context;
    
    // Use the context to load properties and sync
    shading.load("backgroundPatternColor");
    await shadingContext.sync();
    
    console.log("Shading background color: " + shading.backgroundPatternColor);
});
```

---

### foregroundPatternColor

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern. You can provide the value in the '#RRGGBB' format.

#### Examples

**Example**: Set the foreground pattern color of paragraph shading to red

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.shading.foregroundPatternColor = "#FF0000";
    
    await context.sync();
});
```

---

### foregroundPatternColorIndex

**Type:** `Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern.

#### Examples

**Example**: Set the foreground pattern color of paragraph shading to red

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.shading.foregroundPatternColorIndex = "Red";
    
    await context.sync();
});
```

---

### texture

**Type:** `Word.ShadingTextureType | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see [Add, change, or delete the background color in Word](https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515).

#### Examples

**Example**: Apply a light diagonal down texture pattern to the shading of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const shading = firstParagraph.shadingOrNullObject;
    
    shading.texture = "LightDiagonalDown";
    
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
  - `options`: `Word.Interfaces.ShadingUniversalLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ShadingUniversal`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ShadingUniversal`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ShadingUniversal`

#### Examples

**Example**: Load and display the background color of the first paragraph's shading

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.shadingUniversal;
    
    // Load the background color property
    shading.load("backgroundPatternColor");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the loaded property
    console.log("Background color: " + shading.backgroundPatternColor);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ShadingUniversalUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ShadingUniversal` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply yellow background shading with 25% texture pattern to the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const shading = firstParagraph.shadingOrNullObject;
    
    shading.set({
        backgroundPatternColor: "yellow",
        texture: Word.ShadingTextureType.percent25
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ShadingUniversal object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShadingUniversalData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ShadingUniversalData`

#### Examples

**Example**: Get the shading properties of the first paragraph as a plain JavaScript object and log it to the console

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.shading;
    
    // Load the shading properties
    shading.load("backgroundPatternColor");
    
    await context.sync();
    
    // Convert the shading object to a plain JavaScript object
    const shadingData = shading.toJSON();
    
    // Log the plain object (useful for debugging or serialization)
    console.log("Shading properties:", shadingData);
    console.log("Background color:", shadingData.backgroundPatternColor);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ShadingUniversal`

#### Examples

**Example**: Apply background color shading to a paragraph and track the shading object to maintain reference across multiple sync calls for later modification.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    const shading = paragraph.shading;
    
    // Track the shading object for use across sync calls
    shading.track();
    
    // Load and sync to get current properties
    shading.load("backgroundPatternColor");
    await context.sync();
    
    // Set initial background color
    shading.backgroundPatternColor = "yellow";
    await context.sync();
    
    // Later in the same run, modify the shading again
    // The tracked object remains valid across sync calls
    shading.backgroundPatternColor = "lightblue";
    await context.sync();
    
    // Untrack when done to free up memory
    shading.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ShadingUniversal`

#### Examples

**Example**: Apply shading to a paragraph, then untrack the shading object to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the shading object and track it
    const shading = paragraph.shading;
    shading.track();
    
    // Apply shading properties
    shading.backgroundPatternColor = "#FFFF00"; // Yellow background
    
    // Sync to apply changes
    await context.sync();
    
    // Untrack the shading object to release memory
    shading.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word/word.shadinguniversal
