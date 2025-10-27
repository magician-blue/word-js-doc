# Word.Border

**Package:** `word`

**API Set:** WordApiDesktop 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the Border object for text, a paragraph, or a table.

## Class Examples

**Example**: Updates border properties (e.g., type, width, color) of the specified style.

```typescript
// Link to full sample: // Updates border properties (e.g., type, width, color) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update border properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const borders: Word.BorderCollection = style.borders;
    borders.load("items");
    await context.sync();

    borders.outsideBorderType = Word.BorderType.dashed;
    borders.outsideBorderWidth = Word.BorderWidth.pt025;
    borders.outsideBorderColor = "green";
    console.log("Updated outside borders.");
  }
});
```

## Properties

### color

**Type:** `string`

**Since:** WordApiDesktop 1.1

Specifies the color for the border. Color is specified in â#RRGGBBâ format or by using the color name.

#### Examples

**Example**: Set the border color of the first paragraph to red

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    border.color = "#FF0000";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the border's request context to verify the connection between the add-in and Word before applying border formatting

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Access the request context associated with the border object
    const borderContext = border.context;
    
    // Verify the context is valid before proceeding with operations
    if (borderContext) {
        border.type = Word.BorderType.single;
        border.color = "#0000FF";
        border.width = 2;
        
        await context.sync();
        console.log("Border formatting applied successfully");
    }
});
```

---

### location

**Type:** `Word.BorderLocation | "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"`

**Since:** WordApiDesktop 1.1

Gets the location of the border.

#### Examples

**Example**: Get the location of a paragraph's bottom border and display it in the console

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the bottom border of the paragraph
    const border = paragraph.getBorder(Word.BorderLocation.bottom);
    
    // Load the location property
    border.load("location");
    
    await context.sync();
    
    // Display the border location
    console.log("Border location: " + border.location);
});
```

---

### type

**Type:** `Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"`

**Since:** WordApiDesktop 1.1

Specifies the border type for the border.

#### Examples

**Example**: Set a paragraph's bottom border to a double-line style

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.bottom);
    
    border.type = Word.BorderType.double;
    border.load("type");
    
    await context.sync();
    console.log("Border type set to: " + border.type);
});
```

---

### visible

**Type:** `boolean`

**Since:** WordApiDesktop 1.1

Specifies whether the border is visible.

#### Examples

**Example**: Make the border of the first paragraph visible

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    border.visible = true;
    
    await context.sync();
});
```

---

### width

**Type:** `Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"`

**Since:** WordApiDesktop 1.1

Specifies the width for the border.

#### Examples

**Example**: Set a paragraph's bottom border width to 3.0 points

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.load("text");
    
    // Set the bottom border width to 3.0 points
    paragraph.border.bottom.width = Word.BorderWidth.pt300;
    
    await context.sync();
    console.log("Border width set to 3.0 points for paragraph: " + paragraph.text);
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
  - `options`: `Word.Interfaces.BorderLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Border`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Border`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Border`

#### Examples

**Example**: Get and display the border type and color of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border object
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Load border properties
    border.load("type, color");
    
    // Sync to read the properties
    await context.sync();
    
    // Display the border properties
    console.log("Border Type: " + border.type);
    console.log("Border Color: " + border.color);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.BorderUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Border` (required)

  **Returns:** `void`

#### Examples

**Example**: Apply multiple border properties to a paragraph's bottom border, setting its color to blue, line style to single, and width to 2 points.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.bottom);
    
    border.set({
        color: "#0000FF",
        type: Word.BorderType.single,
        width: 2
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Border object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BorderData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BorderData`

#### Examples

**Example**: Get a paragraph's border properties as a plain JavaScript object and log it to the console for debugging or serialization purposes.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border of the paragraph
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Load border properties
    border.load("type,color,width,visible");
    
    await context.sync();
    
    // Convert the border object to a plain JavaScript object
    const borderData = border.toJSON();
    
    // Log the plain object (useful for debugging or serialization)
    console.log("Border properties:", borderData);
    console.log("Border type:", borderData.type);
    console.log("Border color:", borderData.color);
    console.log("Border width:", borderData.width);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Border`

#### Examples

**Example**: Track a paragraph's bottom border object across multiple sync calls to read its properties and then modify its color without encountering InvalidObjectPath errors.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.bottom);
    
    // Track the border object for use across multiple sync calls
    border.track();
    
    // Load and sync to get current border properties
    border.load("type,color,width");
    await context.sync();
    
    console.log(`Current border - Type: ${border.type}, Color: ${border.color}, Width: ${border.width}`);
    
    // Modify the border in a subsequent operation
    border.color = "#FF0000";
    border.width = 2;
    await context.sync();
    
    // Clean up tracking when done
    border.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Border`

#### Examples

**Example**: Get a paragraph's border properties, then untrack the border object to free memory after use

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    const border = paragraph.getBorder(Word.BorderLocation.top);
    
    // Track the border object to work with it
    border.track();
    border.load("type,color,width");
    
    await context.sync();
    
    // Use the border properties
    console.log(`Border type: ${border.type}`);
    console.log(`Border color: ${border.color}`);
    console.log(`Border width: ${border.width}`);
    
    // Untrack the border object to release memory
    border.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.border
