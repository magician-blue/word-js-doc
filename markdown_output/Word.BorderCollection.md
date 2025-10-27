# Word.BorderCollection

**Package:** `word`

**API Set:** WordApiDesktop 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the collection of border styles.

## Class Examples

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

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the border collection's request context to verify the connection to the Word host application before applying border styles to a paragraph.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.getBorders();
    
    // Load the borders collection
    borders.load();
    await context.sync();
    
    // Access the request context associated with the borders collection
    const borderContext = borders.context;
    
    // Verify the context is connected and use it for operations
    if (borderContext) {
        console.log("Border collection is connected to Word host application");
        
        // Now safely apply border styles using the same context
        borders.outsideBorderColor = "blue";
        borders.outsideBorderWidth = 2;
        
        await context.sync();
    }
});
```

---

### insideBorderColor

**Type:** `string`

**Since:** WordApiDesktop 1.1

Specifies the 24-bit color of the inside borders. Color is specified in '#RRGGBB' format or by using the color name.

#### Examples

**Example**: Set the inside borders of a table to red color

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set the inside border color to red
    table.borders.insideBorderColor = "#FF0000";
    
    await context.sync();
});
```

---

### insideBorderType

**Type:** `Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"`

**Since:** WordApiDesktop 1.1

Specifies the border type of the inside borders.

#### Examples

**Example**: Set the inside borders of a table to a dashed line style

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set the inside border type to dashed
    table.borders.insideBorderType = Word.BorderType.dashed;
    
    await context.sync();
    
    console.log("Inside borders set to dashed style");
});
```

---

### insideBorderWidth

**Type:** `Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"`

**Since:** WordApiDesktop 1.1

Specifies the width of the inside borders.

#### Examples

**Example**: Set the inside border width of a table to 2.25 points

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    table.getBorder(Word.BorderLocation.insideHorizontal).load("type");
    table.getBorder(Word.BorderLocation.insideVertical).load("type");
    
    await context.sync();
    
    // Set inside border width to 2.25 points
    table.borders.insideBorderWidth = Word.BorderWidth.pt225;
    
    await context.sync();
});
```

---

### items

**Type:** `Word.Border[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Iterate through all borders of a paragraph and log their types and colors to the console.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.borders;
    
    // Load the items property to access the border collection
    borders.load("items");
    await context.sync();
    
    // Access the loaded border items
    const borderItems = borders.items;
    
    // Iterate through each border in the collection
    for (let i = 0; i < borderItems.length; i++) {
        const border = borderItems[i];
        border.load("type, color");
    }
    
    await context.sync();
    
    // Log border information
    for (let i = 0; i < borderItems.length; i++) {
        console.log(`Border ${i}: Type = ${borderItems[i].type}, Color = ${borderItems[i].color}`);
    }
});
```

---

### outsideBorderColor

**Type:** `string`

**Since:** WordApiDesktop 1.1

Specifies the 24-bit color of the outside borders. Color is specified in '#RRGGBB' format or by using the color name.

#### Examples

**Example**: Update the outside border properties of a specified document style to have a dashed type, 0.25 point width, and green color.

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

---

### outsideBorderType

**Type:** `Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"`

**Since:** WordApiDesktop 1.1

Specifies the border type of the outside borders.

#### Examples

**Example**: Update the outside border properties of a specified document style to be dashed, 0.25 points wide, and green in color.

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

---

### outsideBorderWidth

**Type:** `Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"`

**Since:** WordApiDesktop 1.1

Specifies the width of the outside borders.

#### Examples

**Example**: Update the outside border properties of a specified document style to use a dashed border type with 0.25 point width and green color.

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

---

## Methods

### getByLocation

**Kind:** `read`

Gets the border that has the specified location.

#### Signature

**Parameters:**
- `borderLocation`: `Word.BorderLocation.top | Word.BorderLocation.left | Word.BorderLocation.bottom | Word.BorderLocation.right | Word.BorderLocation.insideHorizontal | Word.BorderLocation.insideVertical | "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical"` (required)

**Returns:** `Word.Border`

#### Examples

**Example**: Get the top border of the first paragraph and change its color to red

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const topBorder = paragraph.borders.getByLocation(Word.BorderLocation.top);
    
    topBorder.color = "#FF0000";
    topBorder.visible = true;
    
    await context.sync();
});
```

---

### getFirst

**Kind:** `read`

Gets the first border in this collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.Border`

#### Examples

**Example**: Get and apply a red color to the first border of the selected paragraph

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const borders = paragraph.borders;
    
    const firstBorder = borders.getFirst();
    firstBorder.color = "#FF0000";
    firstBorder.width = 2;
    
    await context.sync();
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first border in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Border`

#### Examples

**Example**: Check if a paragraph has any borders and display an alert with the first border's type, or indicate that no borders exist.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.getBorders();
    const firstBorder = borders.getFirstOrNullObject();
    
    firstBorder.load("isNullObject, type");
    await context.sync();
    
    if (firstBorder.isNullObject) {
        console.log("This paragraph has no borders.");
    } else {
        console.log(`First border type: ${firstBorder.type}`);
    }
});
```

---

### getItem

**Kind:** `read`

Gets a Border object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a Border object.

**Returns:** `Word.Border`

#### Examples

**Example**: Get the bottom border of the first paragraph and set its color to red and width to 3 points.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.borders;
    
    // Get the bottom border using index (0=top, 1=left, 2=bottom, 3=right)
    const bottomBorder = borders.getItem(2);
    bottomBorder.color = "#FF0000";
    bottomBorder.width = 3;
    
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
  - `options`: `Word.Interfaces.BorderCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.BorderCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BorderCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BorderCollection`

#### Examples

**Example**: Load and display the border properties of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border collection
    const borders = paragraph.borders;
    
    // Load border properties
    borders.load("items");
    
    // Sync to execute the load command
    await context.sync();
    
    // Access the loaded border properties
    console.log(`Number of borders: ${borders.items.length}`);
    borders.items.forEach((border, index) => {
        console.log(`Border ${index}: ${border.type}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BorderCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BorderCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.BorderCollectionData`

#### Examples

**Example**: Serialize the border collection of a paragraph to JSON format and log the border properties to the console for debugging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border collection
    const borders = paragraph.borders;
    
    // Load the border properties
    borders.load("items");
    
    await context.sync();
    
    // Convert the border collection to a plain JavaScript object
    const bordersJSON = borders.toJSON();
    
    // Log the JSON representation
    console.log("Border Collection JSON:", JSON.stringify(bordersJSON, null, 2));
    
    // Access the items array from the JSON object
    console.log("Number of borders:", bordersJSON.items.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BorderCollection`

#### Examples

**Example**: Track a paragraph's border collection across multiple sync calls to safely modify border properties without encountering InvalidObjectPath errors.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.getBorder();
    
    // Track the border collection for use across sync calls
    borders.track();
    
    await context.sync();
    
    // Now safe to modify borders in subsequent operations
    borders.outsideBorderColor = "#FF0000";
    borders.outsideBorderWidth = 2;
    
    await context.sync();
    
    // Untrack when done to free up memory
    borders.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BorderCollection`

#### Examples

**Example**: Get border collection from a paragraph, use it to check border properties, then untrack it to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the border collection and track it
    const borders = paragraph.getBorder();
    borders.load("type, color");
    
    await context.sync();
    
    // Use the border collection
    console.log("Border type:", borders.type);
    console.log("Border color:", borders.color);
    
    // Untrack the border collection to release memory
    borders.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
