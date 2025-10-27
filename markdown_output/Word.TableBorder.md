# Word.TableBorder

**Package:** `word`

**API Set:** WordApi 1.3 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Specifies the border style.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
});
```

## Properties

### color

**Type:** `string`

**Since:** WordApi 1.3

Specifies the table border color.

#### Examples

**Example**: Retrieve and display the color, type, and width properties of the top border of the first table in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a table border object to verify the connection to the Office host application

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const topBorder = table.getBorder(Word.BorderLocation.top);
    
    // Load the border properties
    topBorder.load("type,color");
    await context.sync();
    
    // Access the request context associated with the border object
    const borderContext = topBorder.context;
    
    // Verify the context is valid and connected
    console.log("Border context is connected:", borderContext !== null);
    console.log("Border type:", topBorder.type);
    console.log("Border color:", topBorder.color);
});
```

---

### type

**Type:** `Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"`

**Since:** WordApi 1.3

Specifies the type of the table border.

#### Examples

**Example**: Retrieve and display the type, color, and width properties of the top border of the first table in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
});
```

---

### width

**Type:** `number`

**Since:** WordApi 1.3

Specifies the width, in points, of the table border. Not applicable to table border types that have fixed widths.

#### Examples

**Example**: Retrieve and display the type, color, and width properties of the top border of the first table in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
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
  - `options`: `Word.Interfaces.TableBorderLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TableBorder`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TableBorder`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TableBorder`

#### Examples

**Example**: Load and read the border width property of the first table's top border

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the top border of the table
    const topBorder = table.getBorder(Word.BorderLocation.top);
    
    // Load the width property of the border
    topBorder.load("width");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded property
    console.log(`Top border width: ${topBorder.width} pt`);
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.TableBorderUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.TableBorder` (required)

  **Returns:** `void`

#### Examples

**Example**: Set multiple border properties (color, width, and type) on the top border of the first table in the document

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const topBorder = firstTable.getBorder(Word.BorderLocation.top);
    
    topBorder.set({
        color: "#FF0000",
        width: 3,
        type: Word.BorderType.single
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableBorder object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableBorderData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.TableBorderData`

#### Examples

**Example**: Get the border properties of a table's first cell as a JSON object and log it to the console for inspection or serialization.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first cell's top border
    const topBorder = table.tables.getFirst().rows.getFirst().cells.getFirst().getBorder(Word.BorderLocation.top);
    
    // Load the border properties
    topBorder.load("type,color,width");
    
    await context.sync();
    
    // Convert the TableBorder object to a plain JSON object
    const borderJson = topBorder.toJSON();
    
    // Log the JSON representation
    console.log("Border properties:", borderJson);
    // Output example: { type: "Single", color: "#000000", width: 1 }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TableBorder`

#### Examples

**Example**: Get a table border, track it across multiple sync calls, and modify its properties while maintaining the object reference outside the sequential execution batch.

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const topBorder = firstTable.getBorder(Word.BorderLocation.top);
    
    // Track the border object to use it across multiple sync calls
    topBorder.track();
    
    topBorder.load("type,color,width");
    await context.sync();
    
    console.log(`Current border - Type: ${topBorder.type}, Color: ${topBorder.color}`);
    
    // Modify the border in a subsequent sync
    topBorder.type = Word.BorderType.single;
    topBorder.color = "#FF0000";
    topBorder.width = 3;
    await context.sync();
    
    console.log("Border updated successfully");
    
    // Untrack when done to free memory
    topBorder.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TableBorder`

#### Examples

**Example**: Apply a border style to a table cell, then untrack the border object to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const cell = table.getCell(0, 0);
    const border = cell.getBorder(Word.BorderLocation.top);
    
    // Track the border object for changes
    border.track();
    
    // Modify the border properties
    border.type = Word.BorderType.single;
    border.color = "#FF0000";
    border.width = 2;
    
    // Sync changes
    await context.sync();
    
    // Untrack the border object to release memory
    border.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext
- https://learn.microsoft.com/en-us/javascript/api/word/word.bordertype
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tableborderloadoptions
- https://learn.microsoft.com/en-us/javascript/api/word/word.tableborder
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tableborderupdatedata
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tableborderdata
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
