# Word.TableStyle

**Package:** `word`

**API Set:** WordApi 1.6

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the TableStyle object.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-custom-style.yaml

// Gets the table style properties and displays them in the form.
const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
if (styleName == "") {
  console.warn("Please input a table style name.");
  return;
}

await Word.run(async (context) => {
  const tableStyle: Word.TableStyle =
    context.document.getStyles().getByName(styleName).tableStyle;
  tableStyle.load();
  await context.sync();

  if (tableStyle.isNullObject) {
    console.warn(`There's no existing table style with the name '${styleName}'.`);
    return;
  }

  console.log(tableStyle);
  (document.getElementById("alignment") as HTMLInputElement).value = tableStyle.alignment;
  (document.getElementById("allow-break-across-page") as HTMLInputElement).value =
    tableStyle.allowBreakAcrossPage.toString();
  (document.getElementById("top-cell-margin") as HTMLInputElement).value = tableStyle.topCellMargin;
  (document.getElementById("bottom-cell-margin") as HTMLInputElement).value = tableStyle.bottomCellMargin;
  (document.getElementById("left-cell-margin") as HTMLInputElement).value = tableStyle.leftCellMargin;
  (document.getElementById("right-cell-margin") as HTMLInputElement).value = tableStyle.rightCellMargin;
  (document.getElementById("cell-spacing") as HTMLInputElement).value = tableStyle.cellSpacing;
});
```

## Properties

### alignment

**Type:** `Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"`

**Since:** WordApiDesktop 1.1

Specifies the table's alignment against the page margin.

#### Examples

**Example**: Set a table's alignment to center it on the page

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set the table alignment to centered
    table.style.alignment = Word.Alignment.centered;
    
    await context.sync();
});
```

---

### allowBreakAcrossPage

**Type:** `boolean`

**Since:** WordApiDesktop 1.1

Specifies whether lines in tables formatted with a specified style break across pages.

#### Examples

**Example**: Configure a table style to prevent table rows from breaking across pages

```typescript
await Word.run(async (context) => {
    // Get the table style by name
    const tableStyle = context.document.getStyles().getByNameOrNullObject("MyCustomTableStyle") as Word.TableStyle;
    
    // Prevent rows from breaking across pages
    tableStyle.allowBreakAcrossPage = false;
    
    await context.sync();
    
    console.log("Table style configured to keep rows together on the same page");
});
```

---

### bottomCellMargin

**Type:** `number`

**Since:** WordApi 1.6

Specifies the amount of space to add between the contents and the bottom borders of the cells.

#### Examples

**Example**: Set the bottom cell margin to 10 points for a table style to add spacing between cell content and bottom borders

```typescript
await Word.run(async (context) => {
    // Get the table style by name
    const tableStyle = context.document.getStyles().getByNameOrNullObject("MyCustomTableStyle");
    
    // Set the bottom cell margin to 10 points
    tableStyle.bottomCellMargin = 10;
    
    await context.sync();
    
    console.log("Bottom cell margin set to 10 points");
});
```

---

### cellSpacing

**Type:** `number`

**Since:** WordApi 1.6

Specifies the spacing (in points) between the cells in a table style.

#### Examples

**Example**: Set the cell spacing to 5 points for a table style named "CustomTableStyle"

```typescript
await Word.run(async (context) => {
    const tableStyle = context.document.getStyles().getByNameOrNullObject("CustomTableStyle");
    tableStyle.load("type");
    
    await context.sync();
    
    if (!tableStyle.isNullObject && tableStyle.type === Word.StyleType.table) {
        tableStyle.cellSpacing = 5;
        await context.sync();
        
        console.log("Cell spacing set to 5 points");
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TableStyle object to verify the connection to the Word document before applying style changes.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableStyle = firstTable.style;
    
    // Load the style properties
    tableStyle.load("name");
    await context.sync();
    
    // Access the request context from the TableStyle object
    const styleContext = tableStyle.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (styleContext) {
        console.log("TableStyle is connected to the Office host application");
        console.log("Current table style name: " + tableStyle.name);
    }
    
    await context.sync();
});
```

---

### leftCellMargin

**Type:** `number`

**Since:** WordApi 1.6

Specifies the amount of space to add between the contents and the left borders of the cells.

#### Examples

**Example**: Set the left cell margin to 10 points for a table style to add spacing between cell content and the left border

```typescript
await Word.run(async (context) => {
    // Get the table style by name
    const tableStyle = context.document.getStyles().getByNameOrNullObject("MyCustomTableStyle");
    
    // Set the left cell margin to 10 points
    tableStyle.leftCellMargin = 10;
    
    await context.sync();
    
    console.log("Left cell margin set to 10 points");
});
```

---

### rightCellMargin

**Type:** `number`

**Since:** WordApi 1.6

Specifies the amount of space to add between the contents and the right borders of the cells.

#### Examples

**Example**: Set the right cell margin to 15 points for a table style to add spacing between cell content and right borders

```typescript
await Word.run(async (context) => {
    // Get the table style by name
    const tableStyle = context.document.getStyles().getByNameOrNullObject("MyCustomTableStyle");
    tableStyle.load("rightCellMargin");
    
    await context.sync();
    
    if (!tableStyle.isNullObject) {
        // Set the right cell margin to 15 points
        tableStyle.rightCellMargin = 15;
        
        await context.sync();
        console.log("Right cell margin set to 15 points");
    }
});
```

---

### topCellMargin

**Type:** `number`

**Since:** WordApi 1.6

Specifies the amount of space to add between the contents and the top borders of the cells.

#### Examples

**Example**: Set the top cell margin to 10 points for a table style to add spacing between cell content and top borders

```typescript
await Word.run(async (context) => {
    // Get the table style by name
    const tableStyle = context.document.getStyles().getByNameOrNullObject("MyCustomTableStyle");
    
    // Set the top cell margin to 10 points
    tableStyle.topCellMargin = 10;
    
    await context.sync();
    
    console.log("Top cell margin set to 10 points");
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
  - `options`: `Word.Interfaces.TableStyleLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TableStyle`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TableStyle`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TableStyle`

#### Examples

**Example**: Load and display the name and alignment properties of the first table's style in the document.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableStyle = firstTable.style;
    
    // Load specific properties of the table style
    tableStyle.load("name, alignment");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Table Style Name: " + tableStyle.name);
    console.log("Table Alignment: " + tableStyle.alignment);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.TableStyleUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.TableStyle` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a table style at once, including font name, font size, and alignment settings.

```typescript
await Word.run(async (context) => {
    // Get the table style by name
    const tableStyle = context.document.getStyles().getByNameOrNullObject("MyCustomTableStyle");
    
    // Set multiple properties at once using the set() method
    tableStyle.set({
        font: {
            name: "Arial",
            size: 11,
            bold: true
        },
        alignment: Word.Alignment.centered,
        allowBreakAcrossPage: false
    });
    
    await context.sync();
    console.log("Table style properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableStyle object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableStyleData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.TableStyleData`

#### Examples

**Example**: Retrieve a table style's properties as a plain JavaScript object and log it to the console for debugging or serialization purposes.

```typescript
await Word.run(async (context) => {
    // Get the first table style from the document
    const tableStyles = context.document.getStyles().tableStyles;
    const firstTableStyle = tableStyles.getFirst();
    
    // Load properties we want to serialize
    firstTableStyle.load("name,alignment,allowBreakAcrossPage");
    
    await context.sync();
    
    // Convert the TableStyle object to a plain JavaScript object
    const tableStyleData = firstTableStyle.toJSON();
    
    // Now we can use JSON.stringify or access properties as plain data
    console.log("Table Style as JSON:", JSON.stringify(tableStyleData, null, 2));
    console.log("Style name:", tableStyleData.name);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TableStyle`

#### Examples

**Example**: Get a table style by name, track it across multiple sync calls, and modify its properties while maintaining the object reference

```typescript
await Word.run(async (context) => {
    // Get a table style by name
    const tableStyle = context.document.getStyles().getByNameOrNullObject("Grid Table 1 Light");
    
    // Track the object to use it across multiple sync calls
    tableStyle.track();
    
    // Load properties
    tableStyle.load("name,font/bold");
    await context.sync();
    
    // Check if style exists and modify it
    if (!tableStyle.isNullObject) {
        tableStyle.font.bold = true;
        await context.sync();
        
        // Can safely use the object again after sync because it's tracked
        tableStyle.font.size = 12;
        await context.sync();
        
        console.log(`Modified table style: ${tableStyle.name}`);
    }
    
    // Untrack when done to release memory
    tableStyle.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TableStyle`

#### Examples

**Example**: Get a table style, apply it to a table, then untrack the style object to free memory after use

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get a built-in table style and track it
    const tableStyle = context.document.getStyles().getByNameOrNullObject("Grid Table 1 Light");
    tableStyle.load("name");
    
    await context.sync();
    
    if (!tableStyle.isNullObject) {
        // Apply the style to the table
        table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
        
        await context.sync();
        
        // Untrack the style object to release memory
        tableStyle.untrack();
        
        console.log("Table style applied and untracked");
    }
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-custom-style.yaml
