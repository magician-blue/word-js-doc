# Word.TableCellCollection

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `officeextension.clientobject`

## Description

Contains the collection of the document's TableCell objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets content alignment details about the first cell of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  const firstCell: Word.TableCell = firstTableRow.cells.getFirst();
  firstCell.load(["horizontalAlignment", "verticalAlignment"]);
  await context.sync();

  console.log(
    `Details about the alignment of the first table's first cell:`,
    `- Horizontal alignment of the cell's content: ${firstCell.horizontalAlignment}`,
    `- Vertical alignment of the cell's content: ${firstCell.verticalAlignment}`
  );
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a table cell collection to verify the connection between the add-in and Word before performing operations on table cells.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableCells = firstTable.getRange().cells;
    
    // Access the request context from the table cell collection
    const cellContext = tableCells.context;
    
    // Verify the context is valid by using it to load properties
    tableCells.load("items");
    await cellContext.sync();
    
    console.log(`Successfully accessed context. Found ${tableCells.items.length} cells in the table.`);
});
```

---

### items

**Type:** `Word.TableCell[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all loaded table cells from the first table and highlight cells in the first row with yellow background color.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableCells = firstTable.getRange().getRange("Whole").cells;
    
    // Load the items property to access the collection
    tableCells.load("items");
    await context.sync();
    
    // Access the loaded items array
    const cellItems = tableCells.items;
    
    // Highlight the first row cells (assuming table has cells)
    const cellsPerRow = firstTable.rowCount > 0 ? cellItems.length / firstTable.rowCount : 0;
    
    for (let i = 0; i < Math.min(cellsPerRow, cellItems.length); i++) {
        cellItems[i].body.font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first table cell in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.TableCell`

#### Examples

**Example**: Get the first cell from a table and highlight it with yellow shading to mark it as a header cell.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell from the table's cell collection
    const firstCell = firstTable.tableRows.getFirst().cells.getFirst();
    
    // Highlight the first cell with yellow shading
    firstCell.shadingColor = "#FFFF00";
    
    await context.sync();
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first table cell in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

#### Signature

**Returns:** `Word.TableCell`

#### Examples

**Example**: Check if a table has any cells and highlight the first cell yellow if it exists, otherwise log that the table is empty.

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const tableCells = firstTable.tableRows.getFirst().cells;
    const firstCell = tableCells.getFirstOrNullObject();
    
    firstCell.load("isNullObject");
    await context.sync();
    
    if (firstCell.isNullObject) {
        console.log("The table has no cells.");
    } else {
        firstCell.body.font.highlightColor = "yellow";
        await context.sync();
        console.log("First cell highlighted.");
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TableCellCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TableCellCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TableCellCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TableCellCollection`

#### Examples

**Example**: Load and display the text content of all cells in the first table of the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get all cells in the table
    const tableCells = firstTable.getRange().cells;
    
    // Load the text property of all cells
    tableCells.load("text");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the text content of each cell
    for (let i = 0; i < tableCells.items.length; i++) {
        console.log(`Cell ${i}: ${tableCells.items[i].value}`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.TableCellCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableCellCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.TableCellCollectionData`

#### Examples

**Example**: Export table cell data to JSON format for logging or external processing by serializing the first table's cells.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get all cells from the table
    const tableCells = firstTable.getRange().cells;
    
    // Load properties we want to serialize
    tableCells.load("items/value, items/rowIndex, items/columnIndex");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const cellsData = tableCells.toJSON();
    
    // Log the JSON representation
    console.log("Table cells data:", JSON.stringify(cellsData, null, 2));
    
    // Access the items array
    console.log(`Total cells: ${cellsData.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TableCellCollection`

#### Examples

**Example**: Track table cells in the first table to maintain references across multiple sync calls while modifying cell properties

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableCells = firstTable.tableRows.getFirst().cells;
    
    // Track the cell collection to maintain references across sync calls
    tableCells.track();
    
    // Load cell properties
    tableCells.load("items");
    await context.sync();
    
    // First operation: set background color
    for (let i = 0; i < tableCells.items.length; i++) {
        tableCells.items[i].body.font.color = "blue";
    }
    await context.sync();
    
    // Second operation: modify text (cells remain valid due to tracking)
    for (let i = 0; i < tableCells.items.length; i++) {
        tableCells.items[i].body.insertText(`Cell ${i + 1}`, Word.InsertLocation.replace);
    }
    await context.sync();
    
    // Untrack when done to release memory
    tableCells.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.TableCellCollection`

#### Examples

**Example**: Process table cells to find cells with specific content, then release the tracked TableCellCollection from memory to improve performance.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableCells = firstTable.getRange().cells;
    
    // Load the cell values
    tableCells.load("items/value");
    
    // Track the collection for processing
    tableCells.track();
    
    await context.sync();
    
    // Process the cells (e.g., count cells with specific content)
    let count = 0;
    for (let i = 0; i < tableCells.items.length; i++) {
        if (tableCells.items[i].value.includes("Important")) {
            count++;
        }
    }
    
    console.log(`Found ${count} cells with 'Important'`);
    
    // Release the memory associated with the tracked collection
    tableCells.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.tablecellcollection
