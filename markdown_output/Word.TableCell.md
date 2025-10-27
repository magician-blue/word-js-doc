# Word.TableCell

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a table cell in a Word document.

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

### body

**Type:** `Word.Body`

**Since:** 1.3

Gets the body object of the cell.

#### Examples

**Example**: Add formatted text content to the first cell of the first table in the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first cell (row 0, column 0)
    const cell = table.getCell(0, 0);
    
    // Access the cell's body and insert text
    cell.body.insertText("Product Name", Word.InsertLocation.start);
    cell.body.font.bold = true;
    cell.body.font.color = "#0066CC";
    
    await context.sync();
});
```

---

### cellIndex

**Type:** `number`

**Since:** 1.3

Gets the index of the cell in its row.

#### Examples

**Example**: Highlight the first cell in each table row by setting its background color to yellow, using the cellIndex property to identify first cells.

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const tables = body.tables;
    tables.load("items");
    
    await context.sync();
    
    for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        const cells = table.getRange().cells;
        cells.load("items");
        
        await context.sync();
        
        for (let j = 0; j < cells.items.length; j++) {
            const cell = cells.items[j];
            cell.load("cellIndex");
            
            await context.sync();
            
            if (cell.cellIndex === 0) {
                cell.shadingColor = "yellow";
            }
        }
    }
    
    await context.sync();
});
```

---

### columnWidth

**Type:** `number`

**Since:** 1.3

Specifies the width of the cell's column in points. This is applicable to uniform tables.

#### Examples

**Example**: Set the column width of the first cell in the first table to 100 points

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const firstCell = firstTable.tables.getFirst().getCell(0, 0);
    
    firstCell.columnWidth = 100;
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the table cell's request context to load and read the cell's text content

```typescript
await Word.run(async (context) => {
    // Get the first table cell in the document
    const tableCell = context.document.body.tables.getFirst().getCell(0, 0);
    
    // Access the cell's context to load properties
    tableCell.context.load(tableCell, 'value');
    
    await tableCell.context.sync();
    
    console.log("Cell content: " + tableCell.value);
});
```

---

### horizontalAlignment

**Type:** `Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"`

**Since:** 1.3

Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

#### Examples

**Example**: Retrieve and display the horizontal and vertical alignment settings of the first cell in the first table of the document.

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

---

### parentRow

**Type:** `Word.TableRow`

**Since:** 1.3

Gets the parent row of the cell.

#### Examples

**Example**: Highlight the entire row that contains a specific cell by setting its shading color to light yellow

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get a specific cell (e.g., first cell in second row)
    const cell = table.getCell(1, 0);
    
    // Get the parent row of the cell
    const parentRow = cell.parentRow;
    
    // Set shading color for the entire parent row
    parentRow.shadingColor = "#FFFFE0"; // Light yellow
    
    await context.sync();
});
```

---

### parentTable

**Type:** `Word.Table`

**Since:** 1.3

Gets the parent table of the cell.

#### Examples

**Example**: Get the parent table of a cell and apply a blue border to the entire table

```typescript
await Word.run(async (context) => {
    // Get the first table cell in the document
    const cell = context.document.body.tables.getFirst().getCell(0, 0);
    
    // Get the parent table of the cell
    const parentTable = cell.parentTable;
    
    // Apply a blue border to the parent table
    parentTable.set({
        styleBuiltIn: Word.Style.gridTable1Light,
        styleBandedColumns: false,
        styleBandedRows: true
    });
    
    parentTable.getBorder(Word.BorderLocation.all).set({
        type: Word.BorderType.single,
        color: "0000FF",
        width: 2
    });
    
    await context.sync();
});
```

---

### rowIndex

**Type:** `number`

**Since:** 1.3

Gets the index of the cell's row in the table.

#### Examples

**Example**: Get the row index of a clicked table cell and display it in the cell's text

```typescript
await Word.run(async (context) => {
    // Get the first table cell in the document
    const tableCell = context.document.body.tables.getFirst().getCell(2, 1);
    
    // Load the rowIndex property
    tableCell.load("rowIndex");
    
    await context.sync();
    
    // Display the row index in the cell
    tableCell.body.insertText(`Row Index: ${tableCell.rowIndex}`, Word.InsertLocation.replace);
    
    await context.sync();
});
```

---

### shadingColor

**Type:** `string`

**Since:** 1.3

Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.

#### Examples

**Example**: Set the shading color of the first cell in the first table to light blue

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const firstCell = firstTable.getCell(0, 0);
    
    firstCell.shadingColor = "#ADD8E6";
    
    await context.sync();
});
```

---

### value

**Type:** `string`

**Since:** 1.3

Specifies the text of the cell.

#### Examples

**Example**: Set the text content of the first cell in the first table to "Product Name"

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const firstCell = firstTable.getCell(0, 0);
    
    firstCell.value = "Product Name";
    
    await context.sync();
});
```

---

### verticalAlignment

**Type:** `Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom"`

**Since:** 1.3

Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.

#### Examples

**Example**: Retrieve and display the horizontal and vertical alignment settings of the first cell in the first table of the document.

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

---

### width

**Type:** `number`

**Since:** 1.3

Gets the width of the cell in points.

#### Examples

**Example**: Get the width of the first cell in the first table and display it in the console

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell from the first row
    const firstCell = firstTable.rows.getFirst().cells.getFirst();
    
    // Load the width property
    firstCell.load("width");
    
    await context.sync();
    
    // Display the cell width in points
    console.log(`Cell width: ${firstCell.width} points`);
});
```

---

## Methods

### deleteColumn

**Kind:** `delete`

Deletes the column containing this cell. This is applicable to uniform tables.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete the second column from the first table in the document by selecting a cell in that column

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get a cell from the second column (row 0, column 1)
    const cell = table.getCell(0, 1);
    
    // Delete the column containing this cell
    cell.deleteColumn();
    
    await context.sync();
});
```

---

### deleteRow

**Kind:** `delete`

Deletes the row containing this cell.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete the first row of the first table in the document by accessing a cell in that row

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell of the first row
    const firstCell = firstTable.getCell(0, 0);
    
    // Delete the row containing this cell
    firstCell.deleteRow();
    
    await context.sync();
    
    console.log("First row deleted successfully");
});
```

---

### getBorder

**Kind:** `read`

Gets the border style for the specified border.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `borderLocation`: `Word.BorderLocation` (required)
    The border location.

  **Returns:** `Word.TableBorder`

**Overload 2:**

  **Parameters:**
  - `borderLocation`: `"Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"` (required)
    The border location.

  **Returns:** `Word.TableBorder`

#### Examples

**Example**: Retrieve and display the type, color, and width properties of the left border of the first cell in the first table of the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstCell: Word.TableCell = firstTable.getCell(0, 0);
  const borderLocation = "Left";
  const border: Word.TableBorder = firstCell.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(
    `Details about the ${borderLocation} border of the first table's first cell:`,
    `- Color: ${border.color}`,
    `- Type: ${border.type}`,
    `- Width: ${border.width} points`
  );
});
```

---

### getCellPadding

**Kind:** `read`

Gets cell padding in points.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `cellPaddingLocation`: `Word.CellPaddingLocation` (required)
    The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.

  **Returns:** `OfficeExtension.ClientResult<number>`

**Overload 2:**

  **Parameters:**
  - `cellPaddingLocation`: `"Top" | "Left" | "Bottom" | "Right"` (required)
    The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.

  **Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Retrieve the left border cell padding value in points from the first cell of the first table in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets cell padding details about the first cell of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstCell: Word.TableCell = firstTable.getCell(0, 0);
  const cellPaddingLocation = "Left";
  const cellPadding = firstCell.getCellPadding(cellPaddingLocation);
  await context.sync();

  console.log(
    `Cell padding details about the ${cellPaddingLocation} border of the first table's first cell: ${cellPadding.value} points`
  );
});
```

---

### getNext

**Kind:** `read`

Gets the next cell. Throws an `ItemNotFound` error if this cell is the last one.

#### Signature

**Returns:** `Word.TableCell`

#### Examples

**Example**: Highlight the next cell after the first cell in the first table by setting its shading color to yellow.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell in the table
    const firstCell = firstTable.getCell(0, 0);
    
    // Get the next cell after the first cell
    const nextCell = firstCell.getNext();
    
    // Set the shading color of the next cell to yellow
    nextCell.shadingColor = "#FFFF00";
    
    await context.sync();
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next cell. If this cell is the last one, returns an object with `isNullObject` set to `true`. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.TableCell`

#### Examples

**Example**: Iterate through all cells in the first row of a table and highlight every other cell

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const firstRow = table.rows.getFirst();
    
    // Start with the first cell
    let cell = firstRow.cells.getFirst();
    cell.load("cellIndex");
    await context.sync();
    
    // Iterate through cells using getNextOrNullObject
    let isAlternate = false;
    while (cell) {
        if (isAlternate) {
            cell.body.font.highlightColor = "yellow";
        }
        isAlternate = !isAlternate;
        
        // Get next cell or null
        const nextCell = cell.getNextOrNullObject();
        nextCell.load("isNullObject, cellIndex");
        await context.sync();
        
        // Break if we've reached the end
        if (nextCell.isNullObject) {
            break;
        }
        
        cell = nextCell;
    }
    
    await context.sync();
});
```

---

### insertColumns

**Kind:** `create`

Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.

#### Signature

**Parameters:**
- `insertLocation`: `Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"` (required)
  It must be 'Before' or 'After'.
- `columnCount`: `number` (required)
  Number of columns to add.
- `values`: `string[][]` (optional)
  Optional 2D array. Cells are filled if the corresponding strings are specified in the array.

**Returns:** `void`

#### Examples

**Example**: Add 2 columns to the right of the first cell in the first table and populate them with header values

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell in the table
    const firstCell = firstTable.getCell(0, 0);
    
    // Insert 2 columns to the right of this cell with values
    firstCell.insertColumns(
        Word.InsertLocation.after,
        2,
        [["Column A", "Column B"]]
    );
    
    await context.sync();
});
```

---

### insertRows

**Kind:** `create`

Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.

#### Signature

**Parameters:**
- `insertLocation`: `Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"` (required)
  It must be 'Before' or 'After'.
- `rowCount`: `number` (required)
  Number of rows to add.
- `values`: `string[][]` (optional)
  Optional 2D array. Cells are filled if the corresponding strings are specified in the array.

**Returns:** `Word.TableRowCollection`

#### Examples

**Example**: Insert 2 new rows below the first cell of the first table and populate them with employee data

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell of the table
    const firstCell = firstTable.tables.getFirst().getCell(0, 0);
    
    // Insert 2 rows below this cell with employee data
    const newRows = firstCell.insertRows(
        Word.InsertLocation.after,
        2,
        [
            ["John Doe", "Sales", "50000"],
            ["Jane Smith", "Marketing", "55000"]
        ]
    );
    
    await context.sync();
    
    console.log("Successfully inserted 2 rows with employee data");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. Call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TableCellLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TableCell`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TableCell`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TableCell`

#### Examples

**Example**: Load and display the text content and width of the first cell in the first table

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell from the table
    const firstCell = firstTable.getCell(0, 0);
    
    // Load specific properties of the cell
    firstCell.load("value, width");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Cell text:", firstCell.value);
    console.log("Cell width:", firstCell.width);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. Accepts either a plain object with the appropriate properties or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.TableCellUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.TableCell` (required)

  **Returns:** `void`

#### Examples

**Example**: Format a table cell by setting multiple properties at once, including background color, vertical alignment, and cell padding

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first cell in the first row
    const cell = table.rows.getFirst().cells.getFirst();
    
    // Set multiple properties at once using the set() method
    cell.set({
        shadingColor: "#FFFF00",  // Yellow background
        verticalAlignment: Word.VerticalAlignment.center,
        width: 100
    });
    
    await context.sync();
    
    console.log("Table cell properties have been set");
});
```

---

### setCellPadding

**Kind:** `write`

Sets cell padding in points.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `cellPaddingLocation`: `Word.CellPaddingLocation` (required)
    The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
  - `cellPadding`: `number` (required)
    The cell padding.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `cellPaddingLocation`: `"Top" | "Left" | "Bottom" | "Right"` (required)
    The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
  - `cellPadding`: `number` (required)
    The cell padding.

  **Returns:** `void`

#### Examples

**Example**: Set 10-point padding on all sides of the first cell in the first table

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell (row 0, column 0)
    const firstCell = firstTable.getCell(0, 0);
    
    // Set 10-point padding on all sides
    firstCell.setCellPadding(Word.CellPaddingLocation.top, 10);
    firstCell.setCellPadding(Word.CellPaddingLocation.bottom, 10);
    firstCell.setCellPadding(Word.CellPaddingLocation.left, 10);
    firstCell.setCellPadding(Word.CellPaddingLocation.right, 10);
    
    await context.sync();
});
```

---

### split

Splits the cell into the specified number of rows and columns.

#### Signature

**Parameters:**
- `rowCount`: `number` (required)
  The number of rows to split into. Must be a divisor of the number of underlying rows.
- `columnCount`: `number` (required)
  The number of columns to split into.

**Returns:** `void`

#### Examples

**Example**: Split the first cell of the first table into 2 rows and 3 columns

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell (row 0, column 0)
    const firstCell = firstTable.getCell(0, 0);
    
    // Split the cell into 2 rows and 3 columns
    firstCell.split(2, 3);
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method to provide more useful output when an API object is passed to `JSON.stringify()`. Returns a plain JavaScript object (typed as `Word.Interfaces.TableCellData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.TableCellData`

#### Examples

**Example**: Load table cell properties and serialize them to JSON format for logging or data export purposes

```typescript
await Word.run(async (context) => {
    // Get the first table cell in the document
    const firstTable = context.document.body.tables.getFirst();
    const cell = firstTable.getCell(0, 0);
    
    // Load properties we want to serialize
    cell.load("value,rowIndex,columnIndex,width,cellIndex");
    
    await context.sync();
    
    // Convert the cell object to a plain JSON object
    const cellData = cell.toJSON();
    
    // Now you can stringify and use the data
    console.log(JSON.stringify(cellData, null, 2));
    
    // Example output:
    // {
    //   "value": "Header 1",
    //   "rowIndex": 0,
    //   "columnIndex": 0,
    //   "width": 100,
    //   "cellIndex": 0
    // }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. Shorthand for `context.trackedObjects.add(thisObject)`.

#### Signature

**Returns:** `Word.TableCell`

#### Examples

**Example**: Track a table cell to monitor its properties after making formatting changes, ensuring the object reference remains valid throughout the document session.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first cell in the first row
    const cell = firstTable.rows.getFirst().cells.getFirst();
    
    // Track the cell for automatic adjustment
    cell.track();
    
    // Make changes to the cell
    cell.body.insertText("Tracked Cell Content", Word.InsertLocation.replace);
    cell.shadingColor = "#FFFF00";
    
    // Load cell properties
    cell.load("width, cellIndex");
    
    await context.sync();
    
    // The tracked cell reference remains valid even after sync
    console.log(`Cell index: ${cell.cellIndex}, Width: ${cell.width}`);
    
    // Untrack when done to free up memory
    cell.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. Shorthand for `context.trackedObjects.remove(thisObject)`. Call `context.sync()` for the release to take effect.

#### Signature

**Returns:** `Word.TableCell`

#### Examples

**Example**: Get a reference to the first table cell, perform operations on it, then untrack it to release memory after syncing changes.

```typescript
await Word.run(async (context) => {
    // Get the first table cell in the document
    const firstTable = context.document.body.tables.getFirst();
    const firstCell = firstTable.getCell(0, 0);
    
    // Track the cell for change tracking
    firstCell.track();
    
    // Load and modify the cell
    firstCell.load("value");
    await context.sync();
    
    console.log("Cell value:", firstCell.value);
    firstCell.value = "Updated content";
    
    await context.sync();
    
    // Untrack the cell to release memory
    firstCell.untrack();
    
    await context.sync();
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml
- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties
- /en-us/javascript/api/word/word.body
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.alignment
- /en-us/javascript/api/word/word.tablerow
- /en-us/javascript/api/word/word.table
- /en-us/javascript/api/word/word.verticalalignment
- /en-us/javascript/api/word/word.borderlocation
- /en-us/javascript/api/word/word.tableborder
- /en-us/javascript/api/word/word.cellpaddinglocation
- /en-us/javascript/api/office/officeextension.clientresult
- /en-us/javascript/api/word/word.tablecell
- /en-us/javascript/api/word/word.interfaces.tablecellloadoptions
- /en-us/javascript/api/word/word.interfaces.tablecellupdatedata
- /en-us/javascript/api/office/officeextension.updateoptions
- /en-us/javascript/api/word/word.tablerowcollection
- /en-us/javascript/api/word/word.interfaces.tablecelldata
- /en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
