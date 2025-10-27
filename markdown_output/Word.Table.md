# Word.Table

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a table in a Word document.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/table-cell-access.yaml

await Word.run(async (context) => {
  // Use a two-dimensional array to hold the initial table values.
  const data = [
    ["Tokyo", "Beijing", "Seattle"],
    ["Apple", "Orange", "Pineapple"]
  ];
  const table: Word.Table = context.document.body.insertTable(2, 3, "Start", data);
  table.styleBuiltIn = Word.BuiltInStyleName.gridTable5Dark_Accent2;
  table.styleFirstColumn = false;

  await context.sync();
});
```

## Properties

### alignment

**Type:** `None`

Specifies the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.

#### Examples

**Example**: Center align a table in the document so it appears in the middle of the page column

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set the table alignment to centered
    table.alignment = "Centered";
    
    await context.sync();
});
```

---

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the table's request context to verify the connection between the add-in and Word before performing table operations

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    table.load("values");
    
    // Access the request context associated with the table
    const tableContext = table.context;
    
    // Verify the context is connected before proceeding
    if (tableContext) {
        await context.sync();
        console.log("Table context is connected to Word");
        console.log("Table has " + table.values.length + " rows");
    }
});
```

---

### endnotes

**Type:** `None`

Gets the collection of endnotes in the table.

#### Examples

**Example**: Get all endnotes from a table and display their reference numbers in the console.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the endnotes collection from the table
    const endnotes = table.endnotes;
    
    // Load the endnote reference numbers
    endnotes.load("items/reference");
    
    await context.sync();
    
    // Display the endnote reference numbers
    console.log(`Found ${endnotes.items.length} endnote(s) in the table:`);
    endnotes.items.forEach((endnote, index) => {
        console.log(`Endnote ${index + 1}: Reference number ${endnote.reference}`);
    });
});
```

---

### fields

**Type:** `None`

Gets the collection of field objects in the table.

#### Examples

**Example**: Get all fields in a table and display their types in the console

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the fields collection from the table
    const fields = table.fields;
    
    // Load the items and their types
    fields.load("items/type");
    
    await context.sync();
    
    // Display information about each field
    console.log(`Found ${fields.items.length} field(s) in the table`);
    fields.items.forEach((field, index) => {
        console.log(`Field ${index + 1}: ${field.type}`);
    });
});
```

---

### font

**Type:** `None`

Gets the font. Use this to get and set font name, size, color, and other properties.

#### Examples

**Example**: Set the table font to Arial, size 14, and color it blue

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Access the table's font property and set formatting
    table.font.name = "Arial";
    table.font.size = 14;
    table.font.color = "blue";
    
    await context.sync();
});
```

---

### footnotes

**Type:** `None`

Gets the collection of footnotes in the table.

#### Examples

**Example**: Get all footnotes from a table and display their reference marks and text content in the console.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the footnotes collection from the table
    const footnotes = table.footnotes;
    footnotes.load("items");
    
    await context.sync();
    
    // Log information about each footnote
    console.log(`Found ${footnotes.items.length} footnote(s) in the table`);
    
    for (let i = 0; i < footnotes.items.length; i++) {
        const footnote = footnotes.items[i];
        footnote.load("reference, body/text");
        await context.sync();
        
        console.log(`Footnote ${i + 1}:`);
        console.log(`  Reference: ${footnote.reference}`);
        console.log(`  Text: ${footnote.body.text}`);
    }
});
```

---

### headerRowCount

**Type:** `None`

Specifies the number of header rows.

#### Examples

**Example**: Set the first 2 rows of a table as header rows

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set the number of header rows to 2
    table.headerRowCount = 2;
    
    await context.sync();
});
```

---

### horizontalAlignment

**Type:** `None`

Specifies the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

#### Examples

**Example**: Center align all cells in the first table of the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set horizontal alignment to centered for all cells
    table.horizontalAlignment = "Centered";
    
    await context.sync();
});
```

---

### isUniform

**Type:** `None`

Indicates whether all of the table rows are uniform.

#### Examples

**Example**: Check if a table has uniform rows and display an alert message based on the result

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Load the isUniform property
    table.load("isUniform");
    
    await context.sync();
    
    // Check if all rows are uniform and display appropriate message
    if (table.isUniform) {
        console.log("All table rows are uniform (same number of cells).");
    } else {
        console.log("Table rows are not uniform (different number of cells).");
    }
});
```

---

### nestingLevel

**Type:** `None`

Gets the nesting level of the table. Top-level tables have level 1.

#### Examples

**Example**: Check if a table is a top-level table or nested, and display its nesting level in the console

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    table.load("nestingLevel");
    
    await context.sync();
    
    console.log(`Table nesting level: ${table.nestingLevel}`);
    
    if (table.nestingLevel === 1) {
        console.log("This is a top-level table");
    } else {
        console.log("This is a nested table");
    }
});
```

---

### parentBody

**Type:** `None`

Gets the parent body of the table.

#### Examples

**Example**: Get the table's parent body and highlight it with yellow background color to visually identify which body section contains the table.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the parent body of the table
    const parentBody = table.parentBody;
    
    // Highlight the parent body with yellow background
    parentBody.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### parentContentControl

**Type:** `None`

Gets the content control that contains the table. Throws an ItemNotFound error if there isn't a parent content control.

#### Examples

**Example**: Get the title of the content control that contains a table and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the parent content control that contains the table
    const parentContentControl = table.parentContentControl;
    
    // Load the title property of the parent content control
    parentContentControl.load("title");
    
    await context.sync();
    
    // Display the parent content control's title
    console.log("Parent content control title: " + parentContentControl.title);
});
```

---

### parentContentControlOrNullObject

**Type:** `None`

Gets the content control that contains the table. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if a table is inside a content control and highlight the content control's title if it exists

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the parent content control (or null object if none exists)
    const parentContentControl = table.parentContentControlOrNullObject;
    
    // Load properties
    parentContentControl.load("isNullObject, title");
    
    await context.sync();
    
    // Check if the table has a parent content control
    if (!parentContentControl.isNullObject) {
        console.log("Table is inside a content control with title: " + parentContentControl.title);
        // Highlight the content control
        parentContentControl.appearance = Word.ContentControlAppearance.tags;
    } else {
        console.log("Table is not inside a content control");
    }
    
    await context.sync();
});
```

---

### parentTable

**Type:** `None`

Gets the table that contains this table. Throws an ItemNotFound error if it isn't contained in a table.

#### Examples

**Example**: Check if a table is nested inside another table and highlight the parent table's first cell if it exists.

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    
    try {
        const parentTable = table.parentTable;
        parentTable.load("values");
        await context.sync();
        
        // If we reach here, the table is nested - highlight parent's first cell
        parentTable.getCell(0, 0).body.font.highlightColor = "yellow";
        await context.sync();
        
        console.log("Table is nested inside a parent table");
    } catch (error) {
        if (error.code === "ItemNotFound") {
            console.log("Table is not nested - it's a top-level table");
        } else {
            throw error;
        }
    }
});
```

---

### parentTableCell

**Type:** `None`

Gets the table cell that contains this table. Throws an ItemNotFound error if it isn't contained in a table cell.

#### Examples

**Example**: Check if a table is nested inside another table by accessing its parent table cell and logging the cell's position information.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the parent table cell that contains this table
    const parentCell = table.parentTableCell;
    
    // Load properties of the parent cell
    parentCell.load("rowIndex, columnIndex");
    
    await context.sync();
    
    // Log the parent cell information
    console.log(`This table is nested in cell at row ${parentCell.rowIndex}, column ${parentCell.columnIndex}`);
});
```

---

### parentTableCellOrNullObject

**Type:** `None`

Gets the table cell that contains this table. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if a table is nested inside another table's cell, and if so, highlight the parent cell with yellow shading.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the parent table cell (or null object if not nested)
    const parentCell = table.parentTableCellOrNullObject;
    
    // Load the isNullObject property to check if the table is nested
    parentCell.load("isNullObject");
    
    await context.sync();
    
    // Check if the table is nested in another table
    if (!parentCell.isNullObject) {
        // Table is nested - highlight the parent cell
        parentCell.shadingColor = "yellow";
        console.log("Table is nested. Parent cell highlighted.");
    } else {
        console.log("Table is not nested in another table.");
    }
    
    await context.sync();
});
```

---

### parentTableOrNullObject

**Type:** `None`

Gets the table that contains this table. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if a table is nested inside another table and highlight the parent table's borders in red if it exists.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const parentTable = table.parentTableOrNullObject;
    
    // Load properties to check if parent exists
    parentTable.load("isNullObject");
    await context.sync();
    
    if (!parentTable.isNullObject) {
        // This table is nested - highlight the parent table's borders
        parentTable.load("borders");
        await context.sync();
        
        parentTable.borders.outsideBorderColor = "red";
        parentTable.borders.outsideBorderWidth = 3;
        console.log("This table is nested. Parent table borders highlighted.");
    } else {
        console.log("This table is not nested inside another table.");
    }
    
    await context.sync();
});
```

---

### rowCount

**Type:** `None`

Gets the number of rows in the table.

#### Examples

**Example**: Display an alert showing how many rows are in the first table of the document

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    firstTable.load("rowCount");
    
    await context.sync();
    
    console.log(`The table has ${firstTable.rowCount} rows`);
});
```

---

### rows

**Type:** `None`

Gets all of the table rows.

#### Examples

**Example**: Get all rows from a table and highlight the text in each row with yellow color.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get all rows from the table
    const rows = table.rows;
    rows.load("items");
    
    await context.sync();
    
    // Highlight text in each row
    for (let i = 0; i < rows.items.length; i++) {
        rows.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### shadingColor

**Type:** `None`

Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.

#### Examples

**Example**: Set the shading color of a table to light blue

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set the shading color to light blue
    table.shadingColor = "#ADD8E6";
    
    await context.sync();
});
```

---

### style

**Type:** `None`

Specifies the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

#### Examples

**Example**: Apply a custom table style named "MyCustomTableStyle" to the first table in the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Apply the custom style to the table
    table.style = "MyCustomTableStyle";
    
    await context.sync();
});
```

---

### styleBandedColumns

**Type:** `None`

Specifies whether the table has banded columns.

#### Examples

**Example**: Enable banded columns styling on the first table in the document to create alternating column colors

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Enable banded columns
    table.styleBandedColumns = true;
    
    await context.sync();
});
```

---

### styleBandedRows

**Type:** `None`

Specifies whether the table has banded rows.

#### Examples

**Example**: Enable banded rows styling on a table to alternate row colors for better readability

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Enable banded rows
    table.styleBandedRows = true;
    
    await context.sync();
});
```

---

### styleBuiltIn

**Type:** `None`

Specifies the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

#### Examples

**Example**: Apply the built-in "Grid Table 1 Light" style to the first table in the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Apply a built-in table style
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
    
    await context.sync();
});
```

---

### styleFirstColumn

**Type:** `None`

Specifies whether the table has a first column with a special style.

#### Examples

**Example**: Apply special styling to the first column of a table by enabling the first column style option.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Enable special styling for the first column
    table.styleFirstColumn = true;
    
    await context.sync();
});
```

---

### styleLastColumn

**Type:** `None`

Specifies whether the table has a last column with a special style.

#### Examples

**Example**: Enable special styling for the last column of a table to make it visually distinct from other columns

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Enable special styling for the last column
    table.styleLastColumn = true;
    
    await context.sync();
});
```

---

### styleTotalRow

**Type:** `None`

Specifies whether the table has a total (last) row with a special style.

#### Examples

**Example**: Enable the total row style for a table to highlight the last row with special formatting

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Enable the total row style
    table.styleTotalRow = true;
    
    await context.sync();
});
```

---

### tables

**Type:** `None`

Gets the child tables nested one level deeper.

#### Examples

**Example**: Get all child tables nested within the first table of the document and highlight them in yellow.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the child tables nested one level deeper within the first table
    const childTables = firstTable.tables;
    childTables.load("items");
    
    await context.sync();
    
    // Highlight each child table in yellow
    for (let i = 0; i < childTables.items.length; i++) {
        childTables.items[i].shadingColor = "#FFFF00";
    }
    
    await context.sync();
    
    console.log(`Found ${childTables.items.length} nested table(s)`);
});
```

---

### values

**Type:** `None`

Specifies the text values in the table, as a 2D JavaScript array.

#### Examples

**Example**: Read all text values from a table and display them in the console, then update specific cell values in the table.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Load the values property
    table.load("values");
    await context.sync();
    
    // Read and log the current table values
    console.log("Current table values:", table.values);
    
    // Update the table values (modify specific cells)
    const updatedValues = table.values;
    updatedValues[0][0] = "Updated Header";
    updatedValues[1][1] = "New Value";
    
    // Set the new values back to the table
    table.values = updatedValues;
    
    await context.sync();
});
```

---

### verticalAlignment

**Type:** `None`

Specifies the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.

#### Examples

**Example**: Set the vertical alignment of all cells in the first table to center

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set vertical alignment to center for all cells
    table.verticalAlignment = "Center";
    
    await context.sync();
});
```

---

### width

**Type:** `None`

Specifies the width of the table in points.

#### Examples

**Example**: Set the width of the first table in the document to 400 points

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    firstTable.width = 400;
    
    await context.sync();
});
```

---

## Methods

### addColumns

**Kind:** `create`

Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.

#### Signature

**Parameters:**
- `insertLocation`: `None` (required)
- `columnCount`: `None` (required)
- `values`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Add 2 columns to the end of the first table in the document with header values "Q3" and "Q4"

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Add 2 columns to the end of the table with values
    table.addColumns("End", 2, [["Q3", "Q4"]]);
    
    await context.sync();
});
```

---

### addRows

**Kind:** `create`

Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.

#### Signature

**Parameters:**
- `insertLocation`: `None` (required)
- `rowCount`: `None` (required)
- `values`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Add 3 new rows at the end of the first table with employee data

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    
    // Add 3 rows at the end with employee data
    table.addRows(
        Word.InsertLocation.end,
        3,
        [
            ["John Smith", "Engineering", "Senior Developer"],
            ["Sarah Johnson", "Marketing", "Manager"],
            ["Mike Davis", "Sales", "Representative"]
        ]
    );
    
    await context.sync();
});
```

---

### autoFitWindow

**Kind:** `configure`

Autofits the table columns to the width of the window.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Autofit an existing table's columns to match the window width

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Autofit the table columns to the window width
    table.autoFitWindow();
    
    await context.sync();
});
```

---

### clear

**Kind:** `write`

Clears the contents of the table.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Clear all contents from the first table in the document while keeping the table structure intact

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Clear the contents of the table
    firstTable.clear();
    
    await context.sync();
});
```

---

### delete

**Kind:** `delete`

Deletes the entire table.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Delete the first table found in the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Delete the table
    firstTable.delete();
    
    await context.sync();
});
```

---

### deleteColumns

**Kind:** `delete`

Deletes specific columns. This is applicable to uniform tables.

#### Signature

**Parameters:**
- `columnIndex`: `None` (required)
- `columnCount`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Delete 2 columns starting from column index 1 in the first table of the document

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    
    // Delete 2 columns starting at index 1 (second column)
    firstTable.deleteColumns(1, 2);
    
    await context.sync();
});
```

---

### deleteRows

**Kind:** `delete`

Deletes specific rows.

#### Signature

**Parameters:**
- `rowIndex`: `None` (required)
- `rowCount`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Delete 2 rows starting from row index 1 (the second row) in the first table of the document

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    firstTable.deleteRows(1, 2);
    await context.sync();
});
```

---

### distributeColumns

**Kind:** `configure`

Distributes the column widths evenly. This is applicable to uniform tables.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Distribute the column widths evenly across all columns in the first table of the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Distribute the column widths evenly
    table.distributeColumns();
    
    await context.sync();
});
```

---

### getBorder

**Kind:** `read`

Gets the border style for the specified border.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `borderLocation`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `borderLocation`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Get the border style of the top border of the first table in the document and display its properties in the console.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the top border of the table
    const topBorder = table.getBorder(Word.BorderLocation.top);
    
    // Load border properties
    topBorder.load("type, color, width");
    
    await context.sync();
    
    // Display border properties
    console.log("Top Border Type:", topBorder.type);
    console.log("Top Border Color:", topBorder.color);
    console.log("Top Border Width:", topBorder.width);
});
```

---

### getCell

**Kind:** `read`

Gets the table cell at a specified row and column. Throws an ItemNotFound error if the specified table cell doesn't exist.

#### Signature

**Parameters:**
- `rowIndex`: `None` (required)
- `cellIndex`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Get the cell in the second row and third column of the first table and highlight it with yellow background color.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the cell at row index 1 (second row) and cell index 2 (third column)
    const cell = table.getCell(1, 2);
    
    // Highlight the cell with yellow background
    cell.shadingColor = "#FFFF00";
    
    await context.sync();
});
```

---

### getCellOrNullObject

**Kind:** `read`

Gets the table cell at a specified row and column. If the specified table cell doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Signature

**Parameters:**
- `rowIndex`: `None` (required)
- `cellIndex`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Check if a cell exists at row 2, column 3 of the first table, and if it exists, highlight it yellow; otherwise, insert a paragraph indicating the cell doesn't exist.

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const cell = firstTable.getCellOrNullObject(2, 3);
    
    cell.load("isNullObject");
    await context.sync();
    
    if (cell.isNullObject) {
        const paragraph = context.document.body.insertParagraph(
            "Cell at row 2, column 3 does not exist.",
            Word.InsertLocation.end
        );
        paragraph.font.color = "red";
    } else {
        cell.body.font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### getCellPadding

**Kind:** `read`

Gets cell padding in points.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `cellPaddingLocation`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `cellPaddingLocation`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Get the top cell padding value from the first cell in the first table and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first cell in the table
    const cell = table.getCell(0, 0);
    
    // Get the top cell padding
    const topPadding = cell.getCellPadding(Word.CellPaddingLocation.top);
    
    // Load the value property
    topPadding.load("value");
    
    // Sync to get the actual value
    await context.sync();
    
    // Display the padding value
    console.log(`Top cell padding: ${topPadding.value} points`);
});
```

---

### getNext

**Kind:** `read`

Gets the next table. Throws an ItemNotFound error if this table is the last one.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get the next table after the first table in the document and highlight it with a yellow background color.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the next table after the first one
    const nextTable = firstTable.getNext();
    
    // Apply yellow shading to the next table
    nextTable.shadingColor = "#FFFF00";
    
    await context.sync();
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next table. If this table is the last one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Check if there is a table after the current table and highlight it yellow if it exists

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the next table (or null object if none exists)
    const nextTable = firstTable.getNextOrNullObject();
    
    // Load properties to check if it exists
    nextTable.load("isNullObject");
    
    await context.sync();
    
    // Check if the next table exists
    if (!nextTable.isNullObject) {
        // Highlight the next table with yellow shading
        nextTable.shadingColor = "#FFFF00";
        console.log("Next table found and highlighted");
    } else {
        console.log("No table exists after the first table");
    }
    
    await context.sync();
});
```

---

### getParagraphAfter

**Kind:** `read`

Gets the paragraph after the table. Throws an ItemNotFound error if there isn't a paragraph after the table.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get the paragraph that follows a table and highlight it in yellow to make it stand out.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the paragraph after the table
    const paragraphAfter = table.getParagraphAfter();
    
    // Highlight the paragraph in yellow
    paragraphAfter.font.highlightColor = "Yellow";
    
    await context.sync();
});
```

---

### getParagraphAfterOrNullObject

**Kind:** `read`

Gets the paragraph after the table. If there isn't a paragraph after the table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Check if there is a paragraph after the first table in the document, and if it exists, highlight it in yellow.

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const paragraphAfter = firstTable.getParagraphAfterOrNullObject();
    
    await context.sync();
    
    if (!paragraphAfter.isNullObject) {
        paragraphAfter.font.highlightColor = "yellow";
    } else {
        console.log("No paragraph exists after the table.");
    }
    
    await context.sync();
});
```

---

### getParagraphBefore

**Kind:** `read`

Gets the paragraph before the table. Throws an ItemNotFound error if there isn't a paragraph before the table.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Check if there's a paragraph before the first table and highlight it in yellow

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    
    try {
        const paragraphBefore = firstTable.getParagraphBefore();
        paragraphBefore.load("text");
        
        await context.sync();
        
        // Highlight the paragraph before the table
        paragraphBefore.font.highlightColor = "yellow";
        
        await context.sync();
        
        console.log("Highlighted paragraph before table:", paragraphBefore.text);
    } catch (error) {
        console.log("No paragraph found before the table");
    }
});
```

---

### getParagraphBeforeOrNullObject

**Kind:** `read`

Gets the paragraph before the table. If there isn't a paragraph before the table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Check if there is a paragraph before the first table and highlight it yellow if it exists, otherwise insert a new paragraph before the table.

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const paragraphBefore = firstTable.getParagraphBeforeOrNullObject();
    
    await context.sync();
    
    if (paragraphBefore.isNullObject) {
        // No paragraph exists before the table, so insert one
        const newParagraph = firstTable.insertParagraph("This paragraph was inserted before the table.", Word.InsertLocation.before);
        newParagraph.font.color = "blue";
    } else {
        // Paragraph exists, highlight it
        paragraphBefore.font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### getRange

**Kind:** `read`

Gets the range that contains this table, or the range at the start or end of the table.

#### Signature

**Parameters:**
- `rangeLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Get the range of the first table in the document and highlight it with yellow color

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the range that contains the entire table
    const tableRange = table.getRange(Word.RangeLocation.whole);
    
    // Highlight the table range with yellow color
    tableRange.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### insertContentControl

**Kind:** `create`

Inserts a content control on the table.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Insert a content control around an existing table to make it a reusable building block

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Insert a content control on the table
    const contentControl = table.insertContentControl();
    contentControl.title = "Sales Data Table";
    contentControl.tag = "salesTable";
    contentControl.appearance = "BoundingBox";
    
    // Load and sync to apply changes
    await context.sync();
    
    console.log("Content control inserted on table");
});
```

---

### insertParagraph

**Kind:** `create`

Inserts a paragraph at the specified location.

#### Signature

**Parameters:**
- `paragraphText`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a paragraph with text "Summary of Results" after an existing table in the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Insert a paragraph after the table
    table.insertParagraph("Summary of Results", Word.InsertLocation.after);
    
    await context.sync();
});
```

---

### insertTable

**Kind:** `create`

Inserts a table with the specified number of rows and columns.

#### Signature

**Parameters:**
- `rowCount`: `None` (required)
- `columnCount`: `None` (required)
- `insertLocation`: `None` (required)
- `values`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a 3x4 table below an existing table with predefined header and data values

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Define table data with headers and sample rows
    const tableData = [
        ["Product", "Category", "Price", "Stock"],
        ["Laptop", "Electronics", "$999", "15"],
        ["Mouse", "Accessories", "$25", "50"]
    ];
    
    // Insert a new 3x4 table below the first table
    const newTable = firstTable.insertTable(3, 4, Word.InsertLocation.after, tableData);
    
    // Optional: Format the new table
    newTable.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
    
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
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `None` (required)

  **Returns:** `None`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Load and display the row count and column count of the first table in the document

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    
    // Load specific properties of the table
    firstTable.load("rowCount, columnCount");
    
    await context.sync();
    
    console.log(`Table has ${firstTable.rowCount} rows and ${firstTable.columnCount} columns`);
});
```

---

### mergeCells

**Kind:** `write`

Merges the cells bounded inclusively by a first and last cell.

#### Signature

**Parameters:**
- `topRow`: `None` (required)
- `firstCell`: `None` (required)
- `bottomRow`: `None` (required)
- `lastCell`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Merge a 2x2 range of cells in the first table by combining cells from row 1 column 0 to row 2 column 1

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    
    // Merge cells from row 1, cell 0 to row 2, cell 1 (creating a 2x2 merged cell)
    firstTable.mergeCells(1, 0, 2, 1);
    
    await context.sync();
});
```

---

### search

**Kind:** `read`

Performs a search with the specified SearchOptions on the scope of the table object. The search results are a collection of range objects.

#### Signature

**Parameters:**
- `searchText`: `None` (required)
- `searchOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Search for all instances of the word "Revenue" in the first table and highlight them in yellow

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Search for "Revenue" in the table
    const searchResults = firstTable.search("Revenue", { matchCase: false });
    
    // Load the search results
    searchResults.load("font");
    
    await context.sync();
    
    // Highlight all found instances
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### select

Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `selectionMode`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `selectionMode`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Select the first table in the document and navigate to it in the Word UI

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Select the entire table and navigate to it
    firstTable.select(Word.SelectionMode.select);
    
    await context.sync();
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `None` (required)
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `properties`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Format an existing table by setting multiple properties at once including style, width, and alignment

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Set multiple properties at once
    table.set({
        style: "Grid Table 4 - Accent 1",
        width: 400,
        horizontalAlignment: Word.Alignment.centered,
        shadingColor: "#F0F0F0"
    });
    
    await context.sync();
});
```

---

### setCellPadding

**Kind:** `write`

Sets cell padding in points.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `cellPaddingLocation`: `None` (required)
  - `cellPadding`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `cellPaddingLocation`: `None` (required)
  - `cellPadding`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Set 10-point padding on all sides of cells in the first table of the document

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    
    firstTable.setCellPadding(Word.CellPaddingLocation.all, 10);
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Table object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Export a table's properties to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Load properties you want to export
    table.load("values, rowCount, columnCount, style, width");
    
    await context.sync();
    
    // Convert the table to a plain JavaScript object
    const tableData = table.toJSON();
    
    // Now you can use the plain object (e.g., log it, store it, send it)
    console.log("Table data:", JSON.stringify(tableData, null, 2));
    
    // Example: Access properties from the plain object
    console.log(`Rows: ${tableData.rowCount}, Columns: ${tableData.columnCount}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a table object to maintain its reference across multiple sync calls while modifying its properties in different batches

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    table.load("values");
    await context.sync();
    
    // Track the table to use it across multiple sync calls
    table.track();
    
    // First batch of changes
    table.headerRowCount = 1;
    await context.sync();
    
    // Second batch - table reference is still valid because it's tracked
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
    await context.sync();
    
    // Third batch - still valid
    table.shadingColor = "#E7E6E6";
    await context.sync();
    
    // Untrack when done to free up memory
    table.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so p

#### Signature

**Returns:** `None`

#### Examples

**Example**: Create a table, perform operations on it, then untrack it to release memory when no longer needed

```typescript
await Word.run(async (context) => {
    // Insert and track a table
    const body = context.document.body;
    const table = body.insertTable(3, 3, Word.InsertLocation.end, [
        ["Header 1", "Header 2", "Header 3"],
        ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
        ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"]
    ]);
    
    // Track the table for automatic memory management
    table.track();
    
    // Perform operations on the table
    table.headerRowCount = 1;
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
    
    await context.sync();
    
    // Once done with the table, untrack it to release memory
    table.untrack();
    
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.table
