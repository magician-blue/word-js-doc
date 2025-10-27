# TableColumn

**Package:** `Word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a table column in a Word document.

## Properties

### borders

**Type:** `Word.BorderUniversalCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BorderUniversalCollection object that represents all the borders for the table column.

#### Examples

**Example**: Set all borders of the first column in a table to be solid red lines with 2pt width

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first column
    const firstColumn = table.columns.getFirst();
    
    // Get all borders for the column
    const borders = firstColumn.borders;
    borders.load("items");
    
    await context.sync();
    
    // Set properties for all borders in the column
    borders.items.forEach(border => {
        border.type = Word.BorderType.single;
        border.color = "#FF0000"; // Red
        border.width = 2;
    });
    
    await context.sync();
});
```

---

### columnIndex

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the position of this column in a collection.

#### Examples

**Example**: Get the index position of the first column in a table and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first column
    const firstColumn = table.columns.getFirst();
    
    // Load the columnIndex property
    firstColumn.load("columnIndex");
    
    await context.sync();
    
    // Display the column index
    console.log(`Column index: ${firstColumn.columnIndex}`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TableColumn object to load and read the column's width property

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const firstColumn = table.columns.getFirst();
    
    // Access the request context from the TableColumn object
    const columnContext = firstColumn.context;
    
    // Use the context to load properties
    firstColumn.load("width");
    await columnContext.sync();
    
    console.log(`Column width: ${firstColumn.width}`);
});
```

---

### isFirst

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns true if the column or row is the first one in the table; false otherwise.

#### Examples

**Example**: Highlight the first column of a table with a yellow background color to distinguish it from other columns.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const columns = table.columns;
    columns.load("items");
    
    await context.sync();
    
    // Loop through columns and highlight the first one
    for (let i = 0; i < columns.items.length; i++) {
        const column = columns.items[i];
        column.load("isFirst");
        await context.sync();
        
        if (column.isFirst) {
            column.getRange().font.highlightColor = "yellow";
            break;
        }
    }
    
    await context.sync();
});
```

---

### isLast

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns true if the column or row is the last one in the table; false otherwise.

#### Examples

**Example**: Highlight the last column in a table with a yellow background color to make it stand out from other columns.

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    const columns = table.columns;
    columns.load("items");
    
    await context.sync();
    
    for (let i = 0; i < columns.items.length; i++) {
        const column = columns.items[i];
        column.load("isLast");
        await context.sync();
        
        if (column.isLast) {
            column.getCellPadding(Word.CellPaddingLocation.all);
            column.shadingColor = "#FFFF00"; // Yellow background
            break;
        }
    }
    
    await context.sync();
});
```

---

### nestingLevel

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the nesting level of the column.

#### Examples

**Example**: Get the nesting level of the first column in a table and display it in the console to understand the table's structure depth.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first column from the table
    const firstColumn = table.columns.getFirst();
    
    // Load the nesting level property
    firstColumn.load("nestingLevel");
    
    await context.sync();
    
    // Display the nesting level
    console.log(`Column nesting level: ${firstColumn.nestingLevel}`);
});
```

---

### preferredWidth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the preferredWidthType property.

#### Examples

**Example**: Set the first column of a table to have a preferred width of 150 points

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first column
    const firstColumn = table.columns.getFirst();
    
    // Set the preferred width to 150 points
    firstColumn.preferredWidth = 150;
    
    await context.sync();
});
```

---

### preferredWidthType

**Type:** `Word.PreferredWidthType | "Auto" | "Percent" | "Points"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the preferred unit of measurement to use for the width of the table column.

#### Examples

**Example**: Set the first column's width to 2 inches using points as the preferred width type

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first column
    const firstColumn = table.columns.getFirst();
    
    // Set the preferred width type to Points
    firstColumn.preferredWidthType = Word.PreferredWidthType.points;
    
    // Set the width to 2 inches (144 points)
    firstColumn.width = 144;
    
    await context.sync();
    
    console.log("Column width type set to Points");
});
```

---

### shading

**Type:** `Word.ShadingUniversal`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a ShadingUniversal object that refers to the shading formatting for the column.

#### Examples

**Example**: Apply light gray background shading to the first column of a table

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first column
    const firstColumn = table.columns.getFirst();
    
    // Apply light gray shading to the column
    firstColumn.shading.backgroundPatternColor = "#D3D3D3";
    
    await context.sync();
});
```

---

### width

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width of the column, in points.

#### Examples

**Example**: Set the width of the first column in a table to 150 points

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    const firstColumn = table.columns.getFirst();
    
    firstColumn.width = 150;
    
    await context.sync();
});
```

---

## Methods

### autoFit

Changes the width of the table column to accommodate the width of the text without changing the way text wraps in the cells.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Auto-fit the second column of the first table in the document to accommodate its text content

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the second column (index 1)
    const column = table.columns.getItem(1);
    
    // Auto-fit the column to its content
    column.autoFit();
    
    await context.sync();
});
```

---

### delete

**Kind:** `delete`

Deletes the column.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete the second column from the first table in the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the second column (index 1)
    const secondColumn = firstTable.columns.getItemAt(1);
    
    // Delete the column
    secondColumn.delete();
    
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
  - `options`: `Word.Interfaces.TableColumnLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TableColumn`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TableColumn`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TableColumn`

#### Examples

**Example**: Load and display the width and cell count properties of the first column in the first table

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const firstColumn = firstTable.columns.getFirst();
    
    // Load specific properties of the column
    firstColumn.load("width, cellCount");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log(`Column width: ${firstColumn.width}`);
    console.log(`Number of cells: ${firstColumn.cellCount}`);
});
```

---

### select

Selects the table column.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Select the first column of the first table in the document to highlight it for the user

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first column of the table
    const firstColumn = firstTable.columns.getFirst();
    
    // Select the column
    firstColumn.select();
    
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
  - `properties`: `Interfaces.TableColumnUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.TableColumn` (required)

  **Returns:** `void`

#### Examples

**Example**: Format the first column of the first table by setting multiple properties including width, shading color, and preferred width type

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the first column
    const firstColumn = firstTable.columns.getFirst();
    
    // Set multiple properties at once
    firstColumn.set({
        width: 100,
        shadingColor: "#E7E6E6",
        preferredWidth: Word.PreferredWidth.points(100)
    });
    
    await context.sync();
    
    console.log("Column properties updated successfully");
});
```

---

### setWidth

**Kind:** `write`

Sets the width of the column in a table.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `columnWidth`: `number` (required)
    The width to set.
  - `rulerStyle`: `Word.RulerStyle` (required)
    The ruler style to apply.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `columnWidth`: `number` (required)
    The width to set.
  - `rulerStyle`: `"None" | "Proportional" | "FirstColumn" | "SameWidth"` (required)
    The ruler style to apply.

  **Returns:** `void`

#### Examples

**Example**: Set the width of the second column in the first table to 100 points using a fixed ruler style

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the second column (index 1)
    const column = table.columns.getItem(1);
    
    // Set the column width to 100 points with fixed ruler style
    column.setWidth(100, Word.RulerStyle.fixed);
    
    await context.sync();
});
```

---

### sort

Sorts the table column.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Sort a table column in ascending alphabetical order

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first column of the table
    const firstColumn = table.columns.getFirst();
    
    // Sort the column in ascending order
    firstColumn.sort(true);
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableColumn object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableColumnData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.TableColumnData`

#### Examples

**Example**: Serialize a table column's properties to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first column from the table
    const column = table.columns.getFirst();
    
    // Load properties we want to serialize
    column.load("width,values");
    
    await context.sync();
    
    // Convert the column to a plain JavaScript object
    const columnData = column.toJSON();
    
    // Now you can use the plain object (e.g., log it, send to server, etc.)
    console.log("Column data:", JSON.stringify(columnData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TableColumn`

#### Examples

**Example**: Track a table column object to maintain its reference across multiple sync calls while modifying its properties and the document structure

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const firstColumn = table.columns.getFirst();
    
    // Track the column to prevent InvalidObjectPath errors across sync calls
    firstColumn.track();
    
    // Load properties
    firstColumn.load("width");
    await context.sync();
    
    // Modify the column width
    firstColumn.width = 100;
    await context.sync();
    
    // Make other document changes that might affect object paths
    context.document.body.insertParagraph("New paragraph", Word.InsertLocation.start);
    await context.sync();
    
    // Can still safely access the tracked column object
    firstColumn.width = 120;
    await context.sync();
    
    // Untrack when done to free up memory
    firstColumn.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TableColumn`

#### Examples

**Example**: Get a reference to the first column in a table, perform operations on it, then release it from memory tracking to improve performance.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const firstColumn = table.columns.getFirst();
    
    // Track the column object for operations
    firstColumn.track();
    
    // Load and use the column
    firstColumn.load("width");
    await context.sync();
    
    console.log(`Column width: ${firstColumn.width}`);
    
    // Perform operations on the column
    firstColumn.width = 100;
    await context.sync();
    
    // Release the column from tracking when done
    firstColumn.untrack();
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
