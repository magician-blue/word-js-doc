# Word.TableRowCollection

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains the collection of the document's TableRow objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets content alignment details about the first row of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  firstTableRow.load(["horizontalAlignment", "verticalAlignment"]);
  await context.sync();

  console.log(
    `Details about the alignment of the first table's first row:`,
    `- Horizontal alignment of every cell in the row: ${firstTableRow.horizontalAlignment}`,
    `- Vertical alignment of every cell in the row: ${firstTableRow.verticalAlignment}`
  );
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a table row collection to verify the connection to the Word host application before performing operations on table rows.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const tableRows = table.rows;
    
    // Access the request context from the table row collection
    const rowContext = tableRows.context;
    
    // Verify the context is connected and load row count
    tableRows.load("count");
    await rowContext.sync();
    
    console.log(`Connected to Word. Table has ${tableRows.count} rows.`);
});
```

---

### items

**Type:** `Word.TableRow[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all table rows from the first table in the document and log the count of rows to the console.

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirst();
    const tableRows = firstTable.rows;
    
    tableRows.load("items");
    await context.sync();
    
    console.log(`Total rows in table: ${tableRows.items.length}`);
    
    // Access individual rows from the items array
    tableRows.items.forEach((row, index) => {
        console.log(`Row ${index} found`);
    });
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first row in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.TableRow`

#### Examples

**Example**: Retrieve and display the border properties (type, color, and width) of the bottom border of the first row in the first table of the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first row of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  const borderLocation = Word.BorderLocation.bottom;
  const border: Word.TableBorder = firstTableRow.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(
    `Details about the ${borderLocation} border of the first table's first row:`,
    `- Color: ${border.color}`,
    `- Type: ${border.type}`,
    `- Width: ${border.width} points`
  );
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first row in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.TableRow`

#### Examples

**Example**: Check if a table has any rows and highlight the first row if it exists

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const firstRow = table.rows.getFirstOrNullObject();
    
    // Load the isNullObject property to check if the row exists
    firstRow.load("isNullObject");
    await context.sync();
    
    if (!firstRow.isNullObject) {
        // First row exists - highlight it
        firstRow.font.highlightColor = "yellow";
        console.log("First row highlighted");
    } else {
        console.log("Table has no rows");
    }
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TableRowCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TableRowCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TableRowCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TableRowCollection`

#### Examples

**Example**: Load and display the text content of all cells in the first row of the first table in the document.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the collection of rows from the table
    const tableRows = firstTable.rows;
    
    // Load the items property to access individual rows
    tableRows.load("items");
    await context.sync();
    
    // Get the first row
    const firstRow = tableRows.items[0];
    
    // Get cells from the first row
    const cells = firstRow.cells;
    cells.load("items/value");
    await context.sync();
    
    // Display the cell values
    cells.items.forEach((cell, index) => {
        console.log(`Cell ${index}: ${cell.value}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.TableRowCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableRowCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.TableRowCollectionData`

#### Examples

**Example**: Export table row data to JSON format for logging or external processing by serializing the first table's rows collection.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableRows = firstTable.rows;
    
    // Load properties we want to include in the JSON output
    tableRows.load("items/cellCount, items/rowIndex, items/isHeader");
    
    await context.sync();
    
    // Convert the TableRowCollection to a plain JavaScript object
    const rowsData = tableRows.toJSON();
    
    // Log the JSON representation
    console.log(JSON.stringify(rowsData, null, 2));
    
    // The rowsData object contains an "items" array with row information
    console.log(`Number of rows: ${rowsData.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TableRowCollection`

#### Examples

**Example**: Track table rows from the first table to safely access their properties across multiple sync calls and modify cell values

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableRows = firstTable.rows;
    
    // Track the row collection to prevent InvalidObjectPath errors
    tableRows.track();
    
    // Load row properties
    tableRows.load("items");
    await context.sync();
    
    // Now we can safely work with rows across multiple sync calls
    console.log(`Table has ${tableRows.items.length} rows`);
    
    // Modify cells in tracked rows
    for (let i = 0; i < tableRows.items.length; i++) {
        const cell = tableRows.items[i].cells.getFirst();
        cell.value = `Row ${i + 1}`;
    }
    
    await context.sync();
    
    // Untrack when done to free up memory
    tableRows.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.TableRowCollection`

#### Examples

**Example**: Get all table rows from the first table, process them, and then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const tableRows = firstTable.rows;
    
    // Load the row count
    tableRows.load("count");
    await context.sync();
    
    // Process the rows (e.g., log the count)
    console.log(`Table has ${tableRows.count} rows`);
    
    // Untrack the collection to release memory
    tableRows.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
