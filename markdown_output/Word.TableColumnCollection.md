# Word.TableColumnCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of [Word.TableColumn](/en-us/javascript/api/word/word.tablecolumn) objects in a Word document.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TableColumnCollection to verify the connection between the add-in and Word before performing table operations.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const columns = table.columns;
    
    // Access the request context from the collection
    const requestContext = columns.context;
    
    // Verify the context is valid by using it to load properties
    columns.load("count");
    await requestContext.sync();
    
    console.log(`Connected to Word. Table has ${columns.count} columns.`);
});
```

---

### items

**Type:** `Word.TableColumn[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all columns from the first table and highlight the cells in every other column with yellow background.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const columns = firstTable.columns;
    
    // Load the items property to access the array of columns
    columns.load("items");
    await context.sync();
    
    // Access the loaded columns using the items property
    const columnArray = columns.items;
    
    // Highlight every other column with yellow background
    for (let i = 0; i < columnArray.length; i += 2) {
        const column = columnArray[i];
        column.load("cells");
        await context.sync();
        
        // Apply yellow shading to all cells in this column
        for (const cell of column.cells.items) {
            cell.shadingColor = "#FFFF00";
        }
    }
    
    await context.sync();
});
```

---

## Methods

### add

**Kind:** `create`

Returns a TableColumn object that represents a column added to a table.

#### Signature

**Parameters:**
- `beforeColumn`: `Word.TableColumn` (optional)
  Optional. The column before which the new column is added.

**Returns:** `Word.TableColumn`
A new TableColumn object.

#### Examples

**Example**: Add a new column before the second column in the first table of the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the table columns collection
    const columns = firstTable.columns;
    
    // Get the second column (index 1)
    const secondColumn = columns.getFirst().getNext();
    
    // Add a new column before the second column
    const newColumn = columns.add(secondColumn);
    
    // Load the column properties to verify
    newColumn.load("columnIndex");
    
    await context.sync();
    
    console.log(`New column added at index: ${newColumn.columnIndex}`);
});
```

---

### autoFit

**Kind:** `configure`

Changes the width of a table column to accommodate the width of the text without changing the way text wraps in the cells.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Auto-fit all columns in the first table to accommodate their text content without changing text wrapping

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get all columns in the table
    const columns = firstTable.columns;
    
    // Auto-fit all columns to their content
    columns.autoFit();
    
    await context.sync();
});
```

---

### delete

**Kind:** `delete`

Deletes the specified columns.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete all columns from the first table in the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get all columns from the table
    const columns = firstTable.columns;
    
    // Delete all columns
    columns.delete();
    
    await context.sync();
});
```

---

### distributeWidth

**Kind:** `configure`

Adjusts the width of the specified columns so that they are equal.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Distribute the width equally across all columns in the first table of the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get all columns in the table
    const columns = firstTable.columns;
    
    // Distribute the width equally across all columns
    columns.distributeWidth();
    
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
  - `options`: `Word.Interfaces.TableColumnCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TableColumnCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TableColumnCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TableColumnCollection`

#### Examples

**Example**: Load and display the width of each column in the first table of the document

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const columns = firstTable.columns;
    
    // Load the width property for all columns
    columns.load("items/width");
    
    await context.sync();
    
    // Display the width of each column
    console.log(`Table has ${columns.items.length} columns`);
    columns.items.forEach((column, index) => {
        console.log(`Column ${index + 1} width: ${column.width}`);
    });
});
```

---

### select

**Kind:** `configure`

Selects the specified table columns.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Select all columns in the first table of the document to highlight them for the user

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get all columns in the table
    const columns = firstTable.columns;
    
    // Select all columns
    columns.select();
    
    await context.sync();
});
```

---

### setWidth

**Kind:** `configure`

Sets the width of columns in a table.

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

**Example**: Set all columns in the first table to 100 points width using auto fit ruler style

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    
    // Get the columns collection
    const columns = firstTable.columns;
    
    // Set width to 100 points with auto fit ruler style
    columns.setWidth(100, Word.RulerStyle.autoFit);
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableColumnCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableColumnCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.TableColumnCollectionData`

#### Examples

**Example**: Export table column information to JSON format for logging or external processing

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    const columns = table.columns;
    
    // Load properties we want to include in the JSON output
    columns.load("width,columnIndex");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const columnsJSON = columns.toJSON();
    
    // Now you can use the JSON data (e.g., log it, send to server, etc.)
    console.log("Table columns data:", JSON.stringify(columnsJSON, null, 2));
    console.log("Number of columns:", columnsJSON.items.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TableColumnCollection`

#### Examples

**Example**: Track table columns across multiple sync calls to maintain object references when modifying column properties outside of a single batch operation

```typescript
await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst();
    const columns = table.columns;
    columns.load("items");
    await context.sync();
    
    // Track the columns collection to use it across multiple sync calls
    columns.track();
    
    // First sync: modify first column
    if (columns.items.length > 0) {
        columns.items[0].width = 100;
    }
    await context.sync();
    
    // Second sync: modify second column using the tracked collection
    if (columns.items.length > 1) {
        columns.items[1].width = 150;
    }
    await context.sync();
    
    // Untrack when done to free up memory
    columns.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TableColumnCollection`

#### Examples

**Example**: Load table columns, get their count, then untrack the collection to free memory after use

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const firstTable = context.document.body.tables.getFirst();
    const columns = firstTable.columns;
    
    // Load the column count
    columns.load("count");
    await context.sync();
    
    console.log(`Table has ${columns.count} columns`);
    
    // Untrack the columns collection to release memory
    columns.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
