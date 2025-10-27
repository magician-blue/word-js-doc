# Word.TableCollection

**Package:** `word`

**API Set:** WordApi 1.3 None

**Extends:** `officeextension.clientobject`

## Description

Contains the collection of the document's Table objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets alignment details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  firstTable.load(["alignment", "horizontalAlignment", "verticalAlignment"]);
  await context.sync();

  console.log(
    `Details about the alignment of the first table:`,
    `- Alignment of the table within the containing page column: ${firstTable.alignment}`,
    `- Horizontal alignment of every cell in the table: ${firstTable.horizontalAlignment}`,
    `- Vertical alignment of every cell in the table: ${firstTable.verticalAlignment}`
  );
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TableCollection to verify the connection between the add-in and Word before performing table operations.

```typescript
await Word.run(async (context) => {
    const tables = context.document.body.tables;
    
    // Access the request context associated with the TableCollection
    const requestContext = tables.context;
    
    // Verify the context is valid before proceeding with operations
    if (requestContext) {
        tables.load("items");
        await context.sync();
        
        console.log(`Connected to Word. Found ${tables.items.length} table(s) in the document.`);
    }
});
```

---

### items

**Type:** `Word.Table[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Get all tables in the document and log the count and row count of each table to the console.

```typescript
await Word.run(async (context) => {
    // Get the table collection from the document
    const tables = context.document.body.tables;
    
    // Load the items property to access the array of tables
    tables.load("items");
    
    await context.sync();
    
    // Access the loaded tables through the items property
    console.log(`Total tables found: ${tables.items.length}`);
    
    // Iterate through each table in the items array
    for (let i = 0; i < tables.items.length; i++) {
        const table = tables.items[i];
        table.load("rowCount");
        await context.sync();
        
        console.log(`Table ${i + 1} has ${table.rowCount} rows`);
    }
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first table in this collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.Table`

#### Examples

**Example**: Retrieve and display the text content from the first cell (row 0, column 0) of the first table in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/table-cell-access.yaml

// Gets the content of the first cell in the first table.
await Word.run(async (context) => {
  const firstCell: Word.Body = context.document.body.tables.getFirst().getCell(0, 0).body;
  firstCell.load("text");

  await context.sync();
  console.log("First cell's text is: " + firstCell.text);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first table in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

#### Signature

**Returns:** `Word.Table`

#### Examples

**Example**: Check if the document contains any tables and display an alert with the first table's row count, or notify the user if no tables exist.

```typescript
await Word.run(async (context) => {
    const firstTable = context.document.body.tables.getFirstOrNullObject();
    firstTable.load("isNullObject, rowCount");
    
    await context.sync();
    
    if (firstTable.isNullObject) {
        console.log("No tables found in the document.");
    } else {
        console.log(`First table has ${firstTable.rowCount} rows.`);
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TableCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TableCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TableCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TableCollection`

#### Examples

**Example**: Load all tables in the document and display the count of tables found

```typescript
await Word.run(async (context) => {
    // Get the table collection from the document
    const tables = context.document.body.tables;
    
    // Load the count property of the table collection
    tables.load("count");
    
    // Synchronize the document state
    await context.sync();
    
    // Display the number of tables
    console.log(`Number of tables in document: ${tables.count}`);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.TableCollectionData`

#### Examples

**Example**: Export all tables in the document to JSON format for logging or external processing

```typescript
await Word.run(async (context) => {
    // Get all tables in the document
    const tables = context.document.body.tables;
    
    // Load properties we want to include in the JSON output
    tables.load("items/rowCount, items/columnCount, items/headerRowCount");
    
    await context.sync();
    
    // Convert the table collection to a plain JavaScript object
    const tablesJSON = tables.toJSON();
    
    // Log or process the JSON data
    console.log("Tables data:", JSON.stringify(tablesJSON, null, 2));
    console.log(`Found ${tablesJSON.items.length} table(s) in the document`);
    
    // Access individual table properties from the JSON
    tablesJSON.items.forEach((table, index) => {
        console.log(`Table ${index + 1}: ${table.rowCount} rows, ${table.columnCount} columns`);
    });
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TableCollection`

#### Examples

**Example**: Track all tables in a document and modify their properties across multiple sync calls without losing object references

```typescript
await Word.run(async (context) => {
    // Get all tables in the document
    const tables = context.document.body.tables;
    
    // Track the collection to maintain references across sync calls
    tables.track();
    
    // Load table properties
    tables.load("items");
    await context.sync();
    
    // First sync - add borders to tables
    for (let i = 0; i < tables.items.length; i++) {
        tables.items[i].set({
            styleBuiltIn: Word.BuiltInStyleName.gridTable1Light
        });
    }
    await context.sync();
    
    // Second sync - modify table alignment (object references still valid due to tracking)
    for (let i = 0; i < tables.items.length; i++) {
        tables.items[i].alignment = Word.Alignment.centered;
    }
    await context.sync();
    
    // Untrack when done to release memory
    tables.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TableCollection`

#### Examples

**Example**: Load all tables in the document, process their data, then untrack them to free memory and improve performance.

```typescript
await Word.run(async (context) => {
    // Load all tables in the document
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    
    // Process the tables (e.g., count them)
    console.log(`Found ${tables.items.length} tables in the document`);
    
    // Untrack the collection to release memory
    tables.untrack();
    
    // Sync to apply the memory release
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.tablecollection
