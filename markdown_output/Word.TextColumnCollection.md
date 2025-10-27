# TextColumnCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

A collection of [Word.TextColumn](/en-us/javascript/api/word/word.textcolumn) objects that represent all the columns of text in the document or a section of the document.

## Properties

### context

**Type:** `RequestContext`

**Since:** WordApi BETA (PREVIEW ONLY)

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TextColumnCollection to verify the connection to the Word host application and log its properties.

```typescript
await Word.run(async (context) => {
    // Get the text columns from the first section
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    
    // Load the text columns collection
    textColumns.load("count");
    await context.sync();
    
    // Access the request context from the TextColumnCollection
    const requestContext = textColumns.context;
    
    // Verify the context is available and log information
    console.log("Request context available:", requestContext !== null);
    console.log("Number of text columns:", textColumns.count);
    
    // The context property connects the add-in to the Office host
    // It's the same context object passed to Word.run
    console.log("Context matches Word.run context:", requestContext === context);
});
```

---

### items

**Type:** `Word.TextColumn[]`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the loaded child items in this collection.

#### Examples

**Example**: Display the width of each text column in the first section of the document to the console.

```typescript
await Word.run(async (context) => {
    // Get the first section's text columns
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    
    // Load the items property to access the collection of TextColumn objects
    textColumns.load("items");
    
    await context.sync();
    
    // Iterate through the items array and log each column's width
    textColumns.items.forEach((column, index) => {
        column.load("width");
    });
    
    await context.sync();
    
    textColumns.items.forEach((column, index) => {
        console.log(`Column ${index + 1} width: ${column.width}`);
    });
});
```

---

## Methods

### add

**Kind:** `create`

Returns a TextColumn object that represents a new text column added to a section or document.

#### Signature

**Parameters:**
- `options`: `Word.TextColumnAddOptions` (optional)
  Options for configuring the new text column.

**Returns:** `Word.TextColumn`
A TextColumn object that represents a new text column added to the document.

#### Examples

**Example**: Add a new text column to the first section of the document to create a two-column layout

```typescript
await Word.run(async (context) => {
    // Get the first section of the document
    const firstSection = context.document.sections.getFirst();
    
    // Get the text column collection for the section
    const columns = firstSection.body.textColumns;
    
    // Load the current column count
    columns.load("count");
    await context.sync();
    
    // Add a new column to create multi-column layout
    const newColumn = columns.add();
    
    await context.sync();
    
    console.log("New text column added successfully");
});
```

---

### getFlowDirection

**Kind:** `read`

Gets the direction in which text flows from one text column to the next.

#### Signature

**Returns:** `OfficeExtension.ClientResult<Word.FlowDirection>`

#### Examples

**Example**: Check the text flow direction of columns in the first section and display it in the console

```typescript
await Word.run(async (context) => {
    // Get the first section of the document
    const firstSection = context.document.sections.getFirst();
    
    // Get the text columns collection
    const textColumns = firstSection.textColumns;
    
    // Get the flow direction
    const flowDirection = textColumns.getFlowDirection();
    
    // Load the flow direction value
    await context.sync();
    
    // Display the flow direction
    console.log("Text column flow direction: " + flowDirection.value);
    // Possible values: "LeftToRight" or "RightToLeft"
});
```

---

### getHasLineBetween

**Kind:** `read`

Gets whether vertical lines appear between all the columns in the TextColumnCollection object.

#### Signature

**Returns:** `OfficeExtension.ClientResult<boolean>`

#### Examples

**Example**: Check if vertical lines are displayed between text columns in the active document and display the result in the console.

```typescript
await Word.run(async (context) => {
    // Get the text column collection from the body of the document
    const textColumns = context.document.body.textColumns;
    
    // Get whether lines appear between columns
    const hasLineBetween = textColumns.getHasLineBetween();
    
    // Load the property
    await context.sync();
    
    // Display the result
    console.log(`Vertical lines between columns: ${hasLineBetween.value}`);
});
```

---

### getIsEvenlySpaced

**Kind:** `read`

Gets whether text columns are evenly spaced.

#### Signature

**Returns:** `OfficeExtension.ClientResult<boolean>`

#### Examples

**Example**: Check if the text columns in the first section are evenly spaced and display the result in the console.

```typescript
await Word.run(async (context) => {
    // Get the first section of the document
    const firstSection = context.document.sections.getFirst();
    
    // Get the text columns collection
    const textColumns = firstSection.body.textColumns;
    
    // Check if columns are evenly spaced
    const isEvenlySpaced = textColumns.getIsEvenlySpaced();
    
    // Load the property
    await context.sync();
    
    // Display the result
    console.log(`Text columns are evenly spaced: ${isEvenlySpaced.value}`);
});
```

---

### getItem

**Kind:** `read`

Gets a TextColumn by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a TextColumn object.

**Returns:** `Word.TextColumn`

#### Examples

**Example**: Get the second text column from the first section and highlight its text in yellow

```typescript
await Word.run(async (context) => {
    // Get the first section of the document
    const firstSection = context.document.sections.getFirst();
    
    // Get the text columns collection from the section
    const textColumns = firstSection.body.textColumns;
    
    // Get the second column (index 1)
    const secondColumn = textColumns.getItem(1);
    
    // Load the column's properties
    secondColumn.load("text");
    
    await context.sync();
    
    // Highlight the second column's text in yellow
    secondColumn.font.highlightColor = "yellow";
    
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
  - `options`: `Word.Interfaces.TextColumnCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TextColumnCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TextColumnCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TextColumnCollection`

#### Examples

**Example**: Load and display the number of text columns and their spacing in the first section of the document

```typescript
await Word.run(async (context) => {
    // Get the first section's text column collection
    const firstSection = context.document.sections.getFirst();
    const textColumns = firstSection.body.textColumns;
    
    // Load properties of the text column collection
    textColumns.load("items, spacing");
    
    await context.sync();
    
    // Display the loaded information
    console.log(`Number of columns: ${textColumns.items.length}`);
    console.log(`Column spacing: ${textColumns.spacing} points`);
});
```

---

### setCount

**Kind:** `configure`

Arranges text into the specified number of text columns.

#### Signature

**Parameters:**
- `numColumns`: `number` (required)
  The number of columns the text is to be arranged into.

**Returns:** `void`

#### Examples

**Example**: Set the current section to display text in 3 columns

```typescript
await Word.run(async (context) => {
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    
    textColumns.setCount(3);
    
    await context.sync();
});
```

---

### setFlowDirection

**Kind:** `configure`

Sets the direction in which text flows from one text column to the next.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `value`: `Word.FlowDirection` (required)
    The flow direction to set.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `value`: `"LeftToRight" | "RightToLeft"` (required)
    The flow direction to set.

  **Returns:** `void`

#### Examples

**Example**: Set the text flow direction to right-to-left for the first section's columns (useful for languages like Arabic or Hebrew)

```typescript
await Word.run(async (context) => {
    const firstSection = context.document.sections.getFirst();
    const textColumns = firstSection.body.textColumns;
    
    // Set the flow direction to right-to-left
    textColumns.setFlowDirection(Word.TextFlowDirection.rightToLeft);
    
    await context.sync();
    
    console.log("Text column flow direction set to right-to-left");
});
```

---

### setHasLineBetween

**Kind:** `configure`

Sets whether vertical lines appear between all the columns in the TextColumnCollection object.

#### Signature

**Parameters:**
- `value`: `boolean` (required)
  true to show vertical lines between columns.

**Returns:** `void`

#### Examples

**Example**: Configure a document section to display three text columns with vertical lines between them

```typescript
await Word.run(async (context) => {
    // Get the first section of the document
    const section = context.document.sections.getFirst();
    
    // Get the text columns collection for this section
    const textColumns = section.body.textColumns;
    
    // Load the current columns
    textColumns.load("items");
    await context.sync();
    
    // Set the number of columns to 3
    textColumns.setCount(3);
    
    // Enable vertical lines between the columns
    textColumns.setHasLineBetween(true);
    
    await context.sync();
    
    console.log("Text columns configured with lines between them");
});
```

---

### setIsEvenlySpaced

**Kind:** `configure`

Sets whether text columns are evenly spaced.

#### Signature

**Parameters:**
- `value`: `boolean` (required)
  true to evenly space all the text columns in the document.

**Returns:** `void`

#### Examples

**Example**: Set the text columns in the active document to be unevenly spaced (custom spacing between columns)

```typescript
await Word.run(async (context) => {
    // Get the text columns collection from the body of the document
    const textColumns = context.document.body.textColumns;
    
    // Set columns to be unevenly spaced (custom spacing)
    textColumns.setIsEvenlySpaced(false);
    
    await context.sync();
    
    console.log("Text columns are now set to custom spacing");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TextColumnCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TextColumnCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.TextColumnCollectionData`

#### Examples

**Example**: Export text column configuration to JSON for logging or storage purposes

```typescript
await Word.run(async (context) => {
    // Get the text columns from the first section
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    
    // Load the properties we want to export
    textColumns.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const columnsData = textColumns.toJSON();
    
    // Log or store the JSON data
    console.log("Text Columns Configuration:", JSON.stringify(columnsData, null, 2));
    
    // You can now use this data for storage, comparison, or documentation
    console.log(`Number of columns: ${columnsData.items.length}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TextColumnCollection`

#### Examples

**Example**: Track text columns in a section to monitor and adjust column properties across multiple sync calls without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    const firstSection = context.document.sections.getFirst();
    const textColumns = firstSection.body.textColumns;
    
    // Track the collection to use it across multiple sync calls
    textColumns.track();
    
    textColumns.load("count,spacing,evenlySpaced");
    await context.sync();
    
    console.log(`Current columns: ${textColumns.count}`);
    console.log(`Column spacing: ${textColumns.spacing}`);
    
    // Modify properties across another sync call
    textColumns.spacing = 36; // 0.5 inch spacing
    await context.sync();
    
    console.log("Column spacing updated successfully");
    
    // Untrack when done to free up memory
    textColumns.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TextColumnCollection`

#### Examples

**Example**: Load text column information from a section, use it to log column count, then untrack the collection to free memory.

```typescript
await Word.run(async (context) => {
    // Get the first section's text columns
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    
    // Load and track the collection
    textColumns.load("count");
    await context.sync();
    
    // Use the column information
    console.log(`Number of columns: ${textColumns.count}`);
    
    // Untrack the collection to release memory
    textColumns.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word/word.textcolumncollection
