# Word.TextColumn

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a single text column in a section.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a TextColumn object to verify the connection between the add-in and Word application before performing operations on the column.

```typescript
await Word.run(async (context) => {
    // Get the first section and its text columns
    const section = context.document.sections.getFirst();
    const textColumns = section.textColumns;
    textColumns.load("items");
    
    await context.sync();
    
    if (textColumns.items.length > 0) {
        const firstColumn = textColumns.items[0];
        
        // Access the request context from the TextColumn object
        const columnContext = firstColumn.context;
        
        // Verify the context is valid and connected
        console.log("TextColumn context is connected:", columnContext !== null);
        console.log("Context type:", typeof columnContext);
        
        // Use the context to perform operations
        firstColumn.load("width");
        await columnContext.sync();
        
        console.log("Column width:", firstColumn.width);
    }
});
```

---

### spaceAfter

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the amount of spacing (in points) after the specified paragraph or text column.

#### Examples

**Example**: Set the spacing after a text column to 24 points to add visual separation between columns

```typescript
await Word.run(async (context) => {
    const section = context.document.sections.getFirst();
    const textColumns = section.textColumns;
    textColumns.load("items");
    
    await context.sync();
    
    // Set 24 points of spacing after the first text column
    if (textColumns.items.length > 0) {
        textColumns.items[0].spaceAfter = 24;
    }
    
    await context.sync();
});
```

---

### width

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width, in points, of the specified text columns.

#### Examples

**Example**: Set the width of the first text column in the active document's first section to 200 points.

```typescript
await Word.run(async (context) => {
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    textColumns.load("items");
    
    await context.sync();
    
    if (textColumns.items.length > 0) {
        const firstColumn = textColumns.items[0];
        firstColumn.width = 200;
        
        await context.sync();
    }
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
  - `options`: `Word.Interfaces.TextColumnLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TextColumn`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TextColumn`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TextColumn`

#### Examples

**Example**: Read and display the width and spacing properties of the first text column in the active document's first section.

```typescript
await Word.run(async (context) => {
    // Get the first section and its text columns
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    const firstColumn = textColumns.getFirst();
    
    // Load properties of the first text column
    firstColumn.load("width, spaceAfter");
    
    await context.sync();
    
    // Display the loaded properties
    console.log(`Column width: ${firstColumn.width}`);
    console.log(`Space after column: ${firstColumn.spaceAfter}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.TextColumnUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.TextColumn` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple properties of a text column to set its width to 200 points and spacing to 36 points

```typescript
await Word.run(async (context) => {
    // Get the first section's text columns
    const section = context.document.sections.getFirst();
    const textColumn = section.body.textColumns.getFirst();
    
    // Set multiple properties at once
    textColumn.set({
        width: 200,
        spaceAfter: 36
    });
    
    await context.sync();
    console.log("Text column properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TextColumn object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TextColumnData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.TextColumnData`

#### Examples

**Example**: Serialize a text column's properties to a plain JavaScript object for logging or data transfer purposes.

```typescript
await Word.run(async (context) => {
    // Get the first section's text columns
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    
    // Load the first text column with its properties
    const firstColumn = textColumns.getFirst();
    firstColumn.load("width,spaceAfter");
    
    await context.sync();
    
    // Convert the TextColumn object to a plain JavaScript object
    const columnData = firstColumn.toJSON();
    
    // Now you can use the plain object for logging, storage, or data transfer
    console.log("Column width:", columnData.width);
    console.log("Space after:", columnData.spaceAfter);
    console.log("Full column data:", JSON.stringify(columnData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TextColumn`

#### Examples

**Example**: Track a text column object to safely access its properties across multiple sync calls when working with a multi-column section layout.

```typescript
await Word.run(async (context) => {
    // Get the first section and its text columns
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    textColumns.load("items");
    
    await context.sync();
    
    // Get the first column and track it for use across sync calls
    const firstColumn = textColumns.items[0];
    firstColumn.track();
    
    // Load properties
    firstColumn.load("width,spaceAfter");
    await context.sync();
    
    // Now we can safely use the tracked object after another sync
    console.log(`Column width: ${firstColumn.width}`);
    console.log(`Space after: ${firstColumn.spaceAfter}`);
    
    // Untrack when done to free up memory
    firstColumn.untrack();
    
    await context.sync();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TextColumn`

#### Examples

**Example**: Access a text column in a section, perform operations on it, then untrack it to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the first section and its text columns
    const section = context.document.sections.getFirst();
    const textColumns = section.body.textColumns;
    const firstColumn = textColumns.getFirst();
    
    // Track the column object to work with it
    firstColumn.track();
    
    // Load properties to perform operations
    firstColumn.load("width,spaceAfter");
    await context.sync();
    
    // Perform operations with the column
    console.log(`Column width: ${firstColumn.width}`);
    console.log(`Space after: ${firstColumn.spaceAfter}`);
    
    // Untrack the object to release memory after we're done
    firstColumn.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
