# Break

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a break in a Word document. This could be a page, column, or section break.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Break object to verify the connection between the add-in and Word, then use it to load and log break properties.

```typescript
await Word.run(async (context) => {
    // Get the first break in the document body
    const breaks = context.document.body.getBreaks();
    breaks.load("items");
    await context.sync();
    
    if (breaks.items.length > 0) {
        const firstBreak = breaks.items[0];
        
        // Access the request context from the Break object
        const breakContext = firstBreak.context;
        
        // Use the context to load properties
        firstBreak.load("type");
        await breakContext.sync();
        
        console.log("Break type: " + firstBreak.type);
        console.log("Context is connected: " + (breakContext !== null));
    }
});
```

---

### pageIndex

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the page number on which the break occurs.

#### Examples

**Example**: Display the page number where the first page break occurs in the document

```typescript
await Word.run(async (context) => {
    // Get all page breaks in the document
    const breaks = context.document.body.getBreaks(Word.BreakType.page);
    breaks.load("pageIndex");
    
    await context.sync();
    
    if (breaks.items.length > 0) {
        const firstBreak = breaks.items[0];
        console.log(`The first page break occurs on page ${firstBreak.pageIndex}`);
    } else {
        console.log("No page breaks found in the document");
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a Range object that represents the portion of the document that's contained in the break.

#### Examples

**Example**: Get the text content from a page break's range to verify what content exists at the break location.

```typescript
await Word.run(async (context) => {
    // Get all breaks in the document
    const breaks = context.document.body.getBreaks();
    breaks.load("items");
    
    await context.sync();
    
    if (breaks.items.length > 0) {
        // Get the range of the first break
        const breakRange = breaks.items[0].range;
        breakRange.load("text");
        
        await context.sync();
        
        // Access the text content at the break location
        console.log("Text at break location:", breakRange.text);
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
  - `options`: `Word.Interfaces.BreakLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Break`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Break`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Break`

#### Examples

**Example**: Load and display the type of the first break in the document

```typescript
await Word.run(async (context) => {
    // Get the first break in the document
    const breaks = context.document.body.getRange().breaks;
    const firstBreak = breaks.getFirst();
    
    // Load the type property of the break
    firstBreak.load("type");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the break type
    console.log("Break type: " + firstBreak.type);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.BreakUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Break` (required)

  **Returns:** `void`

#### Examples

**Example**: Insert a page break at the end of the document and configure its type to be a page break

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Insert a page break at the end of the document
    const pageBreak = body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    
    // Set properties on the break
    pageBreak.set({
        type: Word.BreakType.page
    });
    
    await context.sync();
    console.log("Page break inserted and configured successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Break object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BreakData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BreakData`

#### Examples

**Example**: Serialize a page break object to JSON format to log or store its properties

```typescript
await Word.run(async (context) => {
    // Insert a page break at the end of the document
    const body = context.document.body;
    const pageBreak = body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    
    // Load properties of the break
    pageBreak.load("type");
    
    await context.sync();
    
    // Convert the break object to a plain JavaScript object
    const breakData = pageBreak.toJSON();
    
    // Log the serialized break data
    console.log("Break data:", JSON.stringify(breakData, null, 2));
    // Output example: { "type": "Page" }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Break`

#### Examples

**Example**: Insert a page break at the end of the document, track it across multiple sync calls, and then change its type to a column break.

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Insert a page break at the end of the document
    const pageBreak = body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    
    // Track the break object for use across sync calls
    pageBreak.track();
    
    await context.sync();
    
    // Now we can safely modify the break after sync
    // Change it to a column break
    pageBreak.delete();
    const columnBreak = body.insertBreak(Word.BreakType.column, Word.InsertLocation.end);
    
    await context.sync();
    
    // Untrack when done
    pageBreak.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Break`

#### Examples

**Example**: Insert a page break in the document, use it to verify its type, then untrack it to free up memory resources.

```typescript
await Word.run(async (context) => {
    // Insert a page break at the end of the document
    const body = context.document.body;
    const pageBreak = body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    
    // Track the break object to work with it
    pageBreak.track();
    
    // Load and use the break's properties
    pageBreak.load("type");
    await context.sync();
    
    console.log("Break type: " + pageBreak.type);
    
    // Once done using the break object, untrack it to release memory
    pageBreak.untrack();
    await context.sync();
    
    console.log("Break object has been untracked and memory released");
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
