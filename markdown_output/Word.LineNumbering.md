# Word.LineNumbering

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents line numbers in the left margin or to the left of each newspaper-style column.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the line numbering context to ensure the object is properly loaded before reading its properties.

```typescript
await Word.run(async (context) => {
    const lineNumbering = context.document.body.sections.getFirst().lineNumbering;
    
    // Access the context property to verify connection to Office host
    const requestContext = lineNumbering.context;
    
    // Load properties using the context
    lineNumbering.load("restartMode,startingNumber");
    await requestContext.sync();
    
    console.log("Line numbering restart mode: " + lineNumbering.restartMode);
    console.log("Starting number: " + lineNumbering.startingNumber);
});
```

---

### countBy

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the numeric increment for line numbers.

#### Examples

**Example**: Set line numbering to increment by 5 (showing line numbers 5, 10, 15, etc.)

```typescript
await Word.run(async (context) => {
    const lineNumbering = context.document.body.sections.getFirst().lineNumbering;
    lineNumbering.countBy = 5;
    
    await context.sync();
});
```

---

### distanceFromText

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.

#### Examples

**Example**: Set the distance between line numbers and document text to 36 points (half an inch)

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const lineNumbering = body.lineNumbering;
    
    lineNumbering.distanceFromText = 36;
    
    await context.sync();
});
```

---

### isActive

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if line numbering is active for the specified document, section, or sections.

#### Examples

**Example**: Check if line numbering is currently active in the document and display the result in the console.

```typescript
await Word.run(async (context) => {
    const lineNumbering = context.document.body.parentContentControlOrNullObject.parentBody.sections.getFirst().lineNumbering;
    lineNumbering.load("isActive");
    
    await context.sync();
    
    console.log("Line numbering is active: " + lineNumbering.isActive);
});
```

---

### restartMode

**Type:** `Word.NumberingRule | "RestartContinuous" | "RestartSection" | "RestartPage"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.

#### Examples

**Example**: Configure line numbering to restart at the beginning of each page in the document

```typescript
await Word.run(async (context) => {
    // Get the line numbering settings for the document body
    const lineNumbering = context.document.body.sections.getFirst().lineNumbering;
    
    // Set line numbering to restart on each page
    lineNumbering.restartMode = Word.NumberingRule.restartPage;
    
    await context.sync();
    
    console.log("Line numbering set to restart on each page");
});
```

---

### startingNumber

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the starting line number.

#### Examples

**Example**: Set the starting line number to 5 for the document's line numbering

```typescript
await Word.run(async (context) => {
    const lineNumbering = context.document.body.parentContentControlOrNullObject.parentBody.sections.getFirst().lineNumbering;
    
    lineNumbering.startingNumber = 5;
    
    await context.sync();
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
  - `options`: `Word.Interfaces.LineNumberingLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.LineNumbering`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.LineNumbering`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.LineNumbering`

#### Examples

**Example**: Load and read the line numbering restart setting from the current document section

```typescript
await Word.run(async (context) => {
    // Get the first section of the document
    const section = context.document.sections.getFirst();
    const lineNumbering = section.body.lineNumbering;
    
    // Load the restart mode property
    lineNumbering.load("restartMode");
    
    await context.sync();
    
    // Read the loaded property
    console.log("Line numbering restart mode: " + lineNumbering.restartMode);
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.LineNumberingUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.LineNumbering` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure line numbering to restart at 1 for each page and display every 5th line number

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const lineNumbering = body.sections.getFirst().lineNumbering;
    
    lineNumbering.set({
        restartMode: Word.LineNumberingRestartMode.restartPage,
        countBy: 5,
        start: 1
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.LineNumbering object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.LineNumberingData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.LineNumberingData`

#### Examples

**Example**: Get the line numbering settings from the document body and log them as a JSON string to the console.

```typescript
await Word.run(async (context) => {
    // Get the line numbering settings from the document body
    const lineNumbering = context.document.body.lineNumbering;
    
    // Load the properties we want to inspect
    lineNumbering.load("startingNumber,countBy,restartMode");
    
    await context.sync();
    
    // Convert the line numbering object to a plain JavaScript object
    const lineNumberingData = lineNumbering.toJSON();
    
    // Log the JSON representation
    console.log("Line Numbering Settings:", JSON.stringify(lineNumberingData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.LineNumbering`

#### Examples

**Example**: Track line numbering settings across multiple operations to ensure the object remains valid when reading and modifying properties in different sync calls.

```typescript
await Word.run(async (context) => {
    const lineNumbering = context.document.body.sections.getFirst().lineNumbering;
    
    // Track the object to prevent InvalidObjectPath errors across sync calls
    lineNumbering.track();
    
    // Load properties
    lineNumbering.load("restartMode,startingNumber");
    await context.sync();
    
    console.log("Current restart mode:", lineNumbering.restartMode);
    console.log("Starting number:", lineNumbering.startingNumber);
    
    // Modify properties in a subsequent operation
    lineNumbering.startingNumber = 1;
    lineNumbering.restartMode = "RestartPage";
    await context.sync();
    
    // Untrack when done
    lineNumbering.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.LineNumbering`

#### Examples

**Example**: Track line numbering settings, modify them, then untrack to release memory after the changes are complete.

```typescript
await Word.run(async (context) => {
    const lineNumbering = context.document.body.sections.getFirst().lineNumbering;
    
    // Track the object to monitor changes
    lineNumbering.track();
    lineNumbering.load("restartMode");
    
    await context.sync();
    
    // Make changes to line numbering
    lineNumbering.restartMode = Word.LineNumberRestartMode.continuous;
    
    await context.sync();
    
    // Release memory after we're done using the tracked object
    lineNumbering.untrack();
    
    await context.sync();
    
    console.log("Line numbering updated and memory released");
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.linenumbering
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
