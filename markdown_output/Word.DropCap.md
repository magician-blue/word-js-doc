# Word.DropCap

**Package:** `https://learn.microsoft.com/en-us/javascript/api/word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject`

## Description

Represents a dropped capital letter in a Word document.

## Properties

### context

**Type:** `RequestContext`

**Since:** WordApi BETA (PREVIEW ONLY)

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a DropCap object to load and read its properties

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const dropCap = paragraph.dropCap;
    
    // Access the context property to use the same RequestContext
    const requestContext = dropCap.context;
    
    // Use the context to load properties
    dropCap.load("type,linesDropped");
    await requestContext.sync();
    
    console.log(`Drop cap type: ${dropCap.type}`);
    console.log(`Lines dropped: ${dropCap.linesDropped}`);
});
```

---

### distanceFromText

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the distance (in points) between the dropped capital letter and the paragraph text.

#### Examples

**Example**: Get the distance between the drop cap and the paragraph text and display it to the user.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const dropCap = firstParagraph.dropCap;
    
    dropCap.load("distanceFromText");
    await context.sync();
    
    const distance = dropCap.distanceFromText;
    console.log(`Distance from text: ${distance} points`);
});
```

---

### fontName

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the name of the font for the dropped capital letter.

#### Examples

**Example**: Get the font name of the dropped capital letter in the first paragraph and display it in the console

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const dropCap = firstParagraph.dropCap;
    
    dropCap.load("fontName");
    await context.sync();
    
    console.log("Drop cap font name: " + dropCap.fontName);
});
```

---

### linesToDrop

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the height (in lines) of the dropped capital letter.

#### Examples

**Example**: Get the height in lines of the first paragraph's drop cap and display it in the console.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const dropCap = firstParagraph.dropCap;
    
    dropCap.load("linesToDrop");
    await context.sync();
    
    console.log(`Drop cap height: ${dropCap.linesToDrop} lines`);
});
```

---

### position

**Type:** `Word.DropPosition | "None" | "Normal" | "Margin"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the position of the dropped capital letter.

#### Examples

**Example**: Check if a paragraph has a drop cap applied and display its position (None, Normal, or Margin) in the console.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const dropCap = paragraph.dropCap;
    
    dropCap.load("position");
    await context.sync();
    
    console.log("Drop cap position: " + dropCap.position);
});
```

---

## Methods

### clear

**Kind:** `delete`

Removes the dropped capital letter formatting.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove the drop cap formatting from the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the drop cap object
    const dropCap = firstParagraph.dropCap;
    
    // Remove the drop cap formatting
    dropCap.clear();
    
    await context.sync();
});
```

---

### enable

**Kind:** `create`

Formats the first character in the specified paragraph as a dropped capital letter.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Format the first character of the first paragraph in the document as a drop cap

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the drop cap object for the paragraph
    const dropCap = firstParagraph.dropCap;
    
    // Enable the drop cap formatting
    dropCap.enable();
    
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
  - `options`: `Word.Interfaces.DropCapLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.DropCap`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.DropCap`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.DropCap`

#### Examples

**Example**: Load and display the text and number of lines for the first paragraph's drop cap

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const dropCap = firstParagraph.dropCap;
    
    // Load drop cap properties
    dropCap.load("text, linesToDrop");
    
    await context.sync();
    
    console.log("Drop cap text: " + dropCap.text);
    console.log("Lines to drop: " + dropCap.linesToDrop);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DropCap object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DropCapData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.DropCapData`

#### Examples

**Example**: Get the drop cap properties of the first paragraph and log them as a JSON string to the console.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the drop cap object
    const dropCap = firstParagraph.dropCap;
    
    // Load the drop cap properties
    dropCap.load("type,linesToDrop");
    
    await context.sync();
    
    // Convert the drop cap object to a plain JavaScript object
    const dropCapData = dropCap.toJSON();
    
    // Log the JSON representation
    console.log("Drop Cap Properties:", JSON.stringify(dropCapData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.DropCap`

#### Examples

**Example**: Track a drop cap object across multiple sync calls to modify its properties without getting an InvalidObjectPath error

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const dropCap = firstParagraph.dropCap;
    
    // Track the drop cap object for use across multiple sync calls
    dropCap.track();
    
    dropCap.load("linesDropped");
    await context.sync();
    
    // Now we can safely modify the drop cap in subsequent operations
    console.log(`Current lines dropped: ${dropCap.linesDropped}`);
    dropCap.linesDropped = 3;
    
    await context.sync();
    
    // Untrack when done to free up memory
    dropCap.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.DropCap`

#### Examples

**Example**: Track a drop cap object to modify its properties, then untrack it to release memory after the modifications are complete.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph's drop cap
    const paragraph = context.document.body.paragraphs.getFirst();
    const dropCap = paragraph.dropCap;
    
    // Track the drop cap object for property modifications
    dropCap.track();
    
    // Load and modify drop cap properties
    dropCap.load("linesDropped");
    await context.sync();
    
    // Make changes to the drop cap
    dropCap.linesDropped = 3;
    await context.sync();
    
    // Untrack the drop cap to release memory
    dropCap.untrack();
    await context.sync();
    
    console.log("Drop cap modified and untracked successfully");
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
