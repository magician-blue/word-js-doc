# Word.NoteItem

**Package:** `word`

**API Set:** WordApi 1.5

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a footnote or endnote.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the text of the referenced footnote.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items/body");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const footnoteBody: Word.Range = footnotes.items[mark].body.getRange();
  footnoteBody.load("text");
  await context.sync();

  console.log(`Text of footnote ${referenceNumber}: ${footnoteBody.text}`);
});
```

## Properties

### body

**Type:** `Word.Body`

**Since:** WordApi 1.5

Represents the body object of the note item. It's the portion of the text within the footnote or endnote.

#### Examples

**Example**: Retrieve and display the text content of a specific footnote in the document based on its reference number.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the text of the referenced footnote.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items/body");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const footnoteBody: Word.Range = footnotes.items[mark].body.getRange();
  footnoteBody.load("text");
  await context.sync();

  console.log(`Text of footnote ${referenceNumber}: ${footnoteBody.text}`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a footnote to verify the connection between the add-in and Word before performing operations on the note item.

```typescript
await Word.run(async (context) => {
    // Get the first footnote in the document
    const footnotes = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();
    
    if (footnotes.items.length > 0) {
        const firstFootnote = footnotes.items[0];
        
        // Access the request context from the footnote
        const noteContext = firstFootnote.context;
        
        // Use the context to verify connection and load properties
        firstFootnote.load("body/text");
        await noteContext.sync();
        
        console.log("Footnote text: " + firstFootnote.body.text);
        console.log("Context connection verified");
    }
});
```

---

### reference

**Type:** `Word.Range`

**Since:** WordApi 1.5

Represents a footnote or endnote reference in the main document.

#### Examples

**Example**: Select a footnote's reference mark in the document body based on a user-provided reference number.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Selects the footnote's reference mark in the document body.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items/reference");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const item: Word.NoteItem = footnotes.items[mark];
  const reference: Word.Range = item.reference;
  reference.select();
  await context.sync();

  console.log(`Reference ${referenceNumber} is selected.`);
});
```

---

### type

**Type:** `Word.NoteItemType | "Footnote" | "Endnote"`

**Since:** WordApi 1.5

Represents the note item type: footnote or endnote.

#### Examples

**Example**: Retrieve and display the note type and body type of a specific footnote based on its reference number.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the referenced note's item type and body type, which are both "Footnote".
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const item: Word.NoteItem = footnotes.items[mark];
  console.log(`Note type of footnote ${referenceNumber}: ${item.type}`);

  item.body.load("type");
  await context.sync();

  console.log(`Body type of note: ${item.body.type}`);
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the note item.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a specific footnote from the document body based on a user-provided reference number.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Deletes this referenced footnote.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  footnotes.items[mark].delete();
  await context.sync();

  console.log("Footnote deleted.");
});
```

---

### getNext

**Kind:** `read`

Gets the next note item of the same type. Throws an `ItemNotFound` error if this note item is the last one.

#### Signature

**Returns:** `Word.NoteItem`

#### Examples

**Example**: Select the footnote that comes after the footnote at the specified reference number in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Selects the next footnote in the document body.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items/reference");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const reference: Word.Range = footnotes.items[mark].getNext().reference;
  reference.select();
  console.log("Selected is the next footnote: " + (mark + 2));
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next note item of the same type. If this note item is the last one, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.NoteItem`

#### Examples

**Example**: Iterate through all footnotes in the document and log their reference numbers and text content to the console.

```typescript
await Word.run(async (context) => {
    // Get the first footnote in the document
    const footnotes = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();
    
    if (footnotes.items.length > 0) {
        let currentNote = footnotes.items[0];
        currentNote.load("reference, body/text");
        await context.sync();
        
        // Process the first footnote
        console.log(`Footnote ${currentNote.reference}: ${currentNote.body.text}`);
        
        // Iterate through remaining footnotes using getNextOrNullObject
        while (true) {
            let nextNote = currentNote.getNextOrNullObject();
            nextNote.load("isNullObject, reference, body/text");
            await context.sync();
            
            if (nextNote.isNullObject) {
                break;
            }
            
            console.log(`Footnote ${nextNote.reference}: ${nextNote.body.text}`);
            currentNote = nextNote;
        }
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.NoteItemLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.NoteItem`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.NoteItem`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.NoteItem`

#### Examples

**Example**: Load and display the reference text of the first footnote in the document

```typescript
await Word.run(async (context) => {
    // Get the first footnote in the document
    const footnotes = context.document.body.footnotes;
    const firstFootnote = footnotes.getFirst();
    
    // Load the reference property of the footnote
    firstFootnote.load("reference");
    
    await context.sync();
    
    // Display the footnote reference text
    console.log("Footnote reference: " + firstFootnote.reference);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.NoteItemUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.NoteItem` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of the first footnote in the document, setting its reference mark to a custom symbol and making it a superscript

```typescript
await Word.run(async (context) => {
    const footnotes = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();
    
    if (footnotes.items.length > 0) {
        const firstFootnote = footnotes.items[0];
        
        // Set multiple properties at once
        firstFootnote.set({
            reference: "*",
            type: Word.NoteItemType.footnote
        });
        
        await context.sync();
        console.log("Footnote properties updated successfully");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.NoteItem` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.NoteItemData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.NoteItemData`

#### Examples

**Example**: Serialize a footnote's properties to a plain JavaScript object and log it to the console for debugging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first footnote in the document
    const footnotes = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();
    
    if (footnotes.items.length > 0) {
        const firstFootnote = footnotes.items[0];
        
        // Load properties you want to serialize
        firstFootnote.load("type, body/text");
        await context.sync();
        
        // Convert the NoteItem to a plain JavaScript object
        const footnoteData = firstFootnote.toJSON();
        
        // Now you can use the plain object (e.g., log it, store it, send it)
        console.log("Footnote data:", JSON.stringify(footnoteData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.NoteItem`

#### Examples

**Example**: Track a footnote object to maintain its reference across multiple sync calls while modifying its properties

```typescript
await Word.run(async (context) => {
    // Get the first footnote in the document
    const footnotes = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();
    
    if (footnotes.items.length > 0) {
        const firstFootnote = footnotes.items[0];
        
        // Track the footnote to use it across multiple sync calls
        firstFootnote.track();
        
        // First sync - load the footnote's body text
        firstFootnote.body.load("text");
        await context.sync();
        
        console.log("Original footnote text:", firstFootnote.body.text);
        
        // Second sync - modify the footnote
        firstFootnote.body.insertText(" [Updated]", Word.InsertLocation.end);
        await context.sync();
        
        // Untrack when done to free up memory
        firstFootnote.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.NoteItem`

#### Examples

**Example**: Access a footnote, perform operations on it, then release it from memory tracking to improve performance

```typescript
await Word.run(async (context) => {
    // Get the first footnote in the document
    const footnotes = context.document.body.footnotes;
    footnotes.load("items");
    await context.sync();
    
    if (footnotes.items.length > 0) {
        const firstFootnote = footnotes.items[0];
        
        // Track the footnote for operations
        firstFootnote.track();
        firstFootnote.body.load("text");
        await context.sync();
        
        // Perform operations with the footnote
        console.log("Footnote text: " + firstFootnote.body.text);
        
        // Untrack the footnote to release memory
        firstFootnote.untrack();
        await context.sync();
    }
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml
