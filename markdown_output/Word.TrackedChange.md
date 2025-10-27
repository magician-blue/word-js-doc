# TrackedChange

**Package:** `word`

**API Set:** WordApi 1.6

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a tracked change in a Word document.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the next (second) tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  await context.sync();

  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const nextTrackedChange: Word.TrackedChange = trackedChange.getNext();
  await context.sync();

  nextTrackedChange.load(["author", "date", "text", "type"]);
  await context.sync();

  console.log(nextTrackedChange);
});
```

## Properties

### author

**Type:** `string`

**Since:** 1.6

Gets the author of the tracked change.

#### Examples

**Example**: Get and display the author name of the first tracked change in the document

```typescript
await Word.run(async (context) => {
    const trackedChanges = context.document.body.getTrackedChanges();
    trackedChanges.load("items");
    
    await context.sync();
    
    if (trackedChanges.items.length > 0) {
        const firstChange = trackedChanges.items[0];
        firstChange.load("author");
        
        await context.sync();
        
        console.log("Author of the first tracked change: " + firstChange.author);
    } else {
        console.log("No tracked changes found in the document.");
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a tracked change to load and read its properties, then display the author who made the change.

```typescript
await Word.run(async (context) => {
    // Get the first tracked change in the document
    const trackedChanges = context.document.getTrackedChanges();
    trackedChanges.load("items");
    await context.sync();
    
    if (trackedChanges.items.length > 0) {
        const firstChange = trackedChanges.items[0];
        
        // Access the context property to use the same request context
        const requestContext = firstChange.context;
        
        // Use the context to load properties
        firstChange.load("author, text");
        await requestContext.sync();
        
        console.log(`Change made by: ${firstChange.author}`);
        console.log(`Changed text: ${firstChange.text}`);
    }
});
```

---

### date

**Type:** `Date`

**Since:** 1.6

Gets the date of the tracked change.

#### Examples

**Example**: Display the date when each tracked change was made in the document by showing an alert with the dates of all tracked changes.

```typescript
await Word.run(async (context) => {
    const trackedChanges = context.document.getTrackedChanges();
    trackedChanges.load("items");
    
    await context.sync();
    
    const dates: string[] = [];
    for (let i = 0; i < trackedChanges.items.length; i++) {
        const change = trackedChanges.items[i];
        change.load("date");
    }
    
    await context.sync();
    
    for (let i = 0; i < trackedChanges.items.length; i++) {
        const changeDate = trackedChanges.items[i].date;
        dates.push(changeDate.toLocaleString());
    }
    
    console.log("Tracked change dates: " + dates.join(", "));
});
```

---

### text

**Type:** `string`

**Since:** 1.6

Gets the text of the tracked change.

#### Examples

**Example**: Get and display the text content of the first tracked change in the document

```typescript
await Word.run(async (context) => {
    const trackedChanges = context.document.getTrackedChanges();
    trackedChanges.load("items");
    
    await context.sync();
    
    if (trackedChanges.items.length > 0) {
        const firstChange = trackedChanges.items[0];
        firstChange.load("text");
        
        await context.sync();
        
        console.log("Tracked change text: " + firstChange.text);
    }
});
```

---

### type

**Type:** `Word.TrackedChangeType | "None" | "Added" | "Deleted" | "Formatted"`

**Since:** 1.6

Gets the type of the tracked change.

#### Examples

**Example**: Get all tracked changes in the document and display an alert showing the type of each tracked change (whether it's an addition, deletion, or formatting change).

```typescript
await Word.run(async (context) => {
    const trackedChanges = context.document.getTrackedChanges();
    trackedChanges.load("items");
    
    await context.sync();
    
    let changeTypes = [];
    for (let i = 0; i < trackedChanges.items.length; i++) {
        const change = trackedChanges.items[i];
        change.load("type");
    }
    
    await context.sync();
    
    for (let i = 0; i < trackedChanges.items.length; i++) {
        const change = trackedChanges.items[i];
        changeTypes.push(`Change ${i + 1}: ${change.type}`);
    }
    
    console.log(changeTypes.join("\n"));
});
```

---

## Methods

### accept

**Kind:** `write`

Accepts the tracked change.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Accept the first tracked change in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Accepts the first tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  trackedChange.load();
  await context.sync();

  console.log("First tracked change:", trackedChange);
  trackedChange.accept();
  console.log("Accepted the first tracked change.");
});
```

---

### getNext

**Kind:** `read`

Gets the next tracked change. Throws an ItemNotFound error if this tracked change is the last one.

#### Signature

**Returns:** `Word.TrackedChange`

#### Examples

**Example**: Retrieve the second tracked change in the document body and load its author, date, text, and type properties.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the next (second) tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  await context.sync();

  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const nextTrackedChange: Word.TrackedChange = trackedChange.getNext();
  await context.sync();

  nextTrackedChange.load(["author", "date", "text", "type"]);
  await context.sync();

  console.log(nextTrackedChange);
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next tracked change. If this tracked change is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.TrackedChange`

#### Examples

**Example**: Iterate through all tracked changes in the document and log their text content to the console.

```typescript
await Word.run(async (context) => {
    const trackedChanges = context.document.getTrackedChanges();
    trackedChanges.load("items");
    await context.sync();

    if (trackedChanges.items.length > 0) {
        let currentChange = trackedChanges.items[0];
        currentChange.load("text");
        await context.sync();

        while (!currentChange.isNullObject) {
            console.log("Tracked change text: " + currentChange.text);
            
            currentChange = currentChange.getNextOrNullObject();
            currentChange.load("text, isNullObject");
            await context.sync();
        }
        
        console.log("Finished iterating through all tracked changes.");
    } else {
        console.log("No tracked changes found in the document.");
    }
});
```

---

### getRange

**Kind:** `read`

Gets the range of the tracked change.

#### Signature

**Parameters:**
- `rangeLocation`: `Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | "Whole" | "Start" | "End"` (optional)

**Returns:** `Word.Range`

#### Examples

**Example**: Retrieve and display the text content of the range associated with the first tracked change in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets the range of the first tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  await context.sync();

  const range: Word.Range = trackedChange.getRange();
  range.load();
  await context.sync();

  console.log("range.text: " + range.text);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.TrackedChangeLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.TrackedChange`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.TrackedChange`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.TrackedChange`

#### Examples

**Example**: Get and display the author and date of the first tracked change in the document

```typescript
await Word.run(async (context) => {
    // Get the first tracked change in the document
    const trackedChanges = context.document.body.trackedChanges;
    trackedChanges.load("items");
    await context.sync();
    
    if (trackedChanges.items.length > 0) {
        const firstChange = trackedChanges.items[0];
        
        // Load specific properties of the tracked change
        firstChange.load("author, date, type");
        await context.sync();
        
        // Display the tracked change information
        console.log(`Author: ${firstChange.author}`);
        console.log(`Date: ${firstChange.date}`);
        console.log(`Type: ${firstChange.type}`);
    } else {
        console.log("No tracked changes found in the document.");
    }
});
```

---

### reject

**Kind:** `write`

Rejects the tracked change.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Reject the first tracked change in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Rejects the first tracked change.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  const trackedChange: Word.TrackedChange = trackedChanges.getFirst();
  trackedChange.load();
  await context.sync();

  console.log("First tracked change:", trackedChange);
  trackedChange.reject();
  console.log("Rejected the first tracked change.");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TrackedChange object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TrackedChangeData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.TrackedChangeData`

#### Examples

**Example**: Serialize a tracked change to JSON format to log or store its properties for debugging purposes

```typescript
await Word.run(async (context) => {
    // Get the first tracked change in the document
    const trackedChanges = context.document.getTrackedChanges();
    trackedChanges.load("items");
    await context.sync();

    if (trackedChanges.items.length > 0) {
        const firstChange = trackedChanges.items[0];
        
        // Load properties we want to serialize
        firstChange.load("type,author,date,text");
        await context.sync();

        // Convert the TrackedChange object to a plain JavaScript object
        const changeData = firstChange.toJSON();
        
        // Now we can use the plain object (e.g., log it, store it, etc.)
        console.log("Tracked change data:", JSON.stringify(changeData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.TrackedChange`

#### Examples

**Example**: Track a tracked change object to monitor its properties across multiple sync calls and ensure it remains valid when accessed outside the initial batch operation.

```typescript
await Word.run(async (context) => {
    // Get the first tracked change in the document
    const trackedChanges = context.document.body.getTrackedChanges();
    context.load(trackedChanges);
    await context.sync();
    
    if (trackedChanges.items.length > 0) {
        const firstChange = trackedChanges.items[0];
        
        // Track the object for use across multiple sync calls
        firstChange.track();
        
        // Load properties
        context.load(firstChange, "text, author, date, type");
        await context.sync();
        
        console.log(`Change by ${firstChange.author}: ${firstChange.text}`);
        
        // Can safely use the tracked object in subsequent operations
        await context.sync();
        
        // Access the tracked object again without errors
        console.log(`Change type: ${firstChange.type}`);
        
        // Untrack when done
        firstChange.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.TrackedChange`

#### Examples

**Example**: Get a tracked change from the document, perform operations with it, then untrack it to free memory resources when done.

```typescript
await Word.run(async (context) => {
    // Get the first tracked change in the document
    const trackedChanges = context.document.body.getTrackedChanges();
    context.trackedObjects.add(trackedChanges);
    await context.sync();
    
    if (trackedChanges.items.length > 0) {
        const firstChange = trackedChanges.items[0];
        
        // Track the object for use
        context.trackedObjects.add(firstChange);
        
        // Load properties to work with
        firstChange.load("text,type");
        await context.sync();
        
        // Use the tracked change
        console.log(`Change type: ${firstChange.type}, Text: ${firstChange.text}`);
        
        // Untrack when done to free memory
        firstChange.untrack();
        await context.sync();
    }
    
    // Clean up the collection as well
    trackedChanges.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
