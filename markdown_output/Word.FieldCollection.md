# Word.FieldCollection

**Package:** `word`

**API Set:** WordApi 1.4

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.Field](/en-us/javascript/api/word/word.field) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets all fields in the document body.
await Word.run(async (context) => {
  const fields: Word.FieldCollection = context.document.body.fields.load("items");

  await context.sync();

  if (fields.items.length === 0) {
    console.log("No fields in this document.");
  } else {
    fields.load(["code", "result"]);
    await context.sync();

    for (let i = 0; i < fields.items.length; i++) {
      console.log(`Field ${i + 1}'s code: ${fields.items[i].code}`, `Field ${i + 1}'s result: ${JSON.stringify(fields.items[i].result)}`);
    }
  }
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a FieldCollection to verify the connection between the add-in and Word application before performing field operations.

```typescript
await Word.run(async (context) => {
    const fields = context.document.body.fields;
    fields.load("items");
    await context.sync();
    
    // Access the request context associated with the field collection
    const requestContext = fields.context;
    
    // Verify the context is valid and connected
    if (requestContext) {
        console.log("Field collection is connected to Word application");
        console.log(`Number of fields found: ${fields.items.length}`);
    }
});
```

---

### items

**Type:** `Word.Field[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve all fields from the document body and display their code and result values in the console.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets all fields in the document body.
await Word.run(async (context) => {
  const fields: Word.FieldCollection = context.document.body.fields.load("items");

  await context.sync();

  if (fields.items.length === 0) {
    console.log("No fields in this document.");
  } else {
    fields.load(["code", "result"]);
    await context.sync();

    for (let i = 0; i < fields.items.length; i++) {
      console.log(`Field ${i + 1}'s code: ${fields.items[i].code}`, `Field ${i + 1}'s result: ${JSON.stringify(fields.items[i].result)}`);
    }
  }
});
```

---

## Methods

### getByTypes

**Kind:** `read`

Gets the Field object collection including the specified types of fields.

#### Signature

**Parameters:**
- `types`: `Word.FieldType[]` (required)
  An array of field types.

**Returns:** `Word.FieldCollection`

#### Examples

**Example**: Get all hyperlink and page reference fields from the document and display their count in the console.

```typescript
await Word.run(async (context) => {
    // Get all fields of type hyperlink and page reference
    const fields = context.document.body.fields.getByTypes([
        Word.FieldType.hyperlink,
        Word.FieldType.pageRef
    ]);
    
    // Load the count property
    fields.load("items");
    
    await context.sync();
    
    console.log(`Found ${fields.items.length} hyperlink and page reference fields`);
});
```

---

### getFirst

**Kind:** `read`

Gets the first field in this collection. Throws an `ItemNotFound` error if this collection is empty.

#### Signature

**Returns:** `Word.Field`

#### Examples

**Example**: Get the first field in the document and display its code

```typescript
await Word.run(async (context) => {
    // Get the field collection from the document body
    const fields = context.document.body.fields;
    
    // Get the first field in the collection
    const firstField = fields.getFirst();
    
    // Load the code property of the first field
    firstField.load("code");
    
    await context.sync();
    
    // Display the field code
    console.log("First field code: " + firstField.code);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first field in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.Field`

#### Examples

**Example**: Retrieve and display the properties (code, result, type, locked status, data, and kind) of the first field in the document, or indicate if no fields exist.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load(["code", "result", "locked", "type", "data", "kind"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    console.log(
      "Code of first field: " + field.code,
      "Result of first field: " + JSON.stringify(field.result),
      "Type of first field: " + field.type,
      "Is the first field locked? " + field.locked,
      "Kind of the first field: " + field.kind
    );
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
  - `options`: `Word.Interfaces.FieldCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.FieldCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.FieldCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.FieldCollection`

#### Examples

**Example**: Load and display the result text of all fields in the document

```typescript
await Word.run(async (context) => {
    // Get all fields in the document
    const fields = context.document.body.fields;
    
    // Load the result property of all fields in the collection
    fields.load("result");
    
    await context.sync();
    
    // Display the result text of each field
    for (let i = 0; i < fields.items.length; i++) {
        console.log(`Field ${i + 1} result: ${fields.items[i].result.text}`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.FieldCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FieldCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.FieldCollectionData`

#### Examples

**Example**: Serialize field collection data to JSON format for logging or data export purposes

```typescript
await Word.run(async (context) => {
    // Get all fields in the document
    const fields = context.document.body.fields;
    
    // Load properties we want to serialize
    fields.load("items/type,items/code,items/result");
    
    await context.sync();
    
    // Convert the FieldCollection to a plain JavaScript object
    const fieldsJSON = fields.toJSON();
    
    // Now you can use the plain object (e.g., log it, export it, etc.)
    console.log(JSON.stringify(fieldsJSON, null, 2));
    
    // Access the items array from the serialized data
    console.log(`Found ${fieldsJSON.items.length} fields in the document`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.FieldCollection`

#### Examples

**Example**: Track all fields in a document to monitor and update their values across multiple sync calls without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get all fields in the document
    const fields = context.document.body.fields;
    
    // Track the collection for automatic adjustment
    fields.track();
    
    // Load field properties
    fields.load("items");
    await context.sync();
    
    // First sync - work with fields
    console.log(`Found ${fields.items.length} fields`);
    
    // Perform some operations that might change the document
    context.document.body.insertParagraph("New content added", Word.InsertLocation.start);
    await context.sync();
    
    // Second sync - fields are still valid because they're tracked
    for (let i = 0; i < fields.items.length; i++) {
        fields.items[i].load("result");
    }
    await context.sync();
    
    // Access field results without InvalidObjectPath errors
    fields.items.forEach(field => {
        console.log(`Field result: ${field.result.value}`);
    });
    
    // Untrack when done
    fields.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.FieldCollection`

#### Examples

**Example**: Load all fields in a document, process them, then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    // Get the field collection from the document body
    const fields = context.document.body.fields;
    
    // Track the collection for change tracking
    fields.track();
    
    // Load properties to work with the fields
    fields.load("items");
    await context.sync();
    
    // Process the fields (e.g., log count)
    console.log(`Found ${fields.items.length} fields in the document`);
    
    // Untrack the collection to release memory
    fields.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
