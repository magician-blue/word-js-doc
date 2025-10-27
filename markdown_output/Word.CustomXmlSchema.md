# Word.CustomXmlSchema

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a schema in a [Word.CustomXmlSchemaCollection](/en-us/javascript/api/word/word.customxmlschemacollection) object.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlSchema object to verify the connection between the add-in and Word application.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part and its schemas
        const xmlPart = customXmlParts.items[0];
        const schemas = xmlPart.getSchemas();
        schemas.load("items");
        await context.sync();

        if (schemas.items.length > 0) {
            const schema = schemas.items[0];
            
            // Access the context property to verify connection
            const schemaContext = schema.context;
            console.log("Schema context is connected:", schemaContext !== null);
            console.log("Context matches document context:", schemaContext === context);
        }
    }
});
```

---

### location

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the location of the schema on a computer.

#### Examples

**Example**: Get and display the file system location of the first custom XML schema in the document.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const schemas = customXmlParts.items[0].getSchemas();
        schemas.load("items");
        await context.sync();

        if (schemas.items.length > 0) {
            const firstSchema = schemas.items[0];
            firstSchema.load("location");
            await context.sync();

            console.log("Schema location: " + firstSchema.location);
        }
    }
});
```

---

### namespaceUri

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the unique address identifier for the namespace of the CustomXmlSchema object.

#### Examples

**Example**: Retrieve and display the namespace URI of the first custom XML schema in the document's schema collection.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const schemas = customXmlParts.items[0].getSchemas();
        schemas.load("items");
        await context.sync();

        if (schemas.items.length > 0) {
            const firstSchema = schemas.items[0];
            firstSchema.load("namespaceUri");
            await context.sync();

            console.log("Schema namespace URI: " + firstSchema.namespaceUri);
        }
    }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes this schema from the [Word.CustomXmlSchemaCollection](/en-us/javascript/api/word/word.customxmlschemacollection) object.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a specific custom XML schema from the document's schema collection

```typescript
await Word.run(async (context) => {
    // Get the custom XML part collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the schema collection from the first custom XML part
        const schemas = customXmlParts.items[0].getSchemas();
        schemas.load("items");
        await context.sync();

        if (schemas.items.length > 0) {
            // Delete the first schema from the collection
            schemas.items[0].delete();
            await context.sync();
            
            console.log("Schema deleted successfully");
        }
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
  - `options`: `Word.Interfaces.CustomXmlSchemaLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlSchema`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlSchema`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlSchema`

#### Examples

**Example**: Load and display the namespace URI and location properties of the first custom XML schema in the document

```typescript
await Word.run(async (context) => {
    // Get the first custom XML schema from the document
    const schemas = context.document.customXmlParts.getByNamespace("http://example.com/schema").getFirst().schemas;
    const firstSchema = schemas.getFirst();
    
    // Load the schema properties
    firstSchema.load("namespaceUri, location");
    
    await context.sync();
    
    // Display the loaded properties
    console.log("Schema Namespace URI: " + firstSchema.namespaceUri);
    console.log("Schema Location: " + firstSchema.location);
});
```

---

### reload

**Kind:** `load`

Reloads the schema from a file.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Refresh a custom XML schema to ensure it has the latest data from its source file

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the first custom XML part (if it exists)
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the schemas collection for this XML part
        const schemas = customXmlPart.schemas;
        schemas.load("items");
        await context.sync();

        // Reload the first schema to get the latest data from its file
        if (schemas.items.length > 0) {
            const schema = schemas.items[0];
            schema.reload();
            await context.sync();

            console.log("Schema reloaded successfully");
        }
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlSchema object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlSchemaData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlSchemaData`

#### Examples

**Example**: Serialize a custom XML schema to JSON format for logging or debugging purposes

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts, "items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const customXmlPart = customXmlParts.items[0];
        
        // Get the schemas collection for this part
        const schemas = customXmlPart.getSchemas();
        context.load(schemas, "items");
        await context.sync();

        if (schemas.items.length > 0) {
            // Get the first schema
            const schema = schemas.items[0];
            context.load(schema);
            await context.sync();

            // Convert the schema to a plain JavaScript object
            const schemaJSON = schema.toJSON();
            
            // Log the serialized schema data
            console.log("Schema JSON:", JSON.stringify(schemaJSON, null, 2));
        }
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlSchema`

#### Examples

**Example**: Track a custom XML schema object across multiple sync calls to prevent InvalidObjectPath errors when accessing its properties after document changes.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part's schema
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const schemas = customXmlParts.items[0].getSchemas();
        schemas.load("items");
        await context.sync();
        
        if (schemas.items.length > 0) {
            const schema = schemas.items[0];
            
            // Track the schema object to use it across multiple sync calls
            schema.track();
            
            // Load properties
            schema.load("namespaceUri, uri");
            await context.sync();
            
            // Can safely access properties after sync because object is tracked
            console.log("Schema namespace: " + schema.namespaceUri);
            console.log("Schema URI: " + schema.uri);
            
            // Perform other operations...
            await context.sync();
            
            // Still safe to access tracked object properties
            console.log("Still accessible: " + schema.namespaceUri);
            
            // Untrack when done
            schema.untrack();
        }
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlSchema`

#### Examples

**Example**: Get a custom XML schema from a part, use it to verify the schema namespace, then untrack it to free memory after you're done working with it.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the schema collection for this part
        const schemas = customXmlPart.getSchemas();
        schemas.load("items");
        await context.sync();
        
        if (schemas.items.length > 0) {
            const schema = schemas.items[0];
            schema.load("namespaceUri");
            await context.sync();
            
            // Use the schema
            console.log("Schema namespace: " + schema.namespaceUri);
            
            // Untrack the schema object to free memory
            schema.untrack();
            await context.sync();
        }
    }
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
