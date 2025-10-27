# Word.CustomXmlSchemaCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of [Word.CustomXmlSchema](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschema) objects attached to a data stream.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlSchemaCollection to verify the connection between the add-in and Word, then use it to sync changes after loading schema properties.

```typescript
await Word.run(async (context) => {
    // Get a custom XML part and its schema collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const schemas = customXmlParts.items[0].getSchemas();
        schemas.load("items");
        
        // Access the request context from the schema collection
        const schemaContext = schemas.context;
        
        // Use the context to sync and load schema data
        await schemaContext.sync();
        
        console.log(`Schema collection has ${schemas.items.length} schemas`);
        console.log(`Context is connected: ${schemaContext !== null}`);
    }
});
```

---

### items

**Type:** `Word.CustomXmlSchema[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Display the namespace URIs of all custom XML schemas attached to the document in the console.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    const firstPart = customXmlParts.getByNamespace("http://example.com/data").getFirstOrNullObject();
    
    firstPart.load("schemaCollection");
    await context.sync();
    
    if (!firstPart.isNullObject) {
        const schemaCollection = firstPart.schemaCollection;
        schemaCollection.load("items");
        await context.sync();
        
        // Access the items property to get all schemas
        const schemas = schemaCollection.items;
        
        console.log(`Found ${schemas.length} schema(s):`);
        schemas.forEach((schema, index) => {
            schema.load("namespaceUri");
        });
        
        await context.sync();
        
        schemas.forEach((schema, index) => {
            console.log(`Schema ${index + 1}: ${schema.namespaceUri}`);
        });
    }
});
```

---

## Methods

### add

**Kind:** `create`

Adds one or more schemas to the schema collection that can then be added to a stream in the data store and to the schema library.

#### Signature

**Parameters:**
- `options`: `Word.CustomXmlAddSchemaOptions` (optional)
  Optional. The options that define the schema to be added.

**Returns:** `Word.CustomXmlSchema`

#### Examples

**Example**: Add a custom XML schema to the schema collection for validating custom XML parts in the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML part collection
    const customXmlParts = context.document.customXmlParts;
    const customXmlPart = customXmlParts.getByNamespace("http://schemas.contoso.com/customer")[0];
    
    // Get the schema collection for this custom XML part
    const schemaCollection = customXmlPart.schemaCollection;
    
    // Define the XML schema to add
    const schemaXml = `<?xml version="1.0" encoding="utf-8"?>
        <xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema" 
                   targetNamespace="http://schemas.contoso.com/customer">
            <xs:element name="customer">
                <xs:complexType>
                    <xs:sequence>
                        <xs:element name="name" type="xs:string"/>
                        <xs:element name="email" type="xs:string"/>
                    </xs:sequence>
                </xs:complexType>
            </xs:element>
        </xs:schema>`;
    
    // Add the schema to the collection
    schemaCollection.add(schemaXml);
    
    await context.sync();
    console.log("Schema added to the collection successfully");
});
```

---

### addCollection

**Kind:** `create`

Adds an existing schema collection to the current schema collection.

#### Signature

**Parameters:**
- `schemaCollection`: `Word.CustomXmlSchemaCollection` (required)
  The schema collection to add.

**Returns:** `Word.CustomXmlSchemaCollection`

#### Examples

**Example**: Add all schemas from a secondary custom XML part's schema collection to the primary custom XML part's schema collection

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Assume we have at least two custom XML parts
    if (customXmlParts.items.length >= 2) {
        const primaryPart = customXmlParts.items[0];
        const secondaryPart = customXmlParts.items[1];

        // Get the schema collections
        const primarySchemas = primaryPart.schemaCollection;
        const secondarySchemas = secondaryPart.schemaCollection;

        // Add all schemas from secondary part to primary part
        primarySchemas.addCollection(secondarySchemas);

        await context.sync();
        console.log("Schema collection added successfully");
    }
});
```

---

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Get and display the count of custom XML schemas attached to the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML schema collection from the document
    const schemaCollection = context.document.customXmlSchemaCollection;
    
    // Get the count of schemas in the collection
    const count = schemaCollection.getCount();
    
    // Sync to get the actual count value
    await context.sync();
    
    // Display the count
    console.log(`Number of custom XML schemas: ${count.value}`);
});
```

---

### getItem

**Kind:** `read`

Returns a `CustomXmlSchema` object that represents the specified item in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a paragraph object.

**Returns:** `Word.CustomXmlSchema`

#### Examples

**Example**: Get the first custom XML schema from the collection and display its namespace URI in the console.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the schema collection for the first custom XML part
        const schemaCollection = customXmlParts.items[0].schemaCollection;
        schemaCollection.load("items");
        await context.sync();

        if (schemaCollection.items.length > 0) {
            // Get the first schema from the collection using getItem()
            const firstSchema = schemaCollection.getItem(0);
            firstSchema.load("namespaceUri");
            await context.sync();

            console.log("First schema namespace URI: " + firstSchema.namespaceUri);
        }
    }
});
```

---

### getNamespaceUri

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Get the namespace URI of the first custom XML schema in the collection

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the schemas collection for the first custom XML part
        const schemas = customXmlParts.items[0].getSchemas();
        schemas.load("items");
        await context.sync();

        if (schemas.items.length > 0) {
            // Get the namespace URI of the first schema
            const namespaceUri = schemas.items[0].getNamespaceUri();
            await context.sync();

            console.log("Schema namespace URI: " + namespaceUri.value);
        } else {
            console.log("No schemas found in the custom XML part.");
        }
    } else {
        console.log("No custom XML parts found in the document.");
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
  - `options`: `Word.Interfaces.CustomXmlSchemaCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlSchemaCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlSchemaCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlSchemaCollection`

#### Examples

**Example**: Load and display the namespace URIs of all custom XML schemas attached to the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    // Get the schema collection from the first custom XML part
    if (customXmlParts.items.length > 0) {
        const schemaCollection = customXmlParts.items[0].getSchemaCollection();
        
        // Load the schema collection properties
        schemaCollection.load("items");
        await context.sync();

        // Display the namespace URIs of all schemas
        console.log(`Found ${schemaCollection.items.length} schema(s)`);
        schemaCollection.items.forEach((schema, index) => {
            schema.load("namespaceUri");
        });
        await context.sync();

        schemaCollection.items.forEach((schema, index) => {
            console.log(`Schema ${index + 1}: ${schema.namespaceUri}`);
        });
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlSchemaCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlSchemaCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlSchemaCollectionData`

#### Examples

**Example**: Serialize a custom XML schema collection to JSON format for logging or debugging purposes

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    
    // Get the first custom XML part (if it exists)
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const firstPart = customXmlParts.items[0];
        
        // Get the schema collection for this custom XML part
        const schemaCollection = firstPart.schemaCollection;
        schemaCollection.load("items");
        await context.sync();
        
        // Convert the schema collection to a plain JavaScript object
        const schemaCollectionData = schemaCollection.toJSON();
        
        // Log the serialized data
        console.log("Schema Collection Data:", JSON.stringify(schemaCollectionData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlSchemaCollection`

#### Examples

**Example**: Track a CustomXmlSchemaCollection object to maintain its reference across multiple sync calls when working with custom XML parts in a document.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the schema collection for this custom XML part
        const schemaCollection = customXmlPart.getSchemas();
        
        // Track the collection to use it across multiple sync calls
        schemaCollection.track();
        
        // Load properties
        schemaCollection.load("items");
        await context.sync();
        
        // Now we can safely use the collection across syncs
        console.log(`Schema count: ${schemaCollection.items.length}`);
        
        // Perform additional operations...
        await context.sync();
        
        // Still valid to use the tracked collection
        console.log(`Still accessible: ${schemaCollection.items.length} schemas`);
        
        // Untrack when done to free up memory
        schemaCollection.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlSchemaCollection`

#### Examples

**Example**: Load a custom XML schema collection, use it to verify schemas are present, then untrack it to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    // Get the schema collection from the first custom XML part
    if (customXmlParts.items.length > 0) {
        const schemaCollection = customXmlParts.items[0].getSchemas();
        schemaCollection.load("items");
        await context.sync();
        
        // Use the schema collection
        console.log(`Found ${schemaCollection.items.length} schemas`);
        
        // Untrack the schema collection to release memory
        schemaCollection.untrack();
        await context.sync();
        
        console.log("Schema collection untracked and memory released");
    }
});
```

---

### validate

**Kind:** `read`

Specifies whether the schemas in the schema collection are valid (conforms to the syntactic rules of XML and the rules for a specified vocabulary).

#### Signature

**Returns:** `OfficeExtension.ClientResult<boolean>`

#### Examples

**Example**: Validate all XML schemas attached to a custom XML part to ensure they conform to XML syntactic rules and vocabulary specifications.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the schema collection for this XML part
        const schemaCollection = customXmlPart.schemaCollection;
        
        // Validate the schemas in the collection
        const isValid = schemaCollection.validate();
        
        await context.sync();
        
        console.log(`Schema collection is valid: ${isValid.value}`);
    } else {
        console.log("No custom XML parts found in the document.");
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlschemacollection
