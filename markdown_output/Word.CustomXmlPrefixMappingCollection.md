# Word.CustomXmlPrefixMappingCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of Word.CustomXmlPrefixMapping objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlPrefixMappingCollection to verify the connection between the add-in and Word before performing operations on custom XML prefix mappings.

```typescript
await Word.run(async (context) => {
    // Get the custom XML prefix mappings collection
    const prefixMappings = context.document.customXmlParts.getByNamespace("http://example.com/schema").getOnlyItem().prefixMappings;
    
    // Access the request context associated with the collection
    const requestContext = prefixMappings.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (requestContext === context) {
        console.log("Request context is properly connected to the Word application");
    }
    
    // Load the collection to ensure the context connection is working
    prefixMappings.load("items");
    await context.sync();
    
    console.log(`Successfully accessed ${prefixMappings.items.length} prefix mappings through the context`);
});
```

---

### items

**Type:** `Word.CustomXmlPrefixMapping[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve and log all custom XML prefix mappings (namespace prefixes and URIs) from the document.

```typescript
await Word.run(async (context) => {
    // Get the custom XML prefix mappings collection
    const prefixMappings = context.document.customXmlParts.getByNamespace("http://example.com/schema").getOnlyItem().namespaceManager.prefixMappings;
    
    // Load the items property
    prefixMappings.load("items");
    
    await context.sync();
    
    // Access the loaded items array
    const mappingsArray = prefixMappings.items;
    
    // Iterate through the items
    for (let i = 0; i < mappingsArray.length; i++) {
        const mapping = mappingsArray[i];
        mapping.load("prefix, namespaceUri");
    }
    
    await context.sync();
    
    // Log each prefix mapping
    mappingsArray.forEach(mapping => {
        console.log(`Prefix: ${mapping.prefix}, URI: ${mapping.namespaceUri}`);
    });
});
```

---

## Methods

### addNamespace

**Kind:** `configure`

Adds a custom namespace/prefix mapping to use when querying an item.

#### Signature

**Parameters:**
- `prefix`: `string` (required)
  The prefix to associate with the namespace.
- `namespaceUri`: `string` (required)
  The namespace URI to map.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Add a custom XML namespace mapping with prefix "contoso" for the namespace URI "http://schemas.contoso.com/2024/data" to enable querying custom XML parts

```typescript
await Word.run(async (context) => {
    const prefixMappings = context.document.customXmlParts.getByNamespace("http://schemas.contoso.com/2024/data").getOnlyItem().namespaceManager;
    
    // Add a custom namespace/prefix mapping
    prefixMappings.addNamespace("contoso", "http://schemas.contoso.com/2024/data");
    
    await context.sync();
    
    console.log("Custom namespace mapping added successfully");
});
```

---

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Get the count of custom XML prefix mappings in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const prefixMappings = customXmlParts.items[0].namespaceManager.customPrefixes;
        prefixMappings.load("count");
        await context.sync();
        
        const count = prefixMappings.getCount();
        await context.sync();
        
        console.log(`Number of custom XML prefix mappings: ${count.value}`);
    } else {
        console.log("No custom XML parts found in the document.");
    }
});
```

---

### getItem

**Kind:** `read`

Returns a CustomXmlPrefixMapping object that represents the specified item in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a paragraph object.

**Returns:** `Word.CustomXmlPrefixMapping`

#### Examples

**Example**: Get the first prefix mapping from a custom XML part and display its namespace URI in the console.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the prefix mappings collection
        const prefixMappings = customXmlPart.namespaceManager.customPrefixMappings;
        prefixMappings.load("items");
        await context.sync();
        
        // Get the first prefix mapping using getItem()
        if (prefixMappings.items.length > 0) {
            const firstMapping = prefixMappings.getItem(0);
            firstMapping.load("prefix, namespaceUri");
            await context.sync();
            
            console.log(`Prefix: ${firstMapping.prefix}, Namespace URI: ${firstMapping.namespaceUri}`);
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
  - `options`: `Word.Interfaces.CustomXmlPrefixMappingCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlPrefixMappingCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlPrefixMappingCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlPrefixMappingCollection`

#### Examples

**Example**: Load and display all custom XML namespace prefix mappings in the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML prefix mappings collection
    const prefixMappings = context.document.customXmlParts.getByNamespace("http://example.com/schema").getOnlyItem().namespaceManager.prefixMappings;
    
    // Load the prefix and namespaceUri properties for all mappings
    prefixMappings.load("items/prefix, items/namespaceUri");
    
    await context.sync();
    
    // Display the loaded prefix mappings
    console.log("Custom XML Prefix Mappings:");
    for (let i = 0; i < prefixMappings.items.length; i++) {
        const mapping = prefixMappings.items[i];
        console.log(`Prefix: ${mapping.prefix}, Namespace: ${mapping.namespaceUri}`);
    }
});
```

---

### lookupNamespace

**Kind:** `read`

Gets the namespace corresponding to the specified prefix.

#### Signature

**Parameters:**
- `prefix`: `string` (required)
  The prefix to look up.

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Look up and display the namespace URI associated with the "contoso" prefix in the document's custom XML prefix mappings.

```typescript
await Word.run(async (context) => {
    const prefixMappings = context.document.customXmlParts.getByNamespace("http://schemas.contoso.com/example").getOnlyItem().namespaceManager.customPrefixMappings;
    
    const namespace = prefixMappings.lookupNamespace("contoso");
    
    context.load(namespace);
    await context.sync();
    
    console.log(`Namespace for prefix 'contoso': ${namespace.value}`);
});
```

---

### lookupPrefix

**Kind:** `read`

Gets the prefix corresponding to the specified namespace.

#### Signature

**Parameters:**
- `namespaceUri`: `string` (required)
  The namespace URI to look up.

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Look up and display the namespace prefix for a specific XML namespace URI in the document's custom XML mappings

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const prefixMappings = customXmlParts.items[0].namespaceManager.customPrefixes;
        prefixMappings.load("items");
        await context.sync();

        // Look up the prefix for a specific namespace URI
        const namespaceUri = "http://www.example.com/schema";
        const prefix = prefixMappings.lookupPrefix(namespaceUri);
        
        console.log(`Prefix for namespace '${namespaceUri}': ${prefix}`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlPrefixMappingCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlPrefixMappingCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlPrefixMappingCollectionData`

#### Examples

**Example**: Export custom XML prefix mappings to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts, "items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the namespace prefix mappings from the first custom XML part
        const prefixMappings = customXmlParts.items[0].namespaceManager.customPrefixes;
        context.load(prefixMappings, "items");
        await context.sync();

        // Convert the prefix mappings collection to JSON
        const jsonData = prefixMappings.toJSON();
        
        // Log or use the JSON data
        console.log("Custom XML Prefix Mappings:", JSON.stringify(jsonData, null, 2));
        
        // The jsonData object contains an "items" array with prefix mapping details
        jsonData.items.forEach((mapping: any) => {
            console.log(`Prefix: ${mapping.prefix}, Namespace: ${mapping.namespaceUri}`);
        });
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlPrefixMappingCollection`

#### Examples

**Example**: Track a custom XML prefix mapping collection across multiple sync calls to maintain object references when checking and modifying namespace mappings in a document.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    // Get the first custom XML part
    const customXmlPart = customXmlParts.items[0];
    
    if (customXmlPart) {
        const prefixMappings = customXmlPart.namespaceManager.customPrefixes;
        
        // Track the collection to use it across multiple sync calls
        prefixMappings.track();
        
        prefixMappings.load("items");
        await context.sync();
        
        // Now we can safely use the collection after sync
        console.log(`Found ${prefixMappings.items.length} prefix mappings`);
        
        // Perform additional operations
        await context.sync();
        
        // Still safe to access the tracked collection
        prefixMappings.items.forEach(mapping => {
            console.log(`Prefix: ${mapping.prefix}, Namespace: ${mapping.namespaceUri}`);
        });
        
        // Untrack when done
        prefixMappings.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlPrefixMappingCollection`

#### Examples

**Example**: Track a custom XML prefix mapping collection during operations, then untrack it to release memory after use

```typescript
await Word.run(async (context) => {
    // Get the custom XML prefix mappings collection
    const customXmlParts = context.document.customXmlParts;
    const customXmlPart = customXmlParts.getByNamespace("http://example.com/schema").getOnlyItem();
    const prefixMappings = customXmlPart.namespaceManager.customPrefixMappings;
    
    // Track the collection for performance monitoring
    prefixMappings.track();
    
    // Load properties
    prefixMappings.load("items");
    await context.sync();
    
    // Perform operations with the collection
    console.log(`Found ${prefixMappings.items.length} prefix mappings`);
    
    // Untrack the collection to release memory
    prefixMappings.untrack();
    await context.sync();
    
    console.log("Memory released for prefix mappings collection");
});
```

---

## Source

- /en-us/javascript/api/word
