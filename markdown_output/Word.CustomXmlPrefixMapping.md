# Word.CustomXmlPrefixMapping

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a CustomXmlPrefixMapping object.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlPrefixMapping object to verify the connection between the add-in and Word application before performing operations.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts);
    await context.sync();
    
    // Get the first custom XML part (if exists)
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        const namespaceMappings = customXmlPart.namespaceManager.prefixMappings;
        context.load(namespaceMappings);
        await context.sync();
        
        // Access the context property from the first prefix mapping
        if (namespaceMappings.items.length > 0) {
            const prefixMapping = namespaceMappings.items[0];
            const requestContext = prefixMapping.context;
            
            // Verify the context is valid and connected
            console.log("Request context is connected:", requestContext !== null);
            console.log("Context type:", typeof requestContext);
        }
    }
});
```

---

### namespaceUri

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the unique address identifier for the namespace of the CustomXmlPrefixMapping object.

#### Examples

**Example**: Get the namespace URI from a custom XML prefix mapping to verify the namespace associated with a specific prefix in the document.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts, "items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const customXmlPart = customXmlParts.items[0];
        
        // Get the namespace prefix mappings
        const prefixMappings = customXmlPart.namespaceManager.customPrefixes;
        context.load(prefixMappings, "items");
        await context.sync();

        if (prefixMappings.items.length > 0) {
            // Get the namespace URI from the first prefix mapping
            const mapping = prefixMappings.items[0];
            context.load(mapping, "namespaceUri");
            await context.sync();

            console.log("Namespace URI: " + mapping.namespaceUri);
        }
    }
});
```

---

### prefix

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the prefix for the CustomXmlPrefixMapping object.

#### Examples

**Example**: Get the prefix of the first custom XML prefix mapping in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        const prefixMappings = customXmlPart.namespaceManager.prefixMappings;
        prefixMappings.load("items");
        await context.sync();

        if (prefixMappings.items.length > 0) {
            const prefixMapping = prefixMappings.items[0];
            prefixMapping.load("prefix");
            await context.sync();

            console.log("Prefix: " + prefixMapping.prefix);
        }
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
  - `options`: `Word.Interfaces.CustomXmlPrefixMappingLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlPrefixMapping`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlPrefixMapping`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlPrefixMapping`

#### Examples

**Example**: Load and read the namespace URI and prefix properties of a custom XML prefix mapping from a custom XML part in the document.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    
    // Get the namespace manager and its prefix mappings
    const namespaceManager = customXmlPart.namespaceManager;
    const prefixMappings = namespaceManager.customPrefixMappings;
    prefixMappings.load("items");
    await context.sync();
    
    // Load properties of the first prefix mapping
    const prefixMapping = prefixMappings.items[0];
    prefixMapping.load("prefix, namespaceUri");
    await context.sync();
    
    // Now you can read the loaded properties
    console.log("Prefix: " + prefixMapping.prefix);
    console.log("Namespace URI: " + prefixMapping.namespaceUri);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.CustomXmlPrefixMapping object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlPrefixMappingData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlPrefixMappingData`

#### Examples

**Example**: Retrieve a custom XML prefix mapping and serialize it to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the namespace manager and prefix mappings
        const namespaceManager = customXmlPart.namespaceManager;
        const prefixMappings = namespaceManager.customPrefixMappings;
        prefixMappings.load("items");
        await context.sync();
        
        if (prefixMappings.items.length > 0) {
            const mapping = prefixMappings.items[0];
            mapping.load("prefix, namespaceUri");
            await context.sync();
            
            // Convert the mapping to a plain JSON object
            const jsonObject = mapping.toJSON();
            
            // Log or use the JSON representation
            console.log("Prefix Mapping as JSON:", JSON.stringify(jsonObject, null, 2));
        }
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlPrefixMapping`

#### Examples

**Example**: Track a custom XML prefix mapping object across multiple sync calls to maintain its reference while modifying document properties

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        const prefixMappings = customXmlPart.namespaceManager.customPrefixMappings;
        prefixMappings.load("items");
        await context.sync();
        
        if (prefixMappings.items.length > 0) {
            const prefixMapping = prefixMappings.items[0];
            
            // Track the prefix mapping object to use it across sync calls
            prefixMapping.track();
            
            // Load properties
            prefixMapping.load("prefix,uri");
            await context.sync();
            
            console.log(`Prefix: ${prefixMapping.prefix}, URI: ${prefixMapping.uri}`);
            
            // Can safely use the object in subsequent sync calls
            await context.sync();
            
            // Untrack when done
            prefixMapping.untrack();
        }
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlPrefixMapping`

#### Examples

**Example**: Release memory for a tracked CustomXmlPrefixMapping object after retrieving and using its namespace prefix information.

```typescript
await Word.run(async (context) => {
    // Get a custom XML part and its prefix mapping
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        const prefixMappings = xmlPart.namespaceManager.prefixMappings;
        prefixMappings.load("items");
        await context.sync();
        
        if (prefixMappings.items.length > 0) {
            const mapping = prefixMappings.items[0];
            
            // Track the object to monitor it
            mapping.track();
            
            // Load and use the mapping properties
            mapping.load("prefix, namespaceUri");
            await context.sync();
            
            console.log(`Prefix: ${mapping.prefix}, Namespace: ${mapping.namespaceUri}`);
            
            // Release the memory after we're done using it
            mapping.untrack();
            await context.sync();
        }
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlprefixmapping
