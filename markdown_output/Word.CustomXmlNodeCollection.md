# Word.CustomXmlNodeCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode) objects representing the XML nodes in a document.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlNodeCollection to verify the connection to the Word host application and log context information.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts from the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get nodes from the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.getXml();
        await context.sync();
        
        // Access a collection of custom XML nodes (example assumes nodes exist)
        const nodeCollection = xmlPart.query("//*");
        nodeCollection.load("items");
        await context.sync();
        
        // Access the context property from the CustomXmlNodeCollection
        const requestContext = nodeCollection.context;
        
        // Verify the context is valid and connected
        console.log("Context is connected:", requestContext !== null);
        console.log("Context type:", typeof requestContext);
        
        // The context can be used for additional operations
        await requestContext.sync();
    }
});
```

---

### items

**Type:** `Word.CustomXmlNode[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve and log all custom XML nodes from a custom XML part to the console

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get all XML nodes from the custom XML part
        const xmlNodes = customXmlPart.getXml();
        const nodeCollection = customXmlPart.query("*");
        nodeCollection.load("items");
        await context.sync();
        
        // Access the loaded child items using the items property
        const nodes = nodeCollection.items;
        
        // Log information about each node
        console.log(`Found ${nodes.length} XML nodes`);
        for (let i = 0; i < nodes.length; i++) {
            nodes[i].load("nodeType, baseName");
        }
        await context.sync();
        
        nodes.forEach((node, index) => {
            console.log(`Node ${index}: ${node.baseName} (Type: ${node.nodeType})`);
        });
    }
});
```

---

## Methods

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Get the count of custom XML nodes in a custom XML part and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the collection of XML nodes
        const xmlNodes = customXmlPart.getXmlNodes();
        
        // Get the count of nodes in the collection
        const count = xmlNodes.getCount();
        count.load();
        await context.sync();
        
        console.log(`Number of custom XML nodes: ${count.value}`);
    } else {
        console.log("No custom XML parts found in the document.");
    }
});
```

---

### getItem

**Kind:** `read`

Returns a `CustomXmlNode` object that represents the specified item in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a CustomXMLNode object.

**Returns:** `Word.CustomXmlNode`

#### Examples

**Example**: Get the second custom XML node from a collection and display its base name in the console.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const customXmlPart = customXmlParts.items[0];
        
        // Get the collection of XML nodes
        const xmlNodes = customXmlPart.getXml();
        const xmlNodeCollection = customXmlPart.query("//*");
        xmlNodeCollection.load("items");
        await context.sync();

        if (xmlNodeCollection.items.length >= 2) {
            // Get the second node (index 1) from the collection
            const secondNode = xmlNodeCollection.getItem(1);
            secondNode.load("baseName");
            await context.sync();

            console.log("Second node base name: " + secondNode.baseName);
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
  - `options`: `Word.Interfaces.CustomXmlNodeCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlNodeCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlNodeCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlNodeCollection`

#### Examples

**Example**: Load and display the base names of all custom XML nodes in the document's first custom XML part

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const firstPart = customXmlParts.items[0];
        const xmlNodes = firstPart.getXmlNodes();
        
        // Load properties of the custom XML node collection
        xmlNodes.load("items/baseName");
        await context.sync();
        
        // Display the base names of all nodes
        console.log("Custom XML Nodes:");
        xmlNodes.items.forEach((node, index) => {
            console.log(`Node ${index}: ${node.baseName}`);
        });
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. Whereas the original `Word.CustomXmlNodeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlNodeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlNodeCollectionData`

#### Examples

**Example**: Serialize a collection of custom XML nodes to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get all nodes in the custom XML part
        const xmlNodes = customXmlPart.getXml();
        const nodeCollection = customXmlPart.query("//node()");
        nodeCollection.load("items");
        await context.sync();

        // Convert the collection to a plain JavaScript object
        const jsonData = nodeCollection.toJSON();
        
        // Log the serialized data
        console.log("Custom XML Nodes as JSON:", JSON.stringify(jsonData, null, 2));
        
        // The jsonData object contains an "items" array with node properties
        console.log(`Number of nodes: ${jsonData.items.length}`);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlNodeCollection`

#### Examples

**Example**: Track a custom XML node collection across multiple sync calls to prevent "InvalidObjectPath" errors when accessing the collection after document changes.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts from the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    // Get the first custom XML part (if it exists)
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        const xmlNodes = customXmlPart.getXml();
        
        // Track the collection to use it across sync calls
        xmlNodes.track();
        
        await context.sync();
        
        // Now safe to use the collection after sync
        // Perform operations that might change the document
        context.document.body.insertParagraph("New content", Word.InsertLocation.end);
        await context.sync();
        
        // The tracked collection remains valid
        xmlNodes.load("items");
        await context.sync();
        
        console.log(`XML nodes count: ${xmlNodes.items.length}`);
        
        // Untrack when done
        xmlNodes.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlNodeCollection`

#### Examples

**Example**: Query custom XML nodes from a document, use them to perform operations, then untrack the collection to free memory.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts from the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get nodes from the first custom XML part
        const xmlPart = customXmlParts.items[0];
        const xmlNodes = xmlPart.getXml();
        
        // Load the collection for use
        xmlNodes.load("items");
        await context.sync();

        // Perform operations with the nodes
        console.log(`Found ${xmlNodes.items.length} XML nodes`);

        // Untrack the collection to release memory
        xmlNodes.untrack();
        await context.sync();
    }
});
```

---

## Source

- /en-us/javascript/api/word
