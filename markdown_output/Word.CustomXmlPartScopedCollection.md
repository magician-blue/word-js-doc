# Word.CustomXmlPartScopedCollection

**Package:** `word`

**API Set:** WordApi 1.4 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains the collection of [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart) objects with a specific namespace.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Gets the custom XML parts with the specified namespace URI.
await Word.run(async (context) => {
  const namespaceUri = "http://schemas.contoso.com/review/1.0";
  console.log(`Specified namespace URI: ${namespaceUri}`);
  const scopedCustomXmlParts: Word.CustomXmlPartScopedCollection =
    context.document.customXmlParts.getByNamespace(namespaceUri);
  scopedCustomXmlParts.load("items");
  await context.sync();

  console.log(`Number of custom XML parts found with this namespace: ${!scopedCustomXmlParts.items ? 0 : scopedCustomXmlParts.items.length}`);
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlPartScopedCollection to verify the connection between the add-in and Word before performing operations on custom XML parts with a specific namespace.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts with a specific namespace
    const namespace = "http://schemas.contoso.com/customer";
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    
    // Access the request context from the collection
    const requestContext = customXmlParts.context;
    
    // Verify the context is available and matches the Word.run context
    if (requestContext === context) {
        console.log("Request context is properly connected to Word");
    }
    
    // Use the context to load and sync the collection
    customXmlParts.load("items");
    await context.sync();
    
    console.log(`Found ${customXmlParts.items.length} custom XML parts in namespace: ${namespace}`);
});
```

---

### items

**Type:** `Word.CustomXmlPart[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve all custom XML parts with a specific namespace and log their IDs to the console.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts with a specific namespace
    const customXmlParts = context.document.customXmlParts.getByNamespace("http://schemas.contoso.com/customer");
    
    // Load the items property to access the collection
    customXmlParts.load("items");
    
    await context.sync();
    
    // Access the items array and log each part's ID
    console.log(`Found ${customXmlParts.items.length} custom XML parts`);
    
    for (let i = 0; i < customXmlParts.items.length; i++) {
        const part = customXmlParts.items[i];
        part.load("id");
        await context.sync();
        console.log(`Custom XML Part ${i + 1} ID: ${part.id}`);
    }
});
```

---

## Methods

### getCount

**Kind:** `read`

Gets the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Get the count of custom XML parts with a specific namespace and display it in the console.

```typescript
await Word.run(async (context) => {
    // Define the namespace to search for
    const namespace = "http://schemas.contoso.com/customer";
    
    // Get custom XML parts scoped to the specific namespace
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    
    // Get the count of custom XML parts in this namespace
    const count = customXmlParts.getCount();
    
    // Load and sync the count value
    await context.sync();
    
    // Display the count
    console.log(`Number of custom XML parts with namespace '${namespace}': ${count.value}`);
});
```

---

### getItem

**Kind:** `read`

Gets a custom XML part based on its ID.

#### Signature

**Parameters:**
- `id`: `string` (required)
  ID of the custom XML part to be retrieved.

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Get a custom XML part with a specific ID from a scoped collection filtered by namespace and read its XML content.

```typescript
await Word.run(async (context) => {
    // Define the namespace and the ID of the custom XML part to retrieve
    const namespace = "http://schemas.contoso.com/customer";
    const customXmlPartId = "{12345678-1234-1234-1234-123456789012}";
    
    // Get the scoped collection of custom XML parts with the specified namespace
    const scopedCollection = context.document.customXmlParts.getByNamespace(namespace);
    
    // Get the specific custom XML part by its ID
    const customXmlPart = scopedCollection.getItem(customXmlPartId);
    
    // Load the XML content
    customXmlPart.load("xml");
    
    await context.sync();
    
    // Use the retrieved custom XML part
    console.log("Custom XML Part Content:", customXmlPart.xml);
});
```

---

### getItemOrNullObject

**Kind:** `read`

Gets a custom XML part based on its ID. If the CustomXmlPart doesn't exist in the collection, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Parameters:**
- `id`: `string` (required)
  Required. ID of the object to be retrieved.

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Check if a custom XML part with a specific ID exists in a namespace-scoped collection and read its content if found, otherwise log that it doesn't exist.

```typescript
await Word.run(async (context) => {
    const namespace = "http://schemas.contoso.com/customer";
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    const targetId = "{12345678-1234-1234-1234-123456789012}";
    
    const customXmlPart = customXmlParts.getItemOrNullObject(targetId);
    customXmlPart.load("isNullObject, xml");
    
    await context.sync();
    
    if (customXmlPart.isNullObject) {
        console.log(`Custom XML part with ID ${targetId} not found in namespace ${namespace}`);
    } else {
        console.log("Custom XML part found:");
        console.log(customXmlPart.xml);
    }
});
```

---

### getOnlyItem

**Kind:** `read`

If the collection contains exactly one item, this method returns it. Otherwise, this method produces an error.

#### Signature

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Get the single custom XML part with a specific namespace and display its ID, assuming exactly one part exists with that namespace.

```typescript
await Word.run(async (context) => {
    const namespace = "http://schemas.contoso.com/customer";
    
    // Get all custom XML parts with the specific namespace
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    
    // Get the only item (will error if collection doesn't contain exactly one item)
    const customXmlPart = customXmlParts.getOnlyItem();
    customXmlPart.load("id");
    
    await context.sync();
    
    console.log(`Custom XML Part ID: ${customXmlPart.id}`);
});
```

---

### getOnlyItemOrNullObject

**Kind:** `read`

If the collection contains exactly one item, this method returns it. Otherwise, this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Check if there is exactly one custom XML part with a specific namespace and retrieve its ID, or handle the case when there are zero or multiple parts.

```typescript
await Word.run(async (context) => {
    const namespace = "http://schemas.contoso.com/customer";
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    const singlePart = customXmlParts.getOnlyItemOrNullObject();
    
    singlePart.load("id, isNullObject");
    await context.sync();
    
    if (singlePart.isNullObject) {
        console.log("There are zero or multiple custom XML parts with this namespace.");
    } else {
        console.log(`Found exactly one custom XML part with ID: ${singlePart.id}`);
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
  - `options`: `Word.Interfaces.CustomXmlPartScopedCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlPartScopedCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlPartScopedCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlPartScopedCollection`

#### Examples

**Example**: Load and display the namespace URIs of all custom XML parts that match a specific namespace in the document.

```typescript
await Word.run(async (context) => {
    const namespace = "http://schemas.contoso.com/customer";
    
    // Get custom XML parts with the specified namespace
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    
    // Load the namespace property of the scoped collection items
    customXmlParts.load("items/namespaceUri");
    
    await context.sync();
    
    // Display the namespace URIs
    console.log(`Found ${customXmlParts.items.length} custom XML parts with namespace: ${namespace}`);
    customXmlParts.items.forEach((part, index) => {
        console.log(`Part ${index + 1} namespace: ${part.namespaceUri}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlPartScopedCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlPartScopedCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlPartScopedCollectionData`

#### Examples

**Example**: Retrieve custom XML parts with a specific namespace and serialize them to JSON format for logging or external storage.

```typescript
await Word.run(async (context) => {
    // Get custom XML parts with a specific namespace
    const namespace = "http://schemas.contoso.com/customer";
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    
    // Load the properties we want to serialize
    customXmlParts.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const jsonData = customXmlParts.toJSON();
    
    // Now you can use the plain object (e.g., log it, send to server, etc.)
    console.log("Custom XML Parts as JSON:", JSON.stringify(jsonData, null, 2));
    console.log("Number of parts found:", jsonData.items.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlPartScopedCollection`

#### Examples

**Example**: Track a custom XML part with a specific namespace across multiple sync calls to safely access its properties without getting "InvalidObjectPath" errors.

```typescript
await Word.run(async (context) => {
    const namespace = "http://schemas.contoso.com/customer";
    
    // Get custom XML parts with the specific namespace
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    
    // Track the collection to use it across multiple sync calls
    customXmlParts.track();
    
    // Load the collection
    customXmlParts.load("items");
    await context.sync();
    
    // Now we can safely use the collection in subsequent operations
    console.log(`Found ${customXmlParts.items.length} custom XML parts`);
    
    // Perform additional operations after sync
    if (customXmlParts.items.length > 0) {
        const firstPart = customXmlParts.items[0];
        firstPart.load("id");
        await context.sync();
        
        console.log(`First part ID: ${firstPart.id}`);
    }
    
    // Untrack when done
    customXmlParts.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlPartScopedCollection`

#### Examples

**Example**: Load custom XML parts with a specific namespace, use them to perform operations, then untrack the collection to free memory.

```typescript
await Word.run(async (context) => {
    const namespace = "http://schemas.contoso.com/customer";
    
    // Get custom XML parts with the specific namespace
    const customXmlParts = context.document.customXmlParts.getByNamespace(namespace);
    customXmlParts.load("items");
    
    await context.sync();
    
    // Use the collection (e.g., log the count)
    console.log(`Found ${customXmlParts.items.length} custom XML parts`);
    
    // Untrack the collection to release memory
    customXmlParts.untrack();
    
    await context.sync();
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml
