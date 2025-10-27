# Word.CustomXmlPartCollection

**Package:** `word`

**API Set:** WordApi 1.4 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains the collection of [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart) objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlPartCollection to verify the connection to the Word host application before performing operations on custom XML parts.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    
    // Access the request context associated with the collection
    const requestContext = customXmlParts.context;
    
    // Verify the context is available and connected
    if (requestContext) {
        console.log("Request context is available and connected to Word host application");
        
        // Use the context to load and sync custom XML parts
        customXmlParts.load("items");
        await context.sync();
        
        console.log(`Found ${customXmlParts.items.length} custom XML parts`);
    }
});
```

---

### items

**Type:** `Word.CustomXmlPart[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve and log the IDs of all custom XML parts in the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    
    // Load the items property to access the array of custom XML parts
    customXmlParts.load("items");
    
    await context.sync();
    
    // Access the items array and log each custom XML part's ID
    console.log(`Total custom XML parts: ${customXmlParts.items.length}`);
    
    customXmlParts.items.forEach((part, index) => {
        part.load("id");
    });
    
    await context.sync();
    
    customXmlParts.items.forEach((part, index) => {
        console.log(`Custom XML Part ${index + 1} ID: ${part.id}`);
    });
});
```

---

## Methods

### add

**Kind:** `create`

Adds a new custom XML part to the document.

#### Signature

**Parameters:**
- `xml`: `string` (required)
  XML content. Must be a valid XML fragment.

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Add a custom XML part with namespace to the document, retrieve its ID and namespace URI, and store the ID in document settings for later use.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Adds a custom XML part.
// If you want to populate the CustomXml.namespaceUri property, you must include the xmlns attribute.
await Word.run(async (context) => {
  const originalXml =
    "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
  const customXmlPart = context.document.customXmlParts.add(originalXml);
  customXmlPart.load(["id", "namespaceUri"]);
  const xmlBlob = customXmlPart.getXml();

  await context.sync();

  const readableXml = addLineBreaksToXML(xmlBlob.value);
  console.log(`Added custom XML part with namespace URI ${customXmlPart.namespaceUri}:`, readableXml);

  // Store the XML part's ID in a setting so the ID is available to other functions.
  const settings: Word.SettingCollection = context.document.settings;
  settings.add("ContosoReviewXmlPartIdNS", customXmlPart.id);

  await context.sync();
});
```

**Example**: Add a custom XML part containing reviewer data to the document and store its ID in the document settings for later retrieval.

```typescript
// Adds a custom XML part.
await Word.run(async (context) => {
  const originalXml =
    "<Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
  const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.add(originalXml);
  customXmlPart.load("id");
  const xmlBlob = customXmlPart.getXml();

  await context.sync();

  const readableXml = addLineBreaksToXML(xmlBlob.value);
  console.log("Added custom XML part:", readableXml);

  // Store the XML part's ID in a setting so the ID is available to other functions.
  const settings: Word.SettingCollection = context.document.settings;
  settings.add("ContosoReviewXmlPartId", customXmlPart.id);

  await context.sync();
});
```

---

### getByNamespace

**Kind:** `read`

Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.

#### Signature

**Parameters:**
- `namespaceUri`: `string` (required)
  The namespace URI.

**Returns:** `Word.CustomXmlPartScopedCollection`

#### Examples

**Example**: Retrieve all custom XML parts from the document that match the namespace URI 'http://schemas.contoso.com/review/1.0' and display the count of matching parts.

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

---

### getCount

**Kind:** `read`

Gets the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Get the total count of custom XML parts in the document and display it in the console

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    
    // Get the count of custom XML parts
    const count = customXmlParts.getCount();
    
    // Sync to get the count value
    await context.sync();
    
    // Display the count
    console.log(`Total custom XML parts: ${count.value}`);
});
```

---

### getItem

**Kind:** `read`

Gets a custom XML part based on its ID.

#### Signature

**Parameters:**
- `id`: `string` (required)
  ID or index of the custom XML part to be retrieved.

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Retrieve a custom XML part by its stored ID and query it for elements matching an XPath expression with namespace mapping.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Queries a custom XML part for elements matching the search terms.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartIdNS").load("value");

  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
    const xpathToQueryFor = "/contoso:Reviewers";
    const clientResult = customXmlPart.query(xpathToQueryFor, {
      contoso: "http://schemas.contoso.com/review/1.0"
    });

    await context.sync();

    console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
    for (let i = 0; i < clientResult.value.length; i++) {
      console.log(clientResult.value[i]);
    }
  } else {
    console.warn("Didn't find custom XML part to query.");
  }
});
```

**Example**: Retrieve a custom XML part by its stored ID and query it for all Reviewer elements, logging the matching results to the console.

```typescript
// Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Queries a custom XML part for elements matching the search terms.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
    const xpathToQueryFor = "/Reviewers/Reviewer";
    const clientResult = customXmlPart.query(xpathToQueryFor, {
      contoso: "http://schemas.contoso.com/review/1.0"
    });

    await context.sync();

    console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
    for (let i = 0; i < clientResult.value.length; i++) {
      console.log(clientResult.value[i]);
    }
  } else {
    console.warn("Didn't find custom XML part to query.");
  }
});
```

---

### getItemOrNullObject

**Kind:** `read`

Gets a custom XML part based on its ID. If the CustomXmlPart doesn't exist, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Parameters:**
- `id`: `string` (required)
  ID of the object to be retrieved.

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Check if a custom XML part with a specific ID exists in the document and display its namespace URI if found, or log a message if it doesn't exist.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    const targetId = "{12345678-1234-1234-1234-123456789012}";
    
    const customXmlPart = customXmlParts.getItemOrNullObject(targetId);
    customXmlPart.load("isNullObject, namespaceUri");
    
    await context.sync();
    
    if (customXmlPart.isNullObject) {
        console.log(`Custom XML part with ID ${targetId} does not exist.`);
    } else {
        console.log(`Found custom XML part with namespace: ${customXmlPart.namespaceUri}`);
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
  - `options`: `Word.Interfaces.CustomXmlPartCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlPartCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlPartCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlPartCollection`

#### Examples

**Example**: Load and display the namespace URIs of all custom XML parts in the document

```typescript
await Word.run(async (context) => {
    // Get the collection of custom XML parts
    const customXmlParts = context.document.customXmlParts;
    
    // Load the items and their namespaceUri properties
    customXmlParts.load("items/namespaceUri");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the namespace URIs
    console.log(`Found ${customXmlParts.items.length} custom XML parts:`);
    customXmlParts.items.forEach((part, index) => {
        console.log(`Part ${index + 1}: ${part.namespaceUri}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlPartCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlPartCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlPartCollectionData`

#### Examples

**Example**: Serialize the custom XML parts collection to JSON format for logging or debugging purposes

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    
    // Load the properties we want to serialize
    customXmlParts.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const jsonData = customXmlParts.toJSON();
    
    // Log or use the serialized data
    console.log("Custom XML Parts:", JSON.stringify(jsonData, null, 2));
    console.log("Number of custom XML parts:", jsonData.items.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlPartCollection`

#### Examples

**Example**: Track a custom XML part collection across multiple sync calls to safely access and manipulate custom XML parts without encountering "InvalidObjectPath" errors.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    customXmlParts.track();
    
    // Add a new custom XML part
    const xmlString = "<root><item>Sample Data</item></root>";
    customXmlParts.add(xmlString);
    await context.sync();
    
    // Safe to access the tracked collection after sync
    customXmlParts.load("items/id");
    await context.sync();
    
    console.log(`Total custom XML parts: ${customXmlParts.items.length}`);
    
    // Untrack when done
    customXmlParts.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlPartCollection`

#### Examples

**Example**: Add a custom XML part to the document, use it to perform operations, then untrack it to release memory after the work is complete.

```typescript
await Word.run(async (context) => {
    // Add a custom XML part to the collection
    const customXmlParts = context.document.customXmlParts;
    const xmlString = "<root><item>Sample Data</item></root>";
    const customXmlPart = customXmlParts.add(xmlString);
    
    // Track the collection for use
    customXmlParts.load("items");
    await context.sync();
    
    // Perform operations with the collection
    console.log(`Total custom XML parts: ${customXmlParts.items.length}`);
    
    // Untrack the collection to release memory when done
    customXmlParts.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
