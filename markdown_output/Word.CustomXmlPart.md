# Word.CustomXmlPart

**Package:** `word`

**API Set:** WordApi 1.4

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a custom XML part.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part.yaml

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

## Properties

### builtIn

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a value that indicates whether the CustomXmlPart is built-in.

#### Examples

**Example**: Check if a custom XML part is built-in and log the result to the console

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const firstPart = customXmlParts.items[0];
        firstPart.load("builtIn");
        
        await context.sync();
        
        console.log(`Is built-in: ${firstPart.builtIn}`);
    } else {
        console.log("No custom XML parts found in the document.");
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access a custom XML part's request context to verify the connection to the Office host application before performing operations on the XML part.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Access the request context associated with the custom XML part
        const partContext = customXmlPart.context;
        
        // Use the context to verify connection and perform operations
        customXmlPart.load("id");
        await partContext.sync();
        
        console.log("Custom XML part ID: " + customXmlPart.id);
        console.log("Request context is connected to Office host");
    }
});
```

---

### documentElement

**Type:** `Word.CustomXmlNode`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the root element of a bound region of data in the document. If the region is empty, the property returns Nothing.

#### Examples

**Example**: Get and display the tag name of the root element from a custom XML part in the document.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the root element of the custom XML part
        const documentElement = customXmlPart.documentElement;
        documentElement.load("baseName");
        await context.sync();

        // Display the root element's tag name
        console.log("Root element tag name: " + documentElement.baseName);
    } else {
        console.log("No custom XML parts found in the document.");
    }
});
```

---

### errors

**Type:** `Word.CustomXmlValidationErrorCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets a CustomXmlValidationErrorCollection object that provides access to any XML validation errors.

#### Examples

**Example**: Check if a custom XML part has any validation errors and log the error count to the console.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Access the errors collection
        const errors = customXmlPart.errors;
        errors.load("items");
        await context.sync();
        
        // Log the number of validation errors
        console.log(`Number of validation errors: ${errors.items.length}`);
        
        // Optionally, log details of each error
        errors.items.forEach((error, index) => {
            console.log(`Error ${index + 1}: ${error.reason}`);
        });
    }
});
```

---

### id

**Type:** `string`

**Since:** WordApi 1.4

Gets the ID of the custom XML part.

#### Examples

**Example**: Add a custom XML part containing reviewer names to the document and store its ID in the document settings for later retrieval.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part.yaml

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

### namespaceManager

**Type:** `Word.CustomXmlPrefixMappingCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the set of namespace prefix mappings used against the current CustomXmlPart object.

#### Examples

**Example**: Get all namespace prefix mappings from a custom XML part and log them to the console to understand what XML namespaces are defined in the document.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part in the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the namespace manager for this custom XML part
        const namespaceManager = customXmlPart.namespaceManager;
        namespaceManager.load("items");
        await context.sync();

        // Log all prefix mappings
        console.log("Namespace prefix mappings:");
        namespaceManager.items.forEach(mapping => {
            mapping.load(["prefix", "uri"]);
        });
        await context.sync();

        namespaceManager.items.forEach(mapping => {
            console.log(`Prefix: ${mapping.prefix}, URI: ${mapping.uri}`);
        });
    }
});
```

---

### namespaceUri

**Type:** `string`

**Since:** WordApi 1.4

Gets the namespace URI of the custom XML part.

#### Examples

**Example**: Retrieve and display the namespace URI from a custom XML part that was previously stored in the document settings.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Gets the namespace URI from a custom XML part.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartIdNS").load("value");

  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
    customXmlPart.load("namespaceUri");
    await context.sync();

    const namespaceUri = customXmlPart.namespaceUri;
    console.log(`Namespace URI: ${JSON.stringify(namespaceUri)}`);
  } else {
    console.warn("Didn't find custom XML part.");
  }
});
```

---

### schemaCollection

**Type:** `Word.CustomXmlSchemaCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies a CustomXmlSchemaCollection object representing the set of schemas attached to a bound region of data in the document.

#### Examples

**Example**: Get the number of schemas attached to a custom XML part and log the namespace of the first schema in the collection.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get the first custom XML part
        const customXmlPart = customXmlParts.items[0];
        
        // Access the schema collection for this custom XML part
        const schemaCollection = customXmlPart.schemaCollection;
        schemaCollection.load("items");
        await context.sync();

        // Log the count and first schema namespace if available
        console.log(`Number of schemas: ${schemaCollection.items.length}`);
        
        if (schemaCollection.items.length > 0) {
            const firstSchema = schemaCollection.items[0];
            firstSchema.load("namespaceUri");
            await context.sync();
            console.log(`First schema namespace: ${firstSchema.namespaceUri}`);
        }
    }
});
```

---

### xml

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the XML representation of the current CustomXmlPart object.

#### Examples

**Example**: Read and display the XML content from the first custom XML part in the document

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part from the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Load the xml property
        customXmlPart.load("xml");
        await context.sync();
        
        // Display the XML content
        console.log("Custom XML content:", customXmlPart.xml);
    } else {
        console.log("No custom XML parts found in the document");
    }
});
```

---

## Methods

### addNode

**Kind:** `create`

Adds a node to the XML tree.

#### Signature

**Parameters:**
- `parent`: `Word.CustomXmlNode` (required)
  The parent node to which the new node will be added.
- `options`: `Word.CustomXmlAddNodeOptions` (optional)
  Optional. The options that define the node to be added.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Add a new customer element with name and email child nodes to an existing custom XML part in the document

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlPart = context.document.customXmlParts.getByNamespace("http://schemas.contoso.com/customer")[0];
    
    // Load the XML content to work with it
    customXmlPart.load("namespaceUri");
    await context.sync();
    
    // Add a new customer node to the root
    const customerNode = customXmlPart.addNode(
        "/root",
        {
            nodeType: "Element",
            nodeName: "customer",
            nodeValue: ""
        }
    );
    
    // Add child nodes for customer details
    customXmlPart.addNode(
        "/root/customer[last()]",
        {
            nodeType: "Element",
            nodeName: "name",
            nodeValue: "John Doe"
        }
    );
    
    customXmlPart.addNode(
        "/root/customer[last()]",
        {
            nodeType: "Element",
            nodeName: "email",
            nodeValue: "john.doe@contoso.com"
        }
    );
    
    await context.sync();
    console.log("Customer node added successfully");
});
```

---

### delete

**Kind:** `delete`

Deletes the custom XML part.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a custom XML part from a Word document by retrieving its ID from settings, removing the XML part, verifying deletion, and cleaning up the associated setting.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part.yaml

// Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Deletes a custom XML part.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
  await context.sync();

  if (xmlPartIDSetting.value) {
    let customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
    const xmlBlob = customXmlPart.getXml();
    customXmlPart.delete();
    customXmlPart = context.document.customXmlParts.getItemOrNullObject(xmlPartIDSetting.value);

    await context.sync();

    if (customXmlPart.isNullObject) {
      console.log(`The XML part with the ID ${xmlPartIDSetting.value} has been deleted.`);

      // Delete the associated setting too.
      xmlPartIDSetting.delete();

      await context.sync();
    } else {
      const readableXml = addLineBreaksToXML(xmlBlob.value);
      console.error(`This is strange. The XML part with the id ${xmlPartIDSetting.value} wasn't deleted:`, readableXml);
    }
  } else {
    console.warn("Didn't find custom XML part to delete.");
  }
});

...

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Deletes a custom XML part.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartIdNS").load("value");
  await context.sync();

  if (xmlPartIDSetting.value) {
    let customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
    const xmlBlob = customXmlPart.getXml();
    customXmlPart.delete();
    customXmlPart = context.document.customXmlParts.getItemOrNullObject(xmlPartIDSetting.value);

    await context.sync();

    if (customXmlPart.isNullObject) {
      console.log(`The XML part with the ID ${xmlPartIDSetting.value} has been deleted.`);

      // Delete the associated setting too.
      xmlPartIDSetting.delete();

      await context.sync();
    } else {
      const readableXml = addLineBreaksToXML(xmlBlob.value);
      console.error(
        `This is strange. The XML part with the id ${xmlPartIDSetting.value} wasn't deleted:`,
        readableXml
      );
    }
  } else {
    console.warn("Didn't find custom XML part to delete.");
  }
});
```

---

### deleteAttribute

**Kind:** `delete`

Deletes an attribute with the given name from the element identified by xpath.

#### Signature

**Parameters:**
- `xpath`: `string` (required)
  Absolute path to the single element in XPath notation.
- `namespaceMappings`: `{ [key: string]: string; }` (required)
  An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".
- `name`: `string` (required)
  Name of the attribute.

**Returns:** `void`

#### Examples

**Example**: Delete the "status" attribute from a "document" element in a custom XML part that uses a namespace

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlPart = context.document.customXmlParts.getByNamespace("http://schemas.contoso.com/document").getOnlyItem();
    
    // Define namespace mappings for the XPath query
    const namespaceMapping = {
        prefix: "ns",
        uri: "http://schemas.contoso.com/document"
    };
    
    // Delete the "status" attribute from the document element
    customXmlPart.deleteAttribute(
        "/ns:document",
        [namespaceMapping],
        "status"
    );
    
    await context.sync();
    console.log("Attribute 'status' deleted successfully");
});
```

---

### deleteElement

**Kind:** `delete`

Deletes the element identified by xpath.

#### Signature

**Parameters:**
- `xpath`: `string` (required)
  Absolute path to the single element in XPath notation.
- `namespaceMappings`: `{ [key: string]: string; }` (required)
  An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".

**Returns:** `void`

#### Examples

**Example**: Delete a specific customer element from a custom XML part using its ID xpath

```typescript
await Word.run(async (context) => {
    // Get the custom XML part by ID or namespace
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    // Assume we have a custom XML part with customer data
    const customXmlPart = customXmlParts.items[0];
    
    // Define namespace mappings for the XML
    const namespaceMappings = {
        "ns": "http://example.com/customers"
    };
    
    // Delete the customer element with ID "12345"
    const xpath = "//ns:customer[@id='12345']";
    customXmlPart.deleteElement(xpath, namespaceMappings);
    
    await context.sync();
    console.log("Customer element deleted successfully");
});
```

---

### getXml

**Kind:** `read`

Gets the full XML content of the custom XML part.

#### Signature

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Add a custom XML part to the Word document, retrieve its XML content, and store its ID in document settings for later access.

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

...

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

### insertAttribute

**Kind:** `write`

Inserts an attribute with the given name and value to the element identified by xpath.

#### Signature

**Parameters:**
- `xpath`: `string` (required)
  Absolute path to the single element in XPath notation.
- `namespaceMappings`: `{ [key: string]: string; }` (required)
  An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".
- `name`: `string` (required)
  Name of the attribute.
- `value`: `string` (required)
  Value of the attribute.

**Returns:** `void`

#### Examples

**Example**: Add a "Nation" attribute with value "US" to the root element of a custom XML part stored in the Word document, using XPath to locate the target element.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Inserts an attribute into a custom XML part.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartIdNS").load("value");
  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);

    // The insertAttribute method inserts an attribute with the given name and value into the element identified by the xpath parameter.
    customXmlPart.insertAttribute(
      "/contoso:Reviewers",
      { contoso: "http://schemas.contoso.com/review/1.0" },
      "Nation",
      "US"
    );
    const xmlBlob = customXmlPart.getXml();
    await context.sync();

    const readableXml = addLineBreaksToXML(xmlBlob.value);
    console.log("Successfully inserted attribute:", readableXml);
  } else {
    console.warn("Didn't find custom XML part to insert attribute into.");
  }
});

...

// Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Inserts an attribute into a custom XML part.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);

    // The insertAttribute method inserts an attribute with the given name and value into the element identified by the xpath parameter.
    customXmlPart.insertAttribute("/Reviewers", { contoso: "http://schemas.contoso.com/review/1.0" }, "Nation", "US");
    const xmlBlob = customXmlPart.getXml();
    await context.sync();

    const readableXml = addLineBreaksToXML(xmlBlob.value);
    console.log("Successfully inserted attribute:", readableXml);
  } else {
    console.warn("Didn't find custom XML part to insert attribute into.");
  }
});
```

---

### insertElement

**Kind:** `write`

Inserts the given XML under the parent element identified by xpath at child position index.

#### Signature

**Parameters:**
- `xpath`: `string` (required)
  Absolute path to the single parent element in XPath notation.
- `xml`: `string` (required)
  XML content to be inserted.
- `namespaceMappings`: `{ [key: string]: string; }` (required)
  An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".
- `index`: `number` (optional)
  Zero-based position at which the new XML to be inserted. If omitted, the XML will be appended as the last child of this parent.

**Returns:** `void`

#### Examples

**Example**: Insert a new XML element at a specified position within a custom XML part stored in a Word document, using XPath to locate the parent element.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Inserts an element into a custom XML part.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartIdNS").load("value");
  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);

    // The insertElement method inserts the given XML under the parent element identified by the xpath parameter at the provided child position index.
    customXmlPart.insertElement(
      "/contoso:Reviewers",
      "<Lead>Mark</Lead>",
      { contoso: "http://schemas.contoso.com/review/1.0" },
      0
    );
    const xmlBlob = customXmlPart.getXml();
    await context.sync();

    const readableXml = addLineBreaksToXML(xmlBlob.value);
    console.log("Successfully inserted element:", readableXml);
  } else {
    console.warn("Didn't find custom XML part to insert element into.");
  }
});

...

// Original XML: <Reviewers><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Inserts an element into a custom XML part.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");
  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);

    // The insertElement method inserts the given XML under the parent element identified by the xpath parameter at the provided child position index.
    customXmlPart.insertElement(
      "/Reviewers",
      "<Lead>Mark</Lead>",
      { contoso: "http://schemas.contoso.com/review/1.0" },
      0
    );
    const xmlBlob = customXmlPart.getXml();
    await context.sync();

    const readableXml = addLineBreaksToXML(xmlBlob.value);
    console.log("Successfully inserted element:", readableXml);
  } else {
    console.warn("Didn't find custom XML part to insert element into.");
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
  - `options`: `Word.Interfaces.CustomXmlPartLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlPart`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlPart`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Load and display the namespace URI of a custom XML part in a Word document

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part from the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Load the namespaceUri property of the custom XML part
        customXmlPart.load("namespaceUri");
        await context.sync();
        
        // Display the namespace URI
        console.log("Custom XML Part Namespace URI: " + customXmlPart.namespaceUri);
    } else {
        console.log("No custom XML parts found in the document.");
    }
});
```

---

### loadXml

**Kind:** `load`

Populates the CustomXmlPart object from an XML string.

#### Signature

**Parameters:**
- `xml`: `string` (required)
  The XML string to load.

**Returns:** `OfficeExtension.ClientResult<boolean>`

#### Examples

**Example**: Replace the XML content of an existing custom XML part with new product catalog data

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    
    // Define new XML content for a product catalog
    const newXml = `<?xml version="1.0" encoding="UTF-8"?>
        <catalog>
            <product id="101">
                <name>Laptop</name>
                <price>999.99</price>
            </product>
            <product id="102">
                <name>Mouse</name>
                <price>29.99</price>
            </product>
        </catalog>`;
    
    // Load the new XML content into the custom XML part
    customXmlPart.loadXml(newXml);
    
    await context.sync();
    console.log("Custom XML part updated with new product catalog data");
});
```

---

### query

**Kind:** `read`

Queries the XML content of the custom XML part.

#### Signature

**Parameters:**
- `xpath`: `string` (required)
  An XPath query.
- `namespaceMappings`: `{ [key: string]: string; }` (required)
  An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".

**Returns:** `OfficeExtension.ClientResult<string[]>`
An array where each item represents an entry matched by the XPath query.

#### Examples

**Example**: Query a custom XML part in a Word document using XPath expressions to find matching elements, with support for both namespaced and non-namespaced XML.

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

...

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

### selectNodes

**Kind:** `read`

Selects a collection of nodes from a custom XML part.

#### Signature

**Parameters:**
- `xPath`: `string` (required)
  The XPath expression to evaluate.

**Returns:** `Word.CustomXmlNodeCollection`

#### Examples

**Example**: Select all employee nodes from a custom XML part where the department is "Engineering" and log their count to the console.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    
    // Select nodes using XPath query
    const nodes = customXmlPart.selectNodes("//employee[@department='Engineering']");
    nodes.load("items");
    await context.sync();
    
    console.log(`Found ${nodes.items.length} employees in Engineering department`);
});
```

---

### selectSingleNode

**Kind:** `read`

Selects a single node within a custom XML part matching an XPath expression.

#### Signature

**Parameters:**
- `xPath`: `string` (required)
  The XPath expression to evaluate.

**Returns:** `Word.CustomXmlNode`

#### Examples

**Example**: Select and retrieve the text content of a specific employee node from a custom XML part using an XPath expression to find the employee with ID "E001"

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlPart = context.document.customXmlParts.getByNamespace("http://company.com/employees").getOnlyItem();
    
    // Select a single node using XPath
    const node = customXmlPart.selectSingleNode("//employee[@id='E001']");
    
    // Load the node's XML content
    node.load("xml");
    
    await context.sync();
    
    // Display the selected node's XML content
    console.log("Selected employee node:", node.xml);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.CustomXmlPartUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.CustomXmlPart` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a custom XML part by setting its namespace URI and ID at the same time

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    
    // Set multiple properties at once
    customXmlPart.set({
        namespaceUri: "http://schemas.example.com/mydata",
        id: "customXmlPart1"
    });
    
    await context.sync();
    console.log("Custom XML part properties updated");
});
```

---

### setXml

**Kind:** `write`

Sets the full XML content of the custom XML part.

#### Signature

**Parameters:**
- `xml`: `string` (required)
  XML content to be set.

**Returns:** `void`

#### Examples

**Example**: Replace an existing custom XML part in a Word document with new XML content containing a different set of reviewers.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml

// Original XML: <Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>

// Replaces a custom XML part.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  const xmlPartIDSetting: Word.Setting = settings.getItemOrNullObject("ContosoReviewXmlPartIdNS").load("value");
  await context.sync();

  if (xmlPartIDSetting.value) {
    const customXmlPart: Word.CustomXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
    const originalXmlBlob = customXmlPart.getXml();
    await context.sync();

    let readableXml = addLineBreaksToXML(originalXmlBlob.value);
    console.log("Original custom XML part:", readableXml);

    // The setXml method replaces the entire XML part.
    customXmlPart.setXml(
      "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>John</Reviewer><Reviewer>Hitomi</Reviewer></Reviewers>"
    );
    const updatedXmlBlob = customXmlPart.getXml();
    await context.sync();

    readableXml = addLineBreaksToXML(updatedXmlBlob.value);
    console.log("Replaced custom XML part:", readableXml);
  } else {
    console.warn("Didn't find custom XML part to replace.");
  }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlPart object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlPartData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlPartData`

#### Examples

**Example**: Retrieve a custom XML part and serialize it to JSON format for logging or external storage purposes.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part from the document
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        customXmlPart.load("id,namespaceUri");
        await context.sync();
        
        // Convert the CustomXmlPart object to a plain JavaScript object
        const customXmlPartData = customXmlPart.toJSON();
        
        // Now you can use JSON.stringify or log the plain object
        console.log("Custom XML Part as JSON:", JSON.stringify(customXmlPartData, null, 2));
        console.log("ID:", customXmlPartData.id);
        console.log("Namespace URI:", customXmlPartData.namespaceUri);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension.clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Track a custom XML part object across multiple sync calls to prevent "InvalidObjectPath" errors when accessing its properties after synchronization.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts);
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    
    // Track the object to use it across multiple sync calls
    customXmlPart.track();
    
    // Load properties
    context.load(customXmlPart, "id, namespaceUri");
    await context.sync();
    
    // Can safely access properties after sync because object is tracked
    console.log("Custom XML Part ID: " + customXmlPart.id);
    console.log("Namespace URI: " + customXmlPart.namespaceUri);
    
    // Perform additional operations after another sync
    await context.sync();
    
    // Still valid because the object is tracked
    console.log("Still accessible: " + customXmlPart.id);
    
    // Untrack when done to release memory
    customXmlPart.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension.clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlPart`

#### Examples

**Example**: Add a custom XML part to the document, use it to store data, then untrack it to release memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Create a custom XML part with sample data
    const xmlString = "<root><item>Sample Data</item></root>";
    const customXmlPart = context.document.customXmlParts.add(xmlString);
    
    // Load the custom XML part to work with it
    customXmlPart.load("id");
    await context.sync();
    
    // Log the ID to confirm it was created
    console.log("Custom XML Part ID: " + customXmlPart.id);
    
    // Untrack the object to release memory since we're done using it
    customXmlPart.untrack();
    
    // Sync to apply the memory release
    await context.sync();
});
```

---

### updateAttribute

**Kind:** `write`

Updates the value of an attribute with the given name of the element identified by xpath.

#### Signature

**Parameters:**
- `xpath`: `string` (required)
  Absolute path to the single element in XPath notation.
- `namespaceMappings`: `{ [key: string]: string; }` (required)
  An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".
- `name`: `string` (required)
  Name of the attribute.
- `value`: `string` (required)
  New value of the attribute.

**Returns:** `void`

#### Examples

**Example**: Update the author attribute of a book element in a custom XML part to change the author name from "John Smith" to "Jane Doe"

```typescript
await Word.run(async (context) => {
    // Get the custom XML part by namespace
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    // Assume we have a custom XML part with book data
    const customXmlPart = customXmlParts.items[0];
    
    // Define namespace mappings for the XML
    const namespaceMapping = {
        "bk": "http://example.com/books"
    };
    
    // Update the author attribute of the book element
    customXmlPart.updateAttribute(
        "/bk:catalog/bk:book[@id='1']",
        namespaceMapping,
        "author",
        "Jane Doe"
    );
    
    await context.sync();
    console.log("Author attribute updated successfully");
});
```

---

### updateElement

**Kind:** `write`

Updates the XML of the element identified by xpath.

#### Signature

**Parameters:**
- `xpath`: `string` (required)
  Absolute path to the single element in XPath notation.
- `xml`: `string` (required)
  New XML content to be stored.
- `namespaceMappings`: `{ [key: string]: string; }` (required)
  An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".

**Returns:** `void`

#### Examples

**Example**: Update a customer's email address in a custom XML part by finding the customer element with a specific ID and replacing the email value

```typescript
await Word.run(async (context) => {
    // Get the custom XML part by namespace
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    // Assume we have a custom XML part with customer data
    const customXmlPart = customXmlParts.items[0];
    
    // Define namespace mappings for the XPath query
    const namespaceMappings = {
        "ns": "http://example.com/customers"
    };
    
    // Update the email element for customer with ID "12345"
    const xpath = "//ns:customer[@id='12345']/ns:email";
    const newEmailXml = "<email>newemail@example.com</email>";
    
    customXmlPart.updateElement(xpath, newEmailXml, namespaceMappings);
    
    await context.sync();
    console.log("Customer email updated successfully");
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part.yaml
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-custom-xml-part-ns.yaml
