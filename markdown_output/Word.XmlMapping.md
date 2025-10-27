# XmlMapping

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the XML mapping on a [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) object between custom XML and that content control. An XML mapping is a link between the text in a content control and an XML element in the custom XML data store for this document.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from an XML mapping to verify the connection between the add-in and Word, then use it to load and read the XML mapping's properties.

```typescript
await Word.run(async (context) => {
    // Get the first content control with an XML mapping
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const contentControl = contentControls.items[0];
        const xmlMapping = contentControl.xmlMapping;
        
        // Access the request context from the xmlMapping object
        const mappingContext = xmlMapping.context;
        
        // Use the context to load properties of the XML mapping
        xmlMapping.load("xpath, customXmlPart");
        await mappingContext.sync();
        
        console.log("XPath: " + xmlMapping.xpath);
        console.log("Context is connected: " + (mappingContext !== null));
    }
});
```

---

### customXmlNode

**Type:** `Word.CustomXmlNode`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a `CustomXmlNode` object that represents the custom XML node in the data store that the content control in the document maps to.

#### Examples

**Example**: Get the text content of the custom XML node that is mapped to a content control

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the XML mapping and its custom XML node
    const xmlMapping = contentControl.xmlMapping;
    const customXmlNode = xmlMapping.customXmlNode;
    
    // Load the node's text property
    customXmlNode.load("text");
    
    await context.sync();
    
    // Display the XML node's text content
    console.log("Custom XML node text: " + customXmlNode.text);
});
```

---

### customXmlPart

**Type:** `Word.CustomXmlPart`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a `CustomXmlPart` object that represents the custom XML part to which the content control in the document maps.

#### Examples

**Example**: Get the namespace URI of the custom XML part that is mapped to a content control

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the XML mapping and its custom XML part
    const xmlMapping = contentControl.xmlMapping;
    const customXmlPart = xmlMapping.customXmlPart;
    
    // Load the namespace URI property
    customXmlPart.load("namespaceUri");
    
    await context.sync();
    
    // Display the namespace URI of the mapped custom XML part
    console.log("Custom XML Part Namespace URI: " + customXmlPart.namespaceUri);
});
```

---

### isMapped

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns whether the content control in the document is mapped to an XML node in the document's XML data store.

#### Examples

**Example**: Check if a content control is mapped to XML data and display an alert message indicating the mapping status.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Load the XML mapping property
    contentControl.load("xmlMapping");
    
    await context.sync();
    
    // Check if the content control is mapped to XML
    const xmlMapping = contentControl.xmlMapping;
    xmlMapping.load("isMapped");
    
    await context.sync();
    
    // Display the mapping status
    if (xmlMapping.isMapped) {
        console.log("This content control is mapped to an XML node.");
    } else {
        console.log("This content control is not mapped to any XML node.");
    }
});
```

---

### prefixMappings

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the prefix mappings used to evaluate the XPath for the current XML mapping.

#### Examples

**Example**: Get and display the prefix mappings (namespace definitions) used in the XPath expression of a content control's XML mapping.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Load the XML mapping and its prefix mappings
    contentControl.load("xmlMapping");
    await context.sync();
    
    // Get the prefix mappings used for XPath evaluation
    const prefixMappings = contentControl.xmlMapping.prefixMappings;
    
    console.log("Prefix mappings for XPath evaluation:");
    console.log(prefixMappings);
    // Example output: "xmlns:ns0='http://example.com/schema'"
});
```

---

### xpath

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the XPath for the XML mapping, which evaluates to the currently mapped XML node.

#### Examples

**Example**: Get the XPath expression of a content control's XML mapping to verify which XML node it is currently mapped to.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Load the XML mapping and its xpath property
    contentControl.load("xmlMapping");
    const xmlMapping = contentControl.xmlMapping;
    xmlMapping.load("xpath");
    
    await context.sync();
    
    // Display the XPath expression
    console.log("XPath expression: " + xmlMapping.xpath);
    // Example output: "/root/customer/name"
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the XML mapping from the parent content control.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove the XML mapping from a content control to unlink it from custom XML data

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Load the XML mapping
    contentControl.load("xmlMapping");
    await context.sync();
    
    // Delete the XML mapping from the content control
    contentControl.xmlMapping.delete();
    
    await context.sync();
    
    console.log("XML mapping deleted from content control");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.XmlMappingLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.XmlMapping`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.XmlMapping`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `object` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.XmlMapping`

#### Examples

**Example**: Load and display the XML mapping properties of the first content control in the document to check if it's mapped to custom XML data.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the XML mapping object
    const xmlMapping = contentControl.xmlMapping;
    
    // Load properties of the XML mapping
    xmlMapping.load("isMapped, xpath, prefix, customXmlPart");
    
    // Synchronize to read the loaded properties
    await context.sync();
    
    // Display the XML mapping information
    console.log("Is Mapped: " + xmlMapping.isMapped);
    if (xmlMapping.isMapped) {
        console.log("XPath: " + xmlMapping.xpath);
        console.log("Prefix: " + xmlMapping.prefix);
    }
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.XmlMappingUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.XmlMapping` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure an XML mapping on a content control by setting its XPath expression and prefix mappings to link the control to a custom XML part in the document.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    const xmlMapping = contentControl.xmlMapping;
    
    // Set multiple XML mapping properties at once
    xmlMapping.set({
        xpath: "/books/book[1]/title",
        prefixMappings: "xmlns:ns='http://example.com/books'"
    });
    
    await context.sync();
    
    console.log("XML mapping properties configured successfully");
});
```

---

### setMapping

**Kind:** `configure`

Allows creating or changing the XML mapping on the content control.

#### Signature

**Parameters:**
- `xPath`: `string` (required)
  The XPath expression to evaluate.
- `options`: `Word.XmlSetMappingOptions` (optional)
  Optional. The options available for setting the XML mapping.

**Returns:** `OfficeExtension.ClientResult<boolean>`

#### Examples

**Example**: Map a content control to a custom XML element using an XPath expression to bind the control's text to XML data

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Set the XML mapping using an XPath expression
    // This maps the content control to an XML element at the specified path
    contentControl.xmlMapping.setMapping(
        "/root/customer/name",
        { prefixMappings: "xmlns:ns='http://example.com/schema'" }
    );
    
    await context.sync();
    
    console.log("XML mapping successfully set for the content control");
});
```

---

### setMappingByNode

**Kind:** `configure`

Allows creating or changing the XML data mapping on the content control.

#### Signature

**Parameters:**
- `node`: `Word.CustomXmlNode` (required)
  The custom XML node to map.

**Returns:** `OfficeExtension.ClientResult<boolean>`

#### Examples

**Example**: Map a content control to a specific XML node from custom XML data in the document

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Load the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    
    await context.sync();
    
    // Get the first custom XML part (or create one if needed)
    const customXmlPart = customXmlParts.items[0];
    
    // Get the XML nodes from the custom XML part
    const xmlNodes = customXmlPart.getXml();
    
    await context.sync();
    
    // Parse the XML and get a specific node
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlNodes.value, "text/xml");
    const targetNode = xmlDoc.getElementsByTagName("employee")[0];
    
    // Set the mapping using the XML node
    const xmlMapping = contentControl.xmlMapping;
    xmlMapping.setMappingByNode(targetNode);
    
    await context.sync();
    
    console.log("Content control mapped to XML node successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.XmlMapping` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.XmlMappingData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.XmlMappingData`

#### Examples

**Example**: Serialize an XML mapping object to JSON format to inspect or log its properties, such as the XPath expression and custom XML part ID associated with a content control.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    const xmlMapping = contentControl.xmlMapping;
    
    // Load the XML mapping properties
    xmlMapping.load("customXmlPartId, xpath");
    
    await context.sync();
    
    // Convert the XML mapping to a plain JavaScript object
    const xmlMappingData = xmlMapping.toJSON();
    
    // Log or use the serialized data
    console.log("XML Mapping Data:", JSON.stringify(xmlMappingData, null, 2));
    console.log("XPath:", xmlMappingData.xpath);
    console.log("Custom XML Part ID:", xmlMappingData.customXmlPartId);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.XmlMapping`

#### Examples

**Example**: Track an XML mapping object across multiple sync calls to maintain its reference while checking and updating the mapping status of a content control.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    const xmlMapping = contentControl.xmlMapping;
    
    // Track the xmlMapping object to use it across multiple sync calls
    xmlMapping.track();
    
    // Load properties for the first sync
    xmlMapping.load("isMapped");
    await context.sync();
    
    // Check if mapped (using tracked object after sync)
    if (xmlMapping.isMapped) {
        console.log("Content control is mapped to XML");
        
        // Load additional properties in a second sync call
        xmlMapping.load("xpath");
        await context.sync();
        
        // Access the tracked object again after another sync
        console.log("XPath: " + xmlMapping.xpath);
    }
    
    // Untrack when done to free up memory
    xmlMapping.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.XmlMapping`

#### Examples

**Example**: Release memory for an XML mapping object after checking its properties to avoid memory leaks in a long-running add-in

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getByTag("myXmlControl").getFirst();
    const xmlMapping = contentControl.xmlMapping;
    xmlMapping.load("customXmlPart");
    
    await context.sync();
    
    // Use the XML mapping object
    console.log("XML mapping exists: " + (xmlMapping.customXmlPart !== null));
    
    // Untrack the object to release memory after use
    xmlMapping.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
