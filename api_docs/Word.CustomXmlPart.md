# Word.CustomXmlPart class

Package: [word](/en-us/javascript/api/word)

Represents a custom XML part.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.4]

### Examples
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
- builtIn  
  Gets a value that indicates whether the CustomXmlPart is built-in.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- documentElement  
  Gets the root element of a bound region of data in the document. If the region is empty, the property returns Nothing.
- errors  
  Gets a CustomXmlValidationErrorCollection object that provides access to any XML validation errors.
- id  
  Gets the ID of the custom XML part.
- namespaceManager  
  Gets the set of namespace prefix mappings used against the current CustomXmlPart object.
- namespaceUri  
  Gets the namespace URI of the custom XML part.
- schemaCollection  
  Specifies a CustomXmlSchemaCollection object representing the set of schemas attached to a bound region of data in the document.
- xml  
  Gets the XML representation of the current CustomXmlPart object.

## Methods
- addNode(parent, options)  
  Adds a node to the XML tree.
- delete()  
  Deletes the custom XML part.
- deleteAttribute(xpath, namespaceMappings, name)  
  Deletes an attribute with the given name from the element identified by xpath.
- deleteElement(xpath, namespaceMappings)  
  Deletes the element identified by xpath.
- getXml()  
  Gets the full XML content of the custom XML part.
- insertAttribute(xpath, namespaceMappings, name, value)  
  Inserts an attribute with the given name and value to the element identified by xpath.
- insertElement(xpath, xml, namespaceMappings, index)  
  Inserts the given XML under the parent element identified by xpath at child position index.
- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- loadXml(xml)  
  Populates the CustomXmlPart object from an XML string.
- query(xpath, namespaceMappings)  
  Queries the XML content of the custom XML part.
- selectNodes(xPath)  
  Selects a collection of nodes from a custom XML part.
- selectSingleNode(xPath)  
  Selects a single node within a custom XML part matching an XPath expression.
- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- setXml(xml)  
  Sets the full XML content of the custom XML part.
- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlPart object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlPartData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.
- updateAttribute(xpath, namespaceMappings, name, value)  
  Updates the value of an attribute with the given name of the element identified by xpath.
- updateElement(xpath, xml, namespaceMappings)  
  Updates the XML of the element identified by xpath.

## Property Details

### builtIn
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a value that indicates whether the CustomXmlPart is built-in.

```typescript
readonly builtIn: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### documentElement
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the root element of a bound region of data in the document. If the region is empty, the property returns Nothing.

```typescript
readonly documentElement: Word.CustomXmlNode;
```

Property Value
- [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### errors
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a CustomXmlValidationErrorCollection object that provides access to any XML validation errors.

```typescript
readonly errors: Word.CustomXmlValidationErrorCollection;
```

Property Value
- [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### id
Gets the ID of the custom XML part.

```typescript
readonly id: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.4]

#### Examples
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
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the set of namespace prefix mappings used against the current CustomXmlPart object.

```typescript
readonly namespaceManager: Word.CustomXmlPrefixMappingCollection;
```

Property Value
- [Word.CustomXmlPrefixMappingCollection](/en-us/javascript/api/word/word.customxmlprefixmappingcollection)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### namespaceUri
Gets the namespace URI of the custom XML part.

```typescript
readonly namespaceUri: string;
```

Property Value
- string

Remarks  
[API set: WordApi 1.4]

#### Examples
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
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a CustomXmlSchemaCollection object representing the set of schemas attached to a bound region of data in the document.

```typescript
schemaCollection: Word.CustomXmlSchemaCollection;
```

Property Value
- [Word.CustomXmlSchemaCollection](/en-us/javascript/api/word/word.customxmlschemacollection)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### xml
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the XML representation of the current CustomXmlPart object.

```typescript
readonly xml: string;
```

Property Value
- string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

## Method Details

### addNode(parent, options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds a node to the XML tree.

```typescript
addNode(parent: Word.CustomXmlNode, options?: Word.CustomXmlAddNodeOptions): OfficeExtension.ClientResult<number>;
```

Parameters
- parent: [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)  
  The parent node to which the new node will be added.
- options: [Word.CustomXmlAddNodeOptions](/en-us/javascript/api/word/word.customxmladdnodeoptions)  
  Optional. The options that define the node to be added.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### delete()
Deletes the custom XML part.

```typescript
delete(): void;
```

Returns
- void

Remarks  
[API set: WordApi 1.4]

#### Examples
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

### deleteAttribute(xpath, namespaceMappings, name)
Deletes an attribute with the given name from the element identified by xpath.

```typescript
deleteAttribute(xpath: string, namespaceMappings: { [key: string]: string; }, name: string): void;
```

Parameters
- xpath: string  
  Required. Absolute path to the single element in XPath notation.
- namespaceMappings: { [key: string]: string; }  
  Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".
- name: string  
  Required. Name of the attribute.

Returns
- void

Remarks  
[API set: WordApi 1.4]

If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI), an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose the tree is the following:
```xml
<Day>
  <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
    <Week>something</Week>
  </Month>
</Day>
```
The xpath to `<Week>` must be /Day/greg:Month/Week, where greg is an alias that is mapped to "http://calendartypes.org/xsds/GregorianCalendar" in the namespaceMappings parameter.

---

### deleteElement(xpath, namespaceMappings)
Deletes the element identified by xpath.

```typescript
deleteElement(xpath: string, namespaceMappings: { [key: string]: string; }): void;
```

Parameters
- xpath: string  
  Required. Absolute path to the single element in XPath notation.
- namespaceMappings: { [key: string]: string; }  
  Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".

Returns
- void

Remarks  
[API set: WordApi 1.4]

If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI), an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose the tree is the following:
```xml
<Day>
  <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
    <Week>something</Week>
  </Month>
</Day>
```
The xpath to `<Week>` must be /Day/greg:Month/Week, where greg is an alias that is mapped to "http://calendartypes.org/xsds/GregorianCalendar" in the namespaceMappings parameter.

---

### getXml()
Gets the full XML content of the custom XML part.

```typescript
getXml(): OfficeExtension.ClientResult<string>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks  
[API set: WordApi 1.4]

#### Examples
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

### insertAttribute(xpath, namespaceMappings, name, value)
Inserts an attribute with the given name and value to the element identified by xpath.

```typescript
insertAttribute(xpath: string, namespaceMappings: { [key: string]: string; }, name: string, value: string): void;
```

Parameters
- xpath: string  
  Required. Absolute path to the single element in XPath notation.
- namespaceMappings: { [key: string]: string; }  
  Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".
- name: string  
  Required. Name of the attribute.
- value: string  
  Required. Value of the attribute.

Returns
- void

Remarks  
[API set: WordApi 1.4]

If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI), an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose the tree is the following:
```xml
<Day>
  <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
    <Week>something</Week>
  </Month>
</Day>
```
The xpath to `<Week>` must be /Day/greg:Month/Week, where greg is an alias that is mapped to "http://calendartypes.org/xsds/GregorianCalendar" in the namespaceMappings parameter.

#### Examples
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

### insertElement(xpath, xml, namespaceMappings, index)
Inserts the given XML under the parent element identified by xpath at child position index.

```typescript
insertElement(xpath: string, xml: string, namespaceMappings: { [key: string]: string; }, index?: number): void;
```

Parameters
- xpath: string  
  Required. Absolute path to the single parent element in XPath notation.
- xml: string  
  Required. XML content to be inserted.
- namespaceMappings: { [key: string]: string; }  
  Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".
- index: number  
  Optional. Zero-based position at which the new XML to be inserted. If omitted, the XML will be appended as the last child of this parent.

Returns
- void

Remarks  
[API set: WordApi 1.4]

If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI), an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose the tree is the following:
```xml
<Day>
  <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
    <Week>something</Week>
  </Month>
</Day>
```
The xpath to `<Week>` must be /Day/greg:Month/Week, where greg is an alias that is mapped to "http://calendartypes.org/xsds/GregorianCalendar" in the namespaceMappings parameter.

#### Examples
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

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.CustomXmlPartLoadOptions): Word.CustomXmlPart;
```

Parameters
- options: [Word.Interfaces.CustomXmlPartLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlpartloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

---

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomXmlPart;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

---

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: { select?: string; expand?: string; }): Word.CustomXmlPart;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

---

### loadXml(xml)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Populates the CustomXmlPart object from an XML string.

```typescript
loadXml(xml: string): OfficeExtension.ClientResult<boolean>;
```

Parameters
- xml: string  
  The XML string to load.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<boolean>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### query(xpath, namespaceMappings)
Queries the XML content of the custom XML part.

```typescript
query(xpath: string, namespaceMappings: { [key: string]: string; }): OfficeExtension.ClientResult<string[]>;
```

Parameters
- xpath: string  
  Required. An XPath query.
- namespaceMappings: { [key: string]: string; }  
  Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string[]>  
  An array where each item represents an entry matched by the XPath query.

Remarks  
[API set: WordApi 1.4]

If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI), an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose the tree is the following:
```xml
<Day>
  <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
    <Week>something</Week>
  </Month>
</Day>
```
The xpath to `<Week>` must be /Day/greg:Month/Week, where greg is an alias that is mapped to "http://calendartypes.org/xsds/GregorianCalendar" in the namespaceMappings parameter.

#### Examples
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

### selectNodes(xPath)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Selects a collection of nodes from a custom XML part.

```typescript
selectNodes(xPath: string): Word.CustomXmlNodeCollection;
```

Parameters
- xPath: string  
  The XPath expression to evaluate.

Returns
- [Word.CustomXmlNodeCollection](/en-us/javascript/api/word/word.customxmlnodecollection)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### selectSingleNode(xPath)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Selects a single node within a custom XML part matching an XPath expression.

```typescript
selectSingleNode(xPath: string): Word.CustomXmlNode;
```

Parameters
- xPath: string  
  The XPath expression to evaluate.

Returns
- [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

---

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.CustomXmlPartUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.CustomXmlPartUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlpartupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

---

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.CustomXmlPart): void;
```

Parameters
- properties: [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

Returns
- void

---

### setXml(xml)
Sets the full XML content of the custom XML part.

```typescript
setXml(xml: string): void;
```

Parameters
- xml: string  
  Required. XML content to be set.

Returns
- void

Remarks  
[API set: WordApi 1.4]

#### Examples
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

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlPart object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlPartData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CustomXmlPartData;
```

Returns
- [Word.Interfaces.CustomXmlPartData](/en-us/javascript/api/word/word.interfaces.customxmlpartdata)

---

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomXmlPart;
```

Returns
- [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

---

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.CustomXmlPart;
```

Returns
- [Word.CustomXmlPart](/en-us/javascript/api/word/word.customxmlpart)

---

### updateAttribute(xpath, namespaceMappings, name, value)
Updates the value of an attribute with the given name of the element identified by xpath.

```typescript
updateAttribute(xpath: string, namespaceMappings: { [key: string]: string; }, name: string, value: string): void;
```

Parameters
- xpath: string  
  Required. Absolute path to the single element in XPath notation.
- namespaceMappings: { [key: string]: string; }  
  Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".
- name: string  
  Required. Name of the attribute.
- value: string  
  Required. New value of the attribute.

Returns
- void

Remarks  
[API set: WordApi 1.4]

If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI), an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose the tree is the following:
```xml
<Day>
  <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
    <Week>something</Week>
  </Month>
</Day>
```
The xpath to `<Week>` must be /Day/greg:Month/Week, where greg is an alias that is mapped to "http://calendartypes.org/xsds/GregorianCalendar" in the namespaceMappings parameter.

---

### updateElement(xpath, xml, namespaceMappings)
Updates the XML of the element identified by xpath.

```typescript
updateElement(xpath: string, xml: string, namespaceMappings: { [key: string]: string; }): void;
```

Parameters
- xpath: string  
  Required. Absolute path to the single element in XPath notation.
- xml: string  
  Required. New XML content to be stored.
- namespaceMappings: { [key: string]: string; }  
  Required. An object whose property values are namespace names and whose property names are aliases for the corresponding namespaces. For example, `{greg: "http://calendartypes.org/xsds/GregorianCalendar"}`. The property names (such as "greg") can be any string that doesn't used reserved XPath characters, such as the forward slash "/".

Returns
- void

Remarks  
[API set: WordApi 1.4]

If any element in the tree has an xmlns attribute (whose value is typically, but not always, a URI), an alias for that attribute value must prefix the element name in the xpath parameter. For example, suppose the tree is the following:
```xml
<Day>
  <Month xmlns="http://calendartypes.org/xsds/GregorianCalendar">
    <Week>something</Week>
  </Month>
</Day>
```
The xpath to `<Week>` must be /Day/greg:Month/Week, where greg is an alias that is mapped to "http://calendartypes.org/xsds/GregorianCalendar" in the namespaceMappings parameter.