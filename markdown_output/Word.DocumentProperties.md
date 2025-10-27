# Word.DocumentProperties

**Package:** `word`

**API Set:** WordApi 1.3 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents document properties.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/get-built-in-properties.yaml

await Word.run(async (context) => {
    const builtInProperties: Word.DocumentProperties = context.document.properties;
    builtInProperties.load("*"); // Let's get all!

    await context.sync();
    console.log(JSON.stringify(builtInProperties, null, 4));
});
```

## Properties

### applicationName

**Type:** `string`

**Since:** WordApi 1.3

Gets the application name of the document.

#### Examples

**Example**: Display the name of the application that created the Word document in the console.

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.load("applicationName");
    
    await context.sync();
    
    console.log("Application name: " + properties.applicationName);
});
```

---

### author

**Type:** `string`

**Since:** WordApi 1.3

Specifies the author of the document.

#### Examples

**Example**: Set the document author to "Jane Smith"

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.author = "Jane Smith";
    await context.sync();
});
```

---

### category

**Type:** `string`

**Since:** WordApi 1.3

Specifies the category of the document.

#### Examples

**Example**: Set the document category to "Report" to organize and classify the document.

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.category = "Report";
    
    await context.sync();
});
```

---

### comments

**Type:** `string`

**Since:** WordApi 1.3

Specifies the Comments field in the metadata of the document. These have no connection to comments by users made in the document.

#### Examples

**Example**: Set the document's Comments metadata field to "This document requires legal review before publication"

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.comments = "This document requires legal review before publication";
    
    await context.sync();
});
```

---

### company

**Type:** `string`

**Since:** WordApi 1.3

Specifies the company of the document.

#### Examples

**Example**: Set the document's company property to "Contoso Ltd"

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.company = "Contoso Ltd";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the document properties context to verify the add-in is properly connected to the Word host application before performing operations.

```typescript
await Word.run(async (context) => {
    const docProperties = context.document.properties;
    
    // Access the request context to ensure connection to Word host
    const requestContext = docProperties.context;
    
    // Verify the context is valid by loading and syncing properties
    docProperties.load("title,author");
    await requestContext.sync();
    
    console.log("Connected to Word host application");
    console.log(`Document title: ${docProperties.title}`);
    console.log(`Document author: ${docProperties.author}`);
});
```

---

### creationDate

**Type:** `Date`

**Since:** WordApi 1.3

Gets the creation date of the document.

#### Examples

**Example**: Display the document creation date in the console to track when the document was originally created.

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.load("creationDate");
    
    await context.sync();
    
    console.log("Document created on: " + properties.creationDate.toLocaleDateString());
});
```

---

### customProperties

**Type:** `Word.CustomPropertyCollection`

**Since:** WordApi 1.3

Gets the collection of custom properties of the document.

#### Examples

**Example**: Add a new custom property named "ProjectCode" with value "WJS-2024" to the document and read it back to verify

```typescript
await Word.run(async (context) => {
    // Get the custom properties collection
    const customProperties = context.document.properties.customProperties;
    
    // Add a new custom property
    customProperties.add("ProjectCode", "WJS-2024");
    
    // Load the custom properties to verify
    customProperties.load("items");
    
    await context.sync();
    
    // Display the custom property value
    const projectCodeProperty = customProperties.items.find(prop => prop.key === "ProjectCode");
    if (projectCodeProperty) {
        console.log(`Custom property '${projectCodeProperty.key}' = '${projectCodeProperty.value}'`);
    }
});
```

---

### format

**Type:** `string`

**Since:** WordApi 1.3

Specifies the format of the document.

#### Examples

**Example**: Read and display the current document format type in the console

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.load("format");
    
    await context.sync();
    
    console.log("Document format: " + properties.format);
});
```

---

### keywords

**Type:** `string`

**Since:** WordApi 1.3

Specifies the keywords of the document.

#### Examples

**Example**: Set the document keywords to "annual report, financial, 2024" to improve document searchability and categorization.

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.keywords = "annual report, financial, 2024";
    
    await context.sync();
});
```

---

### lastAuthor

**Type:** `string`

**Since:** WordApi 1.3

Gets the last author of the document.

#### Examples

**Example**: Display the last author of the document in a content control

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.load("lastAuthor");
    
    await context.sync();
    
    const contentControl = context.document.body.insertContentControl();
    contentControl.insertText(`Last modified by: ${properties.lastAuthor}`, Word.InsertLocation.end);
    
    await context.sync();
});
```

---

### lastPrintDate

**Type:** `Date`

**Since:** WordApi 1.3

Gets the last print date of the document.

#### Examples

**Example**: Display the last print date of the document in a message box

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.load("lastPrintDate");
    
    await context.sync();
    
    console.log("Document last printed on: " + properties.lastPrintDate);
    // Or display in UI: alert("Last printed: " + properties.lastPrintDate);
});
```

---

### lastSaveTime

**Type:** `Date`

**Since:** WordApi 1.3

Gets the last save time of the document.

#### Examples

**Example**: Display the last save time of the document in a content control

```typescript
await Word.run(async (context) => {
    // Get the document properties
    const properties = context.document.properties;
    
    // Load the last save time
    properties.load("lastSaveTime");
    
    await context.sync();
    
    // Insert the last save time at the end of the document
    const lastSaveDate = properties.lastSaveTime;
    context.document.body.insertParagraph(
        `Document last saved: ${lastSaveDate.toLocaleString()}`,
        Word.InsertLocation.end
    );
    
    await context.sync();
});
```

---

### manager

**Type:** `string`

**Since:** WordApi 1.3

Specifies the manager of the document.

#### Examples

**Example**: Set the document's manager property to "Sarah Johnson"

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.manager = "Sarah Johnson";
    
    await context.sync();
});
```

---

### revisionNumber

**Type:** `string`

**Since:** WordApi 1.3

Gets the revision number of the document.

#### Examples

**Example**: Display the document's revision number in a content control so users can see how many times the document has been saved.

```typescript
await Word.run(async (context) => {
    // Get the document properties
    const properties = context.document.properties;
    
    // Load the revision number
    properties.load("revisionNumber");
    
    await context.sync();
    
    // Insert the revision number at the end of the document
    const body = context.document.body;
    body.insertParagraph(
        `Document Revision: ${properties.revisionNumber}`,
        Word.InsertLocation.end
    );
    
    await context.sync();
});
```

---

### security

**Type:** `number`

**Since:** WordApi 1.3

Gets security settings of the document. Some are access restrictions on the file on disk. Others are Document Protection settings. Some possible values are 0 = File on disk is read/write; 1 = Protect Document: File is encrypted and requires a password to open; 2 = Protect Document: Always Open as Read-Only; 3 = Protect Document: Both #1 and #2; 4 = File on disk is read-only; 5 = Both #1 and #4; 6 = Both #2 and #4; 7 = All of #1, #2, and #4; 8 = Protect Document: Restrict Edit to read-only; 9 = Both #1 and #8; 10 = Both #2 and #8; 11 = All of #1, #2, and #8; 12 = Both #4 and #8; 13 = All of #1, #4, and #8; 14 = All of #2, #4, and #8; 15 = All of #1, #2, #4, and #8.

#### Examples

**Example**: Check if the document has any security restrictions and display an appropriate message to the user

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.load("security");
    
    await context.sync();
    
    const securityLevel = properties.security;
    let message = "";
    
    if (securityLevel === 0) {
        message = "Document has no security restrictions";
    } else if (securityLevel === 1) {
        message = "Document is encrypted and requires a password";
    } else if (securityLevel === 2) {
        message = "Document is set to always open as read-only";
    } else if (securityLevel === 4) {
        message = "Document file on disk is read-only";
    } else if (securityLevel === 8) {
        message = "Document editing is restricted to read-only";
    } else {
        message = `Document has multiple security restrictions (level: ${securityLevel})`;
    }
    
    console.log(message);
});
```

---

### subject

**Type:** `string`

**Since:** WordApi 1.3

Specifies the subject of the document.

#### Examples

**Example**: Set the document subject to "Q4 Financial Report"

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.subject = "Q4 Financial Report";
    
    await context.sync();
});
```

---

### template

**Type:** `string`

**Since:** WordApi 1.3

Gets the template of the document.

#### Examples

**Example**: Get and display the template name of the current document in the console

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.load("template");
    
    await context.sync();
    
    console.log("Document template: " + properties.template);
});
```

---

### title

**Type:** `string`

**Since:** WordApi 1.3

Specifies the title of the document.

#### Examples

**Example**: Set the document title to "Q4 Sales Report"

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    properties.title = "Q4 Sales Report";
    await context.sync();
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.DocumentPropertiesLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.DocumentProperties`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.DocumentProperties`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.DocumentProperties`

#### Examples

**Example**: Load and display the document's title and author properties in the console

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    
    // Load specific properties
    properties.load("title, author");
    
    await context.sync();
    
    console.log("Document Title: " + properties.title);
    console.log("Document Author: " + properties.author);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.DocumentPropertiesUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.DocumentProperties` (required)

  **Returns:** `void`

#### Examples

**Example**: Set multiple document properties including title, author, and subject at once

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    
    properties.set({
        title: "Annual Sales Report",
        author: "John Smith",
        subject: "Q4 2023 Sales Analysis"
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.DocumentProperties` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentPropertiesData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.DocumentPropertiesData`

#### Examples

**Example**: Serialize document properties to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the document properties
    const properties = context.document.properties;
    
    // Load the properties you want to serialize
    properties.load("title,author,subject,keywords,comments,creationDate,lastAuthor");
    
    await context.sync();
    
    // Convert to plain JavaScript object using toJSON()
    const propertiesData = properties.toJSON();
    
    // Now you can use the plain object for logging, storage, etc.
    console.log("Document Properties:", JSON.stringify(propertiesData, null, 2));
    
    // Example: Send to external API or save to local storage
    // await fetch('/api/save-metadata', { 
    //     method: 'POST', 
    //     body: JSON.stringify(propertiesData) 
    // });
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.DocumentProperties`

#### Examples

**Example**: Track document properties object to access its values across multiple sync calls without getting an InvalidObjectPath error

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties;
    
    // Track the object to use it across multiple sync calls
    properties.track();
    
    // Load properties in first sync
    properties.load("title,author");
    await context.sync();
    
    console.log("Title: " + properties.title);
    console.log("Author: " + properties.author);
    
    // Modify properties after sync - tracking prevents InvalidObjectPath error
    properties.title = "Updated Document Title";
    properties.author = "New Author";
    await context.sync();
    
    // Access the tracked object again after another sync
    console.log("New Title: " + properties.title);
    
    // Untrack when done to release memory
    properties.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.DocumentProperties`

#### Examples

**Example**: Load document properties, read the title, then untrack the object to free memory after use

```typescript
await Word.run(async (context) => {
    // Load the document properties
    const properties = context.document.properties;
    properties.load("title");
    
    await context.sync();
    
    // Use the properties
    console.log("Document title: " + properties.title);
    
    // Untrack the object to release memory
    properties.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.documentproperties
