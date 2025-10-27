# Word.CustomPropertyCollection

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains the collection of [Word.CustomProperty](/en-us/javascript/api/word/word.customproperty) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml

await Word.run(async (context) => {
    const properties: Word.CustomPropertyCollection = context.document.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    for (let i = 0; i < properties.items.length; i++)
        console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomPropertyCollection to verify the connection between the add-in and Word before performing operations on custom document properties.

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties.customProperties;
    properties.load("items");
    await context.sync();
    
    // Access the request context associated with the collection
    const requestContext = properties.context;
    
    // Verify the context is valid by checking if it matches the current context
    if (requestContext === context) {
        console.log("CustomPropertyCollection is connected to the current Word context");
        console.log(`Found ${properties.items.length} custom properties`);
    }
});
```

---

### items

**Type:** `Word.CustomProperty[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve and display all custom document properties including their names, types, and values from a Word document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml

await Word.run(async (context) => {
    const properties: Word.CustomPropertyCollection = context.document.properties.customProperties;
    properties.load("key,type,value");

    await context.sync();
    for (let i = 0; i < properties.items.length; i++)
        console.log("Property Name:" + properties.items[i].key + "; Type=" + properties.items[i].type + "; Property Value=" + properties.items[i].value);
});
```

---

## Methods

### add

**Kind:** `create`

Creates a new or sets an existing custom property.

#### Signature

**Parameters:**
- `key`: `string` (required)
  The custom property's key, which is case-insensitive.
- `value`: `any` (required)
  The custom property's value.

**Returns:** `Word.CustomProperty`

#### Examples

**Example**: Add custom properties to a Word document with specified names and values of different types (numeric and string).

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml

await Word.run(async (context) => {
    context.document.properties.customProperties.add("Numeric Property", 1234);

    await context.sync();
    console.log("Property added");
});

...

await Word.run(async (context) => {
    context.document.properties.customProperties.add("String Property", "Hello World!");

    await context.sync();
    console.log("Property added");
});
```

---

### deleteAll

**Kind:** `delete`

Deletes all custom properties in this collection.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete all custom properties from the current document

```typescript
await Word.run(async (context) => {
    // Get the custom properties collection
    const customProperties = context.document.properties.customProperties;
    
    // Delete all custom properties
    customProperties.deleteAll();
    
    await context.sync();
    console.log("All custom properties have been deleted.");
});
```

---

### getCount

**Kind:** `read`

Gets the count of custom properties.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Check how many custom properties exist in the document and display the count in the console.

```typescript
await Word.run(async (context) => {
    const customProperties = context.document.properties.customProperties;
    const count = customProperties.getCount();
    
    await context.sync();
    
    console.log(`The document has ${count.value} custom properties.`);
});
```

---

### getItem

**Kind:** `read`

Gets a custom property object by its key, which is case-insensitive. Throws an ItemNotFound error if the custom property doesn't exist.

#### Signature

**Parameters:**
- `key`: `string` (required)
  The key that identifies the custom property object.

**Returns:** `Word.CustomProperty`

#### Examples

**Example**: Retrieve and display the value of an existing custom property named "ProjectCode" from the document

```typescript
await Word.run(async (context) => {
    const customProperties = context.document.properties.customProperties;
    const projectCodeProperty = customProperties.getItem("ProjectCode");
    
    projectCodeProperty.load("value");
    await context.sync();
    
    console.log("Project Code: " + projectCodeProperty.value);
});
```

---

### getItemOrNullObject

**Kind:** `read`

Gets a custom property object by its key, which is case-insensitive. If the custom property doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Parameters:**
- `key`: `string` (required)
  The key that identifies the custom property object.

**Returns:** `Word.CustomProperty`

#### Examples

**Example**: Check if a custom property named "ProjectCode" exists in the document and display its value, or show a message if it doesn't exist.

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties.customProperties;
    const projectCodeProperty = properties.getItemOrNullObject("ProjectCode");
    
    projectCodeProperty.load("key, value, isNullObject");
    await context.sync();
    
    if (projectCodeProperty.isNullObject) {
        console.log("ProjectCode property does not exist");
    } else {
        console.log(`ProjectCode: ${projectCodeProperty.value}`);
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
  - `options`: `Word.Interfaces.CustomPropertyCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomPropertyCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomPropertyCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomPropertyCollection`

#### Examples

**Example**: Load and display all custom document properties including their keys and values from the current Word document.

```typescript
await Word.run(async (context) => {
    // Get the custom properties collection
    const customProperties = context.document.properties.customProperties;
    
    // Load the collection with specific properties
    customProperties.load("items");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the custom properties
    console.log(`Found ${customProperties.items.length} custom properties:`);
    customProperties.items.forEach(property => {
        property.load("key, value");
    });
    
    await context.sync();
    
    customProperties.items.forEach(property => {
        console.log(`${property.key}: ${property.value}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomPropertyCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomPropertyCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CustomPropertyCollectionData`

#### Examples

**Example**: Serialize custom document properties to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the custom properties collection
    const customProperties = context.document.properties.customProperties;
    customProperties.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const jsonData = customProperties.toJSON();
    
    // Log the serialized data (can be stored or transmitted)
    console.log("Custom Properties as JSON:", JSON.stringify(jsonData, null, 2));
    
    // Access the items array from the serialized data
    jsonData.items.forEach(prop => {
        console.log(`Property: ${prop.key}, Value: ${prop.value}, Type: ${prop.type}`);
    });
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomPropertyCollection`

#### Examples

**Example**: Track custom properties collection to safely access and modify custom properties across multiple sync calls without getting InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    const properties = context.document.properties.customProperties;
    properties.load("items");
    await context.sync();
    
    // Track the collection to use it across multiple sync calls
    properties.track();
    
    // Add a new custom property
    properties.add("ProjectStatus", "Active");
    await context.sync();
    
    // Safe to access the tracked object after sync
    console.log(`Total custom properties: ${properties.items.length}`);
    
    // Modify another property
    properties.add("LastModified", new Date().toISOString());
    await context.sync();
    
    // Still safe to access because it's tracked
    console.log(`Updated count: ${properties.items.length}`);
    
    // Clean up tracking when done
    properties.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomPropertyCollection`

#### Examples

**Example**: Load custom properties, read their values, then untrack the collection to free memory after processing is complete.

```typescript
await Word.run(async (context) => {
    // Get the custom properties collection
    const customProperties = context.document.properties.customProperties;
    customProperties.load("key, value");
    
    await context.sync();
    
    // Process the custom properties
    customProperties.items.forEach(property => {
        console.log(`${property.key}: ${property.value}`);
    });
    
    // Untrack the collection to release memory
    customProperties.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.custompropertycollection
