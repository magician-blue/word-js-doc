# Word.CustomProperty

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a custom property.

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

**Example**: Access the request context from a custom property to verify the connection between the add-in and Word, then use it to load and read the custom property's value.

```typescript
await Word.run(async (context) => {
    // Get a custom property from the document
    const customProperty = context.document.properties.customProperties.getItemOrNullObject("ProjectName");
    
    // Access the request context associated with the custom property
    const propertyContext = customProperty.context;
    
    // Use the context to load the property's value
    customProperty.load("key, value");
    await propertyContext.sync();
    
    // Check if the property exists and log its value
    if (!customProperty.isNullObject) {
        console.log(`Custom property '${customProperty.key}' has value: ${customProperty.value}`);
    } else {
        console.log("Custom property 'ProjectName' does not exist");
    }
});
```

---

### key

**Type:** `string`

**Since:** WordApi 1.3

Gets the key of the custom property.

#### Examples

**Example**: Retrieve and display the key of the first custom property in the document

```typescript
await Word.run(async (context) => {
    const customProperties = context.document.properties.customProperties;
    customProperties.load("items");
    
    await context.sync();
    
    if (customProperties.items.length > 0) {
        const firstProperty = customProperties.items[0];
        firstProperty.load("key");
        
        await context.sync();
        
        console.log("Custom property key: " + firstProperty.key);
    }
});
```

---

### type

**Type:** `Word.DocumentPropertyType | "String" | "Number" | "Date" | "Boolean"`

**Since:** WordApi 1.3

Gets the value type of the custom property. Possible values are: String, Number, Date, Boolean.

#### Examples

**Example**: Check the type of a custom property named "ProjectID" and display different messages based on whether it's a String, Number, Date, or Boolean type.

```typescript
await Word.run(async (context) => {
    const customProperty = context.document.properties.customProperties.getItemOrNullObject("ProjectID");
    customProperty.load("type");
    
    await context.sync();
    
    if (!customProperty.isNullObject) {
        const propertyType = customProperty.type;
        
        switch (propertyType) {
            case Word.DocumentPropertyType.string:
            case "String":
                console.log("ProjectID is a String type property");
                break;
            case Word.DocumentPropertyType.number:
            case "Number":
                console.log("ProjectID is a Number type property");
                break;
            case Word.DocumentPropertyType.date:
            case "Date":
                console.log("ProjectID is a Date type property");
                break;
            case Word.DocumentPropertyType.boolean:
            case "Boolean":
                console.log("ProjectID is a Boolean type property");
                break;
        }
    } else {
        console.log("ProjectID custom property does not exist");
    }
});
```

---

### value

**Type:** `any`

**Since:** WordApi 1.3

Specifies the value of the custom property. Note that even though Word on the web and the docx file format allow these properties to be arbitrarily long, the desktop version of Word will truncate string values to 255 16-bit chars (possibly creating invalid unicode by breaking up a surrogate pair).

#### Examples

**Example**: Set a custom document property named "ProjectCode" with the value "ALPHA-2024" to track project information in the Word document.

```typescript
await Word.run(async (context) => {
    // Get the custom properties collection
    const customProperties = context.document.properties.customProperties;
    
    // Add or update a custom property with a specific value
    const projectCodeProperty = customProperties.add("ProjectCode", "ALPHA-2024");
    
    // Load the value to verify
    projectCodeProperty.load("value");
    
    await context.sync();
    
    console.log(`Custom property value set to: ${projectCodeProperty.value}`);
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the custom property.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a custom property named "ProjectCode" from the document

```typescript
await Word.run(async (context) => {
    // Get the custom properties collection
    const properties = context.document.properties.customProperties;
    
    // Get the specific custom property by key
    const projectCodeProperty = properties.getItem("ProjectCode");
    
    // Delete the custom property
    projectCodeProperty.delete();
    
    await context.sync();
    
    console.log("Custom property 'ProjectCode' has been deleted.");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.CustomPropertyLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomProperty`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomProperty`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomProperty`

#### Examples

**Example**: Load and display the value and type of a custom document property named "ProjectCode"

```typescript
await Word.run(async (context) => {
    // Get the custom property by key
    const customProperty = context.document.properties.customProperties.getItem("ProjectCode");
    
    // Load the value and type properties
    customProperty.load("value, type");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the loaded properties
    console.log(`Property Value: ${customProperty.value}`);
    console.log(`Property Type: ${customProperty.type}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.CustomPropertyUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.CustomProperty` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a custom document property at once, setting both its value and type

```typescript
await Word.run(async (context) => {
    // Get the custom property named "ProjectStatus"
    const customProperty = context.document.properties.customProperties.getItemOrNullObject("ProjectStatus");
    
    // Set multiple properties at once
    customProperty.set({
        value: "In Progress",
        type: Word.DocumentPropertyType.string
    });
    
    await context.sync();
    console.log("Custom property updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomProperty` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomPropertyData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CustomPropertyData`

#### Examples

**Example**: Serialize a custom document property to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get a custom property from the document
    const customProperty = context.document.properties.customProperties.getItemOrNullObject("ProjectName");
    customProperty.load("key, type, value");
    
    await context.sync();
    
    if (!customProperty.isNullObject) {
        // Convert the custom property to a plain JavaScript object
        const propertyData = customProperty.toJSON();
        
        // Now you can use the plain object for logging, storage, or transmission
        console.log("Custom Property as JSON:", JSON.stringify(propertyData, null, 2));
        console.log("Property Key:", propertyData.key);
        console.log("Property Value:", propertyData.value);
        console.log("Property Type:", propertyData.type);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomProperty`

#### Examples

**Example**: Track a custom property object across multiple sync calls to prevent "InvalidObjectPath" errors when accessing it after document changes

```typescript
await Word.run(async (context) => {
    // Get a custom property
    const customProperty = context.document.properties.customProperties.getItemOrNullObject("ProjectName");
    customProperty.load("key,value");
    await context.sync();
    
    // Track the object to use it across multiple sync calls
    customProperty.track();
    
    // Make changes to the document
    context.document.body.insertParagraph("New content", Word.InsertLocation.end);
    await context.sync();
    
    // Access the tracked custom property again without errors
    console.log(`Custom property: ${customProperty.key} = ${customProperty.value}`);
    
    // Update the property value
    customProperty.value = "Updated Project";
    await context.sync();
    
    // Untrack when done
    customProperty.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomProperty`

#### Examples

**Example**: Create a custom property, use it to store document metadata, then untrack it to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    // Add a custom property to the document
    const customProperty = context.document.properties.customProperties.add("ProjectCode", "PRJ-2024-001");
    
    // Track the object to work with it
    customProperty.track();
    
    // Load and sync to ensure the property is created
    customProperty.load("key,value");
    await context.sync();
    
    console.log(`Custom property created: ${customProperty.key} = ${customProperty.value}`);
    
    // Untrack the object to release memory after we're done using it
    customProperty.untrack();
    
    // Sync to apply the memory release
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/read-write-custom-document-properties.yaml
