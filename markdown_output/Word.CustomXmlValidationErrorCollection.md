# Word.CustomXmlValidationErrorCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror) objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a CustomXmlValidationErrorCollection to verify the connection between the add-in and Word before processing validation errors.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const validationErrors = customXmlParts.items[0].validationErrors;
        validationErrors.load("items");
        await context.sync();

        // Access the request context from the collection
        const requestContext = validationErrors.context;
        
        // Verify the context is valid and connected
        console.log("Context is connected:", requestContext !== null);
        console.log("Number of validation errors:", validationErrors.items.length);
    }
});
```

---

### items

**Type:** `Word.CustomXmlValidationError[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Retrieve and log all validation errors from a custom XML part to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the validation errors collection
        const validationErrors = customXmlPart.validationErrors;
        validationErrors.load("items");
        await context.sync();
        
        // Access the items property to get all validation errors
        const errorItems = validationErrors.items;
        
        console.log(`Found ${errorItems.length} validation error(s)`);
        
        // Loop through each error item
        errorItems.forEach((error, index) => {
            error.load("text, type");
        });
        
        await context.sync();
        
        errorItems.forEach((error, index) => {
            console.log(`Error ${index + 1}: ${error.text} (Type: ${error.type})`);
        });
    }
});
```

---

## Methods

### add

**Kind:** `create`

Adds a CustomXmlValidationError object containing an XML validation error to the CustomXmlValidationErrorCollection object.

#### Signature

**Parameters:**
- `node`: `Word.CustomXmlNode` (required)
  The XML node where the error occurred.
- `errorName`: `string` (required)
  The name of the error.
- `options`: `Word.CustomXmlAddValidationErrorOptions` (optional)
  Optional. The options that define the error message and other settings.

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Add a validation error to a custom XML part's error collection when detecting an invalid XML node structure

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlPart = context.document.customXmlParts.getByNamespace("http://example.com/schema").getOnlyItem();
    
    // Get the validation error collection
    const errorCollection = customXmlPart.validationErrors;
    
    // Get a specific XML node (assuming we have a reference to it)
    const xmlNodes = customXmlPart.query("//invalidNode");
    const node = xmlNodes.getItem(0);
    
    // Add a validation error for this node
    errorCollection.add(node, "InvalidNodeStructure", { errorText: "The node structure does not match the schema requirements" });
    
    await context.sync();
    
    console.log("Validation error added to the custom XML part");
});
```

---

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Check if a custom XML part has any validation errors and display the count

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get validation errors for this custom XML part
        const validationErrors = customXmlPart.validationErrors;
        
        // Get the count of validation errors
        const errorCount = validationErrors.getCount();
        await context.sync();
        
        console.log(`Number of validation errors: ${errorCount.value}`);
    }
});
```

---

### getItem

**Kind:** `read`

Returns a CustomXmlValidationError object that represents the specified item in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  A number that identifies the index location of a paragraph object.

**Returns:** `Word.CustomXmlValidationError`

#### Examples

**Example**: Get and display the error message from the first custom XML validation error in the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML validation errors collection
    const validationErrors = context.document.customXmlParts.getByNamespace("http://example.com/schema").getOnlyItem().validationErrors;
    
    // Get the first validation error from the collection
    const firstError = validationErrors.getItem(0);
    
    // Load the error's properties
    firstError.load("errorCode, errorMessage");
    
    await context.sync();
    
    // Display the error information
    console.log(`Error Code: ${firstError.errorCode}`);
    console.log(`Error Message: ${firstError.errorMessage}`);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.CustomXmlValidationErrorCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlValidationErrorCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlValidationErrorCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlValidationErrorCollection`

#### Examples

**Example**: Load and display validation error details from a custom XML part's validation errors collection

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the validation errors collection
        const validationErrors = customXmlPart.validationErrors;
        
        // Load properties of the validation errors
        validationErrors.load("items");
        await context.sync();
        
        // Display validation error information
        console.log(`Total validation errors: ${validationErrors.items.length}`);
        
        validationErrors.items.forEach((error, index) => {
            error.load("errorCode, errorType");
        });
        await context.sync();
        
        validationErrors.items.forEach((error, index) => {
            console.log(`Error ${index + 1}: Code=${error.errorCode}, Type=${error.errorType}`);
        });
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlValidationErrorCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlValidationErrorCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlValidationErrorCollectionData`

#### Examples

**Example**: Serialize custom XML validation errors to JSON format for logging or debugging purposes

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts, "items");
    await context.sync();

    // Get validation errors from the first custom XML part (if it exists)
    if (customXmlParts.items.length > 0) {
        const validationErrors = customXmlParts.items[0].validationErrors;
        context.load(validationErrors, "items");
        await context.sync();

        // Convert the validation errors collection to a plain JSON object
        const errorsJson = validationErrors.toJSON();
        
        // Log or process the JSON representation
        console.log("Validation Errors:", JSON.stringify(errorsJson, null, 2));
        
        // Access the items array from the JSON object
        if (errorsJson.items && errorsJson.items.length > 0) {
            console.log(`Found ${errorsJson.items.length} validation error(s)`);
        }
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlValidationErrorCollection`

#### Examples

**Example**: Track a custom XML validation error collection across multiple sync calls to monitor validation errors without encountering InvalidObjectPath errors

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get the validation errors collection
        const validationErrors = customXmlPart.validationErrors;
        validationErrors.load("items");
        
        // Track the collection for use across sync calls
        validationErrors.track();
        
        await context.sync();
        
        // Now we can safely use the collection across multiple syncs
        console.log(`Found ${validationErrors.items.length} validation errors`);
        
        await context.sync();
        
        // The tracked object remains valid even after additional sync calls
        for (const error of validationErrors.items) {
            error.load("errorCode, errorText");
        }
        
        await context.sync();
        
        // Clean up tracking when done
        validationErrors.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlValidationErrorCollection`

#### Examples

**Example**: Get custom XML validation errors, process them, then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts, "items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const validationErrors = customXmlParts.items[0].validationErrors;
        context.load(validationErrors, "items");
        await context.sync();
        
        // Process the validation errors
        console.log(`Found ${validationErrors.items.length} validation errors`);
        
        // Untrack the collection to release memory
        validationErrors.untrack();
        await context.sync();
    }
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.customxmlvalidationerror
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.customxmlvalidationerror
- /en-us/javascript/api/word/word.customxmlnode
- /en-us/javascript/api/word/word.customxmladdvalidationerroroptions
- /en-us/javascript/api/office/officeextension.clientresult
- /en-us/javascript/api/word/word.interfaces.customxmlvalidationerrorcollectionloadoptions
- /en-us/javascript/api/word/word.interfaces.collectionloadoptions
- /en-us/javascript/api/word/word.customxmlvalidationerrorcollection
- /en-us/javascript/api/office/officeextension.loadoption
- /en-us/javascript/api/word/word.interfaces.customxmlvalidationerrorcollectiondata
- /en-us/javascript/api/office/officeextension.clientrequestcontext
