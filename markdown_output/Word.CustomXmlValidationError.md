# Word.CustomXmlValidationError

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a single validation error in a Word.CustomXmlValidationErrorCollection object.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a custom XML validation error to load and read its properties

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        
        // Get validation errors for the XML part
        const validationErrors = xmlPart.getValidationErrors();
        validationErrors.load("items");
        await context.sync();

        if (validationErrors.items.length > 0) {
            const error = validationErrors.items[0];
            
            // Access the request context from the validation error
            const errorContext = error.context;
            
            // Use the context to load error properties
            error.load("errorCode, reason");
            await errorContext.sync();
            
            console.log(`Error code: ${error.errorCode}`);
            console.log(`Reason: ${error.reason}`);
        }
    }
});
```

---

### errorCode

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets an integer representing the validation error in the CustomXmlValidationError object.

#### Examples

**Example**: Check if a custom XML part has validation errors and log the error code of the first validation error found.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        const validationErrors = customXmlPart.getValidationErrors();
        validationErrors.load("items");
        await context.sync();

        if (validationErrors.items.length > 0) {
            const firstError = validationErrors.items[0];
            firstError.load("errorCode");
            await context.sync();

            console.log(`Validation error code: ${firstError.errorCode}`);
        } else {
            console.log("No validation errors found.");
        }
    }
});
```

---

### name

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the name of the error in the CustomXmlValidationError object.If no errors exist, the property returns Nothing

#### Examples

**Example**: Check if a custom XML part has validation errors and display the name of the first error found, or show a message if no errors exist.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        const validationErrors = customXmlPart.validationErrors;
        validationErrors.load("items");
        await context.sync();

        if (validationErrors.items.length > 0) {
            const firstError = validationErrors.items[0];
            firstError.load("name");
            await context.sync();

            console.log(`Validation error name: ${firstError.name}`);
        } else {
            console.log("No validation errors found.");
        }
    }
});
```

---

### node

**Type:** `Word.CustomXmlNode`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the node associated with this CustomXmlValidationError object, if any exist.If no nodes exist, the property returns Nothing.

#### Examples

**Example**: Get the XML node associated with a validation error and display its base name in the console.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        const validationErrors = customXmlPart.validationErrors;
        validationErrors.load("items");
        await context.sync();

        if (validationErrors.items.length > 0) {
            const error = validationErrors.items[0];
            const errorNode = error.node;
            
            if (errorNode) {
                errorNode.load("baseName");
                await context.sync();
                console.log("Error node base name: " + errorNode.baseName);
            } else {
                console.log("No node associated with this validation error.");
            }
        }
    }
});
```

---

### text

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the text in the CustomXmlValidationError object.

#### Examples

**Example**: Display all validation error messages from custom XML parts in the document to the console.

```typescript
await Word.run(async (context) => {
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();

    for (let i = 0; i < customXmlParts.items.length; i++) {
        const xmlPart = customXmlParts.items[i];
        const errors = xmlPart.validationErrors;
        errors.load("items");
        await context.sync();

        for (let j = 0; j < errors.items.length; j++) {
            const error = errors.items[j];
            error.load("text");
            await context.sync();
            
            console.log(`Validation error: ${error.text}`);
        }
    }
});
```

---

### type

**Type:** `Word.CustomXmlValidationErrorType | "schemaGenerated" | "automaticallyCleared" | "manual"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the type of error generated from the CustomXmlValidationError object.

#### Examples

**Example**: Check if a custom XML validation error was manually created or automatically generated by the system

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts, "items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get validation errors for the first custom XML part
        const validationErrors = customXmlParts.items[0].validationErrors;
        context.load(validationErrors, "items");
        await context.sync();

        // Check the type of each validation error
        for (const error of validationErrors.items) {
            error.load("type");
        }
        await context.sync();

        // Display error types
        validationErrors.items.forEach((error, index) => {
            console.log(`Error ${index + 1} type: ${error.type}`);
            
            if (error.type === "manual") {
                console.log("This error was manually created");
            } else if (error.type === "schemaGenerated") {
                console.log("This error was generated by schema validation");
            } else if (error.type === "automaticallyCleared") {
                console.log("This error was automatically cleared");
            }
        });
    }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes this CustomXmlValidationError object.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete the first validation error from a custom XML part to clear it from the error collection

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlPart = context.document.customXmlParts.getByNamespace("http://example.com/schema").getOnlyItem();
    
    // Get the validation errors collection
    const validationErrors = customXmlPart.validationErrors;
    validationErrors.load("items");
    
    await context.sync();
    
    // Delete the first validation error if it exists
    if (validationErrors.items.length > 0) {
        const firstError = validationErrors.items[0];
        firstError.delete();
        
        await context.sync();
        console.log("First validation error deleted successfully");
    } else {
        console.log("No validation errors to delete");
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
  - `options`: `Word.Interfaces.CustomXmlValidationErrorLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.CustomXmlValidationError`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.CustomXmlValidationError`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.CustomXmlValidationError`

#### Examples

**Example**: Load and display the error type and reason for the first custom XML validation error in the document

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts, "items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get validation errors for the first custom XML part
        const validationErrors = customXmlParts.items[0].validationErrors;
        context.load(validationErrors, "items");
        await context.sync();

        if (validationErrors.items.length > 0) {
            // Get the first validation error
            const error = validationErrors.items[0];
            
            // Load properties of the validation error
            error.load("errorType, reason");
            await context.sync();

            // Display the error details
            console.log(`Error Type: ${error.errorType}`);
            console.log(`Reason: ${error.reason}`);
        }
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
  - `properties`: `Interfaces.CustomXmlValidationErrorUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.CustomXmlValidationError` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a custom XML validation error object at once to configure its error details

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    const customXmlPart = customXmlParts.items[0];
    
    // Get validation errors collection
    const validationErrors = customXmlPart.validationErrors;
    validationErrors.load("items");
    await context.sync();
    
    // Get the first validation error
    const validationError = validationErrors.items[0];
    
    // Set multiple properties at once using the set() method
    validationError.set({
        errorCode: 1001,
        errorType: Word.CustomXmlValidationErrorType.schemaValidation
    });
    
    await context.sync();
    
    console.log("Validation error properties updated");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlValidationError object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlValidationErrorData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.CustomXmlValidationErrorData`

#### Examples

**Example**: Serialize a custom XML validation error to a plain JavaScript object for logging or debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts, "items");
    await context.sync();

    if (customXmlParts.items.length > 0) {
        // Get validation errors for the first custom XML part
        const validationErrors = customXmlParts.items[0].validationErrors;
        context.load(validationErrors, "items");
        await context.sync();

        if (validationErrors.items.length > 0) {
            // Get the first validation error
            const error = validationErrors.items[0];
            context.load(error);
            await context.sync();

            // Convert the error to a plain JavaScript object
            const errorData = error.toJSON();
            
            // Log the serialized error data
            console.log("Validation Error Details:", JSON.stringify(errorData, null, 2));
        }
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.CustomXmlValidationError`

#### Examples

**Example**: Track a custom XML validation error object across multiple sync calls to safely access its properties after document changes occur.

```typescript
await Word.run(async (context) => {
    // Get the first custom XML part
    const customXmlParts = context.document.customXmlParts;
    customXmlParts.load("items");
    await context.sync();
    
    if (customXmlParts.items.length > 0) {
        const customXmlPart = customXmlParts.items[0];
        
        // Get validation errors
        const validationErrors = customXmlPart.validationErrors;
        validationErrors.load("items");
        await context.sync();
        
        if (validationErrors.items.length > 0) {
            const firstError = validationErrors.items[0];
            
            // Track the error object for use across sync calls
            firstError.track();
            
            // Load properties
            firstError.load("errorCode, errorType");
            await context.sync();
            
            // Make document changes that might affect object paths
            context.document.body.insertParagraph("Processing validation errors...", "Start");
            await context.sync();
            
            // Can still safely access the tracked error object
            console.log(`Error Code: ${firstError.errorCode}`);
            console.log(`Error Type: ${firstError.errorType}`);
            
            // Untrack when done
            firstError.untrack();
        }
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.CustomXmlValidationError`

#### Examples

**Example**: Retrieve custom XML validation errors, process them, and then untrack them to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection
    const customXmlParts = context.document.customXmlParts;
    context.load(customXmlParts);
    await context.sync();

    if (customXmlParts.items.length > 0) {
        const xmlPart = customXmlParts.items[0];
        
        // Get validation errors for the first custom XML part
        const validationErrors = xmlPart.validationErrors;
        validationErrors.load("items");
        await context.sync();

        // Process each validation error
        for (let i = 0; i < validationErrors.items.length; i++) {
            const error = validationErrors.items[i];
            error.load("errorCode, errorText");
            await context.sync();

            // Log error details
            console.log(`Error: ${error.errorText} (Code: ${error.errorCode})`);

            // Untrack the error object to free memory
            error.untrack();
        }

        await context.sync();
    }
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
