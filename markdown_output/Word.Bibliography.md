# Word.Bibliography

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the list of available sources attached to the document (in the current list) or the list of sources available in the application (in the master list).

## Properties

### bibliographyStyle

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the name of the active style to use for the bibliography.

#### Examples

**Example**: Set the bibliography style to APA (American Psychological Association) format for the document's bibliography.

```typescript
await Word.run(async (context) => {
    const bibliography = context.document.bibliography;
    bibliography.bibliographyStyle = "APA";
    
    await context.sync();
    console.log("Bibliography style set to APA");
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the bibliography's request context to verify the add-in is properly connected to the Word host application before performing bibliography operations.

```typescript
await Word.run(async (context) => {
    const bibliography = context.document.bibliography;
    
    // Access the request context associated with the bibliography object
    const requestContext = bibliography.context;
    
    // Verify the context is valid by checking if it's defined
    if (requestContext) {
        console.log("Bibliography is connected to Word host application");
        
        // Load bibliography properties using the context
        bibliography.load("sources");
        await context.sync();
        
        console.log(`Number of sources: ${bibliography.sources.items.length}`);
    }
});
```

---

### sources

**Type:** `Word.SourceCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a SourceCollection object that represents all the sources contained in the bibliography.

#### Examples

**Example**: Get all sources from the document's bibliography and display the count and title of each source in the console.

```typescript
await Word.run(async (context) => {
    // Get the bibliography from the document
    const bibliography = context.document.bibliography;
    
    // Get the sources collection
    const sources = bibliography.sources;
    
    // Load the count and title properties
    sources.load("items/title");
    
    await context.sync();
    
    // Display the count and each source title
    console.log(`Total sources: ${sources.items.length}`);
    
    sources.items.forEach((source, index) => {
        console.log(`Source ${index + 1}: ${source.title}`);
    });
});
```

---

## Methods

### generateUniqueTag

Generates a unique identification tag for a bibliography source and returns a string that represents the tag.

#### Signature

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Generate a unique tag for a new bibliography source before adding it to the document's bibliography

```typescript
await Word.run(async (context) => {
    const bibliography = context.document.bibliography;
    
    // Generate a unique tag for a new source
    const uniqueTag = bibliography.generateUniqueTag();
    
    await context.sync();
    
    console.log("Generated unique tag for new bibliography source:", uniqueTag);
    
    // The tag can now be used when creating a new bibliography source
    // Example: "Smi23" or "Jon24" depending on existing sources
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.BibliographyLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Bibliography`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Bibliography`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `object` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Bibliography`

#### Examples

**Example**: Load and display the number of sources in the document's bibliography

```typescript
await Word.run(async (context) => {
    const bibliography = context.document.bibliography;
    
    // Load the sources property of the bibliography
    bibliography.load("sources");
    
    await context.sync();
    
    // Display the count of bibliography sources
    console.log(`Number of bibliography sources: ${bibliography.sources.items.length}`);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.BibliographyUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Bibliography` (required)

  **Returns:** `void`

#### Examples

**Example**: Update bibliography settings to set the bibliography style to APA and enable automatic updates

```typescript
await Word.run(async (context) => {
    const bibliography = context.document.bibliography;
    
    bibliography.set({
        bibliographyStyle: "APA",
        automaticUpdate: true
    });
    
    await context.sync();
    console.log("Bibliography settings updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.Bibliography object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BibliographyData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BibliographyData`

#### Examples

**Example**: Serialize the document's bibliography data to JSON format and log it to the console for debugging or export purposes.

```typescript
await Word.run(async (context) => {
    // Get the bibliography from the document
    const bibliography = context.document.bibliography;
    
    // Load the bibliography properties
    bibliography.load("*");
    
    await context.sync();
    
    // Convert the bibliography to a plain JavaScript object
    const bibliographyData = bibliography.toJSON();
    
    // Log the JSON representation
    console.log("Bibliography data:", JSON.stringify(bibliographyData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Bibliography`

#### Examples

**Example**: Track a bibliography object to maintain its reference across multiple sync calls while adding sources and updating properties

```typescript
await Word.run(async (context) => {
    const bibliography = context.document.bibliography;
    
    // Track the bibliography object to use it across multiple sync calls
    bibliography.track();
    
    // Load properties
    bibliography.load("sources");
    await context.sync();
    
    // Use the bibliography object after sync (tracking prevents InvalidObjectPath error)
    console.log(`Current source count: ${bibliography.sources.items.length}`);
    
    // Perform another sync and continue using the tracked object
    await context.sync();
    
    // Access the bibliography again safely
    console.log("Bibliography is still accessible after multiple syncs");
    
    // Untrack when done to free up memory
    bibliography.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Bibliography`

#### Examples

**Example**: Access the document's bibliography, perform operations with it, then untrack it to free memory after use.

```typescript
await Word.run(async (context) => {
    // Get the bibliography and track it
    const bibliography = context.document.bibliography;
    bibliography.track();
    
    // Load properties to work with the bibliography
    bibliography.load("sources");
    await context.sync();
    
    // Perform operations (e.g., log the number of sources)
    console.log(`Number of sources: ${bibliography.sources.items.length}`);
    
    // Untrack the bibliography to release memory
    bibliography.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.sourcecollection
- /en-us/javascript/api/office/officeextension.clientresult
- /en-us/javascript/api/word/word.interfaces.bibliographyloadoptions
- /en-us/javascript/api/word/word.bibliography
- /en-us/javascript/api/word/word.interfaces.bibliographyupdatedata
- /en-us/javascript/api/office/officeextension.updateoptions
- /en-us/javascript/api/word/word.interfaces.bibliographydata
- /en-us/javascript/api/office/officeextension.clientrequestcontext
