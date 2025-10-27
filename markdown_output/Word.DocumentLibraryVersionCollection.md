# Word.DocumentLibraryVersionCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the collection of [Word.DocumentLibraryVersion](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversion) objects.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a DocumentLibraryVersionCollection to verify the connection between the add-in and Word host application.

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const versions = context.document.properties.documentLibraryVersions;
    
    // Access the request context associated with the collection
    const requestContext = versions.context;
    
    // Verify the context is available and connected to the Word application
    console.log("Request context available:", requestContext !== null);
    console.log("Context type:", requestContext.constructor.name);
    
    // The context can be used to sync changes or manage the connection
    await context.sync();
    
    console.log("Successfully accessed the DocumentLibraryVersionCollection context");
});
```

---

### items

**Type:** `Word.DocumentLibraryVersion[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Display the version history of the current document by iterating through all available library versions and logging their version numbers and modification dates.

```typescript
await Word.run(async (context) => {
    // Get the document library version collection
    const versions = context.document.documentLibraryVersions;
    
    // Load the items property along with version details
    versions.load("items");
    await context.sync();
    
    // Access the items array to iterate through all versions
    const versionItems = versions.items;
    
    console.log(`Total versions: ${versionItems.length}`);
    
    // Loop through each version in the collection
    for (let i = 0; i < versionItems.length; i++) {
        const version = versionItems[i];
        version.load("versionIndex, modified");
        await context.sync();
        
        console.log(`Version ${version.versionIndex}: Modified on ${version.modified}`);
    }
});
```

---

## Methods

### getItem

**Kind:** `read`

Gets a DocumentLibraryVersion object by its index in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  The location of a DocumentLibraryVersion object.

**Returns:** `Word.DocumentLibraryVersion`

#### Examples

**Example**: Get and display the comment from the third version in the document's version history

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const versions = context.document.documentLibraryVersions;
    
    // Get the third version (index 2) from the collection
    const thirdVersion = versions.getItem(2);
    
    // Load the comment property of that version
    thirdVersion.load("comment");
    
    await context.sync();
    
    // Display the version comment
    console.log("Version 3 comment: " + thirdVersion.comment);
});
```

---

### isVersioningEnabled

**Kind:** `read`

Returns whether the document library in which the active document is saved on the server is configured to create a backup copy, or version, each time the file is edited on the website.

#### Signature

**Returns:** `OfficeExtension.ClientResult<boolean>`

#### Examples

**Example**: Check if versioning is enabled for the document library and display the result in the console

```typescript
await Word.run(async (context) => {
    const documentLibraryVersions = context.document.properties.documentLibraryVersions;
    const isEnabled = documentLibraryVersions.isVersioningEnabled();
    
    await context.sync();
    
    console.log(`Document library versioning is ${isEnabled.value ? 'enabled' : 'disabled'}`);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.DocumentLibraryVersionCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.DocumentLibraryVersionCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.DocumentLibraryVersionCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.DocumentLibraryVersionCollection`

#### Examples

**Example**: Load and display version information for all versions of the current document stored in a SharePoint library

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const versions = context.document.properties.documentLibraryVersions;
    
    // Load specific properties of all versions in the collection
    versions.load("items/versionIndex, items/modifiedDate, items/modifiedBy");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display version information
    console.log(`Total versions: ${versions.items.length}`);
    versions.items.forEach(version => {
        console.log(`Version ${version.versionIndex}: Modified on ${version.modifiedDate} by ${version.modifiedBy}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DocumentLibraryVersionCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DocumentLibraryVersionCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.DocumentLibraryVersionCollectionData`

#### Examples

**Example**: Retrieve all document library versions and export them as a plain JavaScript object for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const versions = context.document.properties.documentLibraryVersions;
    
    // Load the properties we want to access
    versions.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const versionsData = versions.toJSON();
    
    // Now you can use the plain object (e.g., log it, store it, send it to a server)
    console.log("Document versions:", versionsData);
    console.log("Number of versions:", versionsData.items.length);
    
    // The versionsData object can be safely serialized
    const jsonString = JSON.stringify(versionsData, null, 2);
    console.log("Serialized versions:", jsonString);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.DocumentLibraryVersionCollection`

#### Examples

**Example**: Track a document library version collection to maintain a reference across multiple sync calls when checking for version updates

```typescript
await Word.run(async (context) => {
    const document = context.document;
    const versions = document.properties.documentLibraryVersions;
    
    // Track the collection to use it across multiple sync calls
    versions.track();
    
    // Load version information
    versions.load("items");
    await context.sync();
    
    // Process versions (can safely use the collection after sync)
    console.log(`Total versions: ${versions.items.length}`);
    
    // Perform another sync operation
    await context.sync();
    
    // Collection is still valid because it's tracked
    for (let version of versions.items) {
        version.load("versionIndex, comments");
    }
    await context.sync();
    
    // Untrack when done to free up resources
    versions.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.DocumentLibraryVersionCollection`

#### Examples

**Example**: Load document library versions, use them to display version information, then untrack the collection to free memory

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const versions = context.document.properties.documentLibraryVersions;
    
    // Load the collection and its properties
    versions.load("items");
    await context.sync();
    
    // Use the versions data (e.g., log version count)
    console.log(`Total versions: ${versions.items.length}`);
    
    // Untrack the collection to release memory
    versions.untrack();
    await context.sync();
    
    console.log("Version collection memory released");
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
- https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversion
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext
- https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversioncollection
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.documentlibraryversioncollectionloadoptions
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.collectionloadoptions
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.documentlibraryversioncollectiondata
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
