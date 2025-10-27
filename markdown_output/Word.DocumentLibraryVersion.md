# Word.DocumentLibraryVersion

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a document library version.

## Properties

### comments

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets any optional comments associated with this version of the shared document.

#### Examples

**Example**: Retrieve and display the comments associated with a specific version of a shared document from the document library.

```typescript
await Word.run(async (context) => {
    // Get the document library versions
    const documentLibraryVersions = context.document.documentLibraryVersions;
    documentLibraryVersions.load("items");
    await context.sync();

    // Get the first version and load its comments
    if (documentLibraryVersions.items.length > 0) {
        const version = documentLibraryVersions.items[0];
        version.load("comments");
        await context.sync();

        // Display the comments associated with this version
        console.log("Version comments: " + version.comments);
    } else {
        console.log("No document library versions found.");
    }
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a DocumentLibraryVersion object to verify the connection between the add-in and Word before performing version operations.

```typescript
await Word.run(async (context) => {
    // Get the document library versions
    const documentLibraryVersions = context.document.documentLibraryVersions;
    context.load(documentLibraryVersions, "items");
    await context.sync();
    
    if (documentLibraryVersions.items.length > 0) {
        const version = documentLibraryVersions.items[0];
        
        // Access the request context from the version object
        const versionContext = version.context;
        
        // Use the context to load version properties
        versionContext.load(version, "versionNumber, comments");
        await versionContext.sync();
        
        console.log(`Version: ${version.versionNumber}`);
        console.log(`Comments: ${version.comments}`);
    }
});
```

---

### modified

**Type:** `any`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the date and time at which this version of the shared document was last saved to the server.

#### Examples

**Example**: Display the last modified date of the current document's latest version in the console.

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const documentLibraryVersions = context.document.documentLibraryVersions;
    
    // Load the items in the collection
    documentLibraryVersions.load("items");
    await context.sync();
    
    // Get the most recent version (first item)
    if (documentLibraryVersions.items.length > 0) {
        const latestVersion = documentLibraryVersions.items[0];
        
        // Load the modified property
        latestVersion.load("modified");
        await context.sync();
        
        // Display the modified date
        console.log("Last modified: " + latestVersion.modified);
    } else {
        console.log("No versions found");
    }
});
```

---

### modifiedBy

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the name of the user who last saved this version of the shared document to the server.

#### Examples

**Example**: Display the name of the user who last modified the current document version in a content control.

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const documentLibrary = context.document.properties.documentLibrary;
    const versions = documentLibrary.versions;
    
    // Load the current version
    const currentVersion = versions.getItemAt(0);
    currentVersion.load("modifiedBy");
    
    await context.sync();
    
    // Insert the modified by information into the document
    const contentControl = context.document.body.insertContentControl();
    contentControl.insertText(
        `Last modified by: ${currentVersion.modifiedBy}`,
        Word.InsertLocation.start
    );
    
    await context.sync();
});
```

---

## Methods

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.DocumentLibraryVersionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.DocumentLibraryVersion`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.DocumentLibraryVersion`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.DocumentLibraryVersion`

#### Examples

**Example**: Load and display the version number and comments of a document library version.

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const documentVersions = context.document.documentLibraryVersions;
    context.load(documentVersions);
    await context.sync();
    
    // Get the first version
    const firstVersion = documentVersions.items[0];
    
    // Load specific properties of the version
    firstVersion.load("versionNumber, comments");
    await context.sync();
    
    // Display the loaded properties
    console.log(`Version Number: ${firstVersion.versionNumber}`);
    console.log(`Comments: ${firstVersion.comments}`);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DocumentLibraryVersion object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DocumentLibraryVersionData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.DocumentLibraryVersionData`

#### Examples

**Example**: Retrieve a document library version object and serialize it to JSON format for logging or external storage.

```typescript
await Word.run(async (context) => {
    // Get the document's library versions collection
    const documentVersions = context.document.properties.documentLibraryVersions;
    
    // Load the first version
    const firstVersion = documentVersions.getFirst();
    firstVersion.load("id,comments,created,createdBy");
    
    await context.sync();
    
    // Convert the version object to a plain JavaScript object
    const versionData = firstVersion.toJSON();
    
    // Now you can use the plain object for logging, storage, or transmission
    console.log("Version as JSON:", JSON.stringify(versionData, null, 2));
    console.log("Version ID:", versionData.id);
    console.log("Created:", versionData.created);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.DocumentLibraryVersion`

#### Examples

**Example**: Track a document library version object to use it across multiple sync calls when checking version properties

```typescript
await Word.run(async (context) => {
    // Get the document library versions
    const documentLibrary = context.document.properties.documentLibrary;
    const versions = documentLibrary.versions;
    versions.load("items");
    await context.sync();
    
    // Get the first version and track it for use across sync calls
    if (versions.items.length > 0) {
        const firstVersion = versions.items[0];
        firstVersion.track();
        
        // Load properties
        firstVersion.load("versionNumber, comments");
        await context.sync();
        
        // Use the tracked object after sync
        console.log(`Version: ${firstVersion.versionNumber}`);
        console.log(`Comments: ${firstVersion.comments}`);
        
        // Can safely use across multiple syncs
        await context.sync();
        console.log(`Still accessible: ${firstVersion.versionNumber}`);
        
        // Untrack when done
        firstVersion.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.DocumentLibraryVersion`

#### Examples

**Example**: Retrieve document library version information, use it to display version details, then untrack the object to free memory.

```typescript
await Word.run(async (context) => {
    // Get the document library versions collection
    const versions = context.document.properties.documentLibraryVersions;
    versions.load("items");
    await context.sync();
    
    if (versions.items.length > 0) {
        // Get the first version
        const version = versions.items[0];
        version.load("versionNumber, comments");
        await context.sync();
        
        // Use the version information
        console.log(`Version: ${version.versionNumber}`);
        console.log(`Comments: ${version.comments}`);
        
        // Untrack the version object to release memory
        version.untrack();
        await context.sync();
        
        console.log("Version object memory released");
    }
});
```

---

## Source

- /en-us/javascript/api/word/word.documentlibraryversion
