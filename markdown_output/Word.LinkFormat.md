# LinkFormat

**Package:** `Word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the linking characteristics for an OLE object or picture.

## Properties

### context

**Type:** `RequestContext`

**Since:** WordApi BETA (PREVIEW ONLY)

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the LinkFormat's RequestContext to verify the connection between the add-in and Word before performing operations on a linked object.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirst();
    const linkFormat = firstPicture.linkFormat;
    
    // Access the RequestContext associated with the LinkFormat object
    const linkContext = linkFormat.context;
    
    // Use the context to load properties and sync
    linkFormat.load("isLinked");
    await linkContext.sync();
    
    console.log("LinkFormat context is connected to Word");
    console.log(`Picture is linked: ${linkFormat.isLinked}`);
});
```

---

### isAutoUpdated

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the link is updated automatically when the container file is opened or when the source file is changed.

#### Examples

**Example**: Check if a linked image in the document is set to auto-update and display the result in the console

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirstOrNullObject();
    firstPicture.load("linkFormat");
    
    await context.sync();
    
    if (!firstPicture.isNullObject) {
        const linkFormat = firstPicture.linkFormat;
        linkFormat.load("isAutoUpdated");
        
        await context.sync();
        
        console.log(`Link is auto-updated: ${linkFormat.isAutoUpdated}`);
    } else {
        console.log("No pictures found in the document");
    }
});
```

---

### isLocked

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if a Field, InlineShape, or Shape object is locked to prevent automatic updating.

#### Examples

**Example**: Lock a linked picture to prevent it from being automatically updated when the source file changes

```typescript
await Word.run(async (context) => {
    // Get the first inline shape (picture) in the document
    const inlineShape = context.document.body.inlineShapes.getFirst();
    
    // Load the linkFormat property
    inlineShape.load("linkFormat");
    await context.sync();
    
    // Lock the linked picture to prevent automatic updates
    if (inlineShape.linkFormat) {
        inlineShape.linkFormat.isLocked = true;
        await context.sync();
        
        console.log("Linked picture has been locked from automatic updates");
    }
});
```

---

### isPictureSavedWithDocument

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies if the linked picture is saved with the document.

#### Examples

**Example**: Check if a linked picture in the document is saved with the document and display the result in the console.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const picture = inlinePictures.items[0];
        const linkFormat = picture.linkFormat;
        linkFormat.load("isPictureSavedWithDocument");
        
        await context.sync();
        
        console.log("Is picture saved with document: " + linkFormat.isPictureSavedWithDocument);
    }
});
```

---

### sourceFullName

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the path and name of the source file for the linked OLE object, picture, or field.

#### Examples

**Example**: Update the source file path of a linked image to point to a new location on the file system.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirst();
    
    // Load the link format properties
    firstPicture.load("linkFormat");
    await context.sync();
    
    // Update the source file path to a new location
    firstPicture.linkFormat.sourceFullName = "C:\\Images\\UpdatedLogo.png";
    
    await context.sync();
    
    console.log("Source file path updated successfully");
});
```

---

### sourceName

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the name of the source file for the linked OLE object, picture, or field.

#### Examples

**Example**: Get the source file name of a linked image in the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const picture = inlinePictures.items[0];
        
        // Get the link format and source name
        const linkFormat = picture.linkFormat;
        linkFormat.load("sourceName");
        
        await context.sync();
        
        // Display the source file name
        console.log("Source file name: " + linkFormat.sourceName);
    } else {
        console.log("No pictures found in the document.");
    }
});
```

---

### sourcePath

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the path of the source file for the linked OLE object, picture, or field.

#### Examples

**Example**: Get the source file path of a linked image in the document and display it in a content control.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const pictures = context.document.body.inlinePictures;
    pictures.load("items");
    
    await context.sync();
    
    if (pictures.items.length > 0) {
        const picture = pictures.items[0];
        
        // Get the link format and source path
        const linkFormat = picture.linkFormat;
        linkFormat.load("sourcePath");
        
        await context.sync();
        
        // Display the source path
        const sourcePath = linkFormat.sourcePath;
        console.log("Linked image source path: " + sourcePath);
        
        // Insert the path into the document
        context.document.body.insertParagraph(
            `Image source: ${sourcePath}`,
            Word.InsertLocation.end
        );
    }
    
    await context.sync();
});
```

---

### type

**Type:** `Word.LinkType | "Ole" | "Picture" | "Text" | "Reference" | "Include" | "Import" | "Dde" | "DdeAuto" | "Chart"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the link type.

#### Examples

**Example**: Check if a shape in the document contains a linked picture and display the link type

```typescript
await Word.run(async (context) => {
    // Get the first inline shape in the document
    const inlineShape = context.document.body.inlinePictures.getFirst();
    
    // Load the link format and its type property
    inlineShape.load("linkFormat");
    const linkFormat = inlineShape.linkFormat;
    linkFormat.load("type");
    
    await context.sync();
    
    // Check and display the link type
    console.log("Link type: " + linkFormat.type);
    
    // Perform actions based on link type
    if (linkFormat.type === Word.LinkType.picture || linkFormat.type === "Picture") {
        console.log("This is a linked picture");
    } else if (linkFormat.type === Word.LinkType.ole || linkFormat.type === "Ole") {
        console.log("This is an OLE object link");
    }
});
```

---

## Methods

### breakLink

**Kind:** `delete`

Breaks the link between the source file and the OLE object, picture, or linked field.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Break the link between a linked picture and its source file to embed it permanently in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirst();
    
    // Load the linkFormat property
    firstPicture.load("linkFormat");
    
    await context.sync();
    
    // Break the link to the source file
    firstPicture.linkFormat.breakLink();
    
    await context.sync();
    
    console.log("Link to source file has been broken. Picture is now embedded.");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.LinkFormatLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.LinkFormat`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.LinkFormat`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.LinkFormat`

#### Examples

**Example**: Load and read the AutoUpdate property of a link format to check if an inline picture's link automatically updates

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirst();
    
    // Get the link format of the picture
    const linkFormat = firstPicture.linkFormat;
    
    // Load the AutoUpdate property
    linkFormat.load("autoUpdate");
    
    // Sync to execute the load command
    await context.sync();
    
    // Read the loaded property
    console.log("Auto Update enabled: " + linkFormat.autoUpdate);
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.LinkFormatUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.LinkFormat` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple linking properties for a picture to make it auto-update and save the picture with the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirst();
    
    // Set multiple link format properties at once
    firstPicture.linkFormat.set({
        autoUpdate: true,
        saveWithDocument: true
    });
    
    await context.sync();
    console.log("Link format properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.LinkFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.LinkFormatData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.LinkFormatData`

#### Examples

**Example**: Get the link format properties of the first inline picture in the document and output them as a JSON string to the console.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirst();
    
    // Get the link format of the picture
    const linkFormat = firstPicture.linkFormat;
    
    // Load the link format properties
    linkFormat.load();
    
    await context.sync();
    
    // Convert the LinkFormat object to a plain JavaScript object
    const linkFormatData = linkFormat.toJSON();
    
    // Output as JSON string
    console.log(JSON.stringify(linkFormatData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.LinkFormat`

#### Examples

**Example**: Track a LinkFormat object to maintain its reference across multiple sync calls when modifying linked object properties

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirst();
    
    // Get the LinkFormat object
    const linkFormat = firstPicture.linkFormat;
    
    // Track the LinkFormat object for use across sync calls
    linkFormat.track();
    
    // Load properties
    linkFormat.load("isLinked");
    
    await context.sync();
    
    // Now we can safely use the object across multiple syncs
    console.log("Is linked: " + linkFormat.isLinked);
    
    // Perform additional operations with the tracked object
    await context.sync();
    
    // Untrack when done to release memory
    linkFormat.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.LinkFormat`

#### Examples

**Example**: Release memory for a tracked LinkFormat object after reading its properties to avoid slowing down the host application.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirstOrNullObject();
    const linkFormat = firstPicture.linkFormat;
    
    // Track the object to work with it
    linkFormat.track();
    
    // Load properties to use
    linkFormat.load("autoUpdate");
    await context.sync();
    
    // Use the linkFormat object
    console.log("Auto update enabled: " + linkFormat.autoUpdate);
    
    // Release the memory after we're done using it
    linkFormat.untrack();
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
