# BuildingBlockEntryCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of building block entries in a Word template.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the building block entries collection and verify the context is properly connected before performing operations on building blocks in the template.

```typescript
await Word.run(async (context) => {
    // Get the building block entries collection from the first gallery
    const gallery = context.document.properties.customXmlParts.getByNamespace("http://schemas.microsoft.com/office/word/2010/wordml")[0];
    const buildingBlockEntries = context.document.buildingBlockEntries;
    
    // Access the context property to ensure connection to Office host
    const requestContext = buildingBlockEntries.context;
    
    // Verify the context is valid by checking if it's the same as the current context
    if (requestContext === context) {
        console.log("Building block entries context is properly connected to Office host");
        
        // Now safe to perform operations on the collection
        buildingBlockEntries.load("items");
        await context.sync();
        
        console.log(`Found ${buildingBlockEntries.items.length} building blocks`);
    }
});
```

---

## Methods

### add

**Kind:** `create`

Creates a new building block entry in a template and returns a BuildingBlock object that represents the new building block entry.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `name`: `string` (required)
    The name of the building block.
  - `type`: `Word.BuildingBlockType` (required)
    The type of the building block.
  - `category`: `string` (required)
    The category of the building block.
  - `range`: `Word.Range` (required)
    The range to insert the building block.
  - `description`: `string` (required)
    The description of the building block.
  - `insertType`: `Word.DocPartInsertType` (required)
    How to insert the contents of the building block.

  **Returns:** `Word.BuildingBlock`

**Overload 2:**

  **Parameters:**
  - `name`: `string` (required)
    The name of the building block.
  - `type`: `"QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography"` (required)
    The type of the building block.
  - `category`: `string` (required)
    The category of the building block.
  - `range`: `Word.Range` (required)
    The range to insert the building block.
  - `description`: `string` (required)
    The description of the building block.
  - `insertType`: `"Content" | "Paragraph" | "Page"` (required)
    How to insert the contents of the building block.

  **Returns:** `Word.BuildingBlock`

#### Examples

**Example**: Create a new building block entry named "Company Header" that stores the selected text as a reusable AutoText block in the General category

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const range = context.document.getSelection();
    
    // Get the building blocks collection from the first template
    const template = context.document.getDefaultTemplate();
    const buildingBlocks = template.buildingBlockEntries;
    
    // Create a new building block entry
    const newBlock = buildingBlocks.add(
        "Company Header",                    // name
        Word.BuildingBlockType.autoText,     // type
        "General",                           // category
        range,                               // range
        "Standard company header text",      // description
        Word.InsertLocation.replace          // insertType
    );
    
    await context.sync();
    console.log("Building block created successfully");
});
```

---

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Display the total number of available building blocks in the "General" gallery to the user in the console.

```typescript
await Word.run(async (context) => {
    const template = context.document.getActiveSectionOrNullObject().body;
    const buildingBlockEntries = context.application.buildingBlockTemplates
        .getFirst()
        .buildingBlockEntries;
    
    const count = buildingBlockEntries.getCount();
    
    await context.sync();
    
    console.log(`Total building blocks available: ${count.value}`);
});
```

---

### getItemAt

**Kind:** `read`

Returns a BuildingBlock object that represents the specified item in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  The index of the item to retrieve.

**Returns:** `Word.BuildingBlock`

#### Examples

**Example**: Get the third building block entry from a collection and insert it into the document at the current selection.

```typescript
await Word.run(async (context) => {
    // Get the building blocks from the first template
    const template = context.document.getTemplate();
    const buildingBlockEntries = template.buildingBlockEntries;
    
    // Get the building block at index 2 (third item)
    const buildingBlock = buildingBlockEntries.getItemAt(2);
    buildingBlock.load("name");
    
    // Insert the building block at the current selection
    const range = context.document.getSelection();
    buildingBlock.insertContent(range, Word.InsertLocation.replace);
    
    await context.sync();
    console.log(`Inserted building block: ${buildingBlock.name}`);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BuildingBlockEntryCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BuildingBlockEntryCollection`

#### Examples

**Example**: Load and display the names of all building block entries in the first gallery of the template

```typescript
await Word.run(async (context) => {
    // Get the first gallery's building block entries
    const gallery = context.document.buildingBlockGalleries.getFirst();
    const buildingBlocks = gallery.buildingBlockEntries;
    
    // Load the 'name' property for all entries in the collection
    buildingBlocks.load("name");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the names of all building block entries
    console.log(`Found ${buildingBlocks.items.length} building blocks:`);
    buildingBlocks.items.forEach(block => {
        console.log(`- ${block.name}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockEntryCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockEntryCollectionData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `{ [key: string]: string; }`

#### Examples

**Example**: Export building block entries from a template to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the building block entries collection from the first gallery
    const gallery = context.document.getDefaultTemplate().galleries.getFirst();
    const buildingBlocks = gallery.buildingBlockEntries;
    
    // Load properties needed for JSON export
    buildingBlocks.load("name,type,category");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const buildingBlocksJSON = buildingBlocks.toJSON();
    
    // Log or store the JSON representation
    console.log(JSON.stringify(buildingBlocksJSON, null, 2));
    
    // The JSON can now be used outside the Word context
    // e.g., sent to a server, saved to local storage, etc.
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BuildingBlockEntryCollection`

#### Examples

**Example**: Track a building block entry collection across multiple sync calls to prevent InvalidObjectPath errors when accessing the collection after document changes.

```typescript
await Word.run(async (context) => {
    // Get the building block entry collection from a gallery
    const gallery = context.document.getDefaultBuildingBlockGallery();
    const buildingBlocks = gallery.getBuildingBlockEntries();
    
    // Track the collection to use it across multiple sync calls
    buildingBlocks.track();
    
    // Load properties
    buildingBlocks.load("items");
    await context.sync();
    
    // First sync - access the collection
    console.log(`Found ${buildingBlocks.items.length} building blocks`);
    
    // Perform some document changes
    context.document.body.insertParagraph("New content", Word.InsertLocation.start);
    await context.sync();
    
    // Second sync - can still safely access the tracked collection
    for (const block of buildingBlocks.items) {
        block.load("name");
    }
    await context.sync();
    
    // Untrack when done to free memory
    buildingBlocks.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BuildingBlockEntryCollection`

#### Examples

**Example**: Retrieve building block entries from a gallery, use them to get information, then untrack the collection to free memory after processing is complete.

```typescript
await Word.run(async (context) => {
    // Get building blocks from a specific gallery
    const template = context.document.getDefaultTemplate();
    const gallery = template.buildingBlockGalleries.getByName("Quick Parts");
    const buildingBlocks = gallery.buildingBlockEntries;
    
    // Load properties to use the collection
    buildingBlocks.load("items");
    await context.sync();
    
    // Process the building blocks (e.g., log their names)
    console.log(`Found ${buildingBlocks.items.length} building blocks`);
    buildingBlocks.items.forEach(block => {
        console.log(block.name);
    });
    
    // Untrack the collection to release memory
    buildingBlocks.untrack();
    await context.sync();
    
    console.log("Building block collection memory released");
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
