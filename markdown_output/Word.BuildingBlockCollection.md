# Word.BuildingBlockCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of [Word.BuildingBlock](/en-us/javascript/api/word/word.buildingblock) objects for a specific building block type and category in a template.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the building block collection's request context to verify the add-in is properly connected to the Word host application before performing operations.

```typescript
await Word.run(async (context) => {
    // Get the building blocks collection for a specific type and category
    const template = context.document.getTemplate();
    const buildingBlockCollection = template.getBuildingBlocksByCategory("General", Word.BuildingBlockType.autoText);
    
    // Access the request context associated with the collection
    const requestContext = buildingBlockCollection.context;
    
    // Verify the context is valid and connected
    if (requestContext && requestContext.application) {
        console.log("Building block collection is properly connected to Word application");
        
        // Use the context to load and sync data
        buildingBlockCollection.load("items");
        await context.sync();
        
        console.log(`Found ${buildingBlockCollection.items.length} building blocks`);
    }
});
```

---

## Methods

### add

**Kind:** `create`

Creates a new building block and returns a BuildingBlock object.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `name`: `string` (required)
    The name of the building block.
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
  - `range`: `Word.Range` (required)
    The range to insert the building block.
  - `description`: `string` (required)
    The description of the building block.
  - `insertType`: `"Content" | "Paragraph" | "Page"` (required)
    How to insert the contents of the building block.

  **Returns:** `Word.BuildingBlock`

#### Examples

**Example**: Create a new building block from the currently selected text in the document, naming it "Company Header" with a description for reuse

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const selection = context.document.getSelection();
    
    // Get the building blocks collection for a specific type and category
    const template = context.document.getActiveTemplate();
    const buildingBlocks = template.getBuildingBlocksByCategory("General");
    
    // Create a new building block from the selection
    const newBuildingBlock = buildingBlocks.add(
        "Company Header",
        selection,
        "Standard company header with logo and contact info",
        Word.InsertLocation.replace
    );
    
    // Load properties to verify creation
    newBuildingBlock.load("name");
    
    await context.sync();
    console.log(`Building block created: ${newBuildingBlock.name}`);
});
```

---

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Get and display the number of AutoText building blocks available in the "General" category

```typescript
await Word.run(async (context) => {
    // Get the building blocks collection for AutoText in the General category
    const buildingBlockCollection = context.application.getTemplate()
        .getBuildingBlocksByCategory("AutoText", "General");
    
    // Get the count of building blocks in the collection
    const count = buildingBlockCollection.getCount();
    
    // Load the count value
    await context.sync();
    
    // Display the count
    console.log(`Number of AutoText building blocks in General category: ${count.value}`);
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

**Example**: Get and insert the third building block from the "General" category of "Quick Parts" type into the document at the current selection.

```typescript
await Word.run(async (context) => {
    // Get the first template (active document's template)
    const template = context.document.getTemplate();
    
    // Get the building block collection for "Quick Parts" type and "General" category
    const buildingBlocks = template.getBuildingBlocksByCategory("Quick Parts", "General");
    
    // Get the building block at index 2 (third item)
    const buildingBlock = buildingBlocks.getItemAt(2);
    
    // Load the building block's name property
    buildingBlock.load("name");
    
    await context.sync();
    
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

  **Returns:** `Word.BuildingBlockCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BuildingBlockCollection`

#### Examples

**Example**: Load and display the names of all building blocks in the "General" category of type "AutoText"

```typescript
await Word.run(async (context) => {
    // Get the building blocks collection for AutoText type and General category
    const buildingBlockCollection = context.application.getTemplate()
        .getBuildingBlocksByType(Word.BuildingBlockType.autoText)
        .getByCategory("General");
    
    // Load the 'name' property for all building blocks in the collection
    buildingBlockCollection.load("items/name");
    
    await context.sync();
    
    // Display the names of the building blocks
    buildingBlockCollection.items.forEach(block => {
        console.log(`Building Block Name: ${block.name}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCollectionData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `{ [key: string]: string; }`

#### Examples

**Example**: Serialize building blocks from a specific category to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get building blocks from a specific type and category
    const buildingBlockCollection = context.application.getTemplate()
        .getBuildingBlocksByType(Word.BuildingBlockTypes.autoText)
        .getByCategory("General");
    
    // Load properties needed for serialization
    buildingBlockCollection.load("items/name,items/type,items/category");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const jsonData = buildingBlockCollection.toJSON();
    
    // Now you can stringify it for logging or storage
    console.log(JSON.stringify(jsonData, null, 2));
    
    // The jsonData object contains shallow copies of loaded properties
    // and can be safely used outside the Word.run context
    return jsonData;
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BuildingBlockCollection`

#### Examples

**Example**: Track a building block collection to safely access its items across multiple sync calls when working with AutoText building blocks from the Normal template.

```typescript
await Word.run(async (context) => {
    // Get the building block collection for AutoText in the Normal template
    const buildingBlockCollection = context.application.getBuiltInStylesAsync()
        .then(() => context.application.getActiveDocument().getBody())
        .then(() => context.application.getBuiltInBuildingBlockCollections());
    
    // More practical example:
    const template = context.application.getActiveDocument();
    const buildingBlocks = template.getBuildingBlockCollections()
        .getByType(Word.BuildingBlockType.autoText)
        .getFirstOrNullObject()
        .getByCategory("General");
    
    // Track the collection for use across multiple sync calls
    buildingBlocks.track();
    
    await context.sync();
    
    // Now safe to use the collection in subsequent operations
    buildingBlocks.load("items");
    await context.sync();
    
    console.log(`Found ${buildingBlocks.items.length} AutoText building blocks`);
    
    // Perform additional operations with the tracked collection
    if (buildingBlocks.items.length > 0) {
        buildingBlocks.items[0].load("name");
        await context.sync();
        console.log(`First building block: ${buildingBlocks.items[0].name}`);
    }
    
    // Untrack when done
    buildingBlocks.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BuildingBlockCollection`

#### Examples

**Example**: Load building blocks from a template, use them to get their names, then untrack the collection to free memory after processing is complete.

```typescript
await Word.run(async (context) => {
    // Get building blocks of a specific type and category
    const template = context.document.getDefaultTemplate();
    const buildingBlocks = template.getBuildingBlocksByCategory("General", Word.BuildingBlockType.autoText);
    
    // Load and track the collection
    buildingBlocks.load("items");
    await context.sync();
    
    // Process the building blocks
    console.log(`Found ${buildingBlocks.items.length} building blocks`);
    buildingBlocks.items.forEach(block => {
        console.log(block.name);
    });
    
    // Release memory by untracking the collection when done
    buildingBlocks.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/word/word.buildingblock
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.range
- /en-us/javascript/api/word/word.docpartinserttype
- /en-us/javascript/api/office/officeextension.clientresult
- /en-us/javascript/api/word/word.buildingblockcollection
- /en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
