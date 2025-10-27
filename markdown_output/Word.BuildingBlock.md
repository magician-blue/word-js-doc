# BuildingBlock

**Package:** `Word`

**API Set:** WordApi BETA

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a building block in a template. A building block is pre-built content, similar to autotext, that may contain text, images, and formatting.

## Properties

### category

**Type:** `Word.BuildingBlockCategory`

**Since:** WordApi BETA

Returns a BuildingBlockCategory object that represents the category for the building block.

#### Examples

**Example**: Get the category name of the first building block in the template and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first building block from the template
    const buildingBlocks = context.document.getBuiltInBuildingBlocks();
    const firstBlock = buildingBlocks.getFirst();
    
    // Get the category of the building block
    const category = firstBlock.category;
    
    // Load the category name
    category.load("name");
    
    await context.sync();
    
    // Display the category name
    console.log(`Building block category: ${category.name}`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access a building block's request context to verify the connection between the add-in and Word host application by logging the context debug information.

```typescript
await Word.run(async (context) => {
    // Get a building block from the gallery
    const gallery = context.document.buildingBlockGalleries.getByName("Quick Parts");
    const buildingBlocks = gallery.buildingBlocks;
    buildingBlocks.load("items");
    
    await context.sync();
    
    if (buildingBlocks.items.length > 0) {
        const buildingBlock = buildingBlocks.items[0];
        
        // Access the request context associated with the building block
        const requestContext = buildingBlock.context;
        
        // Use the context to verify connection (e.g., check if it's the same as the main context)
        console.log("Building block context exists:", requestContext !== null);
        console.log("Context matches main context:", requestContext === context);
        
        // The context can be used for operations requiring the same request context
        buildingBlock.load("name");
        await requestContext.sync();
        
        console.log("Building block name:", buildingBlock.name);
    }
});
```

---

### description

**Type:** `string`

**Since:** WordApi BETA

Specifies the description for the building block.

#### Examples

**Example**: Set the description of a building block to "Company header with logo and contact information"

```typescript
await Word.run(async (context) => {
    // Get the first building block from the template
    const buildingBlock = context.document.buildingBlockLists.getFirst().buildingBlocks.getFirst();
    
    // Set the description for the building block
    buildingBlock.description = "Company header with logo and contact information";
    
    await context.sync();
});
```

---

### id

**Type:** `string`

**Since:** WordApi BETA

Returns the internal identification number for the building block.

#### Examples

**Example**: Retrieve and display the internal identification number of a building block from the document's gallery

```typescript
await Word.run(async (context) => {
    // Get the first building block from the first gallery
    const gallery = context.document.buildingBlockGalleries.getFirst();
    const buildingBlock = gallery.buildingBlocks.getFirst();
    
    // Load the id property
    buildingBlock.load("id");
    
    await context.sync();
    
    // Display the building block's internal identification number
    console.log(`Building block ID: ${buildingBlock.id}`);
});
```

---

### index

**Type:** `number`

**Since:** WordApi BETA

Returns the position of this building block in a collection.

#### Examples

**Example**: Get the position of a specific building block named "Company Header" in the collection and display it to the user.

```typescript
await Word.run(async (context) => {
    // Get all building blocks from the template
    const buildingBlocks = context.document.buildingBlocksCollection.getByName("Company Header");
    buildingBlocks.load("index");
    
    await context.sync();
    
    // Display the position of the building block in the collection
    console.log(`The "Company Header" building block is at position: ${buildingBlocks.index}`);
});
```

---

### insertType

**Type:** `Word.DocPartInsertType | "Content" | "Paragraph" | "Page"`

**Since:** WordApi BETA

Specifies a DocPartInsertType value that represents how to insert the contents of the building block into the document.

#### Examples

**Example**: Set a building block's insert type to "Paragraph" so that when inserted, it will be placed as a complete paragraph in the document.

```typescript
await Word.run(async (context) => {
    // Get the first building block from the first template
    const template = context.document.getBuiltInBuildingBlockTemplates().getFirst();
    const buildingBlock = template.buildingBlocks.getFirst();
    
    // Set the insert type to Paragraph
    buildingBlock.insertType = Word.DocPartInsertType.paragraph;
    
    await context.sync();
    
    console.log("Building block insert type set to Paragraph");
});
```

---

### name

**Type:** `string`

**Since:** WordApi BETA

Specifies the name of the building block.

#### Examples

**Example**: Get the name of the first building block in the "General" gallery and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the building blocks from the General gallery
    const buildingBlockEntries = context.application.templates.getFirst()
        .buildingBlockGalleries.getByName("General").buildingBlocks;
    
    // Load the first building block's name
    const firstBlock = buildingBlockEntries.getFirst();
    firstBlock.load("name");
    
    await context.sync();
    
    // Display the building block name
    console.log("Building block name: " + firstBlock.name);
});
```

---

### type

**Type:** `Word.BuildingBlockTypeItem`

**Since:** WordApi BETA

Returns a BuildingBlockTypeItem object that represents the type for the building block.

#### Examples

**Example**: Get the building block type and display it to the user to verify if it's a "Quick Parts" entry before inserting it into the document.

```typescript
await Word.run(async (context) => {
    // Get the first building block from the first gallery
    const gallery = context.document.buildingBlockGalleries.getFirst();
    const buildingBlock = gallery.buildingBlocks.getFirst();
    
    // Load the type property
    buildingBlock.load("type");
    
    await context.sync();
    
    // Access the type information
    const blockType = buildingBlock.type;
    console.log(`Building block type: ${blockType}`);
    
    // Use the type to make decisions
    if (blockType === Word.BuildingBlockType.quickParts) {
        console.log("This is a Quick Parts building block");
    }
});
```

---

### value

**Type:** `string`

**Since:** WordApi BETA

Specifies the contents of the building block.

#### Examples

**Example**: Set the contents of a building block to include formatted text with a title and description for a company disclaimer.

```typescript
await Word.run(async (context) => {
    // Get the building block from the template
    const buildingBlock = context.document.buildingBlockLists.getFirst()
        .buildingBlocks.getFirst();
    
    // Set the value (contents) of the building block
    buildingBlock.value = "DISCLAIMER\n\nThis document is confidential and intended solely for the use of the individual or entity to whom it is addressed. Any unauthorized review, use, disclosure, or distribution is prohibited.";
    
    await context.sync();
    
    console.log("Building block contents updated successfully");
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the building block.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete a building block named "Company Header" from the template's building blocks collection

```typescript
await Word.run(async (context) => {
    // Get the building blocks collection from the template
    const buildingBlocks = context.document.buildingBlocksCollection;
    buildingBlocks.load("items");
    
    await context.sync();
    
    // Find the building block named "Company Header"
    const companyHeaderBlock = buildingBlocks.items.find(
        block => block.name === "Company Header"
    );
    
    if (companyHeaderBlock) {
        // Delete the building block
        companyHeaderBlock.delete();
        
        await context.sync();
        console.log("Building block 'Company Header' has been deleted.");
    } else {
        console.log("Building block 'Company Header' not found.");
    }
});
```

---

### insert

**Kind:** `create`

Inserts the value of the building block into the document and returns a Range object that represents the contents of the building block within the document.

#### Signature

**Parameters:**
- `range`: `Word.Range` (required)
  The range where the building block should be inserted.
- `richText`: `boolean` (required)
  Indicates whether to insert as rich text.

**Returns:** `Word.Range`

#### Examples

**Example**: Insert a pre-built "Disclaimer" building block from the template into the document at the current selection

```typescript
await Word.run(async (context) => {
    // Get the building block by name from the template
    const buildingBlockCollection = context.application.getBuiltInBuildingBlocksByCategory("General");
    buildingBlockCollection.load("items");
    await context.sync();
    
    const disclaimerBlock = buildingBlockCollection.items.find(bb => bb.name === "Disclaimer");
    
    if (disclaimerBlock) {
        // Get the current selection range
        const range = context.document.getSelection();
        
        // Insert the building block content at the selection
        const insertedRange = disclaimerBlock.insert(range, Word.InsertLocation.replace);
        insertedRange.font.highlightColor = "yellow";
        
        await context.sync();
        console.log("Building block inserted successfully");
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
  - `options`: `Word.Interfaces.BuildingBlockLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.BuildingBlock`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BuildingBlock`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BuildingBlock`

#### Examples

**Example**: Load and display the name and type properties of a building block from the gallery

```typescript
await Word.run(async (context) => {
    // Get the first building block from the first gallery
    const gallery = context.document.buildingBlockGalleries.getFirst();
    const buildingBlock = gallery.buildingBlocks.getFirst();
    
    // Load specific properties of the building block
    buildingBlock.load("name, type");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Building Block Name: " + buildingBlock.name);
    console.log("Building Block Type: " + buildingBlock.type);
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.BuildingBlockUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.BuildingBlock` (required)

  **Returns:** `void`

#### Examples

**Example**: Update a building block's name and description properties simultaneously

```typescript
await Word.run(async (context) => {
    // Get a building block from the template
    const template = context.document.getTemplate();
    const buildingBlock = template.buildingBlockGalleries
        .getByName("Quick Parts")
        .getBuildingBlock("MyBuildingBlock");
    
    // Set multiple properties at once
    buildingBlock.set({
        name: "Updated Building Block",
        description: "This is an updated description for the building block"
    });
    
    await context.sync();
    console.log("Building block properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlock object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BuildingBlockData`

#### Examples

**Example**: Serialize a building block's properties to a plain JavaScript object and log it to the console for debugging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first building block from the first template
    const firstTemplate = context.document.getBuiltInBuildingBlockTemplates().getFirst();
    const buildingBlocks = firstTemplate.buildingBlockEntries;
    const firstBuildingBlock = buildingBlocks.getFirst();
    
    // Load properties we want to serialize
    firstBuildingBlock.load("name,type,category,description");
    
    await context.sync();
    
    // Convert the BuildingBlock API object to a plain JavaScript object
    const buildingBlockData = firstBuildingBlock.toJSON();
    
    // Now we can use standard JavaScript operations on the plain object
    console.log("Building Block Data:", JSON.stringify(buildingBlockData, null, 2));
    console.log("Name:", buildingBlockData.name);
    console.log("Type:", buildingBlockData.type);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BuildingBlock`

#### Examples

**Example**: Track a building block object to safely reuse it across multiple sync calls when inserting it into different locations in the document

```typescript
await Word.run(async (context) => {
    // Get a building block from the template
    const buildingBlocks = context.document.getBuiltInBuildingBlocks();
    const firstBlock = buildingBlocks.getFirst();
    
    // Track the building block to use it across multiple sync calls
    firstBlock.track();
    
    await context.sync();
    
    // Now we can safely use the building block in multiple operations
    const range1 = context.document.body.getRange("Start");
    firstBlock.insertContent(range1, Word.InsertLocation.after);
    
    await context.sync();
    
    // Use the same building block again after another sync
    const range2 = context.document.body.getRange("End");
    firstBlock.insertContent(range2, Word.InsertLocation.before);
    
    await context.sync();
    
    // Clean up - untrack when done
    firstBlock.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BuildingBlock`

#### Examples

**Example**: Get a building block from a gallery, insert it into the document, and then untrack it to free up memory after use.

```typescript
await Word.run(async (context) => {
    // Get a building block from the template
    const template = context.document.getDefaultTemplate();
    const buildingBlockGalleries = template.buildingBlockGalleries;
    const gallery = buildingBlockGalleries.getByName("Quick Parts");
    const buildingBlocks = gallery.buildingBlocks;
    buildingBlocks.load("items");
    
    await context.sync();
    
    if (buildingBlocks.items.length > 0) {
        const buildingBlock = buildingBlocks.items[0];
        
        // Track the building block to use it
        context.trackedObjects.add(buildingBlock);
        buildingBlock.load("name");
        
        await context.sync();
        
        // Insert the building block content
        const range = context.document.body.insertParagraph("", Word.InsertLocation.end).getRange();
        buildingBlock.insertContent(range, Word.InsertLocation.replace);
        
        await context.sync();
        
        // Untrack the building block to release memory
        buildingBlock.untrack();
        
        await context.sync();
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
