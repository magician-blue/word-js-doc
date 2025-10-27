# BuildingBlockTypeItem

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a type of building block in a Word document.

## Properties

### categories

**Type:** `Word.BuildingBlockCategoryCollection`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns a BuildingBlockCategoryCollection object that represents the categories for a building block type.

#### Examples

**Example**: Get and display all category names for the "General" building block type.

```typescript
await Word.run(async (context) => {
    // Get the "General" building block type
    const buildingBlockType = context.document.buildingBlockTypes.getByName("General");
    
    // Get the categories collection for this building block type
    const categories = buildingBlockType.categories;
    categories.load("items/name");
    
    await context.sync();
    
    // Display all category names
    console.log("Categories in General building block type:");
    categories.items.forEach(category => {
        console.log(`- ${category.name}`);
    });
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a BuildingBlockTypeItem to verify the connection between the add-in and Word before performing operations on building blocks.

```typescript
await Word.run(async (context) => {
    // Get a building block type item
    const gallery = context.document.buildingBlockGalleries.getByName("Quick Parts");
    const buildingBlockType = gallery.buildingBlockTypes.getFirst();
    
    // Load the building block type
    buildingBlockType.load("name");
    await context.sync();
    
    // Access the request context from the BuildingBlockTypeItem
    const itemContext = buildingBlockType.context;
    
    // Verify the context is valid and connected
    if (itemContext) {
        console.log("BuildingBlockTypeItem is connected to Word application");
        console.log("Building block type name: " + buildingBlockType.name);
    }
    
    await context.sync();
});
```

---

### index

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the position of an item in a collection.

#### Examples

**Example**: Get the position of the first building block type in the collection and display it in the console.

```typescript
await Word.run(async (context) => {
    const buildingBlockTypes = context.document.buildingBlockTypes;
    const firstType = buildingBlockTypes.getFirst();
    firstType.load("index");
    
    await context.sync();
    
    console.log(`The building block type is at position: ${firstType.index}`);
});
```

---

### name

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the localized name of a building block type.

#### Examples

**Example**: Display the localized name of the first building block type in the document to the console.

```typescript
await Word.run(async (context) => {
    const buildingBlockTypes = context.document.buildingBlockTypes;
    buildingBlockTypes.load("items");
    
    await context.sync();
    
    if (buildingBlockTypes.items.length > 0) {
        const firstType = buildingBlockTypes.items[0];
        firstType.load("name");
        
        await context.sync();
        
        console.log("Building block type name: " + firstType.name);
    }
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
  - `options`: `Word.Interfaces.BuildingBlockTypeItemLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.BuildingBlockTypeItem`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.BuildingBlockTypeItem`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{select?: string; expand?: string}` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BuildingBlockTypeItem`

#### Examples

**Example**: Load and display the name property of a building block type from the first gallery in the document.

```typescript
await Word.run(async (context) => {
    // Get the first building block gallery
    const galleries = context.document.buildingBlockGalleries;
    galleries.load("items");
    await context.sync();
    
    // Get the first type from the first gallery
    const firstGallery = galleries.items[0];
    const types = firstGallery.buildingBlockTypes;
    types.load("items");
    await context.sync();
    
    const firstType = types.items[0];
    
    // Load the name property of the building block type
    firstType.load("name");
    await context.sync();
    
    // Now we can read the loaded property
    console.log("Building block type name: " + firstType.name);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockTypeItem object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockTypeItemData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BuildingBlockTypeItemData`

#### Examples

**Example**: Retrieve a building block type item and serialize it to JSON format for logging or data transfer purposes.

```typescript
await Word.run(async (context) => {
    // Get the first building block type from the gallery
    const buildingBlockTypes = context.document.getBuiltInBuildingBlockGalleries()
        .getByName("Cover Pages")
        .buildingBlockTypes;
    
    buildingBlockTypes.load("items");
    await context.sync();
    
    if (buildingBlockTypes.items.length > 0) {
        const firstType = buildingBlockTypes.items[0];
        firstType.load("name, category");
        await context.sync();
        
        // Convert the building block type item to a plain JSON object
        const jsonData = firstType.toJSON();
        
        // Now you can use the plain JavaScript object for logging or serialization
        console.log("Building Block Type as JSON:", JSON.stringify(jsonData, null, 2));
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BuildingBlockTypeItem`

#### Examples

**Example**: Track a building block type item across multiple sync calls to prevent "InvalidObjectPath" errors when accessing its properties after document changes.

```typescript
await Word.run(async (context) => {
    // Get a building block type item
    const buildingBlockTypes = context.document.buildingBlockTypes;
    context.load(buildingBlockTypes);
    await context.sync();
    
    const firstType = buildingBlockTypes.items[0];
    
    // Track the object to use it across multiple sync calls
    firstType.track();
    
    // First sync - load initial properties
    context.load(firstType, "name");
    await context.sync();
    
    console.log("Building block type: " + firstType.name);
    
    // Second sync - can still access the object without errors
    context.load(firstType, "description");
    await context.sync();
    
    console.log("Description: " + firstType.description);
    
    // Untrack when done to free up memory
    firstType.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BuildingBlockTypeItem`

#### Examples

**Example**: Access a building block type, use it to retrieve information, then untrack it to free memory after you're done working with it.

```typescript
await Word.run(async (context) => {
    // Get a building block type from the gallery
    const buildingBlockType = context.document.settings.buildingBlockTypes.getFirst();
    
    // Track the object to work with it
    buildingBlockType.track();
    
    // Load properties to use the building block type
    buildingBlockType.load("name");
    await context.sync();
    
    // Use the building block type (e.g., log its name)
    console.log("Building block type: " + buildingBlockType.name);
    
    // Untrack the object to release memory when done
    buildingBlockType.untrack();
    await context.sync();
    
    console.log("Building block type has been untracked and memory released.");
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblocktypeitem
