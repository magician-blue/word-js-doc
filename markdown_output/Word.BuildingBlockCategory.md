# Word.BuildingBlockCategory

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a category of building blocks in a Word document.

## Properties

### buildingBlocks

**Type:** `None`

Returns a BuildingBlockCollection object that represents the building blocks for the category.

#### Examples

**Example**: Get all building blocks from the "General" category and insert the first one into the document at the current selection.

```typescript
await Word.run(async (context) => {
    // Get the building block categories
    const categories = context.application.buildingBlockCategories;
    
    // Load the "General" category
    const generalCategory = categories.getByName("General");
    
    // Get the building blocks collection from the category
    const buildingBlocks = generalCategory.buildingBlocks;
    buildingBlocks.load("items");
    
    await context.sync();
    
    // Insert the first building block at the current selection
    if (buildingBlocks.items.length > 0) {
        const firstBlock = buildingBlocks.items[0];
        const selection = context.document.getSelection();
        firstBlock.insertContent(selection, Word.InsertLocation.replace);
    }
    
    await context.sync();
});
```

---

### context

**Type:** `None`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a BuildingBlockCategory object to verify the connection between the add-in and Word application before performing operations.

```typescript
await Word.run(async (context) => {
    // Get a building block category
    const gallery = context.document.buildingBlockGalleries.getByName("Quick Parts");
    const category = gallery.categories.getFirst();
    
    // Load the category
    category.load("name");
    await context.sync();
    
    // Access the context property to verify the connection
    const categoryContext = category.context;
    
    // Use the context to perform additional operations
    console.log("Category context is connected:", categoryContext !== null);
    console.log("Category name:", category.name);
    
    await context.sync();
});
```

---

### index

**Type:** `None`

Returns the position of the BuildingBlockCategory object in a collection.

#### Examples

**Example**: Get the position of the first building block category in the collection and display it in the console.

```typescript
await Word.run(async (context) => {
    const categories = context.document.buildingBlockCategories;
    const firstCategory = categories.getFirst();
    firstCategory.load("index, name");
    
    await context.sync();
    
    console.log(`Category "${firstCategory.name}" is at index: ${firstCategory.index}`);
});
```

---

### name

**Type:** `None`

Returns the name of the BuildingBlockCategory object.

#### Examples

**Example**: Get and display the name of the first building block category in the document's gallery.

```typescript
await Word.run(async (context) => {
    const categories = context.document.buildingBlockCategories;
    categories.load("items");
    
    await context.sync();
    
    if (categories.items.length > 0) {
        const firstCategory = categories.items[0];
        firstCategory.load("name");
        
        await context.sync();
        
        console.log("Category name: " + firstCategory.name);
    }
});
```

---

### type

**Type:** `None`

Returns a BuildingBlockTypeItem object that represents the type of building block for the building block category.

#### Examples

**Example**: Get the building block type (such as AutoText, Quick Parts, etc.) from a building block category and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the first building block category from the template
    const categories = context.document.buildingBlockCategories;
    categories.load("items");
    await context.sync();
    
    if (categories.items.length > 0) {
        const category = categories.items[0];
        
        // Access the type property to get the BuildingBlockTypeItem
        const buildingBlockType = category.type;
        buildingBlockType.load("name");
        await context.sync();
        
        console.log(`Building block type: ${buildingBlockType.name}`);
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
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `None` (required)

  **Returns:** `None`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Load and display the name and type properties of a building block category

```typescript
await Word.run(async (context) => {
    // Get the first building block category from the first gallery
    const gallery = context.document.buildingBlockGalleries.getFirst();
    const category = gallery.categories.getFirst();
    
    // Load the name and type properties
    category.load("name, type");
    
    await context.sync();
    
    console.log(`Category Name: ${category.name}`);
    console.log(`Category Type: ${category.type}`);
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockCategory object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCategoryData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.BuildingBlockCategoryData`
a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCategoryData) that contains shallow copies of any loaded child properties from the original object.

#### Examples

**Example**: Serialize a building block category to JSON format for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first building block category from the gallery
    const gallery = context.document.buildingBlockGalleries.getFirst();
    const category = gallery.categories.getFirst();
    
    // Load properties we want to serialize
    category.load("name, type");
    
    await context.sync();
    
    // Convert the category to a plain JavaScript object
    const categoryData = category.toJSON();
    
    // Now you can use the plain object for logging, storage, or transmission
    console.log("Category as JSON:", JSON.stringify(categoryData, null, 2));
    console.log("Category name:", categoryData.name);
    console.log("Category type:", categoryData.type);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part o

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a building block category object to use it across multiple sync calls when searching for and displaying category information

```typescript
await Word.run(async (context) => {
    // Get a building block category
    const gallery = context.document.buildingBlockGalleries.getByName("Quick Parts");
    const category = gallery.categories.getFirst();
    
    // Track the category object for use across sync calls
    category.track();
    
    // Load properties
    category.load("name");
    
    // First sync
    await context.sync();
    
    console.log("Category name: " + category.name);
    
    // Can safely use the tracked object after another sync
    await context.sync();
    
    console.log("Still accessible: " + category.name);
    
    // Untrack when done to free memory
    category.untrack();
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
