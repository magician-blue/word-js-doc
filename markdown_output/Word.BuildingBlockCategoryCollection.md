# Word.BuildingBlockCategoryCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of [Word.BuildingBlockCategory](https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblockcategory) objects in a Word document.

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a BuildingBlockCategoryCollection to verify the connection between the add-in and Word application before performing operations on building block categories.

```typescript
await Word.run(async (context) => {
    // Get the building block categories collection
    const categories = context.document.buildingBlockCategories;
    
    // Access the request context associated with the collection
    const requestContext = categories.context;
    
    // Verify the context is valid by checking if it's the same as the current context
    if (requestContext === context) {
        console.log("Request context is properly connected to the Word application");
        
        // Now safe to perform operations using this context
        categories.load("items");
        await context.sync();
        
        console.log(`Found ${categories.items.length} building block categories`);
    }
});
```

---

## Methods

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Display the total number of building block categories available in the document

```typescript
await Word.run(async (context) => {
    const buildingBlockCategories = context.application.buildingBlockCategories;
    const count = buildingBlockCategories.getCount();
    
    await context.sync();
    
    console.log(`Total building block categories: ${count.value}`);
});
```

---

### getItemAt

**Kind:** `read`

Returns a BuildingBlockCategory object that represents the specified item in the collection.

#### Signature

**Parameters:**
- `index`: `number` (required)
  The index of the item to retrieve.

**Returns:** `Word.BuildingBlockCategory`

#### Examples

**Example**: Get the second building block category from the collection and log its name to the console.

```typescript
await Word.run(async (context) => {
    // Get the building block categories collection
    const categories = context.document.buildingBlockCategories;
    
    // Get the category at index 1 (second item)
    const category = categories.getItemAt(1);
    
    // Load the name property
    category.load("name");
    
    // Sync to execute the queued commands
    await context.sync();
    
    // Log the category name
    console.log(`Category name: ${category.name}`);
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

  **Returns:** `Word.BuildingBlockCategoryCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BuildingBlockCategoryCollection`

#### Examples

**Example**: Load and display the names of all building block categories available in the document

```typescript
await Word.run(async (context) => {
    // Get the building block categories collection
    const categories = context.document.buildingBlockCategories;
    
    // Load the 'name' property for all categories in the collection
    categories.load("name");
    
    // Synchronize the document state
    await context.sync();
    
    // Display the category names
    console.log("Building Block Categories:");
    for (let i = 0; i < categories.items.length; i++) {
        console.log(`- ${categories.items[i].name}`);
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.BuildingBlockCategoryCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockCategoryCollectionData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `{ [key: string]: string; }`

#### Examples

**Example**: Export building block categories to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the building block categories collection
    const categories = context.application.buildingBlockCategories;
    
    // Load the properties we want to export
    categories.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const categoriesJSON = categories.toJSON();
    
    // Log or store the JSON representation
    console.log("Building Block Categories:", JSON.stringify(categoriesJSON, null, 2));
    
    // You can now use this plain object for storage, transmission, or comparison
    // without maintaining references to the Word API objects
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BuildingBlockCategoryCollection`

#### Examples

**Example**: Track a building block category collection to maintain its reference across multiple sync calls when iterating through categories and their building blocks

```typescript
await Word.run(async (context) => {
    // Get the building block category collection
    const categories = context.document.buildingBlockCategories;
    
    // Track the collection to prevent "InvalidObjectPath" errors
    // when using it across multiple sync calls
    categories.track();
    
    // Load the collection
    categories.load("items");
    await context.sync();
    
    // First sync - get category count
    console.log(`Found ${categories.items.length} categories`);
    await context.sync();
    
    // Second sync - iterate through categories
    for (let i = 0; i < categories.items.length; i++) {
        const category = categories.items[i];
        category.load("name");
        await context.sync();
        console.log(`Category: ${category.name}`);
    }
    
    // Untrack when done to free up memory
    categories.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BuildingBlockCategoryCollection`

#### Examples

**Example**: Load building block categories, process them, then release memory by untracking the collection to optimize performance

```typescript
await Word.run(async (context) => {
    // Get the building block categories collection
    const categories = context.application.buildingBlockCategories;
    
    // Track the collection for memory management
    categories.track();
    
    // Load properties to work with
    categories.load("items");
    await context.sync();
    
    // Process the categories (e.g., log their count)
    console.log(`Found ${categories.items.length} building block categories`);
    
    // Untrack the collection to release memory
    categories.untrack();
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word
