# Word.ListLevelCollection

**Package:** `word`

**API Set:** WordApiDesktop 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.ListLevel](/en-us/javascript/api/word/word.listlevel) objects.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/manage-list-styles.yaml

// Gets the properties of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to get properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load("type");
  await context.sync();

  if (style.isNullObject || style.type != Word.StyleType.list) {
    console.warn(`There's no existing style with the name '${styleName}'. Or this isn't a list style.`);
  } else {
    // Load objects to log properties and their values in the console.
    style.load();
    style.listTemplate.load();
    await context.sync();

    console.log(`Properties of the '${styleName}' style:`, style);

    const listLevels = style.listTemplate.listLevels;
    listLevels.load("items");
    await context.sync();

    console.log(`List levels of the '${styleName}' style:`, listLevels);
  }
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ListLevelCollection to verify the connection to the Word host application and log its diagnostic information.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        const listLevels = firstList.levelTypes;
        
        // Access the request context from the ListLevelCollection
        const requestContext = listLevels.context;
        
        // Verify the context is valid and connected
        console.log("Request context is connected:", requestContext !== null);
        console.log("Context debug info:", requestContext.debugInfo);
        
        await context.sync();
    }
});
```

---

### items

**Type:** `Word.ListLevel[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Iterate through all list levels in a list and display the level number and indentation for each level.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        const listLevels = firstList.levelTypes;
        listLevels.load("items");
        await context.sync();
        
        // Access the items property to get all loaded list levels
        const levels = listLevels.items;
        
        console.log(`Found ${levels.length} list levels:`);
        levels.forEach((level, index) => {
            level.load("alignment");
            console.log(`Level ${index}: ${level.alignment}`);
        });
        
        await context.sync();
    }
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first list level in this collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.ListLevel`

#### Examples

**Example**: Get the first list level from a numbered list and change its alignment to right

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Get the list level collection
        const listLevels = firstList.levelTypes[0].listLevels;
        
        // Get the first list level
        const firstLevel = listLevels.getFirst();
        firstLevel.load("alignment");
        await context.sync();
        
        // Change the alignment to right
        firstLevel.alignment = Word.Alignment.right;
        await context.sync();
        
        console.log("First list level alignment changed to right");
    }
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first list level in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.ListLevel`

#### Examples

**Example**: Check if a list has any levels defined and display the alignment of the first level if it exists.

```typescript
await Word.run(async (context) => {
    const list = context.document.body.lists.getFirst();
    const firstLevel = list.listLevels.getFirstOrNullObject();
    firstLevel.load("isNullObject, alignment");
    
    await context.sync();
    
    if (firstLevel.isNullObject) {
        console.log("This list has no levels defined.");
    } else {
        console.log("First level alignment: " + firstLevel.alignment);
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
  - `options`: `Word.Interfaces.ListLevelCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ListLevelCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ListLevelCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ListLevelCollection`

#### Examples

**Example**: Load and display the alignment property of all list levels in the first list of the document

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Get the list level collection
    const listLevels = list.levelTypes.getFirst().listLevels;
    
    // Load the alignment property for all list levels
    listLevels.load("items/alignment");
    
    await context.sync();
    
    // Display the alignment of each list level
    listLevels.items.forEach((level, index) => {
        console.log(`Level ${index}: ${level.alignment}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListLevelCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListLevelCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.ListLevelCollectionData`

#### Examples

**Example**: Retrieve list level information from the first list in a document and serialize it to JSON for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    
    // Get the list level collection
    const listLevels = list.levelTypes;
    
    // Load the properties we want to serialize
    listLevels.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const listLevelsData = listLevels.toJSON();
    
    // Now you can work with the plain object (e.g., log it, send to server, etc.)
    console.log("List levels data:", JSON.stringify(listLevelsData, null, 2));
    console.log("Number of list levels:", listLevelsData.items.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ListLevelCollection`

#### Examples

**Example**: Track a list level collection object to maintain its reference across multiple sync calls when analyzing list formatting properties

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    if (lists.items.length > 0) {
        const firstList = lists.items[0];
        
        // Get the list level collection
        const listLevels = firstList.listLevels;
        
        // Track the collection to use it across multiple sync calls
        listLevels.track();
        
        // Load properties
        listLevels.load("items");
        await context.sync();
        
        // First sync - log count
        console.log(`Number of list levels: ${listLevels.items.length}`);
        await context.sync();
        
        // Second sync - access the tracked object again
        console.log(`Still accessible: ${listLevels.items.length} levels`);
        
        // Untrack when done
        listLevels.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ListLevelCollection`

#### Examples

**Example**: Access list level collection properties, then untrack the object to free memory after use

```typescript
await Word.run(async (context) => {
    // Get the first list in the document
    const list = context.document.body.lists.getFirst();
    const listLevelCollection = list.levelTypes;
    
    // Track the collection to work with it
    listLevelCollection.track();
    
    // Load properties to use the collection
    listLevelCollection.load("items");
    await context.sync();
    
    // Use the collection (e.g., log the count)
    console.log(`Number of list levels: ${listLevelCollection.items.length}`);
    
    // Untrack the object to release memory
    listLevelCollection.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/manage-list-styles.yaml
