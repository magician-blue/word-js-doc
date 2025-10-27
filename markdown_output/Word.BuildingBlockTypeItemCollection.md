# BuildingBlockTypeItemCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of building block types in a Word document.

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the BuildingBlockTypeItemCollection's request context to verify the connection to the Word host application before performing operations on building block types.

```typescript
await Word.run(async (context) => {
    // Get the building block types collection from the gallery
    const gallery = context.document.properties.customXmlParts.getByNamespace("http://schemas.microsoft.com/office/word/2010/wordml")[0];
    const buildingBlockTypes = context.document.buildingBlockTypes;
    
    // Access the request context associated with the collection
    const requestContext = buildingBlockTypes.context;
    
    // Verify the context is connected to the Word application
    console.log("Context is connected:", requestContext !== null);
    
    // Use the context to load properties
    buildingBlockTypes.load("items");
    await context.sync();
    
    console.log(`Found ${buildingBlockTypes.items.length} building block types`);
});
```

---

## Methods

### getByType

**Kind:** `read`

Gets a BuildingBlockTypeItem object by its type in the collection.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `type`: `Word.BuildingBlockType` (required)
    Specifies the building block type of the item in the collection.

  **Returns:** `Word.BuildingBlockTypeItem`

**Overload 2:**

  **Parameters:**
  - `type`: `"QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography"` (required)
    Specifies the building block type of the item in the collection.

  **Returns:** `Word.BuildingBlockTypeItem`

#### Examples

**Example**: Get the "AutoText" building block type from the document's building block collection and load its name property.

```typescript
await Word.run(async (context) => {
    // Get the building block types collection
    const buildingBlockTypes = context.document.buildingBlockTypes;
    
    // Get the AutoText building block type
    const autoTextType = buildingBlockTypes.getByType(Word.BuildingBlockType.autoText);
    
    // Load the name property
    autoTextType.load("name");
    
    await context.sync();
    
    console.log("Building block type name: " + autoTextType.name);
});
```

---

### getCount

**Kind:** `read`

Returns the number of items in the collection.

#### Signature

**Returns:** `OfficeExtension.ClientResult<number>`

#### Examples

**Example**: Display the total number of building block types available in the document

```typescript
await Word.run(async (context) => {
    const buildingBlockTypes = context.document.buildingBlockTypes;
    const count = buildingBlockTypes.getCount();
    
    await context.sync();
    
    console.log(`Total building block types: ${count.value}`);
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

  **Returns:** `Word.BuildingBlockTypeItemCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.BuildingBlockTypeItemCollection`

#### Examples

**Example**: Load and display the names of all available building block types in the document

```typescript
await Word.run(async (context) => {
    // Get the building block type collection
    const buildingBlockTypes = context.document.buildingBlockTypes;
    
    // Load the 'name' property for all building block types
    buildingBlockTypes.load("items/name");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the building block type names
    console.log("Available building block types:");
    buildingBlockTypes.items.forEach((type) => {
        console.log(`- ${type.name}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BuildingBlockTypeItemCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BuildingBlockTypeItemCollectionData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `{ [key: string]: string; }`

#### Examples

**Example**: Export building block type collection data to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the building block type collection
    const buildingBlockTypes = context.document.buildingBlockTypes;
    
    // Load the properties we want to include in the JSON output
    buildingBlockTypes.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const jsonData = buildingBlockTypes.toJSON();
    
    // Use the JSON data (e.g., log it, send to server, or store locally)
    console.log("Building Block Types as JSON:", JSON.stringify(jsonData, null, 2));
    
    // You can now work with the plain JavaScript object
    console.log(`Total building block types: ${jsonData.items?.length || 0}`);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.BuildingBlockTypeItemCollection`

#### Examples

**Example**: Track a building block type collection to maintain a reference across multiple sync calls when iterating through building block types

```typescript
await Word.run(async (context) => {
    // Get the building block type collection
    const buildingBlockTypes = context.document.buildingBlockTypes;
    
    // Track the collection to use it across multiple sync calls
    buildingBlockTypes.track();
    
    // Load properties
    buildingBlockTypes.load("items");
    await context.sync();
    
    // First sync - process the collection
    console.log(`Found ${buildingBlockTypes.items.length} building block types`);
    
    // Perform additional operations that require another sync
    await context.sync();
    
    // The tracked object remains valid across sync calls
    for (const type of buildingBlockTypes.items) {
        type.load("name");
    }
    await context.sync();
    
    // Untrack when done to release memory
    buildingBlockTypes.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.BuildingBlockTypeItemCollection`

#### Examples

**Example**: Load building block types, use them to get information, then untrack the collection to free memory after processing is complete.

```typescript
await Word.run(async (context) => {
    // Get the building block types collection
    const buildingBlockTypes = context.document.buildingBlockTypes;
    
    // Load properties and track the collection
    buildingBlockTypes.load("items");
    await context.sync();
    
    // Use the collection (e.g., log the count)
    console.log(`Found ${buildingBlockTypes.items.length} building block types`);
    
    // Process the types as needed
    for (const type of buildingBlockTypes.items) {
        type.load("name");
    }
    await context.sync();
    
    // Untrack the collection to release memory
    buildingBlockTypes.untrack();
    await context.sync();
    
    console.log("Building block types collection memory released");
});
```

---

## Source

- /en-us/javascript/api/word
- /en-us/javascript/api/office/officeextension.clientobject
- /en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- /en-us/javascript/api/word/word.requestcontext
- /en-us/javascript/api/word/word.buildingblocktype
- /en-us/javascript/api/word/word.buildingblocktypeitem
- /en-us/javascript/api/office/officeextension.clientresult
- /en-us/javascript/api/word/word.buildingblocktypeitemcollection
- /en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
