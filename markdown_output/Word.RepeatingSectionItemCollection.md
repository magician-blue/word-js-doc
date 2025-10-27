# Word.RepeatingSectionItemCollection

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a collection of https://learn.microsoft.com/en-us/javascript/api/word/word.repeatingsectionitem objects in a Word document.

## Properties

### context

**Type:** `RequestContext`

**Since:** BETA (PREVIEW ONLY)

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a RepeatingSectionItemCollection to verify the connection between the add-in and Word before performing operations on repeating section items.

```typescript
await Word.run(async (context) => {
    // Get the first content control that is a repeating section
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    const repeatingSectionCC = contentControls.items.find(
        cc => cc.type === Word.ContentControlType.repeatingSectionItem
    );
    
    if (repeatingSectionCC) {
        const repeatingSectionItems = repeatingSectionCC.getRange().parentContentControl.repeatingSectionItemCollection;
        
        // Access the context property to verify the connection
        const requestContext = repeatingSectionItems.context;
        
        // Use the context to perform operations
        console.log("Context connected:", requestContext !== null);
        
        // Load items using the same context
        repeatingSectionItems.load("items");
        await requestContext.sync();
        
        console.log(`Found ${repeatingSectionItems.items.length} repeating section items`);
    }
});
```

---

## Methods

### getItemAt

**Kind:** `read`

Returns an individual repeating section item.

#### Signature

**Parameters:**
- `index`: `number` (required)
  The index of the item to retrieve.

**Returns:** `Word.RepeatingSectionItem`
A RepeatingSectionItem object representing the item at the specified index.

#### Examples

**Example**: Get the second repeating section item from the first content control and highlight its text in yellow.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirstOrNullObject();
    contentControl.load("type");
    
    await context.sync();
    
    // Get the repeating section items collection
    const repeatingSectionItems = contentControl.repeatingSectionItems;
    repeatingSectionItems.load("items");
    
    await context.sync();
    
    // Get the second item (index 1) from the collection
    const secondItem = repeatingSectionItems.getItemAt(1);
    
    // Highlight the text in the second repeating section item
    secondItem.body.font.highlightColor = "yellow";
    
    await context.sync();
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

  **Returns:** `Word.RepeatingSectionItemCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.RepeatingSectionItemCollection`

#### Examples

**Example**: Load and display the IDs of all repeating section items in the document

```typescript
await Word.run(async (context) => {
    // Get all repeating section items in the document
    const repeatingSectionItems = context.document.body.getRepeatingSectionItemCollection();
    
    // Load the 'id' property for all items in the collection
    repeatingSectionItems.load("id");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the IDs of all repeating section items
    console.log(`Found ${repeatingSectionItems.items.length} repeating section items:`);
    repeatingSectionItems.items.forEach((item, index) => {
        console.log(`Item ${index + 1} ID: ${item.id}`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.RepeatingSectionItemCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.RepeatingSectionItemCollectionData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `{ [key: string]: string; }`

#### Examples

**Example**: Export repeating section items data to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get all repeating section items in the document
    const repeatingSectionItems = context.document.contentControls
        .getByTypes([Word.ContentControlType.repeatingSectionItem]);
    
    // Load properties needed for the collection
    repeatingSectionItems.load("items");
    
    await context.sync();
    
    // Convert the collection to a plain JavaScript object
    const jsonData = repeatingSectionItems.toJSON();
    
    // Log or use the JSON data
    console.log("Repeating Section Items Data:", JSON.stringify(jsonData, null, 2));
    
    // The jsonData object now contains a plain JavaScript representation
    // that can be easily serialized, stored, or transmitted
    return jsonData;
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.RepeatingSectionItemCollection`

#### Examples

**Example**: Track a repeating section item collection across multiple sync calls to safely access and modify items without encountering "InvalidObjectPath" errors

```typescript
await Word.run(async (context) => {
    // Get the first repeating section in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstSection = repeatingSections.getFirstOrNullObject();
    
    await context.sync();
    
    if (!firstSection.isNullObject) {
        const parentRepeatingSection = firstSection.parentContentControlOrNullObject;
        await context.sync();
        
        if (!parentRepeatingSection.isNullObject) {
            // Get the repeating section items collection
            const items = parentRepeatingSection.repeatingSectionItemCollection;
            
            // Track the collection for use across multiple sync calls
            items.track();
            
            await context.sync();
            
            // Now safe to use the collection across multiple syncs
            items.load("items");
            await context.sync();
            
            console.log(`Found ${items.items.length} repeating section items`);
            
            // Perform additional operations
            items.items.forEach((item, index) => {
                console.log(`Item ${index + 1} ID: ${item.id}`);
            });
            
            // Untrack when done
            items.untrack();
        }
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.RepeatingSectionItemCollection`

#### Examples

**Example**: Get all repeating section items in a document, perform operations on them, then untrack them to free memory and improve performance.

```typescript
await Word.run(async (context) => {
    // Get all repeating section items in the document
    const repeatingSectionItems = context.document.body.getRepeatingSectionItems();
    
    // Track the collection for use
    repeatingSectionItems.load("items");
    await context.sync();
    
    // Perform some operations with the items
    console.log(`Found ${repeatingSectionItems.items.length} repeating section items`);
    
    // Once done with the collection, untrack it to release memory
    repeatingSectionItems.untrack();
    
    // Sync to apply the memory release
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.repeatingsectionitem
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
