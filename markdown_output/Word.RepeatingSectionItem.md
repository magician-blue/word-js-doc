# RepeatingSectionItem

**Package:** `word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a single item in a [Word.RepeatingSectionContentControl](/en-us/javascript/api/word/word.repeatingsectioncontentcontrol).

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a RepeatingSectionItem to load and read its properties

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstItem = repeatingSections.getFirstOrNullObject().getAsRepeatingSectionItemOrNullObject();
    
    // Access the context property to use it for loading properties
    const itemContext = firstItem.context;
    
    // Use the context to load properties
    firstItem.load("id");
    await itemContext.sync();
    
    if (!firstItem.isNullObject) {
        console.log("Repeating section item ID: " + firstItem.id);
    }
});
```

---

### range

**Type:** `Word.Range`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns the range of this repeating section item, excluding the start and end tags.

#### Examples

**Example**: Highlight all text within the first repeating section item by applying a yellow background color to its range.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstItem = repeatingSections.getFirst() as Word.RepeatingSectionContentControl;
    
    // Get the range of the repeating section item (excluding tags)
    const itemRange = firstItem.range;
    
    // Apply yellow highlight to the range
    itemRange.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes this RepeatingSectionItem object.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete the first item from a repeating section content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstRepeatingSection = repeatingSections.getFirst();
    
    // Get the parent repeating section content control
    const parentRepeatingSection = firstRepeatingSection.parentContentControl;
    parentRepeatingSection.load("repeatingSectionItems");
    
    await context.sync();
    
    // Get the first item in the repeating section
    const firstItem = parentRepeatingSection.repeatingSectionItems.getFirst();
    
    // Delete the first item
    firstItem.delete();
    
    await context.sync();
    
    console.log("First repeating section item deleted successfully.");
});
```

---

### insertItemAfter

**Kind:** `create`

Adds a repeating section item after this item and returns the new item.

#### Signature

**Returns:** `Word.RepeatingSectionItem`

#### Examples

**Example**: Add a new item after the first repeating section item in the document

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstItem = repeatingSections.getFirst();
    
    // Insert a new item after the first item
    const newItem = firstItem.insertItemAfter();
    
    // Load the new item to work with it
    newItem.load("id");
    
    await context.sync();
    
    console.log("New repeating section item added with ID: " + newItem.id);
});
```

---

### insertItemBefore

**Kind:** `create`

Adds a repeating section item before this item and returns the new item.

#### Signature

**Returns:** `Word.RepeatingSectionItem`

#### Examples

**Example**: Insert a new repeating section item before the currently selected item in a repeating section

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document (assuming it's a repeating section)
    const repeatingSectionCC = context.document.contentControls.getFirst();
    repeatingSectionCC.load("type");
    
    await context.sync();
    
    // Get the first item in the repeating section
    const firstItem = repeatingSectionCC.repeatingSectionItems.getFirst();
    
    // Insert a new item before the first item
    const newItem = firstItem.insertItemBefore();
    newItem.load("index");
    
    await context.sync();
    
    console.log(`New item inserted at index: ${newItem.index}`);
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.RepeatingSectionItemLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.RepeatingSectionItem`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.RepeatingSectionItem`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.RepeatingSectionItem`

#### Examples

**Example**: Load and display the content of the first repeating section item in the document

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstItem = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    // Load properties of the repeating section item
    firstItem.load("text, tag, title");
    
    await context.sync();
    
    if (!firstItem.isNullObject) {
        console.log("Item text:", firstItem.text);
        console.log("Item tag:", firstItem.tag);
        console.log("Item title:", firstItem.title);
    } else {
        console.log("No repeating section items found");
    }
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.RepeatingSectionItemUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.RepeatingSectionItem` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a repeating section item, including its text content and font formatting

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstItem = repeatingSections.getFirst() as Word.RepeatingSectionContentControl;
    
    // Set multiple properties at once using the set() method
    firstItem.set({
        title: "Updated Item",
        tag: "item-001",
        appearance: Word.ContentControlAppearance.boundingBox,
        color: "#FF0000"
    });
    
    await context.sync();
    console.log("Repeating section item properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.RepeatingSectionItem object is an API object, the toJSON method returns a plain JavaScript object (typed as [Word.Interfaces.RepeatingSectionItemData](/en-us/javascript/api/word/word.interfaces.repeatingsectionitemdata)) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.RepeatingSectionItemData`

#### Examples

**Example**: Serialize a repeating section item to a plain JavaScript object to log or store its properties

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control in the document
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstItem = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionItem;
    
    // Load properties we want to serialize
    firstItem.load("id,tag,title,type");
    
    await context.sync();
    
    if (!firstItem.isNullObject) {
        // Convert the API object to a plain JavaScript object
        const itemData = firstItem.toJSON();
        
        // Now we can use the plain object (e.g., log it, store it, etc.)
        console.log("Repeating Section Item Data:", JSON.stringify(itemData, null, 2));
    } else {
        console.log("No repeating section items found in the document.");
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.RepeatingSectionItem`

#### Examples

**Example**: Track a repeating section item to maintain its reference across multiple sync calls while modifying its content

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    const firstItem = repeatingSections.getFirstOrNullObject() as Word.RepeatingSectionContentControl;
    
    await context.sync();
    
    if (!firstItem.isNullObject) {
        const items = firstItem.repeatingSectionItemCollection;
        items.load("items");
        await context.sync();
        
        if (items.items.length > 0) {
            const item = items.items[0];
            
            // Track the item to use it across multiple sync calls
            item.track();
            
            // First sync - modify content
            item.contentControls.load("items");
            await context.sync();
            
            // Second sync - access the same tracked object
            item.contentControls.items[0].insertText("Updated content", Word.InsertLocation.replace);
            await context.sync();
            
            // Untrack when done
            item.untrack();
        }
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.RepeatingSectionItem`

#### Examples

**Example**: Process repeating section items to collect their text content, then untrack them to free memory after processing is complete.

```typescript
await Word.run(async (context) => {
    // Get the first repeating section content control
    const repeatingSections = context.document.contentControls.getByTypes([Word.ContentControlType.repeatingSectionItem]);
    repeatingSections.load("items");
    await context.sync();

    const textContents: string[] = [];
    
    // Process each repeating section item
    for (let i = 0; i < repeatingSections.items.length; i++) {
        const item = repeatingSections.items[i] as Word.RepeatingSectionItem;
        item.load("text");
        await context.sync();
        
        // Collect the text content
        textContents.push(item.text);
        
        // Untrack the item to release memory
        item.untrack();
    }
    
    await context.sync();
    
    console.log("Processed items:", textContents);
});
```

---

## Source

- /en-us/javascript/api/word/word.repeatingsectionitem
