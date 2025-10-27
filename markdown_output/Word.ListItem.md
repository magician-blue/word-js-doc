# Word.ListItem

**Package:** `word`

**API Set:** WordApi 1.3

**Extends:** `officeextension.clientobject`

## Description

Represents the paragraph list item format.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/insert-list.yaml

// This example starts a new list with the second paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Start new list using the second paragraph.
  const list: Word.List = paragraphs.items[1].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set up list level for the list item.
  paragraph.listItem.level = 4;

  // To add paragraphs outside the list, use Before or After.
  list.insertParagraph("New paragraph goes after (not part of the list)", "After");

  await context.sync();
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ListItem object to verify the connection between the add-in and Word application before performing list operations.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const listItem = paragraph.listItem;
    
    // Load the list item
    listItem.load("level");
    await context.sync();
    
    // Access the request context from the listItem object
    const listItemContext = listItem.context;
    
    // Verify the context is valid and connected
    if (listItemContext) {
        console.log("ListItem context is connected to Word application");
        console.log("List level:", listItem.level);
    }
});
```

---

### level

**Type:** `number`

**Since:** WordApi 1.3

Specifies the level of the item in the list.

#### Examples

**Example**: Set the list level of a newly inserted list item to level 5 (index 4) at the end of a list.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/insert-list.yaml

// This example starts a new list with the second paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Start new list using the second paragraph.
  const list: Word.List = paragraphs.items[1].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set up list level for the list item.
  paragraph.listItem.level = 4;

  // To add paragraphs outside the list, use Before or After.
  list.insertParagraph("New paragraph goes after (not part of the list)", "After");

  await context.sync();
});
```

---

### listString

**Type:** `string`

**Since:** WordApi 1.3

Gets the list item bullet, number, or picture as a string.

#### Examples

**Example**: Get and display the list item string (bullet, number, or picture) from the first paragraph in the document that is part of a list.

```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    
    await context.sync();
    
    // Find the first paragraph that is a list item
    for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        const listItem = paragraph.listItem;
        
        if (listItem) {
            listItem.load("listString");
            await context.sync();
            
            console.log("List item string: " + listItem.listString);
            break;
        }
    }
});
```

---

### siblingIndex

**Type:** `number`

**Since:** WordApi 1.3

Gets the list item order number in relation to its siblings.

#### Examples

**Example**: Get and display the position of the current paragraph within its sibling list items (e.g., whether it's the 1st, 2nd, or 3rd item at the same level).

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const listItem = paragraph.listItem;
    
    listItem.load("siblingIndex");
    await context.sync();
    
    console.log(`This list item is at position ${listItem.siblingIndex + 1} among its siblings`);
});
```

---

## Methods

### getAncestor

**Kind:** `read`

Gets the list item parent, or the closest ancestor if the parent doesn't exist. Throws an ItemNotFound error if the list item has no ancestor.

#### Signature

**Parameters:**
- `parentOnly`: `boolean` (optional)
  Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Get the parent list item of the currently selected list item and highlight it in yellow

```typescript
await Word.run(async (context) => {
    // Get the selected paragraph
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    paragraph.load("listItem");
    
    await context.sync();
    
    // Get the parent list item (or closest ancestor)
    const parentListItem = paragraph.listItem.getAncestor(true);
    const parentParagraph = parentListItem.getParagraph();
    
    // Highlight the parent paragraph
    parentParagraph.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### getAncestorOrNullObject

**Kind:** `read`

Gets the list item parent, or the closest ancestor if the parent doesn't exist. If the list item has no ancestor, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

#### Signature

**Parameters:**
- `parentOnly`: `boolean` (optional)
  Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Get the parent list item of the currently selected paragraph and display its text, or show a message if no parent exists

```typescript
await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const paragraph = selection.paragraphs.getFirst();
    const listItem = paragraph.listItem;
    
    // Get the parent list item (or null object if none exists)
    const parentListItem = listItem.getAncestorOrNullObject(true);
    
    // Load properties
    parentListItem.load("isNullObject");
    await context.sync();
    
    if (parentListItem.isNullObject) {
        console.log("The selected item has no parent list item.");
    } else {
        // Load the parent's paragraph text
        const parentParagraph = parentListItem.paragraph;
        parentParagraph.load("text");
        await context.sync();
        
        console.log("Parent list item text: " + parentParagraph.text);
    }
});
```

---

### getDescendants

**Kind:** `read`

Gets all descendant list items of the list item.

#### Signature

**Parameters:**
- `directChildrenOnly`: `boolean` (optional)
  Optional. Specifies only the list item's direct children will be returned. The default is false that indicates to get all descendant items.

**Returns:** `Word.ParagraphCollection`

#### Examples

**Example**: Get all nested list items under the first list item in the document and highlight them in yellow

```typescript
await Word.run(async (context) => {
    // Get the first list item in the document
    const listItems = context.document.body.lists.getFirst().listItems;
    const firstListItem = listItems.getFirst();
    
    // Get all descendant list items (nested items at any level)
    const descendants = firstListItem.getDescendants(false);
    
    // Load the paragraphs of the descendants
    descendants.load("items");
    
    await context.sync();
    
    // Highlight all descendant list items
    for (let i = 0; i < descendants.items.length; i++) {
        descendants.items[i].getRange().font.highlightColor = "yellow";
    }
    
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
  - `options`: `Word.Interfaces.ListItemLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ListItem`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ListItem`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ListItem`

#### Examples

**Example**: Get and display the list level of the first paragraph in the document if it's a list item

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const listItem = paragraph.listItem;
    
    // Load the list item properties
    listItem.load("level, listString");
    
    await context.sync();
    
    // Check if paragraph is a list item and display its properties
    if (listItem.isNullObject) {
        console.log("The first paragraph is not a list item");
    } else {
        console.log(`List level: ${listItem.level}`);
        console.log(`List string: ${listItem.listString}`);
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
  - `properties`: `Interfaces.ListItemUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ListItem` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple list item properties at once by setting the level to 1 and list type to bullets for the first paragraph

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const listItem = firstParagraph.listItem;
    
    listItem.set({
        level: 1,
        listString: "•"
    });
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListItem object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListItemData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ListItemData`

#### Examples

**Example**: Get the list item properties as a plain JavaScript object and log it to the console for debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const listItem = paragraph.listItem;
    
    // Load the list item properties
    listItem.load("level,listString,siblingIndex");
    
    await context.sync();
    
    // Convert the ListItem object to a plain JavaScript object
    const listItemData = listItem.toJSON();
    
    // Log the plain object (useful for debugging or data export)
    console.log("List Item Data:", listItemData);
    console.log("Level:", listItemData.level);
    console.log("List String:", listItemData.listString);
    console.log("Sibling Index:", listItemData.siblingIndex);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ListItem`

#### Examples

**Example**: Track a list item object across multiple sync calls to safely modify its properties without getting an "InvalidObjectPath" error

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    const listItem = paragraph.listItem;
    
    // Track the list item to use it across multiple sync calls
    listItem.track();
    
    // Load properties
    listItem.load("level");
    await context.sync();
    
    // Now we can safely use the tracked object after sync
    console.log("Current list level:", listItem.level);
    
    // Modify the list item level
    listItem.level = listItem.level + 1;
    await context.sync();
    
    // Untrack when done to release memory
    listItem.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ListItem`

#### Examples

**Example**: Get list item information from the first paragraph and then untrack the list item object to free memory after use.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const listItem = firstParagraph.listItem;
    
    // Track the list item to work with it
    context.trackedObjects.add(listItem);
    
    // Load properties to use
    listItem.load("level,listString");
    await context.sync();
    
    // Use the list item data
    console.log(`List level: ${listItem.level}`);
    console.log(`List string: ${listItem.listString}`);
    
    // Untrack to release memory after we're done
    listItem.untrack();
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.listitem
