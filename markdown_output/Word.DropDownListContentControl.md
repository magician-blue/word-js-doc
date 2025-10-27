# Word.DropDownListContentControl

**Package:** `word`

**API Set:** WordApi 1.9

**Extends:** `OfficeExtension.ClientObject`

## Description

The data specific to content controls of type DropDownList.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-dropdown-list-content-control.yaml

// Places a dropdown list content control at the end of the selection.
await Word.run(async (context) => {
  let selection = context.document.getSelection();
  selection.getRange(Word.RangeLocation.end).insertContentControl(Word.ContentControlType.dropDownList);
  await context.sync();

  console.log("Dropdown list content control inserted at the end of the selection.");
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a dropdown list content control to verify the connection to the Office host application before performing operations.

```typescript
await Word.run(async (context) => {
    // Get the first dropdown list content control in the document
    const dropDownControls = context.document.contentControls.getByTypes([Word.ContentControlType.dropDownList]);
    const dropDownControl = dropDownControls.getFirstOrNullObject();
    
    dropDownControl.load("id");
    await context.sync();
    
    if (!dropDownControl.isNullObject) {
        // Access the dropdown list specific data
        const dropDownData = dropDownControl.dropDownListContentControl;
        
        // Access the request context associated with the dropdown control
        const requestContext = dropDownData.context;
        
        // Use the context to perform operations
        console.log("Request context is connected:", requestContext !== null);
        
        // The context can be used to sync operations
        await requestContext.sync();
    }
});
```

---

### listItems

**Type:** `Word.ContentControlListItemCollection`

**Since:** WordApi 1.9

Gets the collection of list items in the dropdown list content control.

#### Examples

**Example**: Delete a specified list item from the first dropdown list content control in the current selection or its parent content control.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-dropdown-list-content-control.yaml

// Deletes the provided list item from the first dropdown list content control in the selection.
await Word.run(async (context) => {
  const listItemText = (document.getElementById("item-to-delete") as HTMLInputElement).value.trim();
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.dropDownList]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,dropDownListContentControl");
  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,dropDownListContentControl");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
      console.warn("No dropdown list content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  let selectedDropdownList: Word.DropDownListContentControl = selectedContentControl.dropDownListContentControl;
  selectedDropdownList.listItems.load("items/*");
  await context.sync();

  let listItems: Word.ContentControlListItemCollection = selectedContentControl.dropDownListContentControl.listItems;
  let itemToDelete: Word.ContentControlListItem = listItems.items.find((item) => item.displayText === listItemText);
  if (!itemToDelete) {
    console.warn(`List item doesn't exist in control with ID ${selectedContentControl.id}: ${listItemText}`)
    return;
  }
  
  itemToDelete.delete();
  await context.sync();

  console.log(`List item deleted from control with ID ${selectedContentControl.id}: ${listItemText}`);
});
```

---

## Methods

### addListItem

**Kind:** `create`

Adds a new list item to this dropdown list content control and returns a Word.ContentControlListItem object.

#### Signature

**Parameters:**
- `displayText`: `string` (required)
  Display text of the list item.
- `value`: `string` (optional)
  Value of the list item.
- `index`: `number` (optional)
  Index location of the new item in the list. If an item exists at the position specified, the existing item is pushed down in the list. If omitted, the new item is added to the end of the list.

**Returns:** `Word.ContentControlListItem`

#### Examples

**Example**: Add a new list item with user-provided text to the first dropdown list content control found in the current selection or its parent content control.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-dropdown-list-content-control.yaml

// Adds the provided list item to the first dropdown list content control in the selection.
await Word.run(async (context) => {
  const listItemText = (document.getElementById("item-to-add") as HTMLInputElement).value.trim();
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.dropDownList]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,dropDownListContentControl");
  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,dropDownListContentControl");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
      console.warn("No dropdown list content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  selectedContentControl.dropDownListContentControl.addListItem(listItemText);
  await context.sync();

  console.log(`List item added to control with ID ${selectedContentControl.id}: ${listItemText}`);
});
```

---

### deleteAllListItems

**Kind:** `delete`

Deletes all list items in this dropdown list content control.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Remove all list items from the first dropdown list content control found in the current selection or its parent content control.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-dropdown-list-content-control.yaml

// Deletes the list items from first dropdown list content control found in the selection.
await Word.run(async (context) => {
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.dropDownList]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,dropDownListContentControl");
  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,dropDownListContentControl");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
      console.warn("No dropdown list content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  console.log(
    `About to delete the list from the dropdown list content control with ID ${selectedContentControl.id}`
  );
  selectedContentControl.dropDownListContentControl.deleteAllListItems();
  await context.sync();

  console.log("Deleted the list from the dropdown list content control.");
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

  **Returns:** `Word.DropDownListContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.DropDownListContentControl`

#### Examples

**Example**: Load and display the title property of the first dropdown list content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first dropdown list content control
    const dropDownControls = context.document.contentControls.getByTypes([Word.ContentControlType.dropDownList]);
    const dropDownControl = dropDownControls.getFirstOrNullObject();
    
    // Load the title property
    dropDownControl.load("title");
    
    await context.sync();
    
    if (!dropDownControl.isNullObject) {
        console.log("Dropdown title: " + dropDownControl.title);
    } else {
        console.log("No dropdown list content control found");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DropDownListContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DropDownListContentControlData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.DropDownListContentControlData`

#### Examples

**Example**: Serialize a dropdown list content control to JSON format to log or store its properties and list items.

```typescript
await Word.run(async (context) => {
    // Get the first dropdown list content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.dropDownList]);
    const dropdownControl = contentControls.getFirst();
    
    // Load the dropdown list content control properties
    dropdownControl.load("dropDownListContentControl");
    await context.sync();
    
    // Get the dropdown list data and convert to JSON
    const dropdownData = dropdownControl.dropDownListContentControl;
    dropdownData.load("listItems");
    await context.sync();
    
    // Convert to plain JavaScript object
    const jsonObject = dropdownData.toJSON();
    
    // Log the JSON representation
    console.log("Dropdown List Content Control JSON:", JSON.stringify(jsonObject, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.DropDownListContentControl`

#### Examples

**Example**: Track a dropdown list content control across multiple sync calls to safely modify its properties without getting an "InvalidObjectPath" error

```typescript
await Word.run(async (context) => {
    // Get the first dropdown list content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.dropDownList]);
    context.load(contentControls);
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const dropDownControl = contentControls.items[0];
        const dropDownList = dropDownControl.dropDownListContentControl;
        
        // Track the object to use it across multiple sync calls
        dropDownList.track();
        
        await context.sync();
        
        // Now we can safely modify properties across multiple syncs
        dropDownList.addListItem("New Option 1", "value1");
        await context.sync();
        
        dropDownList.addListItem("New Option 2", "value2");
        await context.sync();
        
        // Untrack when done
        dropDownList.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.DropDownListContentControl`

#### Examples

**Example**: Release memory for a tracked dropdown list content control after modifying its properties to prevent memory leaks in the host application.

```typescript
await Word.run(async (context) => {
    // Get the first dropdown list content control in the document
    const dropdownControls = context.document.contentControls.getByTypes([Word.ContentControlType.dropDownList]);
    const dropdownControl = dropdownControls.getFirstOrNullObject();
    
    // Load and track the dropdown control
    dropdownControl.load("title");
    dropdownControl.track();
    
    await context.sync();
    
    if (!dropdownControl.isNullObject) {
        // Modify the dropdown control
        dropdownControl.title = "Updated Dropdown";
        
        await context.sync();
        
        // Untrack the object to release memory after we're done using it
        dropdownControl.untrack();
        
        await context.sync();
        
        console.log("Dropdown control updated and memory released");
    }
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.dropdownlistcontentcontrol
