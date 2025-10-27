# Word.ComboBoxContentControl

**Package:** `word`

**API Set:** WordApi 1.9 None

**Extends:** `OfficeExtension.ClientObject`

## Description

The data specific to content controls of type 'ComboBox'.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-combo-box-content-control.yaml

// Places a combo box content control at the end of the selection.
await Word.run(async (context) => {
  let selection = context.document.getSelection();
  selection.getRange(Word.RangeLocation.end).insertContentControl(Word.ContentControlType.comboBox);
  await context.sync();

  console.log("Combo box content control inserted at the end of the selection.");
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ComboBoxContentControl to verify the connection to the Office host application and log its properties.

```typescript
await Word.run(async (context) => {
    // Get the first combo box content control in the document
    const comboBoxControls = context.document.contentControls.getByTypes([Word.ContentControlType.comboBox]);
    const comboBox = comboBoxControls.getFirstOrNullObject();
    
    // Load the combo box
    comboBox.load("tag");
    await context.sync();
    
    if (!comboBox.isNullObject) {
        // Access the request context from the combo box content control
        const requestContext = comboBox.context;
        
        // Verify the context is valid and connected
        console.log("Request context is available:", requestContext !== null);
        console.log("Context type:", typeof requestContext);
        
        // The context can be used for synchronization operations
        await requestContext.sync();
        console.log("Successfully synchronized using the combo box's context");
    }
});
```

---

### listItems

**Type:** `Word.ContentControlListItemCollection`

**Since:** WordApi 1.9

Gets the collection of list items in the combo box content control.

#### Examples

**Example**: Delete a specified list item from the first combo box content control in the current selection or its parent content control.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-combo-box-content-control.yaml

// Deletes the provided list item from the first combo box content control in the selection.
await Word.run(async (context) => {
  const listItemText = (document.getElementById("item-to-delete") as HTMLInputElement).value.trim();
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.comboBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,comboBoxContentControl");
  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,comboBoxContentControl");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
      console.warn("No combo box content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  let selectedComboBox: Word.ComboBoxContentControl = selectedContentControl.comboBoxContentControl;
  selectedComboBox.listItems.load("items/*");
  await context.sync();

  let listItems: Word.ContentControlListItemCollection = selectedContentControl.comboBoxContentControl.listItems;
  let itemToDelete: Word.ContentControlListItem = listItems.items.find((item) => item.displayText === listItemText);
  if (!itemToDelete) {
    console.warn(`List item doesn't exist in control with ID ${selectedContentControl.id}: ${listItemText}`);
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

Adds a new list item to this combo box content control and returns a [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem) object.

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

**Example**: Add a new list item to the first combo box content control found in the current selection or its parent content control.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-combo-box-content-control.yaml

// Adds the provided list item to the first combo box content control in the selection.
await Word.run(async (context) => {
  const listItemText = (document.getElementById("item-to-add") as HTMLInputElement).value.trim();
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.comboBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,comboBoxContentControl");
  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,comboBoxContentControl");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
      console.warn("No combo box content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  selectedContentControl.comboBoxContentControl.addListItem(listItemText);
  await context.sync();

  console.log(`List item added to control with ID ${selectedContentControl.id}: ${listItemText}`);
});
```

---

### deleteAllListItems

**Kind:** `delete`

Deletes all list items in this combo box content control.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete all list items from the first combo box content control found in the current selection or its parent content control.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-combo-box-content-control.yaml

// Deletes the list items from first combo box content control found in the selection.
await Word.run(async (context) => {
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.comboBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,comboBoxContentControl");
  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,comboBoxContentControl");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
      console.warn("No combo box content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  console.log(`About to delete the list from the combo box content control with ID ${selectedContentControl.id}`);
  selectedContentControl.comboBoxContentControl.deleteAllListItems();
  await context.sync();

  console.log("Deleted the list from the combo box content control.");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ComboBoxContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ComboBoxContentControl`

#### Examples

**Example**: Load and display the title and list items of the first combo box content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first combo box content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.comboBox]);
    const comboBox = contentControls.getFirstOrNullObject();
    
    // Load the combo box properties
    comboBox.load("title");
    comboBox.comboBoxContentControl.load("listItems");
    
    await context.sync();
    
    if (!comboBox.isNullObject) {
        console.log("Combo Box Title:", comboBox.title);
        console.log("List Items:");
        comboBox.comboBoxContentControl.listItems.items.forEach(item => {
            console.log(`- ${item.displayText}: ${item.value}`);
        });
    } else {
        console.log("No combo box content control found");
    }
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ComboBoxContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ComboBoxContentControlData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ComboBoxContentControlData`

#### Examples

**Example**: Serialize a combo box content control's properties to a plain JavaScript object for logging or data transfer purposes.

```typescript
await Word.run(async (context) => {
    // Get the first combo box content control in the document
    const comboBoxControls = context.document.contentControls.getByTypes([Word.ContentControlType.comboBox]);
    const comboBox = comboBoxControls.getFirstOrNullObject();
    
    // Load the combo box properties
    comboBox.load("title,tag");
    const comboBoxData = comboBox.comboBox;
    comboBoxData.load("listItems");
    
    await context.sync();
    
    // Convert the combo box content control data to a plain JavaScript object
    const jsonObject = comboBoxData.toJSON();
    
    // Now you can use the plain object for logging, storage, or transfer
    console.log("Combo box data:", JSON.stringify(jsonObject, null, 2));
    console.log("Number of list items:", jsonObject.listItems?.length);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ComboBoxContentControl`

#### Examples

**Example**: Track a combo box content control across multiple sync calls to safely modify its properties without getting an "InvalidObjectPath" error

```typescript
await Word.run(async (context) => {
    // Get the first combo box content control in the document
    const contentControls = context.document.contentControls.getByTypes([Word.ContentControlType.comboBox]);
    context.load(contentControls, "items");
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const comboBox = contentControls.items[0].comboBoxContentControl;
        
        // Track the combo box to use it across multiple sync calls
        comboBox.track();
        
        // First sync - load current properties
        context.load(comboBox, "listItems");
        await context.sync();
        
        console.log(`Current items: ${comboBox.listItems.items.length}`);
        
        // Second sync - add a new list item (object remains valid because it's tracked)
        comboBox.addListItem("New Option", "newValue");
        await context.sync();
        
        console.log("Successfully added new item to tracked combo box");
        
        // Clean up - untrack when done
        comboBox.untrack();
    }
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ComboBoxContentControl`

#### Examples

**Example**: Release memory for a combo box content control after reading its properties to prevent memory leaks in long-running operations

```typescript
await Word.run(async (context) => {
    // Get the first combo box content control in the document
    const comboBoxControls = context.document.contentControls.getByTypes([Word.ContentControlType.comboBox]);
    const comboBox = comboBoxControls.getFirstOrNullObject().comboBoxContentControl;
    
    // Load properties to work with the combo box
    comboBox.load("listItems");
    await context.sync();
    
    // Check if combo box exists and process it
    if (!comboBox.isNullObject) {
        console.log(`Combo box has ${comboBox.listItems.items.length} items`);
        
        // Untrack the object to release memory after we're done with it
        comboBox.untrack();
        await context.sync();
    }
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.comboboxcontentcontrol
