# ContentControlListItem

**Package:** `word`

**API Set:** WordApi 1.9 None

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a list item in a dropdown list or combo box content control.

## Class Examples

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

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a content control list item to verify the connection to the Office host application before performing operations on the list item.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    await context.sync();
    
    // Get the first list item from the content control
    const listItem = contentControl.dropdownListOrComboBox.listItems.getFirst();
    listItem.load("displayText");
    
    // Access the request context associated with the list item
    const itemContext = listItem.context;
    
    // Use the context to sync and verify the connection
    await itemContext.sync();
    
    console.log("List item display text: " + listItem.displayText);
    console.log("Request context is connected to Office host application");
});
```

---

### displayText

**Type:** `string`

**Since:** WordApi 1.9

Specifies the display text of a list item for a dropdown list or combo box content control.

#### Examples

**Example**: Delete a list item with matching display text from the first dropdown list content control in the current selection.

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

### index

**Type:** `number`

**Since:** WordApi 1.9

Specifies the index location of a content control list item in the collection of list items.

#### Examples

**Example**: Get the index of the currently selected item in a dropdown content control and display it to the user.

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type,listItems");
    
    await context.sync();
    
    if (contentControl.type === Word.ContentControlType.dropDownList || 
        contentControl.type === Word.ContentControlType.comboBox) {
        
        const listItems = contentControl.listItems;
        listItems.load("items");
        
        await context.sync();
        
        // Get the index of the first list item
        const firstItem = listItems.items[0];
        firstItem.load("index");
        
        await context.sync();
        
        console.log(`The index of the first list item is: ${firstItem.index}`);
    }
});
```

---

### value

**Type:** `string`

**Since:** WordApi 1.9

Specifies the programmatic value of a list item for a dropdown list or combo box content control.

#### Examples

**Example**: Set the programmatic value of a dropdown list item to "DEPT001" for the first item in a content control's list

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    await context.sync();
    
    // Ensure it's a dropdown list or combo box
    if (contentControl.type === Word.ContentControlType.dropDownList || 
        contentControl.type === Word.ContentControlType.comboBox) {
        
        // Get the first list item
        const listItem = contentControl.listItems.getFirst();
        
        // Set the programmatic value
        listItem.value = "DEPT001";
        
        await context.sync();
    }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the list item.

#### Signature

**Returns:** `void`

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

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.ContentControlListItemLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ContentControlListItem`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ContentControlListItem`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `object` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ContentControlListItem`

#### Examples

**Example**: Load and display the display name and value of the first list item in the first dropdown content control

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the first list item from the content control
    const listItem = contentControl.listItems.getFirst();
    
    // Load the properties of the list item
    listItem.load("displayName, value");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the loaded properties
    console.log(`Display Name: ${listItem.displayName}`);
    console.log(`Value: ${listItem.value}`);
});
```

---

### select

Selects the list item and sets the text of the content control to the value of the list item.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Select the first item in a dropdown content control to set the control's displayed value

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type,dropdownListItems");
    
    await context.sync();
    
    // Verify it's a dropdown or combo box control
    if (contentControl.type === Word.ContentControlType.dropDownList || 
        contentControl.type === Word.ContentControlType.comboBox) {
        
        // Get the first list item and select it
        const firstItem = contentControl.dropdownListItems.getFirst();
        firstItem.select();
        
        await context.sync();
        console.log("First list item selected");
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
  - `properties`: `Interfaces.ContentControlListItemUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ContentControlListItem` (required)

  **Returns:** `void`

#### Examples

**Example**: Update multiple properties of a content control list item at once, setting both its display text and value

```typescript
await Word.run(async (context) => {
    // Get the first content control (assumed to be a dropdown or combo box)
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    await context.sync();
    
    // Get the first list item from the content control
    const listItem = contentControl.listItems.getFirst();
    
    // Set multiple properties at once using the set() method
    listItem.set({
        displayText: "High Priority",
        value: "priority-high"
    });
    
    await context.sync();
    console.log("List item properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ContentControlListItem object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ContentControlListItemData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ContentControlListItemData`

#### Examples

**Example**: Retrieve a content control's dropdown list items and serialize the first item to JSON for logging or data export purposes.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    // Load the list items
    const listItems = contentControl.dropdownListOrComboBoxListItems;
    listItems.load("items");
    
    await context.sync();
    
    // Check if it's a dropdown or combo box
    if (contentControl.type === "DropDownList" || contentControl.type === "ComboBox") {
        const firstItem = listItems.items[0];
        firstItem.load("displayText,value");
        
        await context.sync();
        
        // Convert the list item to a plain JSON object
        const jsonData = firstItem.toJSON();
        
        // Log the JSON representation
        console.log("List Item JSON:", JSON.stringify(jsonData, null, 2));
        // Output example: { "displayText": "Option 1", "value": "1" }
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ContentControlListItem`

#### Examples

**Example**: Track a content control list item across multiple sync calls to safely modify its properties without getting an "InvalidObjectPath" error.

```typescript
await Word.run(async (context) => {
    // Get the first content control (assuming it's a dropdown or combo box)
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    await context.sync();
    
    // Get the first list item from the content control
    const listItem = contentControl.listItems.getFirst();
    listItem.load("displayText,value");
    await context.sync();
    
    // Track the list item for use across multiple sync calls
    listItem.track();
    
    console.log(`Original: ${listItem.displayText} - ${listItem.value}`);
    
    // Perform additional operations after sync
    await context.sync();
    
    // Can safely access the list item properties after tracking
    console.log(`Still accessible: ${listItem.displayText}`);
    
    // Untrack when done
    listItem.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ContentControlListItem`

#### Examples

**Example**: Get a list item from a dropdown content control, use it to read properties, then untrack it to free memory

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    await context.sync();
    
    // Get the first list item from the dropdown
    const listItems = contentControl.listItems;
    const firstItem = listItems.getFirst();
    firstItem.load("displayText,value");
    
    await context.sync();
    
    // Use the list item
    console.log(`Item: ${firstItem.displayText}, Value: ${firstItem.value}`);
    
    // Untrack the list item to release memory
    firstItem.untrack();
    
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word/word.contentcontrollistitem
