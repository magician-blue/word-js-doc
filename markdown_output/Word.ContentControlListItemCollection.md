# Word.ContentControlListItemCollection

**Package:** `word`

**API Set:** WordApi 1.9

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains a collection of [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem) objects that represent the items in a dropdown list or combo box content control.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-combo-box-content-control.yaml

// Gets the list items from the first combo box content control found in the selection.
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

  let selectedComboBox: Word.ComboBoxContentControl = selectedContentControl.comboBoxContentControl;
  selectedComboBox.listItems.load("items");
  await context.sync();

  const currentItems: Word.ContentControlListItemCollection = selectedComboBox.listItems;
  console.log(`The list from the combo box content control with ID ${selectedContentControl.id}:`, currentItems);
});
```

## Properties

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a ContentControlListItemCollection to verify the connection to the Office host application before performing operations on dropdown list items.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    await context.sync();
    
    // Get the list items collection
    const listItems = contentControl.listItems;
    
    // Access the request context from the collection
    const itemsContext = listItems.context;
    
    // Verify the context is valid by using it to load properties
    listItems.load("items");
    await itemsContext.sync();
    
    console.log(`Found ${listItems.items.length} items in the dropdown list`);
    console.log("Request context is valid and connected to Office host");
});
```

---

### items

**Type:** `Word.ContentControlListItem[]`

Gets the loaded child items in this collection.

#### Examples

**Example**: Display the text values of all items in a dropdown list content control to the console

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    // Get the list items collection
    const listItems = contentControl.listItems;
    listItems.load("items");
    
    await context.sync();
    
    // Access the loaded items array
    const itemsArray = listItems.items;
    
    // Display each item's text
    for (let i = 0; i < itemsArray.length; i++) {
        itemsArray[i].load("displayText");
    }
    
    await context.sync();
    
    itemsArray.forEach((item, index) => {
        console.log(`Item ${index}: ${item.displayText}`);
    });
});
```

---

## Methods

### getFirst

**Kind:** `read`

Gets the first list item in this collection. Throws an ItemNotFound error if this collection is empty.

#### Signature

**Returns:** `Word.ContentControlListItem`

#### Examples

**Example**: Get and display the text of the first item in a dropdown list content control

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document (assuming it's a dropdown or combo box)
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    // Get the list items collection
    const listItems = contentControl.listItems;
    
    // Get the first list item
    const firstItem = listItems.getFirst();
    firstItem.load("displayText,value");
    
    await context.sync();
    
    // Display the first item's properties
    console.log(`First item display text: ${firstItem.displayText}`);
    console.log(`First item value: ${firstItem.value}`);
});
```

---

### getFirstOrNullObject

**Kind:** `read`

Gets the first list item in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.ContentControlListItem`

#### Examples

**Example**: Check if a dropdown content control has any items and display the first item's text, or show a message if the list is empty.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    await context.sync();
    
    // Ensure it's a dropdown or combo box
    if (contentControl.type === Word.ContentControlType.dropDownList || 
        contentControl.type === Word.ContentControlType.comboBox) {
        
        // Get the first item or null if empty
        const firstItem = contentControl.listItems.getFirstOrNullObject();
        firstItem.load("isNullObject,displayText");
        
        await context.sync();
        
        if (firstItem.isNullObject) {
            console.log("The dropdown list is empty.");
        } else {
            console.log("First item: " + firstItem.displayText);
        }
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
  - `options`: `Word.Interfaces.ContentControlListItemCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ContentControlListItemCollection`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ContentControlListItemCollection`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `OfficeExtension.LoadOption` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ContentControlListItemCollection`

#### Examples

**Example**: Load and display the text values of all items in a dropdown list content control

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the list items collection
    const listItems = contentControl.listItems;
    
    // Load the 'value' and 'displayText' properties for all items
    listItems.load("value, displayText");
    
    await context.sync();
    
    // Display the loaded list items
    console.log(`Found ${listItems.items.length} items in the dropdown:`);
    listItems.items.forEach((item, index) => {
        console.log(`Item ${index + 1}: ${item.displayText} (value: ${item.value})`);
    });
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ContentControlListItemCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ContentControlListItemCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

#### Signature

**Returns:** `Word.Interfaces.ContentControlListItemCollectionData`

#### Examples

**Example**: Export dropdown list content control items to JSON format for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();
    
    if (contentControls.items.length > 0) {
        const dropdownControl = contentControls.items[0];
        
        // Get the list items collection
        const listItems = dropdownControl.dropdownListItems;
        listItems.load("displayText, value");
        await context.sync();
        
        // Convert the collection to a plain JavaScript object
        const jsonData = listItems.toJSON();
        
        // Log or store the JSON representation
        console.log("Dropdown items as JSON:", JSON.stringify(jsonData, null, 2));
        
        // Access the items array from the JSON object
        jsonData.items.forEach(item => {
            console.log(`Display: ${item.displayText}, Value: ${item.value}`);
        });
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ContentControlListItemCollection`

#### Examples

**Example**: Track a dropdown content control's list items collection across multiple sync calls to safely access and modify items without getting "InvalidObjectPath" errors.

```typescript
await Word.run(async (context) => {
    // Get the first content control (assumed to be a dropdown)
    const contentControl = context.document.contentControls.getFirst();
    const listItems = contentControl.dropdownListOrComboBoxListItems;
    
    // Track the collection to use it across multiple sync calls
    listItems.track();
    
    // Load the items
    listItems.load("items");
    await context.sync();
    
    // Now we can safely work with the collection after sync
    console.log(`Found ${listItems.items.length} list items`);
    
    // Perform another sync and still access the collection
    await context.sync();
    
    // Add a new item using the tracked collection
    listItems.add("New Option", "newValue");
    
    // Untrack when done to release memory
    listItems.untrack();
    
    await context.sync();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.ContentControlListItemCollection`

#### Examples

**Example**: Get the list items from a dropdown content control, process them, and then untrack the collection to free memory.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    // Get the list items collection
    const listItems = contentControl.listItems;
    listItems.load("items");
    
    await context.sync();
    
    // Process the list items (e.g., log their count)
    console.log(`Found ${listItems.items.length} list items`);
    
    // Untrack the collection to release memory
    listItems.untrack();
    
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word/word.contentcontrollistitemcollection
