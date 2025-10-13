# Word.DropDownListContentControl class

Package: https://learn.microsoft.com/en-us/javascript/api/word

The data specific to content controls of type DropDownList.

Extends
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks
[API set: WordApi 1.9]

#### Examples
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
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- listItems  
  Gets the collection of list items in the dropdown list content control.

## Methods
- addListItem(displayText, value, index)  
  Adds a new list item to this dropdown list content control and returns a Word.ContentControlListItem object.

- deleteAllListItems()  
  Deletes all list items in this dropdown list content control.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### listItems
Gets the collection of list items in the dropdown list content control.

```typescript
readonly listItems: Word.ContentControlListItemCollection;
```

Property Value
- https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrollistitemcollection

Remarks
[API set: WordApi 1.9]

#### Examples
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

## Method Details

### addListItem(displayText, value, index)
Adds a new list item to this dropdown list content control and returns a Word.ContentControlListItem object.

```typescript
addListItem(displayText: string, value?: string, index?: number): Word.ContentControlListItem;
```

Parameters
- displayText (string)  
  Required. Display text of the list item.
- value (string)  
  Optional. Value of the list item.
- index (number)  
  Optional. Index location of the new item in the list. If an item exists at the position specified, the existing item is pushed down in the list. If omitted, the new item is added to the end of the list.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrollistitem

Remarks
[API set: WordApi 1.9]

#### Examples
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

### deleteAllListItems()
Deletes all list items in this dropdown list content control.

```typescript
deleteAllListItems(): void;
```

Returns
- void

Remarks
[API set: WordApi 1.9]

#### Examples
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

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.DropDownListContentControl;
```

Parameters
- propertyNames (string | string[])  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.dropdownlistcontentcontrol

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.DropDownListContentControl;
```

Parameters
- propertyNamesAndPaths ({ select?: string; expand?: string; })  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.dropdownlistcontentcontrol

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DropDownListContentControl object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DropDownListContentControlData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.DropDownListContentControlData;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.dropdownlistcontentcontroldata

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.DropDownListContentControl;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.dropdownlistcontentcontrol

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.DropDownListContentControl;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.dropdownlistcontentcontrol