# Word.ContentControlListItem class

- Package: [word](/en-us/javascript/api/word)

Represents a list item in a dropdown list or combo box content control.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- displayText: Specifies the display text of a list item for a dropdown list or combo box content control.
- index: Specifies the index location of a content control list item in the collection of list items.
- value: Specifies the programmatic value of a list item for a dropdown list or combo box content control.

## Methods
- delete(): Deletes the list item.
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- select(): Selects the list item and sets the text of the content control to the value of the list item.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript toJSON() method to provide more useful output for JSON.stringify().
- track(): Track the object for automatic adjustment based on surrounding changes in the document.
- untrack(): Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```TypeScript
context: RequestContext;
```

- Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### displayText
Specifies the display text of a list item for a dropdown list or combo box content control.

```TypeScript
displayText: string;
```

- Property value: string

Remarks
[API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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

### index
Specifies the index location of a content control list item in the collection of list items.

```TypeScript
index: number;
```

- Property value: number

Remarks
[API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### value
Specifies the programmatic value of a list item for a dropdown list or combo box content control.

```TypeScript
value: string;
```

- Property value: string

Remarks
[API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### delete()
Deletes the list item.

```TypeScript
delete(): void;
```

- Returns: void

Remarks
[API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
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

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```TypeScript
load(options?: Word.Interfaces.ContentControlListItemLoadOptions): Word.ContentControlListItem;
```

- Parameters:
  - options: [Word.Interfaces.ContentControlListItemLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrollistitemloadoptions)  
    Provides options for which properties of the object to load.
- Returns: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```TypeScript
load(propertyNames?: string | string[]): Word.ContentControlListItem;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```TypeScript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.ContentControlListItem;
```

- Parameters:
  - propertyNamesAndPaths:  
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)

### select()
Selects the list item and sets the text of the content control to the value of the list item.

```TypeScript
select(): void;
```

- Returns: void

Remarks
[API set: WordApi 1.9](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```TypeScript
set(properties: Interfaces.ContentControlListItemUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

- Parameters:
  - properties: [Word.Interfaces.ContentControlListItemUpdateData](/en-us/javascript/api/word/word.interfaces.contentcontrollistitemupdatedata)  
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
    Provides an option to suppress errors if the properties object tries to set any read-only properties.
- Returns: void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```TypeScript
set(properties: Word.ContentControlListItem): void;
```

- Parameters:
  - properties: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)
- Returns: void

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ContentControlListItem object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ContentControlListItemData) that contains shallow copies of any loaded child properties from the original object.

```TypeScript
toJSON(): Word.Interfaces.ContentControlListItemData;
```

- Returns: [Word.Interfaces.ContentControlListItemData](/en-us/javascript/api/word/word.interfaces.contentcontrollistitemdata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```TypeScript
track(): Word.ContentControlListItem;
```

- Returns: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```TypeScript
untrack(): Word.ContentControlListItem;
```

- Returns: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)