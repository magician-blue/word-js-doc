# Word.ContentControlListItemCollection class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem) objects that represent the items in a dropdown list or combo box content control.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.9 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
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

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- getFirst()  
  Gets the first list item in this collection. Throws an ItemNotFound error if this collection is empty.

- getFirstOrNullObject()  
  Gets the first list item in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ContentControlListItemCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ContentControlListItemCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.ContentControlListItem[];
```

Property Value: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)[]

## Method Details

### getFirst()

Gets the first list item in this collection. Throws an ItemNotFound error if this collection is empty.

```typescript
getFirst(): Word.ContentControlListItem;
```

Returns: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)

Remarks: [ API set: WordApi 1.9 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getFirstOrNullObject()

Gets the first list item in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.ContentControlListItem;
```

Returns: [Word.ContentControlListItem](/en-us/javascript/api/word/word.contentcontrollistitem)

Remarks: [ API set: WordApi 1.9 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ContentControlListItemCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ContentControlListItemCollection;
```

Parameters:
- options: [Word.Interfaces.ContentControlListItemCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrollistitemcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.ContentControlListItemCollection](/en-us/javascript/api/word/word.contentcontrollistitemcollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ContentControlListItemCollection;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.ContentControlListItemCollection](/en-us/javascript/api/word/word.contentcontrollistitemcollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.ContentControlListItemCollection;
```

Parameters:
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.ContentControlListItemCollection](/en-us/javascript/api/word/word.contentcontrollistitemcollection)

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ContentControlListItemCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ContentControlListItemCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.ContentControlListItemCollectionData;
```

Returns: [Word.Interfaces.ContentControlListItemCollectionData](/en-us/javascript/api/word/word.interfaces.contentcontrollistitemcollectiondata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ContentControlListItemCollection;
```

Returns: [Word.ContentControlListItemCollection](/en-us/javascript/api/word/word.contentcontrollistitemcollection)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.ContentControlListItemCollection;
```

Returns: [Word.ContentControlListItemCollection](/en-us/javascript/api/word/word.contentcontrollistitemcollection)