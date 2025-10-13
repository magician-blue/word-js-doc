# Word.ListItem class

Represents the paragraph list item format.

Package: https://learn.microsoft.com/en-us/javascript/api/word

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[ API set: WordApi 1.3 ]

#### Examples

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

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- level  
  Specifies the level of the item in the list.

- listString  
  Gets the list item bullet, number, or picture as a string.

- siblingIndex  
  Gets the list item order number in relation to its siblings.

## Methods

- getAncestor(parentOnly)  
  Gets the list item parent, or the closest ancestor if the parent doesn't exist. Throws an ItemNotFound error if the list item has no ancestor.

- getAncestorOrNullObject(parentOnly)  
  Gets the list item parent, or the closest ancestor if the parent doesn't exist. If the list item has no ancestor, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

- getDescendants(directChildrenOnly)  
  Gets all descendant list items of the list item.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListItem object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListItemData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### level

Specifies the level of the item in the list.

```typescript
level: number;
```

Property value: number

Remarks: [ API set: WordApi 1.3 ]

#### Examples

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

### listString

Gets the list item bullet, number, or picture as a string.

```typescript
readonly listString: string;
```

Property value: string

Remarks: [ API set: WordApi 1.3 ]

### siblingIndex

Gets the list item order number in relation to its siblings.

```typescript
readonly siblingIndex: number;
```

Property value: number

Remarks: [ API set: WordApi 1.3 ]

## Method details

### getAncestor(parentOnly)

Gets the list item parent, or the closest ancestor if the parent doesn't exist. Throws an ItemNotFound error if the list item has no ancestor.

```typescript
getAncestor(parentOnly?: boolean): Word.Paragraph;
```

Parameters:
- parentOnly: boolean  
  Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph

Remarks: [ API set: WordApi 1.3 ]

### getAncestorOrNullObject(parentOnly)

Gets the list item parent, or the closest ancestor if the parent doesn't exist. If the list item has no ancestor, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties.

```typescript
getAncestorOrNullObject(parentOnly?: boolean): Word.Paragraph;
```

Parameters:
- parentOnly: boolean  
  Optional. Specifies only the list item's parent will be returned. The default is false that specifies to get the lowest ancestor.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.paragraph

Remarks: [ API set: WordApi 1.3 ]

### getDescendants(directChildrenOnly)

Gets all descendant list items of the list item.

```typescript
getDescendants(directChildrenOnly?: boolean): Word.ParagraphCollection;
```

Parameters:
- directChildrenOnly: boolean  
  Optional. Specifies only the list item's direct children will be returned. The default is false that indicates to get all descendant items.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.paragraphcollection

Remarks: [ API set: WordApi 1.3 ]

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ListItemLoadOptions): Word.ListItem;
```

Parameters:
- options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.listitemloadoptions  
  Provides options for which properties of the object to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.listitem

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ListItem;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.listitem

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ListItem;
```

Parameters:
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.listitem

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ListItemUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.listitemupdatedata  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ListItem): void;
```

Parameters:
- properties: https://learn.microsoft.com/en-us/javascript/api/word/word.listitem

Returns: void

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ListItem object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ListItemData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ListItemData;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.listitemdata

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ListItem;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.listitem

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.ListItem;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.listitem