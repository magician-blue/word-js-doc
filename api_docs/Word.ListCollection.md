# Word.ListCollection class

- Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.List](/en-us/javascript/api/word/word.list) objects.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.3]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Gets information about the first list in the document.
await Word.run(async (context) => {
  const lists: Word.ListCollection = context.document.body.lists;
  lists.load("items");

  await context.sync();

  if (lists.items.length === 0) {
    console.warn("There are no lists in this document.");
    return;
  }
  
  // Get the first list.
  const list: Word.List = lists.getFirst();
  list.load("levelTypes,levelExistences");

  await context.sync();

  const levelTypes  = list.levelTypes;
  console.log("Level types of the first list:");
  for (let i = 0; i < levelTypes.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelTypes[i]}`);
  }

  const levelExistences = list.levelExistences;
  console.log("Level existences of the first list:");
  for (let i = 0; i < levelExistences.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelExistences[i]}`);
  }
});
```

## Properties

- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items: Gets the loaded child items in this collection.

## Methods

- getById(id): Gets a list by its identifier. Throws an `ItemNotFound` error if there isn't a list with the identifier in this collection.
- getByIdOrNullObject(id): Gets a list by its identifier. If there isn't a list with the identifier in this collection, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- getFirst(): Gets the first list in this collection. Throws an `ItemNotFound` error if this collection is empty.
- getFirstOrNullObject(): Gets the first list in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- getItem(id): Gets a list object by its ID.
- load(options): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- toJSON(): Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ListCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.List[];
```

- Property Value: [Word.List](/en-us/javascript/api/word/word.list)[]

## Method Details

### getById(id)

Gets a list by its identifier. Throws an `ItemNotFound` error if there isn't a list with the identifier in this collection.

```typescript
getById(id: number): Word.List;
```

- Parameters:
  - id (number): Required. A list identifier.
- Returns: [Word.List](/en-us/javascript/api/word/word.list)

Remarks

[API set: WordApi 1.3]

### getByIdOrNullObject(id)

Gets a list by its identifier. If there isn't a list with the identifier in this collection, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getByIdOrNullObject(id: number): Word.List;
```

- Parameters:
  - id (number): Required. A list identifier.
- Returns: [Word.List](/en-us/javascript/api/word/word.list)

Remarks

[API set: WordApi 1.3]

### getFirst()

Gets the first list in this collection. Throws an `ItemNotFound` error if this collection is empty.

```typescript
getFirst(): Word.List;
```

- Returns: [Word.List](/en-us/javascript/api/word/word.list)

Remarks

[API set: WordApi 1.3]

### getFirstOrNullObject()

Gets the first list in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.List;
```

- Returns: [Word.List](/en-us/javascript/api/word/word.list)

Remarks

[API set: WordApi 1.3]

### getItem(id)

Gets a list object by its ID.

```typescript
getItem(id: number): Word.List;
```

- Parameters:
  - id (number): The list's ID.
- Returns: [Word.List](/en-us/javascript/api/word/word.list)

Remarks

[API set: WordApi 1.3]

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ListCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.ListCollection;
```

- Parameters:
  - options ([Word.Interfaces.ListCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.listcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)): Provides options for which properties of the object to load.
- Returns: [Word.ListCollection](/en-us/javascript/api/word/word.listcollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ListCollection;
```

- Parameters:
  - propertyNames (string | string[]): A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.ListCollection](/en-us/javascript/api/word/word.listcollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.ListCollection;
```

- Parameters:
  - propertyNamesAndPaths ([OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)): `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.ListCollection](/en-us/javascript/api/word/word.listcollection)

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ListCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ListCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.ListCollectionData;
```

- Returns: [Word.Interfaces.ListCollectionData](/en-us/javascript/api/word/word.interfaces.listcollectiondata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ListCollection;
```

- Returns: [Word.ListCollection](/en-us/javascript/api/word/word.listcollection)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ListCollection;
```

- Returns: [Word.ListCollection](/en-us/javascript/api/word/word.listcollection)