# Word.RangeCollection class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Range](/en-us/javascript/api/word/word.range) objects.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.1 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/search.yaml

// Does a basic text search and highlights matches in the document.
await Word.run(async (context) => {
  const results : Word.RangeCollection = context.document.body.search("extend");
  results.load("length");

  await context.sync();

  // Let's traverse the search results and highlight matches.
  for (let i = 0; i < results.items.length; i++) {
    results.items[i].font.highlightColor = "yellow";
  }

  await context.sync();
});
```

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- getFirst()  
  Gets the first range in this collection. Throws an `ItemNotFound` error if this collection is empty.

- getFirstOrNullObject()  
  Gets the first range in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

- load(options)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

- toJSON()  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.RangeCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Range[];
```

Property value
- [Word.Range](/en-us/javascript/api/word/word.range)[]

## Method Details

### getFirst()

Gets the first range in this collection. Throws an `ItemNotFound` error if this collection is empty.

```typescript
getFirst(): Word.Range;
```

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
- [ API set: WordApi 1.3 ]

### getFirstOrNullObject()

Gets the first range in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Range;
```

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
- [ API set: WordApi 1.3 ]

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.RangeCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.RangeCollection;
```

Parameters
- options  
  [Word.Interfaces.RangeCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.rangecollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.RangeCollection;
```

Parameters
- propertyNames  
  string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.RangeCollection;
```

Parameters
- propertyNamesAndPaths  
  [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to