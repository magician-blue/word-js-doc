# Word.SectionCollection class

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Contains the collection of the document's [Word.Section](https://learn.microsoft.com/en-us/javascript/api/word/word.section) objects.

Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.1]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-section-breaks.yaml

// Inserts a section break on the next even page.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.insertBreak(Word.BreakType.sectionEven, Word.InsertLocation.end);

  await context.sync();

  console.log("Inserted section break on next even page.");
});
```

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items  
  Gets the loaded child items in this collection.

## Methods

- `getFirst()`  
  Gets the first section in this collection. Throws an `ItemNotFound` error if this collection is empty.
- `getFirstOrNullObject()`  
  Gets the first section in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- `load(options)`  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- `load(propertyNames)`  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- `load(propertyNamesAndPaths)`  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- `toJSON()`  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- `track()`  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- `untrack()`  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Section[];
```

Property Value: [Word.Section](https://learn.microsoft.com/en-us/javascript/api/word/word.section)[]

## Method Details

### getFirst()

Gets the first section in this collection. Throws an `ItemNotFound` error if this collection is empty.

```typescript
getFirst(): Word.Section;
```

Returns: [Word.Section](https://learn.microsoft.com/en-us/javascript/api/word/word.section)

Remarks: [API set: WordApi 1.3]

### getFirstOrNullObject()

Gets the first section in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [\*OrNullObject methods and properties](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Section;
```

Returns: [Word.Section](https://learn.microsoft.com/en-us/javascript/api/word/word.section)

Remarks: [API set: WordApi 1.3]

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.SectionCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.SectionCollection;
```

Parameters:
- options: [Word.Interfaces.SectionCollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.sectioncollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.SectionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.sectioncollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.SectionCollection;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.SectionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.sectioncollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.SectionCollection;
```

Parameters:
- propertyNamesAndPaths: [OfficeExtension.LoadOption](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.SectionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.sectioncollection)

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SectionCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SectionCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.SectionCollectionData;
```

Returns: [Word.Interfaces.SectionCollectionData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.sectioncollectiondata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.SectionCollection;
```

Returns: [Word.SectionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.sectioncollection)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.SectionCollection;
```

Returns: [Word.SectionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.sectioncollection)