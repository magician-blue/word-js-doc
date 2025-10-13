# Word.InlinePictureCollection class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture) objects.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.1 ]

### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Gets the first image in the document.
await Word.run(async (context) => {
  const firstPicture: Word.InlinePicture = context.document.body.inlinePictures.getFirst();
  firstPicture.load("width, height, imageFormat");

  await context.sync();
  console.log(`Image dimensions: ${firstPicture.width} x ${firstPicture.height}`, `Image format: ${firstPicture.imageFormat}`);
  // Get the image encoded as Base64.
  const base64 = firstPicture.getBase64ImageSrc();

  await context.sync();
  console.log(base64.value);
});
```

## Properties
- [context](#word-word-inlinepicturecollection-context-member) — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#word-word-inlinepicturecollection-items-member) — Gets the loaded child items in this collection.

## Methods
- [getFirst()](#word-word-inlinepicturecollection-getfirst-member1) — Gets the first inline image in this collection. Throws an ItemNotFound error if this collection is empty.
- [getFirstOrNullObject()](#word-word-inlinepicturecollection-getfirstornullobject-member1) — Gets the first inline image in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [load(options)](#word-word-inlinepicturecollection-load-member1) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNames)](#word-word-inlinepicturecollection-load-member2) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-inlinepicturecollection-load-member3) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [toJSON()](#word-word-inlinepicturecollection-tojson-member1) — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.InlinePictureCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.InlinePictureCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- [track()](#word-word-inlinepicturecollection-track-member1) — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#word-word-inlinepicturecollection-untrack-member1) — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context
id: word-word-inlinepicturecollection-context-member

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
id: word-word-inlinepicturecollection-items-member

Gets the loaded child items in this collection.

```typescript
readonly items: Word.InlinePicture[];
```

Property Value: [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)[]

## Method Details

### getFirst()
id: word-word-inlinepicturecollection-getfirst-member1

Gets the first inline image in this collection. Throws an ItemNotFound error if this collection is empty.

```typescript
getFirst(): Word.InlinePicture;
```

Returns: [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

Remarks  
[ API set: WordApi 1.3 ]

### getFirstOrNullObject()
id: word-word-inlinepicturecollection-getfirstornullobject-member1

Gets the first inline image in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.InlinePicture;
```

Returns: [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

Remarks  
[ API set: WordApi 1.3 ]

### load(options)
id: word-word-inlinepicturecollection-load-member1

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.InlinePictureCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.InlinePictureCollection;
```

Parameters
- options: [Word.Interfaces.InlinePictureCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.inlinepicturecollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)

### load(propertyNames)
id: word-word-inlinepicturecollection-load-member2

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.InlinePictureCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)

### load(propertyNamesAndPaths)
id: word-word-inlinepicturecollection-load-member3

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.InlinePictureCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)

### toJSON()
id: word-word-inlinepicturecollection-tojson-member1

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.InlinePictureCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.InlinePictureCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.InlinePictureCollectionData;
```

Returns: [Word.Interfaces.InlinePictureCollectionData](/en-us/javascript/api/word/word.interfaces.inlinepicturecollectiondata)

### track()
id: word-word-inlinepicturecollection-track-member1

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.InlinePictureCollection;
```

Returns: [Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)

### untrack()
id: word-word-inlinepicturecollection-untrack-member1

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.InlinePictureCollection;
```

Returns: [Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)