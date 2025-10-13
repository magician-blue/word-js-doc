# Word.NoteItemCollection class

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.NoteItem](/en-us/javascript/api/word/word.noteitem) objects.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the first footnote in the document body and select its reference mark.
await Word.run(async (context) => {
  const reference: Word.Range = context.document.body.footnotes.getFirst().reference;
  reference.select();
  console.log("The first footnote is selected.");
});
```

## Properties
- [context](#word-word-noteitemcollection-context-member)
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#word-word-noteitemcollection-items-member)
  - Gets the loaded child items in this collection.

## Methods
- [getFirst()](#word-word-noteitemcollection-getfirst-member1)
  - Gets the first note item in this collection. Throws an `ItemNotFound` error if this collection is empty.
- [getFirstOrNullObject()](#word-word-noteitemcollection-getfirstornullobject-member1)
  - Gets the first note item in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [load(options)](#word-word-noteitemcollection-load-member1)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#word-word-noteitemcollection-load-member2)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-noteitemcollection-load-member3)
  - Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [toJSON()](#word-word-noteitemcollection-tojson-member1)
  - Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.NoteItemCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.NoteItemCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- [track()](#word-word-noteitemcollection-track-member1)
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#word-word-noteitemcollection-untrack-member1)
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
Gets the loaded child items in this collection.

```typescript
readonly items: Word.NoteItem[];
```

Property Value  
[Word.NoteItem](/en-us/javascript/api/word/word.noteitem)[]

## Method Details

### getFirst()
Gets the first note item in this collection. Throws an `ItemNotFound` error if this collection is empty.

```typescript
getFirst(): Word.NoteItem;
```

Returns  
[Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the first footnote in the document body and select its reference mark.
await Word.run(async (context) => {
  const reference: Word.Range = context.document.body.footnotes.getFirst().reference;
  reference.select();
  console.log("The first footnote is selected.");
});
```

### getFirstOrNullObject()
Gets the first note item in this collection. If this collection is empty, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.NoteItem;
```

Returns  
[Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.NoteItemCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.NoteItemCollection;
```

Parameters
- options: [Word.Interfaces.NoteItemCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.noteitemcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)

Provides options for which properties of the object to load.

Returns  
[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.NoteItemCollection;
```

Parameters
- propertyNames: string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.NoteItemCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)

`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.NoteItemCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.NoteItemCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.NoteItemCollectionData;
```

Returns  
[Word.Interfaces.NoteItemCollectionData](/en-us/javascript/api/word/word.interfaces.noteitemcollectiondata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.NoteItemCollection;
```

Returns  
[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.NoteItemCollection;
```

Returns  
[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)