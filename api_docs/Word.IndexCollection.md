# Word.IndexCollection class

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

A collection of [Word.Index](/en-us/javascript/api/word/word.index) objects that represents all the indexes in the document.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items
  - Gets the loaded child items in this collection.

## Methods
- add(range, indexAddOptions)
  - Returns an Index object that represents a new index added to the document.
- getFormat()
  - Gets the IndexFormat value that represents the formatting for the indexes in the document.
- getItem(index)
  - Gets an Index object by its index in the collection.
- load(options)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- markAllEntries(range, markAllEntriesOptions)
  - Inserts an XE (Index Entry) field after all instances of the text in the range.
- toJSON()
  - Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify().
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document.
- untrack()
  - Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Index[];
```

Property Value
- [Word.Index](/en-us/javascript/api/word/word.index)[]

## Method Details

### add(range, indexAddOptions)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an Index object that represents a new index added to the document.

```typescript
add(range: Word.Range, indexAddOptions?: Word.IndexAddOptions): Word.Index;
```

Parameters
- range: [Word.Range](/en-us/javascript/api/word/word.range)  
  The range where you want the index to appear. The index replaces the range, if the range is not collapsed.
- indexAddOptions: [Word.IndexAddOptions](/en-us/javascript/api/word/word.indexaddoptions)  
  Optional. The options for adding the index.

Returns
- [Word.Index](/en-us/javascript/api/word/word.index)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getFormat()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the IndexFormat value that represents the formatting for the indexes in the document.

```typescript
getFormat(): OfficeExtension.ClientResult<Word.IndexFormat>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<[Word.IndexFormat](/en-us/javascript/api/word/word.indexformat)>

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getItem(index)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an Index object by its index in the collection.

```typescript
getItem(index: number): Word.Index;
```

Parameters
- index: number  
  A number that identifies the index location of an Index object.

Returns
- [Word.Index](/en-us/javascript/api/word/word.index)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.IndexCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.IndexCollection;
```

Parameters
- options: [Word.Interfaces.IndexCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.indexcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.IndexCollection](/en-us/javascript/api/word/word.indexcollection)

### load(propertyNames)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.IndexCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.IndexCollection](/en-us/javascript/api/word/word.indexcollection)

### load(propertyNamesAndPaths)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.IndexCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.IndexCollection](/en-us/javascript/api/word/word.indexcollection)

### markAllEntries(range, markAllEntriesOptions)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Inserts an [XE (Index Entry) field](https://support.microsoft.com/office/abaf7c78-6e21-418d-bf8b-f8186d2e4d08) after all instances of the text in the range.

```typescript
markAllEntries(range: Word.Range, markAllEntriesOptions?: Word.IndexMarkAllEntriesOptions): void;
```

Parameters
- range: [Word.Range](/en-us/javascript/api/word/word.range)  
  The range whose text is marked with an XE field throughout the document.
- markAllEntriesOptions: [Word.IndexMarkAllEntriesOptions](/en-us/javascript/api/word/word.indexmarkallentriesoptions)  
  Optional. The options for marking all entries.

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### toJSON()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.IndexCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.IndexCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.IndexCollectionData;
```

Returns
- [Word.Interfaces.IndexCollectionData](/en-us/javascript/api/word/word.interfaces.indexcollectiondata)

### track()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.IndexCollection;
```

Returns
- [Word.IndexCollection](/en-us/javascript/api/word/word.indexcollection)

### untrack()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.IndexCollection;
```

Returns
- [Word.IndexCollection](/en-us/javascript/api/word/word.indexcollection)