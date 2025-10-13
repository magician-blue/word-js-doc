# Word.SourceCollection class

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.Source](/en-us/javascript/api/word/word.source) objects.

Extends
[OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

## Properties
- [context](#context) — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#items) — Gets the loaded child items in this collection.

## Methods
- [add(xml)](#addxml) — Adds a new `Source` object to the collection.
- [getItem(index)](#getitemindex) — Gets a `Source` by its index in the collection.
- [load(options)](#loadoptions) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#loadpropertynames) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [toJSON()](#tojson) — Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify()`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SourceCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SourceCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- [track()](#track) — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#untrack) — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

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
readonly items: Word.Source[];
```

Property Value
- [Word.Source](/en-us/javascript/api/word/word.source)[]

## Method Details

### add(xml)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds a new `Source` object to the collection.

```typescript
add(xml: string): Word.Source;
```

Parameters
- xml — string

A string containing the XML data for the source.

Returns
- [Word.Source](/en-us/javascript/api/word/word.source)

A `Source` object that was added to the collection.

Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### getItem(index)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `Source` by its index in the collection.

```typescript
getItem(index: number): Word.Source;
```

Parameters
- index — number

A number that identifies the index location of a `Source` object.

Returns
- [Word.Source](/en-us/javascript/api/word/word.source)

Remarks
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### load(options)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.SourceCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.SourceCollection;
```

Parameters
- options — [Word.Interfaces.SourceCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.sourcecollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)

Provides options for which properties of the object to load.

Returns
- [Word.SourceCollection](/en-us/javascript/api/word/word.sourcecollection)

### load(propertyNames)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.SourceCollection;
```

Parameters
- propertyNames — string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.SourceCollection](/en-us/javascript/api/word/word.sourcecollection)

### load(propertyNamesAndPaths)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.SourceCollection;
```

Parameters
- propertyNamesAndPaths — [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)

`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.SourceCollection](/en-us/javascript/api/word/word.sourcecollection)

### toJSON()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify()`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.SourceCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.SourceCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.SourceCollectionData;
```

Returns
- [Word.Interfaces.SourceCollectionData](/en-us/javascript/api/word/word.interfaces.sourcecollectiondata)

### track()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.SourceCollection;
```

Returns
- [Word.SourceCollection](/en-us/javascript/api/word/word.sourcecollection)

### untrack()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.SourceCollection;
```

Returns
- [Word.SourceCollection](/en-us/javascript/api/word/word.sourcecollection)