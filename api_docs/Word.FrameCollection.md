# Word.FrameCollection class

Package: [word](/en-us/javascript/api/word)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the collection of [Word.Frame](/en-us/javascript/api/word/word.frame) objects.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items — Gets the loaded child items in this collection.

## Methods
- add(range) — Returns a Frame object that represents a new frame added to a range, selection, or document.
- delete() — Deletes the FrameCollection object.
- getItem(index) — Gets a Frame object by its index in the collection.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.FrameCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FrameCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Frame[];
```

Property Value
- [Word.Frame](/en-us/javascript/api/word/word.frame)[]

## Method Details

### add(range)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Frame` object that represents a new frame added to a range, selection, or document.

```typescript
add(range: Word.Range): Word.Frame;
```

Parameters
- range — [Word.Range](/en-us/javascript/api/word/word.range)  
  The range where the frame will be added.

Returns
- [Word.Frame](/en-us/javascript/api/word/word.frame)  
  A `Frame` object that represents the new frame.

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### delete()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes the `FrameCollection` object.

```typescript
delete(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getItem(index)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `Frame` object by its index in the collection.

```typescript
getItem(index: number): Word.Frame;
```

Parameters
- index — number  
  The location of a `Frame` object.

Returns
- [Word.Frame](/en-us/javascript/api/word/word.frame)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.FrameCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.FrameCollection;
```

Parameters
- options — [Word.Interfaces.FrameCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.framecollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.FrameCollection](/en-us/javascript/api/word/word.framecollection)

### load(propertyNames)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.FrameCollection;
```

Parameters
- propertyNames — string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.FrameCollection](/en-us/javascript/api/word/word.framecollection)

### load(propertyNamesAndPaths)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.FrameCollection;
```

Parameters
- propertyNamesAndPaths — [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.FrameCollection](/en-us/javascript/api/word/word.framecollection)

### toJSON()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.FrameCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.FrameCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.FrameCollectionData;
```

Returns
- [Word.Interfaces.FrameCollectionData](/en-us/javascript/api/word/word.interfaces.framecollectiondata)

### track()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.FrameCollection;
```

Returns
- [Word.FrameCollection](/en-us/javascript/api/word/word.framecollection)

### untrack()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.FrameCollection;
```

Returns
- [Word.FrameCollection](/en-us/javascript/api/word/word.framecollection)