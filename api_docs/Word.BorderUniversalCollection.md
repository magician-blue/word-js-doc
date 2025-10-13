# Word.BorderUniversalCollection class

Package: [word](/en-us/javascript/api/word)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the collection of [Word.BorderUniversal](/en-us/javascript/api/word/word.borderuniversal) objects.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- [context](#word-word-borderuniversalcollection-context-member)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [items](#word-word-borderuniversalcollection-items-member)  
  Gets the loaded child items in this collection.

## Methods
- [applyPageBordersToAllSections()](#word-word-borderuniversalcollection-applypageborderstoallsections-member1)  
  Applies the specified page-border formatting to all sections in the document.
- [getItem(index)](#word-word-borderuniversalcollection-getitem-member1)  
  Gets a `Border` object by its index in the collection.
- [load(options)](#word-word-borderuniversalcollection-load-member1)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#word-word-borderuniversalcollection-load-member2)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-borderuniversalcollection-load-member3)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [toJSON()](#word-word-borderuniversalcollection-tojson-member1)  
  Overrides the JavaScript `toJSON()` method to provide more useful output when an API object is passed to `JSON.stringify()`. The method returns a plain JavaScript object (typed as `Word.Interfaces.BorderUniversalCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- [track()](#word-word-borderuniversalcollection-track-member1)  
  Track the object for automatic adjustment based on surrounding changes in the document. This is shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member).
- [untrack()](#word-word-borderuniversalcollection-untrack-member1)  
  Release the memory associated with this object, if it has previously been tracked. This is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). You must call `context.sync()` before the memory release takes effect.

## Property Details

### context
id: word-word-borderuniversalcollection-context-member

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items
id: word-word-borderuniversalcollection-items-member

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.BorderUniversal[];
```

Property Value
- [Word.BorderUniversal](/en-us/javascript/api/word/word.borderuniversal)[]

## Method Details

### applyPageBordersToAllSections()
id: word-word-borderuniversalcollection-applypageborderstoallsections-member1

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Applies the specified page-border formatting to all sections in the document.

```typescript
applyPageBordersToAllSections(): void;
```

Returns
- void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItem(index)
id: word-word-borderuniversalcollection-getitem-member1

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `Border` object by its index in the collection.

```typescript
getItem(index: number): Word.BorderUniversal;
```

Parameters
- index: number  
  The location of a `BorderUniversal` object.

Returns
- [Word.BorderUniversal](/en-us/javascript/api/word/word.borderuniversal)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
id: word-word-borderuniversalcollection-load-member1

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.BorderUniversalCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.BorderUniversalCollection;
```

Parameters
- options: [Word.Interfaces.BorderUniversalCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.borderuniversalcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)

### load(propertyNames)
id: word-word-borderuniversalcollection-load-member2

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BorderUniversalCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)

### load(propertyNamesAndPaths)
id: word-word-borderuniversalcollection-load-member3

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.BorderUniversalCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)

### toJSON()
id: word-word-borderuniversalcollection-tojson-member1

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.BorderUniversalCollection` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BorderUniversalCollectionData`) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.BorderUniversalCollectionData;
```

Returns
- [Word.Interfaces.BorderUniversalCollectionData](/en-us/javascript/api/word/word.interfaces.borderuniversalcollectiondata)

### track()
id: word-word-borderuniversalcollection-track-member1

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BorderUniversalCollection;
```

Returns
- [Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)

### untrack()
id: word-word-borderuniversalcollection-untrack-member1

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.BorderUniversalCollection;
```

Returns
- [Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)