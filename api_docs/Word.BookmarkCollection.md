# Word.BookmarkCollection class

- Package: https://learn.microsoft.com/en-us/javascript/api/word

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

A collection of Word.Bookmark objects that represent the bookmarks in the specified selection, range, or document.

- Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items
  - Gets the loaded child items in this collection.

## Methods
- exists(name)
  - Determines whether the specified bookmark exists.
- getItem(index)
  - Gets a Bookmark object by its index in the collection.
- load(options)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BookmarkCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BookmarkCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- Word.RequestContext: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### items

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Bookmark[];
```

Property Value
- Word.Bookmark[]: https://learn.microsoft.com/en-us/javascript/api/word/word.bookmark

## Method Details

### exists(name)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Determines whether the specified bookmark exists.

```typescript
exists(name: string): OfficeExtension.ClientResult<boolean>;
```

Parameters
- name: string  
  A bookmark name than cannot include more than 40 characters or more than one word.

Returns
- OfficeExtension.ClientResult<boolean>: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult

true if the bookmark exists.

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### getItem(index)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a Bookmark object by its index in the collection.

```typescript
getItem(index: number): Word.Bookmark;
```

Parameters
- index: number  
  A number that identifies the index location of a Bookmark object.

Returns
- Word.Bookmark: https://learn.microsoft.com/en-us/javascript/api/word/word.bookmark

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.BookmarkCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.BookmarkCollection;
```

Parameters
- options: Word.Interfaces.BookmarkCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions  
  Provides options for which properties of the object to load.

Returns
- Word.BookmarkCollection: https://learn.microsoft.com/en-us/javascript/api/word/word.bookmarkcollection

### load(propertyNames)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BookmarkCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- Word.BookmarkCollection: https://learn.microsoft.com/en-us/javascript/api/word/word.bookmarkcollection

### load(propertyNamesAndPaths)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.BookmarkCollection;
```

Parameters
- propertyNamesAndPaths: OfficeExtension.LoadOption: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- Word.BookmarkCollection: https://learn.microsoft.com/en-us/javascript/api/word/word.bookmarkcollection

### toJSON()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BookmarkCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BookmarkCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.BookmarkCollectionData;
```

Returns
- Word.Interfaces.BookmarkCollectionData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.bookmarkcollectiondata

### track()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BookmarkCollection;
```

Returns
- Word.BookmarkCollection: https://learn.microsoft.com/en-us/javascript/api/word/word.bookmarkcollection

### untrack()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BookmarkCollection;
```

Returns
- Word.BookmarkCollection: https://learn.microsoft.com/en-us/javascript/api/word/word.bookmarkcollection