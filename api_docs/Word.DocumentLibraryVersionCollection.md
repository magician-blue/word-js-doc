# Word.DocumentLibraryVersionCollection class

- Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the collection of [Word.DocumentLibraryVersion](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversion) objects.

- Extends: [OfficeExtension.ClientObject](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items  
  Gets the loaded child items in this collection.

## Methods
- getItem(index)  
  Gets a DocumentLibraryVersion object by its index in the collection.
- isVersioningEnabled()  
  Returns whether the document library in which the active document is saved on the server is configured to create a backup copy, or version, each time the file is edited on the website.
- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()  
  Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify().
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
[Word.RequestContext](https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### items
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.DocumentLibraryVersion[];
```

Property Value  
[Word.DocumentLibraryVersion](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversion)[]

## Method Details

### getItem(index)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a DocumentLibraryVersion object by its index in the collection.

```typescript
getItem(index: number): Word.DocumentLibraryVersion;
```

Parameters
- index: number  
  The location of a DocumentLibraryVersion object.

Returns  
[Word.DocumentLibraryVersion](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversion)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVersioningEnabled()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns whether the document library in which the active document is saved on the server is configured to create a backup copy, or version, each time the file is edited on the website.

```typescript
isVersioningEnabled(): OfficeExtension.ClientResult<boolean>;
```

Returns  
[OfficeExtension.ClientResult](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientresult)<boolean>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.DocumentLibraryVersionCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.DocumentLibraryVersionCollection;
```

Parameters
- options: [Word.Interfaces.DocumentLibraryVersionCollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.documentlibraryversioncollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns  
[Word.DocumentLibraryVersionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversioncollection)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.DocumentLibraryVersionCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.DocumentLibraryVersionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversioncollection)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.DocumentLibraryVersionCollection;
```

Parameters
- propertyNamesAndPaths: [OfficeExtension.LoadOption](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.DocumentLibraryVersionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversioncollection)

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DocumentLibraryVersionCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DocumentLibraryVersionCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.DocumentLibraryVersionCollectionData;
```

Returns  
[Word.Interfaces.DocumentLibraryVersionCollectionData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.documentlibraryversioncollectiondata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.DocumentLibraryVersionCollection;
```

Returns  
[Word.DocumentLibraryVersionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversioncollection)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.DocumentLibraryVersionCollection;
```

Returns  
[Word.DocumentLibraryVersionCollection](https://learn.microsoft.com/en-us/javascript/api/word/word.documentlibraryversioncollection)