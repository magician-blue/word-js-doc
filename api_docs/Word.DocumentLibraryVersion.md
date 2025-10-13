# Word.DocumentLibraryVersion class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a document library version.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- comments — Gets any optional comments associated with this version of the shared document.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- modified — Gets the date and time at which this version of the shared document was last saved to the server.
- modifiedBy — Gets the name of the user who last saved this version of the shared document to the server.

## Methods
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.DocumentLibraryVersion object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DocumentLibraryVersionData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### comments
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets any optional comments associated with this version of the shared document.

```typescript
readonly comments: string;
```

Property Value: string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### modified
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the date and time at which this version of the shared document was last saved to the server.

```typescript
readonly modified: any;
```

Property Value: any

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### modifiedBy
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the user who last saved this version of the shared document to the server.

```typescript
readonly modifiedBy: string;
```

Property Value: string

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.DocumentLibraryVersionLoadOptions): Word.DocumentLibraryVersion;
```

Parameters
- options: [Word.Interfaces.DocumentLibraryVersionLoadOptions](/en-us/javascript/api/word/word.interfaces.documentlibraryversionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.DocumentLibraryVersion](/en-us/javascript/api/word/word.documentlibraryversion)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.DocumentLibraryVersion;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.DocumentLibraryVersion](/en-us/javascript/api/word/word.documentlibraryversion)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.DocumentLibraryVersion;
```

Parameters
- propertyNamesAndPaths:  
  - select?: string  
  - expand?: string

propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.DocumentLibraryVersion](/en-us/javascript/api/word/word.documentlibraryversion)

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DocumentLibraryVersion object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DocumentLibraryVersionData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.DocumentLibraryVersionData;
```

Returns: [Word.Interfaces.DocumentLibraryVersionData](/en-us/javascript/api/word/word.interfaces.documentlibraryversiondata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.DocumentLibraryVersion;
```

Returns: [Word.DocumentLibraryVersion](/en-us/javascript/api/word/word.documentlibraryversion)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.DocumentLibraryVersion;
```

Returns: [Word.DocumentLibraryVersion](/en-us/javascript/api/word/word.documentlibraryversion)