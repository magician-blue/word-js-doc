# Word.HyperlinkCollection class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Contains a collection of https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink objects.

Extends
https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items — Gets the loaded child items in this collection.

## Methods
- add(anchor, options) — Returns a Hyperlink object that represents a new hyperlink added to a range, selection, or document.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON() — Overrides the JavaScript toJSON() method for more useful output when passed to JSON.stringify().
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

# Property Details

## context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

## items
Gets the loaded child items in this collection.

```typescript
readonly items: Word.Hyperlink[];
```

Type: https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink[]

# Method Details

## add(anchor, options)
Returns a Hyperlink object that represents a new hyperlink added to a range, selection, or document.

```typescript
add(anchor: Word.Range, options?: Word.HyperlinkAddOptions): Word.Hyperlink;
```

Parameters
- anchor: https://learn.microsoft.com/en-us/javascript/api/word/word.range  
  Required. The range to which the hyperlink is added.
- options: https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlinkaddoptions  
  Optional. The options to further configure the new hyperlink.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlink

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.HyperlinkCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.HyperlinkCollection;
```

Parameters
- options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkcollectionloadoptions & https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.collectionloadoptions  
  Provides options for which properties of the object to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlinkcollection

## load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.HyperlinkCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlinkcollection

## load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.HyperlinkCollection;
```

Parameters
- propertyNamesAndPaths: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlinkcollection

## toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.HyperlinkCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.HyperlinkCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.HyperlinkCollectionData;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.hyperlinkcollectiondata

## track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.HyperlinkCollection;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.hyperlinkcollection

Reference: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member

## untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done us