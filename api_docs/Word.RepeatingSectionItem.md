# Word.RepeatingSectionItem class

- Package: [word](/en-us/javascript/api/word)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single item in a [Word.RepeatingSectionContentControl](/en-us/javascript/api/word/word.repeatingsectioncontentcontrol).

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)]

## Properties

- [context](#context)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [range](#range)  
  Returns the range of this repeating section item, excluding the start and end tags.

## Methods

- [delete()](#delete)  
  Deletes this RepeatingSectionItem object.
- [insertItemAfter()](#insertitemafter)  
  Adds a repeating section item after this item and returns the new item.
- [insertItemBefore()](#insertitembefore)  
  Adds a repeating section item before this item and returns the new item.
- [load(options)](#loadoptions)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNames)](#loadpropertynames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [set(properties, options)](#setproperties-options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#setproperties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- [toJSON()](#tojson)  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- [track()](#track)  
  Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#untrack)  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### range

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the range of this repeating section item, excluding the start and end tags.

```typescript
readonly range: Word.Range;
```

Property value: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### delete()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes this RepeatingSectionItem object.

```typescript
delete(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### insertItemAfter()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds a repeating section item after this item and returns the new item.

```typescript
insertItemAfter(): Word.RepeatingSectionItem;
```

Returns: [Word.RepeatingSectionItem](/en-us/javascript/api/word/word.repeatingsectionitem)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### insertItemBefore()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds a repeating section item before this item and returns the new item.

```typescript
insertItemBefore(): Word.RepeatingSectionItem;
```

Returns: [Word.RepeatingSectionItem](/en-us/javascript/api/word/word.repeatingsectionitem)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.RepeatingSectionItemLoadOptions): Word.RepeatingSectionItem;
```

Parameters:
- options: [Word.Interfaces.RepeatingSectionItemLoadOptions](/en-us/javascript/api/word/word.interfaces.repeatingsectionitemloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.RepeatingSectionItem](/en-us/javascript/api/word/word.repeatingsectionitem)

### load(propertyNames)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.RepeatingSectionItem;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.RepeatingSectionItem](/en-us/javascript/api/word/word.repeatingsectionitem)

### load(propertyNamesAndPaths)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.RepeatingSectionItem;
```

Parameters:
- propertyNamesAndPaths:  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.RepeatingSectionItem](/en-us/javascript/api/word/word.repeatingsectionitem)

### set(properties, options)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.RepeatingSectionItemUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.RepeatingSectionItemUpdateData](/en-us/javascript/api/word/word.interfaces.repeatingsectionitemupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.RepeatingSectionItem): void;
```

Parameters:
- properties: [Word.RepeatingSectionItem](/en-us/javascript/api/word/word.repeatingsectionitem)

Returns: void

### toJSON()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.RepeatingSectionItem object is an API object, the toJSON method returns a plain JavaScript object (typed as [Word.Interfaces.RepeatingSectionItemData](/en-us/javascript/api/word/word.interfaces.repeatingsectionitemdata)) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.RepeatingSectionItemData;
```

Returns: [Word.Interfaces.RepeatingSectionItemData](/en-us/javascript/api/word/word.interfaces.repeatingsectionitemdata)

### track()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.RepeatingSectionItem;
```

Returns: [Word.RepeatingSectionItem](/en-us/javascript/api/word/word.repeatingsectionitem)

### untrack()

> Note
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.RepeatingSectionItem;
```

Returns: [Word.RepeatingSectionItem](/en-us/javascript/api/word/word.repeatingsectionitem)