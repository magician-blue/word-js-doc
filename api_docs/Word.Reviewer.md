# Word.Reviewer class

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single reviewer of a document in which changes have been tracked. The Reviewer object is a member of the [Word.ReviewerCollection](/en-us/javascript/api/word/word.reviewercollection) object.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- isVisible  
  Specifies if the Reviewer object is visible.

## Methods
- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked.

## Property Details

### context
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### isVisible
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the Reviewer object is visible.

```typescript
isVisible: boolean;
```

Property value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ReviewerLoadOptions): Word.Reviewer;
```

Parameters
- options: [Word.Interfaces.ReviewerLoadOptions](/en-us/javascript/api/word/word.interfaces.reviewerloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.Reviewer](/en-us/javascript/api/word/word.reviewer)

### load(propertyNames)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Reviewer;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.Reviewer](/en-us/javascript/api/word/word.reviewer)

### load(propertyNamesAndPaths)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.Reviewer;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.Reviewer](/en-us/javascript/api/word/word.reviewer)

### set(properties, options)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ReviewerUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.ReviewerUpdateData](/en-us/javascript/api/word/word.interfaces.reviewerupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Reviewer): void;
```

Parameters
- properties: [Word.Reviewer](/en-us/javascript/api/word/word.reviewer)

Returns
- void

### toJSON()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Reviewer object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ReviewerData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ReviewerData;
```

Returns
- [Word.Interfaces.ReviewerData](/en-us/javascript/api/word/word.interfaces.reviewerdata)

### track()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Reviewer;
```

Returns
- [Word.Reviewer](/en-us/javascript/api/word/word.reviewer)

### untrack()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Reviewer;
```

Returns
- [Word.Reviewer](/en-us/javascript/api/word/word.reviewer)