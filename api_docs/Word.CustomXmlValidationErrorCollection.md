# Word.CustomXmlValidationErrorCollection class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror) objects.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- items  
  Gets the loaded child items in this collection.

## Methods

- add(node, errorName, options)  
  Adds a CustomXmlValidationError object containing an XML validation error to the CustomXmlValidationErrorCollection object.

- getCount()  
  Returns the number of items in the collection.

- getItem(index)  
  Returns a CustomXmlValidationError object that represents the specified item in the collection.

- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlValidationErrorCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlValidationErrorCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the loaded child items in this collection.

```typescript
readonly items: Word.CustomXmlValidationError[];
```

Property Value
- [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror)[]

## Method Details

### add(node, errorName, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Adds a CustomXmlValidationError object containing an XML validation error to the CustomXmlValidationErrorCollection object.

```typescript
add(
  node: Word.CustomXmlNode,
  errorName: string,
  options?: Word.CustomXmlAddValidationErrorOptions
): OfficeExtension.ClientResult<number>;
```

Parameters
- node  
  [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)  
  The XML node where the error occurred.
- errorName  
  string  
  The name of the error.
- options  
  [Word.CustomXmlAddValidationErrorOptions](/en-us/javascript/api/word/word.customxmladdvalidationerroroptions)  
  Optional. The options that define the error message and other settings.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getCount()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the number of items in the collection.

```typescript
getCount(): OfficeExtension.ClientResult<number>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getItem(index)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a CustomXmlValidationError object that represents the specified item in the collection.

```typescript
getItem(index: number): Word.CustomXmlValidationError;
```

Parameters
- index  
  number  
  A number that identifies the index location of a paragraph object.

Returns
- [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(
  options?: Word.Interfaces.CustomXmlValidationErrorCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions
): Word.CustomXmlValidationErrorCollection;
```

Parameters
- options  
  [Word.Interfaces.CustomXmlValidationErrorCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlvalidationerrorcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomXmlValidationErrorCollection;
```

Parameters
- propertyNames  
  string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CustomXmlValidationErrorCollection;
```

Parameters
- propertyNamesAndPaths  
  [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection)

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CustomXmlValidationErrorCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlValidationErrorCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.CustomXmlValidationErrorCollectionData;
```

Returns
- [Word.Interfaces.CustomXmlValidationErrorCollectionData](/en-us/javascript/api/word/word.interfaces.customxmlvalidationerrorcollectiondata)

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomXmlValidationErrorCollection;
```

Returns
- [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.CustomXmlValidationErrorCollection;
```

Returns
- [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection)