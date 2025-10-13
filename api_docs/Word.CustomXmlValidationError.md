# Word.CustomXmlValidationError class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single validation error in a [Word.CustomXmlValidationErrorCollection](/en-us/javascript/api/word/word.customxmlvalidationerrorcollection) object.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)]

## Properties

- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- errorCode — Gets an integer representing the validation error in the CustomXmlValidationError object.
- name — Gets the name of the error in the CustomXmlValidationError object.If no errors exist, the property returns Nothing
- node — Gets the node associated with this CustomXmlValidationError object, if any exist.If no nodes exist, the property returns Nothing.
- text — Gets the text in the CustomXmlValidationError object.
- type — Gets the type of error generated from the CustomXmlValidationError object.

## Methods

- delete() — Deletes this CustomXmlValidationError object.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). JSON.stringify, in turn, calls the toJSON method of the object that's passed to it. Whereas the original Word.CustomXmlValidationError object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CustomXmlValidationErrorData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property details

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### errorCode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an integer representing the validation error in the `CustomXmlValidationError` object.

```typescript
readonly errorCode: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

### name

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the error in the `CustomXmlValidationError` object.If no errors exist, the property returns `Nothing`

```typescript
readonly name: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

### node

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the node associated with this `CustomXmlValidationError` object, if any exist.If no nodes exist, the property returns `Nothing`.

```typescript
readonly node: Word.CustomXmlNode;
```

Property value: [Word.CustomXmlNode](/en-us/javascript/api/word/word.customxmlnode)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

### text

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the text in the `CustomXmlValidationError` object.

```typescript
readonly text: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the type of error generated from the `CustomXmlValidationError` object.

```typescript
readonly type: Word.CustomXmlValidationErrorType | "schemaGenerated" | "automaticallyCleared" | "manual";
```

Property value: [Word.CustomXmlValidationErrorType](/en-us/javascript/api/word/word.customxmlvalidationerrortype) | "schemaGenerated" | "automaticallyCleared" | "manual"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

## Method details

### delete()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes this `CustomXmlValidationError` object.

```typescript
delete(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CustomXmlValidationErrorLoadOptions): Word.CustomXmlValidationError;
```

Parameters:
- options: [Word.Interfaces.CustomXmlValidationErrorLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlvalidationerrorloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CustomXmlValidationError;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.CustomXmlValidationError;
```

Parameters:
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror)

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.CustomXmlValidationErrorUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.CustomXmlValidationErrorUpdateData](/en-us/javascript/api/word/word.interfaces.customxmlvalidationerrorupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.CustomXmlValidationError): void;
```

Parameters:
- properties: [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror)

Returns: void

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CustomXmlValidationError` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CustomXmlValidationErrorData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CustomXmlValidationErrorData;
```

Returns: [Word.Interfaces.CustomXmlValidationErrorData](/en-us/javascript/api/word/word.interfaces.customxmlvalidationerrordata)

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CustomXmlValidationError;
```

Returns: [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CustomXmlValidationError;
```

Returns: [Word.CustomXmlValidationError](/en-us/javascript/api/word/word.customxmlvalidationerror)