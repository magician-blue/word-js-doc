# Word.ReflectionFormat class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the reflection formatting for a shape in Word.

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- blur
  - Specifies the degree of blur effect applied to the ReflectionFormat object as a value between 0.0 and 100.0.
- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- offset
  - Specifies the amount of separation, in points, of the reflected image from the shape.
- size
  - Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.
- transparency
  - Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).
- type
  - Specifies a ReflectionType value that represents the type and direction of the lighting for a shape reflection.

## Methods

- load(options)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options)
  - Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties)
  - Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON()
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ReflectionFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ReflectionFormatData) that contains shallow copies of any loaded child properties from the original object.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### blur

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of blur effect applied to the ReflectionFormat object as a value between 0.0 and 100.0.

```typescript
blur: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### offset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the amount of separation, in points, of the reflected image from the shape.

```typescript
offset: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### size

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.

```typescript
size: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a ReflectionType value that represents the type and direction of the lighting for a shape reflection.

```typescript
type: Word.ReflectionType | "Mixed" | "None" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.reflectiontype | "Mixed" | "None" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ReflectionFormatLoadOptions): Word.ReflectionFormat;
```

Parameters:
- options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.reflectionformatloadoptions  
  Provides options for which properties of the object to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.reflectionformat

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ReflectionFormat;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.reflectionformat

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.ReflectionFormat;
```

Parameters:
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.reflectionformat

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ReflectionFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.reflectionformatupdatedata  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ReflectionFormat): void;
```

Parameters:
- properties: https://learn.microsoft.com/en-us/javascript/api/word/word.reflectionformat

Returns: void

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ReflectionFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ReflectionFormatData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ReflectionFormatData;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.reflectionformatdata

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ReflectionFormat;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.reflectionformat

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member. Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.ReflectionFormat;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.reflectionformat