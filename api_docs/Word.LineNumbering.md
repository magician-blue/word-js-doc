# Word.LineNumbering class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents line numbers in the left margin or to the left of each newspaper-style column.

Extends
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- countBy
  - Specifies the numeric increment for line numbers.
- distanceFromText
  - Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.
- isActive
  - Specifies if line numbering is active for the specified document, section, or sections.
- restartMode
  - Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.
- startingNumber
  - Specifies the starting line number.

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
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.LineNumbering object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.LineNumberingData) that contains shallow copies of any loaded child properties from the original object.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- Word.RequestContext: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### countBy

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the numeric increment for line numbers.

```typescript
countBy: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### distanceFromText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the distance (in points) between the right edge of line numbers and the left edge of the document text.

```typescript
distanceFromText: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isActive

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if line numbering is active for the specified document, section, or sections.

```typescript
isActive: boolean;
```

Property Value
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### restartMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the way line numbering runs; that is, whether it starts over at the beginning of a new page or section, or runs continuously.

```typescript
restartMode: Word.NumberingRule | "RestartContinuous" | "RestartSection" | "RestartPage";
```

Property Value
- Word.NumberingRule: https://learn.microsoft.com/en-us/javascript/api/word/word.numberingrule | "RestartContinuous" | "RestartSection" | "RestartPage"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### startingNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting line number.

```typescript
startingNumber: number;
```

Property Value
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.LineNumberingLoadOptions): Word.LineNumbering;
```

Parameters
- options: Word.Interfaces.LineNumberingLoadOptions
  - Provides options for which properties of the object to load.

Returns
- Word.LineNumbering: https://learn.microsoft.com/en-us/javascript/api/word/word.linenumbering

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.LineNumbering;
```

Parameters
- propertyNames: string | string[]
  - A comma-delimited string or an array of strings that specify the properties to load.

Returns
- Word.LineNumbering: https://learn.microsoft.com/en-us/javascript/api/word/word.linenumbering

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.LineNumbering;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }
  - propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- Word.LineNumbering: https://learn.microsoft.com/en-us/javascript/api/word/word.linenumbering

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.LineNumberingUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: Word.Interfaces.LineNumberingUpdateData
  - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: OfficeExtension.UpdateOptions
  - Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.LineNumbering): void;
```

Parameters
- properties: Word.LineNumbering

Returns
- void

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.LineNumbering object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.LineNumberingData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.LineNumberingData;
```

Returns
- Word.Interfaces.LineNumberingData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.linenumberingdata

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.LineNumbering;
```

Returns
- Word.LineNumbering: https://learn.microsoft.com/en-us/javascript/api/word/word.linenumbering

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.LineNumbering;
```

Returns
- Word.LineNumbering: https://learn.microsoft.com/en-us/javascript/api/word/word.linenumbering