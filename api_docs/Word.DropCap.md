# Word.DropCap class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a dropped capital letter in a Word document.

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- distanceFromText — Gets the distance (in points) between the dropped capital letter and the paragraph text.
- fontName — Gets the name of the font for the dropped capital letter.
- linesToDrop — Gets the height (in lines) of the dropped capital letter.
- position — Gets the position of the dropped capital letter.

## Methods
- clear() — Removes the dropped capital letter formatting.
- enable() — Formats the first character in the specified paragraph as a dropped capital letter.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DropCap object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DropCapData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- Word.RequestContext (https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext)

### distanceFromText
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the distance (in points) between the dropped capital letter and the paragraph text.

```typescript
readonly distanceFromText: number;
```

Property Value
- number

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### fontName
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the font for the dropped capital letter.

```typescript
readonly fontName: string;
```

Property Value
- string

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### linesToDrop
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the height (in lines) of the dropped capital letter.

```typescript
readonly linesToDrop: number;
```

Property Value
- number

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### position
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of the dropped capital letter.

```typescript
readonly position: Word.DropPosition | "None" | "Normal" | "Margin";
```

Property Value
- Word.DropPosition (https://learn.microsoft.com/en-us/javascript/api/word/word.dropposition) | "None" | "Normal" | "Margin"

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### clear()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the dropped capital letter formatting.

```typescript
clear(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### enable()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Formats the first character in the specified paragraph as a dropped capital letter.

```typescript
enable(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.DropCapLoadOptions): Word.DropCap;
```

Parameters
- options: Word.Interfaces.DropCapLoadOptions (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.dropcaploadoptions)  
  Provides options for which properties of the object to load.

Returns
- Word.DropCap (https://learn.microsoft.com/en-us/javascript/api/word/word.dropcap)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.DropCap;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- Word.DropCap (https://learn.microsoft.com/en-us/javascript/api/word/word.dropcap)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.DropCap;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- Word.DropCap (https://learn.microsoft.com/en-us/javascript/api/word/word.dropcap)

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.DropCap object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.DropCapData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.DropCapData;
```

Returns
- Word.Interfaces.DropCapData (https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.dropcapdata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.DropCap;
```

Returns
- Word.DropCap (https://learn.microsoft.com/en-us/javascript/api/word/word.dropcap)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.DropCap;
```

Returns
- Word.DropCap (https://learn.microsoft.com/en-us/javascript/api/word/word.dropcap)