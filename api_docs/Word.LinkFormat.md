# Word.LinkFormat class

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the linking characteristics for an OLE object or picture.

Package: https://learn.microsoft.com/en-us/javascript/api/word

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- isAutoUpdated
  - Specifies if the link is updated automatically when the container file is opened or when the source file is changed.
- isLocked
  - Specifies if a Field, InlineShape, or Shape object is locked to prevent automatic updating.
- isPictureSavedWithDocument
  - Specifies if the linked picture is saved with the document.
- sourceFullName
  - Specifies the path and name of the source file for the linked OLE object, picture, or field.
- sourceName
  - Gets the name of the source file for the linked OLE object, picture, or field.
- sourcePath
  - Gets the path of the source file for the linked OLE object, picture, or field.
- type
  - Gets the link type.

## Methods

- breakLink()
  - Breaks the link between the source file and the OLE object, picture, or linked field.
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
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.LinkFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.LinkFormatData) that contains shallow copies of any loaded child properties from the original object.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### isAutoUpdated

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the link is updated automatically when the container file is opened or when the source file is changed.

```typescript
isAutoUpdated: boolean;
```

- Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isLocked

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if a Field, InlineShape, or Shape object is locked to prevent automatic updating.

```typescript
isLocked: boolean;
```

- Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isPictureSavedWithDocument

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the linked picture is saved with the document.

```typescript
isPictureSavedWithDocument: boolean;
```

- Property Value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourceFullName

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the path and name of the source file for the linked OLE object, picture, or field.

```typescript
sourceFullName: string;
```

- Property Value: string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourceName

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the source file for the linked OLE object, picture, or field.

```typescript
readonly sourceName: string;
```

- Property Value: string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### sourcePath

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the path of the source file for the linked OLE object, picture, or field.

```typescript
readonly sourcePath: string;
```

- Property Value: string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the link type.

```typescript
readonly type: Word.LinkType | "Ole" | "Picture" | "Text" | "Reference" | "Include" | "Import" | "Dde" | "DdeAuto" | "Chart";
```

- Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.linktype | "Ole" | "Picture" | "Text" | "Reference" | "Include" | "Import" | "Dde" | "DdeAuto" | "Chart"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### breakLink()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Breaks the link between the source file and the OLE object, picture, or linked field.

```typescript
breakLink(): void;
```

- Returns: void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.LinkFormatLoadOptions): Word.LinkFormat;
```

- Parameters:
  - options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.linkformatloadoptions  
    Provides options for which properties of the object to load.
- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.linkformat

### load(propertyNames)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.LinkFormat;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.
- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.linkformat

### load(propertyNamesAndPaths)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.LinkFormat;
```

- Parameters:
  - propertyNamesAndPaths: `{ select?: string; expand?: string; }`  
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.
- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.linkformat

### set(properties, options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.LinkFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

- Parameters:
  - properties: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.linkformatupdatedata  
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - options: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions  
    Provides an option to suppress errors if the properties object tries to set any read-only properties.
- Returns: void

### set(properties)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.LinkFormat): void;
```

- Parameters:
  - properties: https://learn.microsoft.com/en-us/javascript/api/word/word.linkformat
- Returns: void

### toJSON()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.LinkFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.LinkFormatData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.LinkFormatData;
```

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.linkformatdata

### track()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.LinkFormat;
```

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.linkformat

### untrack()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.LinkFormat;
```

- Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.linkformat