# Word.Frame class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a frame. The Frame object is a member of the [Word.FrameCollection](/en-us/javascript/api/word/word.framecollection) object.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- borders: Returns a BorderUniversalCollection object that represents all the borders for the frame.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- height: Specifies the height (in points) of the frame.
- heightRule: Specifies a FrameSizeRule value that represents the rule for determining the height of the frame.
- horizontalDistanceFromText: Specifies the horizontal distance between the frame and the surrounding text, in points.
- horizontalPosition: Specifies the horizontal distance between the edge of the frame and the item specified by the relativeHorizontalPosition property.
- lockAnchor: Specifies if the frame is locked.
- range: Returns a Range object that represents the portion of the document that's contained within the frame.
- relativeHorizontalPosition: Specifies the relative horizontal position of the frame.
- relativeVerticalPosition: Specifies the relative vertical position of the frame.
- shading: Returns a ShadingUniversal object that refers to the shading formatting for the frame.
- textWrap: Specifies if document text wraps around the frame.
- verticalDistanceFromText: Specifies the vertical distance (in points) between the frame and the surrounding text.
- verticalPosition: Specifies the vertical distance between the edge of the frame and the item specified by the relativeVerticalPosition property.
- width: Specifies the width (in points) of the frame.
- widthRule: Specifies the rule used to determine the width of the frame.

## Methods
- copy(): Copies the frame to the Clipboard.
- cut(): Removes the frame from the document and places it on the Clipboard.
- delete(): Deletes the frame.
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- select(): Selects the frame.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). Whereas the original Word.Frame object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FrameData) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### borders
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders for the frame.

```typescript
readonly borders: Word.BorderUniversalCollection;
```

Property value: [Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### height
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the height (in points) of the frame.

```typescript
height: number;
```

Property value: number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### heightRule
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a FrameSizeRule value that represents the rule for determining the height of the frame.

```typescript
heightRule: Word.FrameSizeRule | "Auto" | "AtLeast" | "Exact";
```

Property value: [Word.FrameSizeRule](/en-us/javascript/api/word/word.framesizerule) | "Auto" | "AtLeast" | "Exact"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### horizontalDistanceFromText
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal distance between the frame and the surrounding text, in points.

```typescript
horizontalDistanceFromText: number;
```

Property value: number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### horizontalPosition
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal distance between the edge of the frame and the item specified by the relativeHorizontalPosition property.

```typescript
horizontalPosition: number;
```

Property value: number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lockAnchor
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the frame is locked.

```typescript
lockAnchor: boolean;
```

Property value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Range object that represents the portion of the document that's contained within the frame.

```typescript
readonly range: Word.Range;
```

Property value: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### relativeHorizontalPosition
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the relative horizontal position of the frame.

```typescript
relativeHorizontalPosition: Word.RelativeHorizontalPosition | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin";
```

Property value: [Word.RelativeHorizontalPosition](/en-us/javascript/api/word/word.relativehorizontalposition) | "Margin" | "Page" | "Column" | "Character" | "LeftMargin" | "RightMargin" | "InsideMargin" | "OutsideMargin"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### relativeVerticalPosition
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the relative vertical position of the frame.

```typescript
relativeVerticalPosition: Word.RelativeVerticalPosition | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin";
```

Property value: [Word.RelativeVerticalPosition](/en-us/javascript/api/word/word.relativeverticalposition) | "Margin" | "Page" | "Paragraph" | "Line" | "TopMargin" | "BottomMargin" | "InsideMargin" | "OutsideMargin"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### shading
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the frame.

```typescript
readonly shading: Word.ShadingUniversal;
```

Property value: [Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textWrap
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if document text wraps around the frame.

```typescript
textWrap: boolean;
```

Property value: boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalDistanceFromText
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical distance (in points) between the frame and the surrounding text.

```typescript
verticalDistanceFromText: number;
```

Property value: number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalPosition
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical distance between the edge of the frame and the item specified by the relativeVerticalPosition property.

```typescript
verticalPosition: number;
```

Property value: number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in points) of the frame.

```typescript
width: number;
```

Property value: number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### widthRule
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the rule used to determine the width of the frame.

```typescript
widthRule: Word.FrameSizeRule | "Auto" | "AtLeast" | "Exact";
```

Property value: [Word.FrameSizeRule](/en-us/javascript/api/word/word.framesizerule) | "Auto" | "AtLeast" | "Exact"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### copy()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Copies the frame to the Clipboard.

```typescript
copy(): void;
```

Returns: void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### cut()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes the frame from the document and places it on the Clipboard.

```typescript
cut(): void;
```

Returns: void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### delete()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes the frame.

```typescript
delete(): void;
```

Returns: void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.FrameLoadOptions): Word.Frame;
```

Parameters
- options: [Word.Interfaces.FrameLoadOptions](/en-us/javascript/api/word/word.interfaces.frameloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.Frame](/en-us/javascript/api/word/word.frame)

### load(propertyNames)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Frame;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.Frame](/en-us/javascript/api/word/word.frame)

### load(propertyNamesAndPaths)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.Frame;
```

Parameters
- propertyNamesAndPaths: {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.Frame](/en-us/javascript/api/word/word.frame)

### select()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Selects the frame.

```typescript
select(): void;
```

Returns: void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### set(properties, options)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.FrameUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.FrameUpdateData](/en-us/javascript/api/word/word.interfaces.frameupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Frame): void;
```

Parameters
- properties: [Word.Frame](/en-us/javascript/api/word/word.frame)

Returns: void

### toJSON()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Frame object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.FrameData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.FrameData;
```

Returns: [Word.Interfaces.FrameData](/en-us/javascript/api/word/word.interfaces.framedata)

### track()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Frame;
```

Returns: [Word.Frame](/en-us/javascript/api/word/word.frame)

### untrack()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Frame;
```

Returns: [Word.Frame](/en-us/javascript/api/word/word.frame)