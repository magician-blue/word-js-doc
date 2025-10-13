# Word.ShadowFormat class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the shadow formatting for a shape or text in Word.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- blur  
  Specifies the blur level for a shadow format as a value between 0.0 and 100.0.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- foregroundColor  
  Returns a ColorFormat object that represents the foreground color for the fill, line, or shadow.
- isVisible  
  Specifies whether the object or the formatting applied to it is visible.
- obscured  
  Specifies true if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill, false if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.
- offsetX  
  Specifies the horizontal offset (in points) of the shadow from the shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.
- offsetY  
  Specifies the vertical offset (in points) of the shadow from the shape. A positive value offsets the shadow to the top of the shape; a negative value offsets it to the bottom.
- rotateWithShape  
  Specifies whether to rotate the shadow when rotating the shape.
- size  
  Specifies the width of the shadow.
- style  
  Specifies the type of shadow formatting to apply to a shape.
- transparency  
  Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).
- type  
  Specifies the shape shadow type.

## Methods

- incrementOffsetX(increment)  
  Changes the horizontal offset of the shadow by the number of points. Increment The number of points to adjust.
- incrementOffsetY(increment)  
  Changes the vertical offset of the shadow by the specified number of points. Increment The number of points to adjust.
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
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ShadowFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShadowFormatData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### blur

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the blur level for a shadow format as a value between 0.0 and 100.0.

```typescript
blur: number;
```

Property Value  
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value  
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### foregroundColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ColorFormat` object that represents the foreground color for the fill, line, or shadow.

```typescript
readonly foregroundColor: Word.ColorFormat;
```

Property Value  
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the object or the formatting applied to it is visible.

```typescript
isVisible: boolean;
```

Property Value  
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### obscured

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies `true` if the shadow of the shape appears filled in and is obscured by the shape, even if the shape has no fill, `false` if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill.

```typescript
obscured: boolean;
```

Property Value  
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### offsetX

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal offset (in points) of the shadow from the shape. A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.

```typescript
offsetX: number;
```

Property Value  
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### offsetY

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical offset (in points) of the shadow from the shape. A positive value offsets the shadow to the top of the shape; a negative value offsets it to the bottom.

```typescript
offsetY: number;
```

Property Value  
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rotateWithShape

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to rotate the shadow when rotating the shape.

```typescript
rotateWithShape: boolean;
```

Property Value  
- boolean

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### size

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the shadow.

```typescript
size: number;
```

Property Value  
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### style

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the type of shadow formatting to apply to a shape.

```typescript
style: Word.ShadowStyle | "Mixed" | "OuterShadow" | "InnerShadow";
```

Property Value  
- [Word.ShadowStyle](/en-us/javascript/api/word/word.shadowstyle) | "Mixed" | "OuterShadow" | "InnerShadow"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the shadow as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency: number;
```

Property Value  
- number

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the shape shadow type.

```typescript
type: Word.ShadowType | "Mixed" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9" | "Type10" | "Type11" | "Type12" | "Type13" | "Type14" | "Type15" | "Type16" | "Type17" | "Type18" | "Type19" | "Type20" | "Type21" | "Type22" | "Type23" | "Type24" | "Type25" | "Type26" | "Type27" | "Type28" | "Type29" | "Type30" | "Type31" | "Type32" | "Type33" | "Type34" | "Type35" | "Type36" | "Type37" | "Type38" | "Type39" | "Type40" | "Type41" | "Type42" | "Type43";
```

Property Value  
- [Word.ShadowType](/en-us/javascript/api/word/word.shadowtype) | "Mixed" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9" | "Type10" | "Type11" | "Type12" | "Type13" | "Type14" | "Type15" | "Type16" | "Type17" | "Type18" | "Type19" | "Type20" | "Type21" | "Type22" | "Type23" | "Type24" | "Type25" | "Type26" | "Type27" | "Type28" | "Type29" | "Type30" | "Type31" | "Type32" | "Type33" | "Type34" | "Type35" | "Type36" | "Type37" | "Type38" | "Type39" | "Type40" | "Type41" | "Type42" | "Type43"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### incrementOffsetX(increment)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Changes the horizontal offset of the shadow by the number of points. Increment The number of points to adjust.

```typescript
incrementOffsetX(increment: number): void;
```

Parameters
- increment: number

Returns  
- void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### incrementOffsetY(increment)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Changes the vertical offset of the shadow by the specified number of points. Increment The number of points to adjust.

```typescript
incrementOffsetY(increment: number): void;
```

Parameters
- increment: number

Returns  
- void

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ShadowFormatLoadOptions): Word.ShadowFormat;
```

Parameters
- options: [Word.Interfaces.ShadowFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.shadowformatloadoptions)  
  Provides options for which properties of the object to load.

Returns  
- [Word.ShadowFormat](/en-us/javascript/api/word/word.shadowformat)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ShadowFormat;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns  
- [Word.ShadowFormat](/en-us/javascript/api/word/word.shadowformat)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.ShadowFormat;
```

Parameters
- propertyNamesAndPaths:  
  - select?: string  
  - expand?: string

`propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns  
- [Word.ShadowFormat](/en-us/javascript/api/word/word.shadowformat)

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ShadowFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.ShadowFormatUpdateData](/en-us/javascript/api/word/word.interfaces.shadowformatupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns  
- void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ShadowFormat): void;
```

Parameters
- properties: [Word.ShadowFormat](/en-us/javascript/api/word/word.shadowformat)

Returns  
- void

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShadowFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShadowFormatData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ShadowFormatData;
```

Returns  
- [Word.Interfaces.ShadowFormatData](/en-us/javascript/api/word/word.interfaces.shadowformatdata)

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ShadowFormat;
```

Returns  
- [Word.ShadowFormat](/en-us/javascript/api/word/word.shadowformat)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ShadowFormat;
```

Returns  
- [Word.ShadowFormat](/en-us/javascript/api/word/word.shadowformat)