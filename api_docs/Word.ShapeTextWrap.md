# Word.ShapeTextWrap class

Represents all the properties for wrapping text around a shape.

- Package: [word](/en-us/javascript/api/word)
- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [bottomDistance](#word-word-shapetextwrap-bottomdistance-member) — Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.
- [context](#word-word-shapetextwrap-context-member) — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [leftDistance](#word-word-shapetextwrap-leftdistance-member) — Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.
- [rightDistance](#word-word-shapetextwrap-rightdistance-member) — Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.
- [side](#word-word-shapetextwrap-side-member) — Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.
- [topDistance](#word-word-shapetextwrap-topdistance-member) — Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.
- [type](#word-word-shapetextwrap-type-member) — Specifies the text wrap type around the shape. See Word.ShapeTextWrapType for details.

## Methods

- [load(options)](#word-word-shapetextwrap-load-member1) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#word-word-shapetextwrap-load-member2) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-shapetextwrap-load-member3) — Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [set(properties, options)](#word-word-shapetextwrap-set-member1) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#word-word-shapetextwrap-set-member2) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- [toJSON()](#word-word-shapetextwrap-tojson-member1) — Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeTextWrap` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeTextWrapData`) that contains shallow copies of any loaded child properties from the original object.
- [track()](#word-word-shapetextwrap-track-member1) — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#word-word-shapetextwrap-untrack-member1) — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

<a id="word-word-shapetextwrap-bottomdistance-member"></a>
### bottomDistance

Specifies the distance (in points) between the document text and the bottom edge of the text-free area surrounding the specified shape.

```typescript
bottomDistance: number;
```

- Property Value: number

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-shapetextwrap-context-member"></a>
### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

<a id="word-word-shapetextwrap-leftdistance-member"></a>
### leftDistance

Specifies the distance (in points) between the document text and the left edge of the text-free area surrounding the specified shape.

```typescript
leftDistance: number;
```

- Property Value: number

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-shapetextwrap-rightdistance-member"></a>
### rightDistance

Specifies the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape.

```typescript
rightDistance: number;
```

- Property Value: number

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-shapetextwrap-side-member"></a>
### side

Specifies whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.

```typescript
side: Word.ShapeTextWrapSide | "None" | "Both" | "Left" | "Right" | "Largest";
```

- Property Value: [Word.ShapeTextWrapSide](/en-us/javascript/api/word/word.shapetextwrapside) | "None" | "Both" | "Left" | "Right" | "Largest"

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-shapetextwrap-topdistance-member"></a>
### topDistance

Specifies the distance (in points) between the document text and the top edge of the text-free area surrounding the specified shape.

```typescript
topDistance: number;
```

- Property Value: number

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-shapetextwrap-type-member"></a>
### type

Specifies the text wrap type around the shape. See `Word.ShapeTextWrapType` for details.

```typescript
type: Word.ShapeTextWrapType | "Inline" | "Square" | "Tight" | "Through" | "TopBottom" | "Behind" | "Front";
```

- Property Value: [Word.ShapeTextWrapType](/en-us/javascript/api/word/word.shapetextwraptype) | "Inline" | "Square" | "Tight" | "Through" | "TopBottom" | "Behind" | "Front"

Remarks  
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

<a id="word-word-shapetextwrap-load-member(1)"></a>
### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ShapeTextWrapLoadOptions): Word.ShapeTextWrap;
```

- Parameters:
  - options: [Word.Interfaces.ShapeTextWrapLoadOptions](/en-us/javascript/api/word/word.interfaces.shapetextwraploadoptions)  
    Provides options for which properties of the object to load.
- Returns: [Word.ShapeTextWrap](/en-us/javascript/api/word/word.shapetextwrap)

<a id="word-word-shapetextwrap-load-member(2)"></a>
### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ShapeTextWrap;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.ShapeTextWrap](/en-us/javascript/api/word/word.shapetextwrap)

<a id="word-word-shapetextwrap-load-member(3)"></a>
### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ShapeTextWrap;
```

- Parameters:
  - propertyNamesAndPaths:  
    {
    select?: string;  
    expand?: string;  
    }  
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.ShapeTextWrap](/en-us/javascript/api/word/word.shapetextwrap)

<a id="word-word-shapetextwrap-set-member(1)"></a>
### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ShapeTextWrapUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

- Parameters:
  - properties: [Word.Interfaces.ShapeTextWrapUpdateData](/en-us/javascript/api/word/word.interfaces.shapetextwrapupdatedata)  
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
    Provides an option to suppress errors if the properties object tries to set any read-only properties.
- Returns: void

<a id="word-word-shapetextwrap-set-member(2)"></a>
### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ShapeTextWrap): void;
```

- Parameters:
  - properties: [Word.ShapeTextWrap](/en-us/javascript/api/word/word.shapetextwrap)
- Returns: void

<a id="word-word-shapetextwrap-tojson-member(1)"></a>
### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeTextWrap` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeTextWrapData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ShapeTextWrapData;
```

- Returns: [Word.Interfaces.ShapeTextWrapData](/en-us/javascript/api/word/word.interfaces.shapetextwrapdata)

<a id="word-word-shapetextwrap-track-member(1)"></a>
### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ShapeTextWrap;
```

- Returns: [Word.ShapeTextWrap](/en-us/javascript/api/word/word.shapetextwrap)

<a id="word-word-shapetextwrap-untrack-member(1)"></a>
### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ShapeTextWrap;
```

- Returns: [Word.ShapeTextWrap](/en-us/javascript/api/word/word.shapetextwrap)