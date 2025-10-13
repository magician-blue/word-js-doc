# Word.LineFormat class

Package: [word](/en-us/javascript/api/word)

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents line and arrowhead formatting. For a line, the LineFormat object contains formatting information for the line itself; for a shape with a border, this object contains formatting information for the shape's border.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

| Property | Description |
| --- | --- |
| [backgroundColor](#word-word-lineformat-backgroundcolor-member) | Gets a ColorFormat object that represents the background color for a patterned line. |
| [beginArrowheadLength](#word-word-lineformat-beginarrowheadlength-member) | Specifies the length of the arrowhead at the beginning of the line. |
| [beginArrowheadStyle](#word-word-lineformat-beginarrowheadstyle-member) | Specifies the style of the arrowhead at the beginning of the line. |
| [beginArrowheadWidth](#word-word-lineformat-beginarrowheadwidth-member) | Specifies the width of the arrowhead at the beginning of the line. |
| [context](#word-word-lineformat-context-member) | The request context associated with the object. This connects the add-in's process to the Office host application's process. |
| [dashStyle](#word-word-lineformat-dashstyle-member) | Specifies the dash style for the line. |
| [endArrowheadLength](#word-word-lineformat-endarrowheadlength-member) | Specifies the length of the arrowhead at the end of the line. |
| [endArrowheadStyle](#word-word-lineformat-endarrowheadstyle-member) | Specifies the style of the arrowhead at the end of the line. |
| [endArrowheadWidth](#word-word-lineformat-endarrowheadwidth-member) | Specifies the width of the arrowhead at the end of the line. |
| [foregroundColor](#word-word-lineformat-foregroundcolor-member) | Gets a ColorFormat object that represents the foreground color for the line. |
| [insetPen](#word-word-lineformat-insetpen-member) | Specifies if to draw lines inside a shape. |
| [isVisible](#word-word-lineformat-isvisible-member) | Specifies if the object, or the formatting applied to it, is visible. |
| [pattern](#word-word-lineformat-pattern-member) | Specifies the pattern applied to the line. |
| [style](#word-word-lineformat-style-member) | Specifies the line format style. |
| [transparency](#word-word-lineformat-transparency-member) | Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear). |
| [weight](#word-word-lineformat-weight-member) | Specifies the thickness of the line in points. |

## Methods

| Method | Description |
| --- | --- |
| [load(options)](#word-word-lineformat-load-member1) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [load(propertyNames)](#word-word-lineformat-load-member2) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [load(propertyNamesAndPaths)](#word-word-lineformat-load-member3) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [set(properties, options)](#word-word-lineformat-set-member1) | Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type. |
| [set(properties)](#word-word-lineformat-set-member2) | Sets multiple properties on the object at the same time, based on an existing loaded object. |
| [toJSON()](#word-word-lineformat-tojson-member1) | Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). |
| [track()](#word-word-lineformat-track-member1) | Track the object for automatic adjustment based on surrounding changes in the document. |
| [untrack()](#word-word-lineformat-untrack-member1) | Release the memory associated with this object, if it has previously been tracked. |

## Property Details

<a id="word-word-lineformat-backgroundcolor-member"></a>
### backgroundColor

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a ColorFormat object that represents the background color for a patterned line.

```typescript
readonly backgroundColor: Word.ColorFormat;
```

Property Value
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-beginarrowheadlength-member"></a>
### beginArrowheadLength

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the arrowhead at the beginning of the line.

```typescript
beginArrowheadLength: Word.ArrowheadLength | "Mixed" | "Short" | "Medium" | "Long";
```

Property Value
- [Word.ArrowheadLength](/en-us/javascript/api/word/word.arrowheadlength) | "Mixed" | "Short" | "Medium" | "Long"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-beginarrowheadstyle-member"></a>
### beginArrowheadStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the style of the arrowhead at the beginning of the line.

```typescript
beginArrowheadStyle: Word.ArrowheadStyle | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval";
```

Property Value
- [Word.ArrowheadStyle](/en-us/javascript/api/word/word.arrowheadstyle) | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-beginarrowheadwidth-member"></a>
### beginArrowheadWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the arrowhead at the beginning of the line.

```typescript
beginArrowheadWidth: Word.ArrowheadWidth | "Mixed" | "Narrow" | "Medium" | "Wide";
```

Property Value
- [Word.ArrowheadWidth](/en-us/javascript/api/word/word.arrowheadwidth) | "Mixed" | "Narrow" | "Medium" | "Wide"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-context-member"></a>
### context

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

<a id="word-word-lineformat-dashstyle-member"></a>
### dashStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the dash style for the line.

```typescript
dashStyle: Word.LineDashStyle | "Mixed" | "Solid" | "SquareDot" | "RoundDot" | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "SysDash" | "SysDot" | "SysDashDot";
```

Property Value
- [Word.LineDashStyle](/en-us/javascript/api/word/word.linedashstyle) | "Mixed" | "Solid" | "SquareDot" | "RoundDot" | "Dash" | "DashDot" | "DashDotDot" | "LongDash" | "LongDashDot" | "LongDashDotDot" | "SysDash" | "SysDot" | "SysDashDot"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-endarrowheadlength-member"></a>
### endArrowheadLength

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the length of the arrowhead at the end of the line.

```typescript
endArrowheadLength: Word.ArrowheadLength | "Mixed" | "Short" | "Medium" | "Long";
```

Property Value
- [Word.ArrowheadLength](/en-us/javascript/api/word/word.arrowheadlength) | "Mixed" | "Short" | "Medium" | "Long"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-endarrowheadstyle-member"></a>
### endArrowheadStyle

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the style of the arrowhead at the end of the line.

```typescript
endArrowheadStyle: Word.ArrowheadStyle | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval";
```

Property Value
- [Word.ArrowheadStyle](/en-us/javascript/api/word/word.arrowheadstyle) | "Mixed" | "None" | "Triangle" | "Open" | "Stealth" | "Diamond" | "Oval"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-endarrowheadwidth-member"></a>
### endArrowheadWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the arrowhead at the end of the line.

```typescript
endArrowheadWidth: Word.ArrowheadWidth | "Mixed" | "Narrow" | "Medium" | "Wide";
```

Property Value
- [Word.ArrowheadWidth](/en-us/javascript/api/word/word.arrowheadwidth) | "Mixed" | "Narrow" | "Medium" | "Wide"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-foregroundcolor-member"></a>
### foregroundColor

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a ColorFormat object that represents the foreground color for the line.

```typescript
readonly foregroundColor: Word.ColorFormat;
```

Property Value
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-insetpen-member"></a>
### insetPen

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if to draw lines inside a shape.

```typescript
insetPen: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-isvisible-member"></a>
### isVisible

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the object, or the formatting applied to it, is visible.

```typescript
isVisible: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-pattern-member"></a>
### pattern

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the pattern applied to the line.

```typescript
pattern: Word.PatternType | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross";
```

Property Value
- [Word.PatternType](/en-us/javascript/api/word/word.patterntype) | "Mixed" | "Percent5" | "Percent10" | "Percent20" | "Percent25" | "Percent30" | "Percent40" | "Percent50" | "Percent60" | "Percent70" | "Percent75" | "Percent80" | "Percent90" | "DarkHorizontal" | "DarkVertical" | "DarkDownwardDiagonal" | "DarkUpwardDiagonal" | "SmallCheckerBoard" | "Trellis" | "LightHorizontal" | "LightVertical" | "LightDownwardDiagonal" | "LightUpwardDiagonal" | "SmallGrid" | "DottedDiamond" | "WideDownwardDiagonal" | "WideUpwardDiagonal" | "DashedUpwardDiagonal" | "DashedDownwardDiagonal" | "NarrowVertical" | "NarrowHorizontal" | "DashedVertical" | "DashedHorizontal" | "LargeConfetti" | "LargeGrid" | "HorizontalBrick" | "LargeCheckerBoard" | "SmallConfetti" | "ZigZag" | "SolidDiamond" | "DiagonalBrick" | "OutlinedDiamond" | "Plaid" | "Sphere" | "Weave" | "DottedGrid" | "Divot" | "Shingle" | "Wave" | "Horizontal" | "Vertical" | "Cross" | "DownwardDiagonal" | "UpwardDiagonal" | "DiagonalCross"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-style-member"></a>
### style

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the line format style.

```typescript
style: Word.LineFormatStyle | "Mixed" | "Single" | "ThinThin" | "ThinThick" | "ThickThin" | "ThickBetweenThin";
```

Property Value
- [Word.LineFormatStyle](/en-us/javascript/api/word/word.lineformatstyle) | "Mixed" | "Single" | "ThinThin" | "ThinThick" | "ThickThin" | "ThickBetweenThin"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-transparency-member"></a>
### transparency

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency of the line as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency: number;
```

Property Value
- number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

<a id="word-word-lineformat-weight-member"></a>
### weight

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the thickness of the line in points.

```typescript
weight: number;
```

Property Value
- number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

<a id="word-word-lineformat-load-member(1)"></a>
### load(options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.LineFormatLoadOptions): Word.LineFormat;
```

Parameters
- options: [Word.Interfaces.LineFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.lineformatloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.LineFormat](/en-us/javascript/api/word/word.lineformat)

<a id="word-word-lineformat-load-member(2)"></a>
### load(propertyNames)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.LineFormat;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.LineFormat](/en-us/javascript/api/word/word.lineformat)

<a id="word-word-lineformat-load-member(3)"></a>
### load(propertyNamesAndPaths)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.LineFormat;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.LineFormat](/en-us/javascript/api/word/word.lineformat)

<a id="word-word-lineformat-set-member(1)"></a>
### set(properties, options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.LineFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.LineFormatUpdateData](/en-us/javascript/api/word/word.interfaces.lineformatupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

<a id="word-word-lineformat-set-member(2)"></a>
### set(properties)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.LineFormat): void;
```

Parameters
- properties: [Word.LineFormat](/en-us/javascript/api/word/word.lineformat)

Returns
- void

<a id="word-word-lineformat-tojson-member(1)"></a>
### toJSON()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.LineFormat object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.LineFormatData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.LineFormatData;
```

Returns
- [Word.Interfaces.LineFormatData](/en-us/javascript/api/word/word.interfaces.lineformatdata)

<a id="word-word-lineformat-track-member(1)"></a>
### track()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.LineFormat;
```

Returns
- [Word.LineFormat](/en-us/javascript/api/word/word.lineformat)

<a id="word-word-lineformat-untrack-member(1)"></a>
### untrack()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.LineFormat;
```

Returns
- [Word.LineFormat](/en-us/javascript/api/word/word.lineformat)