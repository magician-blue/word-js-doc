# Word.ColorFormat class

Package: [word](/en-us/javascript/api/word)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the color formatting of a shape or text in Word.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

| Property | Description |
|---|---|
| [brightness](#brightness) | Specifies the brightness of a specified shape color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral. |
| [context](#context) | The request context associated with the object. This connects the add-in's process to the Office host application's process. |
| [objectThemeColor](#objectthemecolor) | Specifies the theme color for a color format. |
| [rgb](#rgb) | Specifies the red-green-blue (RGB) value of the specified color. You can provide the value in the '#RRGGBB' format. |
| [tintAndShade](#tintandshade) | Specifies the lightening or darkening of a specified shape's color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral. |
| [type](#type) | Returns the shape color type. |

## Methods

| Method | Description |
|---|---|
| [load(options)](#loadoptions) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [load(propertyNames)](#loadpropertynames) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [load(propertyNamesAndPaths)](#loadpropertynamesandpaths) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [set(properties, options)](#setproperties-options) | Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type. |
| [set(properties)](#setproperties) | Sets multiple properties on the object at the same time, based on an existing loaded object. |
| [toJSON()](#tojson) | Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. |
| [track()](#track) | Track the object for automatic adjustment based on surrounding changes in the document. |
| [untrack()](#untrack) | Release the memory associated with this object, if it has previously been tracked. |

## Property Details

### brightness

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the brightness of a specified shape color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral.

```typescript
brightness: number;
```

Property Value
- number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### objectThemeColor

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the theme color for a color format.

```typescript
objectThemeColor: Word.ThemeColorIndex | "NotThemeColor" | "MainDark1" | "MainLight1" | "MainDark2" | "MainLight2" | "Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6" | "Hyperlink" | "HyperlinkFollowed" | "Background1" | "Text1" | "Background2" | "Text2";
```

Property Value
- [Word.ThemeColorIndex](/en-us/javascript/api/word/word.themecolorindex) | "NotThemeColor" | "MainDark1" | "MainLight1" | "MainDark2" | "MainLight2" | "Accent1" | "Accent2" | "Accent3" | "Accent4" | "Accent5" | "Accent6" | "Hyperlink" | "HyperlinkFollowed" | "Background1" | "Text1" | "Background2" | "Text2"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### rgb

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the specified color. You can provide the value in the '#RRGGBB' format.

```typescript
rgb: string;
```

Property Value
- string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tintAndShade

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the lightening or darkening of a specified shape's color. Valid values are from `-1` (darkest) to `1` (lightest), `0` represents neutral.

```typescript
tintAndShade: number;
```

Property Value
- number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the shape color type.

```typescript
readonly type: Word.ColorType | "rgb" | "scheme";
```

Property Value
- [Word.ColorType](/en-us/javascript/api/word/word.colortype) | "rgb" | "scheme"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ColorFormatLoadOptions): Word.ColorFormat;
```

Parameters
- options: [Word.Interfaces.ColorFormatLoadOptions](/en-us/javascript/api/word/word.interfaces.colorformatloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

### load(propertyNames)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ColorFormat;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

### load(propertyNamesAndPaths)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
    select?: string;
    expand?: string;
}): Word.ColorFormat;
```

Parameters
- propertyNamesAndPaths:  
  {
  select?: string;  
  expand?: string;  
  }  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

### set(properties, options)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ColorFormatUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.ColorFormatUpdateData](/en-us/javascript/api/word/word.interfaces.colorformatupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ColorFormat): void;
```

Parameters
- properties: [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

Returns
- void

### toJSON()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ColorFormat` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ColorFormatData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ColorFormatData;
```

Returns
- [Word.Interfaces.ColorFormatData](/en-us/javascript/api/word/word.interfaces.colorformatdata)

### track()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ColorFormat;
```

Returns
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)

### untrack()

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ColorFormat;
```

Returns
- [Word.ColorFormat](/en-us/javascript/api/word/word.colorformat)