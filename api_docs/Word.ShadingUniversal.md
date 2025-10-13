# Word.ShadingUniversal class

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the ShadingUniversal object, which manages shading for a range, paragraph, frame, or table.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)]

## Properties

- backgroundPatternColor: Specifies the color that's applied to the background of the ShadingUniversal object. You can provide the value in the '#RRGGBB' format.
- backgroundPatternColorIndex: Specifies the color that's applied to the background of the ShadingUniversal object.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- foregroundPatternColor: Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern. You can provide the value in the '#RRGGBB' format.
- foregroundPatternColorIndex: Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern.
- texture: Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see [Add, change, or delete the background color in Word](https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515).

## Methods

- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ShadingUniversal object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShadingUniversalData) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### backgroundPatternColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the background of the ShadingUniversal object. You can provide the value in the '#RRGGBB' format.

```typescript
backgroundPatternColor: string;
```

Property Value: string

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

### backgroundPatternColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the background of the ShadingUniversal object.

```typescript
backgroundPatternColorIndex: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

Property Value: [Word.ColorIndex](/en-us/javascript/api/word/word.colorindex) | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### foregroundPatternColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern. You can provide the value in the '#RRGGBB' format.

```typescript
foregroundPatternColor: string;
```

Property Value: string

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

### foregroundPatternColorIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the color that's applied to the foreground of the ShadingUniversal object. This color is applied to the dots and lines in the shading pattern.

```typescript
foregroundPatternColorIndex: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

Property Value: [Word.ColorIndex](/en-us/javascript/api/word/word.colorindex) | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

### texture

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see [Add, change, or delete the background color in Word](https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515).

```typescript
texture: Word.ShadingTextureType | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid";
```

Property Value: [Word.ShadingTextureType](/en-us/javascript/api/word/word.shadingtexturetype) | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)]

## Method Details

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ShadingUniversalLoadOptions): Word.ShadingUniversal;
```

Parameters
- options: [Word.Interfaces.ShadingUniversalLoadOptions](/en-us/javascript/api/word/word.interfaces.shadinguniversalloadoptions)
  - Provides options for which properties of the object to load.

Returns
- [Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ShadingUniversal;
```

Parameters
- propertyNames: string | string[]
  - A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ShadingUniversal;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }
  - propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ShadingUniversalUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.ShadingUniversalUpdateData](/en-us/javascript/api/word/word.interfaces.shadinguniversalupdatedata)
  - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)
  - Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ShadingUniversal): void;
```

Parameters
- properties: [Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)

Returns
- void

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ShadingUniversal object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShadingUniversalData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ShadingUniversalData;
```

Returns
- [Word.Interfaces.ShadingUniversalData](/en-us/javascript/api/word/word.interfaces.shadinguniversaldata)

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ShadingUniversal;
```

Returns
- [Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.ShadingUniversal;
```

Returns
- [Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)