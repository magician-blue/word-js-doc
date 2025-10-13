# Word.ShapeFill class

- Package: [word](/en-us/javascript/api/word)

Represents the fill formatting of a shape object.

- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- backgroundColor
  - Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.
- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- foregroundColor
  - Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.
- transparency
  - Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.
- type
  - Returns the fill type of the shape. See Word.ShapeFillType for details.

## Methods

- clear()
  - Clears the fill formatting of this shape and set it to Word.ShapeFillType.NoFill;
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
- setSolidColor(color)
  - Sets the fill formatting of the shape to a uniform color. This changes the fill type to Word.ShapeFillType.Solid.
- toJSON()
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.ShapeFill object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShapeFillData) that contains shallow copies of any loaded child properties from the original object.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### backgroundColor

Specifies the shape fill background color. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
backgroundColor: string;
```

- Property Value: string

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### foregroundColor

Specifies the shape fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
foregroundColor: string;
```

- Property Value: string

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Specifies the transparency percentage of the fill as a value from 0.0 (opaque) through 1.0 (clear). Returns null if the shape type does not support transparency or the shape fill has inconsistent transparency, such as with a gradient fill type.

```typescript
transparency: number;
```

- Property Value: number

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Returns the fill type of the shape. See Word.ShapeFillType for details.

```typescript
readonly type: Word.ShapeFillType | "NoFill" | "Solid" | "Gradient" | "Pattern" | "Picture" | "Texture" | "Mixed";
```

- Property Value: [Word.ShapeFillType](/en-us/javascript/api/word/word.shapefilltype) | "NoFill" | "Solid" | "Gradient" | "Pattern" | "Picture" | "Texture" | "Mixed"

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### clear()

Clears the fill formatting of this shape and set it to `Word.ShapeFillType.NoFill`;

```typescript
clear(): void;
```

- Returns: void

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.ShapeFillLoadOptions): Word.ShapeFill;
```

- Parameters:
  - options: [Word.Interfaces.ShapeFillLoadOptions](/en-us/javascript/api/word/word.interfaces.shapefillloadoptions)  
    Provides options for which properties of the object to load.
- Returns: [Word.ShapeFill](/en-us/javascript/api/word/word.shapefill)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.ShapeFill;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.ShapeFill](/en-us/javascript/api/word/word.shapefill)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ShapeFill;
```

- Parameters:
  - propertyNamesAndPaths:  
    {  
    select?: string;  
    expand?: string;  
    }  
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.ShapeFill](/en-us/javascript/api/word/word.shapefill)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ShapeFillUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

- Parameters:
  - properties: [Word.Interfaces.ShapeFillUpdateData](/en-us/javascript/api/word/word.interfaces.shapefillupdatedata)  
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
    Provides an option to suppress errors if the properties object tries to set any read-only properties.
- Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.ShapeFill): void;
```

- Parameters:
  - properties: [Word.ShapeFill](/en-us/javascript/api/word/word.shapefill)
- Returns: void

### setSolidColor(color)

Sets the fill formatting of the shape to a uniform color. This changes the fill type to `Word.ShapeFillType.Solid`.

```typescript
setSolidColor(color: string): void;
```

- Parameters:
  - color: string  
    A string that represents the fill foreground color. You can provide the value in the '#RRGGBB' format or the color name.
- Returns: void

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ShapeFill` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ShapeFillData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ShapeFillData;
```

- Returns: [Word.Interfaces.ShapeFillData](/en-us/javascript/api/word/word.interfaces.shapefilldata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.ShapeFill;
```

- Returns: [Word.ShapeFill](/en-us/javascript/api/word/word.shapefill)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.ShapeFill;
```

- Returns: [Word.ShapeFill](/en-us/javascript/api/word/word.shapefill)