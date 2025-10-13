# Word.Border class

Package: [word](/en-us/javascript/api/word)

Represents the Border object for text, a paragraph, or a table.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Updates border properties (e.g., type, width, color) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update border properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const borders: Word.BorderCollection = style.borders;
    borders.load("items");
    await context.sync();

    borders.outsideBorderType = Word.BorderType.dashed;
    borders.outsideBorderWidth = Word.BorderWidth.pt025;
    borders.outsideBorderColor = "green";
    console.log("Updated outside borders.");
  }
});
```

## Properties

| Property | Description |
| --- | --- |
| [color](#color) | Specifies the color for the border. Color is specified in â#RRGGBBâ format or by using the color name. |
| [context](#context) | The request context associated with the object. This connects the add-in's process to the Office host application's process. |
| [location](#location) | Gets the location of the border. |
| [type](#type) | Specifies the border type for the border. |
| [visible](#visible) | Specifies whether the border is visible. |
| [width](#width) | Specifies the width for the border. |

## Methods

| Method | Description |
| --- | --- |
| [load(options)](#loadoptions) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [load(propertyNames)](#loadpropertynames) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [load(propertyNamesAndPaths)](#loadpropertynamesandpaths) | Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties. |
| [set(properties, options)](#setproperties-options) | Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type. |
| [set(properties)](#setproperties) | Sets multiple properties on the object at the same time, based on an existing loaded object. |
| [toJSON()](#tojson) | Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). |
| [track()](#track) | Track the object for automatic adjustment based on surrounding changes in the document. |
| [untrack()](#untrack) | Release the memory associated with this object, if it has previously been tracked. |

## Property Details

### color

Specifies the color for the border. Color is specified in â#RRGGBBâ format or by using the color name.

```typescript
color: string;
```

Property Value
- string

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### location

Gets the location of the border.

```typescript
readonly location: Word.BorderLocation | "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All";
```

Property Value
- [Word.BorderLocation](/en-us/javascript/api/word/word.borderlocation) | "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Specifies the border type for the border.

```typescript
type: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
```

Property Value
- [Word.BorderType](/en-us/javascript/api/word/word.bordertype) | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### visible

Specifies whether the border is visible.

```typescript
visible: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Specifies the width for the border.

```typescript
width: Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed";
```

Property Value
- [Word.BorderWidth](/en-us/javascript/api/word/word.borderwidth) | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"

Remarks  
[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.BorderLoadOptions): Word.Border;
```

Parameters
- options: [Word.Interfaces.BorderLoadOptions](/en-us/javascript/api/word/word.interfaces.borderloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.Border](/en-us/javascript/api/word/word.border)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Border;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.Border](/en-us/javascript/api/word/word.border)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Border;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.Border](/en-us/javascript/api/word/word.border)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.BorderUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.BorderUpdateData](/en-us/javascript/api/word/word.interfaces.borderupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Border): void;
```

Parameters
- properties: [Word.Border](/en-us/javascript/api/word/word.border)

Returns
- void

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Border object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BorderData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.BorderData;
```

Returns
- [Word.Interfaces.BorderData](/en-us/javascript/api/word/word.interfaces.borderdata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Border;
```

Returns
- [Word.Border](/en-us/javascript/api/word/word.border)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Border;
```

Returns
- [Word.Border](/en-us/javascript/api/word/word.border)