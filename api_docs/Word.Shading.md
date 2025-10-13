# Word.Shading class

Package: [word](/en-us/javascript/api/word)

Represents the shading object.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.6 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Updates shading properties (e.g., texture, pattern colors) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update shading properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const shading: Word.Shading = style.shading;
    shading.load();
    await context.sync();

    shading.backgroundPatternColor = "blue";
    shading.foregroundPatternColor = "yellow";
    shading.texture = Word.ShadingTextureType.darkTrellis;

    console.log("Updated shading.");
  }
});
```

## Properties

- backgroundPatternColor  
  Specifies the color for the background of the object. You can provide the value in the '#RRGGBB' format or the color name.

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.

- foregroundPatternColor  
  Specifies the color for the foreground of the object. You can provide the value in the '#RRGGBB' format or the color name.

- texture  
  Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see Add, change, or delete the background color in Word.

## Methods

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
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Shading object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShadingData) that contains shallow copies of any loaded child properties from the original object.

- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### backgroundPatternColor

Specifies the color for the background of the object. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
backgroundPatternColor: string;
```

Property Value: string

Remarks  
[ API set: WordApi 1.6 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### foregroundPatternColor

Specifies the color for the foreground of the object. You can provide the value in the '#RRGGBB' format or the color name.

```typescript
foregroundPatternColor: string;
```

Property Value: string

Remarks  
[ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### texture

Specifies the shading texture of the object. To learn more about how to apply backgrounds like textures, see Add, change, or delete the background color in Word.

```typescript
texture: Word.ShadingTextureType | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid";
```

Property Value: [Word.ShadingTextureType](/en-us/javascript/api/word/word.shadingtexturetype) | "DarkDiagonalDown" | "DarkDiagonalUp" | "DarkGrid" | "DarkHorizontal" | "DarkTrellis" | "DarkVertical" | "LightDiagonalDown" | "LightDiagonalUp" | "LightGrid" | "LightHorizontal" | "LightTrellis" | "LightVertical" | "None" | "Percent10" | "Percent12Pt5" | "Percent15" | "Percent20" | "Percent25" | "Percent30" | "Percent35" | "Percent37Pt5" | "Percent40" | "Percent45" | "Percent5" | "Percent50" | "Percent55" | "Percent60" | "Percent62Pt5" | "Percent65" | "Percent70" | "Percent75" | "Percent80" | "Percent85" | "Percent87Pt5" | "Percent90" | "Percent95" | "Solid"

Remarks  
[ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ShadingLoadOptions): Word.Shading;
```

Parameters  
- options: [Word.Interfaces.ShadingLoadOptions](/en-us/javascript/api/word/word.interfaces.shadingloadoptions)  
  Provides options for which properties of the object to load.

Returns  
[Word.Shading](/en-us/javascript/api/word/word.shading)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Shading;
```

Parameters  
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns  
[Word.Shading](/en-us/javascript/api/word/word.shading)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Shading;
```

Parameters  
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns  
[Word.Shading](/en-us/javascript/api/word/word.shading)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ShadingUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters  
- properties: [Word.Interfaces.ShadingUpdateData](/en-us/javascript/api/word/word.interfaces.shadingupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns  
void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Shading): void;
```

Parameters  
- properties: [Word.Shading](/en-us/javascript/api/word/word.shading)

Returns  
void

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Shading object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ShadingData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ShadingData;
```

Returns  
[Word.Interfaces.ShadingData](/en-us/javascript/api/word/word.interfaces.shadingdata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Shading;
```

Returns  
[Word.Shading](/en-us/javascript/api/word/word.shading)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Shading;
```

Returns  
[Word.Shading](/en-us/javascript/api/word/word.shading)