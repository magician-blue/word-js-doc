# Word.BorderCollection class

Package: [word](/en-us/javascript/api/word)

Represents the collection of border styles.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApiDesktop 1.1]

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

- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- insideBorderColor  
  Specifies the 24-bit color of the inside borders. Color is specified in '#RRGGBB' format or by using the color name.
- insideBorderType  
  Specifies the border type of the inside borders.
- insideBorderWidth  
  Specifies the width of the inside borders.
- items  
  Gets the loaded child items in this collection.
- outsideBorderColor  
  Specifies the 24-bit color of the outside borders. Color is specified in '#RRGGBB' format or by using the color name.
- outsideBorderType  
  Specifies the border type of the outside borders.
- outsideBorderWidth  
  Specifies the width of the outside borders.

## Methods

- getByLocation(borderLocation)  
  Gets the border that has the specified location.
- getFirst()  
  Gets the first border in this collection. Throws an ItemNotFound error if this collection is empty.
- getFirstOrNullObject()  
  Gets the first border in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- getItem(index)  
  Gets a Border object by its index in the collection.
- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BorderCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BorderCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### insideBorderColor

Specifies the 24-bit color of the inside borders. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
insideBorderColor: string;
```

Property Value: string

Remarks  
[API set: WordApiDesktop 1.1]

---

### insideBorderType

Specifies the border type of the inside borders.

```typescript
insideBorderType: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
```

Property Value: [Word.BorderType](/en-us/javascript/api/word/word.bordertype) | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"

Remarks  
[API set: WordApiDesktop 1.1]

---

### insideBorderWidth

Specifies the width of the inside borders.

```typescript
insideBorderWidth: Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed";
```

Property Value: [Word.BorderWidth](/en-us/javascript/api/word/word.borderwidth) | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"

Remarks  
[API set: WordApiDesktop 1.1]

---

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Border[];
```

Property Value: [Word.Border](/en-us/javascript/api/word/word.border)[]

---

### outsideBorderColor

Specifies the 24-bit color of the outside borders. Color is specified in '#RRGGBB' format or by using the color name.

```typescript
outsideBorderColor: string;
```

Property Value: string

Remarks  
[API set: WordApiDesktop 1.1]

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

---

### outsideBorderType

Specifies the border type of the outside borders.

```typescript
outsideBorderType: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
```

Property Value: [Word.BorderType](/en-us/javascript/api/word/word.bordertype) | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"

Remarks  
[API set: WordApiDesktop 1.1]

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

---

### outsideBorderWidth

Specifies the width of the outside borders.

```typescript
outsideBorderWidth: Word.BorderWidth | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed";
```

Property Value: [Word.BorderWidth](/en-us/javascript/api/word/word.borderwidth) | "None" | "Pt025" | "Pt050" | "Pt075" | "Pt100" | "Pt150" | "Pt225" | "Pt300" | "Pt450" | "Pt600" | "Mixed"

Remarks  
[API set: WordApiDesktop 1.1]

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

## Method Details

### getByLocation(borderLocation)

Gets the border that has the specified location.

```typescript
getByLocation(
  borderLocation:
    Word.BorderLocation.top |
    Word.BorderLocation.left |
    Word.BorderLocation.bottom |
    Word.BorderLocation.right |
    Word.BorderLocation.insideHorizontal |
    Word.BorderLocation.insideVertical |
    "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical"
): Word.Border;
```

Parameters  
- borderLocation: [top](/en-us/javascript/api/word/word.borderlocation#word-word-borderlocation-top-member) | [left](/en-us/javascript/api/word/word.borderlocation#word-word-borderlocation-left-member) | [bottom](/en-us/javascript/api/word/word.borderlocation#word-word-borderlocation-bottom-member) | [right](/en-us/javascript/api/word/word.borderlocation#word-word-borderlocation-right-member) | [insideHorizontal](/en-us/javascript/api/word/word.borderlocation#word-word-borderlocation-insidehorizontal-member) | [insideVertical](/en-us/javascript/api/word/word.borderlocation#word-word-borderlocation-insidevertical-member) | "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical"

Returns: [Word.Border](/en-us/javascript/api/word/word.border)

Remarks  
[API set: WordApiDesktop 1.1]

---

### getFirst()

Gets the first border in this collection. Throws an ItemNotFound error if this collection is empty.

```typescript
getFirst(): Word.Border;
```

Returns: [Word.Border](/en-us/javascript/api/word/word.border)

Remarks  
[API set: WordApiDesktop 1.1]

---

### getFirstOrNullObject()

Gets the first border in this collection. If this collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Border;
```

Returns: [Word.Border](/en-us/javascript/api/word/word.border)

Remarks  
[API set: WordApiDesktop 1.1]

---

### getItem(index)

Gets a Border object by its index in the collection.

```typescript
getItem(index: number): Word.Border;
```

Parameters  
- index: number

A number that identifies the index location of a Border object.

Returns: [Word.Border](/en-us/javascript/api/word/word.border)

Remarks  
[API set: WordApiDesktop 1.1]

---

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.BorderCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.BorderCollection;
```

Parameters  
- options: [Word.Interfaces.BorderCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.bordercollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.BorderCollection](/en-us/javascript/api/word/word.bordercollection)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.BorderCollection;
```

Parameters  
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.BorderCollection](/en-us/javascript/api/word/word.bordercollection)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.BorderCollection;
```

Parameters  
- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.BorderCollection](/en-us/javascript/api/word/word.bordercollection)

---

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.BorderCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.BorderCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.BorderCollectionData;
```

Returns: [Word.Interfaces.BorderCollectionData](/en-us/javascript/api/word/word.interfaces.bordercollectiondata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.BorderCollection;
```

Returns: [Word.BorderCollection](/en-us/javascript/api/word/word.bordercollection)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.BorderCollection;
```

Returns: [Word.BorderCollection](/en-us/javascript/api/word/word.bordercollection)