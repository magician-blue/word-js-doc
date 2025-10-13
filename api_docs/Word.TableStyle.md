# Word.TableStyle class

Package: [word](/en-us/javascript/api/word)

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

Represents the TableStyle object.

## Remarks

[API set: WordApi 1.6]

### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-custom-style.yaml

// Gets the table style properties and displays them in the form.
const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
if (styleName == "") {
  console.warn("Please input a table style name.");
  return;
}

await Word.run(async (context) => {
  const tableStyle: Word.TableStyle =
    context.document.getStyles().getByName(styleName).tableStyle;
  tableStyle.load();
  await context.sync();

  if (tableStyle.isNullObject) {
    console.warn(`There's no existing table style with the name '${styleName}'.`);
    return;
  }

  console.log(tableStyle);
  (document.getElementById("alignment") as HTMLInputElement).value = tableStyle.alignment;
  (document.getElementById("allow-break-across-page") as HTMLInputElement).value =
    tableStyle.allowBreakAcrossPage.toString();
  (document.getElementById("top-cell-margin") as HTMLInputElement).value = tableStyle.topCellMargin;
  (document.getElementById("bottom-cell-margin") as HTMLInputElement).value = tableStyle.bottomCellMargin;
  (document.getElementById("left-cell-margin") as HTMLInputElement).value = tableStyle.leftCellMargin;
  (document.getElementById("right-cell-margin") as HTMLInputElement).value = tableStyle.rightCellMargin;
  (document.getElementById("cell-spacing") as HTMLInputElement).value = tableStyle.cellSpacing;
});
```

## Properties

- [alignment](#alignment): Specifies the table's alignment against the page margin.
- [allowBreakAcrossPage](#allowbreakacrosspage): Specifies whether lines in tables formatted with a specified style break across pages.
- [bottomCellMargin](#bottomcellmargin): Specifies the amount of space to add between the contents and the bottom borders of the cells.
- [cellSpacing](#cellspacing): Specifies the spacing (in points) between the cells in a table style.
- [context](#context): The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [leftCellMargin](#leftcellmargin): Specifies the amount of space to add between the contents and the left borders of the cells.
- [rightCellMargin](#rightcellmargin): Specifies the amount of space to add between the contents and the right borders of the cells.
- [topCellMargin](#topcellmargin): Specifies the amount of space to add between the contents and the top borders of the cells.

## Methods

- [load(options)](#loadoptions): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNames)](#loadpropertynames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [load(propertyNamesAndPaths)](#loadpropertynamesandpaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- [set(properties, options)](#setproperties-options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#setproperties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- [toJSON()](#tojson): Overrides the JavaScript toJSON() method to provide more useful output for JSON.stringify().
- [track()](#track): Track the object for automatic adjustment based on surrounding changes in the document.
- [untrack()](#untrack): Release the memory associated with this object, if it has previously been tracked.

## Property Details

### alignment

Specifies the table's alignment against the page margin.

```typescript
alignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

Property Value: [Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks: [API set: WordApiDesktop 1.1]

### allowBreakAcrossPage

Specifies whether lines in tables formatted with a specified style break across pages.

```typescript
allowBreakAcrossPage: boolean;
```

Property Value: boolean

Remarks: [API set: WordApiDesktop 1.1]

### bottomCellMargin

Specifies the amount of space to add between the contents and the bottom borders of the cells.

```typescript
bottomCellMargin: number;
```

Property Value: number

Remarks: [API set: WordApi 1.6]

### cellSpacing

Specifies the spacing (in points) between the cells in a table style.

```typescript
cellSpacing: number;
```

Property Value: number

Remarks: [API set: WordApi 1.6]

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### leftCellMargin

Specifies the amount of space to add between the contents and the left borders of the cells.

```typescript
leftCellMargin: number;
```

Property Value: number

Remarks: [API set: WordApi 1.6]

### rightCellMargin

Specifies the amount of space to add between the contents and the right borders of the cells.

```typescript
rightCellMargin: number;
```

Property Value: number

Remarks: [API set: WordApi 1.6]

### topCellMargin

Specifies the amount of space to add between the contents and the top borders of the cells.

```typescript
topCellMargin: number;
```

Property Value: number

Remarks: [API set: WordApi 1.6]

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.TableStyleLoadOptions): Word.TableStyle;
```

Parameters:
- options: [Word.Interfaces.TableStyleLoadOptions](/en-us/javascript/api/word/word.interfaces.tablestyleloadoptions)  
  Provides options for which properties of the object to load.

Returns: [Word.TableStyle](/en-us/javascript/api/word/word.tablestyle)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TableStyle;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.TableStyle](/en-us/javascript/api/word/word.tablestyle)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.TableStyle;
```

Parameters:
- propertyNamesAndPaths:  
  - select?: string  
  - expand?: string  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.TableStyle](/en-us/javascript/api/word/word.tablestyle)

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.TableStyleUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: [Word.Interfaces.TableStyleUpdateData](/en-us/javascript/api/word/word.interfaces.tablestyleupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.TableStyle): void;
```

Parameters:
- properties: [Word.TableStyle](/en-us/javascript/api/word/word.tablestyle)

Returns: void

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableStyle object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableStyleData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.TableStyleData;
```

Returns: [Word.Interfaces.TableStyleData](/en-us/javascript/api/word/word.interfaces.tablestyledata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TableStyle;
```

Returns: [Word.TableStyle](/en-us/javascript/api/word/word.tablestyle)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.TableStyle;
```

Returns: [Word.TableStyle](/en-us/javascript/api/word/word.tablestyle)