# Word.TableBorder class

- Package: https://learn.microsoft.com/en-us/javascript/api/word
- Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

Specifies the border style.

## Remarks

[API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
});
```

## Properties

- color: Specifies the table border color.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- type: Specifies the type of the table border.
- width: Specifies the width, in points, of the table border. Not applicable to table border types that have fixed widths.

## Methods

- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableBorder object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableBorderData) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### color

Specifies the table border color.

```typescript
color: string;
```

Property Value
- string

Remarks
- [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
});
```

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- Word.RequestContext: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### type

Specifies the type of the table border.

```typescript
type: Word.BorderType | "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave";
```

Property Value
- Word.BorderType: https://learn.microsoft.com/en-us/javascript/api/word/word.bordertype
- or one of: "Mixed" | "None" | "Single" | "Double" | "Dotted" | "Dashed" | "DotDashed" | "Dot2Dashed" | "Triple" | "ThinThickSmall" | "ThickThinSmall" | "ThinThickThinSmall" | "ThinThickMed" | "ThickThinMed" | "ThinThickThinMed" | "ThinThickLarge" | "ThickThinLarge" | "ThinThickThinLarge" | "Wave" | "DoubleWave" | "DashedSmall" | "DashDotStroked" | "ThreeDEmboss" | "ThreeDEngrave"

Remarks
- [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
});
```

### width

Specifies the width, in points, of the table border. Not applicable to table border types that have fixed widths.

```typescript
width: number;
```

Property Value
- number

Remarks
- [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const borderLocation = Word.BorderLocation.top;
  const border: Word.TableBorder = firstTable.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(`Details about the ${borderLocation} border of the first table:`, `- Color: ${border.color}`, `- Type: ${border.type}`, `- Width: ${border.width} points`);
});
```

## Method Details

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.TableBorderLoadOptions): Word.TableBorder;
```

Parameters
- options: Word.Interfaces.TableBorderLoadOptions — Provides options for which properties of the object to load.  
  https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tableborderloadoptions

Returns
- Word.TableBorder: https://learn.microsoft.com/en-us/javascript/api/word/word.tableborder

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TableBorder;
```

Parameters
- propertyNames: string | string[] — A comma-delimited string or an array of strings that specify the properties to load.

Returns
- Word.TableBorder: https://learn.microsoft.com/en-us/javascript/api/word/word.tableborder

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.TableBorder;
```

Parameters
- propertyNamesAndPaths:  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- Word.TableBorder: https://learn.microsoft.com/en-us/javascript/api/word/word.tableborder

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.TableBorderUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: Word.Interfaces.TableBorderUpdateData — A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.  
  https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tableborderupdatedata
- options: OfficeExtension.UpdateOptions — Provides an option to suppress errors if the properties object tries to set any read-only properties.  
  https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions

Returns
- void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.TableBorder): void;
```

Parameters
- properties: Word.TableBorder — https://learn.microsoft.com/en-us/javascript/api/word/word.tableborder

Returns
- void

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableBorder object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableBorderData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.TableBorderData;
```

Returns
- Word.Interfaces.TableBorderData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tableborderdata

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TableBorder;
```

Returns
- Word.TableBorder: https://learn.microsoft.com/en-us/javascript/api/word/word.tableborder

Reference for tracked objects:
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.TableBorder;
```

Returns
- Word.TableBorder: https://learn.microsoft.com/en-us/javascript/api/word/word.tableborder

Reference for tracked objects:
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member