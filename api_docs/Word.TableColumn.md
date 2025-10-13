# Word.TableColumn class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a table column in a Word document.

Extends
OfficeExtension.ClientObject: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

## Properties

- borders
  - Returns a BorderUniversalCollection object that represents all the borders for the table column.
- columnIndex
  - Returns the position of this column in a collection.
- context
  - The request context associated with the object. This connects the add-in's process to the Office host application's process.
- isFirst
  - Returns true if the column or row is the first one in the table; false otherwise.
- isLast
  - Returns true if the column or row is the last one in the table; false otherwise.
- nestingLevel
  - Returns the nesting level of the column.
- preferredWidth
  - Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the preferredWidthType property.
- preferredWidthType
  - Specifies the preferred unit of measurement to use for the width of the table column.
- shading
  - Returns a ShadingUniversal object that refers to the shading formatting for the column.
- width
  - Specifies the width of the column, in points.

## Methods

- autoFit()
  - Changes the width of the table column to accommodate the width of the text without changing the way text wraps in the cells.
- delete()
  - Deletes the column.
- load(options)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)
  - Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- select()
  - Selects the table column.
- set(properties, options)
  - Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties)
  - Sets multiple properties on the object at the same time, based on an existing loaded object.
- setWidth(columnWidth, rulerStyle)
  - Sets the width of the column in a table.
- setWidth(columnWidth, rulerStyle)
  - Sets the width of the column in a table.
- sort()
  - Sorts the table column.
- toJSON()
  - Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). JSON.stringify, in turn, calls the toJSON method of the object that's passed to it. Whereas the original Word.TableColumn object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableColumnData) that contains shallow copies of any loaded child properties from the original object.
- track()
  - Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()
  - Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### borders

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders for the table column.

```typescript
readonly borders: Word.BorderUniversalCollection;
```

Property Value
Word.BorderUniversalCollection: https://learn.microsoft.com/en-us/javascript/api/word/word.borderuniversalcollection

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### columnIndex

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the position of this column in a collection.

```typescript
readonly columnIndex: number;
```

Property Value
number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### context

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
Word.RequestContext: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### isFirst

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if the column or row is the first one in the table; false otherwise.

```typescript
readonly isFirst: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isLast

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns true if the column or row is the last one in the table; false otherwise.

```typescript
readonly isLast: boolean;
```

Property Value
boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### nestingLevel

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the nesting level of the column.

```typescript
readonly nestingLevel: number;
```

Property Value
number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### preferredWidth

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the preferred width (in points or as a percentage of the window width) for the column. The unit of measurement can be specified by the preferredWidthType property.

```typescript
preferredWidth: number;
```

Property Value
number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### preferredWidthType

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the preferred unit of measurement to use for the width of the table column.

```typescript
preferredWidthType: Word.PreferredWidthType | "Auto" | "Percent" | "Points";
```

Property Value
Word.PreferredWidthType | "Auto" | "Percent" | "Points": https://learn.microsoft.com/en-us/javascript/api/word/word.preferredwidthtype

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### shading

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the column.

```typescript
readonly shading: Word.ShadingUniversal;
```

Property Value
Word.ShadingUniversal: https://learn.microsoft.com/en-us/javascript/api/word/word.shadinguniversal

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### width

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the column, in points.

```typescript
width: number;
```

Property Value
number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

## Method Details

### autoFit()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Changes the width of the table column to accommodate the width of the text without changing the way text wraps in the cells.

```typescript
autoFit(): void;
```

Returns
void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### delete()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Deletes the column.

```typescript
delete(): void;
```

Returns
void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### load(options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.TableColumnLoadOptions): Word.TableColumn;
```

Parameters
- options: Word.Interfaces.TableColumnLoadOptions
  - Provides options for which properties of the object to load.

Returns
Word.TableColumn: https://learn.microsoft.com/en-us/javascript/api/word/word.tablecolumn

### load(propertyNames)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TableColumn;
```

Parameters
- propertyNames: string | string[]
  - A comma-delimited string or an array of strings that specify the properties to load.

Returns
Word.TableColumn: https://learn.microsoft.com/en-us/javascript/api/word/word.tablecolumn

### load(propertyNamesAndPaths)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.TableColumn;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }
  - propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
Word.TableColumn: https://learn.microsoft.com/en-us/javascript/api/word/word.tablecolumn

### select()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Selects the table column.

```typescript
select(): void;
```

Returns
void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### set(properties, options)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.TableColumnUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: Word.Interfaces.TableColumnUpdateData
  - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: OfficeExtension.UpdateOptions
  - Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
void

### set(properties)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.TableColumn): void;
```

Parameters
- properties: Word.TableColumn

Returns
void

### setWidth(columnWidth, rulerStyle)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the width of the column in a table.

```typescript
setWidth(columnWidth: number, rulerStyle: Word.RulerStyle): void;
```

Parameters
- columnWidth: number
  - The width to set.
- rulerStyle: Word.RulerStyle
  - The ruler style to apply.

Returns
void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### setWidth(columnWidth, rulerStyle)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the width of the column in a table.

```typescript
setWidth(columnWidth: number, rulerStyle: "None" | "Proportional" | "FirstColumn" | "SameWidth"): void;
```

Parameters
- columnWidth: number
  - The width to set.
- rulerStyle: "None" | "Proportional" | "FirstColumn" | "SameWidth"
  - The ruler style to apply.

Returns
void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### sort()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sorts the table column.

```typescript
sort(): void;
```

Returns
void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### toJSON()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableColumn object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableColumnData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.TableColumnData;
```

Returns
Word.Interfaces.TableColumnData: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.tablecolumndata

### track()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TableColumn;
```

Returns
Word.TableColumn: https://learn.microsoft.com/en-us/javascript/api/word/word.tablecolumn

### untrack()

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.TableColumn;
```

Returns
Word.TableColumn: https://learn.microsoft.com/en-us/javascript/api/word/word.tablecolumn