# Word.TableCell class

Package: [word](/en-us/javascript/api/word)

Represents a table cell in a Word document.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets content alignment details about the first cell of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  const firstCell: Word.TableCell = firstTableRow.cells.getFirst();
  firstCell.load(["horizontalAlignment", "verticalAlignment"]);
  await context.sync();

  console.log(
    `Details about the alignment of the first table's first cell:`,
    `- Horizontal alignment of the cell's content: ${firstCell.horizontalAlignment}`,
    `- Vertical alignment of the cell's content: ${firstCell.verticalAlignment}`
  );
});
```

## Properties
- `body` — Gets the body object of the cell.
- `cellIndex` — Gets the index of the cell in its row.
- `columnWidth` — Specifies the width of the cell's column in points. This is applicable to uniform tables.
- `context` — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- `horizontalAlignment` — Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- `parentRow` — Gets the parent row of the cell.
- `parentTable` — Gets the parent table of the cell.
- `rowIndex` — Gets the index of the cell's row in the table.
- `shadingColor` — Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.
- `value` — Specifies the text of the cell.
- `verticalAlignment` — Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.
- `width` — Gets the width of the cell in points.

## Methods
- `deleteColumn()` — Deletes the column containing this cell. This is applicable to uniform tables.
- `deleteRow()` — Deletes the row containing this cell.
- `getBorder(borderLocation)` — Gets the border style for the specified border.
- `getBorder(borderLocation)` — Gets the border style for the specified border.
- `getCellPadding(cellPaddingLocation)` — Gets cell padding in points.
- `getCellPadding(cellPaddingLocation)` — Gets cell padding in points.
- `getNext()` — Gets the next cell. Throws an `ItemNotFound` error if this cell is the last one.
- `getNextOrNullObject()` — Gets the next cell. If this cell is the last one, returns an object with `isNullObject` set to `true`. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- `insertColumns(insertLocation, columnCount, values)` — Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
- `insertRows(insertLocation, rowCount, values)` — Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.
- `load(options)` — Queues up a command to load the specified properties of the object. Call `context.sync()` before reading the properties.
- `load(propertyNames)` — Queues up a command to load the specified properties of the object. Call `context.sync()` before reading the properties.
- `load(propertyNamesAndPaths)` — Queues up a command to load the specified properties of the object. Call `context.sync()` before reading the properties.
- `set(properties, options)` — Sets multiple properties of an object at the same time. Accepts either a plain object with the appropriate properties or another API object of the same type.
- `set(properties)` — Sets multiple properties on the object at the same time, based on an existing loaded object.
- `setCellPadding(cellPaddingLocation, cellPadding)` — Sets cell padding in points.
- `setCellPadding(cellPaddingLocation, cellPadding)` — Sets cell padding in points.
- `split(rowCount, columnCount)` — Splits the cell into the specified number of rows and columns.
- `toJSON()` — Overrides the JavaScript `toJSON()` method to provide more useful output when an API object is passed to `JSON.stringify()`. Returns a plain JavaScript object (typed as `Word.Interfaces.TableCellData`) that contains shallow copies of any loaded child properties from the original object.
- `track()` — Track the object for automatic adjustment based on surrounding changes in the document. Shorthand for `context.trackedObjects.add(thisObject)`.
- `untrack()` — Release the memory associated with this object, if it has previously been tracked. Shorthand for `context.trackedObjects.remove(thisObject)`. Call `context.sync()` for the release to take effect.

## Property Details

### body
Gets the body object of the cell.

```typescript
readonly body: Word.Body;
```

- Property Value: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks
- [API set: WordApi 1.3]

### cellIndex
Gets the index of the cell in its row.

```typescript
readonly cellIndex: number;
```

- Property Value: number

Remarks
- [API set: WordApi 1.3]

### columnWidth
Specifies the width of the cell's column in points. This is applicable to uniform tables.

```typescript
columnWidth: number;
```

- Property Value: number

Remarks
- [API set: WordApi 1.3]

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### horizontalAlignment
Specifies the horizontal alignment of the cell. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

- Property Value: [Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks
- [API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets content alignment details about the first cell of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  const firstCell: Word.TableCell = firstTableRow.cells.getFirst();
  firstCell.load(["horizontalAlignment", "verticalAlignment"]);
  await context.sync();

  console.log(
    `Details about the alignment of the first table's first cell:`,
    `- Horizontal alignment of the cell's content: ${firstCell.horizontalAlignment}`,
    `- Vertical alignment of the cell's content: ${firstCell.verticalAlignment}`
  );
});
```

### parentRow
Gets the parent row of the cell.

```typescript
readonly parentRow: Word.TableRow;
```

- Property Value: [Word.TableRow](/en-us/javascript/api/word/word.tablerow)

Remarks
- [API set: WordApi 1.3]

### parentTable
Gets the parent table of the cell.

```typescript
readonly parentTable: Word.Table;
```

- Property Value: [Word.Table](/en-us/javascript/api/word/word.table)

Remarks
- [API set: WordApi 1.3]

### rowIndex
Gets the index of the cell's row in the table.

```typescript
readonly rowIndex: number;
```

- Property Value: number

Remarks
- [API set: WordApi 1.3]

### shadingColor
Specifies the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor: string;
```

- Property Value: string

Remarks
- [API set: WordApi 1.3]

### value
Specifies the text of the cell.

```typescript
value: string;
```

- Property Value: string

Remarks
- [API set: WordApi 1.3]

### verticalAlignment
Specifies the vertical alignment of the cell. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
```

- Property Value: [Word.VerticalAlignment](/en-us/javascript/api/word/word.verticalalignment) | "Mixed" | "Top" | "Center" | "Bottom"

Remarks
- [API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets content alignment details about the first cell of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  const firstCell: Word.TableCell = firstTableRow.cells.getFirst();
  firstCell.load(["horizontalAlignment", "verticalAlignment"]);
  await context.sync();

  console.log(
    `Details about the alignment of the first table's first cell:`,
    `- Horizontal alignment of the cell's content: ${firstCell.horizontalAlignment}`,
    `- Vertical alignment of the cell's content: ${firstCell.verticalAlignment}`
  );
});
```

### width
Gets the width of the cell in points.

```typescript
readonly width: number;
```

- Property Value: number

Remarks
- [API set: WordApi 1.3]

## Method Details

### deleteColumn()
Deletes the column containing this cell. This is applicable to uniform tables.

```typescript
deleteColumn(): void;
```

Returns
- void

Remarks
- [API set: WordApi 1.3]

### deleteRow()
Deletes the row containing this cell.

```typescript
deleteRow(): void;
```

Returns
- void

Remarks
- [API set: WordApi 1.3]

### getBorder(borderLocation)
Gets the border style for the specified border.

```typescript
getBorder(borderLocation: Word.BorderLocation): Word.TableBorder;
```

Parameters
- borderLocation: [Word.BorderLocation](/en-us/javascript/api/word/word.borderlocation)  
  Required. The border location.

Returns
- [Word.TableBorder](/en-us/javascript/api/word/word.tableborder)

Remarks
- [API set: WordApi 1.3]

### getBorder(borderLocation)
Gets the border style for the specified border.

```typescript
getBorder(borderLocation: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"): Word.TableBorder;
```

Parameters
- borderLocation: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"  
  Required. The border location.

Returns
- [Word.TableBorder](/en-us/javascript/api/word/word.tableborder)

Remarks
- [API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstCell: Word.TableCell = firstTable.getCell(0, 0);
  const borderLocation = "Left";
  const border: Word.TableBorder = firstCell.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(
    `Details about the ${borderLocation} border of the first table's first cell:`,
    `- Color: ${border.color}`,
    `- Type: ${border.type}`,
    `- Width: ${border.width} points`
  );
});
```

### getCellPadding(cellPaddingLocation)
Gets cell padding in points.

```typescript
getCellPadding(cellPaddingLocation: Word.CellPaddingLocation): OfficeExtension.ClientResult<number>;
```

Parameters
- cellPaddingLocation: [Word.CellPaddingLocation](/en-us/javascript/api/word/word.cellpaddinglocation)  
  Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks
- [API set: WordApi 1.3]

### getCellPadding(cellPaddingLocation)
Gets cell padding in points.

```typescript
getCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right"): OfficeExtension.ClientResult<number>;
```

Parameters
- cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right"  
  Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

Remarks
- [API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets cell padding details about the first cell of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstCell: Word.TableCell = firstTable.getCell(0, 0);
  const cellPaddingLocation = "Left";
  const cellPadding = firstCell.getCellPadding(cellPaddingLocation);
  await context.sync();

  console.log(
    `Cell padding details about the ${cellPaddingLocation} border of the first table's first cell: ${cellPadding.value} points`
  );
});
```

### getNext()
Gets the next cell. Throws an `ItemNotFound` error if this cell is the last one.

```typescript
getNext(): Word.TableCell;
```

Returns
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Remarks
- [API set: WordApi 1.3]

### getNextOrNullObject()
Gets the next cell. If this cell is the last one, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getNextOrNullObject(): Word.TableCell;
```

Returns
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Remarks
- [API set: WordApi 1.3]

### insertColumns(insertLocation, columnCount, values)
Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.

```typescript
insertColumns(
  insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After",
  columnCount: number,
  values?: string[][]
): void;
```

Parameters
- insertLocation: Required. It must be 'Before' or 'After'.
- columnCount: number  
  Required. Number of columns to add.
- values: string[][]  
  Optional 2D array. Cells are filled if the corresponding strings are specified in the array.

Returns
- void

Remarks
- [API set: WordApi 1.3]

### insertRows(insertLocation, rowCount, values)
Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows.

```typescript
insertRows(
  insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After",
  rowCount: number,
  values?: string[][]
): Word.TableRowCollection;
```

Parameters
- insertLocation: Required. It must be 'Before' or 'After'.
- rowCount: number  
  Required. Number of rows to add.
- values: string[][]  
  Optional 2D array. Cells are filled if the corresponding strings are specified in the array.

Returns
- [Word.TableRowCollection](/en-us/javascript/api/word/word.tablerowcollection)

Remarks
- [API set: WordApi 1.3]

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.TableCellLoadOptions): Word.TableCell;
```

Parameters
- options: [Word.Interfaces.TableCellLoadOptions](/en-us/javascript/api/word/word.interfaces.tablecellloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TableCell;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.TableCell;
```

Parameters
- propertyNamesAndPaths:  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.TableCellUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.TableCellUpdateData](/en-us/javascript/api/word/word.interfaces.tablecellupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.TableCell): void;
```

Parameters
- properties: [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Returns
- void

### setCellPadding(cellPaddingLocation, cellPadding)
Sets cell padding in points.

```typescript
setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number): void;
```

Parameters
- cellPaddingLocation: [Word.CellPaddingLocation](/en-us/javascript/api/word/word.cellpaddinglocation)  
  Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
- cellPadding: number  
  Required. The cell padding.

Returns
- void

Remarks
- [API set: WordApi 1.3]

### setCellPadding(cellPaddingLocation, cellPadding)
Sets cell padding in points.

```typescript
setCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right", cellPadding: number): void;
```

Parameters
- cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right"  
  Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
- cellPadding: number  
  Required. The cell padding.

Returns
- void

Remarks
- [API set: WordApi 1.3]

### split(rowCount, columnCount)
Splits the cell into the specified number of rows and columns.

```typescript
split(rowCount: number, columnCount: number): void;
```

Parameters
- rowCount: number  
  Required. The number of rows to split into. Must be a divisor of the number of underlying rows.
- columnCount: number  
  Required. The number of columns to split into.

Returns
- void

Remarks
- [API set: WordApi 1.4]

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.TableCell` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.TableCellData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.TableCellData;
```

Returns
- [Word.Interfaces.TableCellData](/en-us/javascript/api/word/word.interfaces.tablecelldata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TableCell;
```

Returns
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.TableCell;
```

Returns
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)