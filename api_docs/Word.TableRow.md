# Word.TableRow class

Package: [word](/en-us/javascript/api/word)

Represents a row in a Word document.

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
- cellCount  
  Gets the number of cells in the row.
- cells  
  Gets cells.
- context  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- endnotes  
  Gets the collection of endnotes in the table row.
- fields  
  Gets the collection of field objects in the table row.
- font  
  Gets the font. Use this to get and set font name, size, color, and other properties.
- footnotes  
  Gets the collection of footnotes in the table row.
- horizontalAlignment  
  Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- isHeader  
  Checks whether the row is a header row. To set the number of header rows, use headerRowCount on the Table object.
- parentTable  
  Gets parent table.
- preferredHeight  
  Specifies the preferred height of the row in points.
- rowIndex  
  Gets the index of the row in its parent table.
- shadingColor  
  Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.
- values  
  Specifies the text values in the row, as a 2D JavaScript array.
- verticalAlignment  
  Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

## Methods
- clear()  
  Clears the contents of the row.
- delete()  
  Deletes the entire row.
- getBorder(borderLocation)  
  Gets the border style of the cells in the row.
- getBorder(borderLocation)  
  Gets the border style of the cells in the row.
- getCellPadding(cellPaddingLocation)  
  Gets cell padding in points.
- getCellPadding(cellPaddingLocation)  
  Gets cell padding in points.
- getNext()  
  Gets the next row. Throws an ItemNotFound error if this row is the last one.
- getNextOrNullObject()  
  Gets the next row. If this row is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- insertContentControl()  
  Inserts a content control on the row.
- insertRows(insertLocation, rowCount, values)  
  Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.
- load(options)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths)  
  Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- merge()  
  Merges the row into one cell.
- search(searchText, searchOptions)  
  Performs a search with the specified SearchOptions on the scope of the row. The search results are a collection of range objects.
- select(selectionMode)  
  Selects the row and navigates the Word UI to it.
- select(selectionMode)  
  Selects the row and navigates the Word UI to it.
- set(properties, options)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- setCellPadding(cellPaddingLocation, cellPadding)  
  Sets cell padding in points.
- setCellPadding(cellPaddingLocation, cellPadding)  
  Sets cell padding in points.
- toJSON()  
  Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableRow object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableRowData) that contains shallow copies of any loaded child properties from the original object.
- track()  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack()  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### cellCount
Gets the number of cells in the row.

```typescript
readonly cellCount: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi 1.3]

---

### cells
Gets cells.

```typescript
readonly cells: Word.TableCellCollection;
```

#### Property Value
[Word.TableCellCollection](/en-us/javascript/api/word/word.tablecellcollection)

#### Remarks
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

---

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

#### Property Value
[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### endnotes
Gets the collection of endnotes in the table row.

```typescript
readonly endnotes: Word.NoteItemCollection;
```

#### Property Value
[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

#### Remarks
[API set: WordApi 1.5]

---

### fields
Gets the collection of field objects in the table row.

```typescript
readonly fields: Word.FieldCollection;
```

#### Property Value
[Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

#### Remarks
[API set: WordApi 1.4]

---

### font
Gets the font. Use this to get and set font name, size, color, and other properties.

```typescript
readonly font: Word.Font;
```

#### Property Value
[Word.Font](/en-us/javascript/api/word/word.font)

#### Remarks
[API set: WordApi 1.3]

---

### footnotes
Gets the collection of footnotes in the table row.

```typescript
readonly footnotes: Word.NoteItemCollection;
```

#### Property Value
[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

#### Remarks
[API set: WordApi 1.5]

---

### horizontalAlignment
Specifies the horizontal alignment of every cell in the row. The value can be 'Left', 'Centered', 'Right', or 'Justified'.

```typescript
horizontalAlignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

#### Property Value
[Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

#### Remarks
[API set: WordApi 1.3]

---

### isHeader
Checks whether the row is a header row. To set the number of header rows, use headerRowCount on the Table object.

```typescript
readonly isHeader: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi 1.3]

---

### parentTable
Gets parent table.

```typescript
readonly parentTable: Word.Table;
```

#### Property Value
[Word.Table](/en-us/javascript/api/word/word.table)

#### Remarks
[API set: WordApi 1.3]

---

### preferredHeight
Specifies the preferred height of the row in points.

```typescript
preferredHeight: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi 1.3]

---

### rowIndex
Gets the index of the row in its parent table.

```typescript
readonly rowIndex: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi 1.3]

---

### shadingColor
Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.

```typescript
shadingColor: string;
```

#### Property Value
string

#### Remarks
[API set: WordApi 1.3]

---

### values
Specifies the text values in the row, as a 2D JavaScript array.

```typescript
values: string[][];
```

#### Property Value
string[][]

#### Remarks
[API set: WordApi 1.3]

---

### verticalAlignment
Specifies the vertical alignment of the cells in the row. The value can be 'Top', 'Center', or 'Bottom'.

```typescript
verticalAlignment: Word.VerticalAlignment | "Mixed" | "Top" | "Center" | "Bottom";
```

#### Property Value
[Word.VerticalAlignment](/en-us/javascript/api/word/word.verticalalignment) | "Mixed" | "Top" | "Center" | "Bottom"

#### Remarks
[API set: WordApi 1.3]

## Method Details

### clear()
Clears the contents of the row.

```typescript
clear(): void;
```

#### Returns
void

#### Remarks
[API set: WordApi 1.3]

---

### delete()
Deletes the entire row.

```typescript
delete(): void;
```

#### Returns
void

#### Remarks
[API set: WordApi 1.3]

---

### getBorder(borderLocation)
Gets the border style of the cells in the row.

```typescript
getBorder(borderLocation: Word.BorderLocation): Word.TableBorder;
```

#### Parameters
- borderLocation  
  [Word.BorderLocation](/en-us/javascript/api/word/word.borderlocation)  
  Required. The border location.

#### Returns
[Word.TableBorder](/en-us/javascript/api/word/word.tableborder)

#### Remarks
[API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets border details about the first row of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  const borderLocation = Word.BorderLocation.bottom;
  const border: Word.TableBorder = firstTableRow.getBorder(borderLocation);
  border.load(["type", "color", "width"]);
  await context.sync();

  console.log(
    `Details about the ${borderLocation} border of the first table's first row:`,
    `- Color: ${border.color}`,
    `- Type: ${border.type}`,
    `- Width: ${border.width} points`
  );
});
```

---

### getBorder(borderLocation)
Gets the border style of the cells in the row.

```typescript
getBorder(borderLocation: "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"): Word.TableBorder;
```

#### Parameters
- borderLocation  
  "Top" | "Left" | "Bottom" | "Right" | "InsideHorizontal" | "InsideVertical" | "Inside" | "Outside" | "All"  
  Required. The border location.

#### Returns
[Word.TableBorder](/en-us/javascript/api/word/word.tableborder)

#### Remarks
[API set: WordApi 1.3]

---

### getCellPadding(cellPaddingLocation)
Gets cell padding in points.

```typescript
getCellPadding(cellPaddingLocation: Word.CellPaddingLocation): OfficeExtension.ClientResult<number>;
```

#### Parameters
- cellPaddingLocation  
  [Word.CellPaddingLocation](/en-us/javascript/api/word/word.cellpaddinglocation)  
  Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.

#### Returns
[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

#### Remarks
[API set: WordApi 1.3]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-formatting.yaml

// Gets cell padding details about the first row of the first table in the document.
await Word.run(async (context) => {
  const firstTable: Word.Table = context.document.body.tables.getFirst();
  const firstTableRow: Word.TableRow = firstTable.rows.getFirst();
  const cellPaddingLocation = Word.CellPaddingLocation.bottom;
  const cellPadding = firstTableRow.getCellPadding(cellPaddingLocation);
  await context.sync();

  console.log(
    `Cell padding details about the ${cellPaddingLocation} border of the first table's first row: ${cellPadding.value} points`
  );
});
```

---

### getCellPadding(cellPaddingLocation)
Gets cell padding in points.

```typescript
getCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right"): OfficeExtension.ClientResult<number>;
```

#### Parameters
- cellPaddingLocation  
  "Top" | "Left" | "Bottom" | "Right"  
  Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.

#### Returns
[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<number>

#### Remarks
[API set: WordApi 1.3]

---

### getNext()
Gets the next row. Throws an ItemNotFound error if this row is the last one.

```typescript
getNext(): Word.TableRow;
```

#### Returns
[Word.TableRow](/en-us/javascript/api/word/word.tablerow)

#### Remarks
[API set: WordApi 1.3]

---

### getNextOrNullObject()
Gets the next row. If this row is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getNextOrNullObject(): Word.TableRow;
```

#### Returns
[Word.TableRow](/en-us/javascript/api/word/word.tablerow)

#### Remarks
[API set: WordApi 1.3]

---

### insertContentControl()
Inserts a content control on the row.

```typescript
insertContentControl(): Word.ContentControl;
```

#### Returns
[Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

#### Remarks
[API set: WordApiDesktop 1.1]

---

### insertRows(insertLocation, rowCount, values)
Inserts rows using this row as a template. If values are specified, inserts the values into the new rows.

```typescript
insertRows(
  insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After",
  rowCount: number,
  values?: string[][]
): Word.TableRowCollection;
```

#### Parameters
- insertLocation  
  [before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) | [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) | "Before" | "After"  
  Required. Where the new rows should be inserted, relative to the current row. It must be 'Before' or 'After'.
- rowCount  
  number  
  Required. Number of rows to add
- values  
  string[][]  
  Optional. Strings to insert in the new rows, specified as a 2D array. The number of cells in each row must not exceed the number of cells in the existing row.

#### Returns
[Word.TableRowCollection](/en-us/javascript/api/word/word.tablerowcollection)

#### Remarks
[API set: WordApi 1.3]

---

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.TableRowLoadOptions): Word.TableRow;
```

#### Parameters
- options  
  [Word.Interfaces.TableRowLoadOptions](/en-us/javascript/api/word/word.interfaces.tablerowloadoptions)  
  Provides options for which properties of the object to load.

#### Returns
[Word.TableRow](/en-us/javascript/api/word/word.tablerow)

---

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.TableRow;
```

#### Parameters
- propertyNames  
  string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

#### Returns
[Word.TableRow](/en-us/javascript/api/word/word.tablerow)

---

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
  select?: string;
  expand?: string;
}): Word.TableRow;
```

#### Parameters
- propertyNamesAndPaths  
  {
  select?: string;
  expand?: string;
  }  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

#### Returns
[Word.TableRow](/en-us/javascript/api/word/word.tablerow)

---

### merge()
Merges the row into one cell.

```typescript
merge(): Word.TableCell;
```

#### Returns
[Word.TableCell](/en-us/javascript/api/word/word.tablecell)

#### Remarks
[API set: WordApi 1.4]

---

### search(searchText, searchOptions)
Performs a search with the specified SearchOptions on the scope of the row. The search results are a collection of range objects.

```typescript
search(
  searchText: string,
  searchOptions?: Word.SearchOptions | {
    ignorePunct?: boolean;
    ignoreSpace?: boolean;
    matchCase?: boolean;
    matchPrefix?: boolean;
    matchSuffix?: boolean;
    matchWholeWord?: boolean;
    matchWildcards?: boolean;
  }
): Word.RangeCollection;
```

#### Parameters
- searchText  
  string  
  Required. The search text.
- searchOptions  
  [Word.SearchOptions](/en-us/javascript/api/word/word.searchoptions) | {
  ignorePunct?: boolean;
  ignoreSpace?: boolean;
  matchCase?: boolean;
  matchPrefix?: boolean;
  matchSuffix?: boolean;
  matchWholeWord?: boolean;
  matchWildcards?: boolean;
  }  
  Optional. Options for the search.

#### Returns
[Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

#### Remarks
[API set: WordApi 1.3]

---

### select(selectionMode)
Selects the row and navigates the Word UI to it.

```typescript
select(selectionMode?: Word.SelectionMode): void;
```

#### Parameters
- selectionMode  
  [Word.SelectionMode](/en-us/javascript/api/word/word.selectionmode)  
  Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

#### Returns
void

#### Remarks
[API set: WordApi 1.3]

---

### select(selectionMode)
Selects the row and navigates the Word UI to it.

```typescript
select(selectionMode?: "Select" | "Start" | "End"): void;
```

#### Parameters
- selectionMode  
  "Select" | "Start" | "End"  
  Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

#### Returns
void

#### Remarks
[API set: WordApi 1.3]

---

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.TableRowUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

#### Parameters
- properties  
  [Word.Interfaces.TableRowUpdateData](/en-us/javascript/api/word/word.interfaces.tablerowupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options  
  [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

#### Returns
void

---

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.TableRow): void;
```

#### Parameters
- properties  
  [Word.TableRow](/en-us/javascript/api/word/word.tablerow)

#### Returns
void

---

### setCellPadding(cellPaddingLocation, cellPadding)
Sets cell padding in points.

```typescript
setCellPadding(cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number): void;
```

#### Parameters
- cellPaddingLocation  
  [Word.CellPaddingLocation](/en-us/javascript/api/word/word.cellpaddinglocation)  
  Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
- cellPadding  
  number  
  Required. The cell padding.

#### Returns
void

#### Remarks
[API set: WordApi 1.3]

---

### setCellPadding(cellPaddingLocation, cellPadding)
Sets cell padding in points.

```typescript
setCellPadding(cellPaddingLocation: "Top" | "Left" | "Bottom" | "Right", cellPadding: number): void;
```

#### Parameters
- cellPaddingLocation  
  "Top" | "Left" | "Bottom" | "Right"  
  Required. The cell padding location must be 'Top', 'Left', 'Bottom', or 'Right'.
- cellPadding  
  number  
  Required. The cell padding.

#### Returns
void

#### Remarks
[API set: WordApi 1.3]

---

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.TableRow object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableRowData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.TableRowData;
```

#### Returns
[Word.Interfaces.TableRowData](/en-us/javascript/api/word/word.interfaces.tablerowdata)

---

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.TableRow;
```

#### Returns
[Word.TableRow](/en-us/javascript/api/word/word.tablerow)

---

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.TableRow;
```

#### Returns
[Word.TableRow](/en-us/javascript/api/word/word.tablerow)