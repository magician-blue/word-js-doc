# Word.Table class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents a table in a Word document.

Extends
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks
[ API set: WordApi 1.3 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/table-cell-access.yaml

await Word.run(async (context) => {
  // Use a two-dimensional array to hold the initial table values.
  const data = [
    ["Tokyo", "Beijing", "Seattle"],
    ["Apple", "Orange", "Pineapple"]
  ];
  const table: Word.Table = context.document.body.insertTable(2, 3, "Start", data);
  table.styleBuiltIn = Word.BuiltInStyleName.gridTable5Dark_Accent2;
  table.styleFirstColumn = false;

  await context.sync();
});
```

## Properties
- alignment: Specifies the alignment of the table against the page column. The value can be 'Left', 'Centered', or 'Right'.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- endnotes: Gets the collection of endnotes in the table.
- fields: Gets the collection of field objects in the table.
- font: Gets the font. Use this to get and set font name, size, color, and other properties.
- footnotes: Gets the collection of footnotes in the table.
- headerRowCount: Specifies the number of header rows.
- horizontalAlignment: Specifies the horizontal alignment of every cell in the table. The value can be 'Left', 'Centered', 'Right', or 'Justified'.
- isUniform: Indicates whether all of the table rows are uniform.
- nestingLevel: Gets the nesting level of the table. Top-level tables have level 1.
- parentBody: Gets the parent body of the table.
- parentContentControl: Gets the content control that contains the table. Throws an ItemNotFound error if there isn't a parent content control.
- parentContentControlOrNullObject: Gets the content control that contains the table. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- parentTable: Gets the table that contains this table. Throws an ItemNotFound error if it isn't contained in a table.
- parentTableCell: Gets the table cell that contains this table. Throws an ItemNotFound error if it isn't contained in a table cell.
- parentTableCellOrNullObject: Gets the table cell that contains this table. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- parentTableOrNullObject: Gets the table that contains this table. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- rowCount: Gets the number of rows in the table.
- rows: Gets all of the table rows.
- shadingColor: Specifies the shading color. Color is specified in "#RRGGBB" format or by using the color name.
- style: Specifies the style name for the table. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBandedColumns: Specifies whether the table has banded columns.
- styleBandedRows: Specifies whether the table has banded rows.
- styleBuiltIn: Specifies the built-in style name for the table. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- styleFirstColumn: Specifies whether the table has a first column with a special style.
- styleLastColumn: Specifies whether the table has a last column with a special style.
- styleTotalRow: Specifies whether the table has a total (last) row with a special style.
- tables: Gets the child tables nested one level deeper.
- values: Specifies the text values in the table, as a 2D JavaScript array.
- verticalAlignment: Specifies the vertical alignment of every cell in the table. The value can be 'Top', 'Center', or 'Bottom'.
- width: Specifies the width of the table in points.

## Methods
- addColumns(insertLocation, columnCount, values): Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows.
- addRows(insertLocation, rowCount, values): Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows.
- autoFitWindow(): Autofits the table columns to the width of the window.
- clear(): Clears the contents of the table.
- delete(): Deletes the entire table.
- deleteColumns(columnIndex, columnCount): Deletes specific columns. This is applicable to uniform tables.
- deleteRows(rowIndex, rowCount): Deletes specific rows.
- distributeColumns(): Distributes the column widths evenly. This is applicable to uniform tables.
- getBorder(borderLocation): Gets the border style for the specified border.
- getBorder(borderLocation): Gets the border style for the specified border.
- getCell(rowIndex, cellIndex): Gets the table cell at a specified row and column. Throws an ItemNotFound error if the specified table cell doesn't exist.
- getCellOrNullObject(rowIndex, cellIndex): Gets the table cell at a specified row and column. If the specified table cell doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- getCellPadding(cellPaddingLocation): Gets cell padding in points.
- getCellPadding(cellPaddingLocation): Gets cell padding in points.
- getNext(): Gets the next table. Throws an ItemNotFound error if this table is the last one.
- getNextOrNullObject(): Gets the next table. If this table is the last one, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- getParagraphAfter(): Gets the paragraph after the table. Throws an ItemNotFound error if there isn't a paragraph after the table.
- getParagraphAfterOrNullObject(): Gets the paragraph after the table. If there isn't a paragraph after the table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- getParagraphBefore(): Gets the paragraph before the table. Throws an ItemNotFound error if there isn't a paragraph before the table.
- getParagraphBeforeOrNullObject(): Gets the paragraph before the table. If there isn't a paragraph before the table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- getRange(rangeLocation): Gets the range that contains this table, or the range at the start or end of the table.
- insertContentControl(): Inserts a content control on the table.
- insertParagraph(paragraphText, insertLocation): Inserts a paragraph at the specified location.
- insertTable(rowCount, columnCount, insertLocation, values): Inserts a table with the specified number of rows and columns.
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- mergeCells(topRow, firstCell, bottomRow, lastCell): Merges the cells bounded inclusively by a first and last cell.
- search(searchText, searchOptions): Performs a search with the specified SearchOptions on the scope of the table object. The search results are a collection of range objects.
- select(selectionMode): Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
- select(selectionMode): Selects the table, or the position at the start or end of the table, and navigates the Word UI to it.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- setCellPadding(cellPaddingLocation, cellPadding): Sets cell padding in points.
- setCellPadding(cellPaddingLocation, cellPadding): Sets cell padding in points.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Table object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.TableData) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so p