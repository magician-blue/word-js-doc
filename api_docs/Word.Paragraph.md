# Word.Paragraph class

Package: [word](/en-us/javascript/api/word)

Represents a single paragraph in a selection, range, content control, or document body.

Extends
- OfficeExtension.ClientObject (https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/get-paragraph-on-insertion-point.yaml

await Word.run(async (context) => {
  // The collection of paragraphs of the current selection returns the full paragraphs contained in it.
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  paragraph.load("text");

  await context.sync();
  console.log(paragraph.text);
});
```

## Properties
- alignment — Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.
- borders — Returns a BorderUniversalCollection object that represents all the borders for the paragraph. Note: Preview API, do not use in production.
- contentControls — Gets the collection of content control objects in the paragraph.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- endnotes — Gets the collection of endnotes in the paragraph.
- fields — Gets the collection of fields in the paragraph.
- firstLineIndent — Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.
- font — Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.
- footnotes — Gets the collection of footnotes in the paragraph.
- inlinePictures — Gets the collection of InlinePicture objects in the paragraph. The collection doesn't include floating images.
- isLastParagraph — Indicates the paragraph is the last one inside its parent body.
- isListItem — Checks whether the paragraph is a list item.
- leftIndent — Specifies the left indent value, in points, for the paragraph.
- lineSpacing — Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.
- lineUnitAfter — Specifies the amount of spacing, in grid lines, after the paragraph.
- lineUnitBefore — Specifies the amount of spacing, in grid lines, before the paragraph.
- list — Gets the List to which this paragraph belongs. Throws an ItemNotFound error if the paragraph isn't in a list.
- listItem — Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.
- listItemOrNullObject — Gets the ListItem for the paragraph. If the paragraph isn't part of a list, this returns an object with isNullObject = true. See “OrNullObject methods and properties”.
- listOrNullObject — Gets the List to which this paragraph belongs. If the paragraph isn't in a list, this returns an object with isNullObject = true. See “OrNullObject methods and properties”.
- outlineLevel — Specifies the outline level for the paragraph.
- parentBody — Gets the parent body of the paragraph.
- parentContentControl — Gets the content control that contains the paragraph. Throws an ItemNotFound error if there isn't a parent content control.
- parentContentControlOrNullObject — Gets the content control that contains the paragraph. If there isn't a parent content control, this returns an object with isNullObject = true. See “OrNullObject methods and properties”.
- parentTable — Gets the table that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table.
- parentTableCell — Gets the table cell that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table cell.
- parentTableCellOrNullObject — Gets the table cell that contains the paragraph. If it isn't contained in a table cell, this returns an object with isNullObject = true. See “OrNullObject methods and properties”.
- parentTableOrNullObject — Gets the table that contains the paragraph. If it isn't contained in a table, this returns an object with isNullObject = true. See “OrNullObject methods and properties”.
- range — Gets a Range object that represents the portion of the document that's contained within the paragraph. Note: Preview API, do not use in production.
- rightIndent — Specifies the right indent value, in points, for the paragraph.
- shading — Returns a ShadingUniversal object that refers to the shading formatting for the paragraph. Note: Preview API, do not use in production.
- shapes — Gets the collection of shape objects anchored in the paragraph, including both inline and floating shapes. Currently supported: text boxes, geometric shapes, groups, pictures, and canvases.
- spaceAfter — Specifies the spacing, in points, after the paragraph.
- spaceBefore — Specifies the spacing, in points, before the paragraph.
- style — Specifies the style name for the paragraph. Use this property for custom styles and localized style names. For built-in styles portable between locales, use styleBuiltIn.
- styleBuiltIn — Specifies the built-in style name for the paragraph. Use for built-in styles portable between locales. For custom or localized style names, use style.
- tableNestingLevel — Gets the level of the paragraph's table. Returns 0 if the paragraph isn't in a table.
- text — Gets the text of the paragraph.
- uniqueLocalId — Gets a string that represents the paragraph identifier in the current session. ID is 8-4-4-4-12 GUID without braces and differs across sessions and coauthors.

## Methods
- attachToList(listId, level) — Lets the paragraph join an existing list at the specified level. Fails if it cannot join or is already a list item.
- clear() — Clears the contents of the paragraph object. The user can undo the cleared content.
- closeUp() — Removes any spacing before the paragraph. Note: Preview API.
- delete() — Deletes the paragraph and its content from the document.
- detachFromList() — Moves this paragraph out of its list, if the paragraph is a list item.
- getAnnotations() — Gets annotations set on this Paragraph object.
- getComments() — Gets comments associated with the paragraph.
- getContentControls(options) — Gets the currently supported content controls in the paragraph.
- getHtml() — Gets an HTML representation of the paragraph object. Formatting is close but not exact and may vary across platforms; use getOoxml() for fidelity and convert to HTML.
- getNext() — Gets the next paragraph. Throws ItemNotFound if this paragraph is the last one.
- getNextOrNullObject() — Gets the next paragraph; if this is last, returns an object with isNullObject = true. See “OrNullObject methods and properties”.
- getOoxml() — Gets the Office Open XML (OOXML) representation of the paragraph object.
- getPrevious() — Gets the previous paragraph. Throws ItemNotFound if this paragraph is the first one.
- getPreviousOrNullObject() — Gets the previous paragraph; if this is first, returns an object with isNullObject = true. See “OrNullObject methods and properties”.
- getRange(rangeLocation) — Gets the whole paragraph, or the start/end/after/content of the paragraph, as a range.
- getReviewedText(changeTrackingVersion) — Gets reviewed text based on ChangeTrackingVersion selection.
- getReviewedText(changeTrackingVersion) — Gets reviewed text based on ChangeTrackingVersion selection (string literal overload).
- getText(options) — Returns the text of the paragraph (excludes equations, graphics, and special content markers; hidden and deleted text excluded by default).
- getTextRanges(endingMarks, trimSpacing) — Gets the text ranges in the paragraph using punctuation and/or other ending marks.
- getTrackedChanges() — Gets the collection of the TrackedChange objects in the paragraph.
- indent() — Indents the paragraph by one level. Note: Preview API.
- indentCharacterWidth(count) — Indents the paragraph by a specified number of characters. Note: Preview API.
- indentFirstLineCharacterWidth(count) — Indents the first line by a specified number of characters. Note: Preview API.
- insertAnnotations(annotations) — Inserts annotations on this Paragraph object.
- insertBreak(breakType, insertLocation) — Inserts a break at the specified location in the main document.
- insertCanvas(insertShapeOptions) — Inserts a floating canvas in front of text with its anchor at the beginning of the paragraph.
- insertContentControl(contentControlType) — Wraps the Paragraph object with a content control.
- insertFileFromBase64(base64File, insertLocation) — Inserts a document into the paragraph at the specified location.
- insertGeometricShape(geometricShapeType, insertShapeOptions) — Inserts a geometric shape in front of text with its anchor at the beginning of the paragraph.
- insertGeometricShape(geometricShapeType, insertShapeOptions) — Inserts a geometric shape (string literal overload).
- insertHtml(html, insertLocation) — Inserts HTML into the paragraph at the specified location.
- insertInlinePictureFromBase64(base64EncodedImage, insertLocation) — Inserts a picture into the paragraph at the specified location.
- insertOoxml(ooxml, insertLocation) — Inserts OOXML into the paragraph at the specified location.
- insertParagraph(paragraphText, insertLocation) — Inserts a paragraph at the specified location.
- insertPictureFromBase64(base64EncodedImage, insertShapeOptions) — Inserts a floating picture in front of text with its anchor at the beginning of the paragraph.
- insertTable(rowCount, columnCount, insertLocation, values) — Inserts a table with the specified number of rows and columns.
- insertText(text, insertLocation) — Inserts text into the paragraph at the specified location.
- insertTextBox(text, insertShapeOptions) — Inserts a floating text box in front of text with its anchor at the beginning of the paragraph.
- joinList() — Joins a list paragraph with the closest list above or below this paragraph. Note: Preview API.
- load(options) — Queues up a command to load specified properties; call context.sync() before reading.
- load(propertyNames) — Queues up a command to load specified properties; call context.sync() before reading.
- load(propertyNamesAndPaths) — Queues up a command to load specified properties; call context.sync() before reading.
- next(count) — Returns the next paragraph object. Note: Preview API.
- openOrCloseUp() — Toggles the spacing before the paragraph. Note: Preview API.
- openUp() — Sets spacing before the paragraph to 12 points. Note: Preview API.
- outdent() — Removes one level of indent for the paragraph. Note: Preview API.
- outlineDemote() — Applies the next heading level style (Heading 1 through Heading 8). Note: Preview API.
- outlineDemoteToBody() — Demotes the paragraph to body text by applying the Normal style. Note: Preview API.
- outlinePromote() — Applies the previous heading level style (Heading 1 through Heading 8). Note: Preview API.
- previous(count) — Returns the previous paragraph as a Paragraph object. Note: Preview API.
- reset() — Removes manual paragraph formatting (formatting not applied using a style). Note: Preview API.
- resetAdvanceTo() — Resets paragraph that uses custom list levels to the original level settings. Note: Preview API.
- search(searchText, searchOptions) — Performs a search with specified SearchOptions; results are a collection of range objects.
- select(selectionMode) — Selects and navigates the Word UI to the paragraph.
- select(selectionMode) — Selects and navigates the Word UI to the paragraph (string literal overload).
- selectNumber() — Selects the number or bullet in a list. Note: Preview API.
- separateList() — Separates a list into two lists; numbering restarts for numbered lists. Note: Preview API.
- set(properties, options) — Sets multiple properties at once with a plain object or another API object of same type.
- set(properties) — Sets multiple properties at once based on an existing loaded object.
- space1() — Sets the paragraph to single spacing. Note: Preview API.
- space1Pt5() — Sets the paragraph to 1.5-line spacing. Note: Preview API.
- space2() — Sets the paragraph to double spacing. Note: Preview API.
- split(delimiters, trimDelimiters, trimSpacing) — Splits the paragraph into child ranges by using delimiters.
- startNewList() — Starts a new list with this paragraph. Fails if the paragraph is already a list item.
- tabHangingIndent(count) — Sets a hanging indent to a specified number of tab stops. Note: Preview API.
- tabIndent(count) — Sets the left indent to a specified number of tab stops. Note: Preview API.
- toJSON() — Returns a plain JavaScript object with shallow copies of any loaded child properties for JSON serialization.
- track() — Track the object for automatic adjustment based on surrounding changes. Shortcut for context.trackedObjects.add(thisObject).
- untrack() — Release memory associated with this object if previously tracked. Shortcut for context.trackedObjects.remove(thisObject). Requires context.sync().

## Events
- onCommentAdded — Occurs when new comments are added. Note: Preview API.
- onCommentChanged — Occurs when a comment or its reply is changed. Note: Preview API.
- onCommentDeleted — Occurs when comments are deleted. Note: Preview API.
- onCommentDeselected — Occurs when a comment is deselected. Note: Preview API.
- onCommentSelected — Occurs when a comment is selected. Note: Preview API.

## Property Details

### alignment
Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.

```typescript
alignment: Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified";
```

Property Value
- [Word.Alignment](/en-us/javascript/api/word/word.alignment) | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Center last paragraph alignment.
  context.document.body.paragraphs.getLast().alignment = "Centered";

  await context.sync();
});
```

### borders
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BorderUniversalCollection object that represents all the borders for the paragraph.

```typescript
readonly borders: Word.BorderUniversalCollection;
```

Property Value
- [Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### contentControls
Gets the collection of content control objects in the paragraph.

```typescript
readonly contentControls: Word.ContentControlCollection;
```

Property Value
- [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks
[ API set: WordApi 1.1 ]

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### endnotes
Gets the collection of endnotes in the paragraph.

```typescript
readonly endnotes: Word.NoteItemCollection;
```

Property Value
- [Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

Remarks
[ API set: WordApi 1.5 ]

### fields
Gets the collection of fields in the paragraph.

```typescript
readonly fields: Word.FieldCollection;
```

Property Value
- [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

Remarks
[ API set: WordApi 1.4 ]

### firstLineIndent
Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

```typescript
firstLineIndent: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

### font
Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.

```typescript
readonly font: Word.Font;
```

Property Value
- [Word.Font](/en-us/javascript/api/word/word.font)

Remarks
[ API set: WordApi 1.1 ]

### footnotes
Gets the collection of footnotes in the paragraph.

```typescript
readonly footnotes: Word.NoteItemCollection;
```

Property Value
- [Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

Remarks
[ API set: WordApi 1.5 ]

### inlinePictures
Gets the collection of InlinePicture objects in the paragraph. The collection doesn't include floating images.

```typescript
readonly inlinePictures: Word.InlinePictureCollection;
```

Property Value
- [Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)

Remarks
[ API set: WordApi 1.1 ]

### isLastParagraph
Indicates the paragraph is the last one inside its parent body.

```typescript
readonly isLastParagraph: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi 1.3 ]

### isListItem
Checks whether the paragraph is a list item.

```typescript
readonly isListItem: boolean;
```

Property Value
- boolean

Remarks
[ API set: WordApi 1.3 ]

### leftIndent
Specifies the left indent value, in points, for the paragraph.

```typescript
leftIndent: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Indent the first paragraph.
  context.document.body.paragraphs.getFirst().leftIndent = 75; //units = points

  return context.sync();
});
```

### lineSpacing
Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.

```typescript
lineSpacing: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Adjust line spacing.
  context.document.body.paragraphs.getFirst().lineSpacing = 20;

  await context.sync();
});
```

### lineUnitAfter
Specifies the amount of spacing, in grid lines, after the paragraph.

```typescript
lineUnitAfter: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Set the space (in line units) after the first paragraph.
  context.document.body.paragraphs.getFirst().lineUnitAfter = 1;

  await context.sync();
});
```

### lineUnitBefore
Specifies the amount of spacing, in grid lines, before the paragraph.

```typescript
lineUnitBefore: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Set the space (in line units) before the first paragraph.
  context.document.body.paragraphs.getFirst().lineUnitBefore = 1;

  await context.sync();
});
```

### list
Gets the List to which this paragraph belongs. Throws an ItemNotFound error if the paragraph isn't in a list.

```typescript
readonly list: Word.List;
```

Property Value
- [Word.List](/en-us/javascript/api/word/word.list)

Remarks
[ API set: WordApi 1.3 ]

### listItem
Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.

```typescript
readonly listItem: Word.ListItem;
```

Property Value
- [Word.ListItem](/en-us/javascript/api/word/word.listitem)

Remarks
[ API set: WordApi 1.3 ]

### listItemOrNullObject
Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

```typescript
readonly listItemOrNullObject: Word.ListItem;
```

Property Value
- [Word.ListItem](/en-us/javascript/api/word/word.listitem)

Remarks
[ API set: WordApi 1.3 ]

### listOrNullObject
Gets the List to which this paragraph belongs. If the paragraph isn't in a list, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

```typescript
readonly listOrNullObject: Word.List;
```

Property Value
- [Word.List](/en-us/javascript/api/word/word.list)

Remarks
[ API set: WordApi 1.3 ]

### outlineLevel
Specifies the outline level for the paragraph.

```typescript
outlineLevel: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

### parentBody
Gets the parent body of the paragraph.

```typescript
readonly parentBody: Word.Body;
```

Property Value
- [Word.Body](/en-us/javascript/api/word/word.body)

Remarks
[ API set: WordApi 1.3 ]

### parentContentControl
Gets the content control that contains the paragraph. Throws an ItemNotFound error if there isn't a parent content control.

```typescript
readonly parentContentControl: Word.ContentControl;
```

Property Value
- [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks
[ API set: WordApi 1.1 ]

### parentContentControlOrNullObject
Gets the content control that contains the paragraph. If there isn't a parent content control, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

```typescript
readonly parentContentControlOrNullObject: Word.ContentControl;
```

Property Value
- [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks
[ API set: WordApi 1.3 ]

### parentTable
Gets the table that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table.

```typescript
readonly parentTable: Word.Table;
```

Property Value
- [Word.Table](/en-us/javascript/api/word/word.table)

Remarks
[ API set: WordApi 1.3 ]

### parentTableCell
Gets the table cell that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table cell.

```typescript
readonly parentTableCell: Word.TableCell;
```

Property Value
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Remarks
[ API set: WordApi 1.3 ]

### parentTableCellOrNullObject
Gets the table cell that contains the paragraph. If it isn't contained in a table cell, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

```typescript
readonly parentTableCellOrNullObject: Word.TableCell;
```

Property Value
- [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Remarks
[ API set: WordApi 1.3 ]

### parentTableOrNullObject
Gets the table that contains the paragraph. If it isn't contained in a table, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

```typescript
readonly parentTableOrNullObject: Word.Table;
```

Property Value
- [Word.Table](/en-us/javascript/api/word/word.table)

Remarks
[ API set: WordApi 1.3 ]

### range
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a Range object that represents the portion of the document that's contained within the paragraph.

```typescript
readonly range: Word.Range;
```

Property Value
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### rightIndent
Specifies the right indent value, in points, for the paragraph.

```typescript
rightIndent: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

### shading
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a ShadingUniversal object that refers to the shading formatting for the paragraph.

```typescript
readonly shading: Word.ShadingUniversal;
```

Property Value
- [Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### shapes
Gets the collection of shape objects anchored in the paragraph, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

```typescript
readonly shapes: Word.ShapeCollection;
```

Property Value
- [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks
[ API set: WordApiDesktop 1.2 ]

### spaceAfter
Specifies the spacing, in points, after the paragraph.

```typescript
spaceAfter: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Set the space (in points) after the first paragraph.
  context.document.body.paragraphs.getFirst().spaceAfter = 20;

  await context.sync();
});
```

### spaceBefore
Specifies the spacing, in points, before the paragraph.

```typescript
spaceBefore: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.1 ]

### style
Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style: string;
```

Property Value
- string

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Applies the specified style to a paragraph.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to apply.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else if (style.type != Word.StyleType.paragraph) {
    console.log(`The '${styleName}' style isn't a paragraph style.`);
  } else {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    const paragraph: Word.Paragraph = body.paragraphs.getFirst();
    paragraph.style = style.nameLocal;
    console.log(`'${styleName}' style applied to first paragraph.`);
  }
});
```

### styleBuiltIn
Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

Property Value
- [Word.BuiltInStyleName](/en-us/javascript/api/word/word.builtinstylename) | (many string literal values as above)

Remarks
[ API set: WordApi 1.3 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/90-scenarios/doc-assembly.yaml

await Word.run(async (context) => {
    const paragraph: Word.Paragraph = context.document.body.insertParagraph("Timeline", "End");
    paragraph.styleBuiltIn = "Heading2";
    const paragraph2: Word.Paragraph = context.document.body.insertParagraph("The Services shall commence on July 31, 2015, and shall continue through July 29, 2015.", "End");
    paragraph2.styleBuiltIn = "Normal";
    const paragraph3: Word.Paragraph = context.document.body.insertParagraph("Project Costs by Phase", "End");
    paragraph3.styleBuiltIn = "Heading2";
    // Note a content control with the title of "ProjectCosts" is added. Content will be replaced later.
    const paragraph4: Word.Paragraph = context.document.body.insertParagraph("<Add Project Costs Here>", "End");
    paragraph4.styleBuiltIn = "Normal";
    paragraph4.font.highlightColor = "#FFFF00";
    const contentControl: Word.ContentControl = paragraph4.insertContentControl();
    contentControl.title = "ProjectCosts";
    const paragraph5: Word.Paragraph = context.document.body.insertParagraph("Project Team", "End");
    paragraph5.styleBuiltIn = "Heading2";
    paragraph5.font.highlightColor = "#FFFFFF";
    const paragraph6: Word.Paragraph = context.document.body.insertParagraph("Terms of Work", "End");
    paragraph6.styleBuiltIn = "Heading1";
    const paragraph7: Word.Paragraph = context.document.body.insertParagraph("Contractor shall provide the Services and Deliverable(s) as follows:", "End");
    paragraph7.styleBuiltIn = "Normal";
    const paragraph8: Word.Paragraph = context.document.body.insertParagraph("Out-of-Pocket Expenses / Invoice Procedures", "End");
    paragraph8.styleBuiltIn = "Heading2";
    const paragraph9 : Word.Paragraph= context.document.body.insertParagraph("Client will be invoiced monthly for the consulting services and T&L expenses. Standard Contractor invoicing is assumed to be acceptable. Invoices are due upon receipt. client will be invoiced all costs associated with out-of-pocket expenses (including, without limitation, costs and expenses associated with meals, lodging, local transportation and any other applicable business expenses) listed on the invoice as a separate line item. Reimbursement for out-of-pocket expenses in connection with performance of this SOW, when authorized and up to the limits set forth in this SOW, shall be in accordance with Client's then-current published policies governing travel and associated business expenses, which information shall be provided by the Client Project Manager.", "End");
    paragraph9.styleBuiltIn = "Normal";
    // Insert a page break at the end of the document.
    context.document.body.insertBreak("Page", "End");

    await context.sync();
});
```

### tableNestingLevel
Gets the level of the paragraph's table. It returns 0 if the paragraph isn't in a table.

```typescript
readonly tableNestingLevel: number;
```

Property Value
- number

Remarks
[ API set: WordApi 1.3 ]

### text
Gets the text of the paragraph.

```typescript
readonly text: string;
```

Property Value
- string

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/get-paragraph-on-insertion-point.yaml

await Word.run(async (context) => {
  // The collection of paragraphs of the current selection returns the full paragraphs contained in it.
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  paragraph.load("text");

  await context.sync();
  console.log(paragraph.text);
});
```

### uniqueLocalId
Gets a string that represents the paragraph identifier in the current session. ID is in standard 8-4-4-4-12 GUID format without curly braces and differs across sessions and coauthors.

```typescript
readonly uniqueLocalId: string;
```

Property Value
- string

Remarks
[ API set: WordApi 1.6 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Registers event handlers.
await Word.run(async (context) => {
  eventContexts[0] = context.document.onParagraphAdded.add(paragraphChanged);
  eventContexts[1] = context.document.onParagraphChanged.add(paragraphChanged);

  eventContexts[2] = context.document.onAnnotationClicked.add(onClickedHandler);
  eventContexts[3] = context.document.onAnnotationHovered.add(onHoveredHandler);
  eventContexts[4] = context.document.onAnnotationInserted.add(onInsertedHandler);
  eventContexts[5] = context.document.onAnnotationRemoved.add(onRemovedHandler);
  eventContexts[6] = context.document.onAnnotationPopupAction.add(onPopupActionHandler);

  await context.sync();

  console.log("Event handlers registered.");
});
```

...

```typescript
async function paragraphChanged(args: Word.ParagraphChangedEventArgs) {
  await Word.run(async (context) => {
    const results = [];
    for (let id of args.uniqueLocalIds) {
      let para = context.document.getParagraphByUniqueLocalId(id);
      para.load("uniqueLocalId");

      results.push({ para: para, text: para.getText() });
    }

    await context.sync();

    for (let result of results) {
      console.log(`${args.type}: ID ${result.para.uniqueLocalId}:-`, result.text.value);
    }
  });
}
```

## Method Details

### attachToList(listId, level)
Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.

```typescript
attachToList(listId: number, level: number): Word.List;
```

Parameters
- listId (number) — Required. The ID of an existing list.
- level (number) — Required. The level in the list.

Returns
- [Word.List](/en-us/javascript/api/word/word.list)

Remarks
[ API set: WordApi 1.3 ]

### clear()
Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.

```typescript
clear(): void;
```

Returns
- void

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the style property for all of the paragraphs.
    paragraphs.load('style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a command to clear the contents of the first paragraph.
    paragraphs.items[0].clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Cleared the contents of the first paragraph.');
});
```

### closeUp()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes any spacing before the paragraph.

```typescript
closeUp(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### delete()
Deletes the paragraph and its content from the document.

```typescript
delete(): void;
```

Returns
- void

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the text property for all of the paragraphs.
    paragraphs.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a command to delete the first paragraph.
    paragraphs.items[0].delete();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Deleted the first paragraph.');
});
```

### detachFromList()
Moves this paragraph out of its list, if the paragraph is a list item.

```typescript
detachFromList(): void;
```

Returns
- void

Remarks
[ API set: WordApi 1.3 ]

### getAnnotations()
Gets annotations set on this Paragraph object.

```typescript
getAnnotations(): Word.AnnotationCollection;
```

Returns
- [Word.AnnotationCollection](/en-us/javascript/api/word/word.annotationcollection)

Remarks
[ API set: WordApi 1.7 ]

Important: This API requires a Microsoft 365 subscription in order to work properly because of an underlying service's requirement. For more about this, see GitHub issue 4953 (https://github.com/OfficeDev/office-js/issues/4953).

### getComments()
Gets comments associated with the paragraph.

```typescript
getComments(): Word.CommentCollection;
```

Returns
- [Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)

Remarks
[ API set: WordApi 1.4 ]

### getContentControls(options)
Gets the currently supported content controls in the paragraph.

```typescript
getContentControls(options?: Word.ContentControlOptions): Word.ContentControlCollection;
```

Parameters
- options ([Word.ContentControlOptions](/en-us/javascript/api/word/word.contentcontroloptions)) — Optional. Options that define which content controls are returned.

Returns
- [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks
[ API set: WordApi 1.5 ]

Important: If specific types are provided in the options parameter, only content controls of supported types are returned. Be aware that an exception will be thrown on using methods of a generic [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) that aren't relevant for the specific type. Over time, additional types may be supported. Your add-in should request and handle specific types of content controls.

### getHtml()
Gets an HTML representation of the paragraph object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use Paragraph.getOoxml() and convert the returned XML to HTML.

```typescript
getHtml(): OfficeExtension.ClientResult<string>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the style property for all of the paragraphs.
    paragraphs.load('style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a set of commands to get the HTML of the first paragraph.
    const html = paragraphs.items[0].getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Paragraph HTML: ' + html.value);
});
```

### getNext()
Gets the next paragraph. Throws an ItemNotFound error if the paragraph is the last one.

```typescript
getNext(): Word.Paragraph;
```

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks
[ API set: WordApi 1.3 ]

### getNextOrNullObject()
Gets the next paragraph. If the paragraph is the last one, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

```typescript
getNextOrNullObject(): Word.Paragraph;
```

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks
[ API set: WordApi 1.3 ]

### getOoxml()
Gets the Office Open XML (OOXML) representation of the paragraph object.

```typescript
getOoxml(): OfficeExtension.ClientResult<string>;
```

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the style property for the top 2 paragraphs.
    paragraphs.load({select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a set of commands to get the OOXML of the first paragraph.
    const ooxml = paragraphs.items[0].getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Paragraph OOXML: ' + ooxml.value);
});
```

### getPrevious()
Gets the previous paragraph. Throws an ItemNotFound error if the paragraph is the first one.

```typescript
getPrevious(): Word.Paragraph;
```

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks
[ API set: WordApi 1.3 ]

### getPreviousOrNullObject()
Gets the previous paragraph. If the paragraph is the first one, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

```typescript
getPreviousOrNullObject(): Word.Paragraph;
```

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks
[ API set: WordApi 1.3 ]

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the text property for all of the paragraphs.
    paragraphs.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue commands to create a proxy object for the next-to-last paragraph.
    const indexOfLastParagraph = paragraphs.items.length - 1;
    const precedingParagraph = paragraphs.items[indexOfLastParagraph].getPreviousOrNullObject();

    // Queue a command to load the text of the preceding paragraph.
    precedingParagraph.load('text');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    if (precedingParagraph.isNullObject) {
        console.log('There are no paragraphs before the current one.');
    } else {
        console.log('The preceding paragraph is: ' + precedingParagraph.text);
    }
});
```

### getRange(rangeLocation)
Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.

```typescript
getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | Word.RangeLocation.after | Word.RangeLocation.content | "Whole" | "Start" | "End" | "After" | "Content"): Word.Range;
```

Parameters
- rangeLocation (Word.RangeLocation.whole | .start | .end | .after | .content | "Whole" | "Start" | "End" | "After" | "Content") — Optional. The range location must be 'Whole', 'Start', 'End', 'After', or 'Content'.

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
[ API set: WordApi 1.3 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/get-paragraph-on-insertion-point.yaml

await Word.run(async (context) => {
  // Get the complete sentence (as range) associated with the insertion point.
  const sentences: Word.RangeCollection = context.document
    .getSelection()
    .getTextRanges(["."] /* Using the "." as delimiter */, false /*means without trimming spaces*/);
  sentences.load("$none");
  await context.sync();

  // Expand the range to the end of the paragraph to get all the complete sentences.
  const sentencesToTheEndOfParagraph: Word.RangeCollection = sentences.items[0]
    .getRange()
    .expandTo(
      context.document
        .getSelection()
        .paragraphs.getFirst()
        .getRange(Word.RangeLocation.end)
    )
    .getTextRanges(["."], false /* Don't trim spaces*/);
  sentencesToTheEndOfParagraph.load("text");
  await context.sync();

  for (let i = 0; i < sentencesToTheEndOfParagraph.items.length; i++) {
    console.log(sentencesToTheEndOfParagraph.items[i].text);
  }
});
```

### getReviewedText(changeTrackingVersion)
Gets reviewed text based on ChangeTrackingVersion selection.

```typescript
getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion): OfficeExtension.ClientResult<string>;
```

Parameters
- changeTrackingVersion ([Word.ChangeTrackingVersion](/en-us/javascript/api/word/word.changetrackingversion)) — Optional. The value must be 'Original' or 'Current'. The default is 'Current'.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks
[ API set: WordApi 1.4 ]

### getReviewedText(changeTrackingVersion) (string literal overload)
Gets reviewed text based on ChangeTrackingVersion selection.

```typescript
getReviewedText(changeTrackingVersion?: "Original" | "Current"): OfficeExtension.ClientResult<string>;
```

Parameters
- changeTrackingVersion ("Original" | "Current") — Optional. The value must be 'Original' or 'Current'. The default is 'Current'.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks
[ API set: WordApi 1.4 ]

### getText(options)
Returns the text of the paragraph. This excludes equations, graphics (e.g., images, videos, drawings), and special characters that mark various content (e.g., for content controls, fields, comments, footnotes, endnotes). By default, hidden text and text marked as deleted are excluded.

```typescript
getText(options?: Word.GetTextOptions | {
            IncludeHiddenText?: boolean;
            IncludeTextMarkedAsDeleted?: boolean;
        }): OfficeExtension.ClientResult<string>;
```

Parameters
- options ([Word.GetTextOptions](/en-us/javascript/api/word/word.gettextoptions) | { IncludeHiddenText?: boolean; IncludeTextMarkedAsDeleted?: boolean; }) — Optional. Options that define whether the final result should include hidden text and text marked as deleted.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks
[ API set: WordApi 1.7 ]

### getTextRanges(endingMarks, trimSpacing)
Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.

```typescript
getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
```

Parameters
- endingMarks (string[]) — Required. The punctuation marks and/or other ending marks as an array of strings.
- trimSpacing (boolean) — Optional. Whether to trim spacing characters from the start and end of the returned ranges. Default is false.

Returns
- [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

Remarks
[ API set: WordApi 1.3 ]

### getTrackedChanges()
Gets the collection of the TrackedChange objects in the paragraph.

```typescript
getTrackedChanges(): Word.TrackedChangeCollection;
```

Returns
- [Word.TrackedChangeCollection](/en-us/javascript/api/word/word.trackedchangecollection)

Remarks
[ API set: WordApi 1.6 ]

### indent()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indents the paragraph by one level.

```typescript
indent(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### indentCharacterWidth(count)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indents the paragraph by a specified number of characters.

```typescript
indentCharacterWidth(count: number): void;
```

Parameters
- count (number) — The number of characters for the indent.

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### indentFirstLineCharacterWidth(count)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indents the first line of the paragraph by the specified number of characters.

```typescript
indentFirstLineCharacterWidth(count: number): void;
```

Parameters
- count (number) — The number of characters for the first line indent.

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### insertAnnotations(annotations)
Inserts annotations on this Paragraph object.

```typescript
insertAnnotations(annotations: Word.AnnotationSet): OfficeExtension.ClientResult<string[]>;
```

Parameters
- annotations ([Word.AnnotationSet](/en-us/javascript/api/word/word.annotationset)) — Annotations to set.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string[]> — An array of the inserted annotations identifiers.

Remarks
[ API set: WordApi 1.7 ]

Important: This API requires a Microsoft 365 subscription in order to work properly because of an underlying service's requirement. For more about this, see GitHub issue 4953 (https://github.com/OfficeDev/office-js/issues/4953).

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Adds annotations to the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const options: Word.CritiquePopupOptions = {
    brandingTextResourceId: "PG.TabLabel",
    subtitleResourceId: "PG.HelpCommand.TipTitle",
    titleResourceId: "PG.HelpCommand.Label",
    suggestions: ["suggestion 1", "suggestion 2", "suggestion 3"]
  };
  const critique1: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.red,
    start: 1,
    length: 3,
    popupOptions: options
  };
  const critique2: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.green,
    start: 6,
    length: 1,
    popupOptions: options
  };
  const critique3: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.blue,
    start: 10,
    length: 3,
    popupOptions: options
  };
  const critique4: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.lavender,
    start: 14,
    length: 3,
    popupOptions: options
  };
  const critique5: Word.Critique = {
    colorScheme: Word.CritiqueColorScheme.berry,
    start: 18,
    length: 10,
    popupOptions: options
  };
  const annotationSet: Word.AnnotationSet = {
    critiques: [critique1, critique2, critique3, critique4, critique5]
  };

  const annotationIds = paragraph.insertAnnotations(annotationSet);

  await context.sync();

  console.log("Annotations inserted:", annotationIds.value);
});
```

### insertBreak(breakType, insertLocation)
Inserts a break at the specified location in the main document.

```typescript
insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): void;
```

Parameters
- breakType ([Word.BreakType](/en-us/javascript/api/word/word.breaktype) | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line") — Required. The break type to add to the document.
- insertLocation ([Word.InsertLocation.before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) | [Word.InsertLocation.after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) | "Before" | "After") — Required. The value must be 'Before' or 'After'.

Returns
- void

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    paragraphs.load({select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a command to get the first paragraph.
    const paragraph = paragraphs.items[0];

    // Queue a command to insert a page break after the first paragraph.
    paragraph.insertBreak(Word.BreakType.page, Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Inserted a page break after the paragraph.');
});
```

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-line-and-page-breaks.yaml

Word.run(async (context) => {
  context.document.body.paragraphs.getFirst().insertBreak(Word.BreakType.line, "After");

  await context.sync();
  console.log("success");
});
```

### insertCanvas(insertShapeOptions)
Inserts a floating canvas in front of text with its anchor at the beginning of the paragraph.

```typescript
insertCanvas(insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

Parameters
- insertShapeOptions ([Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions)) — Optional. The location and size of canvas. The default location and size is (0, 0, 300, 200).

Returns
- [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks
[ API set: WordApiDesktop 1.2 ]

### insertContentControl(contentControlType)
Wraps the Paragraph object with a content control.

```typescript
insertContentControl(contentControlType?: Word.ContentControlType.richText | Word.ContentControlType.plainText | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.comboBox | "RichText" | "PlainText" | "CheckBox" | "DropDownList" | "ComboBox"): Word.ContentControl;
```

Parameters
- contentControlType ([Word.ContentControlType.richText](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-richtext-member) | [plainText](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-plaintext-member) | [checkBox](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-checkbox-member) | [dropDownList](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-dropdownlist-member) | [comboBox](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-combobox-member) | string literals) — Optional. Content control type to insert. Default is 'RichText'.

Returns
- [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks
[ API set: WordApi 1.1 ]

Note: The contentControlType parameter was introduced in WordApi 1.5. PlainText support was added in 1.5. CheckBox support was added in 1.7. DropDownList and ComboBox support was added in 1.9.

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    paragraphs.load({select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a command to get the first paragraph.
    const paragraph = paragraphs.items[0];

    // Queue a command to wrap the first paragraph in a rich text content control.
    paragraph.insertContentControl();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Wrapped the first paragraph in a content control.');
});
```

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-content-controls.yaml

// Traverses each paragraph of the document and wraps a content control on each with either a even or odd tags.
await Word.run(async (context) => {
  let paragraphs = context.document.body.paragraphs;
  paragraphs.load("$none"); // Don't need any properties; just wrap each paragraph with a content control.

  await context.sync();

  for (let i = 0; i < paragraphs.items.length; i++) {
    let contentControl = paragraphs.items[i].insertContentControl();
    // For even, tag "even".
    if (i % 2 === 0) {
      contentControl.tag = "even";
    } else {
      contentControl.tag = "odd";
    }
  }
  console.log("Content controls inserted: " + paragraphs.items.length);

  await context.sync();
});
```

### insertFileFromBase64(base64File, insertLocation)
Inserts a document into the paragraph at the specified location.

```typescript
insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
```

Parameters
- base64File (string) — Required. The Base64-encoded content of a .docx file.
- insertLocation ([Word.InsertLocation.replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | string literals) — Required. The value must be 'Replace', 'Start', or 'End'.

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
[ API set: WordApi 1.1 ]

Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or another option appropriate for your scenario.

### insertGeometricShape(geometricShapeType, insertShapeOptions)
Inserts a geometric shape in front of text with its anchor at the beginning of the paragraph.

```typescript
insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

Parameters
- geometricShapeType ([Word.GeometricShapeType](/en-us/javascript/api/word/word.geometricshapetype)) — The geometric type of the shape to insert.
- insertShapeOptions ([Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions)) — Optional. The location and size of the geometric shape. Default is (0, 0, 100, 100).

Returns
- [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks
[ API set: WordApiDesktop 1.2 ]

### insertGeometricShape(geometricShapeType, insertShapeOptions) (string literal overload)
Inserts a geometric shape in front of text with its anchor at the beginning of the paragraph.

```typescript
insertGeometricShape(geometricShapeType: "LineInverse" | "Triangle" | ... | "ChartPlus", insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

Parameters
- geometricShapeType (string literal union) — The geometric type of the shape to insert.
- insertShapeOptions ([Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions)) — Optional. The location and size of the geometric shape. Default is (0, 0, 100, 100).

Returns
- [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks
[ API set: WordApiDesktop 1.2 ]

### insertHtml(html, insertLocation)
Inserts HTML into the paragraph at the specified location.

```typescript
insertHtml(html: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
```

Parameters
- html (string) — Required. The HTML to be inserted in the paragraph.
- insertLocation ([Word.InsertLocation.replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | string literals) — Required. The value must be 'Replace', 'Start', or 'End'.

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the style property for the top 2 paragraphs.
    // We never perform an empty load. We always must request a property.
    paragraphs.load({select: 'style', top: 2} );

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a command to get the first paragraph.
    const paragraph = paragraphs.items[0];

    // Queue a command to insert HTML content at the end of the first paragraph.
    paragraph.insertHtml('<strong>Inserted HTML.</strong>', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Inserted HTML content at the end of the first paragraph.');
});
```

### insertInlinePictureFromBase64(base64EncodedImage, insertLocation)
Inserts a picture into the paragraph at the specified location.

```typescript
insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.InlinePicture;
```

Parameters
- base64EncodedImage (string) — Required. The Base64-encoded image to be inserted.
- insertLocation ([Word.InsertLocation.replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | string literals) — Required. The value must be 'Replace', 'Start', or 'End'.

Returns
- [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the paragraphs collection.
    const paragraphs = context.document.body.paragraphs;

    // Queue a command to load the style property for all of the paragraphs.
    paragraphs.load('style');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();

    // Queue a command to get the first paragraph.
    const paragraph = paragraphs.items[0];

    const b64encodedImg = "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

    // Queue a command to insert a base64 encoded image at the beginning of the first paragraph.
    paragraph.insertInlinePictureFromBase64(b64encodedImg, Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Added an image to the first paragraph.');
});
```

### insertOoxml(ooxml, insertLocation)
Inserts OOXML into the paragraph at the specified location.

```typescript
insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
```

Parameters
- ooxml (string) — Required. The OOXML to be inserted in the paragraph.
- insertLocation ([Word.InsertLocation.replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | string literals) — Required. The value must be 'Replace', 'Start', or 'End'.

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
[ API set: WordApi 1.1 ]

### insertParagraph(paragraphText, insertLocation)
Inserts a paragraph at the specified location.

```typescript
insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
```

Parameters
- paragraphText (string) — Required. The paragraph text to be inserted.
- insertLocation ([Word.InsertLocation.before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) | [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) | string literals) — Required. The value must be 'Before' or 'After'.

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks
[ API set: WordApi 1.1 ]

### insertPictureFromBase64(base64EncodedImage, insertShapeOptions)
Inserts a floating picture in front of text with its anchor at the beginning of the paragraph.

```typescript
insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

Parameters
- base64EncodedImage (string) — Required. The Base64-encoded image to be inserted.
- insertShapeOptions ([Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions)) — Optional. The location and size of the picture. Default location is (0, 0) and default size is the image's original size.

Returns
- [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks
[ API set: WordApiDesktop 1.2 ]

### insertTable(rowCount, columnCount, insertLocation, values)
Inserts a table with the specified number of rows and columns.

```typescript
insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After", values?: string[][]): Word.Table;
```

Parameters
- rowCount (number) — Required. The number of rows in the table.
- columnCount (number) — Required. The number of columns in the table.
- insertLocation ([Word.InsertLocation.before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) | [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) | string literals) — Required. The value must be 'Before' or 'After'.
- values (string[][]) — Optional 2D array. Cells are filled if corresponding strings are specified.

Returns
- [Word.Table](/en-us/javascript/api/word/word.table)

Remarks
[ API set: WordApi 1.3 ]

### insertText(text, insertLocation)
Inserts text into the paragraph at the specified location.

```typescript
insertText(text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
```

Parameters
- text (string) — Required. Text to be inserted.
- insertLocation ([Word.InsertLocation.replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | string literals) — Required. The value must be 'Replace', 'Start', or 'End'.

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-in-different-locations.yaml

await Word.run(async (context) => {
  // Replace the last paragraph.
  const range: Word.Range = context.document.body.paragraphs.getLast().insertText("Just replaced the last paragraph!", "Replace");
  range.font.highlightColor = "black";
  range.font.color = "white";

  await context.sync();
});
```

### insertTextBox(text, insertShapeOptions)
Inserts a floating text box in front of text with its anchor at the beginning of the paragraph.

```typescript
insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

Parameters
- text (string) — Optional. The text to insert into the text box.
- insertShapeOptions ([Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions)) — Optional. The location and size of the text box. Default is (0, 0, 100, 100).

Returns
- [Word.Shape](/en-us/javascript/api/word/word.shape)

Remarks
[ API set: WordApiDesktop 1.2 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Inserts a text box at the beginning of the first paragraph in header.
  const headerFooterBody: Word.Body = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
  headerFooterBody.load("paragraphs");
  const firstParagraph: Word.Paragraph = headerFooterBody.paragraphs.getFirst();
  const insertShapeOptions: Word.InsertShapeOptions = {
    top: 0,
    left: 0,
    height: 100,
    width: 100
  };
  const newTextBox: Word.Shape = firstParagraph.insertTextBox("placeholder text", insertShapeOptions);
  newTextBox.select();
  await context.sync();

  console.log("Inserted a text box at the beginning of the first paragraph in the header.");
});
```

### joinList()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Joins a list paragraph with the closest list above or below this paragraph.

```typescript
joinList(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ParagraphLoadOptions): Word.Paragraph;
```

Parameters
- options ([Word.Interfaces.ParagraphLoadOptions](/en-us/javascript/api/word/word.interfaces.paragraphloadoptions)) — Provides options for which properties of the object to load.

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Paragraph;
```

Parameters
- propertyNames (string | string[]) — A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Paragraph;
```

Parameters
- propertyNamesAndPaths ({ select?: string; expand?: string; }) — select is a comma-delimited string of properties; expand is a comma-delimited string of navigation properties.

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

### next(count)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Paragraph object that represents the next paragraph.

```typescript
next(count: number): Word.Paragraph;
```

Parameters
- count (number) — Optional. The number of paragraphs to move forward.

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### openOrCloseUp()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Toggles the spacing before the paragraph.

```typescript
openOrCloseUp(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### openUp()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets spacing before the paragraph to 12 points.

```typescript
openUp(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### outdent()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes one level of indent for the paragraph.

```typescript
outdent(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### outlineDemote()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Applies the next heading level style (Heading 1 through Heading 8) to the paragraph.

```typescript
outlineDemote(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### outlineDemoteToBody()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Demotes the paragraph to body text by applying the Normal style.

```typescript
outlineDemoteToBody(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### outlinePromote()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Applies the previous heading level style (Heading 1 through Heading 8) to the paragraph.

```typescript
outlinePromote(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### previous(count)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the previous paragraph as a Paragraph object.

```typescript
previous(count: number): Word.Paragraph;
```

Parameters
- count (number) — Optional. The number of paragraphs to move backward.

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### reset()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Removes manual paragraph formatting (formatting not applied using a style).

```typescript
reset(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### resetAdvanceTo()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Resets the paragraph that uses custom list levels to the original level settings.

```typescript
resetAdvanceTo(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### search(searchText, searchOptions)
Performs a search with the specified SearchOptions on the scope of the paragraph object. The search results are a collection of range objects.

```typescript
search(searchText: string, searchOptions?: Word.SearchOptions | {
            ignorePunct?: boolean;
            ignoreSpace?: boolean;
            matchCase?: boolean;
            matchPrefix?: boolean;
            matchSuffix?: boolean;
            matchWholeWord?: boolean;
            matchWildcards?: boolean;
        }): Word.RangeCollection;
```

Parameters
- searchText (string) — Required. The search text.
- searchOptions ([Word.SearchOptions](/en-us/javascript/api/word/word.searchoptions) | { ignorePunct?: boolean; ignoreSpace?: boolean; matchCase?: boolean; matchPrefix?: boolean; matchSuffix?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean; }) — Optional. Options for the search.

Returns
- [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

Remarks
[ API set: WordApi 1.1 ]

### select(selectionMode)
Selects and navigates the Word UI to the paragraph.

```typescript
select(selectionMode?: Word.SelectionMode): void;
```

Parameters
- selectionMode ([Word.SelectionMode](/en-us/javascript/api/word/word.selectionmode)) — Optional. Must be 'Select', 'Start', or 'End'. 'Select' is default.

Returns
- void

Remarks
[ API set: WordApi 1.1 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/scroll-to-range.yaml

await Word.run(async (context) => {
  // If select is called with no parameters, it selects the object.
  context.document.body.paragraphs.getLast().select();

  await context.sync();
});

...

await Word.run(async (context) => {
  // Select can be at the start or end of a range; this by definition moves the insertion point without selecting the range.
  context.document.body.paragraphs.getLast().select(Word.SelectionMode.end);

  await context.sync();
});
```

### select(selectionMode) (string literal overload)
Selects and navigates the Word UI to the paragraph.

```typescript
select(selectionMode?: "Select" | "Start" | "End"): void;
```

Parameters
- selectionMode ("Select" | "Start" | "End") — Optional. The selection mode. 'Select' is the default.

Returns
- void

Remarks
[ API set: WordApi 1.1 ]

### selectNumber()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Selects the number or bullet in a list.

```typescript
selectNumber(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### separateList()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Separates a list into two separate lists. For numbered lists, the new list restarts numbering at the starting number, usually 1.

```typescript
separateList(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ParagraphUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties ([Word.Interfaces.ParagraphUpdateData](/en-us/javascript/api/word/word.interfaces.paragraphupdatedata)) — A JavaScript object isomorphic to the properties of the object.
- options ([OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)) — Option to suppress errors if setting any read-only properties.

Returns
- void

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/90-scenarios/multiple-property-set.yaml

await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.body.paragraphs.getFirst();
  paragraph.set({
    leftIndent: 30,
    font: {
      bold: true,
      color: "red"
    }
  });

  await context.sync();
});

...

await Word.run(async (context) => {
  const firstParagraph: Word.Paragraph = context.document.body.paragraphs.getFirst();
  const secondParagraph: Word.Paragraph = firstParagraph.getNext();
  firstParagraph.load("text, font/color, font/bold, leftIndent");

  await context.sync();

  secondParagraph.set(firstParagraph);

  await context.sync();
});
```

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Paragraph): void;
```

Parameters
- properties ([Word.Paragraph](/en-us/javascript/api/word/word.paragraph))

Returns
- void

### space1()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the paragraph to single spacing.

```typescript
space1(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### space1Pt5()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the paragraph to 1.5-line spacing.

```typescript
space1Pt5(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### space2()
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the paragraph to double spacing.

```typescript
space2(): void;
```

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### split(delimiters, trimDelimiters, trimSpacing)
Splits the paragraph into child ranges by using delimiters.

```typescript
split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
```

Parameters
- delimiters (string[]) — Required. The delimiters as an array of strings.
- trimDelimiters (boolean) — Optional. Whether to trim delimiters from the ranges. Default is false.
- trimSpacing (boolean) — Optional. Whether to trim spacing characters from the start and end of the ranges. Default is false.

Returns
- [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

Remarks
[ API set: WordApi 1.3 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/split-words-of-first-paragraph.yaml

await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.body.paragraphs.getFirst();
  const words = paragraph.split([" "], true /* trimDelimiters*/, true /* trimSpaces */);
  words.load("text");

  await context.sync();

  for (let i = 0; i < words.items.length; i++) {
    if (i >= 1) {
      words.items[i - 1].font.highlightColor = "#FFFFFF";
    }
    words.items[i].font.highlightColor = "#FFFF00";

    await context.sync();
    await pause(200);
  }
});
```

### startNewList()
Starts a new list with this paragraph. Fails if the paragraph is already a list item.

```typescript
startNewList(): Word.List;
```

Returns
- [Word.List](/en-us/javascript/api/word/word.list)

Remarks
[ API set: WordApi 1.3 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/insert-list.yaml

// This example starts a new list with the second paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Start new list using the second paragraph.
  const list: Word.List = paragraphs.items[1].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set up list level for the list item.
  paragraph.listItem.level = 4;

  // To add paragraphs outside the list, use Before or After.
  list.insertParagraph("New paragraph goes after (not part of the list)", "After");

  await context.sync();
});
```

### tabHangingIndent(count)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets a hanging indent to a specified number of tab stops.

```typescript
tabHangingIndent(count: number): void;
```

Parameters
- count (number) — The number of tab stops for the hanging indent.

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### tabIndent(count)
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the left indent for the paragraph to a specified number of tab stops.

```typescript
tabIndent(count: number): void;
```

Parameters
- count (number) — The number of tab stops for the left indent.

Returns
- void

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Paragraph object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ParagraphData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ParagraphData;
```

Returns
- [Word.Interfaces.ParagraphData](/en-us/javascript/api/word/word.interfaces.paragraphdata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Paragraph;
```

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.Paragraph;
```

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

## Event Details

### onCommentAdded
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when new comments are added.

```typescript
readonly onCommentAdded: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### onCommentChanged
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment or its reply is changed.

```typescript
readonly onCommentChanged: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### onCommentDeleted
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when comments are deleted.

```typescript
readonly onCommentDeleted: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### onCommentDeselected
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment is deselected.

```typescript
readonly onCommentDeselected: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

### onCommentSelected
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment is selected.

```typescript
readonly onCommentSelected: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]