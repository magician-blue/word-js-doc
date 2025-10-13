# Word.ContentControl class

Package: word

Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.

Extends: OfficeExtension.ClientObject

## Remarks

[API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;

    // Queue a command to load the id property for all of the content controls.
    contentControls.load('id');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to load the properties on the first content control.
        contentControls.items[0].load(  'appearance,' +
                                        'cannotDelete,' +
                                        'cannotEdit,' +
                                        'color,' +
                                        'id,' +
                                        'placeHolderText,' +
                                        'removeWhenEdited,' +
                                        'title,' +
                                        'text,' +
                                        'type,' +
                                        'style,' +
                                        'tag,' +
                                        'font/size,' +
                                        'font/name,' +
                                        'font/color');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Property values of the first content control:' +
            '   ----- appearance: ' + contentControls.items[0].appearance +
            '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
            '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
            '   ----- color: ' + contentControls.items[0].color +
            '   ----- id: ' + contentControls.items[0].id +
            '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
            '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
            '   ----- title: ' + contentControls.items[0].title +
            '   ----- text: ' + contentControls.items[0].text +
            '   ----- type: ' + contentControls.items[0].type +
            '   ----- style: ' + contentControls.items[0].style +
            '   ----- tag: ' + contentControls.items[0].tag +
            '   ----- font size: ' + contentControls.items[0].font.size +
            '   ----- font name: ' + contentControls.items[0].font.name +
            '   ----- font color: ' + contentControls.items[0].font.color);
    }
});
```

## Properties

- appearance — Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.
- buildingBlockGalleryContentControl — Gets the building block gallery-related data if the content control's Word.ContentControlType is BuildingBlockGallery. It's null otherwise.
- cannotDelete — Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.
- cannotEdit — Specifies a value that indicates whether the user can edit the contents of the content control.
- checkboxContentControl — Gets the data of the content control when its type is CheckBox. It's null otherwise.
- color — Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.
- comboBoxContentControl — Gets the data of the content control when its type is ComboBox. It's null otherwise.
- contentControls — Gets the collection of content control objects in the content control.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- datePickerContentControl — Gets the date picker-related data if the content control's Word.ContentControlType is DatePicker. It's null otherwise.
- dropDownListContentControl — Gets the data of the content control when its type is DropDownList. It's null otherwise.
- endnotes — Gets the collection of endnotes in the content control.
- fields — Gets the collection of field objects in the content control.
- font — Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.
- footnotes — Gets the collection of footnotes in the content control.
- groupContentControl — Gets the group-related data if the content control's Word.ContentControlType is Group. It's null otherwise.
- id — Gets an integer that represents the content control identifier.
- inlinePictures — Gets the collection of InlinePicture objects in the content control. The collection doesn't include floating images.
- lists — Gets the collection of list objects in the content control.
- paragraphs — Gets the collection of paragraph objects in the content control.
- parentBody — Gets the parent body of the content control.
- parentContentControl — Gets the content control that contains the content control. Throws an ItemNotFound error if there isn't a parent content control.
- parentContentControlOrNullObject — Gets the content control that contains the content control. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- parentTable — Gets the table that contains the content control. Throws an ItemNotFound error if it isn't contained in a table.
- parentTableCell — Gets the table cell that contains the content control. Throws an ItemNotFound error if it isn't contained in a table cell.
- parentTableCellOrNullObject — Gets the table cell that contains the content control. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- parentTableOrNullObject — Gets the table that contains the content control. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.
- pictureContentControl — Gets the picture-related data if the content control's Word.ContentControlType is Picture. It's null otherwise.
- placeholderText — Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.
- removeWhenEdited — Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.
- repeatingSectionContentControl — Gets the repeating section-related data if the content control's Word.ContentControlType is RepeatingSection. It's null otherwise.
- style — Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- styleBuiltIn — Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- subtype — Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.
- tables — Gets the collection of table objects in the content control.
- tag — Specifies a tag to identify a content control.
- text — Gets the text of the content control.
- title — Specifies the title for a content control.
- type — Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.
- xmlMapping — Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

## Methods

- clear() — Clears the contents of the content control. The user can perform the undo operation on the cleared content.
- delete(keepContent) — Deletes the content control and its content. If keepContent is set to true, the content isn't deleted.
- getComments() — Gets comments associated with the content control.
- getContentControls(options) — Gets the currently supported child content controls in this content control.
- getHtml() — Gets an HTML representation of the content control object.
- getOoxml() — Gets the Office Open XML (OOXML) representation of the content control object.
- getRange(rangeLocation) — Gets the whole content control, or the starting or ending point of the content control, as a range.
- getReviewedText(changeTrackingVersion) — Gets reviewed text based on ChangeTrackingVersion selection.
- getTextRanges(endingMarks, trimSpacing) — Gets the text ranges in the content control by using punctuation marks and/or other ending marks.
- getTrackedChanges() — Gets the collection of the TrackedChange objects in the content control.
- insertBreak(breakType, insertLocation) — Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
- insertFileFromBase64(base64File, insertLocation) — Inserts a document into the content control at the specified location.
- insertHtml(html, insertLocation) — Inserts HTML into the content control at the specified location.
- insertInlinePictureFromBase64(base64EncodedImage, insertLocation) — Inserts an inline picture into the content control at the specified location.
- insertOoxml(ooxml, insertLocation) — Inserts OOXML into the content control at the specified location.
- insertParagraph(paragraphText, insertLocation) — Inserts a paragraph at the specified location.
- insertTable(rowCount, columnCount, insertLocation, values) — Inserts a table with the specified number of rows and columns into, or next to, a content control.
- insertText(text, insertLocation) — Inserts text into the content control at the specified location.
- load(...) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- resetState() — Resets the state of the content control.
- search(searchText, searchOptions) — Performs a search with the specified SearchOptions on the scope of the content control object. The search results are a collection of range objects.
- select(selectionMode) — Selects the content control. This causes Word to scroll to the selection.
- set(...) — Sets multiple properties of an object at the same time.
- setState(contentControlState) — Sets the state of the content control.
- split(delimiters, multiParagraphs, trimDelimiters, trimSpacing) — Splits the content control into child ranges by using delimiters.
- toJSON() — Returns a plain JavaScript object with shallow copies of any loaded child properties.
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

## Events

- onCommentAdded — Occurs when new comments are added.
- onCommentChanged — Occurs when a comment or its reply is changed.
- onCommentDeselected — Occurs when a comment is deselected.
- onCommentSelected — Occurs when a comment is selected.
- onDataChanged — Occurs when data within the content control are changed. To get the new text, load this content control in the handler. To get the old text, do not load it.
- onDeleted — Occurs when the content control is deleted. Do not load this content control in the handler, otherwise you won't be able to get its original properties.
- onEntered — Occurs when the content control is entered.
- onExited — Occurs when the content control is exited, for example, when the cursor leaves the content control.
- onSelectionChanged — Occurs when selection within the content control is changed.

## Property Details

### appearance

Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

```ts
appearance: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property Value: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"

Remarks: [API set: WordApi 1.1]

---

### buildingBlockGalleryContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the building block gallery-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is BuildingBlockGallery. It's null otherwise.

```ts
readonly buildingBlockGalleryContentControl: Word.BuildingBlockGalleryContentControl;
```

Property Value: [Word.BuildingBlockGalleryContentControl](/en-us/javascript/api/word/word.buildingblockgallerycontentcontrol)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### cannotDelete

Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

```ts
cannotDelete: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1]

---

### cannotEdit

Specifies a value that indicates whether the user can edit the contents of the content control.

```ts
cannotEdit: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1]

---

### checkboxContentControl

Gets the data of the content control when its type is CheckBox. It's null otherwise.

```ts
readonly checkboxContentControl: Word.CheckboxContentControl;
```

Property Value: [Word.CheckboxContentControl](/en-us/javascript/api/word/word.checkboxcontentcontrol)

Remarks: [API set: WordApi 1.7]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml

// Toggles the isChecked property of the first checkbox content control found in the selection.
await Word.run(async (context) => {
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.checkBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,checkboxContentControl/isChecked");

  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,checkboxContentControl/isChecked");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.checkBox) {
      console.warn("No checkbox content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  const isCheckedBefore = selectedContentControl.checkboxContentControl.isChecked;
  console.log("isChecked state before:", `id: ${selectedContentControl.id} ... isChecked: ${isCheckedBefore}`);
  selectedContentControl.checkboxContentControl.isChecked = !isCheckedBefore;
  selectedContentControl.load("id,checkboxContentControl/isChecked");
  await context.sync();

  console.log(
    "isChecked state after:",
    `id: ${selectedContentControl.id} ... isChecked: ${selectedContentControl.checkboxContentControl.isChecked}`
  );
});
```

---

### color

Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

```ts
color: string;
```

Property Value: string

Remarks: [API set: WordApi 1.1]

---

### comboBoxContentControl

Gets the data of the content control when its type is ComboBox. It's null otherwise.

```ts
readonly comboBoxContentControl: Word.ComboBoxContentControl;
```

Property Value: [Word.ComboBoxContentControl](/en-us/javascript/api/word/word.comboboxcontentcontrol)

Remarks: [API set: WordApi 1.9]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-combo-box-content-control.yaml

// Adds the provided list item to the first combo box content control in the selection.
await Word.run(async (context) => {
  const listItemText = (document.getElementById("item-to-add") as HTMLInputElement).value.trim();
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.comboBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,comboBoxContentControl");
  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,comboBoxContentControl");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.comboBox) {
      console.warn("No combo box content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  selectedContentControl.comboBoxContentControl.addListItem(listItemText);
  await context.sync();

  console.log(`List item added to control with ID ${selectedContentControl.id}: ${listItemText}`);
});
```

---

### contentControls

Gets the collection of content control objects in the content control.

```ts
readonly contentControls: Word.ContentControlCollection;
```

Property Value: [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks: [API set: WordApi 1.1]

---

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```ts
context: RequestContext;
```

Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### datePickerContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the date picker-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is DatePicker. It's null otherwise.

```ts
readonly datePickerContentControl: Word.DatePickerContentControl;
```

Property Value: [Word.DatePickerContentControl](/en-us/javascript/api/word/word.datepickercontentcontrol)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### dropDownListContentControl

Gets the data of the content control when its type is DropDownList. It's null otherwise.

```ts
readonly dropDownListContentControl: Word.DropDownListContentControl;
```

Property Value: [Word.DropDownListContentControl](/en-us/javascript/api/word/word.dropdownlistcontentcontrol)

Remarks: [API set: WordApi 1.9]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-dropdown-list-content-control.yaml

// Adds the provided list item to the first dropdown list content control in the selection.
await Word.run(async (context) => {
  const listItemText = (document.getElementById("item-to-add") as HTMLInputElement).value.trim();
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.dropDownList]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,dropDownListContentControl");
  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,dropDownListContentControl");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.dropDownList) {
      console.warn("No dropdown list content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  selectedContentControl.dropDownListContentControl.addListItem(listItemText);
  await context.sync();

  console.log(`List item added to control with ID ${selectedContentControl.id}: ${listItemText}`);
});
```

---

### endnotes

Gets the collection of endnotes in the content control.

```ts
readonly endnotes: Word.NoteItemCollection;
```

Property Value: [Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

Remarks: [API set: WordApi 1.5]

---

### fields

Gets the collection of field objects in the content control.

```ts
readonly fields: Word.FieldCollection;
```

Property Value: [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

Remarks: [API set: WordApi 1.4]

---

### font

Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

```ts
readonly font: Word.Font;
```

Property Value: [Word.Font](/en-us/javascript/api/word/word.font)

Remarks: [API set: WordApi 1.1]

---

### footnotes

Gets the collection of footnotes in the content control.

```ts
readonly footnotes: Word.NoteItemCollection;
```

Property Value: [Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

Remarks: [API set: WordApi 1.5]

---

### groupContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the group-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is Group. It's null otherwise.

```ts
readonly groupContentControl: Word.GroupContentControl;
```

Property Value: [Word.GroupContentControl](/en-us/javascript/api/word/word.groupcontentcontrol)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### id

Gets an integer that represents the content control identifier.

```ts
readonly id: number;
```

Property Value: number

Remarks: [API set: WordApi 1.1]

---

### inlinePictures

Gets the collection of InlinePicture objects in the content control. The collection doesn't include floating images.

```ts
readonly inlinePictures: Word.InlinePictureCollection;
```

Property Value: [Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)

Remarks: [API set: WordApi 1.1]

---

### lists

Gets the collection of list objects in the content control.

```ts
readonly lists: Word.ListCollection;
```

Property Value: [Word.ListCollection](/en-us/javascript/api/word/word.listcollection)

Remarks: [API set: WordApi 1.3]

---

### paragraphs

Gets the collection of paragraph objects in the content control.

```ts
readonly paragraphs: Word.ParagraphCollection;
```

Property Value: [Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

Remarks: [API set: WordApi 1.1]

Important: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this content control aren't returned. From requirement set 1.3, paragraphs in such tables are also returned.

---

### parentBody

Gets the parent body of the content control.

```ts
readonly parentBody: Word.Body;
```

Property Value: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks: [API set: WordApi 1.3]

---

### parentContentControl

Gets the content control that contains the content control. Throws an ItemNotFound error if there isn't a parent content control.

```ts
readonly parentContentControl: Word.ContentControl;
```

Property Value: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks: [API set: WordApi 1.1]

---

### parentContentControlOrNullObject

Gets the content control that contains the content control. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```ts
readonly parentContentControlOrNullObject: Word.ContentControl;
```

Property Value: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks: [API set: WordApi 1.3]

---

### parentTable

Gets the table that contains the content control. Throws an ItemNotFound error if it isn't contained in a table.

```ts
readonly parentTable: Word.Table;
```

Property Value: [Word.Table](/en-us/javascript/api/word/word.table)

Remarks: [API set: WordApi 1.3]

---

### parentTableCell

Gets the table cell that contains the content control. Throws an ItemNotFound error if it isn't contained in a table cell.

```ts
readonly parentTableCell: Word.TableCell;
```

Property Value: [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Remarks: [API set: WordApi 1.3]

---

### parentTableCellOrNullObject

Gets the table cell that contains the content control. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```ts
readonly parentTableCellOrNullObject: Word.TableCell;
```

Property Value: [Word.TableCell](/en-us/javascript/api/word/word.tablecell)

Remarks: [API set: WordApi 1.3]

---

### parentTableOrNullObject

Gets the table that contains the content control. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```ts
readonly parentTableOrNullObject: Word.Table;
```

Property Value: [Word.Table](/en-us/javascript/api/word/word.table)

Remarks: [API set: WordApi 1.3]

---

### pictureContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the picture-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is Picture. It's null otherwise.

```ts
readonly pictureContentControl: Word.PictureContentControl;
```

Property Value: [Word.PictureContentControl](/en-us/javascript/api/word/word.picturecontentcontrol)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### placeholderText

Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

```ts
placeholderText: string;
```

Property Value: string

Remarks: [API set: WordApi 1.1]

---

### removeWhenEdited

Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

```ts
removeWhenEdited: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1]

---

### repeatingSectionContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the repeating section-related data if the content control's [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) is RepeatingSection. It's null otherwise.

```ts
readonly repeatingSectionContentControl: Word.RepeatingSectionContentControl;
```

Property Value: [Word.RepeatingSectionContentControl](/en-us/javascript/api/word/word.repeatingsectioncontentcontrol)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### style

Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```ts
style: string;
```

Property Value: string

Remarks: [API set: WordApi 1.1]

---

### styleBuiltIn

Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```ts
styleBuiltIn: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

Property Value: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"

Remarks: [API set: WordApi 1.3]

---

### subtype

Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.

```ts
readonly subtype: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group";
```

Property Value: [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group"

Remarks: [API set: WordApi 1.3]

---

### tables

Gets the collection of table objects in the content control.

```ts
readonly tables: Word.TableCollection;
```

Property Value: [Word.TableCollection](/en-us/javascript/api/word/word.tablecollection)

Remarks: [API set: WordApi 1.3]

---

### tag

Specifies a tag to identify a content control.

```ts
tag: string;
```

Property Value: string

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
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

---

### text

Gets the text of the content control.

```ts
readonly text: string;
```

Property Value: string

Remarks: [API set: WordApi 1.1]

---

### title

Specifies the title for a content control.

```ts
title: string;
```

Property Value: string

Remarks: [API set: WordApi 1.1]

---

### type

Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.

```ts
readonly type: Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group";
```

Property Value: [Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype) | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group"

Remarks: [API set: WordApi 1.1]

---

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

```ts
readonly xmlMapping: Word.XmlMapping;
```

Property Value: [Word.XmlMapping](/en-us/javascript/api/word/word.xmlmapping)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

## Method Details

### clear()

Clears the contents of the content control. The user can perform the undo operation on the cleared content.

```ts
clear(): void;
```

Returns: void

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
        
    if (contentControls.items.length === 0) {
        console.log("There isn't a content control in this document.");
    } else {
        // Queue a command to clear the contents of the first content control.
        contentControls.items[0].clear();

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Content control cleared of contents.');
    }
});
```

---

### delete(keepContent)

Deletes the content control and its content. If keepContent is set to true, the content isn't deleted.

```ts
delete(keepContent: boolean): void;
```

Parameters:
- keepContent (boolean) — Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content isn't deleted.

Returns: void

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the content controls collection.
    contentControls.load('text');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
        
    if (contentControls.items.length === 0) {
        console.log("There isn't a content control in this document.");
    } else {            
        // Queue a command to delete the first content control. 
        // The contents will remain in the document.
        contentControls.items[0].delete(true);

        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Content control cleared of contents.'); 
    }
});
```

```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-ondeleted-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls.getByTag("forTesting");
  contentControls.load("items");
  await context.sync();

  if (contentControls.items.length === 0) {
    console.log("There are no content controls in this document.");
  } else {
    console.log("Control to be deleted:", contentControls.items[0]);
    contentControls.items[0].delete(false);
    await context.sync();
  }
});
```

---

### getComments()

Gets comments associated with the content control.

```ts
getComments(): Word.CommentCollection;
```

Returns: [Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)

Remarks: [API set: WordApi 1.4]

---

### getContentControls(options)

Gets the currently supported child content controls in this content control.

```ts
getContentControls(options?: Word.ContentControlOptions): Word.ContentControlCollection;
```

Parameters:
- options ([Word.ContentControlOptions](/en-us/javascript/api/word/word.contentcontroloptions)) — Optional. Options that define which content controls are returned.

Returns: [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks: [API set: WordApi 1.5]

Important: If specific types are provided in the options parameter, only content controls of supported types are returned. Be aware that an exception will be thrown on using methods of a generic [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) that aren't relevant for the specific type. With time, additional types of content controls may be supported. Therefore, your add-in should request and handle specific types of content controls.

---

### getHtml()

Gets an HTML representation of the content control object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `ContentControl.getOoxml()` and convert the returned XML to HTML.

```ts
getHtml(): OfficeExtension.ClientResult<string>;
```

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection that contains a specific tag.
    const contentControlsWithTag = context.document.contentControls.getByTag('Customer-Address');
    
    // Queue a command to load the tag property for all of content controls.
    contentControlsWithTag.load('tag');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControlsWithTag.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to get the HTML contents of the first content control.
        const html = contentControlsWithTag.items[0].getHtml();
    
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Content control HTML: ' + html.value);
    }
});
```

---

### getOoxml()

Gets the Office Open XML (OOXML) representation of the content control object.

```ts
getOoxml(): OfficeExtension.ClientResult<string>;
```

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls.
    contentControls.load('id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to get the OOXML contents of the first content control.
        const ooxml = contentControls.items[0].getOoxml();
    
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Content control OOXML: ' + ooxml.value);
    }
});
```

---

### getRange(rangeLocation)

Gets the whole content control, or the starting or ending point of the content control, as a range.

```ts
getRange(rangeLocation?: Word.RangeLocation | "Whole" | "Start" | "End" | "Before" | "After" | "Content"): Word.Range;
```

Parameters:
- rangeLocation ([Word.RangeLocation](/en-us/javascript/api/word/word.rangelocation) | "Whole" | "Start" | "End" | "Before" | "After" | "Content") — Optional. The range location must be 'Whole', 'Start', 'End', 'Before', 'After', or 'Content'.

Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks: [API set: WordApi 1.3]

---

### getReviewedText(changeTrackingVersion) — overload 1

Gets reviewed text based on ChangeTrackingVersion selection.

```ts
getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion): OfficeExtension.ClientResult<string>;
```

Parameters:
- changeTrackingVersion ([Word.ChangeTrackingVersion](/en-us/javascript/api/word/word.changetrackingversion)) — Optional. The value must be 'Original' or 'Current'. The default is 'Current'.

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks: [API set: WordApi 1.4]

---

### getReviewedText(changeTrackingVersion) — overload 2

Gets reviewed text based on ChangeTrackingVersion selection.

```ts
getReviewedText(changeTrackingVersion?: "Original" | "Current"): OfficeExtension.ClientResult<string>;
```

Parameters:
- changeTrackingVersion ("Original" | "Current") — Optional. The value must be 'Original' or 'Current'. The default is 'Current'.

Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks: [API set: WordApi 1.4]

---

### getTextRanges(endingMarks, trimSpacing)

Gets the text ranges in the content control by using punctuation marks and/or other ending marks.

```ts
getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
```

Parameters:
- endingMarks (string[]) — Required. The punctuation marks and/or other ending marks as an array of strings.
- trimSpacing (boolean) — Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.

Returns: [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

Remarks: [API set: WordApi 1.3]

---

### getTrackedChanges()

Gets the collection of the TrackedChange objects in the content control.

```ts
getTrackedChanges(): Word.TrackedChangeCollection;
```

Returns: [Word.TrackedChangeCollection](/en-us/javascript/api/word/word.trackedchangecollection)

Remarks: [API set: WordApi 1.6]

---

### insertBreak(breakType, insertLocation)

Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.

```ts
insertBreak(
  breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line",
  insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"
): void;
```

Parameters:
- breakType ([Word.BreakType](/en-us/javascript/api/word/word.breaktype) | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line") — Required. Type of break.
- insertLocation ([start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | [before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) | [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) | "Start" | "End" | "Before" | "After") — Required. The value must be 'Start', 'End', 'Before', or 'After'.

Returns: void

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of content controls.
    contentControls.load('id');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    // We now will have access to the content control collection.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to insert a page break after the first content control.
        contentControls.items[0].insertBreak(Word.BreakType.page, Word.InsertLocation.after);
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Inserted a page break after the first content control.');    
    }
});
```

---

### insertFileFromBase64(base64File, insertLocation)

Inserts a document into the content control at the specified location.

```ts
insertFileFromBase64(
  base64File: string,
  insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"
): Word.Range;
```

Parameters:
- base64File (string) — Required. The Base64-encoded content of a .docx file.
- insertLocation ([replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End") — Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks: [API set: WordApi 1.1]

Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.

---

### insertHtml(html, insertLocation)

Inserts HTML into the content control at the specified location.

```ts
insertHtml(
  html: string,
  insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"
): Word.Range;
```

Parameters:
- html (string) — Required. The HTML to be inserted in to the content control.
- insertLocation ([replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End") — Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls.
    contentControls.load('id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to put HTML into the contents of the first content control.
        contentControls.items[0].insertHtml(
            '<strong>HTML content inserted into the content control.</strong>',
            'Start');
    
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Inserted HTML in the first content control.');
    }
});
```

---

### insertInlinePictureFromBase64(base64EncodedImage, insertLocation)

Inserts an inline picture into the content control at the specified location.

```ts
insertInlinePictureFromBase64(
  base64EncodedImage: string,
  insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"
): Word.InlinePicture;
```

Parameters:
- base64EncodedImage (string) — Required. The Base64-encoded image to be inserted in the content control.
- insertLocation ([replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End") — Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

Returns: [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

Remarks: [API set: WordApi 1.2]

---

### insertOoxml(ooxml, insertLocation)

Inserts OOXML into the content control at the specified location.

```ts
insertOoxml(
  ooxml: string,
  insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"
): Word.Range;
```

Parameters:
- ooxml (string) — Required. The OOXML to be inserted in to the content control.
- insertLocation ([replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End") — Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls.
    contentControls.load('id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to put OOXML into the contents of the first content control.
        contentControls.items[0].insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", "End");
    
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Inserted OOXML in the first content control.');
    }
});  

// Read "Create better add-ins for Word with Office Open XML" for guidance on working with OOXML.
// https://learn.microsoft.com/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml
```

---

### insertParagraph(paragraphText, insertLocation)

Inserts a paragraph at the specified location.

```ts
insertParagraph(
  paragraphText: string,
  insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"
): Word.Paragraph;
```

Parameters:
- paragraphText (string) — Required. The paragraph text to be inserted.
- insertLocation ([start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | [before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) | [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) | "Start" | "End" | "Before" | "After") — Required. The value must be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.

Returns: [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls.
    contentControls.load('id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to insert a paragraph after the first content control.
        contentControls.items[0].insertParagraph('Text of the inserted paragraph.', 'After');
    
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Inserted a paragraph after the first content control.');
    }
});  
```

---

### insertTable(rowCount, columnCount, insertLocation, values)

Inserts a table with the specified number of rows and columns into, or next to, a content control.

```ts
insertTable(
  rowCount: number,
  columnCount: number,
  insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After",
  values?: string[][]
): Word.Table;
```

Parameters:
- rowCount (number) — Required. The number of rows in the table.
- columnCount (number) — Required. The number of columns in the table.
- insertLocation ([start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | [before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) | [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) | "Start" | "End" | "Before" | "After") — Required. The value must be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
- values (string[][]) — Optional 2D array. Cells are filled if the corresponding strings are specified in the array.

Returns: [Word.Table](/en-us/javascript/api/word/word.table)

Remarks: [API set: WordApi 1.3]

---

### insertText(text, insertLocation)

Inserts text into the content control at the specified location.

```ts
insertText(
  text: string,
  insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"
): Word.Range;
```

Parameters:
- text (string) — Required. The text to be inserted in to the content control.
- insertLocation ([replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End") — Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls.
    contentControls.load('id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to replace text in the first content control.
        contentControls.items[0].insertText('Replaced text in the first content control.', 'Replace');
    
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Replaced text in the first content control.');
    }
});  

// The Silly stories add-in sample shows how to use the insertText method.
// https://aka.ms/sillystorywordaddin
```

---

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```ts
load(options?: Word.Interfaces.ContentControlLoadOptions): Word.ContentControl;
```

Parameters:
- options ([Word.Interfaces.ContentControlLoadOptions](/en-us/javascript/api/word/word.interfaces.contentcontrolloadoptions)) — Provides options for which properties of the object to load.

Returns: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

#### Examples
```ts
// Load all of the content control properties
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls.
    contentControls.load('id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    } else {
        // Queue a command to load the properties on the first content control.
        contentControls.items[0].load(  'appearance,' +
                                        'cannotDelete,' +
                                        'cannotEdit,' +
                                        'id,' +
                                        'placeHolderText,' +
                                        'removeWhenEdited,' +
                                        'title,' +
                                        'text,' +
                                        'type,' +
                                        'style,' +
                                        'tag,' +
                                        'font/size,' +
                                        'font/name,' +
                                        'font/color');             
    
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Property values of the first content control:' + 
            '   ----- appearance: ' + contentControls.items[0].appearance + 
            '   ----- cannotDelete: ' + contentControls.items[0].cannotDelete +
            '   ----- cannotEdit: ' + contentControls.items[0].cannotEdit +
            '   ----- color: ' + contentControls.items[0].color +
            '   ----- id: ' + contentControls.items[0].id +
            '   ----- placeHolderText: ' + contentControls.items[0].placeholderText +
            '   ----- removeWhenEdited: ' + contentControls.items[0].removeWhenEdited +
            '   ----- title: ' + contentControls.items[0].title +
            '   ----- text: ' + contentControls.items[0].text +
            '   ----- type: ' + contentControls.items[0].type +
            '   ----- style: ' + contentControls.items[0].style +
            '   ----- tag: ' + contentControls.items[0].tag +
            '   ----- font size: ' + contentControls.items[0].font.size +
            '   ----- font name: ' + contentControls.items[0].font.name +
            '   ----- font color: ' + contentControls.items[0].font.color);
    }
});  
```

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```ts
load(propertyNames?: string | string[]): Word.ContentControl;
```

Parameters:
- propertyNames (string | string[]) — A comma-delimited string or an array of strings that specify the properties to load.

Returns: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```ts
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.ContentControl;
```

Parameters:
- propertyNamesAndPaths ({ select?: string; expand?: string; }) — `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

---

### resetState()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Resets the state of the content control.

```ts
resetState(): void;
```

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/insert-and-change-content-controls.yaml

// Resets the state of the first content control.
await Word.run(async (context) => {
  let firstContentControl = context.document.contentControls.getFirstOrNullObject();
  await context.sync();

  if (firstContentControl.isNullObject) {
    console.warn("There are no content controls in this document.");
    return;
  }

  firstContentControl.resetState();
  firstContentControl.load("id");
  await context.sync();

  console.log(`Reset state of first content control with ID: ${firstContentControl.id}`);
});
```

---

### search(searchText, searchOptions)

Performs a search with the specified SearchOptions on the scope of the content control object. The search results are a collection of range objects.

```ts
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

Parameters:
- searchText (string) — Required. The search text.
- searchOptions ([Word.SearchOptions](/en-us/javascript/api/word/word.searchoptions) | { ignorePunct?: boolean; ignoreSpace?: boolean; matchCase?: boolean; matchPrefix?: boolean; matchSuffix?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean; }) — Optional. Options for the search.

Returns: [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

Remarks: [API set: WordApi 1.1]

#### Examples
```ts
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the content controls collection.
    const contentControls = context.document.contentControls;
    
    // Queue a command to load the id property for all of the content controls.
    contentControls.load('id');
     
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (contentControls.items.length === 0) {
        console.log('No content control found.');
    }
    else {
        // Queue a command to select the first content control.
        contentControls.items[0].select();
    
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Selected the first content control.');
    }
});  
```

---

### select(selectionMode) — overload 1

Selects the content control. This causes Word to scroll to the selection.

```ts
select(selectionMode?: Word.SelectionMode): void;
```

Parameters:
- selectionMode ([Word.SelectionMode](/en-us/javascript/api/word/word.selectionmode)) — Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

Returns: void

Remarks: [API set: WordApi 1.1]

---

### select(selectionMode) — overload 2

Selects the content control. This causes Word to scroll to the selection.

```ts
select(selectionMode?: "Select" | "Start" | "End"): void;
```

Parameters:
- selectionMode ("Select" | "Start" | "End") — Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

Returns: void

Remarks: [API set: WordApi 1.1]

---

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```ts
set(properties: Interfaces.ContentControlUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties ([Word.Interfaces.ContentControlUpdateData](/en-us/javascript/api/word/word.interfaces.contentcontrolupdatedata)) — A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options ([OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)) — Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-content-controls.yaml

// Adds title and colors to odd and even content controls and changes their appearance.
await Word.run(async (context) => {
  // Get the complete sentence (as range) associated with the insertion point.
  let evenContentControls = context.document.contentControls.getByTag("even");
  let oddContentControls = context.document.contentControls.getByTag("odd");
  evenContentControls.load("length");
  oddContentControls.load("length");

  await context.sync();

  for (let i = 0; i < evenContentControls.items.length; i++) {
    // Change a few properties and append a paragraph.
    evenContentControls.items[i].set({
      color: "red",
      title: "Odd ContentControl #" + (i + 1),
      appearance: Word.ContentControlAppearance.tags
    });
    evenContentControls.items[i].insertParagraph("This is an odd content control", "End");
  }

  for (let j = 0; j < oddContentControls.items.length; j++) {
    // Change a few properties and append a paragraph.
    oddContentControls.items[j].set({
      color: "green",
      title: "Even ContentControl #" + (j + 1),
      appearance: "Tags"
    });
    oddContentControls.items[j].insertHtml("This is an <b>even</b> content control", "End");
  }

  await context.sync();
});
```

---

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```ts
set(properties: Word.ContentControl): void;
```

Parameters:
- properties ([Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol))

Returns: void

---

### setState(contentControlState) — overload 1

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the state of the content control.

```ts
setState(contentControlState: Word.ContentControlState): void;
```

Parameters:
- contentControlState ([Word.ContentControlState](/en-us/javascript/api/word/word.contentcontrolstate)) — State to be set.

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/insert-and-change-content-controls.yaml

// Sets the state of the first content control.
await Word.run(async (context) => {
  const state = ((document.getElementById("state-to-set") as HTMLSelectElement)
    .value as unknown) as Word.ContentControlState;
  let firstContentControl = context.document.contentControls.getFirstOrNullObject();
  await context.sync();

  if (firstContentControl.isNullObject) {
    console.warn("There are no content controls in this document.");
    return;
  }

  firstContentControl.setState(state);
  firstContentControl.load("id");
  await context.sync();

  console.log(`Set state of first content control with ID ${firstContentControl.id} to ${state}.`);
});
```

---

### setState(contentControlState) — overload 2

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the state of the content control.

```ts
setState(contentControlState: "Error" | "Warning"): void;
```

Parameters:
- contentControlState ("Error" | "Warning") — State to be set.

Returns: void

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### split(delimiters, multiParagraphs, trimDelimiters, trimSpacing)

Splits the content control into child ranges by using delimiters.

```ts
split(
  delimiters: string[],
  multiParagraphs?: boolean,
  trimDelimiters?: boolean,
  trimSpacing?: boolean
): Word.RangeCollection;
```

Parameters:
- delimiters (string[]) — Required. The delimiters as an array of strings.
- multiParagraphs (boolean) — Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
- trimDelimiters (boolean) — Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
- trimSpacing (boolean) — Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.

Returns: [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

Remarks: [API set: WordApi 1.3]

---

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlData`) that contains shallow copies of any loaded child properties from the original object.

```ts
toJSON(): Word.Interfaces.ContentControlData;
```

Returns: [Word.Interfaces.ContentControlData](/en-us/javascript/api/word/word.interfaces.contentcontroldata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```ts
track(): Word.ContentControl;
```

Returns: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```ts
untrack(): Word.ContentControl;
```

Returns: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

---

## Event Details

### onCommentAdded

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when new comments are added.

```ts
readonly onCommentAdded: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### onCommentChanged

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment or its reply is changed.

```ts
readonly onCommentChanged: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### onCommentDeselected

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment is deselected.

```ts
readonly onCommentDeselected: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### onCommentSelected

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment is selected.

```ts
readonly onCommentSelected: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks: [API set: WordApi BETA (PREVIEW ONLY)]

---

### onDataChanged

Occurs when data within the content control are changed. To get the new text, load this content control in the handler. To get the old text, do not load it.

```ts
readonly onDataChanged: OfficeExtension.EventHandlers<Word.ContentControlDataChangedEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ContentControlDataChangedEventArgs](/en-us/javascript/api/word/word.contentcontroldatachangedeventargs)>

Remarks: [API set: WordApi 1.5]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-ondatachanged-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  // Register the onDataChanged event handler on each content control.
  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onDataChanged.add(contentControlDataChanged);
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when data is changed in content controls.");
  }
});

...

async function contentControlDataChanged(event: Word.ContentControlDataChangedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls where data was changed:`, event.ids);
  });
}
```

---

### onDeleted

Occurs when the content control is deleted. Do not load this content control in the handler, otherwise you won't be able to get its original properties.

```ts
readonly onDeleted: OfficeExtension.EventHandlers<Word.ContentControlDeletedEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ContentControlDeletedEventArgs](/en-us/javascript/api/word/word.contentcontroldeletedeventargs)>

Remarks: [API set: WordApi 1.5]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-ondeleted-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  // Register the onDeleted event handler on each content control.
  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onDeleted.add(contentControlDeleted);
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when content controls are deleted.");
  }
});

...

async function contentControlDeleted(event: Word.ContentControlDeletedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls that were deleted:`, event.ids);
  });
}
```

---

### onEntered

Occurs when the content control is entered.

```ts
readonly onEntered: OfficeExtension.EventHandlers<Word.ContentControlEnteredEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ContentControlEnteredEventArgs](/en-us/javascript/api/word/word.contentcontrolenteredeventargs)>

Remarks: [API set: WordApi 1.5]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-onentered-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  // Register the onEntered event handler on each content control.
  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onEntered.add(contentControlEntered);
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when the cursor is placed in content controls.");
  }
});

...

async function contentControlEntered(event: Word.ContentControlEnteredEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. ID of content control that was entered: ${event.ids[0]}`);
  });
}
```

---

### onExited

Occurs when the content control is exited, for example, when the cursor leaves the content control.

```ts
readonly onExited: OfficeExtension.EventHandlers<Word.ContentControlExitedEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ContentControlExitedEventArgs](/en-us/javascript/api/word/word.contentcontrolexitedeventargs)>

Remarks: [API set: WordApi 1.5]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-onexited-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  // Register the onExited event handler on each content control.
  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onExited.add(contentControlExited);
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when the cursor is removed from within content controls.");
  }
});

...

async function contentControlExited(event: Word.ContentControlExitedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. ID of content control that was exited: ${event.ids[0]}`);
  });
}
```

---

### onSelectionChanged

Occurs when selection within the content control is changed.

```ts
readonly onSelectionChanged: OfficeExtension.EventHandlers<Word.ContentControlSelectionChangedEventArgs>;
```

Event Type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ContentControlSelectionChangedEventArgs](/en-us/javascript/api/word/word.contentcontrolselectionchangedeventargs)>

Remarks: [API set: WordApi 1.5]

#### Examples
```ts
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-onselectionchanged-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onSelectionChanged.add(contentControlSelectionChanged);
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when selections are changed in content controls.");
  }
});

...

async function contentControlSelectionChanged(event: Word.ContentControlSelectionChangedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls where selection was changed:`, event.ids);
  });
}
```