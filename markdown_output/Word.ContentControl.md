# ContentControl

**Package:** `word`

**API Set:** WordApi 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents a content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as images, tables, or paragraphs of formatted text. Currently, only rich text, plain text, checkbox, dropdown list, and combo box content controls are supported.

## Class Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### appearance

**Type:** `Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"`

**Since:** WordApi 1.1

Specifies the appearance of the content control. The value can be 'BoundingBox', 'Tags', or 'Hidden'.

#### Examples

**Example**: Change the appearance of a content control to show bounding box borders instead of tags

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Set the appearance to show a bounding box
    contentControl.appearance = Word.ContentControlAppearance.boundingBox;
    // Alternative: contentControl.appearance = "BoundingBox";
    
    await context.sync();
});
```

---

### buildingBlockGalleryContentControl

**Type:** `Word.BuildingBlockGalleryContentControl`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the building block gallery-related data if the content control's Word.ContentControlType is BuildingBlockGallery. It's null otherwise.

#### Examples

**Example**: Check if a content control is a building block gallery type and if so, set its category to "General"

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    await context.sync();
    
    if (contentControl.type === Word.ContentControlType.buildingBlockGallery) {
        const bbGallery = contentControl.buildingBlockGalleryContentControl;
        bbGallery.category = "General";
        
        await context.sync();
        console.log("Building block gallery category set to General");
    } else {
        console.log("Content control is not a building block gallery type");
    }
});
```

---

### cannotDelete

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the user can delete the content control. Mutually exclusive with removeWhenEdited.

#### Examples

**Example**: Protect a content control containing important legal text by preventing users from deleting it while still allowing them to edit its contents.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Prevent users from deleting this content control
    contentControl.cannotDelete = true;
    
    // Load and sync to apply the changes
    await context.sync();
    
    console.log("Content control is now protected from deletion");
});
```

---

### cannotEdit

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the user can edit the contents of the content control.

#### Examples

**Example**: Protect a content control containing a legal disclaimer by making it read-only so users cannot edit its contents.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Make the content control read-only
    contentControl.cannotEdit = true;
    
    await context.sync();
    
    console.log("Content control is now protected from editing");
});
```

---

### checkboxContentControl

**Type:** `Word.CheckboxContentControl`

**Since:** WordApi 1.7

Gets the data of the content control when its type is CheckBox. It's null otherwise.

#### Examples

**Example**: Toggles the isChecked property of the first checkbox content control found in the selection.

```typescript
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

**Type:** `string`

**Since:** WordApi 1.1

Specifies the color of the content control. Color is specified in '#RRGGBB' format or by using the color name.

#### Examples

**Example**: Set the color of the first content control in the document to blue

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Set the content control color to blue
    contentControl.color = "#0000FF";
    
    await context.sync();
});
```

---

### comboBoxContentControl

**Type:** `Word.ComboBoxContentControl`

**Since:** WordApi 1.9

Gets the data of the content control when its type is ComboBox. It's null otherwise.

#### Examples

**Example**: Adds the provided list item to the first combo box content control in the selection.

```typescript
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

**Type:** `Word.ContentControlCollection`

**Since:** WordApi 1.1

Gets the collection of content control objects in the content control.

#### Examples

**Example**: Find and highlight all nested content controls within a parent content control by setting their background color to yellow.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const parentContentControl = context.document.contentControls.getFirst();
    
    // Get all nested content controls within the parent
    const nestedContentControls = parentContentControl.contentControls;
    nestedContentControls.load("items");
    
    await context.sync();
    
    // Highlight each nested content control
    for (let i = 0; i < nestedContentControls.items.length; i++) {
        nestedContentControls.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
    
    console.log(`Found and highlighted ${nestedContentControls.items.length} nested content controls`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a content control to verify the connection and sync changes with the document

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("text");
    
    // Access the request context associated with the content control
    const requestContext = contentControl.context;
    
    // Use the context to sync and retrieve data
    await requestContext.sync();
    
    console.log("Content control text:", contentControl.text);
    console.log("Request context is connected:", requestContext !== null);
});
```

---

### datePickerContentControl

**Type:** `Word.DatePickerContentControl`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the date picker-related data if the content control's Word.ContentControlType is DatePicker. It's null otherwise.

#### Examples

**Example**: Get the selected date from a date picker content control and display it in the console, or show a message if the content control is not a date picker type.

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    await context.sync();
    
    // Access the datePickerContentControl property
    const datePickerCC = contentControl.datePickerContentControl;
    
    if (datePickerCC !== null) {
        datePickerCC.load("selectedDate");
        await context.sync();
        
        console.log("Selected date:", datePickerCC.selectedDate);
    } else {
        console.log("This content control is not a date picker type.");
    }
});
```

---

### dropDownListContentControl

**Type:** `Word.DropDownListContentControl`

**Since:** WordApi 1.9

Gets the data of the content control when its type is DropDownList. It's null otherwise.

#### Examples

**Example**: Adds the provided list item to the first dropdown list content control in the selection.

```typescript
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

**Type:** `Word.NoteItemCollection`

**Since:** WordApi 1.5

Gets the collection of endnotes in the content control.

#### Examples

**Example**: Count and display the number of endnotes within a specific content control in the document.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the endnotes collection from the content control
    const endnotes = contentControl.endnotes;
    endnotes.load("items");
    
    await context.sync();
    
    // Display the count of endnotes
    console.log(`Number of endnotes in content control: ${endnotes.items.length}`);
});
```

---

### fields

**Type:** `Word.FieldCollection`

**Since:** WordApi 1.4

Gets the collection of field objects in the content control.

#### Examples

**Example**: Get all fields within a content control and display their codes in the console.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the fields collection from the content control
    const fields = contentControl.fields;
    fields.load("items");
    
    await context.sync();
    
    // Log the number of fields and their codes
    console.log(`Number of fields in content control: ${fields.items.length}`);
    
    fields.items.forEach((field, index) => {
        field.load("code");
    });
    
    await context.sync();
    
    fields.items.forEach((field, index) => {
        console.log(`Field ${index + 1} code: ${field.code}`);
    });
});
```

---

### font

**Type:** `Word.Font`

**Since:** WordApi 1.1

Gets the text format of the content control. Use this to get and set font name, size, color, and other properties.

#### Examples

**Example**: Set the content control's font to Arial, size 14, bold, and blue color

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Access and configure the font properties
    contentControl.font.name = "Arial";
    contentControl.font.size = 14;
    contentControl.font.bold = true;
    contentControl.font.color = "blue";
    
    await context.sync();
});
```

---

### footnotes

**Type:** `Word.NoteItemCollection`

**Since:** WordApi 1.5

Gets the collection of footnotes in the content control.

#### Examples

**Example**: Count and display the number of footnotes within a specific content control in the document.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the footnotes collection from the content control
    const footnotes = contentControl.footnotes;
    footnotes.load("items");
    
    await context.sync();
    
    // Display the count of footnotes
    console.log(`Number of footnotes in content control: ${footnotes.items.length}`);
});
```

---

### groupContentControl

**Type:** `Word.GroupContentControl`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the group-related data if the content control's Word.ContentControlType is Group. It's null otherwise.

#### Examples

**Example**: Check if a content control is a group type and if so, access its group-related data to log the group's title.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    await context.sync();
    
    // Check if the content control is a group type
    if (contentControl.type === Word.ContentControlType.group) {
        // Access the group content control data
        const groupContentControl = contentControl.groupContentControl;
        groupContentControl.load("title");
        
        await context.sync();
        
        console.log("Group content control title: " + groupContentControl.title);
    } else {
        console.log("This content control is not a group type.");
    }
});
```

---

### id

**Type:** `number`

**Since:** WordApi 1.1

Gets an integer that represents the content control identifier.

#### Examples

**Example**: Find and highlight a content control by retrieving its ID and displaying it in the console, then change its background color to yellow.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Load the id property
    contentControl.load("id");
    
    await context.sync();
    
    // Display the content control ID
    console.log(`Content Control ID: ${contentControl.id}`);
    
    // Use the ID to identify and modify the content control
    contentControl.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### inlinePictures

**Type:** `Word.InlinePictureCollection`

**Since:** WordApi 1.1

Gets the collection of InlinePicture objects in the content control. The collection doesn't include floating images.

#### Examples

**Example**: Get all inline pictures from a content control and resize them to a width of 100 pixels

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    const inlinePictures = contentControl.inlinePictures;
    
    inlinePictures.load("items");
    await context.sync();
    
    for (let i = 0; i < inlinePictures.items.length; i++) {
        inlinePictures.items[i].width = 100;
    }
    
    await context.sync();
    console.log(`Resized ${inlinePictures.items.length} inline picture(s)`);
});
```

---

### lists

**Type:** `Word.ListCollection`

**Since:** WordApi 1.3

Gets the collection of list objects in the content control.

#### Examples

**Example**: Get all lists within a content control and display the count of items in each list to the console.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the collection of lists in the content control
    const lists = contentControl.lists;
    lists.load("items");
    
    await context.sync();
    
    // Iterate through each list and get item count
    for (let i = 0; i < lists.items.length; i++) {
        const list = lists.items[i];
        const listItems = list.paragraphs;
        listItems.load("items/length");
        
        await context.sync();
        
        console.log(`List ${i + 1} has ${listItems.items.length} items`);
    }
});
```

---

### paragraphs

**Type:** `Word.ParagraphCollection`

**Since:** WordApi 1.1

Gets the collection of paragraph objects in the content control.

#### Examples

**Example**: Highlight all paragraphs within a content control by setting their background color to yellow.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the paragraphs collection from the content control
    const paragraphs = contentControl.paragraphs;
    
    // Load the paragraphs
    paragraphs.load("items");
    
    await context.sync();
    
    // Set background color for each paragraph
    for (let i = 0; i < paragraphs.items.length; i++) {
        paragraphs.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### parentBody

**Type:** `Word.Body`

**Since:** WordApi 1.3

Gets the parent body of the content control.

#### Examples

**Example**: Highlight the parent body of a content control by applying a yellow background color to demonstrate the relationship between the content control and its containing body.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the parent body of the content control
    const parentBody = contentControl.parentBody;
    
    // Apply yellow shading to the parent body to highlight it
    parentBody.font.highlightColor = "yellow";
    
    await context.sync();
    
    console.log("Parent body of the content control has been highlighted");
});
```

---

### parentContentControl

**Type:** `Word.ContentControl`

**Since:** WordApi 1.1

Gets the content control that contains the content control. Throws an ItemNotFound error if there isn't a parent content control.

#### Examples

**Example**: Get the parent content control of a nested content control and change its title to "Parent Container"

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document (assumed to be nested)
    const contentControl = context.document.contentControls.getFirst();
    
    // Get its parent content control
    const parentContentControl = contentControl.parentContentControl;
    parentContentControl.load("title");
    
    await context.sync();
    
    // Update the parent's title
    parentContentControl.title = "Parent Container";
    
    await context.sync();
    
    console.log("Parent content control title updated");
});
```

---

### parentContentControlOrNullObject

**Type:** `Word.ContentControl`

**Since:** WordApi 1.3

Gets the content control that contains the content control. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if a content control is nested inside another content control and display the parent's title, or indicate if it's a top-level control.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the parent content control (or null object if none exists)
    const parentContentControl = contentControl.parentContentControlOrNullObject;
    
    // Load properties
    contentControl.load("title");
    parentContentControl.load("isNullObject, title");
    
    await context.sync();
    
    // Check if parent exists
    if (parentContentControl.isNullObject) {
        console.log(`"${contentControl.title}" is a top-level content control (no parent).`);
    } else {
        console.log(`"${contentControl.title}" is nested inside "${parentContentControl.title}".`);
    }
});
```

---

### parentTable

**Type:** `Word.Table`

**Since:** WordApi 1.3

Gets the table that contains the content control. Throws an ItemNotFound error if it isn't contained in a table.

#### Examples

**Example**: Get the parent table of a content control and highlight it by applying a shading color to show which table contains the content control.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the parent table that contains this content control
    const parentTable = contentControl.parentTable;
    
    // Apply shading to the parent table to highlight it
    parentTable.shadingColor = "#FFFF00"; // Yellow background
    
    await context.sync();
    
    console.log("Parent table has been highlighted");
});
```

---

### parentTableCell

**Type:** `Word.TableCell`

**Since:** WordApi 1.3

Gets the table cell that contains the content control. Throws an ItemNotFound error if it isn't contained in a table cell.

#### Examples

**Example**: Highlight the parent table cell of a content control by setting its background color to yellow.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the parent table cell containing the content control
    const parentCell = contentControl.parentTableCell;
    
    // Set the cell's background color to yellow
    parentCell.shadingColor = "#FFFF00";
    
    await context.sync();
});
```

---

### parentTableCellOrNullObject

**Type:** `Word.TableCell`

**Since:** WordApi 1.3

Gets the table cell that contains the content control. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if a content control is inside a table cell, and if so, highlight the parent cell with a yellow background color.

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    const parentCell = contentControl.parentTableCellOrNullObject;
    
    parentCell.load("isNullObject");
    await context.sync();
    
    if (!parentCell.isNullObject) {
        parentCell.shadingColor = "#FFFF00"; // Yellow background
        console.log("Content control is in a table cell - cell highlighted");
    } else {
        console.log("Content control is not in a table cell");
    }
    
    await context.sync();
});
```

---

### parentTableOrNullObject

**Type:** `Word.Table`

**Since:** WordApi 1.3

Gets the table that contains the content control. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if a content control is inside a table, and if so, highlight the entire parent table with a yellow background color.

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    const parentTable = contentControl.parentTableOrNullObject;
    
    parentTable.load("isNullObject");
    await context.sync();
    
    if (!parentTable.isNullObject) {
        // Content control is inside a table
        parentTable.shadingColor = "yellow";
        console.log("Content control is in a table - highlighting it");
    } else {
        console.log("Content control is not in a table");
    }
    
    await context.sync();
});
```

---

### pictureContentControl

**Type:** `Word.PictureContentControl`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the picture-related data if the content control's Word.ContentControlType is Picture. It's null otherwise.

#### Examples

**Example**: Get the picture content control and set its image to a base64-encoded image if the content control is a picture type.

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    await context.sync();

    if (contentControl.type === Word.ContentControlType.picture) {
        const pictureContentControl = contentControl.pictureContentControl;
        pictureContentControl.load("imageFormat");
        await context.sync();

        // Set a new image (example base64 string - replace with actual image data)
        const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
        pictureContentControl.insertPicture(base64Image, Word.InsertLocation.replace);
        await context.sync();
        
        console.log("Picture updated successfully");
    } else {
        console.log("Content control is not a picture type");
    }
});
```

---

### placeholderText

**Type:** `string`

**Since:** WordApi 1.1

Specifies the placeholder text of the content control. Dimmed text will be displayed when the content control is empty.

#### Examples

**Example**: Set placeholder text "Enter your full name here" for a content control so users see dimmed guidance text when the control is empty.

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    contentControl.placeholderText = "Enter your full name here";
    
    await context.sync();
});
```

---

### removeWhenEdited

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the content control is removed after it is edited. Mutually exclusive with cannotDelete.

#### Examples

**Example**: Configure a content control to automatically remove itself after the user edits its content, useful for placeholder text that should disappear once modified.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Set the content control to be removed after editing
    contentControl.removeWhenEdited = true;
    
    // Load the property to verify
    contentControl.load("removeWhenEdited");
    
    await context.sync();
    
    console.log("Content control will be removed after editing:", contentControl.removeWhenEdited);
});
```

---

### repeatingSectionContentControl

**Type:** `Word.RepeatingSectionContentControl`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the repeating section-related data if the content control's Word.ContentControlType is RepeatingSection. It's null otherwise.

#### Examples

**Example**: Check if a content control is a repeating section and log its title if it is

```typescript
await Word.run(async (context) => {
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("type");
    
    const repeatingSectionCC = contentControl.repeatingSectionContentControl;
    repeatingSectionCC.load("title");
    
    await context.sync();
    
    if (contentControl.type === Word.ContentControlType.repeatingSection) {
        console.log("This is a repeating section with title: " + repeatingSectionCC.title);
    } else {
        console.log("This content control is not a repeating section");
    }
});
```

---

### style

**Type:** `string`

**Since:** WordApi 1.1

Specifies the style name for the content control. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

#### Examples

**Example**: Set a custom style named "CustomHeading" to a content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Set the custom style name
    contentControl.style = "CustomHeading";
    
    await context.sync();
});
```

---

### styleBuiltIn

**Type:** `Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"`

**Since:** WordApi 1.3

Specifies the built-in style name for the content control. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

#### Examples

**Example**: Apply the "Title" built-in style to all content controls in the document

```typescript
await Word.run(async (context) => {
    // Get all content controls in the document
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    
    await context.sync();
    
    // Apply the Title style to each content control
    for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].styleBuiltIn = "Title";
    }
    
    await context.sync();
});
```

---

### subtype

**Type:** `Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group"`

**Since:** WordApi 1.3

Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls, or 'PlainTextInline' and 'PlainTextParagraph' for plain text content controls, or 'CheckBox' for checkbox content controls.

#### Examples

**Example**: Get all content controls in the document and log their subtypes to identify which are inline rich text vs paragraph rich text vs checkboxes

```typescript
await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items/subtype, items/title");
    
    await context.sync();
    
    contentControls.items.forEach((cc) => {
        console.log(`Content Control "${cc.title}" has subtype: ${cc.subtype}`);
    });
});
```

---

### tables

**Type:** `Word.TableCollection`

**Since:** WordApi 1.3

Gets the collection of table objects in the content control.

#### Examples

**Example**: Get all tables within a content control and apply a built-in style to each table.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get all tables in the content control
    const tables = contentControl.tables;
    tables.load("items");
    
    await context.sync();
    
    // Apply a style to each table
    for (let i = 0; i < tables.items.length; i++) {
        tables.items[i].styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
    }
    
    await context.sync();
    
    console.log(`Found and styled ${tables.items.length} table(s) in the content control.`);
});
```

---

### tag

**Type:** `string`

**Since:** WordApi 1.1

Specifies a tag to identify a content control.

#### Examples

**Example**: Traverses each paragraph of the document and wraps a content control on each with either a even or odd tags.

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

---

### text

**Type:** `string`

**Since:** WordApi 1.1

Gets the text of the content control.

#### Examples

**Example**: Get the text content from all content controls in the document and display them in the console

```typescript
await Word.run(async (context) => {
    // Get all content controls in the document
    const contentControls = context.document.contentControls;
    contentControls.load("text");
    
    await context.sync();
    
    // Display the text of each content control
    for (let i = 0; i < contentControls.items.length; i++) {
        console.log(`Content Control ${i + 1}: ${contentControls.items[i].text}`);
    }
});
```

---

### title

**Type:** `string`

**Since:** WordApi 1.1

Specifies the title for a content control.

#### Examples

**Example**: Set the title of a content control to "Customer Information" to label it for users

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Set the title of the content control
    contentControl.title = "Customer Information";
    
    await context.sync();
});
```

---

### type

**Type:** `Word.ContentControlType | "Unknown" | "RichTextInline" | "RichTextParagraphs" | "RichTextTableCell" | "RichTextTableRow" | "RichTextTable" | "PlainTextInline" | "PlainTextParagraph" | "Picture" | "BuildingBlockGallery" | "CheckBox" | "ComboBox" | "DropDownList" | "DatePicker" | "RepeatingSection" | "RichText" | "PlainText" | "Group"`

**Since:** WordApi 1.1

Gets the content control type. Only rich text, plain text, and checkbox content controls are supported currently.

#### Examples

**Example**: Check the type of a content control and display different messages based on whether it's a checkbox, dropdown list, or rich text control.

```typescript
await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    if (contentControls.items.length > 0) {
        const contentControl = contentControls.items[0];
        contentControl.load("type");
        await context.sync();

        const type = contentControl.type;
        
        if (type === "CheckBox") {
            console.log("This is a checkbox content control");
        } else if (type === "DropDownList") {
            console.log("This is a dropdown list content control");
        } else if (type === "RichText" || type === "RichTextParagraphs") {
            console.log("This is a rich text content control");
        } else if (type === "PlainText" || type === "PlainTextParagraph") {
            console.log("This is a plain text content control");
        } else {
            console.log(`Content control type: ${type}`);
        }
    }
});
```

---

### xmlMapping

**Type:** `Word.XmlMapping`

**Since:** WordApi BETA (PREVIEW ONLY)

Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

#### Examples

**Example**: Check if a content control is mapped to XML data and display the XPath of the mapping if it exists.

```typescript
await Word.run(async (context) => {
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    if (contentControls.items.length > 0) {
        const contentControl = contentControls.items[0];
        const xmlMapping = contentControl.xmlMapping;
        xmlMapping.load("isMapped, xpath");
        await context.sync();

        if (xmlMapping.isMapped) {
            console.log("Content control is mapped to XML data");
            console.log("XPath: " + xmlMapping.xpath);
        } else {
            console.log("Content control is not mapped to XML data");
        }
    }
});
```

---

## Methods

### clear

**Kind:** `delete`

Clears the contents of the content control. The user can perform the undo operation on the cleared content.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### delete

**Kind:** `delete`

Deletes the content control and its content. If keepContent is set to true, the content isn't deleted.

#### Signature

**Parameters:**
- `keepContent`: `boolean` (required)
  Required. Indicates whether the content should be deleted with the content control. If keepContent is set to true, the content isn't deleted.

**Returns:** `void`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

**Example**: Delete the first content control tagged "forTesting" from the document without removing its content.

```typescript
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

### getComments

**Kind:** `read`

Gets comments associated with the content control.

#### Signature

**Returns:** `Word.CommentCollection`

#### Examples

**Example**: Retrieve and display all comments associated with a specific content control in the document

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get comments associated with the content control
    const comments = contentControl.getComments();
    comments.load("items");
    
    await context.sync();
    
    // Display the comments
    console.log(`Found ${comments.items.length} comment(s) on this content control`);
    
    comments.items.forEach((comment, index) => {
        comment.load("content, authorName");
    });
    
    await context.sync();
    
    comments.items.forEach((comment, index) => {
        console.log(`Comment ${index + 1}: "${comment.content}" by ${comment.authorName}`);
    });
});
```

---

### getContentControls

**Kind:** `read`

Gets the currently supported child content controls in this content control.

#### Signature

**Parameters:**
- `options`: `Word.ContentControlOptions` (optional)
  Optional. Options that define which content controls are returned.

**Returns:** `Word.ContentControlCollection`

#### Examples

**Example**: Get all child content controls within a parent content control and display their titles in the console.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const parentContentControl = context.document.contentControls.getFirst();
    
    // Get all child content controls within the parent
    const childContentControls = parentContentControl.getContentControls();
    
    // Load the title property for each child content control
    childContentControls.load("title");
    
    await context.sync();
    
    // Display the titles of all child content controls
    console.log(`Found ${childContentControls.items.length} child content controls:`);
    childContentControls.items.forEach((cc, index) => {
        console.log(`Child ${index + 1}: ${cc.title}`);
    });
});
```

---

### getHtml

**Kind:** `serialize`

Gets an HTML representation of the content control object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `ContentControl.getOoxml()` and convert the returned XML to HTML.

#### Signature

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### getOoxml

**Kind:** `serialize`

Gets the Office Open XML (OOXML) representation of the content control object.

#### Signature

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### getRange

**Kind:** `read`

Gets the whole content control, or the starting or ending point of the content control, as a range.

#### Signature

**Parameters:**
- `rangeLocation`: `Word.RangeLocation | "Whole" | "Start" | "End" | "Before" | "After" | "Content"` (optional)
  Optional. The range location must be 'Whole', 'Start', 'End', 'Before', 'After', or 'Content'.

**Returns:** `Word.Range`

#### Examples

**Example**: Get the range of the first content control in the document and highlight its entire content in yellow.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get the range of the entire content control
    const range = contentControl.getRange(Word.RangeLocation.whole);
    
    // Highlight the range in yellow
    range.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### getReviewedText

**Kind:** `read`

Gets reviewed text based on ChangeTrackingVersion selection.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `changeTrackingVersion`: `Word.ChangeTrackingVersion` (optional)
    Optional. The value must be 'Original' or 'Current'. The default is 'Current'.

  **Returns:** `OfficeExtension.ClientResult<string>`

**Overload 2:**

  **Parameters:**
  - `changeTrackingVersion`: `"Original" | "Current"` (optional)
    Optional. The value must be 'Original' or 'Current'. The default is 'Current'.

  **Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Get and display the reviewed text from the first content control in the document, showing both the original and current versions based on change tracking.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    if (contentControls.items.length > 0) {
        const contentControl = contentControls.items[0];
        
        // Get the original text (before tracked changes)
        const originalText = contentControl.getReviewedText(Word.ChangeTrackingVersion.original);
        
        // Get the current text (with tracked changes applied)
        const currentText = contentControl.getReviewedText(Word.ChangeTrackingVersion.current);
        
        await context.sync();
        
        console.log("Original text: " + originalText.value);
        console.log("Current text: " + currentText.value);
    } else {
        console.log("No content controls found in the document.");
    }
});
```

---

### getTextRanges

**Kind:** `read`

Gets the text ranges in the content control by using punctuation marks and/or other ending marks.

#### Signature

**Parameters:**
- `endingMarks`: `string[]` (required)
  Required. The punctuation marks and/or other ending marks as an array of strings.
- `trimSpacing`: `boolean` (optional)
  Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.

**Returns:** `Word.RangeCollection`

#### Examples

**Example**: Get all sentences from a content control by splitting on periods, then highlight the first sentence in yellow.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get text ranges split by periods (sentences)
    const textRanges = contentControl.getTextRanges(["."], true);
    textRanges.load("items");
    
    await context.sync();
    
    // Highlight the first sentence if it exists
    if (textRanges.items.length > 0) {
        textRanges.items[0].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### getTrackedChanges

**Kind:** `read`

Gets the collection of the TrackedChange objects in the content control.

#### Signature

**Returns:** `Word.TrackedChangeCollection`

#### Examples

**Example**: Get and display all tracked changes within a specific content control, showing the type and author of each change.

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Get all tracked changes in the content control
    const trackedChanges = contentControl.getTrackedChanges();
    
    // Load properties of the tracked changes
    trackedChanges.load("items");
    context.load(trackedChanges, "type, author, date");
    
    await context.sync();
    
    // Display information about each tracked change
    console.log(`Found ${trackedChanges.items.length} tracked change(s) in the content control`);
    
    trackedChanges.items.forEach((change, index) => {
        console.log(`Change ${index + 1}:`);
        console.log(`  Type: ${change.type}`);
        console.log(`  Author: ${change.author}`);
        console.log(`  Date: ${change.date}`);
    });
});
```

---

### insertBreak

**Kind:** `create`

Inserts a break at the specified location in the main document. This method cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.

#### Signature

**Parameters:**
- `breakType`: `Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line"` (required)
  Required. Type of break.
- `insertLocation`: `Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"` (required)
  Required. The value must be 'Start', 'End', 'Before', or 'After'.

**Returns:** `void`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### insertFileFromBase64

**Kind:** `create`

Inserts a document into the content control at the specified location.

#### Signature

**Parameters:**
- `base64File`: `string` (required)
  Required. The Base64-encoded content of a .docx file.
- `insertLocation`: `Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"` (required)
  Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

**Returns:** `Word.Range`

#### Examples

**Example**: Insert a pre-encoded Word document into an existing content control at the beginning of its content

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Base64-encoded Word document content
    const base64File = "UEsDBBQABgAIAAAAIQDfpNJsWgEAACAFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAAC...";
    
    // Insert the document at the start of the content control
    contentControl.insertFileFromBase64(base64File, Word.InsertLocation.start);
    
    await context.sync();
    
    console.log("Document inserted into content control successfully");
});
```

---

### insertHtml

**Kind:** `create`

Inserts HTML into the content control at the specified location.

#### Signature

**Parameters:**
- `html`: `string` (required)
  Required. The HTML to be inserted in to the content control.
- `insertLocation`: `Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"` (required)
  Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

**Returns:** `Word.Range`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### insertInlinePictureFromBase64

**Kind:** `create`

Inserts an inline picture into the content control at the specified location.

#### Signature

**Parameters:**
- `base64EncodedImage`: `string` (required)
  Required. The Base64-encoded image to be inserted in the content control.
- `insertLocation`: `Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"` (required)
  Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

**Returns:** `Word.InlinePicture`

#### Examples

**Example**: Insert a company logo image into a content control tagged as "LogoPlaceholder" at the beginning of the control

```typescript
await Word.run(async (context) => {
    // Get the content control by tag
    const contentControls = context.document.contentControls.getByTag("LogoPlaceholder");
    contentControls.load("items");
    await context.sync();

    if (contentControls.items.length > 0) {
        const logoControl = contentControls.items[0];
        
        // Base64 encoded image string (example: small PNG image)
        const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
        
        // Insert the inline picture at the start of the content control
        logoControl.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.start);
        
        await context.sync();
        console.log("Logo inserted successfully");
    }
});
```

---

### insertOoxml

**Kind:** `create`

Inserts OOXML into the content control at the specified location.

#### Signature

**Parameters:**
- `ooxml`: `string` (required)
  Required. The OOXML to be inserted in to the content control.
- `insertLocation`: `Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"` (required)
  Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

**Returns:** `Word.Range`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### insertParagraph

**Kind:** `create`

Inserts a paragraph at the specified location.

#### Signature

**Parameters:**
- `paragraphText`: `string` (required)
  Required. The paragraph text to be inserted.
- `insertLocation`: `Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"` (required)
  Required. The value must be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### insertTable

**Kind:** `create`

Inserts a table with the specified number of rows and columns into, or next to, a content control.

#### Signature

**Parameters:**
- `rowCount`: `number` (required)
  Required. The number of rows in the table.
- `columnCount`: `number` (required)
  Required. The number of columns in the table.
- `insertLocation`: `Word.InsertLocation.start | Word.InsertLocation.end | Word.InsertLocation.before | Word.InsertLocation.after | "Start" | "End" | "Before" | "After"` (required)
  Required. The value must be 'Start', 'End', 'Before', or 'After'. 'Before' and 'After' cannot be used with 'RichTextTable', 'RichTextTableRow' and 'RichTextTableCell' content controls.
- `values`: `string[][]` (optional)
  Optional 2D array. Cells are filled if the corresponding strings are specified in the array.

**Returns:** `Word.Table`

#### Examples

**Example**: Insert a 3x4 table with sample data into an existing content control tagged "ReportData"

```typescript
await Word.run(async (context) => {
    // Get the content control by tag
    const contentControls = context.document.contentControls.getByTag("ReportData");
    const contentControl = contentControls.getFirst();
    
    // Define table data
    const tableData = [
        ["Product", "Q1", "Q2", "Q3"],
        ["Widget A", "100", "120", "135"],
        ["Widget B", "85", "90", "95"]
    ];
    
    // Insert a 3x4 table into the content control
    const table = contentControl.insertTable(3, 4, Word.InsertLocation.replace, tableData);
    
    // Optional: Format the table
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
    
    await context.sync();
});
```

---

### insertText

**Kind:** `create`

Inserts text into the content control at the specified location.

#### Signature

**Parameters:**
- `text`: `string` (required)
  Required. The text to be inserted in to the content control.
- `insertLocation`: `Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"` (required)
  Required. The value must be 'Replace', 'Start', or 'End'. 'Replace' cannot be used with 'RichTextTable' and 'RichTextTableRow' content controls.

**Returns:** `Word.Range`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.ContentControlLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.ContentControl`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.ContentControl`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.ContentControl`

#### Examples

**Example**: Load all of the content control properties

```typescript
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

### resetState

**Kind:** `configure`

Resets the state of the content control.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Reset the state of the first content control in the document and log its ID to the console.

```typescript
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

### search

**Kind:** `read`

Performs a search with the specified SearchOptions on the scope of the content control object. The search results are a collection of range objects.

#### Signature

**Parameters:**
- `searchText`: `string` (required)
  Required. The search text.
- `searchOptions`: `Word.SearchOptions | { ignorePunct?: boolean; ignoreSpace?: boolean; matchCase?: boolean; matchPrefix?: boolean; matchSuffix?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean; }` (optional)
  Optional. Options for the search.

**Returns:** `Word.RangeCollection`

#### Examples

**Example**: Run a batch operation against the Word object model.

```typescript
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

### select

**Kind:** `read`

Selects the content control. This causes Word to scroll to the selection.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `selectionMode`: `Word.SelectionMode` (optional)
    Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `selectionMode`: `"Select" | "Start" | "End"` (optional)
    Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

  **Returns:** `void`

#### Examples

**Example**: Select the first content control in the document and scroll to it so the user can see it

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Select the content control with default selection mode
    contentControl.select();
    
    await context.sync();
});
```

---

### set

**Kind:** `configure`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Word.Interfaces.ContentControlUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.ContentControl` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure content controls tagged as "even" and "odd" with different colors, titles, appearances, and append content to each group.

```typescript
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

### setState

**Kind:** `configure`

Sets the state of the content control.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `contentControlState`: `Word.ContentControlState` (required)
    State to be set.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `contentControlState`: `"Error" | "Warning"` (required)
    State to be set.

  **Returns:** `void`

#### Examples

**Example**: Set the state of the first content control in the document to a user-selected value from a dropdown menu.

```typescript
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

### split

**Kind:** `write`

Splits the content control into child ranges by using delimiters.

#### Signature

**Parameters:**
- `delimiters`: `string[]` (required)
  Required. The delimiters as an array of strings.
- `multiParagraphs`: `boolean` (optional)
  Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters.
- `trimDelimiters`: `boolean` (optional)
  Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection.
- `trimSpacing`: `boolean` (optional)
  Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection.

**Returns:** `Word.RangeCollection`

#### Examples

**Example**: Split a content control containing comma-separated values into separate ranges for individual processing

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("text");
    
    await context.sync();
    
    // Split the content control by commas
    const ranges = contentControl.split([","], false, true, true);
    ranges.load("items");
    
    await context.sync();
    
    // Process each split range (e.g., highlight them)
    for (let i = 0; i < ranges.items.length; i++) {
        ranges.items[i].font.highlightColor = i % 2 === 0 ? "yellow" : "lightblue";
    }
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.ContentControl` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.ContentControlData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ContentControlData`

#### Examples

**Example**: Serialize a content control's properties to JSON format for logging or data export purposes

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    
    // Load properties to serialize
    contentControl.load("text,tag,title,type,appearance,cannotDelete,cannotEdit");
    
    await context.sync();
    
    // Convert the content control to a plain JavaScript object
    const contentControlData = contentControl.toJSON();
    
    // Log the serialized data
    console.log("Content Control Data:", JSON.stringify(contentControlData, null, 2));
    
    // Example output structure:
    // {
    //   "text": "Sample content",
    //   "tag": "myTag",
    //   "title": "My Content Control",
    //   "type": "RichText",
    //   "appearance": "BoundingBox",
    //   "cannotDelete": false,
    //   "cannotEdit": false
    // }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.ContentControl`

#### Examples

**Example**: Track a content control across multiple sync calls to modify its properties without getting an "InvalidObjectPath" error

```typescript
await Word.run(async (context) => {
    // Get the first content control in the document
    const contentControl = context.document.contentControls.getFirst();
    contentControl.load("title,tag");
    await context.sync();
    
    // Track the content control to use it across multiple sync calls
    contentControl.track();
    
    // First sync: modify the title
    contentControl.title = "Updated Title";
    await context.sync();
    
    // Second sync: modify the tag (tracking prevents InvalidObjectPath error)
    contentControl.tag = "updated-tag";
    await context.sync();
    
    // Third sync: change appearance
    contentControl.appearance = "Tags";
    await context.sync();
    
    // Untrack when done to free up memory
    contentControl.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `Word.ContentControl`

#### Examples

**Example**: Process multiple content controls to find those with a specific tag, then untrack them to free memory after processing

```typescript
await Word.run(async (context) => {
    // Get all content controls in the document
    const contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    // Process content controls with a specific tag
    const processedControls: Word.ContentControl[] = [];
    
    for (let i = 0; i < contentControls.items.length; i++) {
        const cc = contentControls.items[i];
        cc.load("tag, text");
        context.trackedObjects.add(cc); // Track for processing
        processedControls.push(cc);
    }
    
    await context.sync();

    // Do some work with the content controls
    for (const cc of processedControls) {
        if (cc.tag === "processed") {
            console.log(cc.text);
        }
    }

    // Untrack all processed content controls to free memory
    for (const cc of processedControls) {
        cc.untrack();
    }
    
    await context.sync();
});
```

---
