# Word.Paragraph

**Package:** `word`

**API Set:** None None

## Description

Represents a single paragraph in a selection, range, content control, or document body.

## Class Examples

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

### alignment

**Type:** `Word.Alignment | "Mixed" | "Unknown" | "Left" | "Centered" | "Right" | "Justified"`

Specifies the alignment for a paragraph. The value can be 'left', 'centered', 'right', or 'justified'.

#### Examples

**Example**: Center the alignment of the last paragraph in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Center last paragraph alignment.
  context.document.body.paragraphs.getLast().alignment = "Centered";

  await context.sync();
});
```

---

### borders

**Type:** `Word.BorderUniversalCollection`

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Set a red bottom border with 2pt width on the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const borders = paragraph.borders;
    
    borders.load("items");
    await context.sync();
    
    const bottomBorder = borders.items.find(border => border.type === Word.BorderType.bottom);
    if (bottomBorder) {
        bottomBorder.color = "#FF0000";
        bottomBorder.width = 2;
        bottomBorder.visible = true;
    }
    
    await context.sync();
});
```

---

### contentControls

**Type:** `Word.ContentControlCollection`

Gets the collection of content control objects in the paragraph.

#### Examples

**Example**: Find and highlight all content controls within the first paragraph of the document by setting their appearance to tags visible.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the collection of content controls in the paragraph
    const contentControls = firstParagraph.contentControls;
    
    // Load the content controls
    contentControls.load("items");
    
    await context.sync();
    
    // Set appearance for each content control to make them visible
    for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].appearance = Word.ContentControlAppearance.tags;
    }
    
    await context.sync();
    
    console.log(`Found ${contentControls.items.length} content control(s) in the first paragraph.`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context to synchronize paragraph properties and log the paragraph's text content to the console

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Load the text property using the paragraph's context
    paragraph.load("text");
    
    // Use the context property to sync with the Office host
    await paragraph.context.sync();
    
    // Now we can access the loaded property
    console.log("Paragraph text:", paragraph.text);
});
```

---

### endnotes

**Type:** `Word.NoteItemCollection`

Gets the collection of endnotes in the paragraph.

#### Examples

**Example**: Count the number of endnotes in the first paragraph of the document and display the count in the console.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the endnotes collection from the paragraph
    const endnotes = firstParagraph.endnotes;
    
    // Load the count of endnotes
    endnotes.load("items");
    
    await context.sync();
    
    // Display the count
    console.log(`Number of endnotes in the first paragraph: ${endnotes.items.length}`);
});
```

---

### fields

**Type:** `Word.FieldCollection`

Gets the collection of fields in the paragraph.

#### Examples

**Example**: Get all fields in the first paragraph of the document and display their types in the console.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const fields = firstParagraph.fields;
    
    fields.load("items");
    await context.sync();
    
    console.log(`Found ${fields.items.length} field(s) in the paragraph`);
    
    for (let i = 0; i < fields.items.length; i++) {
        fields.items[i].load("type");
    }
    await context.sync();
    
    fields.items.forEach((field, index) => {
        console.log(`Field ${index + 1}: ${field.type}`);
    });
});
```

---

### firstLineIndent

**Type:** `number`

Specifies the value, in points, for a first line or hanging indent. Use a positive value to set a first-line indent, and use a negative value to set a hanging indent.

#### Examples

**Example**: Set a first-line indent of 36 points for the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.firstLineIndent = 36;
    
    await context.sync();
});
```

---

### font

**Type:** `Word.Font`

Gets the text format of the paragraph. Use this to get and set font name, size, color, and other properties.

#### Examples

**Example**: Set the paragraph font to Calibri with size 14 and blue color

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    
    paragraph.font.name = "Calibri";
    paragraph.font.size = 14;
    paragraph.font.color = "blue";
    
    await context.sync();
});
```

---

### footnotes

**Type:** `Word.NoteItemCollection`

Gets the collection of footnotes in the paragraph.

#### Examples

**Example**: Count and display the number of footnotes in the first paragraph of the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the footnotes collection from the paragraph
    const footnotes = firstParagraph.footnotes;
    
    // Load the count of footnotes
    footnotes.load("items");
    
    await context.sync();
    
    // Display the number of footnotes
    console.log(`The first paragraph contains ${footnotes.items.length} footnote(s)`);
});
```

---

### inlinePictures

**Type:** `Word.InlinePictureCollection`

Gets the collection of InlinePicture objects in the paragraph. The collection doesn't include floating images.

#### Examples

**Example**: Get all inline pictures from the first paragraph and set their width to 150 pixels

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const inlinePictures = firstParagraph.inlinePictures;
    
    inlinePictures.load("items");
    await context.sync();
    
    for (let i = 0; i < inlinePictures.items.length; i++) {
        inlinePictures.items[i].width = 150;
    }
    
    await context.sync();
});
```

---

### isLastParagraph

**Type:** `boolean`

Indicates the paragraph is the last one inside its parent body.

#### Examples

**Example**: Add a page break after the current paragraph only if it's not the last paragraph in the document body.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    paragraph.load("isLastParagraph");
    
    await context.sync();
    
    if (!paragraph.isLastParagraph) {
        paragraph.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
        console.log("Page break added after paragraph");
    } else {
        console.log("This is the last paragraph - no page break added");
    }
    
    await context.sync();
});
```

---

### isListItem

**Type:** `boolean`

Checks whether the paragraph is a list item.

#### Examples

**Example**: Check if a paragraph is a list item and display an alert with the result

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.load("isListItem");
    
    await context.sync();
    
    if (paragraph.isListItem) {
        console.log("This paragraph is a list item.");
    } else {
        console.log("This paragraph is not a list item.");
    }
});
```

---

### leftIndent

**Type:** `number`

Specifies the left indent value, in points, for the paragraph.

#### Examples

**Example**: Set the left indent of the first paragraph in the document to 75 points.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Indent the first paragraph.
  context.document.body.paragraphs.getFirst().leftIndent = 75; //units = points

  return context.sync();
});
```

---

### lineSpacing

**Type:** `number`

Specifies the line spacing, in points, for the specified paragraph. In the Word UI, this value is divided by 12.

#### Examples

**Example**: Set the line spacing of the first paragraph in the document body to 20 points.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Adjust line spacing.
  context.document.body.paragraphs.getFirst().lineSpacing = 20;

  await context.sync();
});
```

---

### lineUnitAfter

**Type:** `number`

Specifies the amount of spacing, in grid lines, after the paragraph.

#### Examples

**Example**: Set the spacing after the first paragraph in the document body to 1 line unit.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Set the space (in line units) after the first paragraph.
  context.document.body.paragraphs.getFirst().lineUnitAfter = 1;

  await context.sync();
});
```

---

### lineUnitBefore

**Type:** `number`

Specifies the amount of spacing, in grid lines, before the paragraph.

#### Examples

**Example**: Set the spacing before the first paragraph in the document body to 1 line unit.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Set the space (in line units) before the first paragraph.
  context.document.body.paragraphs.getFirst().lineUnitBefore = 1;

  await context.sync();
});
```

---

### list

**Type:** `Word.List`

Gets the List to which this paragraph belongs. Throws an ItemNotFound error if the paragraph isn't in a list.

#### Examples

**Example**: Get the list level of the first paragraph in the document and display it to the user.

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const list = firstParagraph.list;
    
    list.load("levelTypes");
    await context.sync();
    
    console.log("List level types:", list.levelTypes);
    // Or display to user in your UI
});
```

---

### listItem

**Type:** `Word.ListItem`

Gets the ListItem for the paragraph. Throws an ItemNotFound error if the paragraph isn't part of a list.

#### Examples

**Example**: Check if a paragraph is part of a list and get its level number, displaying an error message if it's not in a list.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    
    try {
        const listItem = paragraph.listItem;
        listItem.load("level");
        
        await context.sync();
        
        console.log(`This paragraph is at list level: ${listItem.level}`);
    } catch (error) {
        if (error.code === "ItemNotFound") {
            console.log("This paragraph is not part of a list.");
        } else {
            throw error;
        }
    }
});
```

---

### listItemOrNullObject

**Type:** `Word.ListItem`

Gets the ListItem for the paragraph. If the paragraph isn't part of a list, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

#### Examples

**Example**: Check if the current paragraph is part of a list and display its list level if it is, otherwise display a message that it's not in a list.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const listItem = paragraph.listItemOrNullObject;
    
    listItem.load("isNullObject, level");
    await context.sync();
    
    if (listItem.isNullObject) {
        console.log("This paragraph is not part of a list.");
    } else {
        console.log(`This paragraph is in a list at level ${listItem.level}.`);
    }
});
```

---

### listOrNullObject

**Type:** `Word.List`

Gets the List to which this paragraph belongs. If the paragraph isn't in a list, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

#### Examples

**Example**: Check if the current paragraph is part of a list and display the list ID if it is, or show a message if it's not in a list.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const list = paragraph.listOrNullObject;
    
    list.load("id, isNullObject");
    await context.sync();
    
    if (list.isNullObject) {
        console.log("This paragraph is not part of a list.");
    } else {
        console.log(`This paragraph belongs to list with ID: ${list.id}`);
    }
});
```

---

### outlineLevel

**Type:** `number`

Specifies the outline level for the paragraph.

#### Examples

**Example**: Set the outline level of the first paragraph in the document to level 2 to organize it as a subheading in the document structure

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.outlineLevel = 2;
    
    await context.sync();
});
```

---

### parentBody

**Type:** `Word.Body`

Gets the parent body of the paragraph.

#### Examples

**Example**: Get the parent body of a paragraph and highlight it with yellow background color to show the entire document body that contains the paragraph.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the parent body of the paragraph
    const parentBody = paragraph.parentBody;
    
    // Highlight the parent body with yellow background
    parentBody.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### parentContentControl

**Type:** `Word.ContentControl`

Gets the content control that contains the paragraph. Throws an ItemNotFound error if there isn't a parent content control.

#### Examples

**Example**: Check if a paragraph is inside a content control and highlight the parent content control in yellow if it exists.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    
    try {
        const parentContentControl = paragraph.parentContentControl;
        parentContentControl.load("title");
        await context.sync();
        
        // Highlight the parent content control
        parentContentControl.font.highlightColor = "yellow";
        
        console.log(`Paragraph is inside content control: ${parentContentControl.title}`);
        await context.sync();
    } catch (error) {
        console.log("Paragraph is not inside a content control");
    }
});
```

---

### parentContentControlOrNullObject

**Type:** `Word.ContentControl`

Gets the content control that contains the paragraph. If there isn't a parent content control, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

#### Examples

**Example**: Check if a paragraph is inside a content control and highlight the content control in yellow if it exists

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const parentContentControl = paragraph.parentContentControlOrNullObject;
    
    parentContentControl.load("isNullObject");
    await context.sync();
    
    if (!parentContentControl.isNullObject) {
        parentContentControl.font.highlightColor = "yellow";
        console.log("Paragraph is inside a content control - highlighted it");
    } else {
        console.log("Paragraph is not inside a content control");
    }
    
    await context.sync();
});
```

---

### parentTable

**Type:** `Word.Table`

Gets the table that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table.

#### Examples

**Example**: Check if a paragraph is inside a table and highlight the entire table with a yellow background if it is.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    
    try {
        const table = paragraph.parentTable;
        table.load("values");
        await context.sync();
        
        // If we reach here, paragraph is in a table
        table.shadingColor = "#FFFF00"; // Yellow background
        await context.sync();
        
        console.log("Table highlighted successfully");
    } catch (error) {
        if (error.code === "ItemNotFound") {
            console.log("Selected paragraph is not in a table");
        } else {
            throw error;
        }
    }
});
```

---

### parentTableCell

**Type:** `Word.TableCell`

Gets the table cell that contains the paragraph. Throws an ItemNotFound error if it isn't contained in a table cell.

#### Examples

**Example**: Check if a paragraph is inside a table cell and highlight the entire cell in yellow if it is.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    
    try {
        const tableCell = paragraph.parentTableCell;
        tableCell.load("cellIndex, rowIndex");
        
        await context.sync();
        
        // Highlight the parent table cell
        tableCell.shadingColor = "#FFFF00";
        
        await context.sync();
        
        console.log(`Paragraph is in cell at row ${tableCell.rowIndex}, column ${tableCell.cellIndex}`);
    } catch (error) {
        console.log("Paragraph is not inside a table cell");
    }
});
```

---

### parentTableCellOrNullObject

**Type:** `Word.TableCell`

Gets the table cell that contains the paragraph. If it isn't contained in a table cell, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

#### Examples

**Example**: Check if the current paragraph is inside a table cell, and if so, highlight the entire cell in yellow.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the parent table cell (or null object if not in a table)
    const tableCell = paragraph.parentTableCellOrNullObject;
    
    // Load the isNullObject property to check if paragraph is in a table
    tableCell.load("isNullObject");
    
    await context.sync();
    
    // Check if the paragraph is inside a table cell
    if (!tableCell.isNullObject) {
        // Paragraph is in a table - highlight the cell
        tableCell.shadingColor = "yellow";
        console.log("Paragraph is in a table cell - cell highlighted");
    } else {
        console.log("Paragraph is not in a table cell");
    }
    
    await context.sync();
});
```

---

### parentTableOrNullObject

**Type:** `Word.Table`

Gets the table that contains the paragraph. If it isn't contained in a table, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

#### Examples

**Example**: Check if a paragraph is inside a table, and if so, apply a light blue shading to the entire table.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const parentTable = paragraph.parentTableOrNullObject;
    
    // Load the isNullObject property to check if paragraph is in a table
    parentTable.load("isNullObject");
    await context.sync();
    
    if (!parentTable.isNullObject) {
        // Paragraph is inside a table, apply shading
        parentTable.shadingColor = "#ADD8E6"; // Light blue
        await context.sync();
        console.log("Table shading applied");
    } else {
        console.log("Paragraph is not in a table");
    }
});
```

---

### range

**Type:** `Word.Range`

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Highlight all text in the first paragraph by applying a yellow background color to its range

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the range of the paragraph
    const paragraphRange = paragraph.range;
    
    // Apply yellow highlighting to the paragraph's range
    paragraphRange.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### rightIndent

**Type:** `number`

Specifies the right indent value, in points, for the paragraph.

#### Examples

**Example**: Set the right indent of the first paragraph in the document to 36 points (0.5 inches)

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.rightIndent = 36;
    
    await context.sync();
});
```

---

### shading

**Type:** `Word.ShadingUniversal`

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Examples

**Example**: Apply a light gray background shading to the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Apply shading to the paragraph
    paragraph.shading.backgroundPatternColor = "#D3D3D3"; // Light gray
    
    await context.sync();
});
```

---

### shapes

**Type:** `Word.ShapeCollection`

Gets the collection of shape objects anchored in the paragraph, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

#### Examples

**Example**: Get all shapes anchored in the first paragraph and change the fill color of the first shape to blue

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const shapes = firstParagraph.shapes;
    
    shapes.load("items");
    await context.sync();
    
    if (shapes.items.length > 0) {
        const firstShape = shapes.items[0];
        firstShape.fill.setSolidColor("blue");
    }
    
    await context.sync();
});
```

---

### spaceAfter

**Type:** `number`

Specifies the spacing, in points, after the paragraph.

#### Examples

**Example**: Set the spacing after the first paragraph in the document to 20 points.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/paragraph-properties.yaml

await Word.run(async (context) => {
  // Set the space (in points) after the first paragraph.
  context.document.body.paragraphs.getFirst().spaceAfter = 20;

  await context.sync();
});
```

---

### spaceBefore

**Type:** `number`

Specifies the spacing, in points, before the paragraph.

#### Examples

**Example**: Set the spacing before the first paragraph in the document to 24 points

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.spaceBefore = 24;
    
    await context.sync();
});
```

---

### style

**Type:** `string`

Specifies the style name for the paragraph. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

#### Examples

**Example**: Apply a user-specified paragraph style to the first paragraph of the document body after validating the style exists and is of paragraph type.

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

---

### styleBuiltIn

**Type:** `Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"`

Specifies the built-in style name for the paragraph. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

#### Examples

**Example**: Create a structured document section with multiple paragraphs using built-in styles (Heading1, Heading2, Normal), including a highlighted content control placeholder for project costs and a page break at the end.

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

---

### tableNestingLevel

**Type:** `number`

Gets the level of the paragraph's table. It returns 0 if the paragraph isn't in a table.

#### Examples

**Example**: Check if the selected paragraph is in a table and display its nesting level, then highlight paragraphs that are in nested tables (level 2 or deeper)

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    paragraph.load("tableNestingLevel");
    
    await context.sync();
    
    const nestingLevel = paragraph.tableNestingLevel;
    console.log(`Paragraph table nesting level: ${nestingLevel}`);
    
    if (nestingLevel >= 2) {
        paragraph.font.highlightColor = "yellow";
        console.log("This paragraph is in a nested table!");
    } else if (nestingLevel === 1) {
        console.log("This paragraph is in a top-level table.");
    } else {
        console.log("This paragraph is not in a table.");
    }
    
    await context.sync();
});
```

---

### text

**Type:** `string`

Gets the text of the paragraph.

#### Examples

**Example**: Retrieve and display the text content of the first paragraph in the current selection.

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

---

### uniqueLocalId

**Type:** `string`

Gets a string that represents the paragraph identifier in the current session. ID is in standard 8-4-4-4-12 GUID format without curly braces and differs across sessions and coauthors.

#### Examples

**Example**: Register event handlers for paragraph changes and annotation interactions including clicks, hovers, insertions, removals, and popup actions in a Word document.

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

**Example**: Log the event type, unique local ID, and text content of each paragraph that triggered a paragraph changed event.

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

---

## Methods

### attachToList

Lets the paragraph join an existing list at the specified level. Fails if the paragraph cannot join the list or if the paragraph is already a list item.

#### Signature

**Parameters:**
- `listId`: `None` (required)
- `level`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Attach the first paragraph in the document to an existing numbered list at level 0

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get an existing list (assuming list with ID 1 exists)
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();
    
    const listId = lists.items[0].id;
    
    // Attach the paragraph to the list at level 0
    paragraph.attachToList(listId, 0);
    
    await context.sync();
    
    console.log("Paragraph attached to list successfully");
});
```

---

### clear

**Kind:** `delete`

Clears the contents of the paragraph object. The user can perform the undo operation on the cleared content.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Clear the contents of the first paragraph in the document body.

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

---

### closeUp

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Remove spacing before a paragraph to close it up with the preceding content

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Remove spacing before the paragraph
    paragraph.closeUp();
    
    await context.sync();
    
    console.log("Paragraph spacing removed successfully");
});
```

---

### delete

**Kind:** `delete`

Deletes the paragraph and its content from the document.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Delete the first paragraph from the document body.

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

---

### detachFromList

Moves this paragraph out of its list, if the paragraph is a list item.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Remove the first paragraph from its list while keeping the paragraph text in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Detach the paragraph from its list
    firstParagraph.detachFromList();
    
    await context.sync();
    
    console.log("Paragraph has been removed from the list");
});
```

---

### getAnnotations

**Kind:** `read`

Gets annotations set on this Paragraph object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Read and display all annotations that have been set on the first paragraph in the document.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get annotations on the paragraph
    const annotations = firstParagraph.getAnnotations();
    
    // Load the annotations collection
    context.load(annotations, "items");
    
    await context.sync();
    
    // Display the annotations
    console.log(`Found ${annotations.items.length} annotation(s) on the first paragraph`);
    
    annotations.items.forEach((annotation, index) => {
        context.load(annotation, "id, critiqueAnnotation");
        
        context.sync().then(() => {
            console.log(`Annotation ${index + 1}: ID = ${annotation.id}`);
        });
    });
    
    await context.sync();
});
```

---

### getComments

**Kind:** `read`

Gets comments associated with the paragraph.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve and display all comments associated with the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get comments associated with the paragraph
    const comments = firstParagraph.getComments();
    
    // Load comment properties
    comments.load("items");
    
    await context.sync();
    
    // Display comment information
    console.log(`Found ${comments.items.length} comment(s) on the first paragraph`);
    
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

Gets the currently supported content controls in the paragraph.

#### Signature

**Parameters:**
- `options`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Get all content controls within the first paragraph of the document and display their titles in the console.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get all content controls in the paragraph
    const contentControls = paragraph.getContentControls();
    
    // Load the title property of each content control
    contentControls.load("title");
    
    await context.sync();
    
    // Display the titles of the content controls
    console.log(`Found ${contentControls.items.length} content control(s) in the paragraph:`);
    contentControls.items.forEach((cc, index) => {
        console.log(`Content Control ${index + 1}: ${cc.title}`);
    });
});
```

---

### getHtml

**Kind:** `read`

Gets an HTML representation of the paragraph object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use Paragraph.getOoxml() and convert the returned XML to HTML.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve and display the HTML content of the first paragraph in the document body.

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

---

### getNext

**Kind:** `read`

Gets the next paragraph. Throws an ItemNotFound error if the paragraph is the last one.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Highlight the next paragraph after the first paragraph in the document by changing its font color to red.

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the next paragraph after the first one
    const nextParagraph = firstParagraph.getNext();
    
    // Change the font color of the next paragraph to red
    nextParagraph.font.color = "red";
    
    await context.sync();
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next paragraph. If the paragraph is the last one, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Check if a paragraph has a next paragraph and highlight it yellow if it exists, otherwise log that it's the last paragraph

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    const nextParagraph = paragraph.getNextOrNullObject();
    
    nextParagraph.load("isNullObject, text");
    await context.sync();
    
    if (nextParagraph.isNullObject) {
        console.log("This is the last paragraph in the document.");
    } else {
        nextParagraph.font.highlightColor = "yellow";
        console.log("Next paragraph highlighted: " + nextParagraph.text);
    }
    
    await context.sync();
});
```

---

### getOoxml

**Kind:** `read`

Gets the Office Open XML (OOXML) representation of the paragraph object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve and display the OOXML (Office Open XML) markup of the first paragraph in the document body.

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

---

### getPrevious

**Kind:** `read`

Gets the previous paragraph. Throws an ItemNotFound error if the paragraph is the first one.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get the previous paragraph from the current paragraph and highlight it in yellow

```typescript
await Word.run(async (context) => {
    // Get the current paragraph (e.g., the first paragraph in the document)
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.load("text");
    
    // Get the previous paragraph
    const previousParagraph = paragraph.getPrevious();
    previousParagraph.load("text");
    
    // Highlight the previous paragraph in yellow
    previousParagraph.font.highlightColor = "yellow";
    
    await context.sync();
    
    console.log("Previous paragraph text: " + previousParagraph.text);
});
```

---

### getPreviousOrNullObject

**Kind:** `read`

Gets the previous paragraph. If the paragraph is the first one, then this method returns an object with isNullObject = true. For further information, see “OrNullObject methods and properties”.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve and display the text of the second-to-last paragraph in the document, or indicate if no such paragraph exists.

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

---

### getRange

**Kind:** `read`

Gets the whole paragraph, or the starting or ending point of the paragraph, as a range.

#### Signature

**Parameters:**
- `rangeLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Retrieve and display all complete sentences from the current insertion point to the end of the paragraph.

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

---

### getReviewedText

**Kind:** `read`

Gets reviewed text based on ChangeTrackingVersion selection.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `changeTrackingVersion`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `changeTrackingVersion`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Get and display the original text of the first paragraph before any tracked changes were made

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the original text (before tracked changes)
    const originalText = paragraph.getReviewedText(Word.ChangeTrackingVersion.original);
    
    // Load the text property
    await context.sync();
    
    // Display the original text
    console.log("Original text: " + originalText.value);
});
```

---

### getText

**Kind:** `read`

Returns the text of the paragraph. This excludes equations, graphics (e.g., images, videos, drawings), and special characters that mark various content (e.g., for content controls, fields, comments, footnotes, endnotes). By default, hidden text and text marked as deleted are excluded.

#### Signature

**Parameters:**
- `options`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Retrieve and display the text content from the first paragraph in the document body

```typescript
await Word.run(async (context) => {
    // Get the first paragraph from the document body
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get the text from the paragraph
    const paragraphText = firstParagraph.getText();
    
    // Load the text value
    await context.sync();
    
    // Display the retrieved text
    console.log("Paragraph text: " + paragraphText.value);
});
```

---

### getTextRanges

**Kind:** `read`

Gets the text ranges in the paragraph by using punctuation marks and/or other ending marks.

#### Signature

**Parameters:**
- `endingMarks`: `None` (required)
- `trimSpacing`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Split a paragraph into individual sentences by detecting periods and question marks, then highlight each sentence with alternating colors.

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get text ranges split by periods and question marks
    const textRanges = paragraph.getTextRanges(['.', '?'], true);
    
    textRanges.load("items");
    await context.sync();
    
    // Highlight each sentence with alternating colors
    const colors = ["yellow", "lightblue"];
    for (let i = 0; i < textRanges.items.length; i++) {
        textRanges.items[i].font.highlightColor = colors[i % 2];
    }
    
    await context.sync();
});
```

---

### getTrackedChanges

**Kind:** `read`

Gets the collection of the TrackedChange objects in the paragraph.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get and display information about all tracked changes in the first paragraph of the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Get the tracked changes in the paragraph
    const trackedChanges = paragraph.getTrackedChanges();
    
    // Load properties of the tracked changes
    trackedChanges.load("items");
    
    await context.sync();
    
    // Display information about each tracked change
    console.log(`Found ${trackedChanges.items.length} tracked change(s) in the paragraph`);
    
    trackedChanges.items.forEach((change, index) => {
        change.load("type, author, date");
    });
    
    await context.sync();
    
    trackedChanges.items.forEach((change, index) => {
        console.log(`Change ${index + 1}: Type=${change.type}, Author=${change.author}, Date=${change.date}`);
    });
});
```

---

### indent

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Indent the first paragraph in the document to create a visual hierarchy

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Indent the paragraph
    firstParagraph.indent();
    
    await context.sync();
});
```

---

### indentCharacterWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Parameters:**
- `count`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Indent the first paragraph in the document by 10 character widths

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.indentCharacterWidth(10);
    
    await context.sync();
});
```

---

### indentFirstLineCharacterWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Parameters:**
- `count`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Indent the first line of the first paragraph in the document by 5 characters

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.indentFirstLineCharacterWidth(5);
    
    await context.sync();
});
```

---

### insertAnnotations

**Kind:** `create`

Inserts annotations on this Paragraph object.

#### Signature

**Parameters:**
- `annotations`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Add multiple color-coded critique annotations with suggestions to specific character ranges within the selected paragraph.

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

---

### insertBreak

**Kind:** `create`

Inserts a break at the specified location in the main document.

#### Signature

**Parameters:**
- `breakType`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a page break after the first paragraph in the document body.

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

**Example**: Insert a line break after the first paragraph in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-line-and-page-breaks.yaml

Word.run(async (context) => {
  context.document.body.paragraphs.getFirst().insertBreak(Word.BreakType.line, "After");

  await context.sync();
  console.log("success");
});
```

---

### insertCanvas

**Kind:** `create`

Inserts a floating canvas in front of text with its anchor at the beginning of the paragraph.

#### Signature

**Parameters:**
- `insertShapeOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a floating canvas with a rectangle shape at the beginning of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Insert a floating canvas at the beginning of the paragraph
    const canvas = firstParagraph.insertCanvas({
        width: 300,
        height: 200,
        left: 0,
        top: 0
    });
    
    // Sync to apply changes
    await context.sync();
    
    console.log("Canvas inserted at the beginning of the first paragraph");
});
```

---

### insertContentControl

**Kind:** `create`

Wraps the Paragraph object with a content control.

#### Signature

**Parameters:**
- `contentControlType`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Wrap the first paragraph of the document body in a rich text content control.

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

**Example**: Wrap each paragraph in the document with a content control and tag them alternately as "even" or "odd" based on their position.

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

### insertFileFromBase64

**Kind:** `create`

Inserts a document into the paragraph at the specified location.

#### Signature

**Parameters:**
- `base64File`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a document from a base64-encoded file at the end of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Base64-encoded document content (example: a simple .docx file)
    const base64File = "UEsDBBQABgAIAAAAIQDfpNJsWgEAACAFAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAAC...";
    
    // Insert the document at the end of the paragraph
    firstParagraph.insertFileFromBase64(base64File, Word.InsertLocation.end);
    
    await context.sync();
    
    console.log("Document inserted successfully at the end of the first paragraph");
});
```

---

### insertGeometricShape

**Kind:** `create`

Inserts a geometric shape in front of text with its anchor at the beginning of the paragraph.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `geometricShapeType`: `None` (required)
  - `insertShapeOptions`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `geometricShapeType`: `None` (required)
  - `insertShapeOptions`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Insert a blue rectangle shape at the beginning of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    const shapeOptions: Word.InsertShapeOptions = {
        width: 100,
        height: 50,
        left: 0,
        top: 0
    };
    
    const shape = firstParagraph.insertGeometricShape(
        Word.GeometricShapeType.rectangle,
        shapeOptions
    );
    
    shape.fill.setSolidColor("blue");
    
    await context.sync();
});
```

---

### insertHtml

**Kind:** `create`

Inserts HTML into the paragraph at the specified location.

#### Signature

**Parameters:**
- `html`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert HTML content containing bold text at the end of the first paragraph in the document body.

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

---

### insertInlinePictureFromBase64

**Kind:** `create`

Inserts a picture into the paragraph at the specified location.

#### Signature

**Parameters:**
- `base64EncodedImage`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a base64-encoded inline picture at the beginning of the first paragraph in the document body.

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

---

### insertOoxml

**Kind:** `create`

Inserts OOXML into the paragraph at the specified location.

#### Signature

**Parameters:**
- `ooxml`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a formatted text block with bold styling into the current paragraph using OOXML

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // OOXML string that defines a bold text run
    const ooxml = `
        <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
            <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
                <pkg:xmlData>
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                    </Relationships>
                </pkg:xmlData>
            </pkg:part>
            <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
                <pkg:xmlData>
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                        <w:body>
                            <w:p>
                                <w:r>
                                    <w:rPr>
                                        <w:b/>
                                    </w:rPr>
                                    <w:t>This is bold text inserted via OOXML</w:t>
                                </w:r>
                            </w:p>
                        </w:body>
                    </w:document>
                </pkg:xmlData>
            </pkg:part>
        </pkg:package>`;
    
    // Insert the OOXML at the end of the paragraph
    paragraph.insertOoxml(ooxml, Word.InsertLocation.end);
    
    await context.sync();
});
```

---

### insertParagraph

**Kind:** `create`

Inserts a paragraph at the specified location.

#### Signature

**Parameters:**
- `paragraphText`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a new paragraph with text after the currently selected paragraph

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const selection = context.document.getSelection();
    const paragraph = selection.paragraphs.getFirst();
    
    // Insert a new paragraph after the selected paragraph
    paragraph.insertParagraph("This is the newly inserted paragraph.", Word.InsertLocation.after);
    
    await context.sync();
});
```

---

### insertPictureFromBase64

**Kind:** `create`

Inserts a floating picture in front of text with its anchor at the beginning of the paragraph.

#### Signature

**Parameters:**
- `base64EncodedImage`: `None` (required)
- `insertShapeOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a company logo as a floating picture at the beginning of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Base64 encoded image string (example: a small PNG image)
    const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
    
    // Insert the picture as a floating image at the paragraph's beginning
    const picture = firstParagraph.insertPictureFromBase64(
        base64Image,
        {
            width: 100,
            height: 100,
            left: 50,
            top: 50
        }
    );
    
    await context.sync();
    console.log("Picture inserted successfully");
});
```

---

### insertTable

**Kind:** `create`

Inserts a table with the specified number of rows and columns.

#### Signature

**Parameters:**
- `rowCount`: `None` (required)
- `columnCount`: `None` (required)
- `insertLocation`: `None` (required)
- `values`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a 3x4 table after a paragraph and populate it with employee data including headers

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Define the table data (headers + 2 data rows)
    const tableData = [
        ["Name", "Department", "Position", "Years"],
        ["John Smith", "Engineering", "Developer", "3"],
        ["Jane Doe", "Marketing", "Manager", "5"]
    ];
    
    // Insert a 3x4 table after the paragraph
    const table = paragraph.insertTable(3, 4, Word.InsertLocation.after, tableData);
    
    // Optional: Format the header row
    table.rows.getFirst().font.bold = true;
    
    await context.sync();
});
```

---

### insertText

**Kind:** `create`

Inserts text into the paragraph at the specified location.

#### Signature

**Parameters:**
- `text`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Replace the last paragraph in the document with new text and format it with a black highlight and white font color.

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

---

### insertTextBox

**Kind:** `create`

Inserts a floating text box in front of text with its anchor at the beginning of the paragraph.

#### Signature

**Parameters:**
- `text`: `None` (required)
- `insertShapeOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a text box with placeholder text at the beginning of the first paragraph in the primary header section.

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

---

### joinList

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Remove a paragraph from its current list by joining it with the surrounding content

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Load the paragraph to check if it's in a list
    paragraph.load("isListItem");
    await context.sync();
    
    // If the paragraph is part of a list, remove it from the list
    if (paragraph.isListItem) {
        paragraph.joinList();
        await context.sync();
        
        console.log("Paragraph removed from list");
    } else {
        console.log("Paragraph is not part of a list");
    }
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `None` (required)

  **Returns:** `None`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Load and display the text content and style properties of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document body
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Load specific properties of the paragraph
    paragraph.load("text, style, font/size, font/name");
    
    // Sync to execute the load command
    await context.sync();
    
    // Now we can read the loaded properties
    console.log("Paragraph text:", paragraph.text);
    console.log("Paragraph style:", paragraph.style);
    console.log("Font size:", paragraph.font.size);
    console.log("Font name:", paragraph.font.name);
});
```

---

### next

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Parameters:**
- `count`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Get the paragraph that comes 2 positions after the first paragraph in the document and highlight it in yellow

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    const targetParagraph = firstParagraph.next(2);
    
    targetParagraph.font.highlightColor = "yellow";
    
    await context.sync();
});
```

---

### openOrCloseUp

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Toggle the spacing before a paragraph by opening or closing up the space for the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Toggle the spacing before the paragraph
    firstParagraph.openOrCloseUp();
    
    await context.sync();
    
    console.log("Paragraph spacing toggled successfully");
});
```

---

### openUp

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Add spacing before a paragraph by opening it up (increasing the space above it)

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Open up the paragraph (add spacing before it)
    paragraph.openUp();
    
    await context.sync();
});
```

---

### outdent

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Decrease the indentation level of the first paragraph in the document by one level

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Decrease the indentation level
    firstParagraph.outdent();
    
    await context.sync();
});
```

---

### outlineDemote

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Demote the first paragraph in the document to a lower outline level (e.g., from Heading 1 to Heading 2)

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.outlineDemote();
    
    await context.sync();
});
```

---

### outlineDemoteToBody

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Demote the first paragraph in the document to body text level in the outline hierarchy

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Demote the paragraph to body text level
    firstParagraph.outlineDemoteToBody();
    
    await context.sync();
    
    console.log("Paragraph demoted to body text level");
});
```

---

### outlinePromote

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Promote the first paragraph in the document to a higher outline level (e.g., from Heading 2 to Heading 1, or from body text to a heading)

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.outlinePromote();
    
    await context.sync();
    console.log("First paragraph promoted to higher outline level");
});
```

---

### previous

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Parameters:**
- `count`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Get the previous 2 paragraphs before the currently selected paragraph and highlight them in yellow

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const selection = context.document.getSelection();
    const paragraph = selection.paragraphs.getFirst();
    
    // Get the previous 2 paragraphs
    const previousParagraphs = paragraph.previous(2);
    previousParagraphs.load("text");
    
    // Highlight them in yellow
    previousParagraphs.font.highlightColor = "yellow";
    
    await context.sync();
    
    console.log("Highlighted previous paragraphs");
});
```

---

### reset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Reset a modified paragraph back to its original formatting and content state

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Make some modifications to the paragraph
    paragraph.font.bold = true;
    paragraph.font.size = 18;
    paragraph.alignment = Word.Alignment.centered;
    
    await context.sync();
    
    // Reset the paragraph to its original state
    paragraph.reset();
    
    await context.sync();
    
    console.log("Paragraph has been reset to its original state");
});
```

---

### resetAdvanceTo

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Reset the "advance to" setting for the first paragraph in the document to its default state

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Reset the advance to setting
    paragraph.resetAdvanceTo();
    
    await context.sync();
    
    console.log("Advanced to setting has been reset for the paragraph.");
});
```

---

### search

Performs a search with the specified SearchOptions on the scope of the paragraph object. The search results are a collection of range objects.

#### Signature

**Parameters:**
- `searchText`: `None` (required)
- `searchOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Find and highlight all occurrences of the word "TODO" in the first paragraph of the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Search for "TODO" in the paragraph
    const searchResults = firstParagraph.search("TODO", { matchCase: false });
    
    // Load the search results
    searchResults.load("font");
    await context.sync();
    
    // Highlight all found instances
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### select

Selects and navigates the Word UI to the paragraph.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `selectionMode`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `selectionMode`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Select the last paragraph in the document body, either selecting the entire paragraph or positioning the cursor at its end.

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

---

### selectNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Select the numbered list number/bullet of the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Select the number/bullet portion of the paragraph
    firstParagraph.selectNumber();
    
    await context.sync();
});
```

---

### separateList

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Separate the first paragraph from its list, converting it to a normal paragraph while keeping the remaining list items intact

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Separate this paragraph from its list
    firstParagraph.separateList();
    
    await context.sync();
    
    console.log("Paragraph has been separated from the list");
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `None` (required)
  - `options`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `properties`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Set multiple properties of a paragraph including indentation and font formatting, and copy all properties from one paragraph to another.

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

---

### space1

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Set the spacing after a paragraph to single spacing (1.0) for the first paragraph in the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Set spacing to single spacing (1.0)
    paragraph.space1();
    
    await context.sync();
});
```

---

### space1Pt5

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Set the line spacing of the first paragraph in the document to 1.5 lines

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Set line spacing to 1.5
    paragraph.space1Pt5();
    
    await context.sync();
});
```

---

### space2

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Set the spacing after a paragraph to 12 points using the preview space2() method

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Set spacing after the paragraph using the preview space2() method
    paragraph.space2();
    
    await context.sync();
    
    console.log("Spacing applied to paragraph");
});
```

---

### split

Splits the paragraph into child ranges by using delimiters.

#### Signature

**Parameters:**
- `delimiters`: `None` (required)
- `trimDelimiters`: `None` (required)
- `trimSpacing`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Split the first paragraph into individual words and sequentially highlight each word in yellow with a brief pause between highlights, removing the previous word's highlight.

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

---

### startNewList

Starts a new list with this paragraph. Fails if the paragraph is already a list item.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Create a new list starting from the second paragraph in the document, add items at different list levels, and insert a paragraph after the list that is not part of it.

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

---

### tabHangingIndent

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Parameters:**
- `count`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Apply a hanging indent of 2 tab stops to the first paragraph in the document

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.tabHangingIndent(2);
    
    await context.sync();
});
```

---

### tabIndent

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

#### Signature

**Parameters:**
- `count`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Indent the first paragraph in the document by 3 tab stops to the right

```typescript
await Word.run(async (context) => {
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.tabIndent(3);
    
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Paragraph object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ParagraphData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Load paragraph properties and serialize them to JSON for logging or external storage

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const paragraph = context.document.body.paragraphs.getFirst();
    
    // Load properties we want to serialize
    paragraph.load("text,style,alignment,firstLineIndent,leftIndent");
    
    await context.sync();
    
    // Convert the paragraph object to a plain JavaScript object
    const paragraphData = paragraph.toJSON();
    
    // Now you can use the plain object for logging, storage, or transmission
    console.log("Paragraph as JSON:", JSON.stringify(paragraphData, null, 2));
    console.log("Text:", paragraphData.text);
    console.log("Style:", paragraphData.style);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a paragraph object to maintain its reference across multiple sync calls while modifying its properties in different batches

```typescript
await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.load("text");
    
    // Track the paragraph to use it across multiple sync calls
    paragraph.track();
    
    await context.sync();
    console.log("Original text: " + paragraph.text);
    
    // First modification
    paragraph.font.bold = true;
    await context.sync();
    
    // Second modification - paragraph reference still valid due to tracking
    paragraph.font.color = "blue";
    await context.sync();
    
    // Third modification - still works
    paragraph.alignment = Word.Alignment.center;
    await context.sync();
    
    // Untrack when done to free up memory
    paragraph.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a paragraph object to work with it across multiple sync() calls, then untrack it to free memory when done processing

```typescript
await Word.run(async (context) => {
    // Get the first paragraph and track it
    const paragraph = context.document.body.paragraphs.getFirst();
    paragraph.track();
    
    // Load and sync to get paragraph properties
    paragraph.load("text");
    await context.sync();
    
    console.log("Paragraph text: " + paragraph.text);
    
    // Perform additional operations across multiple syncs
    paragraph.font.color = "blue";
    await context.sync();
    
    // Once done with the paragraph, untrack to release memory
    paragraph.untrack();
    await context.sync();
    
    console.log("Paragraph processing complete and memory released");
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
