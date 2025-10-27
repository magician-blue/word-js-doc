# InlinePicture

**Package:** `word`

**API Set:** WordApi 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents an inline picture.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Inserts an image anchored to the last paragraph.
await Word.run(async (context) => {
  context.document.body.paragraphs
    .getLast()
    .insertParagraph("", "After")
    .insertInlinePictureFromBase64(base64Image, "End");

  await context.sync();
});
```

## Properties

### altTextDescription

**Type:** `string`

**Since:** WordApi 1.1

Specifies a string that represents the alternative text associated with the inline image.

#### Examples

**Example**: Set the alternative text description for an inline picture to "Company logo showing a blue mountain peak"

```typescript
await Word.run(async (context) => {
    const firstPicture = context.document.body.inlinePictures.getFirst();
    firstPicture.altTextDescription = "Company logo showing a blue mountain peak";
    
    await context.sync();
});
```

---

### altTextTitle

**Type:** `string`

**Since:** WordApi 1.1

Specifies a string that contains the title for the inline image.

#### Examples

**Example**: Set the alt text title of the first inline picture in the document to "Company Logo"

```typescript
await Word.run(async (context) => {
    const firstPicture = context.document.body.inlinePictures.getFirst();
    firstPicture.altTextTitle = "Company Logo";
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from an inline picture to load and read its width property

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirstOrNullObject();
    
    // Access the request context associated with the inline picture
    const pictureContext = firstPicture.context;
    
    // Use the context to load properties
    firstPicture.load("width");
    
    await pictureContext.sync();
    
    if (!firstPicture.isNullObject) {
        console.log(`Picture width: ${firstPicture.width}`);
    }
});
```

---

### height

**Type:** `number`

**Since:** WordApi 1.1

Specifies a number that describes the height of the inline image.

#### Examples

**Example**: Set the height of an inline picture to 200 pixels

```typescript
await Word.run(async (context) => {
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const picture = inlinePictures.items[0];
        picture.height = 200;
        
        await context.sync();
    }
});
```

---

### hyperlink

**Type:** `string`

**Since:** WordApi 1.1

Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.

#### Examples

**Example**: Add a hyperlink to an inline picture that navigates to a specific website when clicked

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const picture = inlinePictures.items[0];
        
        // Set hyperlink to navigate to Microsoft's website
        picture.hyperlink = "https://www.microsoft.com";
        
        await context.sync();
        console.log("Hyperlink added to the inline picture");
    }
});
```

---

### imageFormat

**Type:** `Word.ImageFormat | "Unsupported" | "Undefined" | "Bmp" | "Jpeg" | "Gif" | "Tiff" | "Png" | "Icon" | "Exif" | "Wmf" | "Emf" | "Pict" | "Pdf" | "Svg"`

**Since:** WordApiDesktop 1.1

Gets the format of the inline image.

#### Examples

**Example**: Retrieve and display the dimensions, format, and Base64-encoded data of the first inline picture in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Gets the first image in the document.
await Word.run(async (context) => {
  const firstPicture: Word.InlinePicture = context.document.body.inlinePictures.getFirst();
  firstPicture.load("width, height, imageFormat");

  await context.sync();
  console.log(`Image dimensions: ${firstPicture.width} x ${firstPicture.height}`, `Image format: ${firstPicture.imageFormat}`);
  // Get the image encoded as Base64.
  const base64 = firstPicture.getBase64ImageSrc();

  await context.sync();
  console.log(base64.value);
});
```

---

### lockAspectRatio

**Type:** `boolean`

**Since:** WordApi 1.1

Specifies a value that indicates whether the inline image retains its original proportions when you resize it.

#### Examples

**Example**: Lock the aspect ratio of the first inline picture in the document to prevent distortion when resizing

```typescript
await Word.run(async (context) => {
    const firstPicture = context.document.body.inlinePictures.getFirst();
    firstPicture.lockAspectRatio = true;
    firstPicture.load("lockAspectRatio");
    
    await context.sync();
    console.log("Aspect ratio locked: " + firstPicture.lockAspectRatio);
});
```

---

### paragraph

**Type:** `Word.Paragraph`

**Since:** WordApi 1.2

Gets the parent paragraph that contains the inline image.

#### Examples

**Example**: Get the text content of the paragraph that contains the first inline picture in the document.

```typescript
await Word.run(async (context) => {
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const firstPicture = inlinePictures.items[0];
        const parentParagraph = firstPicture.paragraph;
        parentParagraph.load("text");
        
        await context.sync();
        
        console.log("Paragraph text: " + parentParagraph.text);
    }
});
```

---

### parentContentControl

**Type:** `Word.ContentControl`

**Since:** WordApi 1.1

Gets the content control that contains the inline image. Throws an ItemNotFound error if there isn't a parent content control.

#### Examples

**Example**: Check if an inline picture is inside a content control and highlight the parent content control in yellow if it exists.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    try {
        // Get the parent content control
        const parentContentControl = inlinePicture.parentContentControl;
        
        // Highlight the parent content control
        parentContentControl.appearance = "Tags";
        parentContentControl.color = "yellow";
        
        await context.sync();
        console.log("Parent content control highlighted successfully");
    } catch (error) {
        console.log("This inline picture is not inside a content control");
    }
});
```

---

### parentContentControlOrNullObject

**Type:** `Word.ContentControl`

**Since:** WordApi 1.3

Gets the content control that contains the inline image. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Examples

**Example**: Check if an inline picture is inside a content control and highlight the content control with a yellow background if it exists.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Get the parent content control (or null object if none exists)
    const parentContentControl = inlinePicture.parentContentControlOrNullObject;
    
    // Load the isNullObject property to check if it exists
    parentContentControl.load("isNullObject");
    
    await context.sync();
    
    // Check if the picture is inside a content control
    if (!parentContentControl.isNullObject) {
        // Picture is inside a content control - highlight it
        parentContentControl.font.highlightColor = "yellow";
        console.log("Picture is inside a content control - highlighted it");
    } else {
        console.log("Picture is not inside a content control");
    }
    
    await context.sync();
});
```

---

### parentTable

**Type:** `Word.Table`

**Since:** WordApi 1.3

Gets the table that contains the inline image. Throws an ItemNotFound error if it isn't contained in a table.

#### Examples

**Example**: Check if an inline picture is inside a table and if so, add a border around the parent table.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    await context.sync();

    if (inlinePictures.items.length > 0) {
        const picture = inlinePictures.items[0];
        
        try {
            // Get the parent table containing the picture
            const parentTable = picture.parentTable;
            parentTable.load("style");
            await context.sync();
            
            // Add a border to the parent table
            parentTable.setBorder(
                Word.BorderLocation.all,
                Word.BorderType.single,
                { color: "#0000FF", width: 2.0 }
            );
            
            await context.sync();
            console.log("Border added to the table containing the picture");
        } catch (error) {
            console.log("Picture is not contained in a table");
        }
    }
});
```

---

### parentTableCell

**Type:** `Word.TableCell`

**Since:** WordApi 1.3

Gets the table cell that contains the inline image. Throws an ItemNotFound error if it isn't contained in a table cell.

#### Examples

**Example**: Check if an inline picture is inside a table cell and highlight that cell with a yellow background color.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const picture = inlinePictures.items[0];
        
        try {
            // Get the parent table cell containing the picture
            const tableCell = picture.parentTableCell;
            tableCell.load("cellIndex, rowIndex");
            
            // Set the cell's background color to yellow
            tableCell.shadingColor = "#FFFF00";
            
            await context.sync();
            
            console.log(`Picture found in cell at row ${tableCell.rowIndex}, column ${tableCell.cellIndex}`);
        } catch (error) {
            console.log("Picture is not inside a table cell");
        }
    }
});
```

---

### parentTableCellOrNullObject

**Type:** `Word.TableCell`

**Since:** WordApi 1.3

Gets the table cell that contains the inline image. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Examples

**Example**: Check if an inline picture is inside a table cell and highlight the cell in yellow if it is.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    await context.sync();

    if (inlinePictures.items.length > 0) {
        const picture = inlinePictures.items[0];
        const tableCell = picture.parentTableCellOrNullObject;
        tableCell.load("isNullObject");
        await context.sync();

        if (!tableCell.isNullObject) {
            // Picture is in a table cell - highlight it
            tableCell.shadingColor = "yellow";
            console.log("Picture is inside a table cell - cell highlighted");
        } else {
            console.log("Picture is not inside a table cell");
        }
        
        await context.sync();
    }
});
```

---

### parentTableOrNullObject

**Type:** `Word.Table`

**Since:** WordApi 1.3

Gets the table that contains the inline image. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Examples

**Example**: Check if an inline picture is inside a table and if so, highlight the table with a light blue background color.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Get the parent table (or null object if not in a table)
    const parentTable = inlinePicture.parentTableOrNullObject;
    parentTable.load("isNullObject");
    
    await context.sync();
    
    // Check if the picture is inside a table
    if (!parentTable.isNullObject) {
        // Picture is in a table - highlight it
        parentTable.shadingColor = "#ADD8E6"; // Light blue
        console.log("Picture is inside a table. Table highlighted.");
    } else {
        console.log("Picture is not inside a table.");
    }
    
    await context.sync();
});
```

---

### width

**Type:** `number`

**Since:** WordApi 1.1

Specifies a number that describes the width of the inline image.

#### Examples

**Example**: Set the width of an inline picture to 200 pixels

```typescript
await Word.run(async (context) => {
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const picture = inlinePictures.items[0];
        picture.width = 200;
        
        await context.sync();
    }
});
```

---

## Methods

### delete

**Kind:** `delete`

Deletes the inline picture from the document.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Delete all inline pictures from the first paragraph of the document

```typescript
await Word.run(async (context) => {
    // Get the first paragraph in the document
    const firstParagraph = context.document.body.paragraphs.getFirst();
    
    // Get all inline pictures in the paragraph
    const inlinePictures = firstParagraph.inlinePictures;
    
    // Load the inline pictures collection
    context.load(inlinePictures);
    await context.sync();
    
    // Delete each inline picture
    for (let i = 0; i < inlinePictures.items.length; i++) {
        inlinePictures.items[i].delete();
    }
    
    await context.sync();
});
```

---

### getBase64ImageSrc

**Kind:** `read`

Gets the Base64-encoded string representation of the inline image.

#### Signature

**Returns:** `OfficeExtension.ClientResult<string>`

#### Examples

**Example**: Retrieve the first inline picture from the document body and obtain its Base64-encoded image source along with its dimensions and format.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/15-images/insert-and-get-pictures.yaml

// Gets the first image in the document.
await Word.run(async (context) => {
  const firstPicture: Word.InlinePicture = context.document.body.inlinePictures.getFirst();
  firstPicture.load("width, height, imageFormat");

  await context.sync();
  console.log(`Image dimensions: ${firstPicture.width} x ${firstPicture.height}`, `Image format: ${firstPicture.imageFormat}`);
  // Get the image encoded as Base64.
  const base64 = firstPicture.getBase64ImageSrc();

  await context.sync();
  console.log(base64.value);
});
```

---

### getNext

**Kind:** `read`

Gets the next inline image. Throws an ItemNotFound error if this inline image is the last one.

#### Signature

**Returns:** `Word.InlinePicture`

#### Examples

**Example**: Retrieve and display the alternative text title of the first inline picture in the document, or indicate if no inline pictures exist.

```typescript
// To use this snippet, add an inline picture to the document and assign it an alt text title.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the first inline picture.
    const firstPicture = context.document.body.inlinePictures.getFirstOrNullObject();

    // Queue a command to load the alternative text title of the picture.
    firstPicture.load('altTextTitle');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    if (firstPicture.isNullObject) {
        console.log('There are no inline pictures in this document.')
    } else {
        console.log(firstPicture.altTextTitle);
    }
});
```

---

### getNextOrNullObject

**Kind:** `read`

Gets the next inline image. If this inline image is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

#### Signature

**Returns:** `Word.InlinePicture`

#### Examples

**Example**: Retrieve and display the alternative text title of the first inline picture in the document, or indicate if no inline pictures exist.

```typescript
// To use this snippet, add an inline picture to the document and assign it an alt text title.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the first inline picture.
    const firstPicture = context.document.body.inlinePictures.getFirstOrNullObject();

    // Queue a command to load the alternative text title of the picture.
    firstPicture.load('altTextTitle');

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    if (firstPicture.isNullObject) {
        console.log('There are no inline pictures in this document.')
    } else {
        console.log(firstPicture.altTextTitle);
    }
});
```

---

### getRange

**Kind:** `read`

Gets the picture, or the starting or ending point of the picture, as a range.

#### Signature

**Parameters:**
- `rangeLocation`: `Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | "Whole" | "Start" | "End"` (optional)
  The range location must be 'Whole', 'Start', or 'End'.

**Returns:** `Word.Range`

#### Examples

**Example**: Get the range of the first inline picture in the document and highlight it with yellow color

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const firstPicture = inlinePictures.items[0];
        
        // Get the range of the inline picture
        const pictureRange = firstPicture.getRange("Whole");
        pictureRange.font.highlightColor = "yellow";
        
        await context.sync();
        console.log("Inline picture range highlighted");
    }
});
```

---

### insertBreak

Inserts a break at the specified location in the main document.

#### Signature

**Parameters:**
- `breakType`: `Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line"` (required)
  The break type to add.
- `insertLocation`: `Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"` (required)
  The value must be 'Before' or 'After'.

**Returns:** `void`

#### Examples

**Example**: Insert a page break after the first inline picture in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    const firstPicture = inlinePictures.getFirst();
    
    // Insert a page break after the picture
    firstPicture.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
    
    await context.sync();
});
```

---

### insertContentControl

**Kind:** `create`

Wraps the inline picture with a rich text content control.

#### Signature

**Returns:** `Word.ContentControl`

#### Examples

**Example**: Wrap an inline picture with a rich text content control and set its title to "Product Image"

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Wrap the inline picture with a content control
    const contentControl = inlinePicture.insertContentControl();
    contentControl.title = "Product Image";
    contentControl.tag = "productImage";
    
    await context.sync();
    
    console.log("Inline picture wrapped with content control");
});
```

---

### insertFileFromBase64

Inserts a document at the specified location.

#### Signature

**Parameters:**
- `base64File`: `string` (required)
  The Base64-encoded content of a .docx file.
- `insertLocation`: `Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"` (required)
  The value must be 'Before' or 'After'.

**Returns:** `Word.Range`

#### Examples

**Example**: Insert a company logo image from base64 data before an inline picture in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Base64 encoded image data (example: small PNG image)
    const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
    
    // Insert the image before the existing inline picture
    inlinePicture.insertFileFromBase64(base64Image, Word.InsertLocation.before);
    
    await context.sync();
});
```

---

### insertHtml

Inserts HTML at the specified location.

#### Signature

**Parameters:**
- `html`: `string` (required)
  The HTML to be inserted.
- `insertLocation`: `Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"` (required)
  The value must be 'Before' or 'After'.

**Returns:** `Word.Range`

#### Examples

**Example**: Insert a formatted HTML snippet with a heading and paragraph after an inline picture in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Insert HTML after the picture
    const htmlContent = "<h2>Image Caption</h2><p>This is a description of the image above.</p>";
    inlinePicture.insertHtml(htmlContent, Word.InsertLocation.after);
    
    await context.sync();
});
```

---

### insertInlinePictureFromBase64

**Kind:** `create`

Inserts an inline picture at the specified location.

#### Signature

**Parameters:**
- `base64EncodedImage`: `string` (required)
  The Base64-encoded image to be inserted.
- `insertLocation`: `Word.InsertLocation.replace | Word.InsertLocation.before | Word.InsertLocation.after | "Replace" | "Before" | "After"` (required)
  The value must be 'Replace', 'Before', or 'After'.

**Returns:** `Word.InlinePicture`

#### Examples

**Example**: Insert a company logo image from base64 encoding before an existing inline picture in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const firstPicture = context.document.body.inlinePictures.getFirst();
    
    // Base64 encoded image string (example: small red square PNG)
    const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";
    
    // Insert a new inline picture before the existing one
    const newPicture = firstPicture.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.before);
    
    await context.sync();
    console.log("Logo inserted successfully");
});
```

---

### insertOoxml

Inserts OOXML at the specified location.

#### Signature

**Parameters:**
- `ooxml`: `string` (required)
  The OOXML to be inserted.
- `insertLocation`: `Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"` (required)
  The value must be 'Before' or 'After'.

**Returns:** `Word.Range`

#### Examples

**Example**: Insert a formatted text box using OOXML before an existing inline picture in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // OOXML for a simple text box with formatted text
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
                                        <w:color w:val="FF0000"/>
                                    </w:rPr>
                                    <w:t>Important Note</w:t>
                                </w:r>
                            </w:p>
                        </w:body>
                    </w:document>
                </pkg:xmlData>
            </pkg:part>
        </pkg:package>`;
    
    // Insert the OOXML before the inline picture
    inlinePicture.insertOoxml(ooxml, Word.InsertLocation.before);
    
    await context.sync();
});
```

---

### insertParagraph

**Kind:** `create`

Inserts a paragraph at the specified location.

#### Signature

**Parameters:**
- `paragraphText`: `string` (required)
  The paragraph text to be inserted.
- `insertLocation`: `Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"` (required)
  The value must be 'Before' or 'After'.

**Returns:** `Word.Paragraph`

#### Examples

**Example**: Insert a paragraph with text below an inline picture in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Insert a paragraph after the inline picture
    inlinePicture.insertParagraph("This is a caption for the image above.", Word.InsertLocation.after);
    
    await context.sync();
});
```

---

### insertText

Inserts text at the specified location.

#### Signature

**Parameters:**
- `text`: `string` (required)
  Text to be inserted.
- `insertLocation`: `Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"` (required)
  The value must be 'Before' or 'After'.

**Returns:** `Word.Range`

#### Examples

**Example**: Insert the text "Figure 1: " before an inline picture in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Insert text before the inline picture
    inlinePicture.insertText("Figure 1: ", Word.InsertLocation.before);
    
    await context.sync();
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `options`: `Word.Interfaces.InlinePictureLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.InlinePicture`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.InlinePicture`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.InlinePicture`

#### Examples

**Example**: Load and display the width and height properties of the first inline picture in the document

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Load the width and height properties
    inlinePicture.load("width, height");
    
    // Sync to execute the load command
    await context.sync();
    
    // Display the loaded properties
    console.log(`Picture width: ${inlinePicture.width}`);
    console.log(`Picture height: ${inlinePicture.height}`);
});
```

---

### select

Selects the inline picture. This causes Word to scroll to the selection.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `selectionMode`: `Word.SelectionMode` (optional)
    The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `selectionMode`: `"Select" | "Start" | "End"` (optional)
    The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

  **Returns:** `void`

#### Examples

**Example**: Select the first inline picture in the document to highlight it and scroll it into view

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("items");
    
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const firstPicture = inlinePictures.items[0];
        
        // Select the inline picture with default selection mode
        firstPicture.select();
        
        await context.sync();
        console.log("First inline picture selected and scrolled into view");
    } else {
        console.log("No inline pictures found in the document");
    }
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.InlinePictureUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.InlinePicture` (required)

  **Returns:** `void`

#### Examples

**Example**: Update an inline picture's dimensions and alt text by setting multiple properties at once

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Set multiple properties at once
    inlinePicture.set({
        width: 200,
        height: 150,
        altTextTitle: "Company Logo",
        altTextDescription: "Official company logo with blue background"
    });
    
    await context.sync();
    console.log("Inline picture properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.InlinePicture object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.InlinePictureData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.InlinePictureData`

#### Examples

**Example**: Get the properties of an inline picture as a plain JavaScript object and log it to the console for debugging or serialization purposes.

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePictures = context.document.body.inlinePictures;
    inlinePictures.load("width,height,altTextTitle");
    
    await context.sync();
    
    if (inlinePictures.items.length > 0) {
        const firstPicture = inlinePictures.items[0];
        
        // Convert the InlinePicture object to a plain JavaScript object
        const pictureData = firstPicture.toJSON();
        
        // Log the plain object (useful for debugging or serialization)
        console.log("Picture data:", JSON.stringify(pictureData, null, 2));
        console.log("Width:", pictureData.width);
        console.log("Height:", pictureData.height);
    }
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.InlinePicture`

#### Examples

**Example**: Track an inline picture object to maintain its reference across multiple sync calls when modifying its properties in separate batches

```typescript
await Word.run(async (context) => {
    // Get the first inline picture in the document
    const inlinePicture = context.document.body.inlinePictures.getFirst();
    
    // Track the object to use it across multiple sync calls
    inlinePicture.track();
    
    // Load properties
    inlinePicture.load("width,height");
    await context.sync();
    
    // First batch: modify width
    console.log(`Original size: ${inlinePicture.width} x ${inlinePicture.height}`);
    inlinePicture.width = 200;
    await context.sync();
    
    // Second batch: modify height (object remains valid because it's tracked)
    inlinePicture.height = 150;
    await context.sync();
    
    console.log("Picture resized successfully");
    
    // Untrack when done to free memory
    inlinePicture.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.InlinePicture`

#### Examples

**Example**: Insert an inline picture, use it to get its dimensions, then untrack it to free memory after the operation is complete.

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Insert an inline picture
    const picture = body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.start);
    
    // Track the picture to work with it across multiple sync calls
    picture.track();
    
    // Load properties we need
    picture.load("width,height");
    await context.sync();
    
    // Use the picture properties
    console.log(`Picture dimensions: ${picture.width} x ${picture.height}`);
    
    // Untrack the picture to release memory since we're done with it
    picture.untrack();
    await context.sync();
});
```

---

## Source

- /en-us/javascript/api/word
