# Word.Body

**Package:** `word`

**API Set:** WordApi 1.1

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the body of a document or a section.

## Class Examples

```typescript
// Get the body object and read its font size.
await Word.run(async (context) => {
    // Create a proxy object for the document body.
    const body = context.document.body;
    body.load("font/size");

    await context.sync();

    console.log("Font size: " + body.font.size);
});
```

## Properties

### contentControls

**Type:** `Word.ContentControlCollection`

**Since:** WordApi 1.1

Gets the collection of rich text content control objects in the body.

#### Examples

**Example**: Find all content controls in the document body and highlight them by setting their background color to yellow.

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const contentControls = body.contentControls;
    
    contentControls.load("items");
    await context.sync();
    
    for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a document body to verify the add-in is properly connected to the Word host application

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Access the request context associated with the body object
    const requestContext = body.context;
    
    // Use the context to verify connection by loading a property
    body.load("text");
    await requestContext.sync();
    
    console.log("Successfully connected to Word. Body text length:", body.text.length);
    console.log("Request context type:", requestContext.constructor.name);
});
```

---

### endnotes

**Type:** `Word.NoteItemCollection`

**Since:** WordApi 1.5

Gets the collection of endnotes in the body.

#### Examples

**Example**: Get all endnotes from the document body and display the count and text of each endnote in the console.

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const endnotes = body.endnotes;
    
    endnotes.load("items");
    await context.sync();
    
    console.log(`Total endnotes: ${endnotes.items.length}`);
    
    for (let i = 0; i < endnotes.items.length; i++) {
        const endnote = endnotes.items[i];
        endnote.body.load("text");
        await context.sync();
        
        console.log(`Endnote ${i + 1}: ${endnote.body.text}`);
    }
});
```

---

### fields

**Type:** `Word.FieldCollection`

**Since:** WordApi 1.4

Gets the collection of field objects in the body.

#### Examples

**Example**: Retrieve and display all fields in the document body, showing their code and result values, or indicate if no fields exist.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets all fields in the document body.
await Word.run(async (context) => {
  const fields: Word.FieldCollection = context.document.body.fields.load("items");

  await context.sync();

  if (fields.items.length === 0) {
    console.log("No fields in this document.");
  } else {
    fields.load(["code", "result"]);
    await context.sync();

    for (let i = 0; i < fields.items.length; i++) {
      console.log(`Field ${i + 1}'s code: ${fields.items[i].code}`, `Field ${i + 1}'s result: ${JSON.stringify(fields.items[i].result)}`);
    }
  }
});
```

---

### font

**Type:** `Word.Font`

**Since:** WordApi 1.1

Gets the text format of the body. Use this to get and set font name, size, color and other properties.

#### Examples

**Example**: Retrieve and display the font size, font name, font color, and style properties of the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Gets the style and the font size, font name, and font color properties on the body object.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to load font and style information for the document body.
  body.load("font/size, font/name, font/color, style");

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  // Show font-related property values on the body object.
  const results =
    "Font size: " +
    body.font.size +
    "; Font name: " +
    body.font.name +
    "; Font color: " +
    body.font.color +
    "; Body style: " +
    body.style;

  console.log(results);
});
```

---

### footnotes

**Type:** `Word.NoteItemCollection`

**Since:** WordApi 1.5

Gets the collection of footnotes in the body.

#### Examples

**Example**: Retrieve and display the total count of footnotes present in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the footnotes in the document body.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("length");
  await context.sync();

  console.log("Number of footnotes in the document body: " + footnotes.items.length);
});
```

---

### inlinePictures

**Type:** `Word.InlinePictureCollection`

**Since:** WordApi 1.1

Gets the collection of InlinePicture objects in the body. The collection doesn't include floating images.

#### Examples

**Example**: Retrieve the first inline picture from the document body and log its dimensions, format, and Base64-encoded image data.

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

### lists

**Type:** `Word.ListCollection`

**Since:** WordApi 1.3

Gets the collection of list objects in the body.

#### Examples

**Example**: Retrieve and display the level types and level existences information for the first list in the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Gets information about the first list in the document.
await Word.run(async (context) => {
  const lists: Word.ListCollection = context.document.body.lists;
  lists.load("items");

  await context.sync();

  if (lists.items.length === 0) {
    console.warn("There are no lists in this document.");
    return;
  }
  
  // Get the first list.
  const list: Word.List = lists.getFirst();
  list.load("levelTypes,levelExistences");

  await context.sync();

  const levelTypes  = list.levelTypes;
  console.log("Level types of the first list:");
  for (let i = 0; i < levelTypes.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelTypes[i]}`);
  }

  const levelExistences = list.levelExistences;
  console.log("Level existences of the first list:");
  for (let i = 0; i < levelExistences.length; i++) {
    console.log(`- Level ${i + 1} (index ${i}): ${levelExistences[i]}`);
  }
});
```

---

### paragraphs

**Type:** `Word.ParagraphCollection`

**Since:** WordApi 1.1

Gets the collection of paragraph objects in the body.

#### Examples

**Example**: Count the number of occurrences of each unique word in the document and display the results.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/get-word-count.yaml

// Counts how many times each term appears in the document.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("text");
  await context.sync();

  // Split up the document text using existing spaces as the delimiter.
  let text = [];
  paragraphs.items.forEach((item) => {
    let paragraph = item.text.trim();
    if (paragraph) {
      paragraph.split(" ").forEach((term) => {
        let currentTerm = term.trim();
        if (currentTerm) {
          text.push(currentTerm);
        }
      });
    }
  });

  // Determine the list of unique terms.
  let makeTextDistinct = new Set(text);
  let distinctText = Array.from(makeTextDistinct);
  let allSearchResults = [];

  for (let i = 0; i < distinctText.length; i++) {
    let results = context.document.body.search(distinctText[i], { matchCase: true, matchWholeWord: true });
    results.load("text");

    // Map each search term with its results.
    let correlatedResults = {
      searchTerm: distinctText[i],
      hits: results
    };

    allSearchResults.push(correlatedResults);
  }

  await context.sync();

  // Display the count for each search term.
  allSearchResults.forEach((result) => {
    let length = result.hits.items.length;

    console.log("Search term: " + result.searchTerm + " => Count: " + length);
  });
});
```

---

### parentBody

**Type:** `Word.Body`

**Since:** WordApi 1.3

Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an `ItemNotFound` error if there isn't a parent body.

#### Examples

**Example**: Get the parent body of a table cell and highlight it in yellow to show the relationship between nested body elements.

```typescript
await Word.run(async (context) => {
    // Get the first table in the document
    const table = context.document.body.tables.getFirst();
    
    // Get the first cell's body
    const cellBody = table.getCell(0, 0).body;
    
    // Get the parent body of the cell (the table's parent body)
    const parentBody = cellBody.parentBody;
    
    // Highlight the parent body to visualize the relationship
    parentBody.font.highlightColor = "yellow";
    
    await context.sync();
    console.log("Parent body highlighted successfully");
});
```

---

### parentBodyOrNullObject

**Type:** `Word.Body`

**Since:** WordApi 1.3

Gets the parent body of the body. For example, a table cell body's parent body could be a header. If there isn't a parent body, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if the current body has a parent body and highlight the parent body in yellow if it exists (useful for identifying whether content is nested within a table cell, text box, or other container).

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const parentBody = body.parentBodyOrNullObject;
    
    // Load the isNullObject property to check if parent exists
    parentBody.load("isNullObject");
    await context.sync();
    
    if (!parentBody.isNullObject) {
        // Parent body exists - highlight it
        parentBody.font.highlightColor = "yellow";
        console.log("Parent body found and highlighted");
    } else {
        console.log("No parent body - this is the top-level document body");
    }
    
    await context.sync();
});
```

---

### parentContentControl

**Type:** `Word.ContentControl`

**Since:** WordApi 1.1

Gets the content control that contains the body. Throws an `ItemNotFound` error if there isn't a parent content control.

#### Examples

**Example**: Check if the document body is inside a content control and highlight the parent content control in yellow if it exists.

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    try {
        const parentContentControl = body.parentContentControl;
        parentContentControl.load("title");
        await context.sync();
        
        // Highlight the parent content control
        parentContentControl.appearance = "Tags";
        parentContentControl.color = "yellow";
        
        await context.sync();
        console.log(`Body is inside content control: ${parentContentControl.title}`);
    } catch (error) {
        if (error.code === "ItemNotFound") {
            console.log("Document body is not inside a content control");
        } else {
            throw error;
        }
    }
});
```

---

### parentContentControlOrNullObject

**Type:** `Word.ContentControl`

**Since:** WordApi 1.3

Gets the content control that contains the body. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if the document body is inside a content control, and if so, change the content control's title to "Document Container"

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const parentContentControl = body.parentContentControlOrNullObject;
    
    // Load the isNullObject property to check if parent exists
    parentContentControl.load("isNullObject, title");
    await context.sync();
    
    if (!parentContentControl.isNullObject) {
        // Body is inside a content control, update its title
        parentContentControl.title = "Document Container";
        console.log("Body is inside a content control. Title updated.");
    } else {
        console.log("Body is not inside a content control.");
    }
    
    await context.sync();
});
```

---

### parentSection

**Type:** `Word.Section`

**Since:** WordApi 1.3

Gets the parent section of the body. Throws an `ItemNotFound` error if there isn't a parent section.

#### Examples

**Example**: Get the header text from the parent section of the document body and display it in the console.

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const parentSection = body.parentSection;
    
    // Load the header of the parent section
    parentSection.load("getHeader");
    await context.sync();
    
    const header = parentSection.getHeader(Word.HeaderFooterType.primary);
    header.load("text");
    await context.sync();
    
    console.log("Parent section header text: " + header.text);
});
```

---

### parentSectionOrNullObject

**Type:** `Word.Section`

**Since:** WordApi 1.3

Gets the parent section of the body. If there isn't a parent section, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Examples

**Example**: Check if the body belongs to a section and display the section's header text, or show a message if there is no parent section (e.g., for the main document body).

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    const parentSection = body.parentSectionOrNullObject;
    parentSection.load("isNullObject");
    
    await context.sync();
    
    if (parentSection.isNullObject) {
        console.log("This body has no parent section (it's the main document body)");
    } else {
        const header = parentSection.getHeader(Word.HeaderFooterType.primary);
        header.load("text");
        await context.sync();
        
        console.log("Parent section header text: " + header.text);
    }
});
```

---

### shapes

**Type:** `Word.ShapeCollection`

**Since:** WordApiDesktop 1.2

Gets the collection of shape objects in the body, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

#### Examples

**Example**: Retrieve all shapes from the document body and log the properties of each shape that is a text box to the console.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Gets text boxes in main document.
  const shapes: Word.ShapeCollection = context.document.body.shapes;
  shapes.load();
  await context.sync();

  if (shapes.items.length > 0) {
    shapes.items.forEach(function(shape, index) {
      if (shape.type === Word.ShapeType.textBox) {
        console.log(`Shape ${index} in the main document has a text box. Properties:`, shape);
      }
    });
  } else {
    console.log("No shapes found in main document.");
  }
});
```

---

### style

**Type:** `string`

**Since:** WordApi 1.1

Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

#### Examples

**Example**: Set the document body to use a custom style named "MyCustomBodyStyle"

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    body.style = "MyCustomBodyStyle";
    
    await context.sync();
});
```

---

### styleBuiltIn

**Type:** `Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"`

**Since:** WordApi 1.3

Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

#### Examples

**Example**: Apply the "Title" built-in style to the document body to format all body content with the Title style

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    body.styleBuiltIn = Word.BuiltInStyleName.title;
    
    await context.sync();
});
```

---

### tables

**Type:** `Word.TableCollection`

**Since:** WordApi 1.3

Gets the collection of table objects in the body.

#### Examples

**Example**: Retrieve and display the text content from the first cell of the first table in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/table-cell-access.yaml

// Gets the content of the first cell in the first table.
await Word.run(async (context) => {
  const firstCell: Word.Body = context.document.body.tables.getFirst().getCell(0, 0).body;
  firstCell.load("text");

  await context.sync();
  console.log("First cell's text is: " + firstCell.text);
});
```

---

### text

**Type:** `string`

**Since:** WordApi 1.1

Gets the text of the body. Use the insertText method to insert text.

#### Examples

**Example**: Retrieve and display the plain text content of the document body in the console.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Gets the text content of the body.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to load the text in document body.
  body.load("text");

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Body contents (text): " + body.text);
});
```

---

### type

**Type:** `Word.BodyType | "Unknown" | "MainDoc" | "Section" | "Header" | "Footer" | "TableCell" | "Footnote" | "Endnote" | "NoteItem" | "Shape"`

**Since:** WordApi 1.3

Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Additional types âFootnoteâ, âEndnoteâ, and âNoteItemâ are supported in WordAPIOnline 1.1 and later.

#### Examples

**Example**: Retrieve and display the item type and body type of a specific footnote in the document based on a user-provided reference number.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the referenced note's item type and body type, which are both "Footnote".
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.body.footnotes;
  footnotes.load("items");
  await context.sync();

  const referenceNumber = (document.getElementById("input-reference") as HTMLInputElement).value;
  const mark = (referenceNumber as number) - 1;
  const item: Word.NoteItem = footnotes.items[mark];
  console.log(`Note type of footnote ${referenceNumber}: ${item.type}`);

  item.body.load("type");
  await context.sync();

  console.log(`Body type of note: ${item.body.type}`);
});
```

---

## Methods

### clear

**Kind:** `delete`

Clears the contents of the body object. The user can perform the undo operation on the cleared content.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Remove all content from the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Clears out the content from the document body.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to clear the contents of the body.
  body.clear();

  console.log("Cleared the body contents.");
});

// The Silly stories add-in sample shows how the clear method can be used to clear the contents of a document.
// https://aka.ms/sillystorywordaddin
```

---

### getComments

**Kind:** `read`

Gets comments associated with the body.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve all comments from the document body and display them in the console.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the comments in the document body.
await Word.run(async (context) => {
  const comments: Word.CommentCollection = context.document.body.getComments();

  // Load objects to log in the console.
  comments.load();
  await context.sync();

  console.log("All comments:", comments);
});
```

---

### getContentControls

**Kind:** `read`

Gets the currently supported content controls in the body.

#### Signature

**Parameters:**
- `options`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Get all content controls in the document body and display their titles in the console

```typescript
await Word.run(async (context) => {
    // Get all content controls in the document body
    const contentControls = context.document.body.getContentControls();
    
    // Load the title property for each content control
    contentControls.load("title");
    
    await context.sync();
    
    // Display the titles of all content controls
    console.log(`Found ${contentControls.items.length} content control(s):`);
    contentControls.items.forEach((cc, index) => {
        console.log(`${index + 1}. ${cc.title || "(No title)"}`);
    });
});
```

---

### getHtml

**Kind:** `read`

Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve the HTML representation of the document body and display it in the console.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Gets the HTML that represents the content of the body.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to get the HTML contents of the body.
  const bodyHTML = body.getHtml();

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Body contents (HTML): " + bodyHTML.value);
});
```

---

### getOoxml

**Kind:** `read`

Gets the OOXML (Office Open XML) representation of the body object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve the OOXML representation of the document body and output it to the console.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Gets the OOXML that represents the content of the body.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to get the OOXML contents of the body.
  const bodyOOXML = body.getOoxml();

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Body contents (OOXML): " + bodyOOXML.value);
});
```

---

### getRange

**Kind:** `read`

Gets the whole body, or the starting or ending point of the body, as a range.

#### Signature

**Parameters:**
- `rangeLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Highlight the entire document body with yellow color to mark it for review

```typescript
await Word.run(async (context) => {
    // Get the body of the document
    const body = context.document.body;
    
    // Get the entire body as a range using getRange()
    const bodyRange = body.getRange(Word.RangeLocation.whole);
    
    // Apply yellow highlight to the entire body
    bodyRange.font.highlightColor = "yellow";
    
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
  - `changeTrackingVersion`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `changeTrackingVersion`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Get and display the original text of a document before any tracked changes were applied

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Get the original text (before tracked changes)
    const originalText = body.getReviewedText(Word.ChangeTrackingVersion.original);
    
    await context.sync();
    
    console.log("Original text:", originalText.value);
});
```

---

### getTrackedChanges

**Kind:** `read`

Gets the collection of the TrackedChange objects in the body.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve all tracked changes from the document body and log them to the console.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-tracked-changes.yaml

// Gets all tracked changes.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  const trackedChanges: Word.TrackedChangeCollection = body.getTrackedChanges();
  trackedChanges.load();
  await context.sync();

  console.log(trackedChanges);
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

**Example**: Insert a page break at the beginning of the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Inserts a page break at the beginning of the document.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to insert a page break at the start of the document body.
  body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Added a page break at the start of the document body.");
});
```

---

### insertContentControl

**Kind:** `create`

Wraps the Body object with a content control.

#### Signature

**Parameters:**
- `contentControlType`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Wrap the entire document body in a content control.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Creates a content control using the document body.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to wrap the body in a content control.
  body.insertContentControl();

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Wrapped the body in a content control.");
});
```

---

### insertFileFromBase64

**Kind:** `create`

Inserts a document into the body at the specified location.

#### Signature

**Parameters:**
- `base64File`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert the contents of an external Word document (provided as a Base64-encoded string) at the beginning of the current document's body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Inserts the body from the external document at the beginning of this document.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to insert the Base64-encoded string representation of the body of the selected .docx file at the beginning of the current document.
  body.insertFileFromBase64(externalDocument, Word.InsertLocation.start);

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Added Base64-encoded text to the beginning of the document body.");
});
```

---

### insertHtml

**Kind:** `create`

Inserts HTML at the specified location.

#### Signature

**Parameters:**
- `html`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert HTML content containing bold text at the beginning of the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Inserts the HTML at the beginning of this document.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to insert HTML at the beginning of the document.
  body.insertHtml("<strong>This is text inserted with body.insertHtml()</strong>", Word.InsertLocation.start);

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("HTML added to the beginning of the document body.");
});
```

---

### insertInlinePictureFromBase64

**Kind:** `create`

Inserts a picture into the body at the specified location.

#### Signature

**Parameters:**
- `base64EncodedImage`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert a Base64-encoded inline picture at the beginning of the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Inserts an image inline at the beginning of this document.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Base64-encoded image to insert inline.
  const base64EncodedImg =
    "iVBORw0KGgoAAAANSUhEUgAAAB4AAAANCAIAAAAxEEnAAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACFSURBVDhPtY1BEoQwDMP6/0+XgIMTBAeYoTqso9Rkx1zG+tNj1H94jgGzeNSjteO5vtQQuG2seO0av8LzGbe3anzRoJ4ybm/VeKEerAEbAUpW4aWQCmrGFWykRzGBCnYy2ha3oAIq2MloW9yCCqhgJ6NtcQsqoIKdjLbFLaiACnYyf2fODbrjZcXfr2F4AAAAAElFTkSuQmCC";

  // Queue a command to insert a Base64-encoded image at the beginning of the current document.
  body.insertInlinePictureFromBase64(base64EncodedImg, Word.InsertLocation.start);

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Added a Base64-encoded image to the beginning of the document body.");
});
```

---

### insertOoxml

**Kind:** `create`

Inserts OOXML at the specified location.

#### Signature

**Parameters:**
- `ooxml`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert formatted text with custom font size, color, line spacing, and paragraph spacing at the beginning of the document body using OOXML.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Inserts OOXML at the beginning of this document.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to insert OOXML at the beginning of the body.
  body.insertOoxml(
    "<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>",
    Word.InsertLocation.start
  );

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Added OOXML to the beginning of the document body.");
});

// Read "Understand when and how to use Office Open XML in your Word add-in" for guidance on working with OOXML.
// https://learn.microsoft.com/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml

// The Word-Add-in-DocumentAssembly sample shows how you can use this API to assemble a document.
// https://github.com/OfficeDev/Word-Add-in-DocumentAssembly
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

**Example**: Insert a paragraph with custom text at the end of the document body and format it with italic, blue, 30-point Berlin Sans FB font.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/insert-formatted-text.yaml

await Word.run(async (context) => {
  // Second sentence, let's insert it as a paragraph after the previously inserted one.
  const secondSentence: Word.Paragraph = context.document.body.insertParagraph(
    "This is the first text with a custom style.",
    "End"
  );
  secondSentence.font.set({
    bold: false,
    italic: true,
    name: "Berlin Sans FB",
    color: "blue",
    size: 30
  });

  await context.sync();
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

**Example**: Insert a 2x3 table at the start of the document body with city and fruit data, apply the Grid Table 5 Dark - Accent 2 style, and disable first column formatting.

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

---

### insertText

**Kind:** `create`

Inserts text into the body at the specified location.

#### Signature

**Parameters:**
- `text`: `None` (required)
- `insertLocation`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Insert text at the beginning of the document body.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Inserts text at the beginning of this document.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to insert text at the beginning of the current document.
  body.insertText('This is text inserted with body.insertText()', Word.InsertLocation.start);

  // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
  await context.sync();

  console.log("Text added to the beginning of the document body.");
});
```

---

### load

**Kind:** `load`

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

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

**Example**: Load and display the text content from the document body

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Load the text property of the body
    body.load("text");
    
    // Synchronize to read the loaded property
    await context.sync();
    
    // Now we can access the text property
    console.log("Document body text:", body.text);
});
```

---

### search

Performs a search with the specified SearchOptions on the scope of the body object. The search results are a collection of range objects.

#### Signature

**Parameters:**
- `searchText`: `None` (required)
- `searchOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Search for text in the document body using basic text matching or wildcard patterns and highlight all matching results.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/search.yaml

// Does a basic text search and highlights matches in the document.
await Word.run(async (context) => {
  const results : Word.RangeCollection = context.document.body.search("extend");
  results.load("length");

  await context.sync();

  // Let's traverse the search results and highlight matches.
  for (let i = 0; i < results.items.length; i++) {
    results.items[i].font.highlightColor = "yellow";
  }

  await context.sync();
});

...

// Does a wildcard search and highlights matches in the document.
await Word.run(async (context) => {
  // Construct a wildcard expression and set matchWildcards to true in order to use wildcards.
  const results : Word.RangeCollection = context.document.body.search("$*.[0-9][0-9]", { matchWildcards: true });
  results.load("length");

  await context.sync();

  // Let's traverse the search results and highlight matches.
  for (let i = 0; i < results.items.length; i++) {
    results.items[i].font.highlightColor = "red";
    results.items[i].font.color = "white";
  }

  await context.sync();
});
```

---

### select

Selects the body and navigates the Word UI to it.

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

**Example**: Select the entire document body and move the Word UI focus to it.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-body.yaml

// Selects the entire body.
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
  // Create a proxy object for the document body.
  const body: Word.Body = context.document.body;

  // Queue a command to select the document body.
  // The Word UI will move to the selected document body.
  body.select();

  console.log("Selected the document body.");
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

**Example**: Configure multiple body properties at once to set the style and alignment for the document body

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Set multiple properties at once using the set() method
    body.set({
        style: "Normal",
        styleBuiltIn: Word.BuiltInStyleName.normal
    });
    
    await context.sync();
    console.log("Body properties have been set successfully.");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Body` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BodyData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Serialize the document body properties to JSON format for logging or data transfer purposes

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    
    // Load properties you want to serialize
    body.load("text,style,font/name,font/size");
    
    await context.sync();
    
    // Convert the body object to a plain JavaScript object
    const bodyData = body.toJSON();
    
    // Now you can use the plain object for logging, storage, or transfer
    console.log("Body as JSON:", JSON.stringify(bodyData, null, 2));
    console.log("Body text:", bodyData.text);
    console.log("Body style:", bodyData.style);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Track a document body object to maintain its reference across multiple sync calls when modifying its properties in separate operations

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    
    // Track the body object to use it across multiple sync calls
    body.track();
    
    await context.sync();
    console.log("Original text length:", body.text.length);
    
    // Make changes in a separate sync operation
    body.insertParagraph("New paragraph at the end", Word.InsertLocation.end);
    
    await context.sync();
    
    // The tracked object can still be used reliably
    body.load("text");
    await context.sync();
    console.log("Updated text length:", body.text.length);
    
    // Untrack when done to free up memory
    body.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get the document body text and then untrack the body object to free memory after use

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    body.track();
    body.load("text");
    
    await context.sync();
    
    console.log("Body text:", body.text);
    
    // Release memory associated with the tracked body object
    body.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
