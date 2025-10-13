# Word.Body class

Package: [word](/en-us/javascript/api/word)

Represents the body of a document or a section.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.1]

### Examples
```TypeScript
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
- [contentControls](#word-word-body-contentcontrols-member)  
  Gets the collection of rich text content control objects in the body.
- [context](#word-word-body-context-member)  
  The request context associated with the object. This connects the add-in's process to the Office host application's process.
- [endnotes](#word-word-body-endnotes-member)  
  Gets the collection of endnotes in the body.
- [fields](#word-word-body-fields-member)  
  Gets the collection of field objects in the body.
- [font](#word-word-body-font-member)  
  Gets the text format of the body. Use this to get and set font name, size, color and other properties.
- [footnotes](#word-word-body-footnotes-member)  
  Gets the collection of footnotes in the body.
- [inlinePictures](#word-word-body-inlinepictures-member)  
  Gets the collection of InlinePicture objects in the body. The collection doesn't include floating images.
- [lists](#word-word-body-lists-member)  
  Gets the collection of list objects in the body.
- [paragraphs](#word-word-body-paragraphs-member)  
  Gets the collection of paragraph objects in the body.
- [parentBody](#word-word-body-parentbody-member)  
  Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an `ItemNotFound` error if there isn't a parent body.
- [parentBodyOrNullObject](#word-word-body-parentbodyornullobject-member)  
  Gets the parent body of the body. For example, a table cell body's parent body could be a header. If there isn't a parent body, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [parentContentControl](#word-word-body-parentcontentcontrol-member)  
  Gets the content control that contains the body. Throws an `ItemNotFound` error if there isn't a parent content control.
- [parentContentControlOrNullObject](#word-word-body-parentcontentcontrolornullobject-member)  
  Gets the content control that contains the body. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [parentSection](#word-word-body-parentsection-member)  
  Gets the parent section of the body. Throws an `ItemNotFound` error if there isn't a parent section.
- [parentSectionOrNullObject](#word-word-body-parentsectionornullobject-member)  
  Gets the parent section of the body. If there isn't a parent section, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- [shapes](#word-word-body-shapes-member)  
  Gets the collection of shape objects in the body, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.
- [style](#word-word-body-style-member)  
  Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.
- [styleBuiltIn](#word-word-body-stylebuiltin-member)  
  Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.
- [tables](#word-word-body-tables-member)  
  Gets the collection of table objects in the body.
- [text](#word-word-body-text-member)  
  Gets the text of the body. Use the insertText method to insert text.
- [type](#word-word-body-type-member)  
  Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Additional types âFootnoteâ, âEndnoteâ, and âNoteItemâ are supported in WordAPIOnline 1.1 and later.

## Methods
- [clear()](#word-word-body-clear-member1)  
  Clears the contents of the body object. The user can perform the undo operation on the cleared content.
- [getComments()](#word-word-body-getcomments-member1)  
  Gets comments associated with the body.
- [getContentControls(options)](#word-word-body-getcontentcontrols-member1)  
  Gets the currently supported content controls in the body.
- [getHtml()](#word-word-body-gethtml-member1)  
  Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML.
- [getOoxml()](#word-word-body-getooxml-member1)  
  Gets the OOXML (Office Open XML) representation of the body object.
- [getRange(rangeLocation)](#word-word-body-getrange-member1)  
  Gets the whole body, or the starting or ending point of the body, as a range.
- [getReviewedText(changeTrackingVersion)](#word-word-body-getreviewedtext-member1)  
  Gets reviewed text based on ChangeTrackingVersion selection.
- [getReviewedText(changeTrackingVersion)](#word-word-body-getreviewedtext-member2)  
  Gets reviewed text based on ChangeTrackingVersion selection.
- [getTrackedChanges()](#word-word-body-gettrackedchanges-member1)  
  Gets the collection of the TrackedChange objects in the body.
- [insertBreak(breakType, insertLocation)](#word-word-body-insertbreak-member1)  
  Inserts a break at the specified location in the main document.
- [insertContentControl(contentControlType)](#word-word-body-insertcontentcontrol-member1)  
  Wraps the Body object with a content control.
- [insertFileFromBase64(base64File, insertLocation)](#word-word-body-insertfilefrombase64-member1)  
  Inserts a document into the body at the specified location.
- [insertHtml(html, insertLocation)](#word-word-body-inserthtml-member1)  
  Inserts HTML at the specified location.
- [insertInlinePictureFromBase64(base64EncodedImage, insertLocation)](#word-word-body-insertinlinepicturefrombase64-member1)  
  Inserts a picture into the body at the specified location.
- [insertOoxml(ooxml, insertLocation)](#word-word-body-insertooxml-member1)  
  Inserts OOXML at the specified location.
- [insertParagraph(paragraphText, insertLocation)](#word-word-body-insertparagraph-member1)  
  Inserts a paragraph at the specified location.
- [insertTable(rowCount, columnCount, insertLocation, values)](#word-word-body-inserttable-member1)  
  Inserts a table with the specified number of rows and columns.
- [insertText(text, insertLocation)](#word-word-body-inserttext-member1)  
  Inserts text into the body at the specified location.
- [load(options)](#word-word-body-load-member1)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNames)](#word-word-body-load-member2)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [load(propertyNamesAndPaths)](#word-word-body-load-member3)  
  Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- [search(searchText, searchOptions)](#word-word-body-search-member1)  
  Performs a search with the specified SearchOptions on the scope of the body object. The search results are a collection of range objects.
- [select(selectionMode)](#word-word-body-select-member1)  
  Selects the body and navigates the Word UI to it.
- [select(selectionMode)](#word-word-body-select-member2)  
  Selects the body and navigates the Word UI to it.
- [set(properties, options)](#word-word-body-set-member1)  
  Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- [set(properties)](#word-word-body-set-member2)  
  Sets multiple properties on the object at the same time, based on an existing loaded object.
- [toJSON()](#word-word-body-tojson-member1)  
  Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Body` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BodyData`) that contains shallow copies of any loaded child properties from the original object.
- [track()](#word-word-body-track-member1)  
  Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- [untrack()](#word-word-body-untrack-member1)  
  Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Events
- [onCommentAdded](#word-word-body-oncommentadded-member)  
  Occurs when new comments are added.
- [onCommentChanged](#word-word-body-oncommentchanged-member)  
  Occurs when a comment or its reply is changed.
- [onCommentDeleted](#word-word-body-oncommentdeleted-member)  
  Occurs when comments are deleted.
- [onCommentDeselected](#word-word-body-oncommentdeselected-member)  
  Occurs when a comment is deselected.
- [onCommentSelected](#word-word-body-oncommentselected-member)  
  Occurs when a comment is selected.

## Property Details

### contentControls
Gets the collection of rich text content control objects in the body.

```typescript
readonly contentControls: Word.ContentControlCollection;
```

- Property value: [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks  
[API set: WordApi 1.1]

---

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### endnotes
Gets the collection of endnotes in the body.

```typescript
readonly endnotes: Word.NoteItemCollection;
```

- Property value: [Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

Remarks  
[API set: WordApi 1.5]

---

### fields
Gets the collection of field objects in the body.

```typescript
readonly fields: Word.FieldCollection;
```

- Property value: [Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

Remarks  
[API set: WordApi 1.4]

#### Examples
```TypeScript
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
Gets the text format of the body. Use this to get and set font name, size, color and other properties.

```typescript
readonly font: Word.Font;
```

- Property value: [Word.Font](/en-us/javascript/api/word/word.font)

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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
Gets the collection of footnotes in the body.

```typescript
readonly footnotes: Word.NoteItemCollection;
```

- Property value: [Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

Remarks  
[API set: WordApi 1.5]

#### Examples
```TypeScript
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
Gets the collection of InlinePicture objects in the body. The collection doesn't include floating images.

```typescript
readonly inlinePictures: Word.InlinePictureCollection;
```

- Property value: [Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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
Gets the collection of list objects in the body.

```typescript
readonly lists: Word.ListCollection;
```

- Property value: [Word.ListCollection](/en-us/javascript/api/word/word.listcollection)

Remarks  
[API set: WordApi 1.3]

#### Examples
```TypeScript
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
Gets the collection of paragraph objects in the body.

```typescript
readonly paragraphs: Word.ParagraphCollection;
```

- Property value: [Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

Remarks  
[API set: WordApi 1.1]  
Important: Paragraphs in tables aren't returned for requirement sets 1.1 and 1.2. From requirement set 1.3, paragraphs in tables are also returned.

#### Examples
```TypeScript
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
Gets the parent body of the body. For example, a table cell body's parent body could be a header. Throws an `ItemNotFound` error if there isn't a parent body.

```typescript
readonly parentBody: Word.Body;
```

- Property value: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks  
[API set: WordApi 1.3]

---

### parentBodyOrNullObject
Gets the parent body of the body. For example, a table cell body's parent body could be a header. If there isn't a parent body, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentBodyOrNullObject: Word.Body;
```

- Property value: [Word.Body](/en-us/javascript/api/word/word.body)

Remarks  
[API set: WordApi 1.3]

---

### parentContentControl
Gets the content control that contains the body. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
readonly parentContentControl: Word.ContentControl;
```

- Property value: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks  
[API set: WordApi 1.1]

---

### parentContentControlOrNullObject
Gets the content control that contains the body. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentContentControlOrNullObject: Word.ContentControl;
```

- Property value: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks  
[API set: WordApi 1.3]

---

### parentSection
Gets the parent section of the body. Throws an `ItemNotFound` error if there isn't a parent section.

```typescript
readonly parentSection: Word.Section;
```

- Property value: [Word.Section](/en-us/javascript/api/word/word.section)

Remarks  
[API set: WordApi 1.3]

---

### parentSectionOrNullObject
Gets the parent section of the body. If there isn't a parent section, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentSectionOrNullObject: Word.Section;
```

- Property value: [Word.Section](/en-us/javascript/api/word/word.section)

Remarks  
[API set: WordApi 1.3]

---

### shapes
Gets the collection of shape objects in the body, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

```typescript
readonly shapes: Word.ShapeCollection;
```

- Property value: [Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

Remarks  
[API set: WordApiDesktop 1.2]

#### Examples
```TypeScript
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
Specifies the style name for the body. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style: string;
```

- Property value: string

Remarks  
[API set: WordApi 1.1]

---

### styleBuiltIn
Specifies the built-in style name for the body. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

- Property value: [Word.BuiltInStyleName](/en-us/javascript/api/word/word.builtinstylename) | the literal strings listed above

Remarks  
[API set: WordApi 1.3]

---

### tables
Gets the collection of table objects in the body.

```typescript
readonly tables: Word.TableCollection;
```

- Property value: [Word.TableCollection](/en-us/javascript/api/word/word.tablecollection)

Remarks  
[API set: WordApi 1.3]

#### Examples
```TypeScript
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
Gets the text of the body. Use the insertText method to insert text.

```typescript
readonly text: string;
```

- Property value: string

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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
Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Additional types âFootnoteâ, âEndnoteâ, and âNoteItemâ are supported in WordAPIOnline 1.1 and later.

```typescript
readonly type: Word.BodyType | "Unknown" | "MainDoc" | "Section" | "Header" | "Footer" | "TableCell" | "Footnote" | "Endnote" | "NoteItem" | "Shape";
```

- Property value: [Word.BodyType](/en-us/javascript/api/word/word.bodytype) | the literal strings listed above

Remarks  
[API set: WordApi 1.3]

#### Examples
```TypeScript
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

## Method Details

### clear()
Clears the contents of the body object. The user can perform the undo operation on the cleared content.

```typescript
clear(): void;
```

- Returns: void

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### getComments()
Gets comments associated with the body.

```typescript
getComments(): Word.CommentCollection;
```

- Returns: [Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)

Remarks  
[API set: WordApi 1.4]

#### Examples
```TypeScript
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

### getContentControls(options)
Gets the currently supported content controls in the body.

```typescript
getContentControls(options?: Word.ContentControlOptions): Word.ContentControlCollection;
```

- Parameters:
  - options: [Word.ContentControlOptions](/en-us/javascript/api/word/word.contentcontroloptions)  
    Optional. Options that define which content controls are returned.
- Returns: [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks  
[API set: WordApi 1.5]  
Important: If specific types are provided in the options parameter, only content controls of supported types are returned. Be aware that an exception will be thrown on using methods of a generic [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) that aren't relevant for the specific type. With time, additional types of content controls may be supported. Therefore, your add-in should request and handle specific types of content controls.

---

### getHtml()
Gets an HTML representation of the body object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Body.getOoxml()` and convert the returned XML to HTML.

```typescript
getHtml(): OfficeExtension.ClientResult<string>;
```

- Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### getOoxml()
Gets the OOXML (Office Open XML) representation of the body object.

```typescript
getOoxml(): OfficeExtension.ClientResult<string>;
```

- Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### getRange(rangeLocation)
Gets the whole body, or the starting or ending point of the body, as a range.

```typescript
getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | Word.RangeLocation.after | Word.RangeLocation.content | "Whole" | "Start" | "End" | "After" | "Content"): Word.Range;
```

- Parameters:
  - rangeLocation: [whole](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-whole-member) | [start](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-start-member) | [end](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-end-member) | [after](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-after-member) | [content](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-content-member) | "Whole" | "Start" | "End" | "After" | "Content"  
    Optional. The range location must be 'Whole', 'Start', 'End', 'After', or 'Content'.
- Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks  
[API set: WordApi 1.3]

---

### getReviewedText(changeTrackingVersion)
Gets reviewed text based on ChangeTrackingVersion selection.

```typescript
getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion): OfficeExtension.ClientResult<string>;
```

- Parameters:
  - changeTrackingVersion: [Word.ChangeTrackingVersion](/en-us/javascript/api/word/word.changetrackingversion)  
    Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
- Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks  
[API set: WordApi 1.4]

---

### getReviewedText(changeTrackingVersion)
Gets reviewed text based on ChangeTrackingVersion selection.

```typescript
getReviewedText(changeTrackingVersion?: "Original" | "Current"): OfficeExtension.ClientResult<string>;
```

- Parameters:
  - changeTrackingVersion: "Original" | "Current"  
    Optional. The value must be 'Original' or 'Current'. The default is 'Current'.
- Returns: [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

Remarks  
[API set: WordApi 1.4]

---

### getTrackedChanges()
Gets the collection of the TrackedChange objects in the body.

```typescript
getTrackedChanges(): Word.TrackedChangeCollection;
```

- Returns: [Word.TrackedChangeCollection](/en-us/javascript/api/word/word.trackedchangecollection)

Remarks  
[API set: WordApi 1.6]

#### Examples
```TypeScript
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

### insertBreak(breakType, insertLocation)
Inserts a break at the specified location in the main document.

```typescript
insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): void;
```

- Parameters:
  - breakType: [Word.BreakType](/en-us/javascript/api/word/word.breaktype) | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line"  
    Required. The break type to add to the body.
  - insertLocation: [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Start" | "End"  
    Required. The value must be 'Start' or 'End'.
- Returns: void

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### insertContentControl(contentControlType)
Wraps the Body object with a content control.

```typescript
insertContentControl(contentControlType?: Word.ContentControlType.richText | Word.ContentControlType.plainText | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.comboBox | "RichText" | "PlainText" | "CheckBox" | "DropDownList" | "ComboBox"): Word.ContentControl;
```

- Parameters:
  - contentControlType: [richText](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-richtext-member) | [plainText](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-plaintext-member) | [checkBox](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-checkbox-member) | [dropDownList](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-dropdownlist-member) | [comboBox](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-combobox-member) | "RichText" | "PlainText" | "CheckBox" | "DropDownList" | "ComboBox"  
    Optional. Content control type to insert. Must be 'RichText', 'PlainText', 'CheckBox', 'DropDownList', or 'ComboBox'. The default is 'RichText'.
- Returns: [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

Remarks  
[API set: WordApi 1.1]  
Note: The `contentControlType` parameter was introduced in WordApi 1.5. `PlainText` support was added in WordApi 1.5. `CheckBox` support was added in WordApi 1.7. `DropDownList` and `ComboBox` support was added in WordApi 1.9.

#### Examples
```TypeScript
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

### insertFileFromBase64(base64File, insertLocation)
Inserts a document into the body at the specified location.

```typescript
insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
```

- Parameters:
  - base64File: string  
    Required. The Base64-encoded content of a .docx file.
  - insertLocation: [replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End"  
    Required. The value must be 'Replace', 'Start', or 'End'.
- Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks  
[API set: WordApi 1.1]  
Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.

#### Examples
```TypeScript
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

### insertHtml(html, insertLocation)
Inserts HTML at the specified location.

```typescript
insertHtml(html: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
```

- Parameters:
  - html: string  
    Required. The HTML to be inserted in the document.
  - insertLocation: [replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End"  
    Required. The value must be 'Replace', 'Start', or 'End'.
- Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### insertInlinePictureFromBase64(base64EncodedImage, insertLocation)
Inserts a picture into the body at the specified location.

```typescript
insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): Word.InlinePicture;
```

- Parameters:
  - base64EncodedImage: string  
    Required. The Base64-encoded image to be inserted in the body.
  - insertLocation: [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Start" | "End"  
    Required. The value must be 'Start' or 'End'.
- Returns: [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

Remarks  
[API set: WordApi 1.2]

#### Examples
```TypeScript
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

### insertOoxml(ooxml, insertLocation)
Inserts OOXML at the specified location.

```typescript
insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
```

- Parameters:
  - ooxml: string  
    Required. The OOXML to be inserted.
  - insertLocation: [replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End"  
    Required. The value must be 'Replace', 'Start', or 'End'.
- Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### insertParagraph(paragraphText, insertLocation)
Inserts a paragraph at the specified location.

```typescript
insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End"): Word.Paragraph;
```

- Parameters:
  - paragraphText: string  
    Required. The paragraph text to be inserted.
  - insertLocation: [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Start" | "End"  
    Required. The value must be 'Start' or 'End'.
- Returns: [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### insertTable(rowCount, columnCount, insertLocation, values)
Inserts a table with the specified number of rows and columns.

```typescript
insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.start | Word.InsertLocation.end | "Start" | "End", values?: string[][]): Word.Table;
```

- Parameters:
  - rowCount: number  
    Required. The number of rows in the table.
  - columnCount: number  
    Required. The number of columns in the table.
  - insertLocation: [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Start" | "End"  
    Required. The value must be 'Start' or 'End'.
  - values: string[][]  
    Optional 2D array. Cells are filled if the corresponding strings are specified in the array.
- Returns: [Word.Table](/en-us/javascript/api/word/word.table)

Remarks  
[API set: WordApi 1.3]

#### Examples
```TypeScript
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

### insertText(text, insertLocation)
Inserts text into the body at the specified location.

```typescript
insertText(text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End"): Word.Range;
```

- Parameters:
  - text: string  
    Required. Text to be inserted.
  - insertLocation: [replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End"  
    Required. The value must be 'Replace', 'Start', or 'End'.
- Returns: [Word.Range](/en-us/javascript/api/word/word.range)

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.BodyLoadOptions): Word.Body;
```

- Parameters:
  - options: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)  
    Provides options for which properties of the object to load.
- Returns: [Word.Body](/en-us/javascript/api/word/word.body)

---

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Body;
```

- Parameters:
  - propertyNames: string | string[]  
    A comma-delimited string or an array of strings that specify the properties to load.
- Returns: [Word.Body](/en-us/javascript/api/word/word.body)

---

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Body;
```

- Parameters:
  - propertyNamesAndPaths:  
    - select?: string  
      A comma-delimited string that specifies the properties to load.  
    - expand?: string  
      A comma-delimited string that specifies the navigation properties to load.
- Returns: [Word.Body](/en-us/javascript/api/word/word.body)

---

### search(searchText, searchOptions)
Performs a search with the specified SearchOptions on the scope of the body object. The search results are a collection of range objects.

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

- Parameters:
  - searchText: string  
    Required. The search text. Can be a maximum of 255 characters.
  - searchOptions: [Word.SearchOptions](/en-us/javascript/api/word/word.searchoptions) | object with the listed optional boolean properties  
    Optional. Options for the search.
- Returns: [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### select(selectionMode)
Selects the body and navigates the Word UI to it.

```typescript
select(selectionMode?: Word.SelectionMode): void;
```

- Parameters:
  - selectionMode: [Word.SelectionMode](/en-us/javascript/api/word/word.selectionmode)  
    Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
- Returns: void

Remarks  
[API set: WordApi 1.1]

#### Examples
```TypeScript
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

### select(selectionMode)
Selects the body and navigates the Word UI to it.

```typescript
select(selectionMode?: "Select" | "Start" | "End"): void;
```

- Parameters:
  - selectionMode: "Select" | "Start" | "End"  
    Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.
- Returns: void

Remarks  
[API set: WordApi 1.1]

---

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.BodyUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

- Parameters:
  - properties: [Word.Interfaces.BodyUpdateData](/en-us/javascript/api/word/word.interfaces.bodyupdatedata)  
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
    Provides an option to suppress errors if the properties object tries to set any read-only properties.
- Returns: void

---

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Body): void;
```

- Parameters:
  - properties: [Word.Body](/en-us/javascript/api/word/word.body)
- Returns: void

---

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Body` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.BodyData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.BodyData;
```

- Returns: [Word.Interfaces.BodyData](/en-us/javascript/api/word/word.interfaces.bodydata)

---

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Body;
```

- Returns: [Word.Body](/en-us/javascript/api/word/word.body)

---

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.Body;
```

- Returns: [Word.Body](/en-us/javascript/api/word/word.body)

---

## Event Details

### onCommentAdded
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when new comments are added.

```typescript
readonly onCommentAdded: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

- Event type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/manage-comments.yaml

// Registers event handlers.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.track();
  await context.sync();

  eventContexts[0] = body.onCommentAdded.add(onEventHandler);
  eventContexts[1] = body.onCommentChanged.add(onChangedHandler);
  eventContexts[2] = body.onCommentDeleted.add(onEventHandler);
  eventContexts[3] = body.onCommentDeselected.add(onEventHandler);
  eventContexts[4] = body.onCommentSelected.add(onEventHandler);
  await context.sync();

  console.log("Event handlers registered.");
});

...

async function onEventHandler(event: Word.CommentEventArgs) {
  // Handler for all events except onCommentChanged.
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. Event source: ${event.source}. Comment info:`, event.commentDetails);
  });
}
```

---

### onCommentChanged
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment or its reply is changed.

```typescript
readonly onCommentChanged: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

- Event type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/manage-comments.yaml

// Registers event handlers.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.track();
  await context.sync();

  eventContexts[0] = body.onCommentAdded.add(onEventHandler);
  eventContexts[1] = body.onCommentChanged.add(onChangedHandler);
  eventContexts[2] = body.onCommentDeleted.add(onEventHandler);
  eventContexts[3] = body.onCommentDeselected.add(onEventHandler);
  eventContexts[4] = body.onCommentSelected.add(onEventHandler);
  await context.sync();

  console.log("Event handlers registered.");
});

...

async function onChangedHandler(event: Word.CommentEventArgs) {
  await Word.run(async (context) => {
    console.log(
      `${event.type} event detected. ${event.changeType} change made. Event source: ${event.source}. Comment info:`, event.commentDetails
    );
  });
}
```

---

### onCommentDeleted
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when comments are deleted.

```typescript
readonly onCommentDeleted: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

- Event type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/manage-comments.yaml

// Registers event handlers.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.track();
  await context.sync();

  eventContexts[0] = body.onCommentAdded.add(onEventHandler);
  eventContexts[1] = body.onCommentChanged.add(onChangedHandler);
  eventContexts[2] = body.onCommentDeleted.add(onEventHandler);
  eventContexts[3] = body.onCommentDeselected.add(onEventHandler);
  eventContexts[4] = body.onCommentSelected.add(onEventHandler);
  await context.sync();

  console.log("Event handlers registered.");
});

...

async function onEventHandler(event: Word.CommentEventArgs) {
  // Handler for all events except onCommentChanged.
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. Event source: ${event.source}. Comment info:`, event.commentDetails);
  });
}
```

---

### onCommentDeselected
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment is deselected.

```typescript
readonly onCommentDeselected: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

- Event type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/manage-comments.yaml

// Registers event handlers.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.track();
  await context.sync();

  eventContexts[0] = body.onCommentAdded.add(onEventHandler);
  eventContexts[1] = body.onCommentChanged.add(onChangedHandler);
  eventContexts[2] = body.onCommentDeleted.add(onEventHandler);
  eventContexts[3] = body.onCommentDeselected.add(onEventHandler);
  eventContexts[4] = body.onCommentSelected.add(onEventHandler);
  await context.sync();

  console.log("Event handlers registered.");
});

...

async function onEventHandler(event: Word.CommentEventArgs) {
  // Handler for all events except onCommentChanged.
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. Event source: ${event.source}. Comment info:`, event.commentDetails);
  });
}
```

---

### onCommentSelected
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment is selected.

```typescript
readonly onCommentSelected: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

- Event type: [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/manage-comments.yaml

// Registers event handlers.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.track();
  await context.sync();

  eventContexts[0] = body.onCommentAdded.add(onEventHandler);
  eventContexts[1] = body.onCommentChanged.add(onChangedHandler);
  eventContexts[2] = body.onCommentDeleted.add(onEventHandler);
  eventContexts[3] = body.onCommentDeselected.add(onEventHandler);
  eventContexts[4] = body.onCommentSelected.add(onEventHandler);
  await context.sync();

  console.log("Event handlers registered.");
});

...

async function onEventHandler(event: Word.CommentEventArgs) {
  // Handler for all events except onCommentChanged.
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. Event source: ${event.source}. Comment info:`, event.commentDetails);
  });
}
```