# Word.Document

**Package:** `word`

**API Set:** None None

## Description

The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.

## Class Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-change-tracking.yaml

// Gets the current change tracking mode.
await Word.run(async (context) => {
  const document: Word.Document = context.document;
  document.load("changeTrackingMode");
  await context.sync();

  if (document.changeTrackingMode === Word.ChangeTrackingMode.trackMineOnly) {
    console.log("Only my changes are being tracked.");
  } else if (document.changeTrackingMode === Word.ChangeTrackingMode.trackAll) {
    console.log("Everyone's changes are being tracked.");
  } else {
    console.log("No changes are being tracked.");
  }
});
```

## Properties

### activeWindow

**Type:** `Word.Window`

Gets the active window for the document.

#### Examples

**Example**: Retrieve and display the page index, total number of paragraphs, and first paragraph text for each page in the active document window.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets the first paragraph of each page.
  console.log("Getting first paragraph of each page...");

  // Get the active window.
  const activeWindow: Word.Window = context.document.activeWindow;
  activeWindow.load();

  // Get the active pane.
  const activePane: Word.Pane = activeWindow.activePane;
  activePane.load();

  // Get all pages.
  const pages: Word.PageCollection = activePane.pages;
  pages.load();

  await context.sync();

  // Get page index and paragraphs of each page.
  const pagesIndexes = [];
  const pagesNumberOfParagraphs = [];
  const pagesFirstParagraphText = [];
  for (let i = 0; i < pages.items.length; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);

    const paragraphs = page.getRange().paragraphs;
    paragraphs.load('items/length');
    pagesNumberOfParagraphs.push(paragraphs);

    const firstParagraph = paragraphs.getFirst();
    firstParagraph.load('text');
    pagesFirstParagraphText.push(firstParagraph);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Page index: ${pagesIndexes[i].index}`);
    console.log(`Number of paragraphs: ${pagesNumberOfParagraphs[i].items.length}`);
    console.log("First paragraph's text:", pagesFirstParagraphText[i].text);
  }
});
```

---

### attachedTemplate

**Type:** `Word.Template`

Note

#### Examples

**Example**: Get the name and path of the template attached to the current document

```typescript
await Word.run(async (context) => {
    const document = context.document;
    const template = document.attachedTemplate;
    
    template.load("name, path");
    
    await context.sync();
    
    console.log("Template Name: " + template.name);
    console.log("Template Path: " + template.path);
});
```

---

### autoHyphenation

**Type:** `boolean`

Note

#### Examples

**Example**: Enable automatic hyphenation for the entire document to improve text flow and reduce spacing issues in justified paragraphs

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Enable automatic hyphenation
    document.autoHyphenation = true;
    
    await context.sync();
    console.log("Automatic hyphenation has been enabled for the document.");
});
```

---

### autoSaveOn

**Type:** `boolean`

Note

#### Examples

**Example**: Check if AutoSave is currently enabled for the document and display the status to the user

```typescript
await Word.run(async (context) => {
    const document = context.document;
    document.load("autoSaveOn");
    
    await context.sync();
    
    if (document.autoSaveOn) {
        console.log("AutoSave is currently enabled for this document.");
    } else {
        console.log("AutoSave is currently disabled for this document.");
    }
});
```

---

### bibliography

**Type:** `Word.Bibliography`

Note

#### Examples

**Example**: Get the bibliography from the document and access its title property to display it in the console.

```typescript
await Word.run(async (context) => {
    const doc = context.document;
    const bibliography = doc.bibliography;
    
    // Load the bibliography's title property
    bibliography.load("title");
    
    await context.sync();
    
    console.log("Bibliography title: " + bibliography.title);
});
```

---

### body

**Type:** `Word.Body`

Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

#### Examples

**Example**: Add a paragraph with text to the main document body

```typescript
await Word.run(async (context) => {
    const body = context.document.body;
    body.insertParagraph("This is a new paragraph added to the document body.", Word.InsertLocation.end);
    
    await context.sync();
});
```

---

### bookmarks

**Type:** `Word.BookmarkCollection`

Note

#### Examples

**Example**: Get all bookmarks in the document and display their names in the console

```typescript
await Word.run(async (context) => {
    // Get the bookmarks collection from the document
    const bookmarks = context.document.bookmarks;
    
    // Load the bookmark names
    bookmarks.load("items/name");
    
    await context.sync();
    
    // Display each bookmark name
    console.log(`Found ${bookmarks.items.length} bookmark(s):`);
    bookmarks.items.forEach((bookmark) => {
        console.log(`- ${bookmark.name}`);
    });
});
```

---

### changeTrackingMode

**Type:** `Word.ChangeTrackingMode | "Off" | "TrackAll" | "TrackMineOnly"`

Specifies the ChangeTracking mode.

#### Examples

**Example**: Retrieve the current change tracking mode of the document and display whether only the user's changes, everyone's changes, or no changes are being tracked.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-change-tracking.yaml

// Gets the current change tracking mode.
await Word.run(async (context) => {
  const document: Word.Document = context.document;
  document.load("changeTrackingMode");
  await context.sync();

  if (document.changeTrackingMode === Word.ChangeTrackingMode.trackMineOnly) {
    console.log("Only my changes are being tracked.");
  } else if (document.changeTrackingMode === Word.ChangeTrackingMode.trackAll) {
    console.log("Everyone's changes are being tracked.");
  } else {
    console.log("No changes are being tracked.");
  }
});
```

---

### consecutiveHyphensLimit

**Type:** `number`

Note

#### Examples

**Example**: Set the maximum number of consecutive lines that can end with hyphens to 2 to improve document readability

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Set the consecutive hyphens limit to 2
    document.consecutiveHyphensLimit = 2;
    
    await context.sync();
    console.log("Consecutive hyphens limit set to 2");
});
```

---

### contentControls

**Type:** `Word.ContentControlCollection`

Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.

#### Examples

**Example**: Find and highlight all content controls in the document by changing their background color to yellow

```typescript
await Word.run(async (context) => {
    // Get all content controls in the document
    const contentControls = context.document.contentControls;
    
    // Load the content controls
    contentControls.load("items");
    
    await context.sync();
    
    // Set yellow background for each content control
    for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
    
    console.log(`Highlighted ${contentControls.items.length} content controls`);
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the document's request context to synchronize changes and load properties from the Word document.

```typescript
await Word.run(async (context) => {
    const document = context.document;
    const body = document.body;
    
    // Access the request context associated with the document
    const requestContext = document.context;
    
    // Use the context to load properties
    body.load("text");
    
    // Synchronize the context to read loaded properties
    await requestContext.sync();
    
    console.log("Document body text:", body.text);
});
```

---

### customXmlParts

**Type:** `Word.CustomXmlPartCollection`

Gets the custom XML parts in the document.

#### Examples

**Example**: Retrieve and log the count of custom XML parts in the document to verify if any custom XML data exists.

```typescript
await Word.run(async (context) => {
    // Get the custom XML parts collection from the document
    const customXmlParts = context.document.customXmlParts;
    
    // Load the items property to access the collection
    customXmlParts.load("items");
    
    await context.sync();
    
    // Log the count of custom XML parts
    console.log(`Number of custom XML parts: ${customXmlParts.items.length}`);
});
```

---

### documentLibraryVersions

**Type:** `Word.DocumentLibraryVersionCollection`

Note

#### Examples

**Example**: Retrieve and display the number of saved versions available for the current document in the SharePoint document library.

```typescript
await Word.run(async (context) => {
    const document = context.document;
    const versions = document.documentLibraryVersions;
    
    versions.load("items");
    await context.sync();
    
    console.log(`Total versions: ${versions.items.length}`);
    
    // Optionally, display details of each version
    for (let i = 0; i < versions.items.length; i++) {
        const version = versions.items[i];
        version.load("versionIndex, comments, modifiedBy, modifiedDate");
    }
    
    await context.sync();
    
    versions.items.forEach(version => {
        console.log(`Version ${version.versionIndex}: Modified by ${version.modifiedBy} on ${version.modifiedDate}`);
    });
});
```

---

### frames

**Type:** `Word.FrameCollection`

Note

#### Examples

**Example**: Get all frames in the document and log their count to the console

```typescript
await Word.run(async (context) => {
    const frames = context.document.frames;
    frames.load("items");
    
    await context.sync();
    
    console.log(`Total frames in document: ${frames.items.length}`);
});
```

---

### hyperlinks

**Type:** `Word.HyperlinkCollection`

Note

#### Examples

**Example**: Get all hyperlinks in the document and display their text and URLs in the console

```typescript
await Word.run(async (context) => {
    const hyperlinks = context.document.hyperlinks;
    hyperlinks.load("items");
    
    await context.sync();
    
    console.log(`Found ${hyperlinks.items.length} hyperlinks in the document:`);
    
    for (let i = 0; i < hyperlinks.items.length; i++) {
        const hyperlink = hyperlinks.items[i];
        hyperlink.load("textToDisplay, address");
        await context.sync();
        
        console.log(`Text: ${hyperlink.textToDisplay}, URL: ${hyperlink.address}`);
    }
});
```

---

### hyphenateCaps

**Type:** `boolean`

Note

#### Examples

**Example**: Disable hyphenation for capitalized words in the document

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Disable hyphenation for capitalized words
    document.hyphenateCaps = false;
    
    await context.sync();
});
```

---

### indexes

**Type:** `Word.IndexCollection`

Note

#### Examples

**Example**: Access and retrieve all indexes in the document to check if any indexes exist and get the count of indexes.

```typescript
await Word.run(async (context) => {
    const indexes = context.document.indexes;
    indexes.load("items");
    
    await context.sync();
    
    console.log(`Number of indexes in document: ${indexes.items.length}`);
    
    if (indexes.items.length > 0) {
        console.log("Document contains indexes");
    } else {
        console.log("No indexes found in document");
    }
});
```

---

### languageDetected

**Type:** `boolean`

Note

#### Examples

**Example**: Check if the document's language has been automatically detected and display the result to the user

```typescript
await Word.run(async (context) => {
    const document = context.document;
    document.load("languageDetected");
    
    await context.sync();
    
    if (document.languageDetected) {
        console.log("Language has been automatically detected for this document");
    } else {
        console.log("Language has not been automatically detected for this document");
    }
});
```

---

### pageSetup

**Type:** `Word.PageSetup`

Note

#### Examples

**Example**: Configure the document's page margins to 1 inch on all sides and set the page orientation to landscape

```typescript
await Word.run(async (context) => {
    const pageSetup = context.document.pageSetup;
    
    // Set all margins to 1 inch (72 points)
    pageSetup.topMargin = 72;
    pageSetup.bottomMargin = 72;
    pageSetup.leftMargin = 72;
    pageSetup.rightMargin = 72;
    
    // Set orientation to landscape
    pageSetup.orientation = Word.PageOrientation.landscape;
    
    await context.sync();
});
```

---

### properties

**Type:** `Word.DocumentProperties`

Gets the properties of the document.

#### Examples

**Example**: Retrieve and display all built-in document properties from the current Word document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/get-built-in-properties.yaml

await Word.run(async (context) => {
    const builtInProperties: Word.DocumentProperties = context.document.properties;
    builtInProperties.load("*"); // Let's get all!

    await context.sync();
    console.log(JSON.stringify(builtInProperties, null, 4));
});
```

---

### saved

**Type:** `boolean`

Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

#### Examples

**Example**: Check if the document has unsaved changes and display an alert to the user before proceeding with an operation

```typescript
await Word.run(async (context) => {
    const document = context.document;
    document.load("saved");
    
    await context.sync();
    
    if (!document.saved) {
        console.log("Warning: Document has unsaved changes");
        // Prompt user to save before continuing
    } else {
        console.log("Document is saved and up to date");
    }
});
```

---

### sections

**Type:** `Word.SectionCollection`

Gets the collection of section objects in the document.

#### Examples

**Example**: Add a page break between each section in the document by inserting a continuous section break at the end of each section except the last one.

```typescript
await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    
    await context.sync();
    
    // Insert a page break at the end of each section except the last
    for (let i = 0; i < sections.items.length - 1; i++) {
        const section = sections.items[i];
        const body = section.body;
        body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    }
    
    await context.sync();
});
```

---

### settings

**Type:** `Word.SettingCollection`

Gets the add-in's settings in the document.

#### Examples

**Example**: Retrieve and display all custom settings that the add-in has stored in the current Word document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-settings.yaml

// Gets all custom settings this add-in set on this document.
await Word.run(async (context) => {
  const settings: Word.SettingCollection = context.document.settings;
  settings.load("items");
  await context.sync();

  if (settings.items.length == 0) {
    console.log("There are no settings.");
  } else {
    console.log("All settings:");
    for (let i = 0; i < settings.items.length; i++) {
      console.log(settings.items[i]);
    }
  }
});
```

---

### windows

**Type:** `Word.WindowCollection`

Gets the collection of `Word.Window` objects for the document.

#### Examples

**Example**: Get the number of windows currently open for the document and display it in the console.

```typescript
await Word.run(async (context) => {
    // Get the windows collection for the document
    const windows = context.document.windows;
    
    // Load the count property
    windows.load("count");
    
    await context.sync();
    
    // Display the number of windows
    console.log(`Number of windows open for this document: ${windows.count}`);
});
```

---

## Methods

### addStyle

**Kind:** `create`

Adds a style into the document by name and type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `name`: `None` (required)
  - `type`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `name`: `None` (required)
  - `type`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Add a new custom style to the Word document with a user-specified name and type, after verifying that no style with that name already exists.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Adds a new style.
await Word.run(async (context) => {
  const newStyleName = (document.getElementById("new-style-name") as HTMLInputElement).value;
  if (newStyleName == "") {
    console.warn("Enter a style name to add.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(newStyleName);
  style.load();
  await context.sync();

  if (!style.isNullObject) {
    console.warn(
      `There's an existing style with the same name '${newStyleName}'! Please provide another style name.`
    );
    return;
  }

  const newStyleType = ((document.getElementById("new-style-type") as HTMLSelectElement).value as unknown) as Word.StyleType;
  context.document.addStyle(newStyleName, newStyleType);
  await context.sync();

  console.log(newStyleName + " has been added to the style list.");
});
```

---

### close

Closes the current document.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `closeBehavior`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `closeBehavior`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Close the current Word document with default behavior based on its current state.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/save-close.yaml

// Closes the document with default behavior
// for current state of the document.
await Word.run(async (context) => {
  context.document.close();
});
```

---

### compare

Displays revision marks that indicate where the specified document differs from another document.

#### Signature

**Parameters:**
- `filePath`: `None` (required)
- `documentCompareOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Compare the current Word document with an external document specified by file path and display the differences in the current document without detecting format changes.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/compare-documents.yaml

// Compares the current document with a specified external document.
await Word.run(async (context) => {
  // Absolute path of an online or local document.
  const filePath = (document.getElementById("filePath") as HTMLInputElement).value;
  // Options that configure the compare operation.
  const options: Word.DocumentCompareOptions = {
    compareTarget: Word.CompareTarget.compareTargetCurrent,
    detectFormatChanges: false
    // Other options you choose...
    };
  context.document.compare(filePath, options);

  await context.sync();

  console.log("Differences shown in the current document.");
});
```

---

### compareFromBase64

Displays revision marks that indicate where the specified document differs from another document.

#### Signature

**Parameters:**
- `base64File`: `None` (required)
- `documentCompareOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Compare the current document with a base document stored as a base64 string and display revision marks showing the differences

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Base64 string of the document to compare against
    const base64Document = "UEsDBBQABgAIAAAAIQDd..."; // Truncated for brevity
    
    // Configure comparison options
    const compareOptions: Word.DocumentCompareOptions = {
        compareTarget: Word.CompareTarget.current,
        detectFormatChanges: true,
        ignoreWhitespace: false,
        removeDateAndTime: false,
        removePersonalInformation: false
    };
    
    // Compare the current document with the base64 document
    document.compareFromBase64(base64Document, compareOptions);
    
    await context.sync();
    
    console.log("Document comparison complete. Revision marks are now visible.");
});
```

---

### deleteBookmark

**Kind:** `delete`

Deletes a bookmark, if it exists, from the document.

#### Signature

**Parameters:**
- `name`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Delete a bookmark named "Section1Start" from the document if it exists

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Delete the bookmark named "Section1Start"
    document.deleteBookmark("Section1Start");
    
    await context.sync();
    console.log("Bookmark 'Section1Start' has been deleted (if it existed).");
});
```

---

### detectLanguage

Note

#### Signature

**Returns:** `None`

#### Examples

**Example**: Detect the language of all text content in the document to identify what language the user has written in

```typescript
await Word.run(async (context) => {
    // Get the document
    const document = context.document;
    
    // Detect the language of the document content
    document.detectLanguage();
    
    // Sync to execute the detection
    await context.sync();
    
    console.log("Language detection completed for the document");
});
```

---

### getAnnotationById

**Kind:** `read`

Gets the annotation by ID. Throws an `ItemNotFound` error if annotation isn't found.

#### Signature

**Parameters:**
- `id`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Retrieve and display the text of a specific annotation by its ID "annotation123"

```typescript
await Word.run(async (context) => {
    const annotation = context.document.getAnnotationById("annotation123");
    annotation.load("critiqueAnnotation");
    
    await context.sync();
    
    console.log("Annotation text:", annotation.critiqueAnnotation);
});
```

---

### getBookmarkRange

**Kind:** `read`

Gets a bookmark's range. Throws an `ItemNotFound` error if the bookmark doesn't exist.

#### Signature

**Parameters:**
- `name`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Highlight and modify the text content within a bookmark named "CompanyName" in the document.

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Get the range of the bookmark named "CompanyName"
    const bookmarkRange = document.getBookmarkRange("CompanyName");
    
    // Load the range properties
    bookmarkRange.load("text");
    await context.sync();
    
    // Modify the bookmark content
    bookmarkRange.insertText("Contoso Corporation", Word.InsertLocation.replace);
    bookmarkRange.font.highlightColor = "yellow";
    bookmarkRange.font.bold = true;
    
    await context.sync();
});
```

---

### getBookmarkRangeOrNullObject

**Kind:** `read`

Gets a bookmark's range. If the bookmark doesn't exist, then this method will return an object with its `isNullObject` property set to `true`. For further information, see *OrNullObject methods and properties.

#### Signature

**Parameters:**
- `name`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Check if a bookmark named "Introduction" exists in the document and highlight its text in yellow if found, otherwise log that it doesn't exist.

```typescript
await Word.run(async (context) => {
    const bookmarkRange = context.document.getBookmarkRangeOrNullObject("Introduction");
    bookmarkRange.load("isNullObject");
    
    await context.sync();
    
    if (bookmarkRange.isNullObject) {
        console.log("Bookmark 'Introduction' does not exist.");
    } else {
        bookmarkRange.font.highlightColor = "yellow";
        console.log("Bookmark 'Introduction' found and highlighted.");
    }
    
    await context.sync();
});
```

---

### getContentControls

**Kind:** `read`

Gets the currently supported content controls in the document.

#### Signature

**Parameters:**
- `options`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Toggle the checked state of all checkbox content controls in the document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml

// Toggles the isChecked property on all checkbox content controls.
await Word.run(async (context) => {
  let contentControls = context.document.getContentControls({
    types: [Word.ContentControlType.checkBox]
  });
  contentControls.load("items");

  await context.sync();

  const length = contentControls.items.length;
  console.log(`Number of checkbox content controls: ${length}`);

  if (length <= 0) {
    return;
  }

  const checkboxContentControls = [];
  for (let i = 0; i < length; i++) {
    let contentControl = contentControls.items[i];
    contentControl.load("id,checkboxContentControl/isChecked");
    checkboxContentControls.push(contentControl);
  }

  await context.sync();

  console.log("isChecked state before:");
  const updatedCheckboxContentControls = [];
  for (let i = 0; i < checkboxContentControls.length; i++) {
    const currentCheckboxContentControl = checkboxContentControls[i];
    const isCheckedBefore = currentCheckboxContentControl.checkboxContentControl.isChecked;
    console.log(`id: ${currentCheckboxContentControl.id} ... isChecked: ${isCheckedBefore}`);

    currentCheckboxContentControl.checkboxContentControl.isChecked = !isCheckedBefore;
    currentCheckboxContentControl.load("id,checkboxContentControl/isChecked");
    updatedCheckboxContentControls.push(currentCheckboxContentControl);
  }

  await context.sync();

  console.log("isChecked state after:");
  for (let i = 0; i < updatedCheckboxContentControls.length; i++) {
    const currentCheckboxContentControl = updatedCheckboxContentControls[i];
    console.log(
      `id: ${currentCheckboxContentControl.id} ... isChecked: ${currentCheckboxContentControl.checkboxContentControl.isChecked}`
    );
  }
});
```

---

### getEndnoteBody

**Kind:** `read`

Gets the document's endnotes in a single body.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Add a new paragraph with text to the document's endnotes section

```typescript
await Word.run(async (context) => {
    // Get the endnote body
    const endnoteBody = context.document.getEndnoteBody();
    
    // Add a paragraph to the endnote body
    const paragraph = endnoteBody.insertParagraph(
        "This text is added to the endnotes section.",
        Word.InsertLocation.end
    );
    
    // Sync to apply changes
    await context.sync();
    
    console.log("Paragraph added to endnotes successfully");
});
```

---

### getFootnoteBody

**Kind:** `read`

Gets the document's footnotes in a single body.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Add a new paragraph with text to the document's footnotes body.

```typescript
await Word.run(async (context) => {
    // Get the footnotes body
    const footnoteBody = context.document.getFootnoteBody();
    
    // Add a paragraph to the footnotes body
    const paragraph = footnoteBody.insertParagraph(
        "This is additional content in the footnotes section.",
        Word.InsertLocation.end
    );
    
    // Load and sync to verify
    footnoteBody.load("text");
    await context.sync();
    
    console.log("Footnote body content:", footnoteBody.text);
});
```

---

### getParagraphByUniqueLocalId

**Kind:** `read`

Gets the paragraph by its unique local ID. Throws an `ItemNotFound` error if the collection is empty.

#### Signature

**Parameters:**
- `id`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Retrieve and display a paragraph from the document using its unique local identifier obtained from user input.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/onadded-event.yaml

await Word.run(async (context) => {
  const paragraphId = (document.getElementById("paragraph-id") as HTMLInputElement).value;
  const paragraph: Word.Paragraph = context.document.getParagraphByUniqueLocalId(paragraphId);
  paragraph.load();
  await paragraph.context.sync();

  console.log(paragraph);
});
```

---

### getSelection

**Kind:** `read`

Gets the current selection of the document. Multiple selections aren't supported.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Insert explanatory text about the insert text method at the end of the current selection in the document.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    const textSample = 'This is an example of the insert text method. This is a method ' + 
        'which allows users to insert text into a selection. It can insert text into a ' +
        'relative location or it can overwrite the current selection. Since the ' +
        'getSelection method returns a range object, look up the range object documentation ' +
        'for everything you can do with a selection.';
    
    // Create a range proxy object for the current selection.
    const range = context.document.getSelection();
    
    // Queue a command to insert text at the end of the selection.
    range.insertText(textSample, Word.InsertLocation.end);
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Inserted the text at the end of the selection.');
});
```

---

### getStyles

**Kind:** `read`

Gets a StyleCollection object that represents the whole style set of the document.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Retrieve and display the total count of styles available in the Word document.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Gets the number of available styles stored with the document.
await Word.run(async (context) => {
  const styles: Word.StyleCollection = context.document.getStyles();
  const count = styles.getCount();
  await context.sync();

  console.log(`Number of styles: ${count.value}`);
});
```

---

### importStylesFromJson

Import styles from a JSON-formatted string.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `stylesJson`: `None` (required)
  - `importedStylesConflictBehavior`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `stylesJson`: `None` (required)
  - `importedStylesConflictBehavior`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Import custom character, paragraph, and table styles into a Word document from a JSON string containing style definitions with formatting properties.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/40-tables/manage-custom-style.yaml

// Imports styles from JSON.
await Word.run(async (context) => {
  const str =
    '{"styles":[{"baseStyle":"Default Paragraph Font","builtIn":false,"inUse":true,"linked":false,"nameLocal":"NewCharStyle","priority":2,"quickStyle":true,"type":"Character","unhideWhenUsed":false,"visibility":false,"paragraphFormat":null,"font":{"name":"DengXian Light","size":16.0,"bold":true,"italic":false,"color":"#F1A983","underline":"None","subscript":false,"superscript":true,"strikeThrough":true,"doubleStrikeThrough":false,"highlightColor":null,"hidden":false},"shading":{"backgroundPatternColor":"#FF0000"}},{"baseStyle":"Normal","builtIn":false,"inUse":true,"linked":false,"nextParagraphStyle":"NewParaStyle","nameLocal":"NewParaStyle","priority":1,"quickStyle":true,"type":"Paragraph","unhideWhenUsed":false,"visibility":false,"paragraphFormat":{"alignment":"Centered","firstLineIndent":0.0,"keepTogether":false,"keepWithNext":false,"leftIndent":72.0,"lineSpacing":18.0,"lineUnitAfter":0.0,"lineUnitBefore":0.0,"mirrorIndents":false,"outlineLevel":"OutlineLevelBodyText","rightIndent":72.0,"spaceAfter":30.0,"spaceBefore":30.0,"widowControl":true},"font":{"name":"DengXian","size":14.0,"bold":true,"italic":true,"color":"#8DD873","underline":"Single","subscript":false,"superscript":false,"strikeThrough":false,"doubleStrikeThrough":true,"highlightColor":null,"hidden":false},"shading":{"backgroundPatternColor":"#00FF00"}},{"baseStyle":"Table Normal","builtIn":false,"inUse":true,"linked":false,"nextParagraphStyle":"NewTableStyle","nameLocal":"NewTableStyle","priority":100,"type":"Table","unhideWhenUsed":false,"visibility":false,"paragraphFormat":{"alignment":"Left","firstLineIndent":0.0,"keepTogether":false,"keepWithNext":false,"leftIndent":0.0,"lineSpacing":12.0,"lineUnitAfter":0.0,"lineUnitBefore":0.0,"mirrorIndents":false,"outlineLevel":"OutlineLevelBodyText","rightIndent":0.0,"spaceAfter":0.0,"spaceBefore":0.0,"widowControl":true},"font":{"name":"DengXian","size":20.0,"bold":false,"italic":true,"color":"#D86DCB","underline":"None","subscript":false,"superscript":false,"strikeThrough":false,"doubleStrikeThrough":false,"highlightColor":null,"hidden":false},"tableStyle":{"allowBreakAcrossPage":true,"alignment":"Left","bottomCellMargin":0.0,"leftCellMargin":0.08,"rightCellMargin":0.08,"topCellMargin":0.0,"cellSpacing":0.0},"shading":{"backgroundPatternColor":"#60CAF3"}}]}';
  const styles = context.document.importStylesFromJson(str);
  await context.sync();
  console.log("Styles imported from JSON:", styles);
});
```

---

### insertFileFromBase64

**Kind:** `create`

Inserts a document into the target document at a specific location with additional properties. Headers, footers, watermarks, and other section properties are copied by default.

#### Signature

**Parameters:**
- `base64File`: `None` (required)
- `insertLocation`: `None` (required)
- `insertFileOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Replace the current document content with content from a Base64-encoded external document file while importing its theme, styles, paragraph spacing, page color, change tracking mode, custom properties, custom XML parts, and different odd/even page settings.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/insert-external-document.yaml

// Inserts content (applying selected settings) from another document passed in as a Base64-encoded string.
await Word.run(async (context) => {
  // Use the Base64-encoded string representation of the selected .docx file.
  context.document.insertFileFromBase64(externalDocument, "Replace", {
    importTheme: true,
    importStyles: true,
    importParagraphSpacing: true,
    importPageColor: true,
    importChangeTrackingMode: true,
    importCustomProperties: true,
    importCustomXmlParts: true,
    importDifferentOddEvenPages: true
  });
  await context.sync();
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

**Example**: Retrieve and display the ID, text content, and tag properties of all content controls in the document, or indicate if no content controls exist.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the document.
    const thisDocument = context.document;
    
    // Queue a command to load content control properties.
    thisDocument.load('contentControls/id, contentControls/text, contentControls/tag');
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
    if (thisDocument.contentControls.items.length !== 0) {
        for (let i = 0; i < thisDocument.contentControls.items.length; i++) {
            console.log(thisDocument.contentControls.items[i].id);
            console.log(thisDocument.contentControls.items[i].text);
            console.log(thisDocument.contentControls.items[i].tag);
        }
    } else {
        console.log('No content controls in this document.');
    }
});
```

---

### manualHyphenation

Note

#### Signature

**Returns:** `None`

#### Examples

**Example**: Enable manual hyphenation for the document to allow users to manually insert hyphens at line breaks

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Enable manual hyphenation for the document
    document.manualHyphenation();
    
    await context.sync();
    console.log("Manual hyphenation has been enabled for the document.");
});
```

---

### save

Saves the document.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `saveBehavior`: `None` (required)
  - `fileName`: `None` (required)

  **Returns:** `None`

**Overload 2:**

  **Parameters:**
  - `saveBehavior`: `None` (required)
  - `fileName`: `None` (required)

  **Returns:** `None`

#### Examples

**Example**: Check if the document has unsaved changes and save it if necessary, logging the appropriate status message.

```typescript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {
    
    // Create a proxy object for the document.
    const thisDocument = context.document;

    // Queue a command to load the document save state (on the saved property).
    thisDocument.load('saved');    
    
    // Synchronize the document state by executing the queued commands, 
    // and return a promise to indicate task completion.
    await context.sync();
        
    if (thisDocument.saved === false) {
        // Queue a command to save this document.
        thisDocument.save();
        
        // Synchronize the document state by executing the queued commands, 
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Saved the document');
    } else {
        console.log('The document has not changed since the last save.');
    }
});
```

**Example**: Save the current Word document with its default save behavior.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/save-close.yaml

// Saves the document with default behavior
// for current state of the document.
await Word.run(async (context) => {
  context.document.save();
  await context.sync();
});
```

---

### search

Performs a search with the specified search options on the scope of the whole document. The search results are a collection of range objects.

#### Signature

**Parameters:**
- `searchText`: `None` (required)
- `searchOptions`: `None` (required)

**Returns:** `None`

#### Examples

**Example**: Find all occurrences of the word "TODO" in the document and highlight them in yellow

```typescript
await Word.run(async (context) => {
    // Search for all occurrences of "TODO" in the document
    const searchResults = context.document.search("TODO", {
        matchCase: false,
        matchWholeWord: true
    });
    
    // Load the search results
    searchResults.load("font");
    await context.sync();
    
    // Highlight each occurrence in yellow
    for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.highlightColor = "yellow";
    }
    
    await context.sync();
    
    console.log(`Found and highlighted ${searchResults.items.length} occurrences of "TODO"`);
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

**Example**: Configure multiple document properties at once, including setting the document title and subject metadata

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Set multiple document properties at once
    document.set({
        properties: {
            title: "Q4 Sales Report",
            subject: "Quarterly Financial Analysis",
            author: "Sales Team"
        }
    });
    
    await context.sync();
    console.log("Document properties updated successfully");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Document` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentData`) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Serialize a Word document's properties to JSON format for logging or debugging purposes

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Load properties you want to serialize
    document.load("body/text,properties/title,properties/author");
    
    await context.sync();
    
    // Convert the document object to a plain JavaScript object
    const documentData = document.toJSON();
    
    // Now you can use JSON.stringify or log the plain object
    console.log(JSON.stringify(documentData, null, 2));
    
    // The documentData object contains only the loaded properties
    // as a plain JavaScript object, not the API proxy object
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Get a document reference, track it to prevent "InvalidObjectPath" errors across multiple sync calls, and read its properties multiple times in separate sync operations.

```typescript
await Word.run(async (context) => {
    const document = context.document;
    
    // Track the document object to use it across multiple sync calls
    document.track();
    
    // Load and sync properties first time
    document.load("properties/title");
    await context.sync();
    
    console.log("Document title:", document.properties.title);
    
    // Use the document again after sync (tracking prevents InvalidObjectPath error)
    document.body.load("text");
    await context.sync();
    
    console.log("Document body length:", document.body.text.length);
    
    // Clean up tracked object when done
    document.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

#### Signature

**Returns:** `None`

#### Examples

**Example**: Load a document object to read its properties, then untrack it to free memory after use

```typescript
await Word.run(async (context) => {
    const document = context.document;
    document.track();
    document.load("properties");
    
    await context.sync();
    
    // Use the document properties
    console.log("Document loaded and used");
    
    // Release memory after we're done using the tracked object
    document.untrack();
    await context.sync();
});
```

---

## Source

- https://docs.microsoft.com/en-us/javascript/api/word
