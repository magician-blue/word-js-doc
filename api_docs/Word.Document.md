# Word.Document class

Package: [word](/en-us/javascript/api/word)

The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.

Extends
- [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[ API set: WordApi 1.1 ]

#### Examples
```TypeScript
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
- activeWindow — Gets the active window for the document.
- attachedTemplate — Specifies a Template object that represents the template attached to the document.
- autoHyphenation — Specifies if automatic hyphenation is turned on for the document.
- autoSaveOn — Specifies if the edits in the document are automatically saved.
- bibliography — Returns a Bibliography object that represents the bibliography references contained within the document.
- body — Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
- bookmarks — Returns a BookmarkCollection object that represents all the bookmarks in the document.
- changeTrackingMode — Specifies the ChangeTracking mode.
- consecutiveHyphensLimit — Specifies the maximum number of consecutive lines that can end with hyphens.
- contentControls — Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- customXmlParts — Gets the custom XML parts in the document.
- documentLibraryVersions — Returns a DocumentLibraryVersionCollection object that represents the collection of versions of a shared document that has versioning enabled and that's stored in a document library on a server.
- frames — Returns a FrameCollection object that represents all the frames in the document.
- hyperlinks — Returns a HyperlinkCollection object that represents all the hyperlinks in the document.
- hyphenateCaps — Specifies whether words in all capital letters can be hyphenated.
- indexes — Returns an IndexCollection object that represents all the indexes in the document.
- languageDetected — Specifies whether Microsoft Word has detected the language of the document text.
- pageSetup — Returns a PageSetup object that's associated with the document.
- properties — Gets the properties of the document.
- saved — Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.
- sections — Gets the collection of section objects in the document.
- settings — Gets the add-in's settings in the document.
- windows — Gets the collection of Word.Window objects for the document.

## Methods
- addStyle(name, type) — Adds a style into the document by name and type.
- addStyle(name, type) — Adds a style into the document by name and type.
- close(closeBehavior) — Closes the current document. Note: This API isn't supported in Word on the web.
- close(closeBehavior) — Closes the current document. Note: This API isn't supported in Word on the web.
- compare(filePath, documentCompareOptions) — Displays revision marks that indicate where the specified document differs from another document.
- compareFromBase64(base64File, documentCompareOptions) — Displays revision marks that indicate where the specified document differs from another document.
- deleteBookmark(name) — Deletes a bookmark, if it exists, from the document.
- detectLanguage() — Analyzes the document text to determine the language.
- getAnnotationById(id) — Gets the annotation by ID. Throws an ItemNotFound error if annotation isn't found.
- getBookmarkRange(name) — Gets a bookmark's range. Throws an ItemNotFound error if the bookmark doesn't exist.
- getBookmarkRangeOrNullObject(name) — Gets a bookmark's range. If the bookmark doesn't exist, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- getContentControls(options) — Gets the currently supported content controls in the document.
- getEndnoteBody() — Gets the document's endnotes in a single body.
- getFootnoteBody() — Gets the document's footnotes in a single body.
- getParagraphByUniqueLocalId(id) — Gets the paragraph by its unique local ID. Throws an ItemNotFound error if the collection is empty.
- getSelection() — Gets the current selection of the document. Multiple selections aren't supported.
- getStyles() — Gets a StyleCollection object that represents the whole style set of the document.
- importStylesFromJson(stylesJson, importedStylesConflictBehavior) — Import styles from a JSON-formatted string.
- importStylesFromJson(stylesJson, importedStylesConflictBehavior) — Import styles from a JSON-formatted string.
- insertFileFromBase64(base64File, insertLocation, insertFileOptions) — Inserts a document into the target document at a specific location with additional properties. Headers, footers, watermarks, and other section properties are copied by default.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- manualHyphenation() — Initiates manual hyphenation of a document, one line at a time.
- save(saveBehavior, fileName) — Saves the document.
- save(saveBehavior, fileName) — Saves the document.
- search(searchText, searchOptions) — Performs a search with the specified search options on the scope of the whole document. The search results are a collection of range objects.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

## Events
- onAnnotationClicked — Occurs when the user clicks an annotation (or selects it using **Alt+Down**).
- onAnnotationHovered — Occurs when the user hovers the cursor over an annotation.
- onAnnotationInserted — Occurs when the user adds one or more annotations.
- onAnnotationPopupAction — Occurs when the user performs an action in an annotation pop-up menu.
- onAnnotationRemoved — Occurs when the user deletes one or more annotations.
- onContentControlAdded — Occurs when a content control is added. Run context.sync() in the handler to get the new content control's properties.
- onParagraphAdded — Occurs when the user adds new paragraphs.
- onParagraphChanged — Occurs when the user changes paragraphs.
- onParagraphDeleted — Occurs when the user deletes paragraphs.

## Property Details

### activeWindow
Gets the active window for the document.

```typescript
readonly activeWindow: Word.Window;
```

Property Value
- [Word.Window](/en-us/javascript/api/word/word.window)

Remarks
- [ API set: WordApiDesktop 1.2 ]

#### Examples
```TypeScript
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

### attachedTemplate
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `Template` object that represents the template attached to the document.

```typescript
attachedTemplate: Word.Template;
```

Property Value
- [Word.Template](/en-us/javascript/api/word/word.template)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### autoHyphenation
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if automatic hyphenation is turned on for the document.

```typescript
autoHyphenation: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### autoSaveOn
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the edits in the document are automatically saved.

```typescript
autoSaveOn: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### bibliography
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Bibliography` object that represents the bibliography references contained within the document.

```typescript
readonly bibliography: Word.Bibliography;
```

Property Value
- [Word.Bibliography](/en-us/javascript/api/word/word.bibliography)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### body
Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

```typescript
readonly body: Word.Body;
```

Property Value
- [Word.Body](/en-us/javascript/api/word/word.body)

Remarks
- [ API set: WordApi 1.1 ]

### bookmarks
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BookmarkCollection` object that represents all the bookmarks in the document.

```typescript
readonly bookmarks: Word.BookmarkCollection;
```

Property Value
- [Word.BookmarkCollection](/en-us/javascript/api/word/word.bookmarkcollection)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### changeTrackingMode
Specifies the ChangeTracking mode.

```typescript
changeTrackingMode: Word.ChangeTrackingMode | "Off" | "TrackAll" | "TrackMineOnly";
```

Property Value
- [Word.ChangeTrackingMode](/en-us/javascript/api/word/word.changetrackingmode) | "Off" | "TrackAll" | "TrackMineOnly"

Remarks
- [ API set: WordApi 1.4 ]

#### Examples
```TypeScript
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

### consecutiveHyphensLimit
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the maximum number of consecutive lines that can end with hyphens.

```typescript
consecutiveHyphensLimit: number;
```

Property Value
- number

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### contentControls
Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.

```typescript
readonly contentControls: Word.ContentControlCollection;
```

Property Value
- [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks
- [ API set: WordApi 1.1 ]

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### customXmlParts
Gets the custom XML parts in the document.

```typescript
readonly customXmlParts: Word.CustomXmlPartCollection;
```

Property Value
- [Word.CustomXmlPartCollection](/en-us/javascript/api/word/word.customxmlpartcollection)

Remarks
- [ API set: WordApi 1.4 ]

### documentLibraryVersions
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `DocumentLibraryVersionCollection` object that represents the collection of versions of a shared document that has versioning enabled and that's stored in a document library on a server.

```typescript
readonly documentLibraryVersions: Word.DocumentLibraryVersionCollection;
```

Property Value
- [Word.DocumentLibraryVersionCollection](/en-us/javascript/api/word/word.documentlibraryversioncollection)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### frames
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `FrameCollection` object that represents all the frames in the document.

```typescript
readonly frames: Word.FrameCollection;
```

Property Value
- [Word.FrameCollection](/en-us/javascript/api/word/word.framecollection)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### hyperlinks
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `HyperlinkCollection` object that represents all the hyperlinks in the document.

```typescript
readonly hyperlinks: Word.HyperlinkCollection;
```

Property Value
- [Word.HyperlinkCollection](/en-us/javascript/api/word/word.hyperlinkcollection)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### hyphenateCaps
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether words in all capital letters can be hyphenated.

```typescript
hyphenateCaps: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### indexes
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an `IndexCollection` object that represents all the indexes in the document.

```typescript
readonly indexes: Word.IndexCollection;
```

Property Value
- [Word.IndexCollection](/en-us/javascript/api/word/word.indexcollection)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### languageDetected
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word has detected the language of the document text.

```typescript
languageDetected: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### pageSetup
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `PageSetup` object that's associated with the document.

```typescript
readonly pageSetup: Word.PageSetup;
```

Property Value
- [Word.PageSetup](/en-us/javascript/api/word/word.pagesetup)

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### properties
Gets the properties of the document.

```typescript
readonly properties: Word.DocumentProperties;
```

Property Value
- [Word.DocumentProperties](/en-us/javascript/api/word/word.documentproperties)

Remarks
- [ API set: WordApi 1.3 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/30-properties/get-built-in-properties.yaml

await Word.run(async (context) => {
    const builtInProperties: Word.DocumentProperties = context.document.properties;
    builtInProperties.load("*"); // Let's get all!

    await context.sync();
    console.log(JSON.stringify(builtInProperties, null, 4));
});
```

### saved
Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

```typescript
readonly saved: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi 1.1 ]

### sections
Gets the collection of section objects in the document.

```typescript
readonly sections: Word.SectionCollection;
```

Property Value
- [Word.SectionCollection](/en-us/javascript/api/word/word.sectioncollection)

Remarks
- [ API set: WordApi 1.1 ]

### settings
Gets the add-in's settings in the document.

```typescript
readonly settings: Word.SettingCollection;
```

Property Value
- [Word.SettingCollection](/en-us/javascript/api/word/word.settingcollection)

Remarks
- [ API set: WordApi 1.4 ]

#### Examples
```TypeScript
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

### windows
Gets the collection of `Word.Window` objects for the document.

```typescript
readonly windows: Word.WindowCollection;
```

Property Value
- [Word.WindowCollection](/en-us/javascript/api/word/word.windowcollection)

Remarks
- [ API set: WordApiDesktop 1.2 ]

## Method Details

### addStyle(name, type)
Adds a style into the document by name and type.

```typescript
addStyle(name: string, type: Word.StyleType): Word.Style;
```

Parameters
- name: string  
  Required. A string representing the style name.
- type: [Word.StyleType](/en-us/javascript/api/word/word.styletype)  
  Required. The style type, including character, list, paragraph, or table.

Returns
- [Word.Style](/en-us/javascript/api/word/word.style)

Remarks
- [ API set: WordApi 1.5 ]

#### Examples
```TypeScript
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

### addStyle(name, type)
Adds a style into the document by name and type.

```typescript
addStyle(name: string, type: "Character" | "List" | "Paragraph" | "Table"): Word.Style;
```

Parameters
- name: string  
  Required. A string representing the style name.
- type: "Character" | "List" | "Paragraph" | "Table"  
  Required. The style type, including character, list, paragraph, or table.

Returns
- [Word.Style](/en-us/javascript/api/word/word.style)

Remarks
- [ API set: WordApi 1.5 ]

### close(closeBehavior)
Closes the current document.  
Note: This API isn't supported in Word on the web.

```typescript
close(closeBehavior?: Word.CloseBehavior): void;
```

Parameters
- closeBehavior: [Word.CloseBehavior](/en-us/javascript/api/word/word.closebehavior)  
  Optional. The close behavior must be 'Save' or 'SkipSave'. Default value is 'Save'.

Returns
- void

Remarks
- [ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/save-close.yaml

// Closes the document with default behavior
// for current state of the document.
await Word.run(async (context) => {
  context.document.close();
});
```

### close(closeBehavior)
Closes the current document.  
Note: This API isn't supported in Word on the web.

```typescript
close(closeBehavior?: "Save" | "SkipSave"): void;
```

Parameters
- closeBehavior: "Save" | "SkipSave"  
  Optional. The close behavior must be 'Save' or 'SkipSave'. Default value is 'Save'.

Returns
- void

Remarks
- [ API set: WordApi 1.5 ]

### compare(filePath, documentCompareOptions)
Displays revision marks that indicate where the specified document differs from another document.

```typescript
compare(filePath: string, documentCompareOptions?: Word.DocumentCompareOptions): void;
```

Parameters
- filePath: string  
  Required. The path of the document with which the specified document is compared.
- documentCompareOptions: [Word.DocumentCompareOptions](/en-us/javascript/api/word/word.documentcompareoptions)  
  Optional. The additional options that specifies the behavior of comparing document.

Returns
- void

Remarks
- [ API set: WordApiDesktop 1.1 ]

#### Examples
```TypeScript
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

### compareFromBase64(base64File, documentCompareOptions)
Displays revision marks that indicate where the specified document differs from another document.

```typescript
compareFromBase64(base64File: string, documentCompareOptions?: Word.DocumentCompareOptions): void;
```

Parameters
- base64File: string  
  Required. The Base64-encoded content of the document with which the specified document is compared.
- documentCompareOptions: [Word.DocumentCompareOptions](/en-us/javascript/api/word/word.documentcompareoptions)  
  Optional. The additional options that specify the behavior for comparing the documents. Note that the `compareTarget` option isn't allowed to be `CompareTargetSelected` in this API.

Returns
- void

Remarks
- [ API set: WordApiDesktop 1.2 ]

### deleteBookmark(name)
Deletes a bookmark, if it exists, from the document.

```typescript
deleteBookmark(name: string): void;
```

Parameters
- name: string  
  Required. The case-insensitive bookmark name.

Returns
- void

Remarks
- [ API set: WordApi 1.4 ]

### detectLanguage()
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Analyzes the document text to determine the language.

```typescript
detectLanguage(): void;
```

Returns
- void

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### getAnnotationById(id)
Gets the annotation by ID. Throws an `ItemNotFound` error if annotation isn't found.

```typescript
getAnnotationById(id: string): Word.Annotation;
```

Parameters
- id: string  
  The ID of the annotation to get.

Returns
- [Word.Annotation](/en-us/javascript/api/word/word.annotation)

Remarks
- [ API set: WordApi 1.7 ]

### getBookmarkRange(name)
Gets a bookmark's range. Throws an `ItemNotFound` error if the bookmark doesn't exist.

```typescript
getBookmarkRange(name: string): Word.Range;
```

Parameters
- name: string  
  Required. The case-insensitive bookmark name.

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
- [ API set: WordApi 1.4 ]

### getBookmarkRangeOrNullObject(name)
Gets a bookmark's range. If the bookmark doesn't exist, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getBookmarkRangeOrNullObject(name: string): Word.Range;
```

Parameters
- name: string  
  Required. The case-insensitive bookmark name.

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
- [ API set: WordApi 1.4 ]

### getContentControls(options)
Gets the currently supported content controls in the document.

```typescript
getContentControls(options?: Word.ContentControlOptions): Word.ContentControlCollection;
```

Parameters
- options: [Word.ContentControlOptions](/en-us/javascript/api/word/word.contentcontroloptions)  
  Optional. Options that define which content controls are returned.

Returns
- [Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

Remarks
- [ API set: WordApi 1.5 ]

Important: If specific types are provided in the options parameter, only content controls of supported types are returned. Be aware that an exception will be thrown on using methods of a generic [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) that aren't relevant for the specific type. With time, additional types of content controls may be supported. Therefore, your add-in should request and handle specific types of content controls.

#### Examples
```TypeScript
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

### getEndnoteBody()
Gets the document's endnotes in a single body.

```typescript
getEndnoteBody(): Word.Body;
```

Returns
- [Word.Body](/en-us/javascript/api/word/word.body)

Remarks
- [ API set: WordApi 1.5 ]

### getFootnoteBody()
Gets the document's footnotes in a single body.

```typescript
getFootnoteBody(): Word.Body;
```

Returns
- [Word.Body](/en-us/javascript/api/word/word.body)

Remarks
- [ API set: WordApi 1.5 ]

### getParagraphByUniqueLocalId(id)
Gets the paragraph by its unique local ID. Throws an `ItemNotFound` error if the collection is empty.

```typescript
getParagraphByUniqueLocalId(id: string): Word.Paragraph;
```

Parameters
- id: string  
  Required. Unique local ID in standard 8-4-4-4-12 GUID format without curly braces. Note that the ID differs across sessions and coauthors.

Returns
- [Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

Remarks
- [ API set: WordApi 1.6 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/onadded-event.yaml

await Word.run(async (context) => {
  const paragraphId = (document.getElementById("paragraph-id") as HTMLInputElement).value;
  const paragraph: Word.Paragraph = context.document.getParagraphByUniqueLocalId(paragraphId);
  paragraph.load();
  await paragraph.context.sync();

  console.log(paragraph);
});
```

### getSelection()
Gets the current selection of the document. Multiple selections aren't supported.

```typescript
getSelection(): Word.Range;
```

Returns
- [Word.Range](/en-us/javascript/api/word/word.range)

Remarks
- [ API set: WordApi 1.1 ]

#### Examples
```TypeScript
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

### getStyles()
Gets a StyleCollection object that represents the whole style set of the document.

```typescript
getStyles(): Word.StyleCollection;
```

Returns
- [Word.StyleCollection](/en-us/javascript/api/word/word.stylecollection)

Remarks
- [ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Gets the number of available styles stored with the document.
await Word.run(async (context) => {
  const styles: Word.StyleCollection = context.document.getStyles();
  const count = styles.getCount();
  await context.sync();

  console.log(`Number of styles: ${count.value}`);
});
```

### importStylesFromJson(stylesJson, importedStylesConflictBehavior)
Import styles from a JSON-formatted string.

```typescript
importStylesFromJson(stylesJson: string, importedStylesConflictBehavior?: Word.ImportedStylesConflictBehavior): OfficeExtension.ClientResult<string[]>;
```

Parameters
- stylesJson: string  
  Required. A JSON-formatted string representing the styles.
- importedStylesConflictBehavior: [Word.ImportedStylesConflictBehavior](/en-us/javascript/api/word/word.importedstylesconflictbehavior)  
  Optional. Specifies how to handle any imported styles with the same name as existing styles in the current document.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string[]>

Remarks
- [ API set: WordApi 1.6 ]  
  Note: The `importedStylesConflictBehavior` parameter was introduced in WordApiDesktop 1.1.

#### Examples
```TypeScript
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

### importStylesFromJson(stylesJson, importedStylesConflictBehavior)
Import styles from a JSON-formatted string.

```typescript
importStylesFromJson(stylesJson: string, importedStylesConflictBehavior?: "Ignore" | "Overwrite" | "CreateNew"): OfficeExtension.ClientResult<string[]>;
```

Parameters
- stylesJson: string  
  Required. A JSON-formatted string representing the styles.
- importedStylesConflictBehavior: "Ignore" | "Overwrite" | "CreateNew"  
  Optional. Specifies how to handle any imported styles with the same name as existing styles in the current document.

Returns
- [OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string[]>

Remarks
- [ API set: WordApi 1.6 ]  
  Note: The `importedStylesConflictBehavior` parameter was introduced in WordApiDesktop 1.1.

### insertFileFromBase64(base64File, insertLocation, insertFileOptions)
Inserts a document into the target document at a specific location with additional properties. Headers, footers, watermarks, and other section properties are copied by default.

```typescript
insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End", insertFileOptions?: Word.InsertFileOptions): Word.SectionCollection;
```

Parameters
- base64File: string  
  Required. The Base64-encoded content of a .docx file.
- insertLocation: [replace](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-replace-member) | [start](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-start-member) | [end](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-end-member) | "Replace" | "Start" | "End"  
  Required. The value must be 'Replace', 'Start', or 'End'.
- insertFileOptions: [Word.InsertFileOptions](/en-us/javascript/api/word/word.insertfileoptions)  
  Optional. The additional properties that should be imported to the destination document.

Returns
- [Word.SectionCollection](/en-us/javascript/api/word/word.sectioncollection)

Remarks
- [ API set: WordApi 1.5 ]  
  Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.

#### Examples
```TypeScript
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

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.DocumentLoadOptions): Word.Document;
```

Parameters
- options: [Word.Interfaces.DocumentLoadOptions](/en-us/javascript/api/word/word.interfaces.documentloadoptions)  
  Provides options for which properties of the object to load.

Returns
- [Word.Document](/en-us/javascript/api/word/word.document)

#### Examples
```TypeScript
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

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Document;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- [Word.Document](/en-us/javascript/api/word/word.document)

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Document;
```

Parameters
- propertyNamesAndPaths:  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns
- [Word.Document](/en-us/javascript/api/word/word.document)

### manualHyphenation()
Note
- This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Initiates manual hyphenation of a document, one line at a time.

```typescript
manualHyphenation(): void;
```

Returns
- void

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ]

### save(saveBehavior, fileName)
Saves the document.

```typescript
save(saveBehavior?: Word.SaveBehavior, fileName?: string): void;
```

Parameters
- saveBehavior: [Word.SaveBehavior](/en-us/javascript/api/word/word.savebehavior)  
  Optional. The save behavior must be 'Save' or 'Prompt'. Default value is 'Save'.
- fileName: string  
  Optional. The file name (exclude file extension). Only takes effect for a new document.

Returns
- void

Remarks
- [ API set: WordApi 1.1 ]  
  Note: The `saveBehavior` and `fileName` parameters were introduced in WordApi 1.5.

#### Examples
```TypeScript
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
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/save-close.yaml

// Saves the document with default behavior
// for current state of the document.
await Word.run(async (context) => {
  context.document.save();
  await context.sync();
});
```

### save(saveBehavior, fileName)
Saves the document.

```typescript
save(saveBehavior?: "Save" | "Prompt", fileName?: string): void;
```

Parameters
- saveBehavior: "Save" | "Prompt"  
  Optional. The save behavior must be 'Save' or 'Prompt'. Default value is 'Save'.
- fileName: string  
  Optional. The file name (exclude file extension). Only takes effect for a new document.

Returns
- void

Remarks
- [ API set: WordApi 1.1 ]  
  Note: The `saveBehavior` and `fileName` parameters were introduced in WordApi 1.5.

### search(searchText, searchOptions)
Performs a search with the specified search options on the scope of the whole document. The search results are a collection of range objects.

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
- searchText: string
- searchOptions: [Word.SearchOptions](/en-us/javascript/api/word/word.searchoptions) | `{ ignorePunct?: boolean; ignoreSpace?: boolean; matchCase?: boolean; matchPrefix?: boolean; matchSuffix?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean; }`

Returns
- [Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

Remarks
- [ API set: WordApi 1.7 ]

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.DocumentUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: [Word.Interfaces.DocumentUpdateData](/en-us/javascript/api/word/word.interfaces.documentupdatedata)  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions)  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns
- void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Document): void;
```

Parameters
- properties: [Word.Document](/en-us/javascript/api/word/word.document)

Returns
- void

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Document` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.DocumentData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.DocumentData;
```

Returns
- [Word.Interfaces.DocumentData](/en-us/javascript/api/word/word.interfaces.documentdata)

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Document;
```

Returns
- [Word.Document](/en-us/javascript/api/word/word.document)

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.Document;
```

Returns
- [Word.Document](/en-us/javascript/api/word/word.document)

## Event Details

### onAnnotationClicked
Occurs when the user clicks an annotation (or selects it using **Alt+Down**).

```typescript
readonly onAnnotationClicked: OfficeExtension.EventHandlers<Word.AnnotationClickedEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.AnnotationClickedEventArgs](/en-us/javascript/api/word/word.annotationclickedeventargs)>

Remarks
- [ API set: WordApi 1.7 ]

#### Examples
```TypeScript
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

...

async function onClickedHandler(args: Word.AnnotationClickedEventArgs) {
  await Word.run(async (context) => {
    const annotation: Word.Annotation = context.document.getAnnotationById(args.id);
    annotation.load("critiqueAnnotation");

    await context.sync();

    console.log(`AnnotationClicked: ID ${args.id}:`, annotation.critiqueAnnotation.critique);
  });
}
```

### onAnnotationHovered
Occurs when the user hovers the cursor over an annotation.

```typescript
readonly onAnnotationHovered: OfficeExtension.EventHandlers<Word.AnnotationHoveredEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.AnnotationHoveredEventArgs](/en-us/javascript/api/word/word.annotationhoveredeventargs)>

Remarks
- [ API set: WordApi 1.7 ]

#### Examples
```TypeScript
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

...

async function onHoveredHandler(args: Word.AnnotationHoveredEventArgs) {
  await Word.run(async (context) => {
    const annotation: Word.Annotation = context.document.getAnnotationById(args.id);
    annotation.load("critiqueAnnotation");

    await context.sync();

    console.log(`AnnotationHovered: ID ${args.id}:`, annotation.critiqueAnnotation.critique);
  });
}
```

### onAnnotationInserted
Occurs when the user adds one or more annotations.

```typescript
readonly onAnnotationInserted: OfficeExtension.EventHandlers<Word.AnnotationInsertedEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.AnnotationInsertedEventArgs](/en-us/javascript/api/word/word.annotationinsertedeventargs)>

Remarks
- [ API set: WordApi 1.7 ]

#### Examples
```TypeScript
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

...

async function onInsertedHandler(args: Word.AnnotationInsertedEventArgs) {
  await Word.run(async (context) => {
    const annotations = [];
    for (let i = 0; i < args.ids.length; i++) {
      let annotation: Word.Annotation = context.document.getAnnotationById(args.ids[i]);
      annotation.load("id,critiqueAnnotation");

      annotations.push(annotation);
    }

    await context.sync();

    for (let annotation of annotations) {
      console.log(`AnnotationInserted: ID ${annotation.id}:`, annotation.critiqueAnnotation.critique);
    }
  });
}
```

### onAnnotationPopupAction
Occurs when the user performs an action in an annotation pop-up menu.

```typescript
readonly onAnnotationPopupAction: OfficeExtension.EventHandlers<Word.AnnotationPopupActionEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.AnnotationPopupActionEventArgs](/en-us/javascript/api/word/word.annotationpopupactioneventargs)>

Remarks
- [ API set: WordApi 1.8 ]

#### Examples
```TypeScript
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

...

async function onPopupActionHandler(args: Word.AnnotationPopupActionEventArgs) {
  await Word.run(async (context) => {
    let message = `AnnotationPopupAction: ID ${args.id} = `;
    if (args.action === "Accept") {
      message += `Accepted: ${args.critiqueSuggestion}`;
    } else {
      message += "Rejected";
    }

    console.log(message);
  });
}
```

### onAnnotationRemoved
Occurs when the user deletes one or more annotations.

```typescript
readonly onAnnotationRemoved: OfficeExtension.EventHandlers<Word.AnnotationRemovedEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.AnnotationRemovedEventArgs](/en-us/javascript/api/word/word.annotationremovedeventargs)>

Remarks
- [ API set: WordApi 1.7 ]

#### Examples
```TypeScript
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

...

async function onRemovedHandler(args: Word.AnnotationRemovedEventArgs) {
  await Word.run(async (context) => {
    for (let id of args.ids) {
      console.log(`AnnotationRemoved: ID ${id}`);
    }
  });
}
```

### onContentControlAdded
Occurs when a content control is added. Run context.sync() in the handler to get the new content control's properties.

```typescript
readonly onContentControlAdded: OfficeExtension.EventHandlers<Word.ContentControlAddedEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ContentControlAddedEventArgs](/en-us/javascript/api/word/word.contentcontroladdedeventargs)>

Remarks
- [ API set: WordApi 1.5 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-onadded-event.yaml

// Registers the onAdded event handler on the document.
await Word.run(async (context) => {
  eventContext = context.document.onContentControlAdded.add(contentControlAdded);
  await context.sync();

  console.log("Added event handler for when content controls are added.");
});

...

async function contentControlAdded(event: Word.ContentControlAddedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls that were added:`, event.ids);
  });
}
```

### onParagraphAdded
Occurs when the user adds new paragraphs.

```typescript
readonly onParagraphAdded: OfficeExtension.EventHandlers<Word.ParagraphAddedEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ParagraphAddedEventArgs](/en-us/javascript/api/word/word.paragraphaddedeventargs)>

Remarks
- [ API set: WordApi 1.6 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/onadded-event.yaml

// Registers the onParagraphAdded event handler on the document.
await Word.run(async (context) => {
  eventContext = context.document.onParagraphAdded.add(paragraphAdded);
  await context.sync();

  console.log("Added event handler for when paragraphs are added.");
});

...

async function paragraphAdded(event: Word.ParagraphAddedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. IDs of paragraphs that were added:`, event.uniqueLocalIds);
  });
}
```

### onParagraphChanged
Occurs when the user changes paragraphs.

```typescript
readonly onParagraphChanged: OfficeExtension.EventHandlers<Word.ParagraphChangedEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ParagraphChangedEventArgs](/en-us/javascript/api/word/word.paragraphchangedeventargs)>

Remarks
- [ API set: WordApi 1.6 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/onchanged-event.yaml

// Registers the onParagraphChanged event handler on the document.
await Word.run(async (context) => {
  eventContext = context.document.onParagraphChanged.add(paragraphChanged);
  await context.sync();

  console.log("Added event handler for when content is changed in paragraphs.");
});

...

async function paragraphChanged(event: Word.ParagraphChangedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. IDs of paragraphs where content was changed:`, event.uniqueLocalIds);
  });
}
```

### onParagraphDeleted
Occurs when the user deletes paragraphs.

```typescript
readonly onParagraphDeleted: OfficeExtension.EventHandlers<Word.ParagraphDeletedEventArgs>;
```

Event Type
- [OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.ParagraphDeletedEventArgs](/en-us/javascript/api/word/word.paragraphdeletedeventargs)>

Remarks
- [ API set: WordApi 1.6 ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/ondeleted-event.yaml

// Registers the onParagraphDeleted event handler on the document.
await Word.run(async (context) => {
  eventContext = context.document.onParagraphDeleted.add(paragraphDeleted);
  await context.sync();

  console.log("Added event handlers for when paragraphs are deleted.");
});

...

async function paragraphDeleted(event: Word.ParagraphDeletedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. IDs of paragraphs that were deleted:`, event.uniqueLocalIds);
  });
}
```