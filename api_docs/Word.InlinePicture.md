# Word.InlinePicture class

Represents an inline picture.

- Package: [word](/en-us/javascript/api/word)
- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.1 ]

### Examples

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

- altTextDescription — Specifies a string that represents the alternative text associated with the inline image.
- altTextTitle — Specifies a string that contains the title for the inline image.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- height — Specifies a number that describes the height of the inline image.
- hyperlink — Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.
- imageFormat — Gets the format of the inline image.
- lockAspectRatio — Specifies a value that indicates whether the inline image retains its original proportions when you resize it.
- paragraph — Gets the parent paragraph that contains the inline image.
- parentContentControl — Gets the content control that contains the inline image. Throws an ItemNotFound error if there isn't a parent content control.
- parentContentControlOrNullObject — Gets the content control that contains the inline image. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTable — Gets the table that contains the inline image. Throws an ItemNotFound error if it isn't contained in a table.
- parentTableCell — Gets the table cell that contains the inline image. Throws an ItemNotFound error if it isn't contained in a table cell.
- parentTableCellOrNullObject — Gets the table cell that contains the inline image. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- parentTableOrNullObject — Gets the table that contains the inline image. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- width — Specifies a number that describes the width of the inline image.

## Methods

- delete() — Deletes the inline picture from the document.
- getBase64ImageSrc() — Gets the Base64-encoded string representation of the inline image.
- getNext() — Gets the next inline image. Throws an ItemNotFound error if this inline image is the last one.
- getNextOrNullObject() — Gets the next inline image. If this inline image is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- getRange(rangeLocation) — Gets the picture, or the starting or ending point of the picture, as a range.
- insertBreak(breakType, insertLocation) — Inserts a break at the specified location in the main document.
- insertContentControl() — Wraps the inline picture with a rich text content control.
- insertFileFromBase64(base64File, insertLocation) — Inserts a document at the specified location.
- insertHtml(html, insertLocation) — Inserts HTML at the specified location.
- insertInlinePictureFromBase64(base64EncodedImage, insertLocation) — Inserts an inline picture at the specified location.
- insertOoxml(ooxml, insertLocation) — Inserts OOXML at the specified location.
- insertParagraph(paragraphText, insertLocation) — Inserts a paragraph at the specified location.
- insertText(text, insertLocation) — Inserts text at the specified location.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- select(selectionMode) — Selects the inline picture. This causes Word to scroll to the selection.
- select(selectionMode) — Selects the inline picture. This causes Word to scroll to the selection.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- track() — Track the object for automatic adjustment based on surrounding changes in the document.
- untrack() — Release the memory associated with this object, if it has previously been tracked.

## Property Details

### altTextDescription

Specifies a string that represents the alternative text associated with the inline image.

```typescript
altTextDescription: string;
```

#### Property Value

string

#### Remarks

[ API set: WordApi 1.1 ]

---

### altTextTitle

Specifies a string that contains the title for the inline image.

```typescript
altTextTitle: string;
```

#### Property Value

string

#### Remarks

[ API set: WordApi 1.1 ]

---

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

#### Property Value

[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

---

### height

Specifies a number that describes the height of the inline image.

```typescript
height: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi 1.1 ]

---

### hyperlink

Specifies a hyperlink on the image. Use a '#' to separate the address part from the optional location part.

```typescript
hyperlink: string;
```

#### Property Value

string

#### Remarks

[ API set: WordApi 1.1 ]

---

### imageFormat

Gets the format of the inline image.

```typescript
readonly imageFormat: Word.ImageFormat | "Unsupported" | "Undefined" | "Bmp" | "Jpeg" | "Gif" | "Tiff" | "Png" | "Icon" | "Exif" | "Wmf" | "Emf" | "Pict" | "Pdf" | "Svg";
```

#### Property Value

[Word.ImageFormat](/en-us/javascript/api/word/word.imageformat) | "Unsupported" | "Undefined" | "Bmp" | "Jpeg" | "Gif" | "Tiff" | "Png" | "Icon" | "Exif" | "Wmf" | "Emf" | "Pict" | "Pdf" | "Svg"

#### Remarks

[ API set: WordApiDesktop 1.1 ]

#### Examples

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

Specifies a value that indicates whether the inline image retains its original proportions when you resize it.

```typescript
lockAspectRatio: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi 1.1 ]

---

### paragraph

Gets the parent paragraph that contains the inline image.

```typescript
readonly paragraph: Word.Paragraph;
```

#### Property Value

[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

#### Remarks

[ API set: WordApi 1.2 ]

---

### parentContentControl

Gets the content control that contains the inline image. Throws an ItemNotFound error if there isn't a parent content control.

```typescript
readonly parentContentControl: Word.ContentControl;
```

#### Property Value

[Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

#### Remarks

[ API set: WordApi 1.1 ]

---

### parentContentControlOrNullObject

Gets the content control that contains the inline image. If there isn't a parent content control, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentContentControlOrNullObject: Word.ContentControl;
```

#### Property Value

[Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentTable

Gets the table that contains the inline image. Throws an ItemNotFound error if it isn't contained in a table.

```typescript
readonly parentTable: Word.Table;
```

#### Property Value

[Word.Table](/en-us/javascript/api/word/word.table)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentTableCell

Gets the table cell that contains the inline image. Throws an ItemNotFound error if it isn't contained in a table cell.

```typescript
readonly parentTableCell: Word.TableCell;
```

#### Property Value

[Word.TableCell](/en-us/javascript/api/word/word.tablecell)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentTableCellOrNullObject

Gets the table cell that contains the inline image. If it isn't contained in a table cell, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentTableCellOrNullObject: Word.TableCell;
```

#### Property Value

[Word.TableCell](/en-us/javascript/api/word/word.tablecell)

#### Remarks

[ API set: WordApi 1.3 ]

---

### parentTableOrNullObject

Gets the table that contains the inline image. If it isn't contained in a table, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentTableOrNullObject: Word.Table;
```

#### Property Value

[Word.Table](/en-us/javascript/api/word/word.table)

#### Remarks

[ API set: WordApi 1.3 ]

---

### width

Specifies a number that describes the width of the inline image.

```typescript
width: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi 1.1 ]

## Method Details

### delete()

Deletes the inline picture from the document.

```typescript
delete(): void;
```

#### Returns

void

#### Remarks

[ API set: WordApi 1.2 ]

---

### getBase64ImageSrc()

Gets the Base64-encoded string representation of the inline image.

```typescript
getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
```

#### Returns

[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

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

### getNext()

Gets the next inline image. Throws an ItemNotFound error if this inline image is the last one.

```typescript
getNext(): Word.InlinePicture;
```

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

#### Remarks

[ API set: WordApi 1.3 ]

#### Examples

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

### getNextOrNullObject()

Gets the next inline image. If this inline image is the last one, then this method will return an object with its isNullObject property set to true. For further information, see [OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getNextOrNullObject(): Word.InlinePicture;
```

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

#### Remarks

[ API set: WordApi 1.3 ]

#### Examples

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

### getRange(rangeLocation)

Gets the picture, or the starting or ending point of the picture, as a range.

```typescript
getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | "Whole" | "Start" | "End"): Word.Range;
```

#### Parameters

- rangeLocation — Optional. The range location must be 'Whole', 'Start', or 'End'.

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.3 ]

---

### insertBreak(breakType, insertLocation)

Inserts a break at the specified location in the main document.

```typescript
insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): void;
```

#### Parameters

- breakType — Required. The break type to add.
- insertLocation — Required. The value must be 'Before' or 'After'.

#### Returns

void

#### Remarks

[ API set: WordApi 1.2 ]

---

### insertContentControl()

Wraps the inline picture with a rich text content control.

```typescript
insertContentControl(): Word.ContentControl;
```

#### Returns

[Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

#### Remarks

[ API set: WordApi 1.1 ]

---

### insertFileFromBase64(base64File, insertLocation)

Inserts a document at the specified location.

```typescript
insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
```

#### Parameters

- base64File — Required. The Base64-encoded content of a .docx file.
- insertLocation — Required. The value must be 'Before' or 'After'.

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.2 ]

Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.

---

### insertHtml(html, insertLocation)

Inserts HTML at the specified location.

```typescript
insertHtml(html: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
```

#### Parameters

- html — Required. The HTML to be inserted.
- insertLocation — Required. The value must be 'Before' or 'After'.

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.2 ]

---

### insertInlinePictureFromBase64(base64EncodedImage, insertLocation)

Inserts an inline picture at the specified location.

```typescript
insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.before | Word.InsertLocation.after | "Replace" | "Before" | "After"): Word.InlinePicture;
```

#### Parameters

- base64EncodedImage — Required. The Base64-encoded image to be inserted.
- insertLocation — Required. The value must be 'Replace', 'Before', or 'After'.

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

#### Remarks

[ API set: WordApi 1.2 ]

---

### insertOoxml(ooxml, insertLocation)

Inserts OOXML at the specified location.

```typescript
insertOoxml(ooxml: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
```

#### Parameters

- ooxml — Required. The OOXML to be inserted.
- insertLocation — Required. The value must be 'Before' or 'After'.

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.2 ]

---

### insertParagraph(paragraphText, insertLocation)

Inserts a paragraph at the specified location.

```typescript
insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
```

#### Parameters

- paragraphText — Required. The paragraph text to be inserted.
- insertLocation — Required. The value must be 'Before' or 'After'.

#### Returns

[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

#### Remarks

[ API set: WordApi 1.2 ]

---

### insertText(text, insertLocation)

Inserts text at the specified location.

```typescript
insertText(text: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Range;
```

#### Parameters

- text — Required. Text to be inserted.
- insertLocation — Required. The value must be 'Before' or 'After'.

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.2 ]

---

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.InlinePictureLoadOptions): Word.InlinePicture;
```

#### Parameters

- options — Provides options for which properties of the object to load.

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

---

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.InlinePicture;
```

#### Parameters

- propertyNames — A comma-delimited string or an array of strings that specify the properties to load.

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

---

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.InlinePicture;
```

#### Parameters

- propertyNamesAndPaths — propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

---

### select(selectionMode)

Selects the inline picture. This causes Word to scroll to the selection.

```typescript
select(selectionMode?: Word.SelectionMode): void;
```

#### Parameters

- selectionMode — Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

#### Returns

void

#### Remarks

[ API set: WordApi 1.2 ]

---

### select(selectionMode)

Selects the inline picture. This causes Word to scroll to the selection.

```typescript
select(selectionMode?: "Select" | "Start" | "End"): void;
```

#### Parameters

- selectionMode — Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default.

#### Returns

void

#### Remarks

[ API set: WordApi 1.2 ]

---

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.InlinePictureUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

#### Parameters

- properties — [Word.Interfaces.InlinePictureUpdateData](/en-us/javascript/api/word/word.interfaces.inlinepictureupdatedata). A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options — [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions). Provides an option to suppress errors if the properties object tries to set any read-only properties.

#### Returns

void

---

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.InlinePicture): void;
```

#### Parameters

- properties — [Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

#### Returns

void

---

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.InlinePicture object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.InlinePictureData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.InlinePictureData;
```

#### Returns

[Word.Interfaces.InlinePictureData](/en-us/javascript/api/word/word.interfaces.inlinepicturedata)

---

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.InlinePicture;
```

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

---

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.InlinePicture;
```

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)