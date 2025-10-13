# Word.Range class

**Package:** [word](/en-us/javascript/api/word)

Represents a contiguous area in a document.

**Extends:** [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the range of the first comment in the selected content.
await Word.run(async (context) => {
  const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
  comment.load("contentRange");
  const range: Word.Range = comment.getRange();
  range.load("text");
  await context.sync();

  if (comment.isNullObject) {
    console.warn("No comments in the selection, so no range to get.");
    return;
  }

  console.log(`Comment location: ${range.text}`);
  const contentRange: Word.CommentContentRange = comment.contentRange;
  console.log("Comment content range:", contentRange);
});
```

## Properties

| Property | Description |
|---|---|
| [bold](#word-word-range-bold-member) | Specifies whether the range is formatted as bold. |
| [boldBidirectional](#word-word-range-boldbidirectional-member) | Specifies whether the range is formatted as bold in a right-to-left language document. |
| [bookmarks](#word-word-range-bookmarks-member) | Returns a `BookmarkCollection` object that represents all the bookmarks in the range. |
| [borders](#word-word-range-borders-member) | Returns a `BorderUniversalCollection` object that represents all the borders for the range. |
| [case](#word-word-range-case-member) | Specifies a `CharacterCase` value that represents the case of the text in the range. |
| [characterWidth](#word-word-range-characterwidth-member) | Specifies the character width of the range. |
| [combineCharacters](#word-word-range-combinecharacters-member) | Specifies if the range contains combined characters. |
| [contentControls](#word-word-range-contentcontrols-member) | Gets the collection of content control objects in the range. |
| [context](#word-word-range-context-member) | The request context associated with the object. This connects the add-in's process to the Office host application's process. |
| [disableCharacterSpaceGrid](#word-word-range-disablecharacterspacegrid-member) | Specifies if Microsoft Word ignores the number of characters per line for the corresponding `Range` object. |
| [emphasisMark](#word-word-range-emphasismark-member) | Specifies the emphasis mark for a character or designated character string. |
| [end](#word-word-range-end-member) | Specifies the ending character position of the range. |
| [endnotes](#word-word-range-endnotes-member) | Gets the collection of endnotes in the range. |
| [fields](#word-word-range-fields-member) | Gets the collection of field objects in the range. |
| [fitTextWidth](#word-word-range-fittextwidth-member) | Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range. |
| [font](#word-word-range-font-member) | Gets the text format of the range. Use this to get and set font name, size, color, and other properties. |
| [footnotes](#word-word-range-footnotes-member) | Gets the collection of footnotes in the range. |
| [frames](#word-word-range-frames-member) | Gets a `FrameCollection` object that represents all the frames in the range. |
| [grammarChecked](#word-word-range-grammarchecked-member) | Specifies if a grammar check has been run on the range or document. |
| [hasNoProofing](#word-word-range-hasnoproofing-member) | Specifies the proofing status (spelling and grammar checking) of the range. |
| [highlightColorIndex](#word-word-range-highlightcolorindex-member) | Specifies the highlight color for the range. |
| [horizontalInVertical](#word-word-range-horizontalinvertical-member) | Specifies the formatting for horizontal text set within vertical text. |
| [hyperlink](#word-word-range-hyperlink-member) | Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part. |
| [hyperlinks](#word-word-range-hyperlinks-member) | Returns a `HyperlinkCollection` object that represents all the hyperlinks in the range. |
| [id](#word-word-range-id-member) | Specifies the ID for the range. |
| [inlinePictures](#word-word-range-inlinepictures-member) | Gets the collection of inline picture objects in the range. |
| [isEmpty](#word-word-range-isempty-member) | Checks whether the range length is zero. |
| [isEndOfRowMark](#word-word-range-isendofrowmark-member) | Gets if the range is collapsed and is located at the end-of-row mark in a table. |
| [isTextVisibleOnScreen](#word-word-range-istextvisibleonscreen-member) | Gets whether the text in the range is visible on the screen. |
| [italic](#word-word-range-italic-member) | Specifies if the font or range is formatted as italic. |
| [italicBidirectional](#word-word-range-italicbidirectional-member) | Specifies if the font or range is formatted as italic (right-to-left languages). |
| [kana](#word-word-range-kana-member) | Specifies whether the range of Japanese language text is hiragana or katakana. |
| [languageDetected](#word-word-range-languagedetected-member) | Specifies whether Microsoft Word has detected the language of the text in the range. |
| [languageId](#word-word-range-languageid-member) | Specifies a `LanguageId` value that represents the language for the range. |
| [languageIdFarEast](#word-word-range-languageidfareast-member) | Specifies an East Asian language for the range. |
| [languageIdOther](#word-word-range-languageidother-member) | Specifies a language for the range that isn't classified as an East Asian language. |
| [listFormat](#word-word-range-listformat-member) | Returns a `ListFormat` object that represents all the list formatting characteristics of the range. |
| [lists](#word-word-range-lists-member) | Gets the collection of list objects in the range. |
| [pages](#word-word-range-pages-member) | Gets the collection of pages in the range. |
| [paragraphs](#word-word-range-paragraphs-member) | Gets the collection of paragraph objects in the range. |
| [parentBody](#word-word-range-parentbody-member) | Gets the parent body of the range. |
| [parentContentControl](#word-word-range-parentcontentcontrol-member) | Gets the currently supported content control that contains the range. Throws an `ItemNotFound` error if there isn't a parent content control. |
| [parentContentControlOrNullObject](#word-word-range-parentcontentcontrolornullobject-member) | Gets the currently supported content control that contains the range. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties). |
| [parentTable](#word-word-range-parenttable-member) | Gets the table that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table. |
| [parentTableCell](#word-word-range-parenttablecell-member) | Gets the table cell that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table cell. |
| [parentTableCellOrNullObject](#word-word-range-parenttablecellornullobject-member) | Gets the table cell that contains the range. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties). |
| [parentTableOrNullObject](#word-word-range-parenttableornullobject-member) | Gets the table that contains the range. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties). |
| [sections](#word-word-range-sections-member) | Gets the collection of sections in the range. |
| [shading](#word-word-range-shading-member) | Returns a `ShadingUniversal` object that refers to the shading formatting for the range. |
| [shapes](#word-word-range-shapes-member) | Gets the collection of shape objects anchored in the range, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases. |
| [showAll](#word-word-range-showall-member) | Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed. |
| [spellingChecked](#word-word-range-spellingchecked-member) | Specifies if spelling has been checked throughout the range or document. |
| [start](#word-word-range-start-member) | Specifies the starting character position of the range. |
| [storyLength](#word-word-range-storylength-member) | Gets the number of characters in the story that contains the range. |
| [storyType](#word-word-range-storytype-member) | Gets the story type for the range. |
| [style](#word-word-range-style-member) | Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property. |
| [styleBuiltIn](#word-word-range-stylebuiltin-member) | Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property. |
| [tableColumns](#word-word-range-tablecolumns-member) | Gets a `TableColumnCollection` object that represents all the table columns in the range. |
| [tables](#word-word-range-tables-member) | Gets the collection of table objects in the range. |
| [text](#word-word-range-text-member) | Gets the text of the range. |
| [twoLinesInOne](#word-word-range-twolinesinone-member) | Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any. |
| [underline](#word-word-range-underline-member) | Specifies the type of underline applied to the range. |

## Methods

| Method | Description |
|---|---|
| [clear()](#word-word-range-clear-member(1)) | Clears the contents of the range object. The user can perform the undo operation on the cleared content. |
| [compareLocationWith(range)](#word-word-range-comparelocationwith-member(1)) | Compares this range's location with another range's location. |
| [delete()](#word-word-range-delete-member(1)) | Deletes the range and its content from the document. |
| [detectLanguage()](#word-word-range-detectlanguage-member(1)) | Analyzes the range text to determine the language that it's written in. |
| [expandTo(range)](#word-word-range-expandto-member(1)) | Returns a new range that extends from this range in either direction to cover another range. This range isn't changed. Throws an `ItemNotFound` error if the two ranges don't have a union. |
| [expandToOrNullObject(range)](#word-word-range-expandtoornullobject-member(1)) | Returns a new range that extends from this range in either direction to cover another range. This range isn't changed. If the two ranges don't have a union, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties). |
| [getBookmarks(includeHidden, includeAdjacent)](#word-word-range-getbookmarks-member(1)) | Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with the underscore character. |
| [getComments()](#word-word-range-getcomments-member(1)) | Gets comments associated with the range. |
| [getContentControls(options)](#word-word-range-getcontentcontrols-member(1)) | Gets the currently supported content controls in the range. |
| [getHtml()](#word-word-range-gethtml-member(1)) | Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML. |
| [getHyperlinkRanges()](#word-word-range-gethyperlinkranges-member(1)) | Gets hyperlink child ranges within the range. |
| [getNextTextRange(endingMarks, trimSpacing)](#word-word-range-getnexttextrange-member(1)) | Gets the next text range by using punctuation marks and/or other ending marks. Throws an `ItemNotFound` error if this text range is the last one. |
| [getNextTextRangeOrNullObject(endingMarks, trimSpacing)](#word-word-range-getnexttextrangeornullobject-member(1)) | Gets the next text range by using punctuation marks and/or other ending marks. If this text range is the last one, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties). |
| [getOoxml()](#word-word-range-getooxml-member(1)) | Gets the OOXML representation of the range object. |
| [getRange(rangeLocation)](#word-word-range-getrange-member(1)) | Clones the range, or gets the starting or ending point of the range as a new range. |
| [getReviewedText(changeTrackingVersion)](#word-word-range-getreviewedtext-member(1)) | Gets reviewed text based on ChangeTrackingVersion selection. |
| [getReviewedText(changeTrackingVersion)](#word-word-range-getreviewedtext-member(2)) | Gets reviewed text based on ChangeTrackingVersion selection. |
| [getTextRanges(endingMarks, trimSpacing)](#word-word-range-gettextranges-member(1)) | Gets the text child ranges in the range by using punctuation marks and/or other ending marks. |
| [getTrackedChanges()](#word-word-range-gettrackedchanges-member(1)) | Gets the collection of the TrackedChange objects in the range. |
| [highlight()](#word-word-range-highlight-member(1)) | Highlights the range temporarily without changing document content. To highlight the text permanently, set the range's Font.HighlightColor. |
| [insertBookmark(name)](#word-word-range-insertbookmark-member(1)) | Inserts a bookmark on the range. If a bookmark of the same name exists somewhere, it is deleted first. |
| [insertBreak(breakType, insertLocation)](#word-word-range-insertbreak-member(1)) | Inserts a break at the specified location in the main document. |
| [insertCanvas(insertShapeOptions)](#word-word-range-insertcanvas-member(1)) | Inserts a floating canvas in front of text with its anchor at the beginning of the range. |
| [insertComment(commentText)](#word-word-range-insertcomment-member(1)) | Insert a comment on the range. |
| [insertContentControl(contentControlType)](#word-word-range-insertcontentcontrol-member(1)) | Wraps the Range object with a content control. |
| [insertEndnote(insertText)](#word-word-range-insertendnote-member(1)) | Inserts an endnote. The endnote reference is placed after the range. |
| [insertField(insertLocation, fieldType, text, removeFormatting)](#word-word-range-insertfield-member(1)) | Inserts a field at the specified location. |
| [insertField(insertLocation, fieldType, text, removeFormatting)](#word-word-range-insertfield-member(2)) | Inserts a field at the specified location. |
| [insertFileFromBase64(base64File, insertLocation)](#word-word-range-insertfilefrombase64-member(1)) | Inserts a document at the specified location. |
| [insertFootnote(insertText)](#word-word-range-insertfootnote-member(1)) | Inserts a footnote. The footnote reference is placed after the range. |
| [insertGeometricShape(geometricShapeType, insertShapeOptions)](#word-word-range-insertgeometricshape-member(1)) | Inserts a geometric shape in front of text with its anchor at the beginning of the range. |
| [insertGeometricShape(geometricShapeType, insertShapeOptions)](#word-word-range-insertgeometricshape-member(2)) | Inserts a geometric shape in front of text with its anchor at the beginning of the range. |
| [insertHtml(html, insertLocation)](#word-word-range-inserthtml-member(1)) | Inserts HTML at the specified location. |
| [insertInlinePictureFromBase64(base64EncodedImage, insertLocation)](#word-word-range-insertinlinepicturefrombase64-member(1)) | Inserts a picture at the specified location. |
| [insertOoxml(ooxml, insertLocation)](#word-word-range-insertooxml-member(1)) | Inserts OOXML at the specified location. |
| [insertParagraph(paragraphText, insertLocation)](#word-word-range-insertparagraph-member(1)) | Inserts a paragraph at the specified location. |
| [insertPictureFromBase64(base64EncodedImage, insertShapeOptions)](#word-word-range-insertpicturefrombase64-member(1)) | Inserts a floating picture in front of text with its anchor at the beginning of the range. |
| [insertTable(rowCount, columnCount, insertLocation, values)](#word-word-range-inserttable-member(1)) | Inserts a table with the specified number of rows and columns. |
| [insertText(text, insertLocation)](#word-word-range-inserttext-member(1)) | Inserts text at the specified location. |
| [insertTextBox(text, insertShapeOptions)](#word-word-range-inserttextbox-member(1)) | Inserts a floating text box in front of text with its anchor at the beginning of the range. |
| [intersectWith(range)](#word-word-range-intersectwith-member(1)) | Returns a new range as the intersection of this range with another range. This range isn't changed. Throws an `ItemNotFound` error if the two ranges aren't overlapped or adjacent. |
| [intersectWithOrNullObject(range)](#word-word-range-intersectwithornullobject-member(1)) | Returns a new range as the intersection of this range with another range. This range isn't changed. If the two ranges aren't overlapped or adjacent, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties). |
| [load(options)](#word-word-range-load-member(1)) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [load(propertyNames)](#word-word-range-load-member(2)) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [load(propertyNamesAndPaths)](#word-word-range-load-member(3)) | Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties. |
| [removeHighlight()](#word-word-range-removehighlight-member(1)) | Removes the highlight added by the Highlight function if any. |
| [search(searchText, searchOptions)](#word-word-range-search-member(1)) | Performs a search with the specified SearchOptions on the scope of the range object. The search results are a collection of range objects. |
| [select(selectionMode)](#word-word-range-select-member(1)) | Selects and navigates the Word UI to the range. |
| [select(selectionMode)](#word-word-range-select-member(2)) | Selects and navigates the Word UI to the range. |
| [set(properties, options)](#word-word-range-set-member(1)) | Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type. |
| [set(properties)](#word-word-range-set-member(2)) | Sets multiple properties on the object at the same time, based on an existing loaded object. |
| [split(delimiters, multiParagraphs, trimDelimiters, trimSpacing)](#word-word-range-split-member(1)) | Splits the range into child ranges by using delimiters. |
| [toJSON()](#word-word-range-tojson-member(1)) | Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Range` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeData`) that contains shallow copies of any loaded child properties from the original object. |
| [track()](#word-word-range-track-member(1)) | Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection. |
| [untrack()](#word-word-range-untrack-member(1)) | Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect. |

## Events

| Event | Description |
|---|---|
| [onCommentAdded](#word-word-range-oncommentadded-member) | Occurs when new comments are added. |
| [onCommentChanged](#word-word-range-oncommentchanged-member) | Occurs when a comment or its reply is changed. |
| [onCommentDeselected](#word-word-range-oncommentdeselected-member) | Occurs when a comment is deselected. |
| [onCommentSelected](#word-word-range-oncommentselected-member) | Occurs when a comment is selected. |

## Property Details

### bold

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range is formatted as bold.

```typescript
readonly bold: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### boldBidirectional

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range is formatted as bold in a right-to-left language document.

```typescript
readonly boldBidirectional: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### bookmarks

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BookmarkCollection` object that represents all the bookmarks in the range.

```typescript
readonly bookmarks: Word.BookmarkCollection;
```

#### Property Value

[Word.BookmarkCollection](/en-us/javascript/api/word/word.bookmarkcollection)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### borders

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BorderUniversalCollection` object that represents all the borders for the range.

```typescript
readonly borders: Word.BorderUniversalCollection;
```

#### Property Value

[Word.BorderUniversalCollection](/en-us/javascript/api/word/word.borderuniversalcollection)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### case

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `CharacterCase` value that represents the case of the text in the range.

```typescript
case: Word.CharacterCase | "Next" | "Lower" | "Upper" | "TitleWord" | "TitleSentence" | "Toggle" | "HalfWidth" | "FullWidth" | "Katakana" | "Hiragana";
```

#### Property Value

[Word.CharacterCase](/en-us/javascript/api/word/word.charactercase) | "Next" | "Lower" | "Upper" | "TitleWord" | "TitleSentence" | "Toggle" | "HalfWidth" | "FullWidth" | "Katakana" | "Hiragana"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### characterWidth

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the character width of the range.

```typescript
characterWidth: Word.CharacterWidth | "Half" | "Full";
```

#### Property Value

[Word.CharacterWidth](/en-us/javascript/api/word/word.characterwidth) | "Half" | "Full"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### combineCharacters

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the range contains combined characters.

```typescript
combineCharacters: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### contentControls

Gets the collection of content control objects in the range.

```typescript
readonly contentControls: Word.ContentControlCollection;
```

#### Property Value

[Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

#### Remarks

[ API set: WordApi 1.1 ]

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

#### Property Value

[Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### disableCharacterSpaceGrid

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if Microsoft Word ignores the number of characters per line for the corresponding `Range` object.

```typescript
readonly disableCharacterSpaceGrid: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### emphasisMark

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the emphasis mark for a character or designated character string.

```typescript
readonly emphasisMark: Word.EmphasisMark | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle";
```

#### Property Value

[Word.EmphasisMark](/en-us/javascript/api/word/word.emphasismark) | "None" | "OverSolidCircle" | "OverComma" | "OverWhiteCircle" | "UnderSolidCircle"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### end

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ending character position of the range.

```typescript
end: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### endnotes

Gets the collection of endnotes in the range.

```typescript
readonly endnotes: Word.NoteItemCollection;
```

#### Property Value

[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

#### Remarks

[ API set: WordApi 1.5 ]

### fields

Gets the collection of field objects in the range.

```typescript
readonly fields: Word.FieldCollection;
```

#### Property Value

[Word.FieldCollection](/en-us/javascript/api/word/word.fieldcollection)

#### Remarks

[ API set: WordApi 1.4 ]

### fitTextWidth

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width (in the current measurement units) in which Microsoft Word fits the text in the current selection or range.

```typescript
fitTextWidth: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### font

Gets the text format of the range. Use this to get and set font name, size, color, and other properties.

```typescript
readonly font: Word.Font;
```

#### Property Value

[Word.Font](/en-us/javascript/api/word/word.font)

#### Remarks

[ API set: WordApi 1.1 ]

### footnotes

Gets the collection of footnotes in the range.

```typescript
readonly footnotes: Word.NoteItemCollection;
```

#### Property Value

[Word.NoteItemCollection](/en-us/javascript/api/word/word.noteitemcollection)

#### Remarks

[ API set: WordApi 1.5 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Gets the footnotes in the selected document range.
await Word.run(async (context) => {
  const footnotes: Word.NoteItemCollection = context.document.getSelection().footnotes;
  footnotes.load("length");
  await context.sync();

  console.log("Number of footnotes in the selected range: " + footnotes.items.length);
});
```

### frames

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `FrameCollection` object that represents all the frames in the range.

```typescript
readonly frames: Word.FrameCollection;```

#### Property Value

[Word.FrameCollection](/en-us/javascript/api/word/word.framecollection)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### grammarChecked

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if a grammar check has been run on the range or document.

```typescript
grammarChecked: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### hasNoProofing

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the proofing status (spelling and grammar checking) of the range.

```typescript
hasNoProofing: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### highlightColorIndex

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the highlight color for the range.

```typescript
readonly highlightColorIndex: Word.ColorIndex | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor";
```

#### Property Value

[Word.ColorIndex](/en-us/javascript/api/word/word.colorindex) | "Auto" | "Black" | "Blue" | "Turquoise" | "BrightGreen" | "Pink" | "Red" | "Yellow" | "White" | "DarkBlue" | "Teal" | "Green" | "Violet" | "DarkRed" | "DarkYellow" | "Gray50" | "Gray25" | "ClassicRed" | "ClassicBlue" | "ByAuthor"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### horizontalInVertical

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the formatting for horizontal text set within vertical text.

```typescript
horizontalInVertical: Word.HorizontalInVerticalType | "None" | "FitInLine" | "ResizeLine";
```

#### Property Value

[Word.HorizontalInVerticalType](/en-us/javascript/api/word/word.horizontalinverticaltype) | "None" | "FitInLine" | "ResizeLine"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### hyperlink

Gets the first hyperlink in the range, or sets a hyperlink on the range. All hyperlinks in the range are deleted when you set a new hyperlink on the range. Use a '#' to separate the address part from the optional location part.

```typescript
hyperlink: string;
```

#### Property Value

string

#### Remarks

[ API set: WordApi 1.3 ]

### hyperlinks

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `HyperlinkCollection` object that represents all the hyperlinks in the range.

```typescript
readonly hyperlinks: Word.HyperlinkCollection;
```

#### Property Value

[Word.HyperlinkCollection](/en-us/javascript/api/word/word.hyperlinkcollection)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### id

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the ID for the range.

```typescript
id: string;
```

#### Property Value

string

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### inlinePictures

Gets the collection of inline picture objects in the range.

```typescript
readonly inlinePictures: Word.InlinePictureCollection;
```

#### Property Value

[Word.InlinePictureCollection](/en-us/javascript/api/word/word.inlinepicturecollection)

#### Remarks

[ API set: WordApi 1.2 ]

### isEmpty

Checks whether the range length is zero.

```typescript
readonly isEmpty: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi 1.3 ]

### isEndOfRowMark

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets if the range is collapsed and is located at the end-of-row mark in a table.

```typescript
readonly isEndOfRowMark: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### isTextVisibleOnScreen

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the text in the range is visible on the screen.

```typescript
readonly isTextVisibleOnScreen: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### italic

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font or range is formatted as italic.

```typescript
readonly italic: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### italicBidirectional

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the font or range is formatted as italic (right-to-left languages).

```typescript
readonly italicBidirectional: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### kana

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the range of Japanese language text is hiragana or katakana.

```typescript
kana: Word.Kana | "Katakana" | "Hiragana";
```

#### Property Value

[Word.Kana](/en-us/javascript/api/word/word.kana) | "Katakana" | "Hiragana"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### languageDetected

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word has detected the language of the text in the range.

```typescript
languageDetected: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### languageId

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `LanguageId` value that represents the language for the range.

```typescript
languageId: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

#### Property Value

[Word.LanguageId](/en-us/javascript/api/word/word.languageid) | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### languageIdFarEast

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies an East Asian language for the range.

```typescript
languageIdFarEast: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

#### Property Value

[Word.LanguageId](/en-us/javascript/api/word/word.languageid) | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### languageIdOther

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a language for the range that isn't classified as an East Asian language.

```typescript
languageIdOther: Word.LanguageId | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu";
```

#### Property Value

[Word.LanguageId](/en-us/javascript/api/word/word.languageid) | "Afrikaans" | "Albanian" | "Amharic" | "Arabic" | "ArabicAlgeria" | "ArabicBahrain" | "ArabicEgypt" | "ArabicIraq" | "ArabicJordan" | "ArabicKuwait" | "ArabicLebanon" | "ArabicLibya" | "ArabicMorocco" | "ArabicOman" | "ArabicQatar" | "ArabicSyria" | "ArabicTunisia" | "ArabicUAE" | "ArabicYemen" | "Armenian" | "Assamese" | "AzeriCyrillic" | "AzeriLatin" | "Basque" | "BelgianDutch" | "BelgianFrench" | "Bengali" | "Bulgarian" | "Burmese" | "Belarusian" | "Catalan" | "Cherokee" | "ChineseHongKongSAR" | "ChineseMacaoSAR" | "ChineseSingapore" | "Croatian" | "Czech" | "Danish" | "Divehi" | "Dutch" | "Edo" | "EnglishAUS" | "EnglishBelize" | "EnglishCanadian" | "EnglishCaribbean" | "EnglishIndonesia" | "EnglishIreland" | "EnglishJamaica" | "EnglishNewZealand" | "EnglishPhilippines" | "EnglishSouthAfrica" | "EnglishTrinidadTobago" | "EnglishUK" | "EnglishUS" | "EnglishZimbabwe" | "Estonian" | "Faeroese" | "Filipino" | "Finnish" | "French" | "FrenchCameroon" | "FrenchCanadian" | "FrenchCongoDRC" | "FrenchCotedIvoire" | "FrenchHaiti" | "FrenchLuxembourg" | "FrenchMali" | "FrenchMonaco" | "FrenchMorocco" | "FrenchReunion" | "FrenchSenegal" | "FrenchWestIndies" | "FrisianNetherlands" | "Fulfulde" | "GaelicIreland" | "GaelicScotland" | "Galician" | "Georgian" | "German" | "GermanAustria" | "GermanLiechtenstein" | "GermanLuxembourg" | "Greek" | "Guarani" | "Gujarati" | "Hausa" | "Hawaiian" | "Hebrew" | "Hindi" | "Hungarian" | "Ibibio" | "Icelandic" | "Igbo" | "Indonesian" | "Inuktitut" | "Italian" | "Japanese" | "Kannada" | "Kanuri" | "Kashmiri" | "Kazakh" | "Khmer" | "Kirghiz" | "Konkani" | "Korean" | "Kyrgyz" | "LanguageNone" | "Lao" | "Latin" | "Latvian" | "Lithuanian" | "MacedonianFYROM" | "Malayalam" | "MalayBruneiDarussalam" | "Malaysian" | "Maltese" | "Manipuri" | "Marathi" | "MexicanSpanish" | "Mongolian" | "Nepali" | "NoProofing" | "NorwegianBokmol" | "NorwegianNynorsk" | "Oriya" | "Oromo" | "Pashto" | "Persian" | "Polish" | "Portuguese" | "PortugueseBrazil" | "Punjabi" | "RhaetoRomanic" | "Romanian" | "RomanianMoldova" | "Russian" | "RussianMoldova" | "SamiLappish" | "Sanskrit" | "SerbianCyrillic" | "SerbianLatin" | "Sesotho" | "SimplifiedChinese" | "Sindhi" | "SindhiPakistan" | "Sinhalese" | "Slovak" | "Slovenian" | "Somali" | "Sorbian" | "Spanish" | "SpanishArgentina" | "SpanishBolivia" | "SpanishChile" | "SpanishColombia" | "SpanishCostaRica" | "SpanishDominicanRepublic" | "SpanishEcuador" | "SpanishElSalvador" | "SpanishGuatemala" | "SpanishHonduras" | "SpanishModernSort" | "SpanishNicaragua" | "SpanishPanama" | "SpanishParaguay" | "SpanishPeru" | "SpanishPuertoRico" | "SpanishUruguay" | "SpanishVenezuela" | "Sutu" | "Swahili" | "Swedish" | "SwedishFinland" | "SwissFrench" | "SwissGerman" | "SwissItalian" | "Syriac" | "Tajik" | "Tamazight" | "TamazightLatin" | "Tamil" | "Tatar" | "Telugu" | "Thai" | "Tibetan" | "TigrignaEritrea" | "TigrignaEthiopic" | "TraditionalChinese" | "Tsonga" | "Tswana" | "Turkish" | "Turkmen" | "Ukrainian" | "Urdu" | "UzbekCyrillic" | "UzbekLatin" | "Venda" | "Vietnamese" | "Welsh" | "Xhosa" | "Yi" | "Yiddish" | "Yoruba" | "Zulu"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### listFormat

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ListFormat` object that represents all the list formatting characteristics of the range.

```typescript
readonly listFormat: Word.ListFormat;
```

#### Property Value

[Word.ListFormat](/en-us/javascript/api/word/word.listformat)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### lists

Gets the collection of list objects in the range.

```typescript
readonly lists: Word.ListCollection;
```

#### Property Value

[Word.ListCollection](/en-us/javascript/api/word/word.listcollection)

#### Remarks

[ API set: WordApi 1.3 ]

### pages

Gets the collection of pages in the range.

```typescript
readonly pages: Word.PageCollection;
```

#### Property Value

[Word.PageCollection](/en-us/javascript/api/word/word.pagecollection)

#### Remarks

[ API set: WordApiDesktop 1.2 ]

### paragraphs

Gets the collection of paragraph objects in the range.

```typescript
readonly paragraphs: Word.ParagraphCollection;
```

#### Property Value

[Word.ParagraphCollection](/en-us/javascript/api/word/word.paragraphcollection)

#### Remarks

[ API set: WordApi 1.1 ]

Important: For requirement sets 1.1 and 1.2, paragraphs in tables wholly contained within this range aren't returned. From requirement set 1.3, paragraphs in such tables are also returned.

### parentBody

Gets the parent body of the range.

```typescript
readonly parentBody: Word.Body;
```

#### Property Value

[Word.Body](/en-us/javascript/api/word/word.body)

#### Remarks

[ API set: WordApi 1.3 ]

### parentContentControl

Gets the currently supported content control that contains the range. Throws an `ItemNotFound` error if there isn't a parent content control.

```typescript
readonly parentContentControl: Word.ContentControl;
```

#### Property Value

[Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
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

### parentContentControlOrNullObject

Gets the currently supported content control that contains the range. If there isn't a parent content control, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentContentControlOrNullObject: Word.ContentControl;
```

#### Property Value

[Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

#### Remarks

[ API set: WordApi 1.3 ]

### parentTable

Gets the table that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table.

```typescript
readonly parentTable: Word.Table;
```

#### Property Value

[Word.Table](/en-us/javascript/api/word/word.table)

#### Remarks

[ API set: WordApi 1.3 ]

### parentTableCell

Gets the table cell that contains the range. Throws an `ItemNotFound` error if it isn't contained in a table cell.

```typescript
readonly parentTableCell: Word.TableCell;
```

#### Property Value

[Word.TableCell](/en-us/javascript/api/word/word.tablecell)

#### Remarks

[ API set: WordApi 1.3 ]

### parentTableCellOrNullObject

Gets the table cell that contains the range. If it isn't contained in a table cell, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentTableCellOrNullObject: Word.TableCell;
```

#### Property Value

[Word.TableCell](/en-us/javascript/api/word/word.tablecell)

#### Remarks

[ API set: WordApi 1.3 ]

### parentTableOrNullObject

Gets the table that contains the range. If it isn't contained in a table, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
readonly parentTableOrNullObject: Word.Table;
```

#### Property Value

[Word.Table](/en-us/javascript/api/word/word.table)

#### Remarks

[ API set: WordApi 1.3 ]

### sections

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the collection of sections in the range.

```typescript
readonly sections: Word.SectionCollection;
```

#### Property Value

[Word.SectionCollection](/en-us/javascript/api/word/word.sectioncollection)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### shading

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `ShadingUniversal` object that refers to the shading formatting for the range.

```typescript
readonly shading: Word.ShadingUniversal;
```

#### Property Value

[Word.ShadingUniversal](/en-us/javascript/api/word/word.shadinguniversal)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### shapes

Gets the collection of shape objects anchored in the range, including both inline and floating shapes. Currently, only the following shapes are supported: text boxes, geometric shapes, groups, pictures, and canvases.

```typescript
readonly shapes: Word.ShapeCollection;
```

#### Property Value

[Word.ShapeCollection](/en-us/javascript/api/word/word.shapecollection)

#### Remarks

[ API set: WordApiDesktop 1.2 ]

### showAll

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if all nonprinting characters (such as hidden text, tab marks, space marks, and paragraph marks) are displayed.

```typescript
showAll: boolean;```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### spellingChecked

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if spelling has been checked throughout the range or document.

```typescript
spellingChecked: boolean;
```

#### Property Value

boolean

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### start

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the starting character position of the range.

```typescript
start: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### storyLength

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the number of characters in the story that contains the range.

```typescript
readonly storyLength: number;
```

#### Property Value

number

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### storyType

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the story type for the range.

```typescript
readonly storyType: Word.StoryType | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice";
```

#### Property Value

[Word.StoryType](/en-us/javascript/api/word/word.storytype) | "MainText" | "Footnotes" | "Endnotes" | "Comments" | "TextFrame" | "EvenPagesHeader" | "PrimaryHeader" | "EvenPagesFooter" | "PrimaryFooter" | "FirstPageHeader" | "FirstPageFooter" | "FootnoteSeparator" | "FootnoteContinuationSeparator" | "FootnoteContinuationNotice" | "EndnoteSeparator" | "EndnoteContinuationSeparator" | "EndnoteContinuationNotice"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### style

Specifies the style name for the range. Use this property for custom styles and localized style names. To use the built-in styles that are portable between locales, see the "styleBuiltIn" property.

```typescript
style: string;
```

#### Property Value

string

#### Remarks

[ API set: WordApi 1.1 ]

### styleBuiltIn

Specifies the built-in style name for the range. Use this property for built-in styles that are portable between locales. To use custom styles or localized style names, see the "style" property.

```typescript
styleBuiltIn: Word.BuiltInStyleName | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6";
```

#### Property Value

[Word.BuiltInStyleName](/en-us/javascript/api/word/word.builtinstylename) | "Other" | "Normal" | "Heading1" | "Heading2" | "Heading3" | "Heading4" | "Heading5" | "Heading6" | "Heading7" | "Heading8" | "Heading9" | "Toc1" | "Toc2" | "Toc3" | "Toc4" | "Toc5" | "Toc6" | "Toc7" | "Toc8" | "Toc9" | "FootnoteText" | "Header" | "Footer" | "Caption" | "FootnoteReference" | "EndnoteReference" | "EndnoteText" | "Title" | "Subtitle" | "Hyperlink" | "Strong" | "Emphasis" | "NoSpacing" | "ListParagraph" | "Quote" | "IntenseQuote" | "SubtleEmphasis" | "IntenseEmphasis" | "SubtleReference" | "IntenseReference" | "BookTitle" | "Bibliography" | "TocHeading" | "TableGrid" | "PlainTable1" | "PlainTable2" | "PlainTable3" | "PlainTable4" | "PlainTable5" | "TableGridLight" | "GridTable1Light" | "GridTable1Light_Accent1" | "GridTable1Light_Accent2" | "GridTable1Light_Accent3" | "GridTable1Light_Accent4" | "GridTable1Light_Accent5" | "GridTable1Light_Accent6" | "GridTable2" | "GridTable2_Accent1" | "GridTable2_Accent2" | "GridTable2_Accent3" | "GridTable2_Accent4" | "GridTable2_Accent5" | "GridTable2_Accent6" | "GridTable3" | "GridTable3_Accent1" | "GridTable3_Accent2" | "GridTable3_Accent3" | "GridTable3_Accent4" | "GridTable3_Accent5" | "GridTable3_Accent6" | "GridTable4" | "GridTable4_Accent1" | "GridTable4_Accent2" | "GridTable4_Accent3" | "GridTable4_Accent4" | "GridTable4_Accent5" | "GridTable4_Accent6" | "GridTable5Dark" | "GridTable5Dark_Accent1" | "GridTable5Dark_Accent2" | "GridTable5Dark_Accent3" | "GridTable5Dark_Accent4" | "GridTable5Dark_Accent5" | "GridTable5Dark_Accent6" | "GridTable6Colorful" | "GridTable6Colorful_Accent1" | "GridTable6Colorful_Accent2" | "GridTable6Colorful_Accent3" | "GridTable6Colorful_Accent4" | "GridTable6Colorful_Accent5" | "GridTable6Colorful_Accent6" | "GridTable7Colorful" | "GridTable7Colorful_Accent1" | "GridTable7Colorful_Accent2" | "GridTable7Colorful_Accent3" | "GridTable7Colorful_Accent4" | "GridTable7Colorful_Accent5" | "GridTable7Colorful_Accent6" | "ListTable1Light" | "ListTable1Light_Accent1" | "ListTable1Light_Accent2" | "ListTable1Light_Accent3" | "ListTable1Light_Accent4" | "ListTable1Light_Accent5" | "ListTable1Light_Accent6" | "ListTable2" | "ListTable2_Accent1" | "ListTable2_Accent2" | "ListTable2_Accent3" | "ListTable2_Accent4" | "ListTable2_Accent5" | "ListTable2_Accent6" | "ListTable3" | "ListTable3_Accent1" | "ListTable3_Accent2" | "ListTable3_Accent3" | "ListTable3_Accent4" | "ListTable3_Accent5" | "ListTable3_Accent6" | "ListTable4" | "ListTable4_Accent1" | "ListTable4_Accent2" | "ListTable4_Accent3" | "ListTable4_Accent4" | "ListTable4_Accent5" | "ListTable4_Accent6" | "ListTable5Dark" | "ListTable5Dark_Accent1" | "ListTable5Dark_Accent2" | "ListTable5Dark_Accent3" | "ListTable5Dark_Accent4" | "ListTable5Dark_Accent5" | "ListTable5Dark_Accent6" | "ListTable6Colorful" | "ListTable6Colorful_Accent1" | "ListTable6Colorful_Accent2" | "ListTable6Colorful_Accent3" | "ListTable6Colorful_Accent4" | "ListTable6Colorful_Accent5" | "ListTable6Colorful_Accent6" | "ListTable7Colorful" | "ListTable7Colorful_Accent1" | "ListTable7Colorful_Accent2" | "ListTable7Colorful_Accent3" | "ListTable7Colorful_Accent4" | "ListTable7Colorful_Accent5" | "ListTable7Colorful_Accent6"

#### Remarks

[ API set: WordApi 1.3 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/90-scenarios/doc-assembly.yaml

await Word.run(async (context) => {
    const header: Word.Range = context.document.body.insertText("This is a sample Heading 1 Title!!\n",
        "Start" /*this means at the beginning of the body */);
    header.styleBuiltIn = Word.BuiltInStyleName.heading1;

    await context.sync();
});```

### tableColumns

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `TableColumnCollection` object that represents all the table columns in the range.

```typescript
readonly tableColumns: Word.TableColumnCollection;
```

#### Property Value

[Word.TableColumnCollection](/en-us/javascript/api/word/word.tablecolumncollection)

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### tables

Gets the collection of table objects in the range.

```typescript
readonly tables: Word.TableCollection;
```

#### Property Value

[Word.TableCollection](/en-us/javascript/api/word/word.tablecollection)

#### Remarks

[ API set: WordApi 1.3 ]

### text

Gets the text of the range.

```typescript
readonly text: string;
```

#### Property Value

string

#### Remarks

[ API set: WordApi 1.1 ]

### twoLinesInOne

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word sets two lines of text in one and specifies the characters that enclose the text, if any.

```typescript
twoLinesInOne: Word.TwoLinesInOneType | "None" | "NoBrackets" | "Parentheses" | "SquareBrackets" | "AngleBrackets" | "CurlyBrackets";
```

#### Property Value

[Word.TwoLinesInOneType](/en-us/javascript/api/word/word.twolinesinonetype) | "None" | "NoBrackets" | "Parentheses" | "SquareBrackets" | "AngleBrackets" | "CurlyBrackets"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### underline

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the type of underline applied to the range.

```typescript
readonly underline: Word.Underline | "None" | "Single" | "Words" | "Double" | "Dotted" | "Thick" | "Dash" | "DotDash" | "DotDotDash" | "Wavy" | "WavyHeavy" | "DottedHeavy" | "DashHeavy" | "DotDashHeavy" | "DotDotDashHeavy" | "DashLong" | "DashLongHeavy" | "WavyDouble";
```

#### Property Value

[Word.Underline](/en-us/javascript/api/word/word.underline) | "None" | "Single" | "Words" | "Double" | "Dotted" | "Thick" | "Dash" | "DotDash" | "DotDotDash" | "Wavy" | "WavyHeavy" | "DottedHeavy" | "DashHeavy" | "DotDashHeavy" | "DotDotDashHeavy" | "DashLong" | "DashLongHeavy" | "WavyDouble"

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Method Details

### clear()

Clears the contents of the range object. The user can perform the undo operation on the cleared content.

```typescript
clear(): void;
```

#### Returns

void

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to clear the contents of the proxy range object.
    range.clear();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Cleared the selection (range object)');
});```

### compareLocationWith(range)

Compares this range's location with another range's location.

```typescript
compareLocationWith(range: Word.Range): OfficeExtension.ClientResult<Word.LocationRelation>;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| range | [Word.Range](/en-us/javascript/api/word/word.range) | Required. The range to compare with this range. |

#### Returns

[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<[Word.LocationRelation](/en-us/javascript/api/word/word.locationrelation)>

#### Remarks

[ API set: WordApi 1.3 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/compare-location.yaml

// Compares the location of one paragraph in relation to another paragraph.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("items");

  await context.sync();

  const firstParagraphAsRange: Word.Range = paragraphs.items[0].getRange();
  const secondParagraphAsRange: Word.Range = paragraphs.items[1].getRange();

  const comparedLocation = firstParagraphAsRange.compareLocationWith(secondParagraphAsRange);

  await context.sync();

  const locationValue: Word.LocationRelation = comparedLocation.value;
  console.log(`Location of the first paragraph in relation to the second paragraph: ${locationValue}`);
});```

### delete()

Deletes the range and its content from the document.

```typescript
delete(): void;
```

#### Returns

void

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to delete the range object.
    range.delete();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Deleted the selection (range object)');
});
```

### detectLanguage()

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Analyzes the range text to determine the language that it's written in.

```typescript
detectLanguage(): OfficeExtension.ClientResult<boolean>;
```

#### Returns

[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<boolean>

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### expandTo(range)

Returns a new range that extends from this range in either direction to cover another range. This range isn't changed. Throws an `ItemNotFound` error if the two ranges don't have a union.

```typescript
expandTo(range: Word.Range): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| range | [Word.Range](/en-us/javascript/api/word/word.range) | Required. Another range. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.3 ]

#### Examples

```TypeScript
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

### expandToOrNullObject(range)

Returns a new range that extends from this range in either direction to cover another range. This range isn't changed. If the two ranges don't have a union, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
expandToOrNullObject(range: Word.Range): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| range | [Word.Range](/en-us/javascript/api/word/word.range) | Required. Another range. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.3 ]

### getBookmarks(includeHidden, includeAdjacent)

Gets the names all bookmarks in or overlapping the range. A bookmark is hidden if its name starts with the underscore character.

```typescript
getBookmarks(includeHidden?: boolean, includeAdjacent?: boolean): OfficeExtension.ClientResult<string[]>;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| includeHidden | boolean | Optional. Indicates whether to include hidden bookmarks. Default is false which indicates that the hidden bookmarks are excluded. |
| includeAdjacent | boolean | Optional. Indicates whether to include bookmarks that are adjacent to the range. Default is false which indicates that the adjacent bookmarks are excluded. |

#### Returns

[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string[]>

#### Remarks

[ API set: WordApi 1.4 ]

### getComments()

Gets comments associated with the range.

```typescript
getComments(): Word.CommentCollection;
```

#### Returns

[Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)

#### Remarks

[ API set: WordApi 1.4 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the comments in the selected content.
await Word.run(async (context) => {
  const comments: Word.CommentCollection = context.document.getSelection().getComments();

  // Load objects to log in the console.
  comments.load();
  await context.sync();

  console.log("Comments:", comments);
});
```

### getContentControls(options)

Gets the currently supported content controls in the range.

```typescript
getContentControls(options?: Word.ContentControlOptions): Word.ContentControlCollection;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| options | [Word.ContentControlOptions](/en-us/javascript/api/word/word.contentcontroloptions) | Optional. Options that define which content controls are returned. |

#### Returns

[Word.ContentControlCollection](/en-us/javascript/api/word/word.contentcontrolcollection)

#### Remarks

[ API set: WordApi 1.5 ]

**Important**: If specific types are provided in the options parameter, only content controls of supported types are returned. Be aware that an exception will be thrown on using methods of a generic [Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol) that aren't relevant for the specific type. With time, additional types of content controls may be supported. Therefore, your add-in should request and handle specific types of content controls.

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml

// Deletes the first checkbox content control found in the selection.
await Word.run(async (context) => {
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.checkBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id");

  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.checkBox) {
      console.warn("No checkbox content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  console.log(`About to delete checkbox content control with id: ${selectedContentControl.id}`);
  selectedContentControl.delete(false);
  await context.sync();

  console.log("Deleted checkbox content control.");
});
```

### getHtml()

Gets an HTML representation of the range object. When rendered in a web page or HTML viewer, the formatting will be a close, but not exact, match for of the formatting of the document. This method doesn't return the exact same HTML for the same document on different platforms (Windows, Mac, Word on the web, etc.). If you need exact fidelity, or consistency across platforms, use `Range.getOoxml()` and convert the returned XML to HTML.

```typescript
getHtml(): OfficeExtension.ClientResult<string>;
```

#### Returns

[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to get the HTML of the current selection.
    const html = range.getHtml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The HTML read from the document was: ' + html.value);
});
```

### getHyperlinkRanges()

Gets hyperlink child ranges within the range.

```typescript
getHyperlinkRanges(): Word.RangeCollection;
```

#### Returns

[Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

#### Remarks

[ API set: WordApi 1.3 ]

#### Examples

```TypeScript
await Word.run(async (context) => {
    // Get the entire document body.
    const bodyRange = context.document.body.getRange(Word.RangeLocation.whole);

    // Get all the ranges that only consist of hyperlinks.
    const hyperLinks = bodyRange.getHyperlinkRanges();
    hyperLinks.load("hyperlink");
    await context.sync();

    // Log each hyperlink.
    hyperLinks.items.forEach((linkRange) => {
        console.log(linkRange.hyperlink);
    });
});
```

### getNextTextRange(endingMarks, trimSpacing)

Gets the next text range by using punctuation marks and/or other ending marks. Throws an `ItemNotFound` error if this text range is the last one.

```typescript
getNextTextRange(endingMarks: string[], trimSpacing?: boolean): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| endingMarks | string[] | Required. The punctuation marks and/or other ending marks as an array of strings. |
| trimSpacing | boolean | Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.3 ]

### getNextTextRangeOrNullObject(endingMarks, trimSpacing)

Gets the next text range by using punctuation marks and/or other ending marks. If this text range is the last one, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getNextTextRangeOrNullObject(endingMarks: string[], trimSpacing?: boolean): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| endingMarks | string[] | Required. The punctuation marks and/or other ending marks as an array of strings. |
| trimSpacing | boolean | Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the returned range. Default is false which indicates that spacing characters at the start and end of the range are included. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.3 ]

### getOoxml()

Gets the OOXML representation of the range object.

```typescript
getOoxml(): OfficeExtension.ClientResult<string>;
```

#### Returns

[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to get the OOXML of the current selection.
    const ooxml = range.getOoxml();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The OOXML read from the document was:  ' + ooxml.value);
});
```

### getRange(rangeLocation)

Clones the range, or gets the starting or ending point of the range as a new range.

```typescript
getRange(rangeLocation?: Word.RangeLocation.whole | Word.RangeLocation.start | Word.RangeLocation.end | Word.RangeLocation.after | Word.RangeLocation.content | "Whole" | "Start" | "End" | "After" | "Content"): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| rangeLocation | [whole](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-whole-member) \| [start](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-start-member) \| [end](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-end-member) \| [after](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-after-member) \| [content](/en-us/javascript/api/word/word.rangelocation#word-word-rangelocation-content-member) \| "Whole" \| "Start" \| "End" \| "After" \| "Content" | Optional. The range location must be 'Whole', 'Start', 'End', 'After', or 'Content'. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.3 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-dropdown-list-content-control.yaml

// Places a dropdown list content control at the end of the selection.
await Word.run(async (context) => {
  let selection = context.document.getSelection();
  selection.getRange(Word.RangeLocation.end).insertContentControl(Word.ContentControlType.dropDownList);
  await context.sync();

  console.log("Dropdown list content control inserted at the end of the selection.");
});
```

### getReviewedText(changeTrackingVersion)

Gets reviewed text based on ChangeTrackingVersion selection.

```typescript
getReviewedText(changeTrackingVersion?: Word.ChangeTrackingVersion): OfficeExtension.ClientResult<string>;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| changeTrackingVersion | [Word.ChangeTrackingVersion](/en-us/javascript/api/word/word.changetrackingversion) | Optional. The value must be 'Original' or 'Current'. The default is 'Current'. |

#### Returns

[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

#### Remarks

[ API set: WordApi 1.4 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-change-tracking.yaml

// Gets the reviewed text.
await Word.run(async (context) => {
  const range: Word.Range = context.document.getSelection();
  const before = range.getReviewedText(Word.ChangeTrackingVersion.original);
  const after = range.getReviewedText(Word.ChangeTrackingVersion.current);

  await context.sync();

  console.log("Reviewed text (before):", before.value, "Reviewed text (after):", after.value);
});
```

### getReviewedText(changeTrackingVersion)

Gets reviewed text based on ChangeTrackingVersion selection.

```typescript
getReviewedText(changeTrackingVersion?: "Original" | "Current"): OfficeExtension.ClientResult<string>;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| changeTrackingVersion | "Original" \| "Current" | Optional. The value must be 'Original' or 'Current'. The default is 'Current'. |

#### Returns

[OfficeExtension.ClientResult](/en-us/javascript/api/office/officeextension.clientresult)<string>

#### Remarks

[ API set: WordApi 1.4 ]

### getTextRanges(endingMarks, trimSpacing)

Gets the text child ranges in the range by using punctuation marks and/or other ending marks.

```typescript
getTextRanges(endingMarks: string[], trimSpacing?: boolean): Word.RangeCollection;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| endingMarks | string[] | Required. The punctuation marks and/or other ending marks as an array of strings. |
| trimSpacing | boolean | Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection. |

#### Returns

[Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

#### Remarks

[ API set: WordApi 1.3 ]

#### Examples

```TypeScript
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
});```

### getTrackedChanges()

Gets the collection of the TrackedChange objects in the range.

```typescript
getTrackedChanges(): Word.TrackedChangeCollection;
```

#### Returns

[Word.TrackedChangeCollection](/en-us/javascript/api/word/word.trackedchangecollection)

#### Remarks

[ API set: WordApi 1.6 ]

### highlight()

Highlights the range temporarily without changing document content. To highlight the text permanently, set the range's Font.HighlightColor.

```typescript
highlight(): void;
```

#### Returns

void

#### Remarks

[ API set: WordApi 1.8 ]

### insertBookmark(name)

Inserts a bookmark on the range. If a bookmark of the same name exists somewhere, it is deleted first.

```typescript
insertBookmark(name: string): void;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| name | string | Required. The case-insensitive bookmark name. Only alphanumeric and underscore characters are supported. It must begin with a letter but if you want to tag the bookmark as hidden, then start the name with an underscore character. Names can't be longer than 40 characters. |

#### Returns

void

#### Remarks

[ API set: WordApi 1.4 ]

Note: The conditions of inserting a bookmark are similar to doing so in the Word UI. To learn more about managing bookmarks in the Word UI, see [Add or delete bookmarks in a Word document or Outlook message](https://support.microsoft.com/office/f68d781f-0150-4583-a90e-a4009d99c2a0).

### insertBreak(breakType, insertLocation)

Inserts a break at the specified location in the main document.

```typescript
insertBreak(breakType: Word.BreakType | "Page" | "Next" | "SectionNext" | "SectionContinuous" | "SectionEven" | "SectionOdd" | "Line", insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): void;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| breakType | [Word.BreakType](/en-us/javascript/api/word/word.breaktype) \| "Page" \| "Next" \| "SectionNext" \| "SectionContinuous" \| "SectionEven" \| "SectionOdd" \| "Line" | Required. The break type to add. |
| insertLocation | [before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) \| [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) \| "Before" \| "After" | Required. The value must be 'Before' or 'After'. |

#### Returns

void

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert a page break after the selected text.
    range.insertBreak(Word.BreakType.page, Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Inserted a page break after the selected text.');
});
```

### insertCanvas(insertShapeOptions)

Inserts a floating canvas in front of text with its anchor at the beginning of the range.

```typescript
insertCanvas(insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| insertShapeOptions | [Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions) | Optional. The location and size of the canvas. The default location and size is (0, 0, 300, 200). |

#### Returns

[Word.Shape](/en-us/javascript/api/word/word.shape)

#### Remarks

[ API set: WordApiDesktop 1.2 ]

### insertComment(commentText)

Insert a comment on the range.

```typescript
insertComment(commentText: string): Word.Comment;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| commentText | string | Required. The comment text to be inserted. |

#### Returns

[Word.Comment](/en-us/javascript/api/word/word.comment)

comment object

#### Remarks

[ API set: WordApi 1.4 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Sets a comment on the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("comment-text") as HTMLInputElement).value;
  const comment: Word.Comment = context.document.getSelection().insertComment(text);

  // Load object to log in the console.
  comment.load();
  await context.sync();

  console.log("Comment inserted:", comment);
});
```

### insertContentControl(contentControlType)

Wraps the Range object with a content control.

```typescript
insertContentControl(contentControlType?: Word.ContentControlType.richText | Word.ContentControlType.plainText | Word.ContentControlType.checkBox | Word.ContentControlType.dropDownList | Word.ContentControlType.comboBox | "RichText" | "PlainText" | "CheckBox" | "DropDownList" | "ComboBox"): Word.ContentControl;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| contentControlType | [richText](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-richtext-member) \| [plainText](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-plaintext-member) \| [checkBox](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-checkbox-member) \| [dropDownList](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-dropdownlist-member) \| [comboBox](/en-us/javascript/api/word/word.contentcontroltype#word-word-contentcontroltype-combobox-member) \| "RichText" \| "PlainText" \| "CheckBox" \| "DropDownList" \| "ComboBox" | Optional. Content control type to insert. Must be 'RichText', 'PlainText', 'CheckBox', 'DropDownList', or 'ComboBox'. The default is 'RichText'. |

#### Returns

[Word.ContentControl](/en-us/javascript/api/word/word.contentcontrol)

#### Remarks

[ API set: WordApi 1.1 ]

Note: The `contentControlType` parameter was introduced in WordApi 1.5. `PlainText` support was added in WordApi 1.5. `CheckBox` support was added in WordApi 1.7. `DropDownList` and `ComboBox` support was added in WordApi 1.9.

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/90-scenarios/doc-assembly.yaml

// Simulates creation of a template. First searches the document for instances of the string "Contractor",
// then changes the format  of each search result,
// then wraps each search result within a content control,
// finally sets a tag and title property on each content control.
await Word.run(async (context) => {
    const results: Word.RangeCollection = context.document.body.search("Contractor");
    results.load("font/bold");

    // Check to make sure these content controls haven't been added yet.
    const customerContentControls: Word.ContentControlCollection = context.document.contentControls.getByTag("customer");
    customerContentControls.load("text");
    await context.sync();

  if (customerContentControls.items.length === 0) {
    for (let i = 0; i < results.items.length; i++) { 
        results.items[i].font.bold = true;
        let cc: Word.ContentControl = results.items[i].insertContentControl();
        cc.tag = "customer";  // This value is used in the next step of this sample.
        cc.title = "Customer Name " + i;
    }
  }
    await context.sync();
});
```

### insertEndnote(insertText)

Inserts an endnote. The endnote reference is placed after the range.

```typescript
insertEndnote(insertText?: string): Word.NoteItem;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| insertText | string | Optional. Text to be inserted into the endnote body. The default is "". |

#### Returns

[Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

#### Remarks

[ API set: WordApi 1.5 ]

### insertField(insertLocation, fieldType, text, removeFormatting)

Inserts a field at the specified location.

```typescript
insertField(insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After", fieldType?: Word.FieldType, text?: string, removeFormatting?: boolean): Word.Field;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| insertLocation | [Word.InsertLocation](/en-us/javascript/api/word/word.insertlocation) \| "Replace" \| "Start" \| "End" \| "Before" \| "After" | Required. The location relative to the range where the field will be inserted. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'. |
| fieldType | [Word.FieldType](/en-us/javascript/api/word/word.fieldtype) | Optional. Can be any FieldType constant. The default value is Empty. |
| text | string | Optional. Additional properties or options if needed for specified field type. |
| removeFormatting | boolean | Optional. `true` to remove the formatting that's applied to the field during updates, `false` otherwise. The default value is `false`. |

#### Returns

[Word.Field](/en-us/javascript/api/word/word.field)

#### Remarks

[ API set: WordApi 1.5 ]

Important: In Word on Windows and on Mac, the API supports inserting and managing all types listed in [Word.FieldType](/en-us/javascript/api/word/word.fieldtype) except `Word.FieldType.others`. In Word on the web, fields are mainly read-only. To learn more, see [Use fields in your Word add-in](/en-us/office/dev/add-ins/word/fields-guidance).

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Inserts a Date field before selection.
await Word.run(async (context) => {
  const range: Word.Range = context.document.getSelection().getRange();

  const field: Word.Field = range.insertField(Word.InsertLocation.before, Word.FieldType.date, '\\@ "M/d/yyyy h:mm am/pm"', true);

  field.load("result,code");
  await context.sync();

  if (field.isNullObject) {
    console.log("There are no fields in this document.");
  } else {
    console.log("Code of the field: " + field.code, "Result of the field: " + JSON.stringify(field.result));
  }
});
```

**TOC Field Best Practices:**
- For Table of Contents insertion, use simple sequential Range operations rather than complex chained manipulations
- Insert TOC field after title text using `getRange(Word.RangeLocation.end)` for reliable positioning
- Always call `updateResult()` on TOC fields after insertion to populate content
- See complete TOC insertion example in [Word.FieldType.toc](/en-us/javascript/api/word/word.fieldtype#word-word-fieldtype-toc-member)

### insertField(insertLocation, fieldType, text, removeFormatting)

Inserts a field at the specified location.

```typescript
insertField(insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After", fieldType?: "Addin" | "AddressBlock" | "Advance" | "Ask" | "Author" | "AutoText" | "AutoTextList" | "BarCode" | "Bibliography" | "BidiOutline" | "Citation" | "Comments" | "Compare" | "CreateDate" | "Data" | "Database" | "Date" | "DisplayBarcode" | "DocProperty" | "DocVariable" | "EditTime" | "Embedded" | "EQ" | "Expression" | "FileName" | "FileSize" | "FillIn" | "FormCheckbox" | "FormDropdown" | "FormText" | "GotoButton" | "GreetingLine" | "Hyperlink" | "If" | "Import" | "Include" | "IncludePicture" | "IncludeText" | "Index" | "Info" | "Keywords" | "LastSavedBy" | "Link" | "ListNum" | "MacroButton" | "MergeBarcode" | "MergeField" | "MergeRec" | "MergeSeq" | "Next" | "NextIf" | "NoteRef" | "NumChars" | "NumPages" | "NumWords" | "OCX" | "Page" | "PageRef" | "Print" | "PrintDate" | "Private" | "Quote" | "RD" | "Ref" | "RevNum" | "SaveDate" | "Section" | "SectionPages" | "Seq" | "Set" | "Shape" | "SkipIf" | "StyleRef" | "Subject" | "Subscriber" | "Symbol" | "TA" | "TC" | "Template" | "Time" | "Title" | "TOA" | "TOC" | "UserAddress" | "UserInitials" | "UserName" | "XE" | "Empty" | "Others" | "Undefined", text?: string, removeFormatting?: boolean): Word.Field;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| insertLocation | [Word.InsertLocation](/en-us/javascript/api/word/word.insertlocation) \| "Replace" \| "Start" \| "End" \| "Before" \| "After" | Required. The location relative to the range where the field will be inserted. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'. |
| fieldType | "Addin" \| "AddressBlock" \| "Advance" \| "Ask" \| "Author" \| "AutoText" \| "AutoTextList" \| "BarCode" \| "Bibliography" \| "BidiOutline" \| "Citation" \| "Comments" \| "Compare" \| "CreateDate" \| "Data" \| "Database" \| "Date" \| "DisplayBarcode" \| "DocProperty" \| "DocVariable" \| "EditTime" \| "Embedded" \| "EQ" \| "Expression" \| "FileName" \| "FileSize" \| "FillIn" \| "FormCheckbox" \| "FormDropdown" \| "FormText" \| "GotoButton" \| "GreetingLine" \| "Hyperlink" \| "If" \| "Import" \| "Include" \| "IncludePicture" \| "IncludeText" \| "Index" \| "Info" \| "Keywords" \| "LastSavedBy" \| "Link" \| "ListNum" \| "MacroButton" \| "MergeBarcode" \| "MergeField" \| "MergeRec" \| "MergeSeq" \| "Next" \| "NextIf" \| "NoteRef" \| "NumChars" \| "NumPages" \| "NumWords" \| "OCX" \| "Page" \| "PageRef" \| "Print" \| "PrintDate" \| "Private" \| "Quote" \| "RD" \| "Ref" \| "RevNum" \| "SaveDate" \| "Section" \| "SectionPages" \| "Seq" \| "Set" \| "Shape" \| "SkipIf" \| "StyleRef" \| "Subject" \| "Subscriber" \| "Symbol" \| "TA" \| "TC" \| "Template" \| "Time" \| "Title" \| "TOA" \| "TOC" \| "UserAddress" \| "UserInitials" \| "UserName" \| "XE" \| "Empty" \| "Others" \| "Undefined" | Optional. Can be any FieldType constant. The default value is Empty. |
| text | string | Optional. Additional properties or options if needed for specified field type. |
| removeFormatting | boolean | Optional. `true` to remove the formatting that's applied to the field during updates, `false` otherwise. The default value is `false`. |

#### Returns

[Word.Field](/en-us/javascript/api/word/word.field)

#### Remarks

[ API set: WordApi 1.5 ]

Important: In Word on Windows and on Mac, the API supports inserting and managing all types listed in [Word.FieldType](/en-us/javascript/api/word/word.fieldtype) except `Word.FieldType.others`. In Word on the web, fields are mainly read-only. To learn more, see [Use fields in your Word add-in](/en-us/office/dev/add-ins/word/fields-guidance).

### insertFileFromBase64(base64File, insertLocation)

Inserts a document at the specified location.

```typescript
insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| base64File | string | Required. The Base64-encoded content of a .docx file. |
| insertLocation | [Word.InsertLocation](/en-us/javascript/api/word/word.insertlocation) \| "Replace" \| "Start" \| "End" \| "Before" \| "After" | Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.1 ]

Insertion isn't supported if the document being inserted contains an ActiveX control (likely in a form field). Consider replacing such a form field with a content control or other option appropriate for your scenario.

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert base64 encoded .docx at the beginning of the range.
    // You'll need to implement getBase64() to make this work.
    range.insertFileFromBase64(getBase64(), Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Added base64 encoded text to the beginning of the range.');
});
```

### insertFootnote(insertText)

Inserts a footnote. The footnote reference is placed after the range.

```typescript
insertFootnote(insertText?: string): Word.NoteItem;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| insertText | string | Optional. Text to be inserted into the footnote body. The default is "". |

#### Returns

[Word.NoteItem](/en-us/javascript/api/word/word.noteitem)

#### Remarks

[ API set: WordApi 1.5 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-footnotes.yaml

// Sets a footnote on the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("input-footnote") as HTMLInputElement).value;
  const footnote: Word.NoteItem = context.document.getSelection().insertFootnote(text);
  await context.sync();

  console.log("Inserted footnote.");
});
```

### insertGeometricShape(geometricShapeType, insertShapeOptions)

Inserts a geometric shape in front of text with its anchor at the beginning of the range.

```typescript
insertGeometricShape(geometricShapeType: Word.GeometricShapeType, insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| geometricShapeType | [Word.GeometricShapeType](/en-us/javascript/api/word/word.geometricshapetype) | The geometric type of the shape to insert. |
| insertShapeOptions | [Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions) | Optional. The location and size of the geometric shape. The default location and size is (0, 0, 100, 100). |

#### Returns

[Word.Shape](/en-us/javascript/api/word/word.shape)

#### Remarks

[ API set: WordApiDesktop 1.2 ]

### insertGeometricShape(geometricShapeType, insertShapeOptions)

Inserts a geometric shape in front of text with its anchor at the beginning of the range.

```typescript
insertGeometricShape(geometricShapeType: "LineInverse" | "Triangle" | "RightTriangle" | "Rectangle" | "Diamond" | "Parallelogram" | "Trapezoid" | "NonIsoscelesTrapezoid" | "Pentagon" | "Hexagon" | "Heptagon" | "Octagon" | "Decagon" | "Dodecagon" | "Star4" | "Star5" | "Star6" | "Star7" | "Star8" | "Star10" | "Star12" | "Star16" | "Star24" | "Star32" | "RoundRectangle" | "Round1Rectangle" | "Round2SameRectangle" | "Round2DiagonalRectangle" | "SnipRoundRectangle" | "Snip1Rectangle" | "Snip2SameRectangle" | "Snip2DiagonalRectangle" | "Plaque" | "Ellipse" | "Teardrop" | "HomePlate" | "Chevron" | "PieWedge" | "Pie" | "BlockArc" | "Donut" | "NoSmoking" | "RightArrow" | "LeftArrow" | "UpArrow" | "DownArrow" | "StripedRightArrow" | "NotchedRightArrow" | "BentUpArrow" | "LeftRightArrow" | "UpDownArrow" | "LeftUpArrow" | "LeftRightUpArrow" | "QuadArrow" | "LeftArrowCallout" | "RightArrowCallout" | "UpArrowCallout" | "DownArrowCallout" | "LeftRightArrowCallout" | "UpDownArrowCallout" | "QuadArrowCallout" | "BentArrow" | "UturnArrow" | "CircularArrow" | "LeftCircularArrow" | "LeftRightCircularArrow" | "CurvedRightArrow" | "CurvedLeftArrow" | "CurvedUpArrow" | "CurvedDownArrow" | "SwooshArrow" | "Cube" | "Can" | "LightningBolt" | "Heart" | "Sun" | "Moon" | "SmileyFace" | "IrregularSeal1" | "IrregularSeal2" | "FoldedCorner" | "Bevel" | "Frame" | "HalfFrame" | "Corner" | "DiagonalStripe" | "Chord" | "Arc" | "LeftBracket" | "RightBracket" | "LeftBrace" | "RightBrace" | "BracketPair" | "BracePair" | "Callout1" | "Callout2" | "Callout3" | "AccentCallout1" | "AccentCallout2" | "AccentCallout3" | "BorderCallout1" | "BorderCallout2" | "BorderCallout3" | "AccentBorderCallout1" | "AccentBorderCallout2" | "AccentBorderCallout3" | "WedgeRectCallout" | "WedgeRRectCallout" | "WedgeEllipseCallout" | "CloudCallout" | "Cloud" | "Ribbon" | "Ribbon2" | "EllipseRibbon" | "EllipseRibbon2" | "LeftRightRibbon" | "VerticalScroll" | "HorizontalScroll" | "Wave" | "DoubleWave" | "Plus" | "FlowChartProcess" | "FlowChartDecision" | "FlowChartInputOutput" | "FlowChartPredefinedProcess" | "FlowChartInternalStorage" | "FlowChartDocument" | "FlowChartMultidocument" | "FlowChartTerminator" | "FlowChartPreparation" | "FlowChartManualInput" | "FlowChartManualOperation" | "FlowChartConnector" | "FlowChartPunchedCard" | "FlowChartPunchedTape" | "FlowChartSummingJunction" | "FlowChartOr" | "FlowChartCollate" | "FlowChartSort" | "FlowChartExtract" | "FlowChartMerge" | "FlowChartOfflineStorage" | "FlowChartOnlineStorage" | "FlowChartMagneticTape" | "FlowChartMagneticDisk" | "FlowChartMagneticDrum" | "FlowChartDisplay" | "FlowChartDelay" | "FlowChartAlternateProcess" | "FlowChartOffpageConnector" | "ActionButtonBlank" | "ActionButtonHome" | "ActionButtonHelp" | "ActionButtonInformation" | "ActionButtonForwardNext" | "ActionButtonBackPrevious" | "ActionButtonEnd" | "ActionButtonBeginning" | "ActionButtonReturn" | "ActionButtonDocument" | "ActionButtonSound" | "ActionButtonMovie" | "Gear6" | "Gear9" | "Funnel" | "MathPlus" | "MathMinus" | "MathMultiply" | "MathDivide" | "MathEqual" | "MathNotEqual" | "CornerTabs" | "SquareTabs" | "PlaqueTabs" | "ChartX" | "ChartStar" | "ChartPlus", insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| geometricShapeType | "LineInverse" \| "Triangle" \| "RightTriangle" \| "Rectangle" \| "Diamond" \| "Parallelogram" \| "Trapezoid" \| "NonIsoscelesTrapezoid" \| "Pentagon" \| "Hexagon" \| "Heptagon" \| "Octagon" \| "Decagon" \| "Dodecagon" \| "Star4" \| "Star5" \| "Star6" \| "Star7" \| "Star8" \| "Star10" \| "Star12" \| "Star16" \| "Star24" \| "Star32" \| "RoundRectangle" \| "Round1Rectangle" \| "Round2SameRectangle" \| "Round2DiagonalRectangle" \| "SnipRoundRectangle" \| "Snip1Rectangle" \| "Snip2SameRectangle" \| "Snip2DiagonalRectangle" \| "Plaque" \| "Ellipse" \| "Teardrop" \| "HomePlate" \| "Chevron" \| "PieWedge" \| "Pie" \| "BlockArc" \| "Donut" \| "NoSmoking" \| "RightArrow" \| "LeftArrow" \| "UpArrow" \| "DownArrow" \| "StripedRightArrow" \| "NotchedRightArrow" \| "BentUpArrow" \| "LeftRightArrow" \| "UpDownArrow" \| "LeftUpArrow" \| "LeftRightUpArrow" \| "QuadArrow" \| "LeftArrowCallout" \| "RightArrowCallout" \| "UpArrowCallout" \| "DownArrowCallout" \| "LeftRightArrowCallout" \| "UpDownArrowCallout" \| "QuadArrowCallout" \| "BentArrow" \| "UturnArrow" \| "CircularArrow" \| "LeftCircularArrow" \| "LeftRightCircularArrow" \| "CurvedRightArrow" \| "CurvedLeftArrow" \| "CurvedUpArrow" \| "CurvedDownArrow" \| "SwooshArrow" \| "Cube" \| "Can" \| "LightningBolt" \| "Heart" \| "Sun" \| "Moon" \| "SmileyFace" \| "IrregularSeal1" \| "IrregularSeal2" \| "FoldedCorner" \| "Bevel" \| "Frame" \| "HalfFrame" \| "Corner" \| "DiagonalStripe" \| "Chord" \| "Arc" \| "LeftBracket" \| "RightBracket" \| "LeftBrace" \| "RightBrace" \| "BracketPair" \| "BracePair" \| "Callout1" \| "Callout2" \| "Callout3" \| "AccentCallout1" \| "AccentCallout2" \| "AccentCallout3" \| "BorderCallout1" \| "BorderCallout2" \| "BorderCallout3" \| "AccentBorderCallout1" \| "AccentBorderCallout2" \| "AccentBorderCallout3" \| "WedgeRectCallout" \| "WedgeRRectCallout" \| "WedgeEllipseCallout" \| "CloudCallout" \| "Cloud" \| "Ribbon" \| "Ribbon2" \| "EllipseRibbon" \| "EllipseRibbon2" \| "LeftRightRibbon" \| "VerticalScroll" \| "HorizontalScroll" \| "Wave" \| "DoubleWave" \| "Plus" \| "FlowChartProcess" \| "FlowChartDecision" \| "FlowChartInputOutput" \| "FlowChartPredefinedProcess" \| "FlowChartInternalStorage" \| "FlowChartDocument" \| "FlowChartMultidocument" \| "FlowChartTerminator" \| "FlowChartPreparation" \| "FlowChartManualInput" \| "FlowChartManualOperation" \| "FlowChartConnector" \| "FlowChartPunchedCard" \| "FlowChartPunchedTape" \| "FlowChartSummingJunction" \| "FlowChartOr" \| "FlowChartCollate" \| "FlowChartSort" \| "FlowChartExtract" \| "FlowChartMerge" \| "FlowChartOfflineStorage" \| "FlowChartOnlineStorage" \| "FlowChartMagneticTape" \| "FlowChartMagneticDisk" \| "FlowChartMagneticDrum" \| "FlowChartDisplay" \| "FlowChartDelay" \| "FlowChartAlternateProcess" \| "FlowChartOffpageConnector" \| "ActionButtonBlank" \| "ActionButtonHome" \| "ActionButtonHelp" \| "ActionButtonInformation" \| "ActionButtonForwardNext" \| "ActionButtonBackPrevious" \| "ActionButtonEnd" \| "ActionButtonBeginning" \| "ActionButtonReturn" \| "ActionButtonDocument" \| "ActionButtonSound" \| "ActionButtonMovie" \| "Gear6" \| "Gear9" \| "Funnel" \| "MathPlus" \| "MathMinus" \| "MathMultiply" \| "MathDivide" \| "MathEqual" \| "MathNotEqual" \| "CornerTabs" \| "SquareTabs" \| "PlaqueTabs" \| "ChartX" \| "ChartStar" \| "ChartPlus" | The geometric type of the shape to insert. |
| insertShapeOptions | [Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions) | Optional. The location and size of the geometric shape. The default location and size is (0, 0, 100, 100). |

#### Returns

[Word.Shape](/en-us/javascript/api/word/word.shape)

#### Remarks

[ API set: WordApiDesktop 1.2 ]

### insertHtml(html, insertLocation)

Inserts HTML at the specified location.

```typescript
insertHtml(html: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| html | string | Required. The HTML to be inserted. |
| insertLocation | [Word.InsertLocation](/en-us/javascript/api/word/word.insertlocation) \| "Replace" \| "Start" \| "End" \| "Before" \| "After" | Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('HTML added to the beginning of the range.');
});
```

### insertInlinePictureFromBase64(base64EncodedImage, insertLocation)

Inserts a picture at the specified location.

```typescript
insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.InlinePicture;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| base64EncodedImage | string | Required. The Base64-encoded image to be inserted. |
| insertLocation | [Word.InsertLocation](/en-us/javascript/api/word/word.insertlocation) \| "Replace" \| "Start" \| "End" \| "Before" \| "After" | Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'. |

#### Returns

[Word.InlinePicture](/en-us/javascript/api/word/word.inlinepicture)

#### Remarks

[ API set: WordApi 1.2 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Inserts a picture at the start of the first text box.
  const firstShapeWithTextBox: Word.Shape = context.document.body.shapes
    .getByTypes([Word.ShapeType.textBox])
    .getFirst();
  firstShapeWithTextBox.load("type/body");
  await context.sync();

  const startRange: Word.Range = firstShapeWithTextBox.body.getRange(Word.RangeLocation.start);
  const newPic: Word.InlinePicture = startRange.insertInlinePictureFromBase64(
    getPictureBase64(),
    Word.InsertLocation.start
  );
  newPic.load();
  await context.sync();

  console.log("New inline picture properties:", newPic);
});

...

// Returns Base64-encoded image data for a sample picture.
const pictureBase64 =
"iVBORw0KGgoAAAANSUhEUgAAAOEAAADhCAMAAAAJbSJIAAABblBMVEX+7tEYMFlyg5v8zHXVgof///+hrL77qRnIWmBEWXq6MDgAF0/i1b//8dP+79QKJ1MAIFL8yWpugZz/+O/VzLwzTXR+jaP/z3PHzdjNaWvuxrLFT1n8znmMj5fFTFP25OHlsa2wqqJGW3z7pgCbqsH936oAJlWnssRzdoLTd1HTfINbY3a7tar90IxJVG0AH1ecmJH//90gN14AFU/nxInHVFL80YQAD03qv3LUrm7cwJLWjoLenpPRdXTQgoj15sz+57/7szr93KPbiWjUvZj95LnwzLmMX3L8wmz7rib8xnP8vVz91JT8ukvTz8i8vsORkJKvsLIAD1YwPViWnKZVYHbKuqHjwo3ur2/Pa2O+OTvHVETfj1tybm9qdYlsYlnkmmC0DSPirpvAq4bj5uuono7tu5vgpannnX3ksbSKg5bv0tTclJNFSlyZgpPqwsW4go2giWdbWV+3mmuWgpRcbolURmReS2embHkiRHBcZ6c8AAALcElFTVR4nO3di1cTVx4H8AyThmC484ghFzSxEDRhIRBIMEFQA1qoVhAqYBVd3UXcri1dd7fLdv3vdybJZF73zr2TufPyzPccew49hc6H331nZkylkiRJkiRJkiRJkiRJkiRJkiRJkiRJkiQJ6wj2hH1JLKNo9p/sPB3X8rRUau/f2f56kML2k/n5+XFDSjzPQ7l95+swCqkfzDy1hnwvsLT9FRCF1I7Fpwt5Xt6PfRmF1LgNaBAqZdyNOVGwV9AkVMq4HOshR3iCAJqFalONr1HYRQGtQsXYvrONmjKj7xae0QnVuaO0/OiOlv3lfqI/1G4jgShhnzkIfzA/SNgAUoR9d0I9g/9wfjtsAiHocWZ8fIckLA1ad/SFB0jg+AGxhgNi9FvpU7TwGVHIl+QdtR9GfaTBCOdlIlA18vIzPqZC8kCjZT+mQnI31HInpkKqRqpGDhtADFpInCuGaUe9hBghrY+Xo7+xQgnn6Xth9EuIFNIPpDDsy6cISvg1tVGkkB4Y+ZlCjU34lBrIx6GCitAyyOzQ8mA7+nvfXixCigV33xf9tYwWg3B+/ICnAsbrKFwY8nae0figwnsUq3M34aCXZ3KphPa12+2SWjYZ8v0Pa1Jx4ikRSv1ga2Y8MIzH6aElAqFlRn/vQApRuB32FXoNSRiTad0hgkxI5E8piLlOStgX6DnfkBL7GhKFsS8iUfhN2FfoNWRh3ItIFsa9iBTCmBeRQhjz4ZRGGG8ilfB6jInEVVs/MTj5xUWwbSbUQNs2sZ2Kq9EilNup60qj3LUReT4mR2u2mIXyrtbx2nbjI/P+HpgTFoAYAQlU0rYJYXt3aASg+/zw8HBlkKWFuW5UkSbhsnH4RHxIKmtG8Lx2O5PJ1DhxkKqUW+hGk2gUyoJxhniE6Ivq3W0pAXQPVZ8ibHJ6qrl6JImmGppnecwn3XK7kBnEJOS4zlEUiUZh2zzLI4UQrv94GyPkOnMRJBqFyzghHKa0qfvsQk6KYF90bqUb93pZ72fz5Y+3DT6EsFqOtlC+bh1pXjSUtCq3tWTMsQm5VrSF/L6lkW7k1KsWM7jUjq3CXCFyRPOMb9hpLCtfb7TUvlWsYYUrVqG0Gm2hgbjfG2c61erxCRaYqS2J1o4YvQnDuvJeFtSV9zbfm+7hSTGD9ykpVq3ChagL1d1T/09PWLeOLdZYW2kchKbpfZMgrJ2K8RbyPKGEmRMp5kL40mURYyckFzHTjLkQrpPGmhMx3kIe/kRqp0Ux3kKlihlnY+2EE6MuhIYgiPxL25LbTMysSFEWQvjq8evs3Wu9nL15+4MdCdsvM47IWvG42q9j9c+RE4JXr29ms5pQzVtkHX9S94aG2JrquxVRqlZz7yN2Og5SW6rPJLz2BtkdlbTXN797qeS7zXX7YqdWq2VOTk7monTzBgDgPNsHmoTX3qBO2TRmP9hJpA7lRyESzafUe/c1n0V47S/EARa3YL1dh2He/Q26W2ruq9l6kL059FmFZ7giDoW41Zwq5PmwgClw/lf1+hWaEYcQXntFEMrPpzEpqBuv0EabvjCLikX4liA0n6zazpFhWLdIK8KzW0hgNmsW/sm5mcrbzsLQnjQBXWvj1HPmRshjgdpnAaFNGVhg9pYLofFDOIxQDunzVHAfX0QXwhIeOPw8J6TBBnRx3dAy1jgKzUfjGGEUi3hGKZSBA1D/TC6sngjSVEQHIfxQdMqq9p2hPbgHtvAN9YxCCD/mxwzJ54tF5R/617owtOUpuDGDLeMZSQhLRybg2LTaMi/G8nYhXwpvdQpupO3LtsFwc+YkhHBzzAzUel8RIQzzOQYAUnvnWw9mZlTUayvy7q2zM5QQ8ptlsy9/oQkv8nZhyE+3DW/zAfAtopaPrUJlR/jRUr+xsaI+hBYRwohshQX4mCyEGx+KeatvLF/ThYd5uzC8jmiKAO/esscoVMq3auepmkNdOI0QRuSRKaH0LSJd/TrhehnpUzQZXVhDCGFEHijadVyZwPUjjE/l6N+AGEvD2yVaglxkDoRww8FnLGINNZaGN+ebIqCAg506/9HJZ+iJ06gZPyqDKRLYE9qmdxSxOH1xMV1ErdqULEdAiNsmCDLkV4m+HilvqrNJGIHjbzD76dMsKn+D6+QCIsGREgJwf1HPw59/1r/4+4eRfBETgu7lYlrL4rdq4/yk/YtfRgSahaEuagDozuq+AVAjPhyRFyEhAHuzi0bgJ22IWfQGtAoBMv7zurNpo08R/qoJL70BLUJQL6Pi72226kdOZp5F6AloERZazQlbpqqnPgoV36XNZ26lnoAWIcdxUxWrsMk1/LuBUfXZeL0MgJ8Xf2Eo/E20EyvqHUadgj+9EqTuY3zp9GUP+OuDf4w6TdiF8H3/Dg0TsTK4hao+TIGdEewh2qehoX7+fLn4T49A42nivxqDO1AmKjYgJw2TqzJ6EMWpgH2i4vc2ypiE8J4GNBArtjvfuX6bZQF0LKAWj53QKNxoGAwTlUpF+TOBBHLiCgMhuEHhS3tuowbhsemGvuaUOk0gfeptRl3vQEILZVZCTQj/bb0B3CmSZyElkEEJB0J9lKHKsddWCnCTIPsS9oXw95YboOe7/SgrmH7IoIR94T1XFeQ6k96EYJYOmPY62Q+FJVc+ruPxMRtlmqADMmmkPeFv1gdpHJuo5PmZRUpfOs2ihKrwvUR2aRE7np8epu2EbEZSVfh7jt7XWimseQVSt1FGwrF3tBNhVWotMVh1g0vqRvofJsA8uQ9WG51WQ1wp11k8we+ihGwGmjH0ytPYMnPlgrqEYbQxpO+FaY97+0GwS88h8HiS7UkUPZCJcILYRptsT6HcNFIWwisisMX4MWHq5QwbIRnI/HkTFyMpCyHJx2QjaBG6KKH3AwziMMrlmL9UohukcIrYRpmcVpjiaqDxKqyQp3rWw0ywQvIo48djbQEKKRZrnMTa51boZeGdJ48yXMOHd9eMKLyqTDVFlyEDOebDzIjCqymqy3UfyY+XSNEdAxuFFc4fnpIOe59bIdWAP3o8n4l6F141/QSKvjwB7Ur4vZ8+LgI1/K/PQC4XstB3INfw4wVS9EL/gf50RGrhH/4DlWbq8dMJL0K/B5l+/HifBKXwf4EAlTmf9QafWkixamYSH17lRicMpo1yfmzxKYVBAZWxhnkzpRIGVkI/3qlIJQzMp3RE5ntgGmFQA6ka9u9UpBH+ERzQh9e3gm52BpMh3c2NPZ6FPhy2YZ9pzmYfBN5IfRGe4x9Nz84EPJL69B4whyL2iEF2Q39Wpnv4h+97RNt7gOMmVIZTh3aaDW5N2k9zjb1QqSL+/QLZmYeBApVlmy9HGeD8wU1MsotBDjT+vShafb/ADXT2XNygxSKiL8A+Ep1uwMLqgh890SlBC7ncasDErqt7eVmkVQ70L2sBddc11J8EaeRGWtNKTfVvpAnqmT3gfsJfG6ZbKEujGTunC6tz1tQ93g2G/qUtub/CJS0LR3WQKo/WysWqZE/reG5Uo4qZLNh+aXNlcYQS6B/7VhvS0Vqd/nZZchrHIx0aK7q5dxNThoiDX5r3raF0nKqzHKtEyf1JDgD1d1+m7A8Asrqk47VyR29o3n9nbtd1im/CzMMLR1u/SUdAb/ar5aa7By0QV+HuTBVMXtl8GGGzezraxXXMQ3+96bGOru6bAnNf7D608EUBgNXWKGW0nJ8BsOCtY4or1Ise5f+FKCBa2HtqBUwujWK0LqbBXMfThqVFO56CbgUNtAulwa0uYK2wkHM9WtiOecHkqRcj7UEAqH+ZwkVq5fS0ctzRcPxSNhtzC5yUc5NO03pFABQWRFc/w5jWC7oSpgr4TJoDLB0JdCfdBfH7VSbh0UPbSqnj5XvxK2aXP4P485IkSZIkSZIkSZIkSZIkSZIkSZIk8Tv/B3bBREdOWYS3AAAAAElFTkSuQmCC";
return pictureBase64;
```

### insertOoxml(ooxml, insertLocation)

Inserts OOXML at the specified location.

```typescript
insertOoxml(ooxml: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| ooxml | string | Required. The OOXML to be inserted. |
| insertLocation | [Word.InsertLocation](/en-us/javascript/api/word/word.insertlocation) \| "Replace" \| "Start" \| "End" \| "Before" \| "After" | Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert OOXML in to the beginning of the range.
    range.insertOoxml("<pkg:package xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'><pkg:part pkg:name='/_rels/.rels' pkg:contentType='application/vnd.openxmlformats-package.relationships+xml' pkg:padding='512'><pkg:xmlData><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name='/word/document.xml' pkg:contentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'><pkg:xmlData><w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ><w:body><w:p><w:pPr><w:spacing w:before='360' w:after='0' w:line='480' w:lineRule='auto'/><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val='70AD47' w:themeColor='accent6'/><w:sz w:val='28'/></w:rPr><w:t>This text has formatting directly applied to achieve its font size, color, line spacing, and paragraph spacing.</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>", Word.InsertLocation.start);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('OOXML added to the beginning of the range.');
});

// Read "Create better add-ins for Word with Office Open XML" for guidance on working with OOXML.
// https://learn.microsoft.com/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml
```

### insertParagraph(paragraphText, insertLocation)

Inserts a paragraph at the specified location.

```typescript
insertParagraph(paragraphText: string, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After"): Word.Paragraph;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| paragraphText | string | Required. The paragraph text to be inserted. |
| insertLocation | [before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) \| [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) \| "Before" \| "After" | Required. The value must be 'Before' or 'After'. |

#### Returns

[Word.Paragraph](/en-us/javascript/api/word/word.paragraph)

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert the paragraph after the range.
    range.insertParagraph('Content of a new paragraph', Word.InsertLocation.after);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Paragraph added to the end of the range.');
});
```

### insertPictureFromBase64(base64EncodedImage, insertShapeOptions)

Inserts a floating picture in front of text with its anchor at the beginning of the range.

```typescript
insertPictureFromBase64(base64EncodedImage: string, insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| base64EncodedImage | string | Required. The Base64-encoded image to be inserted. |
| insertShapeOptions | [Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions) | Required. The location and size of the picture. The default location is (0, 0) and the default size is the image's original size. |

#### Returns

[Word.Shape](/en-us/javascript/api/word/word.shape)

#### Remarks

[ API set: WordApiDesktop 1.2 ]

### insertTable(rowCount, columnCount, insertLocation, values)

Inserts a table with the specified number of rows and columns.

```typescript
insertTable(rowCount: number, columnCount: number, insertLocation: Word.InsertLocation.before | Word.InsertLocation.after | "Before" | "After", values?: string[][]): Word.Table;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| rowCount | number | Required. The number of rows in the table. |
| columnCount | number | Required. The number of columns in the table. |
| insertLocation | [before](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-before-member) \| [after](/en-us/javascript/api/word/word.insertlocation#word-word-insertlocation-after-member) \| "Before" \| "After" | Required. The value must be 'Before' or 'After'. |
| values | string[][] | Optional 2D array. Cells are filled if the corresponding strings are specified in the array. |

#### Returns

[Word.Table](/en-us/javascript/api/word/word.table)

#### Remarks

[ API set: WordApi 1.3 ]

### insertText(text, insertLocation)

Inserts text at the specified location.

```typescript
insertText(text: string, insertLocation: Word.InsertLocation | "Replace" | "Start" | "End" | "Before" | "After"): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| text | string | Required. Text to be inserted. |
| insertLocation | [Word.InsertLocation](/en-us/javascript/api/word/word.insertlocation) \| "Replace" \| "Start" \| "End" \| "Before" \| "After" | Required. The value must be 'Replace', 'Start', 'End', 'Before', or 'After'. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert the paragraph at the end of the range.
    range.insertText('New text inserted into the range.', Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Text added to the end of the range.');
});
```

### insertTextBox(text, insertShapeOptions)

Inserts a floating text box in front of text with its anchor at the beginning of the range.

```typescript
insertTextBox(text?: string, insertShapeOptions?: Word.InsertShapeOptions): Word.Shape;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| text | string | Optional. The text to insert into the text box. |
| insertShapeOptions | [Word.InsertShapeOptions](/en-us/javascript/api/word/word.insertshapeoptions) | Optional. The location and size of the text box. The default location and size is (0, 0, 100, 100). |

#### Returns

[Word.Shape](/en-us/javascript/api/word/word.shape)

#### Remarks

[ API set: WordApiDesktop 1.2 ]

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-shapes-text-boxes.yaml

await Word.run(async (context) => {
  // Inserts a text box at the beginning of the selection.
  const range: Word.Range = context.document.getSelection();
  const insertShapeOptions: Word.InsertShapeOptions = {
    top: 0,
    left: 0,
    height: 100,
    width: 100
  };

  const newTextBox: Word.Shape = range.insertTextBox("placeholder text", insertShapeOptions);
  await context.sync();

  console.log("Inserted a text box at the beginning of the current selection.");
});
```

### intersectWith(range)

Returns a new range as the intersection of this range with another range. This range isn't changed. Throws an `ItemNotFound` error if the two ranges aren't overlapped or adjacent.

```typescript
intersectWith(range: Word.Range): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| range | [Word.Range](/en-us/javascript/api/word/word.range) | Required. Another range. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.3 ]

### intersectWithOrNullObject(range)

Returns a new range as the intersection of this range with another range. This range isn't changed. If the two ranges aren't overlapped or adjacent, then this method will return an object with its `isNullObject` property set to `true`. For further information, see [*OrNullObject methods and properties](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
intersectWithOrNullObject(range: Word.Range): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| range | [Word.Range](/en-us/javascript/api/word/word.range) | Required. Another range. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

#### Remarks

[ API set: WordApi 1.3 ]

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.RangeLoadOptions): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| options | [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions) | Provides options for which properties of the object to load. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| propertyNames | string \| string[] | A comma-delimited string or an array of strings that specify the properties to load. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Range;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| propertyNamesAndPaths | { select?: string; expand?: string; } | `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load. |

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

### removeHighlight()

Removes the highlight added by the Highlight function if any.

```typescript
removeHighlight(): void;
```

#### Returns

void

#### Remarks

[ API set: WordApi 1.8 ]

### search(searchText, searchOptions)

Performs a search with the specified SearchOptions on the scope of the range object. The search results are a collection of range objects.

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

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| searchText | string | Required. The search text. |
| searchOptions | [Word.SearchOptions](/en-us/javascript/api/word/word.searchoptions) \| { ignorePunct?: boolean; ignoreSpace?: boolean; matchCase?: boolean; matchPrefix?: boolean; matchSuffix?: boolean; matchWholeWord?: boolean; matchWildcards?: boolean; } | Optional. Options for the search. |

#### Returns

[Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

#### Remarks

[ API set: WordApi 1.1 ]

### select(selectionMode)

Selects and navigates the Word UI to the range.

```typescript
select(selectionMode?: Word.SelectionMode): void;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| selectionMode | [Word.SelectionMode](/en-us/javascript/api/word/word.selectionmode) | Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default. |

#### Returns

void

#### Remarks

[ API set: WordApi 1.1 ]

#### Examples

```TypeScript
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Queue a command to get the current selection and then
    // create a proxy range object with the results.
    const range = context.document.getSelection();

    // Queue a command to insert HTML in to the beginning of the range.
    range.insertHtml('<strong>This is text inserted with range.insertHtml()</strong>', Word.InsertLocation.start);

    // Queue a command to select the HTML that was inserted.
    range.select();

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('Selected the range.');
});
```

### select(selectionMode)

Selects and navigates the Word UI to the range.

```typescript
select(selectionMode?: "Select" | "Start" | "End"): void;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| selectionMode | "Select" \| "Start" \| "End" | Optional. The selection mode must be 'Select', 'Start', or 'End'. 'Select' is the default. |

#### Returns

void

#### Remarks

[ API set: WordApi 1.1 ]

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.RangeUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| properties | [Word.Interfaces.RangeUpdateData](/en-us/javascript/api/word/word.interfaces.rangeupdatedata) | A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called. |
| options | [OfficeExtension.UpdateOptions](/en-us/javascript/api/office/officeextension.updateoptions) | Provides an option to suppress errors if the properties object tries to set any read-only properties. |

#### Returns

void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Range): void;
```

#### Parameters

| Parameter | Type |
|---|---|
| properties | [Word.Range](/en-us/javascript/api/word/word.range) |

#### Returns

void

### split(delimiters, multiParagraphs, trimDelimiters, trimSpacing)

Splits the range into child ranges by using delimiters.

```typescript
split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean): Word.RangeCollection;
```

#### Parameters

| Parameter | Type | Description |
|---|---|---|
| delimiters | string[] | Required. The delimiters as an array of strings. |
| multiParagraphs | boolean | Optional. Indicates whether a returned child range can cover multiple paragraphs. Default is false which indicates that the paragraph boundaries are also used as delimiters. |
| trimDelimiters | boolean | Optional. Indicates whether to trim delimiters from the ranges in the range collection. Default is false which indicates that the delimiters are included in the ranges returned in the range collection. |
| trimSpacing | boolean | Optional. Indicates whether to trim spacing characters (spaces, tabs, column breaks, and paragraph end marks) from the start and end of the ranges returned in the range collection. Default is false which indicates that spacing characters at the start and end of the ranges are included in the range collection. |

#### Returns

[Word.RangeCollection](/en-us/javascript/api/word/word.rangecollection)

#### Remarks

[ API set: WordApi 1.3 ]

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Range` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.RangeData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.RangeData;
```

#### Returns

[Word.Interfaces.RangeData](/en-us/javascript/api/word/word.interfaces.rangedata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Range;```

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.Range;
```

#### Returns

[Word.Range](/en-us/javascript/api/word/word.range)

## Event Details

### onCommentAdded

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when new comments are added.

```typescript
readonly onCommentAdded: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

#### Event Type

[OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### onCommentChanged

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment or its reply is changed.

```typescript
readonly onCommentChanged: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

#### Event Type

[OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### onCommentDeselected

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment is deselected.

```typescript
readonly onCommentDeselected: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

#### Event Type

[OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

### onCommentSelected

> **Note**: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Occurs when a comment is selected.

```typescript
readonly onCommentSelected: OfficeExtension.EventHandlers<Word.CommentEventArgs>;
```

#### Event Type

[OfficeExtension.EventHandlers](/en-us/javascript/api/office/officeextension.eventhandlers)<[Word.CommentEventArgs](/en-us/javascript/api/word/word.commenteventargs)>

#### Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## 成功的目录(TOC)插入示例

以下是经过验证的分步骤TOC插入代码，可以避免常见错误：

### 步骤1: 插入目录标题
```typescript
await Word.run(async (context) => {
    const docStart = context.document.body.getRange("Start");
    const titleRange = docStart.insertText("目录\n\n", "Before");
    titleRange.font.size = 16;
    titleRange.font.bold = true;
    await context.sync();
});
```

### 步骤2: 插入TOC字段
```typescript
await Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    
    let insertPosition = null;
    for (let i = 0; i < paragraphs.items.length; i++) {
        if (paragraphs.items[i].text.includes("目录")) {
            if (i + 1 < paragraphs.items.length) {
                insertPosition = paragraphs.items[i + 1].getRange("Start");
            } else {
                insertPosition = paragraphs.items[i].getRange("End");
            }
            break;
        }
    }
    
    if (insertPosition) {
        const tocField = insertPosition.insertField("Before", "TOC", "\\o \"1-3\" \\h \\z \\u", false);
        await context.sync();
    }
});
```

### 步骤3: 添加分页符并更新TOC
```typescript
await Word.run(async (context) => {
    const fields = context.document.body.fields;
    fields.load("type");
    await context.sync();
    
    let tocField = null;
    for (let i = 0; i < fields.items.length; i++) {
        if (fields.items[i].type === "TOC") {
            tocField = fields.items[i];
            break;
        }
    }
    
    if (tocField) {
        tocField.load("result");
        await context.sync();
        tocField.result.insertBreak("Page", "After");
        tocField.updateResult();
        await context.sync();
    }
});
```

### 重要说明：
- **必须分步执行**：不能在一个Word.run中完成所有操作，会导致InvalidArgument错误
- **TOC字段参数**：`"\\o \"1-3\" \\h \\z \\u"` 表示显示1-3级标题，包含超链接和页码
- **避免重复标题**：不要在TOC字段中再次指定标题，会导致重复的"目录"文字
- **必须调用updateResult()**：确保TOC字段显示最新的文档结构