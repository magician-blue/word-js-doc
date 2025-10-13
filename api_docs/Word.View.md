# Word.View class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Contains the view attributes (such as show all, field shading, and table gridlines) for a window or pane.

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

## Properties

- areAllNonprintingCharactersDisplayed — Specifies whether all nonprinting characters are displayed.
- areBackgroundsDisplayed — Gets whether background colors and images are shown when the document is displayed in print layout view.
- areBookmarksIndicated — Gets whether square brackets are displayed at the beginning and end of each bookmark.
- areCommentsDisplayed — Specifies whether Microsoft Word displays the comments in the document.
- areConnectingLinesToRevisionsBalloonDisplayed — Specifies whether Microsoft Word displays connecting lines from the text to the revision and comment balloons.
- areCropMarksDisplayed — Gets whether crop marks are shown in the corners of pages to indicate where margins are located.
- areDrawingsDisplayed — Gets whether objects created with the drawing tools are displayed in print layout view.
- areEditableRangesShaded — Specifies whether shading is applied to the ranges in the document that users have permission to modify.
- areFieldCodesDisplayed — Specifies whether field codes are displayed.
- areFormatChangesDisplayed — Specifies whether Microsoft Word displays formatting changes made to the document with Track Changes enabled.
- areInkAnnotationsDisplayed — Specifies whether handwritten ink annotations are shown or hidden.
- areInsertionsAndDeletionsDisplayed — Specifies whether Microsoft Word displays insertions and deletions made to the document with Track Changes enabled.
- areLinesWrappedToWindow — Gets whether lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.
- areObjectAnchorsDisplayed — Gets whether object anchors are displayed next to items that can be positioned in print layout view.
- areOptionalBreaksDisplayed — Gets whether Microsoft Word displays optional line breaks.
- areOptionalHyphensDisplayed — Gets whether optional hyphens are displayed.
- areOtherAuthorsVisible — Gets whether other authors' presence should be visible in the document.
- arePageBoundariesDisplayed — Gets whether the top and bottom margins and the gray area between pages in the document are displayed.
- areParagraphsMarksDisplayed — Gets whether paragraph marks are displayed.
- arePicturePlaceholdersDisplayed — Gets whether blank boxes are displayed as placeholders for pictures.
- areRevisionsAndCommentsDisplayed — Specifies whether Microsoft Word displays revisions and comments made to the document with Track Changes enabled.
- areSpacesIndicated — Gets whether space characters are displayed.
- areTableGridlinesDisplayed — Specifies whether table gridlines are displayed.
- areTabsDisplayed — Gets whether tab characters are displayed.
- areTextBoundariesDisplayed — Gets whether dotted lines are displayed around page margins, text columns, objects, and frames in print layout view.
- columnWidth — Specifies the column width in Reading mode.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- fieldShading — Gets on-screen shading for fields.
- isDraft — Specifies whether all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.
- isFirstLineOnlyDisplayed — Specifies whether only the first line of body text is shown in outline view.
- isFormatDisplayed — Specifies whether character formatting is visible in outline view.
- isFullScreen — Specifies whether the window is in full-screen view.
- isHiddenTextDisplayed — Gets whether text formatted as hidden text is displayed.
- isHighlightingDisplayed — Gets whether highlight formatting is displayed and printed with the document.
- isInConflictMode — Specifies whether the document is in conflict mode view.
- isInPanning — Specifies whether Microsoft Word is in Panning mode.
- isInReadingLayout — Specifies whether the document is being viewed in reading layout view.
- isMailMergeDataView — Specifies whether mail merge data is displayed instead of mail merge fields.
- isMainTextLayerVisible — Specifies whether the text in the document is visible when the header and footer areas are displayed.
- isPointerShownAsMagnifier — Specifies whether the pointer is displayed as a magnifying glass in print preview.
- isReadingLayoutActualView — Specifies whether pages displayed in reading layout view are displayed using the same layout as printed pages.
- isXmlMarkupVisible — Specifies whether XML tags are visible in the document.
- markupMode — Specifies the display mode for tracked changes.
- pageColor — Specifies the page color in Reading mode.
- pageMovementType — Specifies the page movement type.
- readingLayoutTruncateMargins — Specifies whether margins are visible or hidden when the document is viewed in Full Screen Reading view.
- revisionsBalloonSide — Gets whether Word displays revision balloons in the left or right margin in the document.
- revisionsBalloonWidth — Specifies the width of the revision balloons.
- revisionsBalloonWidthType — Specifies how Microsoft Word measures the width of revision balloons.
- revisionsFilter — Gets the instance of a RevisionsFilter object.
- seekView — Specifies the document element displayed in print layout view.
- splitSpecial — Specifies the active window pane.
- type — Specifies the view type.

## Methods

- collapseAllHeadings() — Collapses all the headings in the document.
- collapseOutline(range) — Collapses the text under the selection or the specified range by one heading level.
- expandAllHeadings() — Expands all the headings in the document.
- expandOutline(range) — Expands the text under the selection by one heading level.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- nextHeaderFooter() — Moves to the next header or footer, depending on whether a header or footer is displayed in the view.
- previousHeaderFooter() — Moves to the previous header or footer, depending on whether a header or footer is displayed in the view.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- showAllHeadings() — Switches between showing all text (headings and body text) and showing only headings.
- showHeading(level) — Shows all headings up to the specified heading level and hides subordinate headings and body text.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.View object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ViewData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### areAllNonprintingCharactersDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether all nonprinting characters are displayed.

```typescript
areAllNonprintingCharactersDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areBackgroundsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether background colors and images are shown when the document is displayed in print layout view.

```typescript
areBackgroundsDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areBookmarksIndicated

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether square brackets are displayed at the beginning and end of each bookmark.

```typescript
readonly areBookmarksIndicated: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areCommentsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays the comments in the document.

```typescript
areCommentsDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areConnectingLinesToRevisionsBalloonDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays connecting lines from the text to the revision and comment balloons.

```typescript
areConnectingLinesToRevisionsBalloonDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview

### areCropMarksDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether crop marks are shown in the corners of pages to indicate where margins are located.

```typescript
readonly areCropMarksDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areDrawingsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether objects created with the drawing tools are displayed in print layout view.

```typescript
readonly areDrawingsDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areEditableRangesShaded

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether shading is applied to the ranges in the document that users have permission to modify.

```typescript
areEditableRangesShaded: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areFieldCodesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether field codes are displayed.

```typescript
areFieldCodesDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areFormatChangesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays formatting changes made to the document with Track Changes enabled.

```typescript
areFormatChangesDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areInkAnnotationsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether handwritten ink annotations are shown or hidden.

```typescript
areInkAnnotationsDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areInsertionsAndDeletionsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays insertions and deletions made to the document with Track Changes enabled.

```typescript
areInsertionsAndDeletionsDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areLinesWrappedToWindow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.

```typescript
readonly areLinesWrappedToWindow: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areObjectAnchorsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether object anchors are displayed next to items that can be positioned in print layout view.

```typescript
readonly areObjectAnchorsDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areOptionalBreaksDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether Microsoft Word displays optional line breaks.

```typescript
readonly areOptionalBreaksDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areOptionalHyphensDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether optional hyphens are displayed.

```typescript
readonly areOptionalHyphensDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areOtherAuthorsVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether other authors' presence should be visible in the document.

```typescript
areOtherAuthorsVisible: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### arePageBoundariesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the top and bottom margins and the gray area between pages in the document are displayed.

```typescript
readonly arePageBoundariesDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areParagraphsMarksDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether paragraph marks are displayed.

```typescript
readonly areParagraphsMarksDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### arePicturePlaceholdersDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether blank boxes are displayed as placeholders for pictures.

```typescript
readonly arePicturePlaceholdersDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areRevisionsAndCommentsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays revisions and comments made to the document with Track Changes enabled.

```typescript
areRevisionsAndCommentsDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview

### areSpacesIndicated

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether space characters are displayed.

```typescript
readonly areSpacesIndicated: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areTableGridlinesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether table gridlines are displayed.

```typescript
areTableGridlinesDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areTabsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether tab characters are displayed.

```typescript
readonly areTabsDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### areTextBoundariesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether dotted lines are displayed around page margins, text columns, objects, and frames in print layout view.

```typescript
readonly areTextBoundariesDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### columnWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the column width in Reading mode.

```typescript
columnWidth: Word.ColumnWidth | "Narrow" | "Default" | "Wide";
```

#### Property Value
Word.ColumnWidth | "Narrow" | "Default" | "Wide"
- https://learn.microsoft.com/en-us/javascript/api/word/word.columnwidth

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### context

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

#### Property Value
Word.RequestContext
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### fieldShading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets on-screen shading for fields.

```typescript
readonly fieldShading: Word.FieldShading | "Never" | "Always" | "WhenSelected";
```

#### Property Value
Word.FieldShading | "Never" | "Always" | "WhenSelected"
- https://learn.microsoft.com/en-us/javascript/api/word/word.fieldshading

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isDraft

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.

```typescript
isDraft: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isFirstLineOnlyDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether only the first line of body text is shown in outline view.

```typescript
isFirstLineOnlyDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isFormatDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether character formatting is visible in outline view.

```typescript
isFormatDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isFullScreen

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is in full-screen view.

```typescript
isFullScreen: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isHiddenTextDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether text formatted as hidden text is displayed.

```typescript
readonly isHiddenTextDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isHighlightingDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether highlight formatting is displayed and printed with the document.

```typescript
readonly isHighlightingDisplayed: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isInConflictMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document is in conflict mode view.

```typescript
isInConflictMode: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isInPanning

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word is in Panning mode.

```typescript
isInPanning: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isInReadingLayout

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document is being viewed in reading layout view.

```typescript
isInReadingLayout: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isMailMergeDataView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether mail merge data is displayed instead of mail merge fields.

```typescript
isMailMergeDataView: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isMainTextLayerVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the text in the document is visible when the header and footer areas are displayed.

```typescript
isMainTextLayerVisible: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isPointerShownAsMagnifier

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the pointer is displayed as a magnifying glass in print preview.

```typescript
isPointerShownAsMagnifier: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isReadingLayoutActualView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether pages displayed in reading layout view are displayed using the same layout as printed pages.

```typescript
isReadingLayoutActualView: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### isXmlMarkupVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether XML tags are visible in the document.

```typescript
isXmlMarkupVisible: boolean;
```

#### Property Value
boolean

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### markupMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the display mode for tracked changes.

```typescript
markupMode: Word.RevisionsMode | "Balloon" | "Inline" | "Mixed";
```

#### Property Value
Word.RevisionsMode | "Balloon" | "Inline" | "Mixed"
- https://learn.microsoft.com/en-us/javascript/api/word/word.revisionsmode

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### pageColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page color in Reading mode.

```typescript
pageColor: Word.PageColor | "None" | "Sepia" | "Inverse";
```

#### Property Value
Word.PageColor | "None" | "Sepia" | "Inverse"
- https://learn.microsoft.com/en-us/javascript/api/word/word.pagecolor

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### pageMovementType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page movement type.

```typescript
pageMovementType: Word.PageMovementType | "Vertical" | "SideToSide";
```

#### Property Value
Word.PageMovementType | "Vertical" | "SideToSide"
- https://learn.microsoft.com/en-us/javascript/api/word/word.pagemovementtype

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### readingLayoutTruncateMargins

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether margins are visible or hidden when the document is viewed in Full Screen Reading view.

```typescript
readingLayoutTruncateMargins: Word.ReadingLayoutMargin | "Automatic" | "Suppress" | "Full";
```

#### Property Value
Word.ReadingLayoutMargin | "Automatic" | "Suppress" | "Full"
- https://learn.microsoft.com/en-us/javascript/api/word/word.readinglayoutmargin

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### revisionsBalloonSide

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether Word displays revision balloons in the left or right margin in the document.

```typescript
readonly revisionsBalloonSide: Word.RevisionsBalloonMargin | "Left" | "Right";
```

#### Property Value
Word.RevisionsBalloonMargin | "Left" | "Right"
- https://learn.microsoft.com/en-us/javascript/api/word/word.revisionsballoonmargin

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview

### revisionsBalloonWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the revision balloons.

```typescript
revisionsBalloonWidth: number;
```

#### Property Value
number

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview

### revisionsBalloonWidthType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies how Microsoft Word measures the width of revision balloons.

```typescript
revisionsBalloonWidthType: Word.RevisionsBalloonWidthType | "Percent" | "Points";
```

#### Property Value
Word.RevisionsBalloonWidthType | "Percent" | "Points"
- https://learn.microsoft.com/en-us/javascript/api/word/word.revisionsballoonwidthtype

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview

### revisionsFilter

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the instance of a RevisionsFilter object.

```typescript
readonly revisionsFilter: Word.RevisionsFilter;
```

#### Property Value
Word.RevisionsFilter
- https://learn.microsoft.com/en-us/javascript/api/word/word.revisionsfilter

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview

### seekView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the document element displayed in print layout view.

```typescript
seekView: Word.SeekView | "MainDocument" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "CurrentPageHeader" | "CurrentPageFooter";
```

#### Property Value
Word.SeekView | "MainDocument" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "CurrentPageHeader" | "CurrentPageFooter"
- https://learn.microsoft.com/en-us/javascript/api/word/word.seekview

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### splitSpecial

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the active window pane.

```typescript
splitSpecial: Word.SpecialPane | "None" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "FootnoteContinuationNotice" | "FootnoteContinuationSeparator" | "FootnoteSeparator" | "EndnoteContinuationNotice" | "EndnoteContinuationSeparator" | "EndnoteSeparator" | "Comments" | "CurrentPageHeader" | "CurrentPageFooter" | "Revisions" | "RevisionsHoriz" | "RevisionsVert";
```

#### Property Value
Word.SpecialPane | "None" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "FootnoteContinuationNotice" | "FootnoteContinuationSeparator" | "FootnoteSeparator" | "EndnoteContinuationNotice" | "EndnoteContinuationSeparator" | "EndnoteSeparator" | "Comments" | "CurrentPageHeader" | "CurrentPageFooter" | "Revisions" | "RevisionsHoriz" | "RevisionsVert"
- https://learn.microsoft.com/en-us/javascript/api/word/word.specialpane

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the view type.

```typescript
type: Word.ViewType | "Normal" | "Outline" | "Print" | "PrintPreview" | "Master" | "Web" | "Reading" | "Conflict";
```

#### Property Value
Word.ViewType | "Normal" | "Outline" | "Print" | "PrintPreview" | "Master" | "Web" | "Reading" | "Conflict"
- https://learn.microsoft.com/en-us/javascript/api/word/word.viewtype

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

## Method Details

### collapseAllHeadings()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Collapses all the headings in the document.

```typescript
collapseAllHeadings(): void;
```

#### Returns
void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### collapseOutline(range)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Collapses the text under the selection or the specified range by one heading level.

```typescript
collapseOutline(range: Word.Range): void;
```

Parameters
- range: Word.Range
  - https://learn.microsoft.com/en-us/javascript/api/word/word.range
  - A Range object that specifies the range to collapse.

#### Returns
void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### expandAllHeadings()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Expands all the headings in the document.

```typescript
expandAllHeadings(): void;
```

#### Returns
void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### expandOutline(range)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Expands the text under the selection by one heading level.

```typescript
expandOutline(range: Word.Range): void;
```

Parameters
- range: Word.Range
  - https://learn.microsoft.com/en-us/javascript/api/word/word.range
  - A Range object that specifies the range to expand.

#### Returns
void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### load(options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.ViewLoadOptions): Word.View;
```

Parameters
- options: Word.Interfaces.ViewLoadOptions
  - https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.viewloadoptions
  - Provides options for which properties of the object to load.

#### Returns
Word.View
- https://learn.microsoft.com/en-us/javascript/api/word/word.view

### load(propertyNames)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.View;
```

Parameters
- propertyNames: string | string[]
  - A comma-delimited string or an array of strings that specify the properties to load.

#### Returns
Word.View
- https://learn.microsoft.com/en-us/javascript/api/word/word.view

### load(propertyNamesAndPaths)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.View;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }
  - propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

#### Returns
Word.View
- https://learn.microsoft.com/en-us/javascript/api/word/word.view

### nextHeaderFooter()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Moves to the next header or footer, depending on whether a header or footer is displayed in the view.

```typescript
nextHeaderFooter(): void;
```

#### Returns
void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### previousHeaderFooter()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Moves to the previous header or footer, depending on whether a header or footer is displayed in the view.

```typescript
previousHeaderFooter(): void;
```

#### Returns
void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### set(properties, options)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.ViewUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: Word.Interfaces.ViewUpdateData
  - https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.viewupdatedata
  - A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: OfficeExtension.UpdateOptions
  - https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions
  - Provides an option to suppress errors if the properties object tries to set any read-only properties.

#### Returns
void

### set(properties)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.View): void;
```

Parameters
- properties: Word.View
  - https://learn.microsoft.com/en-us/javascript/api/word/word.view

#### Returns
void

### showAllHeadings()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Switches between showing all text (headings and body text) and showing only headings.

```typescript
showAllHeadings(): void;
```

#### Returns
void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### showHeading(level)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Shows all headings up to the specified heading level and hides subordinate headings and body text.

```typescript
showHeading(level: number): void;
```

Parameters
- level: number
  - The heading level to show.

#### Returns
void

Remarks
- API set: WordApi BETA (PREVIEW ONLY)
  https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### toJSON()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.View object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ViewData) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.ViewData;
```

#### Returns
Word.Interfaces.ViewData
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.viewdata

### track()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.View;
```

#### Returns
Word.View
- https://learn.microsoft.com/en-us/javascript/api/word/word.view

### untrack()

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.View;
```

#### Returns
Word.View
- https://learn.microsoft.com/en-us/javascript/api/word/word.view