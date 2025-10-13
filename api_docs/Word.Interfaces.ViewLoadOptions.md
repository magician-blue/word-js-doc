# Word.Interfaces.ViewLoadOptions interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Contains the view attributes (such as show all, field shading, and table gridlines) for a window or pane.

## Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
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

## Property Details

### $all
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property value: boolean

---

### areAllNonprintingCharactersDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether all nonprinting characters are displayed.

```typescript
areAllNonprintingCharactersDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areBackgroundsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether background colors and images are shown when the document is displayed in print layout view.

```typescript
areBackgroundsDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areBookmarksIndicated
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether square brackets are displayed at the beginning and end of each bookmark.

```typescript
areBookmarksIndicated?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areCommentsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays the comments in the document.

```typescript
areCommentsDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areConnectingLinesToRevisionsBalloonDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays connecting lines from the text to the revision and comment balloons.

```typescript
areConnectingLinesToRevisionsBalloonDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

---

### areCropMarksDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether crop marks are shown in the corners of pages to indicate where margins are located.

```typescript
areCropMarksDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areDrawingsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether objects created with the drawing tools are displayed in print layout view.

```typescript
areDrawingsDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areEditableRangesShaded
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether shading is applied to the ranges in the document that users have permission to modify.

```typescript
areEditableRangesShaded?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areFieldCodesDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether field codes are displayed.

```typescript
areFieldCodesDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areFormatChangesDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays formatting changes made to the document with Track Changes enabled.

```typescript
areFormatChangesDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areInkAnnotationsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether handwritten ink annotations are shown or hidden.

```typescript
areInkAnnotationsDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areInsertionsAndDeletionsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays insertions and deletions made to the document with Track Changes enabled.

```typescript
areInsertionsAndDeletionsDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areLinesWrappedToWindow
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.

```typescript
areLinesWrappedToWindow?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areObjectAnchorsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether object anchors are displayed next to items that can be positioned in print layout view.

```typescript
areObjectAnchorsDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areOptionalBreaksDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether Microsoft Word displays optional line breaks.

```typescript
areOptionalBreaksDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areOptionalHyphensDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether optional hyphens are displayed.

```typescript
areOptionalHyphensDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areOtherAuthorsVisible
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether other authors' presence should be visible in the document.

```typescript
areOtherAuthorsVisible?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### arePageBoundariesDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the top and bottom margins and the gray area between pages in the document are displayed.

```typescript
arePageBoundariesDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areParagraphsMarksDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether paragraph marks are displayed.

```typescript
areParagraphsMarksDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### arePicturePlaceholdersDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether blank boxes are displayed as placeholders for pictures.

```typescript
arePicturePlaceholdersDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areRevisionsAndCommentsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays revisions and comments made to the document with Track Changes enabled.

```typescript
areRevisionsAndCommentsDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

---

### areSpacesIndicated
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether space characters are displayed.

```typescript
areSpacesIndicated?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areTableGridlinesDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether table gridlines are displayed.

```typescript
areTableGridlinesDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areTabsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether tab characters are displayed.

```typescript
areTabsDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areTextBoundariesDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether dotted lines are displayed around page margins, text columns, objects, and frames in print layout view.

```typescript
areTextBoundariesDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### columnWidth
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the column width in Reading mode.

```typescript
columnWidth?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### fieldShading
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets on-screen shading for fields.

```typescript
fieldShading?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isDraft
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.

```typescript
isDraft?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isFirstLineOnlyDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether only the first line of body text is shown in outline view.

```typescript
isFirstLineOnlyDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isFormatDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether character formatting is visible in outline view.

```typescript
isFormatDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isFullScreen
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is in full-screen view.

```typescript
isFullScreen?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isHiddenTextDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether text formatted as hidden text is displayed.

```typescript
isHiddenTextDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isHighlightingDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether highlight formatting is displayed and printed with the document.

```typescript
isHighlightingDisplayed?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isInConflictMode
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document is in conflict mode view.

```typescript
isInConflictMode?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isInPanning
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word is in Panning mode.

```typescript
isInPanning?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isInReadingLayout
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document is being viewed in reading layout view.

```typescript
isInReadingLayout?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isMailMergeDataView
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether mail merge data is displayed instead of mail merge fields.

```typescript
isMailMergeDataView?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isMainTextLayerVisible
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the text in the document is visible when the header and footer areas are displayed.

```typescript
isMainTextLayerVisible?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isPointerShownAsMagnifier
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the pointer is displayed as a magnifying glass in print preview.

```typescript
isPointerShownAsMagnifier?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isReadingLayoutActualView
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether pages displayed in reading layout view are displayed using the same layout as printed pages.

```typescript
isReadingLayoutActualView?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isXmlMarkupVisible
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether XML tags are visible in the document.

```typescript
isXmlMarkupVisible?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### markupMode
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the display mode for tracked changes.

```typescript
markupMode?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pageColor
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page color in Reading mode.

```typescript
pageColor?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pageMovementType
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page movement type.

```typescript
pageMovementType?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### readingLayoutTruncateMargins
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether margins are visible or hidden when the document is viewed in Full Screen Reading view.

```typescript
readingLayoutTruncateMargins?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### revisionsBalloonSide
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether Word displays revision balloons in the left or right margin in the document.

```typescript
revisionsBalloonSide?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

---

### revisionsBalloonWidth
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the revision balloons.

```typescript
revisionsBalloonWidth?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

---

### revisionsBalloonWidthType
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies how Microsoft Word measures the width of revision balloons.

```typescript
revisionsBalloonWidthType?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

---

### revisionsFilter
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the instance of a RevisionsFilter object.

```typescript
revisionsFilter?: Word.Interfaces.RevisionsFilterLoadOptions;
```

Property value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.revisionsfilterloadoptions

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

---

### seekView
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the document element displayed in print layout view.

```typescript
seekView?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### splitSpecial
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the active window pane.

```typescript
splitSpecial?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the view type.

```typescript
type?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)