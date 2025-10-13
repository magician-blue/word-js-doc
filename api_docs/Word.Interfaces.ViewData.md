# Word.Interfaces.ViewData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface describing the data returned by calling view.toJSON().

## Properties

- [areAllNonprintingCharactersDisplayed](#areallnonprintingcharactersdisplayed): Specifies whether all nonprinting characters are displayed.
- [areBackgroundsDisplayed](#arebackgroundsdisplayed): Gets whether background colors and images are shown when the document is displayed in print layout view.
- [areBookmarksIndicated](#arebookmarksindicated): Gets whether square brackets are displayed at the beginning and end of each bookmark.
- [areCommentsDisplayed](#arecommentsdisplayed): Specifies whether Microsoft Word displays the comments in the document.
- [areConnectingLinesToRevisionsBalloonDisplayed](#areconnectinglinestorevisionsballoondisplayed): Specifies whether Microsoft Word displays connecting lines from the text to the revision and comment balloons.
- [areCropMarksDisplayed](#arecropmarksdisplayed): Gets whether crop marks are shown in the corners of pages to indicate where margins are located.
- [areDrawingsDisplayed](#aredrawingsdisplayed): Gets whether objects created with the drawing tools are displayed in print layout view.
- [areEditableRangesShaded](#areeditablerangesshaded): Specifies whether shading is applied to the ranges in the document that users have permission to modify.
- [areFieldCodesDisplayed](#arefieldcodesdisplayed): Specifies whether field codes are displayed.
- [areFormatChangesDisplayed](#areformatchangesdisplayed): Specifies whether Microsoft Word displays formatting changes made to the document with Track Changes enabled.
- [areInkAnnotationsDisplayed](#areinkannotationsdisplayed): Specifies whether handwritten ink annotations are shown or hidden.
- [areInsertionsAndDeletionsDisplayed](#areinsertionsanddeletionsdisplayed): Specifies whether Microsoft Word displays insertions and deletions made to the document with Track Changes enabled.
- [areLinesWrappedToWindow](#arelineswrappedtowindow): Gets whether lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.
- [areObjectAnchorsDisplayed](#areobjectanchorsdisplayed): Gets whether object anchors are displayed next to items that can be positioned in print layout view.
- [areOptionalBreaksDisplayed](#areoptionalbreaksdisplayed): Gets whether Microsoft Word displays optional line breaks.
- [areOptionalHyphensDisplayed](#areoptionalhyphensdisplayed): Gets whether optional hyphens are displayed.
- [areOtherAuthorsVisible](#areotherauthorsvisible): Gets whether other authors' presence should be visible in the document.
- [arePageBoundariesDisplayed](#arepageboundariesdisplayed): Gets whether the top and bottom margins and the gray area between pages in the document are displayed.
- [areParagraphsMarksDisplayed](#areparagraphsmarksdisplayed): Gets whether paragraph marks are displayed.
- [arePicturePlaceholdersDisplayed](#arepictureplaceholdersdisplayed): Gets whether blank boxes are displayed as placeholders for pictures.
- [areRevisionsAndCommentsDisplayed](#arerevisionsandcommentsdisplayed): Specifies whether Microsoft Word displays revisions and comments made to the document with Track Changes enabled.
- [areSpacesIndicated](#arespacesindicated): Gets whether space characters are displayed.
- [areTableGridlinesDisplayed](#aretablegridlinesdisplayed): Specifies whether table gridlines are displayed.
- [areTabsDisplayed](#aretabsdisplayed): Gets whether tab characters are displayed.
- [areTextBoundariesDisplayed](#aretextboundariesdisplayed): Gets whether dotted lines are displayed around page margins, text columns, objects, and frames in print layout view.
- [columnWidth](#columnwidth): Specifies the column width in Reading mode.
- [fieldShading](#fieldshading): Gets on-screen shading for fields.
- [isDraft](#isdraft): Specifies whether all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.
- [isFirstLineOnlyDisplayed](#isfirstlineonlydisplayed): Specifies whether only the first line of body text is shown in outline view.
- [isFormatDisplayed](#isformatdisplayed): Specifies whether character formatting is visible in outline view.
- [isFullScreen](#isfullscreen): Specifies whether the window is in full-screen view.
- [isHiddenTextDisplayed](#ishiddentextdisplayed): Gets whether text formatted as hidden text is displayed.
- [isHighlightingDisplayed](#ishighlightingdisplayed): Gets whether highlight formatting is displayed and printed with the document.
- [isInConflictMode](#isinconflictmode): Specifies whether the document is in conflict mode view.
- [isInPanning](#isinpanning): Specifies whether Microsoft Word is in Panning mode.
- [isInReadingLayout](#isinreadinglayout): Specifies whether the document is being viewed in reading layout view.
- [isMailMergeDataView](#ismailmergedataview): Specifies whether mail merge data is displayed instead of mail merge fields.
- [isMainTextLayerVisible](#ismaintextlayervisible): Specifies whether the text in the document is visible when the header and footer areas are displayed.
- [isPointerShownAsMagnifier](#ispointershownasmagnifier): Specifies whether the pointer is displayed as a magnifying glass in print preview.
- [isReadingLayoutActualView](#isreadinglayoutactualview): Specifies whether pages displayed in reading layout view are displayed using the same layout as printed pages.
- [isXmlMarkupVisible](#isxmlmarkupvisible): Specifies whether XML tags are visible in the document.
- [markupMode](#markupmode): Specifies the display mode for tracked changes.
- [pageColor](#pagecolor): Specifies the page color in Reading mode.
- [pageMovementType](#pagemovementtype): Specifies the page movement type.
- [readingLayoutTruncateMargins](#readinglayouttruncatemargins): Specifies whether margins are visible or hidden when the document is viewed in Full Screen Reading view.
- [revisionsBalloonSide](#revisionsballoonside): Gets whether Word displays revision balloons in the left or right margin in the document.
- [revisionsBalloonWidth](#revisionsballoonwidth): Specifies the width of the revision balloons.
- [revisionsBalloonWidthType](#revisionsballoonwidthtype): Specifies how Microsoft Word measures the width of revision balloons.
- [seekView](#seekview): Specifies the document element displayed in print layout view.
- [splitSpecial](#splitspecial): Specifies the active window pane.
- [type](#type): Specifies the view type.

## Property Details

### areAllNonprintingCharactersDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether all nonprinting characters are displayed.

```typescript
areAllNonprintingCharactersDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areBackgroundsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether background colors and images are shown when the document is displayed in print layout view.

```typescript
areBackgroundsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areBookmarksIndicated

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether square brackets are displayed at the beginning and end of each bookmark.

```typescript
areBookmarksIndicated?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areCommentsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays the comments in the document.

```typescript
areCommentsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areConnectingLinesToRevisionsBalloonDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays connecting lines from the text to the revision and comment balloons.

```typescript
areConnectingLinesToRevisionsBalloonDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

### areCropMarksDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether crop marks are shown in the corners of pages to indicate where margins are located.

```typescript
areCropMarksDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areDrawingsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether objects created with the drawing tools are displayed in print layout view.

```typescript
areDrawingsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areEditableRangesShaded

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether shading is applied to the ranges in the document that users have permission to modify.

```typescript
areEditableRangesShaded?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areFieldCodesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether field codes are displayed.

```typescript
areFieldCodesDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areFormatChangesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays formatting changes made to the document with Track Changes enabled.

```typescript
areFormatChangesDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areInkAnnotationsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether handwritten ink annotations are shown or hidden.

```typescript
areInkAnnotationsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areInsertionsAndDeletionsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays insertions and deletions made to the document with Track Changes enabled.

```typescript
areInsertionsAndDeletionsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areLinesWrappedToWindow

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.

```typescript
areLinesWrappedToWindow?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areObjectAnchorsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether object anchors are displayed next to items that can be positioned in print layout view.

```typescript
areObjectAnchorsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areOptionalBreaksDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether Microsoft Word displays optional line breaks.

```typescript
areOptionalBreaksDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areOptionalHyphensDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether optional hyphens are displayed.

```typescript
areOptionalHyphensDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areOtherAuthorsVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether other authors' presence should be visible in the document.

```typescript
areOtherAuthorsVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### arePageBoundariesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the top and bottom margins and the gray area between pages in the document are displayed.

```typescript
arePageBoundariesDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areParagraphsMarksDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether paragraph marks are displayed.

```typescript
areParagraphsMarksDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### arePicturePlaceholdersDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether blank boxes are displayed as placeholders for pictures.

```typescript
arePicturePlaceholdersDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areRevisionsAndCommentsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word displays revisions and comments made to the document with Track Changes enabled.

```typescript
areRevisionsAndCommentsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

### areSpacesIndicated

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether space characters are displayed.

```typescript
areSpacesIndicated?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areTableGridlinesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether table gridlines are displayed.

```typescript
areTableGridlinesDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areTabsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether tab characters are displayed.

```typescript
areTabsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areTextBoundariesDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether dotted lines are displayed around page margins, text columns, objects, and frames in print layout view.

```typescript
areTextBoundariesDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### columnWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the column width in Reading mode.

```typescript
columnWidth?: Word.ColumnWidth | "Narrow" | "Default" | "Wide";
```

Property Value: [Word.ColumnWidth](https://learn.microsoft.com/en-us/javascript/api/word/word.columnwidth) | "Narrow" | "Default" | "Wide"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### fieldShading

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets on-screen shading for fields.

```typescript
fieldShading?: Word.FieldShading | "Never" | "Always" | "WhenSelected";
```

Property Value: [Word.FieldShading](https://learn.microsoft.com/en-us/javascript/api/word/word.fieldshading) | "Never" | "Always" | "WhenSelected"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isDraft

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.

```typescript
isDraft?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isFirstLineOnlyDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether only the first line of body text is shown in outline view.

```typescript
isFirstLineOnlyDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isFormatDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether character formatting is visible in outline view.

```typescript
isFormatDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isFullScreen

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is in full-screen view.

```typescript
isFullScreen?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isHiddenTextDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether text formatted as hidden text is displayed.

```typescript
isHiddenTextDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isHighlightingDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether highlight formatting is displayed and printed with the document.

```typescript
isHighlightingDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isInConflictMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document is in conflict mode view.

```typescript
isInConflictMode?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isInPanning

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word is in Panning mode.

```typescript
isInPanning?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isInReadingLayout

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document is being viewed in reading layout view.

```typescript
isInReadingLayout?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isMailMergeDataView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether mail merge data is displayed instead of mail merge fields.

```typescript
isMailMergeDataView?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isMainTextLayerVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the text in the document is visible when the header and footer areas are displayed.

```typescript
isMainTextLayerVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isPointerShownAsMagnifier

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the pointer is displayed as a magnifying glass in print preview.

```typescript
isPointerShownAsMagnifier?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isReadingLayoutActualView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether pages displayed in reading layout view are displayed using the same layout as printed pages.

```typescript
isReadingLayoutActualView?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isXmlMarkupVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether XML tags are visible in the document.

```typescript
isXmlMarkupVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### markupMode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the display mode for tracked changes.

```typescript
markupMode?: Word.RevisionsMode | "Balloon" | "Inline" | "Mixed";
```

Property Value: [Word.RevisionsMode](https://learn.microsoft.com/en-us/javascript/api/word/word.revisionsmode) | "Balloon" | "Inline" | "Mixed"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pageColor

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page color in Reading mode.

```typescript
pageColor?: Word.PageColor | "None" | "Sepia" | "Inverse";
```

Property Value: [Word.PageColor](https://learn.microsoft.com/en-us/javascript/api/word/word.pagecolor) | "None" | "Sepia" | "Inverse"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pageMovementType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the page movement type.

```typescript
pageMovementType?: Word.PageMovementType | "Vertical" | "SideToSide";
```

Property Value: [Word.PageMovementType](https://learn.microsoft.com/en-us/javascript/api/word/word.pagemovementtype) | "Vertical" | "SideToSide"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### readingLayoutTruncateMargins

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether margins are visible or hidden when the document is viewed in Full Screen Reading view.

```typescript
readingLayoutTruncateMargins?: Word.ReadingLayoutMargin | "Automatic" | "Suppress" | "Full";
```

Property Value: [Word.ReadingLayoutMargin](https://learn.microsoft.com/en-us/javascript/api/word/word.readinglayoutmargin) | "Automatic" | "Suppress" | "Full"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### revisionsBalloonSide

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether Word displays revision balloons in the left or right margin in the document.

```typescript
revisionsBalloonSide?: Word.RevisionsBalloonMargin | "Left" | "Right";
```

Property Value: [Word.RevisionsBalloonMargin](https://learn.microsoft.com/en-us/javascript/api/word/word.revisionsballoonmargin) | "Left" | "Right"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

### revisionsBalloonWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the revision balloons.

```typescript
revisionsBalloonWidth?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

### revisionsBalloonWidthType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies how Microsoft Word measures the width of revision balloons.

```typescript
revisionsBalloonWidthType?: Word.RevisionsBalloonWidthType | "Percent" | "Points";
```

Property Value: [Word.RevisionsBalloonWidthType](https://learn.microsoft.com/en-us/javascript/api/word/word.revisionsballoonwidthtype) | "Percent" | "Points"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

### seekView

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the document element displayed in print layout view.

```typescript
seekView?: Word.SeekView | "MainDocument" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "CurrentPageHeader" | "CurrentPageFooter";
```

Property Value: [Word.SeekView](https://learn.microsoft.com/en-us/javascript/api/word/word.seekview) | "MainDocument" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "CurrentPageHeader" | "CurrentPageFooter"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### splitSpecial

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the active window pane.

```typescript
splitSpecial?: Word.SpecialPane | "None" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "FootnoteContinuationNotice" | "FootnoteContinuationSeparator" | "FootnoteSeparator" | "EndnoteContinuationNotice" | "EndnoteContinuationSeparator" | "EndnoteSeparator" | "Comments" | "CurrentPageHeader" | "CurrentPageFooter" | "Revisions" | "RevisionsHoriz" | "RevisionsVert";
```

Property Value: [Word.SpecialPane](https://learn.microsoft.com/en-us/javascript/api/word/word.specialpane) | "None" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "FootnoteContinuationNotice" | "FootnoteContinuationSeparator" | "FootnoteSeparator" | "EndnoteContinuationNotice" | "EndnoteContinuationSeparator" | "EndnoteSeparator" | "Comments" | "CurrentPageHeader" | "CurrentPageFooter" | "Revisions" | "RevisionsHoriz" | "RevisionsVert"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the view type.

```typescript
type?: Word.ViewType | "Normal" | "Outline" | "Print" | "PrintPreview" | "Master" | "Web" | "Reading" | "Conflict";
```

Property Value: [Word.ViewType](https://learn.microsoft.com/en-us/javascript/api/word/word.viewtype) | "Normal" | "Outline" | "Print" | "PrintPreview" | "Master" | "Web" | "Reading" | "Conflict"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)