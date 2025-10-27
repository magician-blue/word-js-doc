# Word.View

**Package:** `https://learn.microsoft.com/en-us/javascript/api/word`

**API Set:** WordApi BETA (PREVIEW ONLY)

**Extends:** `OfficeExtension.ClientObject`

## Description

Contains the view attributes (such as show all, field shading, and table gridlines) for a window or pane.

## Properties

### areAllNonprintingCharactersDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether all nonprinting characters are displayed.

#### Examples

**Example**: Display all nonprinting characters (such as spaces, paragraph marks, and tabs) in the document to review formatting

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Enable display of all nonprinting characters
    view.areAllNonprintingCharactersDisplayed = true;
    
    await context.sync();
    
    console.log("All nonprinting characters are now displayed");
});
```

---

### areBackgroundsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether background colors and images are shown when the document is displayed in print layout view.

#### Examples

**Example**: Check if background colors and images are currently displayed in print layout view and show the result in a message

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areBackgroundsDisplayed");
    
    await context.sync();
    
    console.log(`Backgrounds displayed: ${view.areBackgroundsDisplayed}`);
});
```

---

### areBookmarksIndicated

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether square brackets are displayed at the beginning and end of each bookmark.

#### Examples

**Example**: Check if bookmark indicators are currently displayed and toggle them on to show square brackets around bookmarks in the document.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areBookmarksIndicated");
    
    await context.sync();
    
    if (!view.areBookmarksIndicated) {
        view.areBookmarksIndicated = true;
        console.log("Bookmark indicators enabled - square brackets will now display around bookmarks");
    } else {
        console.log("Bookmark indicators are already enabled");
    }
    
    await context.sync();
});
```

---

### areCommentsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether Microsoft Word displays the comments in the document.

#### Examples

**Example**: Hide all comments in the active document to get a cleaner view of the content

```typescript
await Word.run(async (context) => {
    // Get the view of the active document
    const view = context.document.getActiveView();
    
    // Hide comments in the document
    view.areCommentsDisplayed = false;
    
    await context.sync();
});
```

---

### areConnectingLinesToRevisionsBalloonDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether Microsoft Word displays connecting lines from the text to the revision and comment balloons.

#### Examples

**Example**: Hide the connecting lines between text and revision balloons in the document view

```typescript
await Word.run(async (context) => {
    // Get the view of the active document
    const view = context.document.getActiveView();
    
    // Hide connecting lines to revision balloons
    view.areConnectingLinesToRevisionsBalloonDisplayed = false;
    
    await context.sync();
});
```

---

### areCropMarksDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether crop marks are shown in the corners of pages to indicate where margins are located.

#### Examples

**Example**: Check if crop marks are currently displayed in the document and log the result to the console

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areCropMarksDisplayed");
    
    await context.sync();
    
    console.log(`Crop marks are ${view.areCropMarksDisplayed ? 'displayed' : 'not displayed'}`);
});
```

---

### areDrawingsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether objects created with the drawing tools are displayed in print layout view.

#### Examples

**Example**: Check if drawing objects are currently displayed in the document's print layout view and log the result to the console.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areDrawingsDisplayed");
    
    await context.sync();
    
    console.log(`Drawing objects are ${view.areDrawingsDisplayed ? 'displayed' : 'hidden'} in print layout view`);
});
```

---

### areEditableRangesShaded

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether shading is applied to the ranges in the document that users have permission to modify.

#### Examples

**Example**: Enable shading for editable ranges in the document so users can visually identify which parts they have permission to modify

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Enable shading for editable ranges
    view.areEditableRangesShaded = true;
    
    await context.sync();
    
    console.log("Editable ranges shading has been enabled");
});
```

---

### areFieldCodesDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether field codes are displayed.

#### Examples

**Example**: Toggle the display of field codes in the document to show the underlying field code syntax instead of field results

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Display field codes instead of field results
    view.areFieldCodesDisplayed = true;
    
    await context.sync();
    
    console.log("Field codes are now displayed");
});
```

---

### areFormatChangesDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether Microsoft Word displays formatting changes made to the document with Track Changes enabled.

#### Examples

**Example**: Check if formatting changes are currently being displayed in the document and toggle the display of formatting changes made with Track Changes enabled.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areFormatChangesDisplayed");
    
    await context.sync();
    
    console.log("Current state:", view.areFormatChangesDisplayed);
    
    // Toggle the display of formatting changes
    view.areFormatChangesDisplayed = !view.areFormatChangesDisplayed;
    
    await context.sync();
    
    console.log("Formatting changes display toggled to:", view.areFormatChangesDisplayed);
});
```

---

### areInkAnnotationsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether handwritten ink annotations are shown or hidden.

#### Examples

**Example**: Hide handwritten ink annotations in the current document view

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Hide ink annotations
    view.areInkAnnotationsDisplayed = false;
    
    await context.sync();
});
```

---

### areInsertionsAndDeletionsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether Microsoft Word displays insertions and deletions made to the document with Track Changes enabled.

#### Examples

**Example**: Check if tracked changes (insertions and deletions) are currently displayed in the document, and if not, enable their display to review all edits.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areInsertionsAndDeletionsDisplayed");
    
    await context.sync();
    
    if (!view.areInsertionsAndDeletionsDisplayed) {
        view.areInsertionsAndDeletionsDisplayed = true;
        await context.sync();
        console.log("Tracked changes are now visible");
    } else {
        console.log("Tracked changes are already visible");
    }
});
```

---

### areLinesWrappedToWindow

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether lines wrap at the right edge of the document window rather than at the right margin or the right column boundary.

#### Examples

**Example**: Check if text wrapping is set to window edge and display the result in the console

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areLinesWrappedToWindow");
    
    await context.sync();
    
    console.log(`Lines wrapped to window: ${view.areLinesWrappedToWindow}`);
});
```

---

### areObjectAnchorsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether object anchors are displayed next to items that can be positioned in print layout view.

#### Examples

**Example**: Check if object anchors are currently displayed in the document and show an alert with the result

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areObjectAnchorsDisplayed");
    
    await context.sync();
    
    console.log(`Object anchors are ${view.areObjectAnchorsDisplayed ? 'displayed' : 'hidden'}`);
});
```

---

### areOptionalBreaksDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether Microsoft Word displays optional line breaks.

#### Examples

**Example**: Check if optional line breaks are currently displayed in the document and show an alert with the result

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areOptionalBreaksDisplayed");
    
    await context.sync();
    
    console.log(`Optional line breaks displayed: ${view.areOptionalBreaksDisplayed}`);
});
```

---

### areOptionalHyphensDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether optional hyphens are displayed.

#### Examples

**Example**: Check if optional hyphens are currently displayed in the document and show an alert with the result

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areOptionalHyphensDisplayed");
    
    await context.sync();
    
    console.log(`Optional hyphens displayed: ${view.areOptionalHyphensDisplayed}`);
    
    if (view.areOptionalHyphensDisplayed) {
        console.log("Optional hyphens are currently visible in the document.");
    } else {
        console.log("Optional hyphens are currently hidden in the document.");
    }
});
```

---

### areOtherAuthorsVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether other authors' presence should be visible in the document.

#### Examples

**Example**: Check if other authors are visible in the document and display the result in the console.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areOtherAuthorsVisible");
    
    await context.sync();
    
    console.log(`Other authors visible: ${view.areOtherAuthorsVisible}`);
});
```

---

### arePageBoundariesDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether the top and bottom margins and the gray area between pages in the document are displayed.

#### Examples

**Example**: Check if page boundaries are currently displayed in the document view and log the result to the console.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("arePageBoundariesDisplayed");
    
    await context.sync();
    
    console.log(`Page boundaries displayed: ${view.arePageBoundariesDisplayed}`);
});
```

---

### areParagraphsMarksDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether paragraph marks are displayed.

#### Examples

**Example**: Check if paragraph marks are currently displayed in the document and show an alert with the result

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areParagraphsMarksDisplayed");
    
    await context.sync();
    
    if (view.areParagraphsMarksDisplayed) {
        console.log("Paragraph marks are currently displayed");
    } else {
        console.log("Paragraph marks are currently hidden");
    }
});
```

---

### arePicturePlaceholdersDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether blank boxes are displayed as placeholders for pictures.

#### Examples

**Example**: Check if picture placeholders are currently displayed in the document view and log the result to the console.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("arePicturePlaceholdersDisplayed");
    
    await context.sync();
    
    console.log("Picture placeholders displayed: " + view.arePicturePlaceholdersDisplayed);
});
```

---

### areRevisionsAndCommentsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether Microsoft Word displays revisions and comments made to the document with Track Changes enabled.

#### Examples

**Example**: Check if revisions and comments are currently displayed in the document, and if not, enable their display to show all tracked changes and comments.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areRevisionsAndCommentsDisplayed");
    
    await context.sync();
    
    if (!view.areRevisionsAndCommentsDisplayed) {
        view.areRevisionsAndCommentsDisplayed = true;
        await context.sync();
        console.log("Revisions and comments are now displayed");
    } else {
        console.log("Revisions and comments are already displayed");
    }
});
```

---

### areSpacesIndicated

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether space characters are displayed.

#### Examples

**Example**: Check if space characters are currently displayed in the document view and log the result to the console

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areSpacesIndicated");
    
    await context.sync();
    
    console.log(`Space characters are ${view.areSpacesIndicated ? 'visible' : 'hidden'} in the document view`);
});
```

---

### areTableGridlinesDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether table gridlines are displayed.

#### Examples

**Example**: Toggle the display of table gridlines in the document to make them visible for easier table editing

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Enable table gridlines display
    view.areTableGridlinesDisplayed = true;
    
    await context.sync();
    
    console.log("Table gridlines are now displayed");
});
```

---

### areTabsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether tab characters are displayed.

#### Examples

**Example**: Check if tab characters are currently displayed in the document view and log the result to the console.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areTabsDisplayed");
    
    await context.sync();
    
    console.log(`Tab characters are ${view.areTabsDisplayed ? 'visible' : 'hidden'} in the document view.`);
});
```

---

### areTextBoundariesDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether dotted lines are displayed around page margins, text columns, objects, and frames in print layout view.

#### Examples

**Example**: Check if text boundaries are currently displayed in the document and show an alert with the result

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("areTextBoundariesDisplayed");
    
    await context.sync();
    
    if (view.areTextBoundariesDisplayed) {
        console.log("Text boundaries are currently displayed");
    } else {
        console.log("Text boundaries are not displayed");
    }
});
```

---

### columnWidth

**Type:** `Word.ColumnWidth | "Narrow" | "Default" | "Wide"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the column width in Reading mode.

#### Examples

**Example**: Set the Reading mode column width to "Narrow" for easier reading of dense content

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Set the column width to Narrow in Reading mode
    view.columnWidth = "Narrow";
    
    await context.sync();
    
    console.log("Column width set to Narrow");
});
```

---

### context

**Type:** `RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a View object to load and read the current view's showParagraphMarks property

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    
    // Access the request context associated with the view object
    const requestContext = view.context;
    
    // Use the context to load properties
    view.load("showParagraphMarks");
    
    await requestContext.sync();
    
    console.log("Show paragraph marks: " + view.showParagraphMarks);
});
```

---

### fieldShading

**Type:** `Word.FieldShading | "Never" | "Always" | "WhenSelected"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets on-screen shading for fields.

#### Examples

**Example**: Check the current field shading setting and display it to the user, then change it to always show field shading

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Load the fieldShading property
    view.load("fieldShading");
    
    await context.sync();
    
    // Display current setting
    console.log("Current field shading: " + view.fieldShading);
    
    // Set field shading to always show
    view.fieldShading = Word.FieldShading.always;
    
    await context.sync();
    
    console.log("Field shading set to: Always");
});
```

---

### isDraft

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether all the text in a window is displayed in the same sans-serif font with minimal formatting to speed up display.

#### Examples

**Example**: Enable draft view mode to display all document text in a simple sans-serif font for faster rendering and editing performance.

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Enable draft mode for faster display
    view.isDraft = true;
    
    await context.sync();
    
    console.log("Draft view mode enabled");
});
```

---

### isFirstLineOnlyDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether only the first line of body text is shown in outline view.

#### Examples

**Example**: Set the outline view to display only the first line of each paragraph's body text

```typescript
await Word.run(async (context) => {
    // Get the view of the active document
    const view = context.document.getActiveView();
    
    // Set to display only the first line in outline view
    view.isFirstLineOnlyDisplayed = true;
    
    // Sync to apply the changes
    await context.sync();
    
    console.log("Outline view set to display only first line of body text");
});
```

---

### isFormatDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether character formatting is visible in outline view.

#### Examples

**Example**: Show character formatting in outline view by enabling the format display setting

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Enable character formatting display in outline view
    view.isFormatDisplayed = true;
    
    await context.sync();
    console.log("Character formatting is now visible in outline view");
});
```

---

### isFullScreen

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the window is in full-screen view.

#### Examples

**Example**: Toggle the Word document window to full-screen view mode

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveWindow().view;
    
    // Set the window to full-screen mode
    view.isFullScreen = true;
    
    await context.sync();
    
    console.log("Window is now in full-screen mode");
});
```

---

### isHiddenTextDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether text formatted as hidden text is displayed.

#### Examples

**Example**: Check if hidden text is currently displayed in the document and show an alert with the result

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("isHiddenTextDisplayed");
    
    await context.sync();
    
    if (view.isHiddenTextDisplayed) {
        console.log("Hidden text is currently displayed in the document.");
    } else {
        console.log("Hidden text is currently hidden in the document.");
    }
});
```

---

### isHighlightingDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether highlight formatting is displayed and printed with the document.

#### Examples

**Example**: Check if text highlighting is currently displayed in the document and show an alert with the result.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("isHighlightingDisplayed");
    
    await context.sync();
    
    if (view.isHighlightingDisplayed) {
        console.log("Highlighting is currently displayed in the document.");
    } else {
        console.log("Highlighting is not displayed in the document.");
    }
});
```

---

### isInConflictMode

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the document is in conflict mode view.

#### Examples

**Example**: Check if the document is currently in conflict mode view and display an alert to the user if conflicts are detected.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("isInConflictMode");
    
    await context.sync();
    
    if (view.isInConflictMode) {
        console.log("Warning: Document is in conflict mode. Please resolve conflicts before continuing.");
    } else {
        console.log("Document is not in conflict mode.");
    }
});
```

---

### isInPanning

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether Microsoft Word is in Panning mode.

#### Examples

**Example**: Check if the document view is currently in panning mode and display an alert message to the user.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("isInPanning");
    
    await context.sync();
    
    if (view.isInPanning) {
        console.log("The document is currently in panning mode.");
    } else {
        console.log("The document is not in panning mode.");
    }
});
```

---

### isInReadingLayout

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the document is being viewed in reading layout view.

#### Examples

**Example**: Check if the document is currently in reading layout view and display an alert with the result

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("isInReadingLayout");
    
    await context.sync();
    
    if (view.isInReadingLayout) {
        console.log("The document is currently in reading layout view.");
    } else {
        console.log("The document is not in reading layout view.");
    }
});
```

---

### isMailMergeDataView

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether mail merge data is displayed instead of mail merge fields.

#### Examples

**Example**: Check if mail merge data is currently displayed and toggle the view to show mail merge fields instead of the actual data

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("isMailMergeDataView");
    
    await context.sync();
    
    if (view.isMailMergeDataView) {
        // Currently showing data, switch to show fields
        view.isMailMergeDataView = false;
        console.log("Switched to mail merge fields view");
    } else {
        // Currently showing fields, switch to show data
        view.isMailMergeDataView = true;
        console.log("Switched to mail merge data view");
    }
    
    await context.sync();
});
```

---

### isMainTextLayerVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the text in the document is visible when the header and footer areas are displayed.

#### Examples

**Example**: Hide the main document text while viewing headers and footers to focus only on the header/footer content

```typescript
await Word.run(async (context) => {
    // Get the active document's view
    const view = context.document.getActiveView();
    
    // Hide the main text layer when headers/footers are displayed
    view.isMainTextLayerVisible = false;
    
    // Sync to apply the changes
    await context.sync();
    
    console.log("Main text layer is now hidden when viewing headers/footers");
});
```

---

### isPointerShownAsMagnifier

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the pointer is displayed as a magnifying glass in print preview.

#### Examples

**Example**: Check if the pointer is shown as a magnifying glass in print preview and display the result in the console

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("isPointerShownAsMagnifier");
    
    await context.sync();
    
    console.log(`Pointer shown as magnifier: ${view.isPointerShownAsMagnifier}`);
});
```

---

### isReadingLayoutActualView

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether pages displayed in reading layout view are displayed using the same layout as printed pages.

#### Examples

**Example**: Check if the reading layout view is using the actual page layout, and if not, enable it to match the printed page appearance.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("isReadingLayoutActualView");
    
    await context.sync();
    
    if (!view.isReadingLayoutActualView) {
        view.isReadingLayoutActualView = true;
        console.log("Reading layout now displays actual page layout");
    } else {
        console.log("Reading layout already displays actual page layout");
    }
    
    await context.sync();
});
```

---

### isXmlMarkupVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether XML tags are visible in the document.

#### Examples

**Example**: Hide XML markup tags in the current document view to provide a cleaner reading experience

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Hide XML markup tags
    view.isXmlMarkupVisible = false;
    
    await context.sync();
});
```

---

### markupMode

**Type:** `Word.RevisionsMode | "Balloon" | "Inline" | "Mixed"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the display mode for tracked changes.

#### Examples

**Example**: Set the tracked changes display mode to show revisions inline within the document text

```typescript
await Word.run(async (context) => {
    // Get the view of the active document
    const view = context.document.getActiveView();
    
    // Set the markup mode to display revisions inline
    view.markupMode = Word.RevisionsMode.inline;
    
    await context.sync();
    
    console.log("Markup mode set to inline");
});
```

---

### pageColor

**Type:** `Word.PageColor | "None" | "Sepia" | "Inverse"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the page color in Reading mode.

#### Examples

**Example**: Set the page color to sepia in Reading mode to reduce eye strain

```typescript
await Word.run(async (context) => {
    // Get the view of the active document
    const view = context.document.getActiveView();
    
    // Set the page color to sepia for Reading mode
    view.pageColor = Word.PageColor.sepia;
    
    await context.sync();
});
```

---

### pageMovementType

**Type:** `Word.PageMovementType | "Vertical" | "SideToSide"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the page movement type.

#### Examples

**Example**: Set the page movement type to side-to-side reading mode for the active document view

```typescript
await Word.run(async (context) => {
    // Get the view of the active document
    const view = context.document.getActiveView();
    
    // Set the page movement type to side-to-side
    view.pageMovementType = Word.PageMovementType.sideToSide;
    
    await context.sync();
    
    console.log("Page movement type set to side-to-side");
});
```

---

### readingLayoutTruncateMargins

**Type:** `Word.ReadingLayoutMargin | "Automatic" | "Suppress" | "Full"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether margins are visible or hidden when the document is viewed in Full Screen Reading view.

#### Examples

**Example**: Set the reading layout to suppress margins when viewing the document in Full Screen Reading view

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Suppress margins in Full Screen Reading view
    view.readingLayoutTruncateMargins = Word.ReadingLayoutMargin.suppress;
    
    await context.sync();
    
    console.log("Reading layout margins have been suppressed");
});
```

---

### revisionsBalloonSide

**Type:** `Word.RevisionsBalloonMargin | "Left" | "Right"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets whether Word displays revision balloons in the left or right margin in the document.

#### Examples

**Example**: Check which side of the document displays revision balloons and log the result to the console.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("revisionsBalloonSide");
    
    await context.sync();
    
    console.log(`Revision balloons are displayed on the ${view.revisionsBalloonSide} side`);
});
```

---

### revisionsBalloonWidth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width of the revision balloons.

#### Examples

**Example**: Set the revision balloons width to 300 pixels to make tracked changes more readable in the document.

```typescript
await Word.run(async (context) => {
    // Get the current view
    const view = context.document.getActiveView();
    
    // Set the revision balloons width to 300 pixels
    view.revisionsBalloonWidth = 300;
    
    await context.sync();
});
```

---

### revisionsBalloonWidthType

**Type:** `Word.RevisionsBalloonWidthType | "Percent" | "Points"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies how Microsoft Word measures the width of revision balloons.

#### Examples

**Example**: Set the revision balloons to use a percentage-based width measurement instead of points

```typescript
await Word.run(async (context) => {
    // Get the active document's view
    const view = context.document.getActiveView();
    
    // Set the revisions balloon width type to percentage
    view.revisionsBalloonWidthType = Word.RevisionsBalloonWidthType.percent;
    // Alternative: view.revisionsBalloonWidthType = "Percent";
    
    await context.sync();
    
    console.log("Revision balloons will now use percentage-based width measurement");
});
```

---

### revisionsFilter

**Type:** `Word.RevisionsFilter`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the instance of a RevisionsFilter object.

#### Examples

**Example**: Configure the revisions filter to show only insertions and deletions while hiding formatting changes in the document view.

```typescript
await Word.run(async (context) => {
    // Get the revisions filter from the document view
    const revisionsFilter = context.document.getActiveView().revisionsFilter;
    
    // Configure which revision types to display
    revisionsFilter.showInsertions = true;
    revisionsFilter.showDeletions = true;
    revisionsFilter.showFormatting = false;
    
    // Sync to apply the filter settings
    await context.sync();
    
    console.log("Revisions filter configured to show insertions and deletions only");
});
```

---

### seekView

**Type:** `Word.SeekView | "MainDocument" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "CurrentPageHeader" | "CurrentPageFooter"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the document element displayed in print layout view.

#### Examples

**Example**: Switch the view to display the primary header section of the document in print layout view

```typescript
await Word.run(async (context) => {
    // Get the active document's view
    const view = context.document.getActiveView();
    
    // Switch to the primary header view
    view.seekView = Word.SeekView.primaryHeader;
    
    await context.sync();
    
    console.log("View switched to primary header");
});
```

---

### splitSpecial

**Type:** `Word.SpecialPane | "None" | "PrimaryHeader" | "FirstPageHeader" | "EvenPagesHeader" | "PrimaryFooter" | "FirstPageFooter" | "EvenPagesFooter" | "Footnotes" | "Endnotes" | "FootnoteContinuationNotice" | "FootnoteContinuationSeparator" | "FootnoteSeparator" | "EndnoteContinuationNotice" | "EndnoteContinuationSeparator" | "EndnoteSeparator" | "Comments" | "CurrentPageHeader" | "CurrentPageFooter" | "Revisions" | "RevisionsHoriz" | "RevisionsVert"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the active window pane.

#### Examples

**Example**: Check if the current active pane is showing footnotes and display an alert message

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("splitSpecial");
    
    await context.sync();
    
    if (view.splitSpecial === "Footnotes") {
        console.log("The active pane is currently showing footnotes");
    } else {
        console.log(`The active pane is: ${view.splitSpecial}`);
    }
});
```

---

### type

**Type:** `Word.ViewType | "Normal" | "Outline" | "Print" | "PrintPreview" | "Master" | "Web" | "Reading" | "Conflict"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the view type.

#### Examples

**Example**: Check the current view type of the document and switch it to Print Layout view if it's not already in that mode.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.load("type");
    
    await context.sync();
    
    console.log("Current view type:", view.type);
    
    if (view.type !== Word.ViewType.print) {
        view.type = Word.ViewType.print;
        await context.sync();
        console.log("View changed to Print Layout");
    } else {
        console.log("Already in Print Layout view");
    }
});
```

---

## Methods

### collapseAllHeadings

**Kind:** `configure`

Collapses all the headings in the document.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Collapse all headings in the document to show only the top-level structure

```typescript
await Word.run(async (context) => {
    // Get the document's view
    const view = context.document.getActiveView();
    
    // Collapse all headings in the document
    view.collapseAllHeadings();
    
    await context.sync();
});
```

---

### collapseOutline

**Kind:** `configure`

Collapses the text under the selection or the specified range by one heading level.

#### Signature

**Parameters:**
- `range`: `Word.Range` (required)
  A Range object that specifies the range to collapse.

**Returns:** `void`

#### Examples

**Example**: Collapse all content under the first heading in the document to show only the heading text

```typescript
await Word.run(async (context) => {
    // Get the first heading in the document
    const headings = context.document.body.paragraphs;
    headings.load("items");
    await context.sync();
    
    // Find the first heading paragraph
    const firstHeading = headings.items[0];
    const range = firstHeading.getRange();
    
    // Collapse the outline under this heading
    context.document.getActiveView().collapseOutline(range);
    
    await context.sync();
});
```

---

### expandAllHeadings

**Kind:** `configure`

Expands all the headings in the document.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Expand all collapsed headings in the document to show all content under each heading level

```typescript
await Word.run(async (context) => {
    // Get the view of the active document
    const view = context.document.getActiveView();
    
    // Expand all headings in the document
    view.expandAllHeadings();
    
    await context.sync();
    
    console.log("All headings have been expanded.");
});
```

---

### expandOutline

**Kind:** `configure`

Expands the text under the selection by one heading level.

#### Signature

**Parameters:**
- `range`: `Word.Range` (required)
  A Range object that specifies the range to expand.

**Returns:** `void`

#### Examples

**Example**: Expand all collapsed content under the currently selected heading in the document by one outline level

```typescript
await Word.run(async (context) => {
    // Get the current selection
    const selection = context.document.getSelection();
    
    // Get the document view
    const view = context.document.getActiveView();
    
    // Expand the outline under the selection by one level
    view.expandOutline(selection);
    
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
  - `options`: `Word.Interfaces.ViewLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.View`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.View`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.View`

#### Examples

**Example**: Check if field shading is currently enabled in the document view and display the result

```typescript
await Word.run(async (context) => {
    // Get the document view
    const view = context.document.getActiveView();
    
    // Load the field shading property
    view.load("fieldShading");
    
    // Sync to read the loaded property
    await context.sync();
    
    // Display the field shading status
    console.log(`Field shading is ${view.fieldShading ? 'enabled' : 'disabled'}`);
});
```

---

### nextHeaderFooter

**Kind:** `configure`

Moves to the next header or footer, depending on whether a header or footer is displayed in the view.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Navigate through all headers and footers in the active document by moving to the next header/footer section sequentially

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    
    // Switch to print layout to work with headers/footers
    view.type = Word.ViewType.printLayout;
    
    // Move to the first header
    view.showHeaderFooter = true;
    
    // Navigate to the next header/footer (e.g., from header to footer)
    view.nextHeaderFooter();
    
    await context.sync();
    console.log("Moved to the next header or footer section");
});
```

---

### previousHeaderFooter

**Kind:** `configure`

Moves to the previous header or footer, depending on whether a header or footer is displayed in the view.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Navigate to the previous header or footer section in the document to modify its content

```typescript
await Word.run(async (context) => {
    // Get the active document's view
    const view = context.document.getActiveView();
    
    // Move to the previous header or footer
    view.previousHeaderFooter();
    
    // Sync to apply the navigation
    await context.sync();
    
    console.log("Navigated to the previous header or footer section");
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.ViewUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.View` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure the document view to show all formatting marks and enable field shading

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    
    view.set({
        showAllFieldCodes: false,
        showFieldCodes: false,
        showHiddenText: true,
        showParagraphMarks: true,
        showSpaces: true,
        showTabs: true,
        fieldShading: Word.FieldShading.always
    });
    
    await context.sync();
});
```

---

### showAllHeadings

**Kind:** `configure`

Switches between showing all text (headings and body text) and showing only headings.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Toggle the document view to show only headings (outline view) by hiding body text, useful for reviewing document structure

```typescript
await Word.run(async (context) => {
    // Get the active document's view
    const view = context.document.getActiveView();
    
    // Toggle between showing all text and showing only headings
    view.showAllHeadings();
    
    await context.sync();
});
```

---

### showHeading

**Kind:** `configure`

Shows all headings up to the specified heading level and hides subordinate headings and body text.

#### Signature

**Parameters:**
- `level`: `number` (required)
  The heading level to show.

**Returns:** `void`

#### Examples

**Example**: Show only the top 2 heading levels in the document and hide all level 3+ headings and body text

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    view.showHeading(2);
    await context.sync();
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.View object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.ViewData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.ViewData`

#### Examples

**Example**: Serialize the current document view settings to JSON format for logging or storage purposes.

```typescript
await Word.run(async (context) => {
    // Get the active document's view
    const view = context.document.getActiveView();
    
    // Load the view properties
    view.load("type,zoom,showParagraphMarks,showHiddenText,showFieldCodes");
    
    await context.sync();
    
    // Convert the view object to a plain JavaScript object
    const viewData = view.toJSON();
    
    // Log or store the serialized view data
    console.log("Current view settings:", JSON.stringify(viewData, null, 2));
    
    // You can now use viewData as a plain object
    console.log("View type:", viewData.type);
    console.log("Zoom level:", viewData.zoom);
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.View`

#### Examples

**Example**: Track a view object to maintain its reference across multiple sync calls while toggling field shading on and off with a delay between operations.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    
    // Track the view object to use it across multiple sync calls
    view.track();
    
    // Load and toggle field shading
    view.load("fieldShading");
    await context.sync();
    
    const originalShading = view.fieldShading;
    view.fieldShading = !originalShading;
    await context.sync();
    
    console.log(`Field shading changed from ${originalShading} to ${view.fieldShading}`);
    
    // Simulate some delay or additional operations
    view.fieldShading = originalShading;
    await context.sync();
    
    // Clean up tracking when done
    view.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.View`

#### Examples

**Example**: Track a view object to modify its properties, then untrack it to release memory after the changes are complete.

```typescript
await Word.run(async (context) => {
    const view = context.document.getActiveView();
    
    // Track the view object to work with it
    view.track();
    
    // Load and modify view properties
    view.load("showParagraphMarks");
    await context.sync();
    
    view.showParagraphMarks = true;
    await context.sync();
    
    // Release the memory associated with the tracked view object
    view.untrack();
    await context.sync();
});
```

---

## Source

- https://learn.microsoft.com/en-us/javascript/api/word/word.view
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject
