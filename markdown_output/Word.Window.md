# Word.Window

**Package:** `word`

**API Set:** WordApiDesktop 1.2

**Extends:** `OfficeExtension.ClientObject`

## Description

Represents the window that displays the document. A window can be split to contain multiple reading panes.

## Class Examples

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

## Properties

### activePane

**Type:** `Word.Pane`

**Since:** WordApiDesktop 1.2

Gets the active pane in the window.

#### Examples

**Example**: Retrieve and log the number and index values of all pages that are currently visible in the active document window's viewport.

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml

await Word.run(async (context) => {
  // Gets the pages enclosing the viewport.

  // Get the active window.
  const activeWindow: Word.Window = context.document.activeWindow;
  activeWindow.load();

  // Get the active pane.
  const activePane: Word.Pane = activeWindow.activePane;
  activePane.load();

  // Get pages enclosing the viewport.
  const pages: Word.PageCollection = activePane.pagesEnclosingViewport;
  pages.load();

  await context.sync();

  // Log the number of pages.
  const pageCount = pages.items.length;
  console.log(`Number of pages enclosing the viewport: ${pageCount}`);

  // Log index info of these pages.
  const pagesIndexes = [];
  for (let i = 0; i < pageCount; i++) {
    const page = pages.items[i];
    page.load('index');
    pagesIndexes.push(page);
  }

  await context.sync();

  for (let i = 0; i < pagesIndexes.length; i++) {
    console.log(`Page index: ${pagesIndexes[i].index}`);
  }
});
```

---

### areRulersDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether rulers are displayed for the window or pane.

#### Examples

**Example**: Toggle the display of rulers in the active Word document window

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Load the current ruler display state
    window.load("areRulersDisplayed");
    await context.sync();
    
    // Toggle the ruler display
    window.areRulersDisplayed = !window.areRulersDisplayed;
    
    await context.sync();
    
    console.log(`Rulers are now ${window.areRulersDisplayed ? 'displayed' : 'hidden'}`);
});
```

---

### areScreenTipsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.

#### Examples

**Example**: Check if screen tips are currently enabled in the document window, and if not, enable them to show helpful tooltips for comments, footnotes, endnotes, and hyperlinks.

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.load("areScreenTipsDisplayed");
    
    await context.sync();
    
    if (!window.areScreenTipsDisplayed) {
        window.areScreenTipsDisplayed = true;
        console.log("Screen tips have been enabled.");
    } else {
        console.log("Screen tips are already enabled.");
    }
    
    await context.sync();
});
```

---

### areThumbnailsDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.

#### Examples

**Example**: Check if thumbnails are currently displayed in the document window and toggle them on if they are off

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.load("areThumbnailsDisplayed");
    
    await context.sync();
    
    if (!window.areThumbnailsDisplayed) {
        window.areThumbnailsDisplayed = true;
        await context.sync();
        console.log("Thumbnails have been enabled");
    } else {
        console.log("Thumbnails are already displayed");
    }
});
```

---

### caption

**Type:** `string`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the caption text for the window that is displayed in the title bar of the document or application window.

#### Examples

**Example**: Set the window caption to "Q4 Sales Report - Draft Version"

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Set the caption text
    window.caption = "Q4 Sales Report - Draft Version";
    
    await context.sync();
});
```

---

### context

**Type:** `Word.RequestContext`

The request context associated with the object. This connects the add-in's process to the Office host application's process.

#### Examples

**Example**: Access the request context from a Window object to verify the connection between the add-in and Word application

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.application.activeWindow;
    
    // Access the request context associated with the window
    const windowContext = window.context;
    
    // Use the context to load window properties
    window.load("isActive");
    await windowContext.sync();
    
    console.log("Window is active:", window.isActive);
    console.log("Request context successfully accessed from window object");
});
```

---

### height

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the height of the window (in points).

#### Examples

**Example**: Set the Word window height to 600 points

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.height = 600;
    await context.sync();
});
```

---

### horizontalPercentScrolled

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the horizontal scroll position as a percentage of the document width.

#### Examples

**Example**: Scroll the document window horizontally to 50% of the document width

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Set horizontal scroll position to 50% of document width
    window.horizontalPercentScrolled = 50;
    
    await context.sync();
    
    console.log("Document scrolled to 50% horizontally");
});
```

---

### imemode

**Type:** `Word.ImeMode | "NoControl" | "On" | "Off" | "Hiragana" | "Katakana" | "KatakanaHalf" | "AlphaFull" | "Alpha" | "HangulFull" | "Hangul"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the default start-up mode for the Japanese Input Method Editor (IME).

#### Examples

**Example**: Set the IME mode to Hiragana for Japanese text input in the active Word window

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Set the IME mode to Hiragana for Japanese input
    window.imeMode = Word.ImeMode.hiragana;
    
    await context.sync();
    
    console.log("IME mode set to Hiragana");
});
```

---

### index

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the position of an item in a collection.

#### Examples

**Example**: Display an alert showing the position (index) of the active window in the windows collection

```typescript
await Word.run(async (context) => {
    const activeWindow = context.document.application.activeWindow;
    activeWindow.load("index");
    
    await context.sync();
    
    console.log(`The active window is at position: ${activeWindow.index}`);
});
```

---

### isActive

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the window is active.

#### Examples

**Example**: Check if the current window is active and display a message to the user indicating the window's active state.

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.load("isActive");
    
    await context.sync();
    
    if (window.isActive) {
        console.log("The current window is active");
    } else {
        console.log("The current window is not active");
    }
});
```

---

### isDocumentMapVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the document map is visible.

#### Examples

**Example**: Toggle the document map visibility to show the navigation pane in the Word window

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.application.activeWindow;
    
    // Show the document map
    window.isDocumentMapVisible = true;
    
    await context.sync();
    
    console.log("Document map is now visible");
});
```

---

### isEnvelopeVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the email message header is visible in the document window. The default value is False.

#### Examples

**Example**: Show the email envelope header in the document window to display recipient and subject information

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Make the envelope header visible
    window.isEnvelopeVisible = true;
    
    await context.sync();
    
    console.log("Email envelope header is now visible");
});
```

---

### isHorizontalScrollBarDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether a horizontal scroll bar is displayed for the window.

#### Examples

**Example**: Check if the horizontal scroll bar is displayed in the active document window and show an alert with the result.

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Load the horizontal scroll bar display property
    window.load("isHorizontalScrollBarDisplayed");
    
    await context.sync();
    
    // Display the result
    console.log(`Horizontal scroll bar is ${window.isHorizontalScrollBarDisplayed ? 'displayed' : 'hidden'}`);
});
```

---

### isLeftScrollBarDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the vertical scroll bar appears on the left side of the document window.

#### Examples

**Example**: Check if the vertical scroll bar is displayed on the left side of the document window and display the result in the console.

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.load("isLeftScrollBarDisplayed");
    
    await context.sync();
    
    console.log(`Scroll bar on left side: ${window.isLeftScrollBarDisplayed}`);
});
```

---

### isRightRulerDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the vertical ruler appears on the right side of the document window in print layout view.

#### Examples

**Example**: Check if the right ruler is displayed and toggle it off if it's currently visible in the document window.

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.load("isRightRulerDisplayed");
    
    await context.sync();
    
    if (window.isRightRulerDisplayed) {
        window.isRightRulerDisplayed = false;
        console.log("Right ruler was visible, now hidden");
    } else {
        console.log("Right ruler is already hidden");
    }
    
    await context.sync();
});
```

---

### isSplit

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the window is split into multiple panes.

#### Examples

**Example**: Check if the current document window is split into multiple panes and display the result in the console

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.load("isSplit");
    
    await context.sync();
    
    console.log(`Window is split: ${window.isSplit}`);
});
```

---

### isVerticalRulerDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether a vertical ruler is displayed for the window or pane.

#### Examples

**Example**: Toggle the vertical ruler display in the active Word document window to make it visible for precise layout measurements.

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Enable the vertical ruler display
    window.isVerticalRulerDisplayed = true;
    
    await context.sync();
    
    console.log("Vertical ruler is now displayed");
});
```

---

### isVerticalScrollBarDisplayed

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether a vertical scroll bar is displayed for the window.

#### Examples

**Example**: Check if the vertical scroll bar is displayed in the current document window and display the result in the console.

```typescript
await Word.run(async (context) => {
    // Get the current window
    const window = context.document.window;
    
    // Load the isVerticalScrollBarDisplayed property
    window.load("isVerticalScrollBarDisplayed");
    
    // Sync to get the property value
    await context.sync();
    
    // Display the result
    console.log(`Vertical scroll bar is displayed: ${window.isVerticalScrollBarDisplayed}`);
});
```

---

### isVisible

**Type:** `boolean`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies whether the window is visible.

#### Examples

**Example**: Check if the current document window is visible and display an alert with the visibility status

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.load("isVisible");
    
    await context.sync();
    
    if (window.isVisible) {
        console.log("The document window is currently visible");
    } else {
        console.log("The document window is currently hidden");
    }
});
```

---

### left

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the horizontal position of the window, measured in points.

#### Examples

**Example**: Move the Word window to a horizontal position of 100 points from the left edge of the screen

```typescript
await Word.run(async (context) => {
    // Get the current window
    const window = context.document.window;
    
    // Set the horizontal position to 100 points from the left
    window.left = 100;
    
    await context.sync();
});
```

---

### next

**Type:** `Word.Window`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the next document window in the collection of open document windows.

#### Examples

**Example**: Navigate to and activate the next open document window in Word.

```typescript
await Word.run(async (context) => {
    // Get the current active window
    const currentWindow = context.document.window;
    
    // Get the next window in the collection
    const nextWindow = currentWindow.next;
    
    // Load the next window's properties
    nextWindow.load("isActive");
    
    await context.sync();
    
    // Activate the next window to bring it to focus
    nextWindow.activate();
    
    await context.sync();
    
    console.log("Switched to the next document window");
});
```

---

### panes

**Type:** `Word.PaneCollection`

**Since:** WordApiDesktop 1.2

Gets the collection of panes in the window.

#### Examples

**Example**: Get the number of panes in the current window and display information about each pane

```typescript
await Word.run(async (context) => {
    // Get the active window and its panes
    const window = context.document.window;
    const panes = window.panes;
    
    // Load the panes collection and count
    panes.load("items");
    
    await context.sync();
    
    // Display information about the panes
    console.log(`Number of panes in window: ${panes.items.length}`);
    
    // Iterate through each pane
    panes.items.forEach((pane, index) => {
        console.log(`Pane ${index + 1} detected`);
    });
});
```

---

### previous

**Type:** `Word.Window`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the previous document window in the collection open document windows.

#### Examples

**Example**: Navigate to the previous open document window and activate it to bring it into focus.

```typescript
await Word.run(async (context) => {
    // Get the current window
    const currentWindow = context.document.window;
    
    // Get the previous window in the collection
    const previousWindow = currentWindow.previous;
    
    // Load the previous window's properties
    previousWindow.load("isActive");
    
    await context.sync();
    
    // Activate the previous window to bring it into focus
    previousWindow.activate();
    
    await context.sync();
    
    console.log("Navigated to the previous document window");
});
```

---

### showSourceDocuments

**Type:** `Word.ShowSourceDocuments | "None" | "Original" | "Revised" | "Both"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies how Microsoft Word displays source documents after a compare and merge process.

#### Examples

**Example**: Configure the window to display both the original and revised documents side-by-side after a compare and merge operation.

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.application.activeWindow;
    
    // Set to display both original and revised documents
    window.showSourceDocuments = Word.ShowSourceDocuments.both;
    
    await context.sync();
    
    console.log("Window configured to show both source documents");
});
```

---

### splitVertical

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical split percentage for the window.

#### Examples

**Example**: Set the window's vertical split to 50% so the document is divided equally between top and bottom panes

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Set vertical split to 50%
    window.splitVertical = 50;
    
    await context.sync();
    console.log("Window split vertically at 50%");
});
```

---

### styleAreaWidth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width of the style area in points.

#### Examples

**Example**: Set the style area width to 100 points to display paragraph styles in the left margin of the document window

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.windows.getActiveOrNullObject();
    
    // Set the style area width to 100 points
    window.styleAreaWidth = 100;
    
    await context.sync();
    
    console.log("Style area width set to 100 points");
});
```

---

### top

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical position of the document window, in points.

#### Examples

**Example**: Position the document window 200 points from the top of the screen

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Set the vertical position to 200 points from the top
    window.top = 200;
    
    await context.sync();
});
```

---

### type

**Type:** `Word.WindowType | "Document" | "Template"`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the window type.

#### Examples

**Example**: Check if the current window is displaying a document or a template and display an alert message accordingly.

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.load("type");
    
    await context.sync();
    
    if (window.type === Word.WindowType.document || window.type === "Document") {
        console.log("This window is displaying a document.");
    } else if (window.type === Word.WindowType.template || window.type === "Template") {
        console.log("This window is displaying a template.");
    }
});
```

---

### usableHeight

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the height (in points) of the active working area in the document window.

#### Examples

**Example**: Display an alert showing the available working height in the document window to help determine if content will fit in the visible area.

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Load the usableHeight property
    window.load("usableHeight");
    
    await context.sync();
    
    // Display the usable height
    console.log(`Available working height: ${window.usableHeight} points`);
    
    // Optional: Show in Office dialog or use for layout decisions
    if (window.usableHeight < 400) {
        console.log("Limited vertical space available");
    }
});
```

---

### usableWidth

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the width (in points) of the active working area in the document window.

#### Examples

**Example**: Display an alert showing the usable width of the document window in points to help determine available space for content layout.

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Load the usableWidth property
    window.load("usableWidth");
    
    // Sync to get the property value
    await context.sync();
    
    // Display the usable width
    console.log(`Usable width: ${window.usableWidth} points`);
});
```

---

### verticalPercentScrolled

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the vertical scroll position as a percentage of the document length.

#### Examples

**Example**: Scroll the document window to 50% of its vertical length to jump to the middle of the document.

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Set the vertical scroll position to 50% (middle of document)
    window.verticalPercentScrolled = 50;
    
    await context.sync();
    
    console.log("Document scrolled to 50% position");
});
```

---

### view

**Type:** `Word.View`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets the View object that represents the view for the window.

#### Examples

**Example**: Check if the current window view is in print layout mode and log the view type to the console.

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Get the view object from the window
    const view = window.view;
    
    // Load the view type property
    view.load("type");
    
    await context.sync();
    
    // Log the view type
    console.log("Current view type: " + view.type);
    // View type can be: "Print", "Outline", "Web", "Reading", etc.
});
```

---

### width

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the width of the document window, in points.

#### Examples

**Example**: Set the document window width to 600 points

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.width = 600;
    await context.sync();
});
```

---

### windowNumber

**Type:** `number`

**Since:** WordApi BETA (PREVIEW ONLY)

Gets an integer that represents the position of the window.

#### Examples

**Example**: Display the position number of the current window in a message to help users identify which window they're working in when multiple windows are open.

```typescript
await Word.run(async (context) => {
    // Get the current window
    const window = context.document.window;
    
    // Load the windowNumber property
    window.load("windowNumber");
    
    // Sync to retrieve the property value
    await context.sync();
    
    // Display the window position
    console.log(`You are working in window number: ${window.windowNumber}`);
});
```

---

### windowState

**Type:** `Word.WindowState | "Normal" | "Maximize" | "Minimize"`

**Since:** WordApi BETA (PREVIEW ONLY)

Specifies the state of the document window or task window.

#### Examples

**Example**: Maximize the current document window to fill the entire screen

```typescript
await Word.run(async (context) => {
    // Get the current window
    const window = context.document.window;
    
    // Set the window state to maximized
    window.windowState = Word.WindowState.maximize;
    
    await context.sync();
    
    console.log("Window has been maximized");
});
```

---

## Methods

### activate

Activates the window.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Activate a specific document window to bring it to the front and give it focus

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Activate the window to bring it to the front
    window.activate();
    
    await context.sync();
    
    console.log("Window has been activated");
});
```

---

### close

Closes the window.

#### Signature

**Parameters:**
- `options`: `Word.WindowCloseOptions` (optional)
  The options that define whether to save changes before closing and whether to route the document.

**Returns:** `void`

#### Examples

**Example**: Close the current document window programmatically

```typescript
await Word.run(async (context) => {
    // Get the current window
    const window = context.document.window;
    
    // Close the window
    window.close();
    
    await context.sync();
});
```

---

### largeScroll

Scrolls the window by the specified number of screens.

#### Signature

**Parameters:**
- `options`: `Word.WindowScrollOptions` (optional)
  The options for scrolling the window by the specified number of screens. If no options are specified, the window is scrolled down one screen.

**Returns:** `void`

#### Examples

**Example**: Scroll the document window down by 2 screens to quickly navigate through a long document

```typescript
await Word.run(async (context) => {
    const window = context.document.getActiveWindow();
    
    // Scroll down by 2 screens
    window.largeScroll({ down: 2 });
    
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
  - `options`: `Word.Interfaces.WindowLoadOptions` (optional)
    Provides options for which properties of the object to load.

  **Returns:** `Word.Window`

**Overload 2:**

  **Parameters:**
  - `propertyNames`: `string | string[]` (optional)
    A comma-delimited string or an array of strings that specify the properties to load.

  **Returns:** `Word.Window`

**Overload 3:**

  **Parameters:**
  - `propertyNamesAndPaths`: `{ select?: string; expand?: string; }` (optional)
    propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

  **Returns:** `Word.Window`

#### Examples

**Example**: Load and display the current window's split state to check if the document window is split into multiple panes

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    
    // Load the split property of the window
    window.load("split");
    
    await context.sync();
    
    // Check if the window is split
    if (window.split) {
        console.log("The window is split into multiple panes");
    } else {
        console.log("The window is not split");
    }
});
```

---

### pageScroll

Scrolls through the window page by page.

#### Signature

**Parameters:**
- `options`: `Word.WindowPageScrollOptions` (optional)
  The options for scrolling through the window page by page.

**Returns:** `void`

#### Examples

**Example**: Scroll down through the document by one page to view content that is currently below the visible area

```typescript
await Word.run(async (context) => {
    const window = context.document.getActiveWindow();
    
    // Scroll down by one page
    window.pageScroll({ scrollDirection: Word.PageScrollDirection.down });
    
    await context.sync();
});
```

---

### set

**Kind:** `write`

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

#### Signatures

**Overload 1:**

  **Parameters:**
  - `properties`: `Interfaces.WindowUpdateData` (required)
    A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
  - `options`: `OfficeExtension.UpdateOptions` (optional)
    Provides an option to suppress errors if the properties object tries to set any read-only properties.

  **Returns:** `void`

**Overload 2:**

  **Parameters:**
  - `properties`: `Word.Window` (required)

  **Returns:** `void`

#### Examples

**Example**: Configure multiple window properties at once to set up a split view with specific dimensions

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    
    // Set multiple window properties at once
    window.set({
        split: true,
        splitPercentage: 50
    });
    
    await context.sync();
    console.log("Window configured with split view at 50%");
});
```

---

### setFocus

Sets the focus of the document window to the body of an email message.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Set focus to the email message body when a user clicks a button in a Word add-in for composing emails

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Set focus to the email message body
    window.setFocus();
    
    await context.sync();
    
    console.log("Focus has been set to the email message body");
});
```

---

### smallScroll

Scrolls the window by the specified number of lines. A "line" corresponds to the distance scrolled by clicking the scroll arrow on the scroll bar once.

#### Signature

**Parameters:**
- `options`: `Word.WindowScrollOptions` (optional)
  The options for scrolling the window by the specified number of lines. If no options are specified, the window is scrolled down by one line.

**Returns:** `void`

#### Examples

**Example**: Scroll the active document window down by 5 lines to view content below the current viewport

```typescript
await Word.run(async (context) => {
    const window = context.document.getActiveWindow();
    
    // Scroll down by 5 lines
    window.smallScroll({ down: 5 });
    
    await context.sync();
});
```

---

### toggleRibbon

Shows or hides the ribbon.

#### Signature

**Returns:** `void`

#### Examples

**Example**: Toggle the ribbon visibility to maximize the document editing area

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Toggle the ribbon visibility
    window.toggleRibbon();
    
    await context.sync();
    
    console.log("Ribbon visibility toggled");
});
```

---

### toJSON

**Kind:** `serialize`

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Window object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.WindowData) that contains shallow copies of any loaded child properties from the original object.

#### Signature

**Returns:** `Word.Interfaces.WindowData`

#### Examples

**Example**: Serialize the active window's properties to JSON format for logging or debugging purposes.

```typescript
await Word.run(async (context) => {
    // Get the active window
    const window = context.document.window;
    
    // Load properties to serialize
    window.load("width,height");
    
    await context.sync();
    
    // Convert the window object to a plain JavaScript object
    const windowData = window.toJSON();
    
    // Output the serialized data (e.g., for logging or debugging)
    console.log("Window data:", JSON.stringify(windowData, null, 2));
});
```

---

### track

**Kind:** `track`

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

#### Signature

**Returns:** `Word.Window`

#### Examples

**Example**: Track a window object across multiple sync calls to maintain its reference while modifying document properties and reading window state

```typescript
await Word.run(async (context) => {
    const window = context.document.window;
    window.track();
    
    // Load window properties
    window.load("activePane");
    await context.sync();
    
    // Perform operations that might change the document
    const body = context.document.body;
    body.insertParagraph("New content added", Word.InsertLocation.end);
    await context.sync();
    
    // Access the tracked window object again after sync
    console.log("Active pane type: " + window.activePane);
    
    // Clean up tracking when done
    window.untrack();
});
```

---

### untrack

**Kind:** `untrack`

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

#### Signature

**Returns:** `Word.Window`

#### Examples

**Example**: Get a reference to the active window, track it for change monitoring, then untrack it to release memory after use

```typescript
await Word.run(async (context) => {
    // Get the active window and track it
    const window = context.document.window;
    context.trackedObjects.add(window);
    
    // Load properties to use the window
    window.load("width,height");
    await context.sync();
    
    console.log(`Window dimensions: ${window.width}x${window.height}`);
    
    // Untrack the window to release memory
    window.untrack();
    await context.sync();
    
    console.log("Window object memory released");
});
```

---

## Source

- https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/35-ranges/get-pages.yaml
