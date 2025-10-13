# Word.Window class

Represents the window that displays the document. A window can be split to contain multiple reading panes.

Package: word

Extends: OfficeExtension.ClientObject

## Remarks
[API set: WordApiDesktop 1.2]

#### Examples
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
- activePane: Gets the active pane in the window.
- areRulersDisplayed: Specifies whether rulers are displayed for the window or pane.
- areScreenTipsDisplayed: Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.
- areThumbnailsDisplayed: Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.
- caption: Specifies the caption text for the window that is displayed in the title bar of the document or application window.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- height: Specifies the height of the window (in points).
- horizontalPercentScrolled: Specifies the horizontal scroll position as a percentage of the document width.
- imemode: Specifies the default start-up mode for the Japanese Input Method Editor (IME).
- index: Gets the position of an item in a collection.
- isActive: Specifies whether the window is active.
- isDocumentMapVisible: Specifies whether the document map is visible.
- isEnvelopeVisible: Specifies whether the email message header is visible in the document window. The default value is False.
- isHorizontalScrollBarDisplayed: Specifies whether a horizontal scroll bar is displayed for the window.
- isLeftScrollBarDisplayed: Specifies whether the vertical scroll bar appears on the left side of the document window.
- isRightRulerDisplayed: Specifies whether the vertical ruler appears on the right side of the document window in print layout view.
- isSplit: Specifies whether the window is split into multiple panes.
- isVerticalRulerDisplayed: Specifies whether a vertical ruler is displayed for the window or pane.
- isVerticalScrollBarDisplayed: Specifies whether a vertical scroll bar is displayed for the window.
- isVisible: Specifies whether the window is visible.
- left: Specifies the horizontal position of the window, measured in points.
- next: Gets the next document window in the collection of open document windows.
- panes: Gets the collection of panes in the window.
- previous: Gets the previous document window in the collection open document windows.
- showSourceDocuments: Specifies how Microsoft Word displays source documents after a compare and merge process.
- splitVertical: Specifies the vertical split percentage for the window.
- styleAreaWidth: Specifies the width of the style area in points.
- top: Specifies the vertical position of the document window, in points.
- type: Gets the window type.
- usableHeight: Gets the height (in points) of the active working area in the document window.
- usableWidth: Gets the width (in points) of the active working area in the document window.
- verticalPercentScrolled: Specifies the vertical scroll position as a percentage of the document length.
- view: Gets the View object that represents the view for the window.
- width: Specifies the width of the document window, in points.
- windowNumber: Gets an integer that represents the position of the window.
- windowState: Specifies the state of the document window or task window.

## Methods
- activate(): Activates the window.
- close(options): Closes the window.
- largeScroll(options): Scrolls the window by the specified number of screens.
- load(options): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- pageScroll(options): Scrolls through the window page by page.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- setFocus(): Sets the focus of the document window to the body of an email message.
- smallScroll(options): Scrolls the window by the specified number of lines. A "line" corresponds to the distance scrolled by clicking the scroll arrow on the scroll bar once.
- toggleRibbon(): Shows or hides the ribbon.
- toJSON(): Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify().
- track(): Track the object for automatic adjustment based on surrounding changes in the document.
- untrack(): Release the memory associated with this object, if it has previously been tracked.

## Property Details

### activePane
Gets the active pane in the window.

```typescript
readonly activePane: Word.Pane;
```

Type: Word.Pane

Remarks
[API set: WordApiDesktop 1.2]

#### Examples
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

### areRulersDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether rulers are displayed for the window or pane.

```typescript
areRulersDisplayed: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### areScreenTipsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.

```typescript
readonly areScreenTipsDisplayed: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### areThumbnailsDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.

```typescript
areThumbnailsDisplayed: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### caption
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the caption text for the window that is displayed in the title bar of the document or application window.

```typescript
caption: string;
```

Type: string

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Type: Word.RequestContext

### height
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the height of the window (in points).

```typescript
readonly height: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### horizontalPercentScrolled
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal scroll position as a percentage of the document width.

```typescript
horizontalPercentScrolled: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### imemode
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the default start-up mode for the Japanese Input Method Editor (IME).

```typescript
imemode: Word.ImeMode | "NoControl" | "On" | "Off" | "Hiragana" | "Katakana" | "KatakanaHalf" | "AlphaFull" | "Alpha" | "HangulFull" | "Hangul";
```

Type: Word.ImeMode | "NoControl" | "On" | "Off" | "Hiragana" | "Katakana" | "KatakanaHalf" | "AlphaFull" | "Alpha" | "HangulFull" | "Hangul"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### index
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of an item in a collection.

```typescript
readonly index: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isActive
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is active.

```typescript
readonly isActive: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isDocumentMapVisible
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document map is visible.

```typescript
isDocumentMapVisible: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isEnvelopeVisible
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the email message header is visible in the document window. The default value is `False`.

```typescript
isEnvelopeVisible: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isHorizontalScrollBarDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a horizontal scroll bar is displayed for the window.

```typescript
isHorizontalScrollBarDisplayed: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isLeftScrollBarDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the vertical scroll bar appears on the left side of the document window.

```typescript
isLeftScrollBarDisplayed: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isRightRulerDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the vertical ruler appears on the right side of the document window in print layout view.

```typescript
isRightRulerDisplayed: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isSplit
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is split into multiple panes.

```typescript
isSplit: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isVerticalRulerDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a vertical ruler is displayed for the window or pane.

```typescript
isVerticalRulerDisplayed: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isVerticalScrollBarDisplayed
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a vertical scroll bar is displayed for the window.

```typescript
isVerticalScrollBarDisplayed: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### isVisible
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is visible.

```typescript
isVisible: boolean;
```

Type: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### left
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal position of the window, measured in points.

```typescript
readonly left: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### next
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the next document window in the collection of open document windows.

```typescript
readonly next: Word.Window;
```

Type: Word.Window

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### panes
Gets the collection of panes in the window.

```typescript
readonly panes: Word.PaneCollection;
```

Type: Word.PaneCollection

Remarks
[API set: WordApiDesktop 1.2]

### previous
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the previous document window in the collection open document windows.

```typescript
readonly previous: Word.Window;
```

Type: Word.Window

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### showSourceDocuments
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies how Microsoft Word displays source documents after a compare and merge process.

```typescript
showSourceDocuments: Word.ShowSourceDocuments | "None" | "Original" | "Revised" | "Both";
```

Type: Word.ShowSourceDocuments | "None" | "Original" | "Revised" | "Both"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### splitVertical
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical split percentage for the window.

```typescript
splitVertical: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### styleAreaWidth
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the style area in points.

```typescript
styleAreaWidth: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### top
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical position of the document window, in points.

```typescript
readonly top: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### type
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the window type.

```typescript
readonly type: Word.WindowType | "Document" | "Template";
```

Type: Word.WindowType | "Document" | "Template"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### usableHeight
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the height (in points) of the active working area in the document window.

```typescript
readonly usableHeight: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### usableWidth
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the width (in points) of the active working area in the document window.

```typescript
readonly usableWidth: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### verticalPercentScrolled
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical scroll position as a percentage of the document length.

```typescript
verticalPercentScrolled: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### view
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the `View` object that represents the view for the window.

```typescript
readonly view: Word.View;
```

Type: Word.View

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### width
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the document window, in points.

```typescript
readonly width: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### windowNumber
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an integer that represents the position of the window.

```typescript
readonly windowNumber: number;
```

Type: number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### windowState
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the state of the document window or task window.

```typescript
windowState: Word.WindowState | "Normal" | "Maximize" | "Minimize";
```

Type: Word.WindowState | "Normal" | "Maximize" | "Minimize"

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

## Method Details

### activate()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Activates the window.

```typescript
activate(): void;
```

Returns: void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### close(options)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Closes the window.

```typescript
close(options?: Word.WindowCloseOptions): void;
```

Parameters
- options: Word.WindowCloseOptions  
  The options that define whether to save changes before closing and whether to route the document.

Returns: void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### largeScroll(options)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Scrolls the window by the specified number of screens.

```typescript
largeScroll(options?: Word.WindowScrollOptions): void;
```

Parameters
- options: Word.WindowScrollOptions  
  The options for scrolling the window by the specified number of screens. If no options are specified, the window is scrolled down one screen.

Returns: void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### load(options)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.WindowLoadOptions): Word.Window;
```

Parameters
- options: Word.Interfaces.WindowLoadOptions  
  Provides options for which properties of the object to load.

Returns: Word.Window

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.Window;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: Word.Window

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.Window;
```

Parameters
- propertyNamesAndPaths: { select?: string; expand?: string; }  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: Word.Window

### pageScroll(options)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Scrolls through the window page by page.

```typescript
pageScroll(options?: Word.WindowPageScrollOptions): void;
```

Parameters
- options: Word.WindowPageScrollOptions  
  The options for scrolling through the window page by page.

Returns: void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### set(properties, options)
Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.WindowUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters
- properties: Word.Interfaces.WindowUpdateData  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: OfficeExtension.UpdateOptions  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)
Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.Window): void;
```

Parameters
- properties: Word.Window

Returns: void

### setFocus()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Sets the focus of the document window to the body of an email message.

```typescript
setFocus(): void;
```

Returns: void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### smallScroll(options)
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Scrolls the window by the specified number of lines. A "line" corresponds to the distance scrolled by clicking the scroll arrow on the scroll bar once.

```typescript
smallScroll(options?: Word.WindowScrollOptions): void;
```

Parameters
- options: Word.WindowScrollOptions  
  The options for scrolling the window by the specified number of lines. If no options are specified, the window is scrolled down by one line.

Returns: void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### toggleRibbon()
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Shows or hides the ribbon.

```typescript
toggleRibbon(): void;
```

Returns: void

Remarks
[API set: WordApi BETA (PREVIEW ONLY)]

### toJSON()
Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.Window` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.WindowData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.WindowData;
```

Returns: Word.Interfaces.WindowData

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.Window;
```

Returns: Word.Window

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.Window;
```

Returns: Word.Window