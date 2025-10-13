# Word.Interfaces.WindowLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the window that displays the document. A window can be split to contain multiple reading panes.

## Remarks

[ API set: WordApiDesktop 1.2 ]

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- areRulersDisplayed  
  Specifies whether rulers are displayed for the window or pane.

- areScreenTipsDisplayed  
  Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.

- areThumbnailsDisplayed  
  Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.

- caption  
  Specifies the caption text for the window that is displayed in the title bar of the document or application window.

- height  
  Specifies the height of the window (in points).

- horizontalPercentScrolled  
  Specifies the horizontal scroll position as a percentage of the document width.

- imemode  
  Specifies the default start-up mode for the Japanese Input Method Editor (IME).

- index  
  Gets the position of an item in a collection.

- isActive  
  Specifies whether the window is active.

- isDocumentMapVisible  
  Specifies whether the document map is visible.

- isEnvelopeVisible  
  Specifies whether the email message header is visible in the document window. The default value is False.

- isHorizontalScrollBarDisplayed  
  Specifies whether a horizontal scroll bar is displayed for the window.

- isLeftScrollBarDisplayed  
  Specifies whether the vertical scroll bar appears on the left side of the document window.

- isRightRulerDisplayed  
  Specifies whether the vertical ruler appears on the right side of the document window in print layout view.

- isSplit  
  Specifies whether the window is split into multiple panes.

- isVerticalRulerDisplayed  
  Specifies whether a vertical ruler is displayed for the window or pane.

- isVerticalScrollBarDisplayed  
  Specifies whether a vertical scroll bar is displayed for the window.

- isVisible  
  Specifies whether the window is visible.

- left  
  Specifies the horizontal position of the window, measured in points.

- next  
  Gets the next document window in the collection of open document windows.

- previous  
  Gets the previous document window in the collection open document windows.

- showSourceDocuments  
  Specifies how Microsoft Word displays source documents after a compare and merge process.

- splitVertical  
  Specifies the vertical split percentage for the window.

- styleAreaWidth  
  Specifies the width of the style area in points.

- top  
  Specifies the vertical position of the document window, in points.

- type  
  Gets the window type.

- usableHeight  
  Gets the height (in points) of the active working area in the document window.

- usableWidth  
  Gets the width (in points) of the active working area in the document window.

- verticalPercentScrolled  
  Specifies the vertical scroll position as a percentage of the document length.

- view  
  Gets the View object that represents the view for the window.

- width  
  Specifies the width of the document window, in points.

- windowNumber  
  Gets an integer that represents the position of the window.

- windowState  
  Specifies the state of the document window or task window.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property value  
boolean

---

### areRulersDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether rulers are displayed for the window or pane.

```typescript
areRulersDisplayed?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### areScreenTipsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.

```typescript
areScreenTipsDisplayed?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### areThumbnailsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.

```typescript
areThumbnailsDisplayed?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### caption

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the caption text for the window that is displayed in the title bar of the document or application window.

```typescript
caption?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### height

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the height of the window (in points).

```typescript
height?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### horizontalPercentScrolled

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal scroll position as a percentage of the document width.

```typescript
horizontalPercentScrolled?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### imemode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the default start-up mode for the Japanese Input Method Editor (IME).

```typescript
imemode?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### index

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of an item in a collection.

```typescript
index?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isActive

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is active.

```typescript
isActive?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isDocumentMapVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document map is visible.

```typescript
isDocumentMapVisible?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isEnvelopeVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the email message header is visible in the document window. The default value is `False`.

```typescript
isEnvelopeVisible?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isHorizontalScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a horizontal scroll bar is displayed for the window.

```typescript
isHorizontalScrollBarDisplayed?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isLeftScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the vertical scroll bar appears on the left side of the document window.

```typescript
isLeftScrollBarDisplayed?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isRightRulerDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the vertical ruler appears on the right side of the document window in print layout view.

```typescript
isRightRulerDisplayed?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isSplit

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is split into multiple panes.

```typescript
isSplit?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isVerticalRulerDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a vertical ruler is displayed for the window or pane.

```typescript
isVerticalRulerDisplayed?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isVerticalScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a vertical scroll bar is displayed for the window.

```typescript
isVerticalScrollBarDisplayed?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is visible.

```typescript
isVisible?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### left

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal position of the window, measured in points.

```typescript
left?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### next

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the next document window in the collection of open document windows.

```typescript
next?: Word.Interfaces.WindowLoadOptions;
```

Property value  
[Word.Interfaces.WindowLoadOptions](/en-us/javascript/api/word/word.interfaces.windowloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### previous

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the previous document window in the collection open document windows.

```typescript
previous?: Word.Interfaces.WindowLoadOptions;
```

Property value  
[Word.Interfaces.WindowLoadOptions](/en-us/javascript/api/word/word.interfaces.windowloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### showSourceDocuments

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies how Microsoft Word displays source documents after a compare and merge process.

```typescript
showSourceDocuments?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### splitVertical

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical split percentage for the window.

```typescript
splitVertical?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### styleAreaWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the style area in points.

```typescript
styleAreaWidth?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### top

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical position of the document window, in points.

```typescript
top?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the window type.

```typescript
type?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### usableHeight

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the height (in points) of the active working area in the document window.

```typescript
usableHeight?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### usableWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the width (in points) of the active working area in the document window.

```typescript
usableWidth?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### verticalPercentScrolled

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical scroll position as a percentage of the document length.

```typescript
verticalPercentScrolled?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### view

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the `View` object that represents the view for the window.

```typescript
view?: Word.Interfaces.ViewLoadOptions;
```

Property value  
[Word.Interfaces.ViewLoadOptions](/en-us/javascript/api/word/word.interfaces.viewloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### width

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the document window, in points.

```typescript
width?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### windowNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an integer that represents the position of the window.

```typescript
windowNumber?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### windowState

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the state of the document window or task window.

```typescript
windowState?: boolean;
```

Property value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]