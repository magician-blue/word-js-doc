# Word.Interfaces.WindowCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents the collection of window objects.

## Remarks

[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- areRulersDisplayed: For EACH ITEM in the collection: Specifies whether rulers are displayed for the window or pane.
- areScreenTipsDisplayed: For EACH ITEM in the collection: Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.
- areThumbnailsDisplayed: For EACH ITEM in the collection: Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.
- caption: For EACH ITEM in the collection: Specifies the caption text for the window that is displayed in the title bar of the document or application window.
- height: For EACH ITEM in the collection: Specifies the height of the window (in points).
- horizontalPercentScrolled: For EACH ITEM in the collection: Specifies the horizontal scroll position as a percentage of the document width.
- imemode: For EACH ITEM in the collection: Specifies the default start-up mode for the Japanese Input Method Editor (IME).
- index: For EACH ITEM in the collection: Gets the position of an item in a collection.
- isActive: For EACH ITEM in the collection: Specifies whether the window is active.
- isDocumentMapVisible: For EACH ITEM in the collection: Specifies whether the document map is visible.
- isEnvelopeVisible: For EACH ITEM in the collection: Specifies whether the email message header is visible in the document window. The default value is `False`.
- isHorizontalScrollBarDisplayed: For EACH ITEM in the collection: Specifies whether a horizontal scroll bar is displayed for the window.
- isLeftScrollBarDisplayed: For EACH ITEM in the collection: Specifies whether the vertical scroll bar appears on the left side of the document window.
- isRightRulerDisplayed: For EACH ITEM in the collection: Specifies whether the vertical ruler appears on the right side of the document window in print layout view.
- isSplit: For EACH ITEM in the collection: Specifies whether the window is split into multiple panes.
- isVerticalRulerDisplayed: For EACH ITEM in the collection: Specifies whether a vertical ruler is displayed for the window or pane.
- isVerticalScrollBarDisplayed: For EACH ITEM in the collection: Specifies whether a vertical scroll bar is displayed for the window.
- isVisible: For EACH ITEM in the collection: Specifies whether the window is visible.
- left: For EACH ITEM in the collection: Specifies the horizontal position of the window, measured in points.
- next: For EACH ITEM in the collection: Gets the next document window in the collection of open document windows.
- previous: For EACH ITEM in the collection: Gets the previous document window in the collection open document windows.
- showSourceDocuments: For EACH ITEM in the collection: Specifies how Microsoft Word displays source documents after a compare and merge process.
- splitVertical: For EACH ITEM in the collection: Specifies the vertical split percentage for the window.
- styleAreaWidth: For EACH ITEM in the collection: Specifies the width of the style area in points.
- top: For EACH ITEM in the collection: Specifies the vertical position of the document window, in points.
- type: For EACH ITEM in the collection: Gets the window type.
- usableHeight: For EACH ITEM in the collection: Gets the height (in points) of the active working area in the document window.
- usableWidth: For EACH ITEM in the collection: Gets the width (in points) of the active working area in the document window.
- verticalPercentScrolled: For EACH ITEM in the collection: Specifies the vertical scroll position as a percentage of the document length.
- view: For EACH ITEM in the collection: Gets the `View` object that represents the view for the window.
- width: For EACH ITEM in the collection: Specifies the width of the document window, in points.
- windowNumber: For EACH ITEM in the collection: Gets an integer that represents the position of the window.
- windowState: For EACH ITEM in the collection: Specifies the state of the document window or task window.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### areRulersDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether rulers are displayed for the window or pane.

```typescript
areRulersDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areScreenTipsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.

```typescript
areScreenTipsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### areThumbnailsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.

```typescript
areThumbnailsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### caption

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the caption text for the window that is displayed in the title bar of the document or application window.

```typescript
caption?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### height

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the height of the window (in points).

```typescript
height?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### horizontalPercentScrolled

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the horizontal scroll position as a percentage of the document width.

```typescript
horizontalPercentScrolled?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### imemode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the default start-up mode for the Japanese Input Method Editor (IME).

```typescript
imemode?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### index

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the position of an item in a collection.

```typescript
index?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isActive

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the window is active.

```typescript
isActive?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isDocumentMapVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the document map is visible.

```typescript
isDocumentMapVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isEnvelopeVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the email message header is visible in the document window. The default value is `False`.

```typescript
isEnvelopeVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isHorizontalScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether a horizontal scroll bar is displayed for the window.

```typescript
isHorizontalScrollBarDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isLeftScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the vertical scroll bar appears on the left side of the document window.

```typescript
isLeftScrollBarDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isRightRulerDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the vertical ruler appears on the right side of the document window in print layout view.

```typescript
isRightRulerDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isSplit

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the window is split into multiple panes.

```typescript
isSplit?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isVerticalRulerDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether a vertical ruler is displayed for the window or pane.

```typescript
isVerticalRulerDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isVerticalScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether a vertical scroll bar is displayed for the window.

```typescript
isVerticalScrollBarDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies whether the window is visible.

```typescript
isVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### left

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the horizontal position of the window, measured in points.

```typescript
left?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### next

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the next document window in the collection of open document windows.

```typescript
next?: Word.Interfaces.WindowLoadOptions;
```

Property Value: [Word.Interfaces.WindowLoadOptions](/en-us/javascript/api/word/word.interfaces.windowloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### previous

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the previous document window in the collection open document windows.

```typescript
previous?: Word.Interfaces.WindowLoadOptions;
```

Property Value: [Word.Interfaces.WindowLoadOptions](/en-us/javascript/api/word/word.interfaces.windowloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### showSourceDocuments

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies how Microsoft Word displays source documents after a compare and merge process.

```typescript
showSourceDocuments?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### splitVertical

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the vertical split percentage for the window.

```typescript
splitVertical?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### styleAreaWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the width of the style area in points.

```typescript
styleAreaWidth?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### top

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the vertical position of the document window, in points.

```typescript
top?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the window type.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### usableHeight

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the height (in points) of the active working area in the document window.

```typescript
usableHeight?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### usableWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the width (in points) of the active working area in the document window.

```typescript
usableWidth?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### verticalPercentScrolled

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the vertical scroll position as a percentage of the document length.

```typescript
verticalPercentScrolled?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### view

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the `View` object that represents the view for the window.

```typescript
view?: Word.Interfaces.ViewLoadOptions;
```

Property Value: [Word.Interfaces.ViewLoadOptions](/en-us/javascript/api/word/word.interfaces.viewloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### width

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the width of the document window, in points.

```typescript
width?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### windowNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets an integer that represents the position of the window.

```typescript
windowNumber?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### windowState

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies the state of the document window or task window.

```typescript
windowState?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)