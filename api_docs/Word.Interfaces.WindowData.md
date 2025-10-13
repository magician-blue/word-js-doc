# Word.Interfaces.WindowData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling window.toJSON().

## Properties

- areRulersDisplayed: Specifies whether rulers are displayed for the window or pane.
- areScreenTipsDisplayed: Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.
- areThumbnailsDisplayed: Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.
- caption: Specifies the caption text for the window that is displayed in the title bar of the document or application window.
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
- previous: Gets the previous document window in the collection open document windows.
- showSourceDocuments: Specifies how Microsoft Word displays source documents after a compare and merge process.
- splitVertical: Specifies the vertical split percentage for the window.
- styleAreaWidth: Specifies the width of the style area in points.
- top: Specifies the vertical position of the document window, in points.
- type: Gets the window type.
- usableHeight: Gets the height (in points) of the active working area in the document window.
- usableWidth: Gets the width (in points) of the active working area in the document window.
- verticalPercentScrolled: Specifies the vertical scroll position as a percentage of the document length.
- width: Specifies the width of the document window, in points.
- windowNumber: Gets an integer that represents the position of the window.
- windowState: Specifies the state of the document window or task window.

## Property Details

### areRulersDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether rulers are displayed for the window or pane.

```typescript
areRulersDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areScreenTipsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether comments, footnotes, endnotes, and hyperlinks are displayed as tips.

```typescript
areScreenTipsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### areThumbnailsDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether thumbnail images of the pages in a document are displayed along the left side of the Microsoft Word document window.

```typescript
areThumbnailsDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### caption

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the caption text for the window that is displayed in the title bar of the document or application window.

```typescript
caption?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### height

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the height of the window (in points).

```typescript
height?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### horizontalPercentScrolled

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal scroll position as a percentage of the document width.

```typescript
horizontalPercentScrolled?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### imemode

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the default start-up mode for the Japanese Input Method Editor (IME).

```typescript
imemode?: Word.ImeMode | "NoControl" | "On" | "Off" | "Hiragana" | "Katakana" | "KatakanaHalf" | "AlphaFull" | "Alpha" | "HangulFull" | "Hangul";
```

Property Value: [Word.ImeMode](/en-us/javascript/api/word/word.imemode) | "NoControl" | "On" | "Off" | "Hiragana" | "Katakana" | "KatakanaHalf" | "AlphaFull" | "Alpha" | "HangulFull" | "Hangul"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### index

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of an item in a collection.

```typescript
index?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isActive

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is active.

```typescript
isActive?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isDocumentMapVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the document map is visible.

```typescript
isDocumentMapVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isEnvelopeVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the email message header is visible in the document window. The default value is False.

```typescript
isEnvelopeVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isHorizontalScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a horizontal scroll bar is displayed for the window.

```typescript
isHorizontalScrollBarDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isLeftScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the vertical scroll bar appears on the left side of the document window.

```typescript
isLeftScrollBarDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isRightRulerDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the vertical ruler appears on the right side of the document window in print layout view.

```typescript
isRightRulerDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isSplit

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is split into multiple panes.

```typescript
isSplit?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVerticalRulerDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a vertical ruler is displayed for the window or pane.

```typescript
isVerticalRulerDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVerticalScrollBarDisplayed

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether a vertical scroll bar is displayed for the window.

```typescript
isVerticalScrollBarDisplayed?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether the window is visible.

```typescript
isVisible?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### left

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the horizontal position of the window, measured in points.

```typescript
left?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### next

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the next document window in the collection of open document windows.

```typescript
next?: Word.Interfaces.WindowData;
```

Property Value: [Word.Interfaces.WindowData](/en-us/javascript/api/word/word.interfaces.windowdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### previous

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the previous document window in the collection open document windows.

```typescript
previous?: Word.Interfaces.WindowData;
```

Property Value: [Word.Interfaces.WindowData](/en-us/javascript/api/word/word.interfaces.windowdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### showSourceDocuments

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies how Microsoft Word displays source documents after a compare and merge process.

```typescript
showSourceDocuments?: Word.ShowSourceDocuments | "None" | "Original" | "Revised" | "Both";
```

Property Value: [Word.ShowSourceDocuments](/en-us/javascript/api/word/word.showsourcedocuments) | "None" | "Original" | "Revised" | "Both"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### splitVertical

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical split percentage for the window.

```typescript
splitVertical?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### styleAreaWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the style area in points.

```typescript
styleAreaWidth?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### top

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical position of the document window, in points.

```typescript
top?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the window type.

```typescript
type?: Word.WindowType | "Document" | "Template";
```

Property Value: [Word.WindowType](/en-us/javascript/api/word/word.windowtype) | "Document" | "Template"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### usableHeight

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the height (in points) of the active working area in the document window.

```typescript
usableHeight?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### usableWidth

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the width (in points) of the active working area in the document window.

```typescript
usableWidth?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### verticalPercentScrolled

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the vertical scroll position as a percentage of the document length.

```typescript
verticalPercentScrolled?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the width of the document window, in points.

```typescript
width?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### windowNumber

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an integer that represents the position of the window.

```typescript
windowNumber?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### windowState

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the state of the document window or task window.

```typescript
windowState?: Word.WindowState | "Normal" | "Maximize" | "Minimize";
```

Property Value: [Word.WindowState](/en-us/javascript/api/word/word.windowstate) | "Normal" | "Maximize" | "Minimize"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)