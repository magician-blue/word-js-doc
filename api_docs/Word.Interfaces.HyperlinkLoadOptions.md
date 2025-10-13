# Word.Interfaces.HyperlinkLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a hyperlink in a Word document.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- address — Specifies the address (for example, a file name or URL) of the hyperlink.
- emailSubject — Specifies the text string for the hyperlink's subject line.
- isExtraInfoRequired — Returns `true` if extra information is required to resolve the hyperlink.
- name — Returns the name of the `Hyperlink` object.
- range — Returns a `Range` object that represents the portion of the document that's contained within the hyperlink.
- screenTip — Specifies the text that appears as a ScreenTip when the mouse pointer is positioned over the hyperlink.
- subAddress — Specifies a named location in the destination of the hyperlink.
- target — Specifies the name of the frame or window in which to load the hyperlink.
- textToDisplay — Specifies the hyperlink's visible text in the document.
- type — Returns the hyperlink type.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property value: boolean

### address

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the address (for example, a file name or URL) of the hyperlink.

```typescript
address?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### emailSubject

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text string for the hyperlink's subject line.

```typescript
emailSubject?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isExtraInfoRequired

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns `true` if extra information is required to resolve the hyperlink.

```typescript
isExtraInfoRequired?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### name

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the name of the `Hyperlink` object.

```typescript
name?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Range` object that represents the portion of the document that's contained within the hyperlink.

```typescript
range?: Word.Interfaces.RangeLoadOptions;
```

Property value: [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### screenTip

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text that appears as a ScreenTip when the mouse pointer is positioned over the hyperlink.

```typescript
screenTip?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### subAddress

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a named location in the destination of the hyperlink.

```typescript
subAddress?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### target

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the name of the frame or window in which to load the hyperlink.

```typescript
target?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### textToDisplay

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the hyperlink's visible text in the document.

```typescript
textToDisplay?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the hyperlink type.

```typescript
type?: boolean;
```

Property value: boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)