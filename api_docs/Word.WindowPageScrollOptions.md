# Word.WindowPageScrollOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The options for scrolling through the specified pane or window page by page.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- down: If provided, specifies the number of pages to scroll the window down. If down and up are both provided, the contents of the window are scrolled by the difference of the property values. For example, if down is 3 and up is 6, the contents are scrolled up three pages.
- up: If provided, specifies the number of pages to scroll the window up. If down and up are both provided, the contents of the window are scrolled by the difference of the property values. For example, if down is 3 and up is 6, the contents are scrolled up three pages.

## Property Details

### down

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the number of pages to scroll the window down. If down and up are both provided, the contents of the window are scrolled by the difference of the property values. For example, if down is 3 and up is 6, the contents are scrolled up three pages.

```typescript
down?: number;
```

Property Value
- number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### up

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the number of pages to scroll the window up. If down and up are both provided, the contents of the window are scrolled by the difference of the property values. For example, if down is 3 and up is 6, the contents are scrolled up three pages.

```typescript
up?: number;
```

Property Value
- number

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)