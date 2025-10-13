# Word.Interfaces.PaneData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling pane.toJSON().

## Properties

- [pages](#pages): Gets the collection of pages in the pane.
- [pagesEnclosingViewport](#pagesenclosingviewport): Gets the PageCollection shown in the viewport of the pane. If a page is partially visible in the pane, the whole page is returned.

## Property Details

### pages

Gets the collection of pages in the pane.

```typescript
pages?: Word.Interfaces.PageData[];
```

Property Value
- [Word.Interfaces.PageData](/en-us/javascript/api/word/word.interfaces.pagedata)[]

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### pagesEnclosingViewport

Gets the PageCollection shown in the viewport of the pane. If a page is partially visible in the pane, the whole page is returned.

```typescript
pagesEnclosingViewport?: Word.Interfaces.PageData[];
```

Property Value
- [Word.Interfaces.PageData](/en-us/javascript/api/word/word.interfaces.pagedata)[]

Remarks
- [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)