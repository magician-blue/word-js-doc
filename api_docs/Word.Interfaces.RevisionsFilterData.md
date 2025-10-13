# Word.Interfaces.RevisionsFilterData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `revisionsFilter.toJSON()`.

## Properties

- markup  
  Specifies a `RevisionsMarkup` value that represents the extent of reviewer markup displayed in the document.
- view  
  Specifies a `RevisionsView` value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.

## Property Details

### markup

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `RevisionsMarkup` value that represents the extent of reviewer markup displayed in the document.

```typescript
markup?: Word.RevisionsMarkup | "None" | "Simple" | "All";
```

Property Value  
[Word.RevisionsMarkup](/en-us/javascript/api/word/word.revisionsmarkup) | "None" | "Simple" | "All"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

### view

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `RevisionsView` value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.

```typescript
view?: Word.RevisionsView | "Final" | "Original";
```

Property Value  
[Word.RevisionsView](/en-us/javascript/api/word/word.revisionsview) | "Final" | "Original"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)