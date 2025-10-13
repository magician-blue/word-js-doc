# Word.Interfaces.RevisionsFilterLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the current settings related to the display of reviewers' comments and revision marks in the document.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- markup  
  Specifies a `RevisionsMarkup` value that represents the extent of reviewer markup displayed in the document.

- view  
  Specifies a `RevisionsView` value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### markup

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `RevisionsMarkup` value that represents the extent of reviewer markup displayed in the document.

```typescript
markup?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)

---

### view

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `RevisionsView` value that represents globally whether Word displays the original version of the document or the final version, which might have revisions and formatting changes applied.

```typescript
view?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/office/dev/add-ins/reference/overview/visio-javascript-reference-overview)