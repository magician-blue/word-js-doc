# Word.Interfaces.ReviewerCollectionLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

A collection of [Word.Reviewer](/en-us/javascript/api/word/word.reviewer) objects that represents the reviewers of one or more documents. The ReviewerCollection object contains the names of all reviewers who have reviewed documents opened or edited on a computer.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- isVisible: For EACH ITEM in the collection: Specifies if the `Reviewer` object is visible.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

#### Property Value

boolean

### isVisible

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if the `Reviewer` object is visible.

```typescript
isVisible?: boolean;
```

#### Property Value

boolean

#### Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)