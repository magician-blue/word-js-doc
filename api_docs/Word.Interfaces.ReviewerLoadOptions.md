# Word.Interfaces.ReviewerLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single reviewer of a document in which changes have been tracked. The Reviewer object is a member of the [Word.ReviewerCollection](/en-us/javascript/api/word/word.reviewercollection) object.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- isVisible: Specifies if the `Reviewer` object is visible.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

### isVisible

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the `Reviewer` object is visible.

```typescript
isVisible?: boolean;
```

Property Value
- boolean

Remarks
- [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)