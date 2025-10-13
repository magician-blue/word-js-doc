# Word.Interfaces.SourceCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of [Word.Source](/en-us/javascript/api/word/word.source) objects.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- isCited: For EACH ITEM in the collection: Gets if the `Source` object has been cited in the document.
- tag: For EACH ITEM in the collection: Gets the tag of the source.
- xml: For EACH ITEM in the collection: Gets the XML representation of the source.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property value: boolean

### isCited

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets if the `Source` object has been cited in the document.

```typescript
isCited?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tag

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the tag of the source.

```typescript
tag?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xml

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the XML representation of the source.

```typescript
xml?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)