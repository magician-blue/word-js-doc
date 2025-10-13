# Word.Interfaces.CustomXmlPrefixMappingLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a CustomXmlPrefixMapping object.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- `$all`: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- `namespaceUri`: Gets the unique address identifier for the namespace of the `CustomXmlPrefixMapping` object.
- `prefix`: Gets the prefix for the `CustomXmlPrefixMapping` object.

## Property Details

### $all

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

### namespaceUri

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the unique address identifier for the namespace of the `CustomXmlPrefixMapping` object.

```typescript
namespaceUri?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### prefix

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the prefix for the `CustomXmlPrefixMapping` object.

```typescript
prefix?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)