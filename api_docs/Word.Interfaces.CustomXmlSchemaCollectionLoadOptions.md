# Word.Interfaces.CustomXmlSchemaCollectionLoadOptions interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a collection of Word.CustomXmlSchema objects attached to a data stream.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- location: For EACH ITEM in the collection: Gets the location of the schema on a computer.
- namespaceUri: For EACH ITEM in the collection: Gets the unique address identifier for the namespace of the `CustomXmlSchema` object.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

### location

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the location of the schema on a computer.

```typescript
location?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

### namespaceUri

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the unique address identifier for the namespace of the `CustomXmlSchema` object.

```typescript
namespaceUri?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]