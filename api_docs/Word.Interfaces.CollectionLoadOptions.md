# Word.Interfaces.CollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Provides ways to load properties of only a subset of members of a collection.

## Properties

- $skip
  - Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.

- $top
  - Specify the number of items in the queried collection to be included in the result.

## Property Details

### $skip

Specify the number of items in the collection that are to be skipped and not included in the result. If top is specified, the selection of result will start after skipping the specified number of items.

```typescript
$skip?: number;
```

- Property Value: number

### $top

Specify the number of items in the queried collection to be included in the result.

```typescript
$top?: number;
```

- Property Value: number