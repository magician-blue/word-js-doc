# Word.Interfaces.IndexUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the Index object, for use in index.set({ ... }).

## Properties

- range: Returns a Range object that represents the portion of the document that is contained within the index.
- tabLeader: Specifies the leader character between entries in the index and their associated page numbers.

## Property Details

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Range object that represents the portion of the document that is contained within the index.

```typescript
range?: Word.Interfaces.RangeUpdateData;
```

Property Value
- [Word.Interfaces.RangeUpdateData](/en-us/javascript/api/word/word.interfaces.rangeupdatedata)

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tabLeader

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the leader character between entries in the index and their associated page numbers.

```typescript
tabLeader?: Word.TabLeader | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot";
```

Property Value
- [Word.TabLeader](/en-us/javascript/api/word/word.tableader) | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot"

Remarks
- [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)