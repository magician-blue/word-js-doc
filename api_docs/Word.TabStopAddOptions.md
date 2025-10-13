# Word.TabStopAddOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the options for adding to a [Word.TabStopCollection](/en-us/javascript/api/word/word.tabstopcollection) object.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [alignment](#alignment)  
  If provided, specifies the alignment of the tab stop. The default value is `left`.

- [leader](#leader)  
  If provided, specifies the leader character for the tab stop. The default value is `spaces`.

## Property Details

### alignment

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the alignment of the tab stop. The default value is `left`.

```typescript
alignment?: Word.TabAlignment | "Left" | "Center" | "Right" | "Decimal" | "Bar" | "List";
```

Property Value: [Word.TabAlignment](/en-us/javascript/api/word/word.tabalignment) | "Left" | "Center" | "Right" | "Decimal" | "Bar" | "List"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### leader

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the leader character for the tab stop. The default value is `spaces`.

```typescript
leader?: Word.TabLeader | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot";
```

Property Value: [Word.TabLeader](/en-us/javascript/api/word/word.tableader) | "Spaces" | "Dots" | "Dashes" | "Lines" | "Heavy" | "MiddleDot"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)