# Word.IndexMarkEntryOptions interface

- Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents options for marking an index entry in a Word document.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- bold  
  If provided, specifies whether to add bold formatting to page numbers for index entries. The default value is `false`.

- bookmarkName  
  If provided, specifies the bookmark name that marks the range of pages you want to appear in the index. If this property is omitted, the number of the page that contains the `XE` field appears in the index. The default value is "".

- crossReference  
  If provided, specifies the cross-reference that will appear in the index. The default value is "".

- crossReferenceAutoText  
  If provided, specifies the name of the `AutoText` entry that contains the text for a cross-reference (if this property is specified, `crossReference` is ignored). The default value is "".

- entry  
  If provided, specifies the text you want to appear in the index, in the form `MainEntry[:Subentry]`. The default value is "". Either this property or `entryAutoText` must be provided.

- entryAutoText  
  If provided, specifies the `AutoText` entry that contains the text you want to appear in the index (if this property is specified, `entry` is ignored). The default value is "". Either this property or `entry` must be provided.

- italic  
  If provided, specifies whether to add italic formatting to page numbers for index entries. The default value is `false`.

- reading  
  If provided, specifies whether to show an index entry in the right location when indexes are sorted phonetically (East Asian languages only). The default value is `false`.

## Property Details

### bold

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies whether to add bold formatting to page numbers for index entries. The default value is `false`.

```typescript
bold?: boolean;
```

- Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### bookmarkName

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the bookmark name that marks the range of pages you want to appear in the index. If this property is omitted, the number of the page that contains the `XE` field appears in the index. The default value is "".

```typescript
bookmarkName?: string;
```

- Property Value: string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### crossReference

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the cross-reference that will appear in the index. The default value is "".

```typescript
crossReference?: string;
```

- Property Value: string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### crossReferenceAutoText

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the name of the `AutoText` entry that contains the text for a cross-reference (if this property is specified, `crossReference` is ignored). The default value is "".

```typescript
crossReferenceAutoText?: string;
```

- Property Value: string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### entry

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the text you want to appear in the index, in the form `MainEntry[:Subentry]`. The default value is "". Either this property or `entryAutoText` must be provided.

```typescript
entry?: string;
```

- Property Value: string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### entryAutoText

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies the `AutoText` entry that contains the text you want to appear in the index (if this property is specified, `entry` is ignored). The default value is "". Either this property or `entry` must be provided.

```typescript
entryAutoText?: string;
```

- Property Value: string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### italic

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies whether to add italic formatting to page numbers for index entries. The default value is `false`.

```typescript
italic?: boolean;
```

- Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### reading

Note  
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

If provided, specifies whether to show an index entry in the right location when indexes are sorted phonetically (East Asian languages only). The default value is `false`.

```typescript
reading?: boolean;
```

- Property Value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)