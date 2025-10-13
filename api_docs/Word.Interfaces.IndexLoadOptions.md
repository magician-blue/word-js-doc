# Word.Interfaces.IndexLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents a single index. The `Index` object is a member of the [Word.IndexCollection](/en-us/javascript/api/word/word.indexcollection). The `IndexCollection` includes all the indexes in the document.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- filter  
  Gets a value that represents how Microsoft Word classifies the first character of entries in the index. See `IndexFilter` for available values.

- headingSeparator  
  Gets the text between alphabetical groups (entries that start with the same letter) in the index. Corresponds to the \h switch for an [INDEX field](https://support.microsoft.com/office/adafcf4a-cb30-43f6-85c7-743da1635d9e).

- indexLanguage  
  Gets a `LanguageId` value that represents the sorting language to use for the index.

- numberOfColumns  
  Gets the number of columns for each page of the index.

- range  
  Returns a `Range` object that represents the portion of the document that is contained within the index.

- rightAlignPageNumbers  
  Specifies if page numbers are aligned with the right margin in the index.

- separateAccentedLetterHeadings  
  Gets if the index contains separate headings for accented letters (for example, words that begin with "Ã" are under one heading and words that begin with "A" are under another).

- sortBy  
  Specifies the sorting criteria for the index.

- tabLeader  
  Specifies the leader character between entries in the index and their associated page numbers.

- type  
  Gets the index type.

## Property Details

### $all

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### filter

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a value that represents how Microsoft Word classifies the first character of entries in the index. See `IndexFilter` for available values.

```typescript
filter?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### headingSeparator

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the text between alphabetical groups (entries that start with the same letter) in the index. Corresponds to the \h switch for an [INDEX field](https://support.microsoft.com/office/adafcf4a-cb30-43f6-85c7-743da1635d9e).

```typescript
headingSeparator?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### indexLanguage

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `LanguageId` value that represents the sorting language to use for the index.

```typescript
indexLanguage?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### numberOfColumns

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the number of columns for each page of the index.

```typescript
numberOfColumns?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Range` object that represents the portion of the document that is contained within the index.

```typescript
range?: Word.Interfaces.RangeLoadOptions;
```

Property Value: [Word.Interfaces.RangeLoadOptions](/en-us/javascript/api/word/word.interfaces.rangeloadoptions)

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### rightAlignPageNumbers

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if page numbers are aligned with the right margin in the index.

```typescript
rightAlignPageNumbers?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### separateAccentedLetterHeadings

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets if the index contains separate headings for accented letters (for example, words that begin with "Ã" are under one heading and words that begin with "A" are under another).

```typescript
separateAccentedLetterHeadings?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### sortBy

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the sorting criteria for the index.

```typescript
sortBy?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tabLeader

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the leader character between entries in the index and their associated page numbers.

```typescript
tabLeader?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the index type.

```typescript
type?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)