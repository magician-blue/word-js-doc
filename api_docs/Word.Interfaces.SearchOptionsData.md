# Word.Interfaces.SearchOptionsData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `searchOptions.toJSON()`.

## Properties

- ignorePunct  
  Specifies a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.

- ignoreSpace  
  Specifies a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.

- matchCase  
  Specifies a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.

- matchPrefix  
  Specifies a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.

- matchSuffix  
  Specifies a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.

- matchWholeWord  
  Specifies a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.

- matchWildcards  
  Specifies a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.

## Property Details

### ignorePunct

Specifies a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.

```typescript
ignorePunct?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### ignoreSpace

Specifies a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.

```typescript
ignoreSpace?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### matchCase

Specifies a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.

```typescript
matchCase?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### matchPrefix

Specifies a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.

```typescript
matchPrefix?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### matchSuffix

Specifies a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.

```typescript
matchSuffix?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### matchWholeWord

Specifies a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.

```typescript
matchWholeWord?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### matchWildcards

Specifies a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.

```typescript
matchWildcards?: boolean;
```

- Property value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)