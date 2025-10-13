# Word.Interfaces.SearchOptionsLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Specifies the options to be included in a search operation. To learn more about how to use search options in the Word JavaScript APIs, read [Use search options to find text in your Word add-in](/en-us/office/dev/add-ins/word/search-option-guidance).

## Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

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

### $all
Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### ignorePunct
Specifies a value that indicates whether to ignore all punctuation characters between words. Corresponds to the Ignore punctuation check box in the Find and Replace dialog box.

```typescript
ignorePunct?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### ignoreSpace
Specifies a value that indicates whether to ignore all whitespace between words. Corresponds to the Ignore whitespace characters check box in the Find and Replace dialog box.

```typescript
ignoreSpace?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### matchCase
Specifies a value that indicates whether to perform a case sensitive search. Corresponds to the Match case check box in the Find and Replace dialog box.

```typescript
matchCase?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### matchPrefix
Specifies a value that indicates whether to match words that begin with the search string. Corresponds to the Match prefix check box in the Find and Replace dialog box.

```typescript
matchPrefix?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### matchSuffix
Specifies a value that indicates whether to match words that end with the search string. Corresponds to the Match suffix check box in the Find and Replace dialog box.

```typescript
matchSuffix?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### matchWholeWord
Specifies a value that indicates whether to find operation only entire words, not text that is part of a larger word. Corresponds to the Find whole words only check box in the Find and Replace dialog box.

```typescript
matchWholeWord?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### matchWildcards
Specifies a value that indicates whether the search will be performed using special search operators. Corresponds to the Use wildcards check box in the Find and Replace dialog box.

```typescript
matchWildcards?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)