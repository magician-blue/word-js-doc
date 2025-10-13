# Word.Interfaces.CommentUpdateData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface for updating data on the Comment object, for use in comment.set({ ... }).

## Properties

- content  
  Specifies the comment's content as plain text.

- contentRange  
  Specifies the comment's content range.

- resolved  
  Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

## Property Details

### content

Specifies the comment's content as plain text.

```typescript
content?: string;
```

#### Property Value
string

#### Remarks
[ API set: WordApi 1.4 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contentRange

Specifies the comment's content range.

```typescript
contentRange?: Word.Interfaces.CommentContentRangeUpdateData;
```

#### Property Value
[Word.Interfaces.CommentContentRangeUpdateData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.commentcontentrangeupdatedata)

#### Remarks
[ API set: WordApi 1.4 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### resolved

Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

```typescript
resolved?: boolean;
```

#### Property Value
boolean

#### Remarks
[ API set: WordApi 1.4 ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)