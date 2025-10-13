# Word.Interfaces.CommentData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `comment.toJSON()`.

## Properties

- authorEmail: Gets the email of the comment's author.
- authorName: Gets the name of the comment's author.
- content: Specifies the comment's content as plain text.
- contentRange: Specifies the comment's content range.
- creationDate: Gets the creation date of the comment.
- id: Gets the ID of the comment.
- replies: Gets the collection of reply objects associated with the comment.
- resolved: Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

## Property Details

### authorEmail

Gets the email of the comment's author.

TypeScript
```typescript
authorEmail?: string;
```

Property Value
- string

Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### authorName

Gets the name of the comment's author.

TypeScript
```typescript
authorName?: string;
```

Property Value
- string

Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### content

Specifies the comment's content as plain text.

TypeScript
```typescript
content?: string;
```

Property Value
- string

Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contentRange

Specifies the comment's content range.

TypeScript
```typescript
contentRange?: Word.Interfaces.CommentContentRangeData;
```

Property Value
- [Word.Interfaces.CommentContentRangeData](/en-us/javascript/api/word/word.interfaces.commentcontentrangedata)

Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### creationDate

Gets the creation date of the comment.

TypeScript
```typescript
creationDate?: Date;
```

Property Value
- Date

Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

Gets the ID of the comment.

TypeScript
```typescript
id?: string;
```

Property Value
- string

Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### replies

Gets the collection of reply objects associated with the comment.

TypeScript
```typescript
replies?: Word.Interfaces.CommentReplyData[];
```

Property Value
- [Word.Interfaces.CommentReplyData](/en-us/javascript/api/word/word.interfaces.commentreplydata)[]

Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### resolved

Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

TypeScript
```typescript
resolved?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)