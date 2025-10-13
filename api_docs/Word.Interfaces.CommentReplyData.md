# Word.Interfaces.CommentReplyData interface

- Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `commentReply.toJSON()`.

## Properties

- authorEmail — Gets the email of the comment reply's author.
- authorName — Gets the name of the comment reply's author.
- content — Specifies the comment reply's content. The string is plain text.
- contentRange — Specifies the commentReply's content range.
- creationDate — Gets the creation date of the comment reply.
- id — Gets the ID of the comment reply.
- parentComment — Gets the parent comment of this reply.

## Property Details

### authorEmail

Gets the email of the comment reply's author.

```typescript
authorEmail?: string;
```

#### Property Value
string

#### Remarks
[ [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### authorName

Gets the name of the comment reply's author.

```typescript
authorName?: string;
```

#### Property Value
string

#### Remarks
[ [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### content

Specifies the comment reply's content. The string is plain text.

```typescript
content?: string;
```

#### Property Value
string

#### Remarks
[ [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### contentRange

Specifies the commentReply's content range.

```typescript
contentRange?: Word.Interfaces.CommentContentRangeData;
```

#### Property Value
[Word.Interfaces.CommentContentRangeData](/en-us/javascript/api/word/word.interfaces.commentcontentrangedata)

#### Remarks
[ [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### creationDate

Gets the creation date of the comment reply.

```typescript
creationDate?: Date;
```

#### Property Value
Date

#### Remarks
[ [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### id

Gets the ID of the comment reply.

```typescript
id?: string;
```

#### Property Value
string

#### Remarks
[ [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

---

### parentComment

Gets the parent comment of this reply.

```typescript
parentComment?: Word.Interfaces.CommentData;
```

#### Property Value
[Word.Interfaces.CommentData](/en-us/javascript/api/word/word.interfaces.commentdata)

#### Remarks
[ [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]