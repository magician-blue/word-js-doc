# Word.Interfaces.CommentReplyCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.CommentReply](/en-us/javascript/api/word/word.commentreply) objects. Represents all comment replies in one comment thread.

## Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- authorEmail — For EACH ITEM in the collection: Gets the email of the comment reply's author.
- authorName — For EACH ITEM in the collection: Gets the name of the comment reply's author.
- content — For EACH ITEM in the collection: Specifies the comment reply's content. The string is plain text.
- contentRange — For EACH ITEM in the collection: Specifies the commentReply's content range.
- creationDate — For EACH ITEM in the collection: Gets the creation date of the comment reply.
- id — For EACH ITEM in the collection: Gets the ID of the comment reply.
- parentComment — For EACH ITEM in the collection: Gets the parent comment of this reply.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### authorEmail

For EACH ITEM in the collection: Gets the email of the comment reply's author.

```typescript
authorEmail?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### authorName

For EACH ITEM in the collection: Gets the name of the comment reply's author.

```typescript
authorName?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### content

For EACH ITEM in the collection: Specifies the comment reply's content. The string is plain text.

```typescript
content?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contentRange

For EACH ITEM in the collection: Specifies the commentReply's content range.

```typescript
contentRange?: Word.Interfaces.CommentContentRangeLoadOptions;
```

Property Value: [Word.Interfaces.CommentContentRangeLoadOptions](/en-us/javascript/api/word/word.interfaces.commentcontentrangeloadoptions)

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### creationDate

For EACH ITEM in the collection: Gets the creation date of the comment reply.

```typescript
creationDate?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

For EACH ITEM in the collection: Gets the ID of the comment reply.

```typescript
id?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### parentComment

For EACH ITEM in the collection: Gets the parent comment of this reply.

```typescript
parentComment?: Word.Interfaces.CommentLoadOptions;
```

Property Value: [Word.Interfaces.CommentLoadOptions](/en-us/javascript/api/word/word.interfaces.commentloadoptions)

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)