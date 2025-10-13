# Word.Interfaces.CommentReplyUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the `CommentReply` object, for use in `commentReply.set({ ... })`.

## Properties

- [content](#content) — Specifies the comment reply's content. The string is plain text.
- [contentRange](#contentrange) — Specifies the commentReply's content range.
- [parentComment](#parentcomment) — Gets the parent comment of this reply.

## Property details

### content

Specifies the comment reply's content. The string is plain text.

```typescript
content?: string;
```

Property value: string

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### contentRange

Specifies the commentReply's content range.

```typescript
contentRange?: Word.Interfaces.CommentContentRangeUpdateData;
```

Property value: [Word.Interfaces.CommentContentRangeUpdateData](/en-us/javascript/api/word/word.interfaces.commentcontentrangeupdatedata)

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### parentComment

Gets the parent comment of this reply.

```typescript
parentComment?: Word.Interfaces.CommentUpdateData;
```

Property value: [Word.Interfaces.CommentUpdateData](/en-us/javascript/api/word/word.interfaces.commentupdatedata)

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)