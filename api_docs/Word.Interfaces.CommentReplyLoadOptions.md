# Word.Interfaces.CommentReplyLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a comment reply in the document.

## Remarks

[ API set: WordApi 1.4 ]

## Properties

- $all: Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- authorEmail: Gets the email of the comment reply's author.
- authorName: Gets the name of the comment reply's author.
- content: Specifies the comment reply's content. The string is plain text.
- contentRange: Specifies the commentReply's content range.
- creationDate: Gets the creation date of the comment reply.
- id: Gets the ID of the comment reply.
- parentComment: Gets the parent comment of this reply.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value: boolean

### authorEmail

Gets the email of the comment reply's author.

```typescript
authorEmail?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### authorName

Gets the name of the comment reply's author.

```typescript
authorName?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### content

Specifies the comment reply's content. The string is plain text.

```typescript
content?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### contentRange

Specifies the commentReply's content range.

```typescript
contentRange?: Word.Interfaces.CommentContentRangeLoadOptions;
```

Property Value: [Word.Interfaces.CommentContentRangeLoadOptions](/en-us/javascript/api/word/word.interfaces.commentcontentrangeloadoptions)

Remarks  
[ API set: WordApi 1.4 ]

### creationDate

Gets the creation date of the comment reply.

```typescript
creationDate?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### id

Gets the ID of the comment reply.

```typescript
id?: boolean;
```

Property Value: boolean

Remarks  
[ API set: WordApi 1.4 ]

### parentComment

Gets the parent comment of this reply.

```typescript
parentComment?: Word.Interfaces.CommentLoadOptions;
```

Property Value: [Word.Interfaces.CommentLoadOptions](/en-us/javascript/api/word/word.interfaces.commentloadoptions)

Remarks  
[ API set: WordApi 1.4 ]