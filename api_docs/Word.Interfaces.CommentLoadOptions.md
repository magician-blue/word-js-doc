# Word.Interfaces.CommentLoadOptions interface

- Package: [word](/en-us/javascript/api/word)

Represents a comment in the document.

## Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- authorEmail  
  Gets the email of the comment's author.

- authorName  
  Gets the name of the comment's author.

- content  
  Specifies the comment's content as plain text.

- contentRange  
  Specifies the comment's content range.

- creationDate  
  Gets the creation date of the comment.

- id  
  Gets the ID of the comment.

- resolved  
  Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value
- boolean

---

### authorEmail

Gets the email of the comment's author.

```typescript
authorEmail?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### authorName

Gets the name of the comment's author.

```typescript
authorName?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### content

Specifies the comment's content as plain text.

```typescript
content?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contentRange

Specifies the comment's content range.

```typescript
contentRange?: Word.Interfaces.CommentContentRangeLoadOptions;
```

Property Value
- [Word.Interfaces.CommentContentRangeLoadOptions](/en-us/javascript/api/word/word.interfaces.commentcontentrangeloadoptions)

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### creationDate

Gets the creation date of the comment.

```typescript
creationDate?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

Gets the ID of the comment.

```typescript
id?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### resolved

Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

```typescript
resolved?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)