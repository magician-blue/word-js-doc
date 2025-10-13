# Word.Interfaces.CommentCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Comment](/en-us/javascript/api/word/word.comment) objects.

## Remarks

- [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- authorEmail  
  For EACH ITEM in the collection: Gets the email of the comment's author.

- authorName  
  For EACH ITEM in the collection: Gets the name of the comment's author.

- content  
  For EACH ITEM in the collection: Specifies the comment's content as plain text.

- contentRange  
  For EACH ITEM in the collection: Specifies the comment's content range.

- creationDate  
  For EACH ITEM in the collection: Gets the creation date of the comment.

- id  
  For EACH ITEM in the collection: Gets the ID of the comment.

- resolved  
  For EACH ITEM in the collection: Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### authorEmail

For EACH ITEM in the collection: Gets the email of the comment's author.

```typescript
authorEmail?: boolean;
```

Property Value: boolean

Remarks

- [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### authorName

For EACH ITEM in the collection: Gets the name of the comment's author.

```typescript
authorName?: boolean;
```

Property Value: boolean

Remarks

- [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### content

For EACH ITEM in the collection: Specifies the comment's content as plain text.

```typescript
content?: boolean;
```

Property Value: boolean

Remarks

- [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contentRange

For EACH ITEM in the collection: Specifies the comment's content range.

```typescript
contentRange?: Word.Interfaces.CommentContentRangeLoadOptions;
```

Property Value: [Word.Interfaces.CommentContentRangeLoadOptions](/en-us/javascript/api/word/word.interfaces.commentcontentrangeloadoptions)

Remarks

- [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### creationDate

For EACH ITEM in the collection: Gets the creation date of the comment.

```typescript
creationDate?: boolean;
```

Property Value: boolean

Remarks

- [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

For EACH ITEM in the collection: Gets the ID of the comment.

```typescript
id?: boolean;
```

Property Value: boolean

Remarks

- [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### resolved

For EACH ITEM in the collection: Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

```typescript
resolved?: boolean;
```

Property Value: boolean

Remarks

- [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)