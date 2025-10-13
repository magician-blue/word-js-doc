# Word.CommentReply class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents a comment reply in the document.

Extends: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks

[API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Replies to the first active comment in the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("reply-text") as HTMLInputElement).value;
  const comments: Word.CommentCollection = context.document.getSelection().getComments();
  comments.load("items");
  await context.sync();

  const firstActiveComment: Word.Comment = comments.items.find((item) => item.resolved !== true);
  if (firstActiveComment) {
    const reply: Word.CommentReply = firstActiveComment.reply(text);
    console.log("Reply added.");
  } else {
    console.warn("No active comment was found in the selection, so couldn't reply.");
  }
});
```

## Properties

- authorEmail: Gets the email of the comment reply's author.
- authorName: Gets the name of the comment reply's author.
- content: Specifies the comment reply's content. The string is plain text.
- contentRange: Specifies the commentReply's content range.
- context: The request context associated with the object. This connects the add-in's process to the Office host application's process.
- creationDate: Gets the creation date of the comment reply.
- id: Gets the ID of the comment reply.
- parentComment: Gets the parent comment of this reply.

## Methods

- delete(): Deletes the comment reply.
- load(options): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNames): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- load(propertyNamesAndPaths): Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.
- set(properties, options): Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties): Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON(): Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CommentReply` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentReplyData`) that contains shallow copies of any loaded child properties from the original object.
- track(): Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack(): Release the memory associated with this object, if it has previously been tracked. This call is shorthand for context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

## Property Details

### authorEmail

Gets the email of the comment reply's author.

```typescript
readonly authorEmail: string;
```

Property Value: string

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### authorName

Gets the name of the comment reply's author.

```typescript
readonly authorName: string;
```

Property Value: string

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### content

Specifies the comment reply's content. The string is plain text.

```typescript
content: string;
```

Property Value: string

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### contentRange

Specifies the commentReply's content range.

```typescript
contentRange: Word.CommentContentRange;
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.commentcontentrange

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext

### creationDate

Gets the creation date of the comment reply.

```typescript
readonly creationDate: Date;
```

Property Value: Date

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### id

Gets the ID of the comment reply.

```typescript
readonly id: string;
```

Property Value: string

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### parentComment

Gets the parent comment of this reply.

```typescript
readonly parentComment: Word.Comment;
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.comment

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Method Details

### delete()

Deletes the comment reply.

```typescript
delete(): void;
```

Returns: void

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### load(options)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(options?: Word.Interfaces.CommentReplyLoadOptions): Word.CommentReply;
```

Parameters:
- options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.commentreplyloadoptions  
  Provides options for which properties of the object to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CommentReply;
```

Parameters:
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call `context.sync()` before reading the properties.

```typescript
load(propertyNamesAndPaths?: {
            select?: string;
            expand?: string;
        }): Word.CommentReply;
```

Parameters:
- propertyNamesAndPaths: `{ select?: string; expand?: string; }`  
  `propertyNamesAndPaths.select` is a comma-delimited string that specifies the properties to load, and `propertyNamesAndPaths.expand` is a comma-delimited string that specifies the navigation properties to load.

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply

### set(properties, options)

Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.

```typescript
set(properties: Interfaces.CommentReplyUpdateData, options?: OfficeExtension.UpdateOptions): void;
```

Parameters:
- properties: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.commentreplyupdatedata  
  A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
- options: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.updateoptions  
  Provides an option to suppress errors if the properties object tries to set any read-only properties.

Returns: void

### set(properties)

Sets multiple properties on the object at the same time, based on an existing loaded object.

```typescript
set(properties: Word.CommentReply): void;
```

Parameters:
- properties: https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply

Returns: void

### toJSON()

Overrides the JavaScript `toJSON()` method in order to provide more useful output when an API object is passed to `JSON.stringify()`. (`JSON.stringify`, in turn, calls the `toJSON` method of the object that's passed to it.) Whereas the original `Word.CommentReply` object is an API object, the `toJSON` method returns a plain JavaScript object (typed as `Word.Interfaces.CommentReplyData`) that contains shallow copies of any loaded child properties from the original object.

```typescript
toJSON(): Word.Interfaces.CommentReplyData;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.commentreplydata

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across `.sync` calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CommentReply;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call `context.sync()` before the memory release takes effect.

```typescript
untrack(): Word.CommentReply;
```

Returns: https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply