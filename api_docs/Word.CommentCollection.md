# Word.CommentCollection class

Contains a collection of [Word.Comment](/en-us/javascript/api/word/word.comment) objects.

- Package: [word](/en-us/javascript/api/word)
- Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

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

- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items — Gets the loaded child items in this collection.

## Methods

- getFirst() — Gets the first comment in the collection. Throws an ItemNotFound error if this collection is empty.
- getFirstOrNullObject() — Gets the first comment in the collection. If the collection is empty, returns an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties*](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON() — Overrides the JavaScript toJSON() method to provide more useful output when an API object is passed to JSON.stringify(). Returns a plain JavaScript object (typed as Word.Interfaces.CommentCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. Shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If using this object across .sync calls and outside the sequential execution of a ".run" batch and you get an "InvalidObjectPath" error, add the object to the tracked object collection when first created. If this object is part of a collection, also track the parent collection.
- untrack() — Release the memory associated with this object if previously tracked. Shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so free any objects you add once you're done using them. Call context.sync() before the memory release takes effect.

## Property Details

### context

The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

- Property Value: [Word.RequestContext](/en-us/javascript/api/word/word.requestcontext)

### items

Gets the loaded child items in this collection.

```typescript
readonly items: Word.Comment[];
```

- Property Value: [Word.Comment](/en-us/javascript/api/word/word.comment)[]

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

## Method Details

### getFirst()

Gets the first comment in the collection. Throws an ItemNotFound error if this collection is empty.

```typescript
getFirst(): Word.Comment;
```

- Returns: [Word.Comment](/en-us/javascript/api/word/word.comment)

Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### getFirstOrNullObject()

Gets the first comment in the collection. If the collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see [*OrNullObject methods and properties*](/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties).

```typescript
getFirstOrNullObject(): Word.Comment;
```

- Returns: [Word.Comment](/en-us/javascript/api/word/word.comment)

Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the range of the first comment in the selected content.
await Word.run(async (context) => {
  const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
  comment.load("contentRange");
  const range: Word.Range = comment.getRange();
  range.load("text");
  await context.sync();

  if (comment.isNullObject) {
    console.warn("No comments in the selection, so no range to get.");
    return;
  }

  console.log(`Comment location: ${range.text}`);
  const contentRange: Word.CommentContentRange = comment.contentRange;
  console.log("Comment content range:", contentRange);
});
```

### load(options)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.CommentCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CommentCollection;
```

Parameters

- options: [Word.Interfaces.CommentCollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.commentcollectionloadoptions) & [Word.Interfaces.CollectionLoadOptions](/en-us/javascript/api/word/word.interfaces.collectionloadoptions)

Provides options for which properties of the object to load.

- Returns: [Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)

### load(propertyNames)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CommentCollection;
```

Parameters

- propertyNames: string | string[]

A comma-delimited string or an array of strings that specify the properties to load.

- Returns: [Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)

### load(propertyNamesAndPaths)

Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CommentCollection;
```

Parameters

- propertyNamesAndPaths: [OfficeExtension.LoadOption](/en-us/javascript/api/office/officeextension.loadoption)

propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

- Returns: [Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)

### toJSON()

Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CommentCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CommentCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.CommentCollectionData;
```

- Returns: [Word.Interfaces.CommentCollectionData](/en-us/javascript/api/word/word.interfaces.commentcollectiondata)

### track()

Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for [context.trackedObjects.add(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CommentCollection;
```

- Returns: [Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)

### untrack()

Release the memory associated with this object, if it has previously been tracked. This call is shorthand for [context.trackedObjects.remove(thisObject)](/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.CommentCollection;
```

- Returns: [Word.CommentCollection](/en-us/javascript/api/word/word.commentcollection)