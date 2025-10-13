# Word.CommentReplyCollection class

Package: https://learn.microsoft.com/en-us/javascript/api/word

Contains a collection of https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply objects. Represents all comment replies in one comment thread.

Extends
- https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientobject

## Remarks
[ https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets API set: WordApi 1.4 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Gets the replies to the first comment in the selected content.
await Word.run(async (context) => {
  const comment: Word.Comment = context.document.getSelection().getComments().getFirstOrNullObject();
  comment.load("replies");
  await context.sync();

  if (comment.isNullObject) {
    console.warn("No comments in the selection, so no replies to get.");
    return;
  }

  const replies: Word.CommentReplyCollection = comment.replies;
  console.log("Replies to the first comment:", replies);
});
```

## Properties
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- items — Gets the loaded child items in this collection.

## Methods
- getFirst() — Gets the first comment reply in the collection. Throws an ItemNotFound error if this collection is empty.
- getFirstOrNullObject() — Gets the first comment reply in the collection. If the collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties *OrNullObject methods and properties.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CommentReplyCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CommentReplyCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.
- untrack() — Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

## Property Details

### context
The request context associated with the object. This connects the add-in's process to the Office host application's process.

```typescript
context: RequestContext;
```

Property Value
- https://learn.microsoft.com/en-us/javascript/api/word/word.requestcontext Word.RequestContext

### items
Gets the loaded child items in this collection.

```typescript
readonly items: Word.CommentReply[];
```

Property Value
- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply Word.CommentReply[]

## Method Details

### getFirst()
Gets the first comment reply in the collection. Throws an ItemNotFound error if this collection is empty.

```typescript
getFirst(): Word.CommentReply;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply Word.CommentReply

Remarks
- [ https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets API set: WordApi 1.4 ]

### getFirstOrNullObject()
Gets the first comment reply in the collection. If the collection is empty, then this method will return an object with its isNullObject property set to true. For further information, see https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model#ornullobject-methods-and-properties *OrNullObject methods and properties.

```typescript
getFirstOrNullObject(): Word.CommentReply;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreply Word.CommentReply

Remarks
- [ https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets API set: WordApi 1.4 ]

### load(options)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(options?: Word.Interfaces.CommentReplyCollectionLoadOptions & Word.Interfaces.CollectionLoadOptions): Word.CommentReplyCollection;
```

Parameters
- options: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.commentreplycollectionloadoptions Word.Interfaces.CommentReplyCollectionLoadOptions & https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.collectionloadoptions Word.Interfaces.CollectionLoadOptions  
  Provides options for which properties of the object to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreplycollection Word.CommentReplyCollection

### load(propertyNames)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNames?: string | string[]): Word.CommentReplyCollection;
```

Parameters
- propertyNames: string | string[]  
  A comma-delimited string or an array of strings that specify the properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreplycollection Word.CommentReplyCollection

### load(propertyNamesAndPaths)
Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.

```typescript
load(propertyNamesAndPaths?: OfficeExtension.LoadOption): Word.CommentReplyCollection;
```

Parameters
- propertyNamesAndPaths: https://learn.microsoft.com/en-us/javascript/api/office/officeextension.loadoption OfficeExtension.LoadOption  
  propertyNamesAndPaths.select is a comma-delimited string that specifies the properties to load, and propertyNamesAndPaths.expand is a comma-delimited string that specifies the navigation properties to load.

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreplycollection Word.CommentReplyCollection

### toJSON()
Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.CommentReplyCollection object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CommentReplyCollectionData) that contains an "items" array with shallow copies of any loaded properties from the collection's items.

```typescript
toJSON(): Word.Interfaces.CommentReplyCollectionData;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.commentreplycollectiondata Word.Interfaces.CommentReplyCollectionData

### track()
Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent collection.

```typescript
track(): Word.CommentReplyCollection;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreplycollection Word.CommentReplyCollection

### untrack()
Release the memory associated with this object, if it has previously been tracked. This call is shorthand for https://learn.microsoft.com/en-us/javascript/api/office/officeextension.clientrequestcontext#office-officeextension-clientrequestcontext-trackedobjects-member context.trackedObjects.remove(thisObject). Having many tracked objects slows down the host application, so please remember to free any objects you add, once you're done using them. You'll need to call context.sync() before the memory release takes effect.

```typescript
untrack(): Word.CommentReplyCollection;
```

Returns
- https://learn.microsoft.com/en-us/javascript/api/word/word.commentreplycollection Word.CommentReplyCollection