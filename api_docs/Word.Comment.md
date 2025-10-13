# Word.Comment class

Package: [word](/en-us/javascript/api/word)

Represents a comment in the document.

Extends: [OfficeExtension.ClientObject](/en-us/javascript/api/office/officeextension.clientobject)

## Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// Sets a comment on the selected content.
await Word.run(async (context) => {
  const text = (document.getElementById("comment-text") as HTMLInputElement).value;
  const comment: Word.Comment = context.document.getSelection().insertComment(text);

  // Load object to log in the console.
  comment.load();
  await context.sync();

  console.log("Comment inserted:", comment);
});
```

## Properties
- authorEmail — Gets the email of the comment's author.
- authorName — Gets the name of the comment's author.
- content — Specifies the comment's content as plain text.
- contentRange — Specifies the comment's content range.
- context — The request context associated with the object. This connects the add-in's process to the Office host application's process.
- creationDate — Gets the creation date of the comment.
- id — Gets the ID of the comment.
- replies — Gets the collection of reply objects associated with the comment.
- resolved — Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.

## Methods
- delete() — Deletes the comment and its replies.
- getRange() — Gets the range in the main document where the comment is on.
- load(options) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNames) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- load(propertyNamesAndPaths) — Queues up a command to load the specified properties of the object. You must call context.sync() before reading the properties.
- reply(replyText) — Adds a new reply to the end of the comment thread.
- set(properties, options) — Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
- set(properties) — Sets multiple properties on the object at the same time, based on an existing loaded object.
- toJSON() — Overrides the JavaScript toJSON() method in order to provide more useful output when an API object is passed to JSON.stringify(). (JSON.stringify, in turn, calls the toJSON method of the object that's passed to it.) Whereas the original Word.Comment object is an API object, the toJSON method returns a plain JavaScript object (typed as Word.Interfaces.CommentData) that contains shallow copies of any loaded child properties from the original object.
- track() — Track the object for automatic adjustment based on surrounding changes in the document. This call is a shorthand for context.trackedObjects.add(thisObject). If you're using this object across .sync calls and outside the sequential execution of a ".run" batch, and get an "InvalidObjectPath" error when setting a property or invoking a method on the object, you need to add the object to the tracked object collection when the object was first created. If this object is part of a collection, you should also track the parent 