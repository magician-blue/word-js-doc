# Word.CommentChangeType enum

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents how the comments in the event were changed.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/manage-comments.yaml

// Registers event handlers.
await Word.run(async (context) => {
  const body: Word.Body = context.document.body;
  body.track();
  await context.sync();

  eventContexts[0] = body.onCommentAdded.add(onEventHandler);
  eventContexts[1] = body.onCommentChanged.add(onChangedHandler);
  eventContexts[2] = body.onCommentDeleted.add(onEventHandler);
  eventContexts[3] = body.onCommentDeselected.add(onEventHandler);
  eventContexts[4] = body.onCommentSelected.add(onEventHandler);
  await context.sync();

  console.log("Event handlers registered.");
});

...

async function onChangedHandler(event: Word.CommentEventArgs) {
  await Word.run(async (context) => {
    console.log(
      `${event.type} event detected. ${event.changeType} change made. Event source: ${event.source}. Comment info:`, event.commentDetails
    );
  });
}
```

## Fields

- edited = "edited"
  - A comment was edited.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- none = "none"
  - No comment changed event is triggered.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- reopened = "reopened"
  - A comment was reopened.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- replyAdded = "replyAdded"
  - A reply was added.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- replyDeleted = "replyDeleted"
  - A reply was deleted.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- replyEdited = "replyEdited"
  - A reply was edited.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- resolved = "resolved"
  - A comment was resolved.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)