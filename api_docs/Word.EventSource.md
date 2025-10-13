# Word.EventSource enum

Package: [word](/en-us/javascript/api/word)

An enum that specifies an event's source. It can be local or remote (through coauthoring).

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

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

async function onEventHandler(event: Word.CommentEventArgs) {
  // Handler for all events except onCommentChanged.
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. Event source: ${event.source}. Comment info:`, event.commentDetails);
  });
}
```

## Fields

- local = "Local"
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- remote = "Remote"
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)