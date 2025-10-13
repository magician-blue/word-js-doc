# Word.CommentEventArgs interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Provides information about the comments that raised the comment event.

## Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

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

async function onChangedHandler(event: Word.CommentEventArgs) {
  await Word.run(async (context) => {
    console.log(
      `${event.type} event detected. ${event.changeType} change made. Event source: ${event.source}. Comment info:`, event.commentDetails
    );
  });
}
```

## Properties

- changeType  
  Represents how the comment changed event is triggered.

- commentDetails  
  Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.

- source  
  The source of the event. It can be local or remote (through coauthoring).

- type  
  The event type. See Word.EventType for details.

## Property Details

### changeType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents how the comment changed event is triggered.

```typescript
changeType: Word.CommentChangeType | "none" | "edited" | "resolved" | "reopened" | "replyAdded" | "replyDeleted" | "replyEdited";
```

Property Value  
[Word.CommentChangeType](https://learn.microsoft.com/en-us/javascript/api/word/word.commentchangetype) | "none" | "edited" | "resolved" | "reopened" | "replyAdded" | "replyDeleted" | "replyEdited"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### commentDetails

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the CommentDetail array which contains the IDs and reply IDs of the involved comments.

```typescript
commentDetails: Word.CommentDetail[];
```

Property Value  
[Word.CommentDetail](https://learn.microsoft.com/en-us/javascript/api/word/word.commentdetail)[]

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### source

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The source of the event. It can be local or remote (through coauthoring).

```typescript
source: Word.EventSource | "Local" | "Remote";
```

Property Value  
[Word.EventSource](https://learn.microsoft.com/en-us/javascript/api/word/word.eventsource) | "Local" | "Remote"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

The event type. See Word.EventType for details.

```typescript
type: Word.EventType | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction";
```

Property Value  
[Word.EventType](https://learn.microsoft.com/en-us/javascript/api/word/word.eventtype) | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction"

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)