# Word.ContentControlDeletedEventArgs interface

Package: [word](/en-us/javascript/api/word)

Provides information about the content control that raised contentControlDeleted event.

## Remarks

[ API set: WordApi 1.5 ]

#### Examples
```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-ondeleted-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  // Register the onDeleted event handler on each content control.
  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onDeleted.add(contentControlDeleted);
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when content controls are deleted.");
  }
});

...

async function contentControlDeleted(event: Word.ContentControlDeletedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls that were deleted:`, event.ids);
  });
}
```

## Properties

- eventType  
  The event type. See Word.EventType for details.

- ids  
  Gets the content control IDs.

- source  
  The source of the event. It can be local or remote (through coauthoring).

## Property Details

### eventType

The event type. See Word.EventType for details.

```typescript
eventType: Word.EventType | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction";
```

Property value: [Word.EventType](/en-us/javascript/api/word/word.eventtype) | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction"

Remarks  
[ API set: WordApi 1.5 ]

### ids

Gets the content control IDs.

```typescript
ids: number[];
```

Property value: number[]

Remarks  
[ API set: WordApi 1.5 ]

### source

The source of the event. It can be local or remote (through coauthoring).

```typescript
source: Word.EventSource | "Local" | "Remote";
```

Property value: [Word.EventSource](/en-us/javascript/api/word/word.eventsource) | "Local" | "Remote"

Remarks  
[ API set: WordApi 1.5 ]