# Word.ContentControlDataChangedEventArgs interface

Package: [word](/en-us/javascript/api/word)

Provides information about the content control that raised contentControlDataChanged event.

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-ondatachanged-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  // Register the onDataChanged event handler on each content control.
  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onDataChanged.add(contentControlDataChanged);
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when data is changed in content controls.");
  }
});

...

async function contentControlDataChanged(event: Word.ContentControlDataChangedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls where data was changed:`, event.ids);
  });
}
```

## Properties

- eventType: The event type. See Word.EventType for details.
- ids: Gets the content control IDs.
- source: The source of the event. It can be local or remote (through coauthoring).

## Property Details

### eventType

The event type. See Word.EventType for details.

```TypeScript
eventType: Word.EventType | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction";
```

Property value:
[Word.EventType](/en-us/javascript/api/word/word.eventtype) | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction"

Remarks:
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### ids

Gets the content control IDs.

```TypeScript
ids: number[];
```

Property value:
number[]

Remarks:
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### source

The source of the event. It can be local or remote (through coauthoring).

```TypeScript
source: Word.EventSource | "Local" | "Remote";
```

Property value:
[Word.EventSource](/en-us/javascript/api/word/word.eventsource) | "Local" | "Remote"

Remarks:
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)