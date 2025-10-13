# Word.ParagraphChangedEventArgs interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Provides information about the paragraphs that raised the paragraphChanged event.

## Remarks

[ [API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/onchanged-event.yaml

// Registers the onParagraphChanged event handler on the document.
await Word.run(async (context) => {
  eventContext = context.document.onParagraphChanged.add(paragraphChanged);
  await context.sync();

  console.log("Added event handler for when content is changed in paragraphs.");
});

...

async function paragraphChanged(event: Word.ParagraphChangedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. IDs of paragraphs where content was changed:`, event.uniqueLocalIds);
  });
}
```

## Properties

- source: The source of the event. It can be local or remote (through coauthoring).
- type: The event type. See Word.EventType for details.
- uniqueLocalIds: Gets the unique IDs of the involved paragraphs. IDs are in standard 8-4-4-4-12 GUID format without curly braces and differ across sessions and coauthors.

## Property Details

### source

The source of the event. It can be local or remote (through coauthoring).

```typescript
source: Word.EventSource | "Local" | "Remote";
```

Property Value:
[Word.EventSource](https://learn.microsoft.com/en-us/javascript/api/word/word.eventsource) | "Local" | "Remote"

Remarks:
[ [API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### type

The event type. See Word.EventType for details.

```typescript
type: Word.EventType | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction";
```

Property Value:
[Word.EventType](https://learn.microsoft.com/en-us/javascript/api/word/word.eventtype) | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction"

Remarks:
[ [API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### uniqueLocalIds

Gets the unique IDs of the involved paragraphs. IDs are in standard 8-4-4-4-12 GUID format without curly braces and differ across sessions and coauthors.

```typescript
uniqueLocalIds: string[];
```

Property Value:
string[]

Remarks:
[ [API set: WordApi 1.6](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]