# Word.ParagraphDeletedEventArgs interface

Package: [word](/en-us/javascript/api/word)

Provides information about the paragraphs that raised the paragraphDeleted event.

## Remarks

[API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/25-paragraph/ondeleted-event.yaml

// Registers the onParagraphDeleted event handler on the document.
await Word.run(async (context) => {
  eventContext = context.document.onParagraphDeleted.add(paragraphDeleted);
  await context.sync();

  console.log("Added event handlers for when paragraphs are deleted.");
});

...

async function paragraphDeleted(event: Word.ParagraphDeletedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.type} event detected. IDs of paragraphs that were deleted:`, event.uniqueLocalIds);
  });
}
```

## Properties

- `source`: The source of the event. It can be local or remote (through coauthoring).
- `type`: The event type. See Word.EventType for details.
- `uniqueLocalIds`: Gets the unique IDs of the involved paragraphs. IDs are in standard 8-4-4-4-12 GUID format without curly braces and differ across sessions and coauthors.

## Property Details

### source

The source of the event. It can be local or remote (through coauthoring).

```typescript
source: Word.EventSource | "Local" | "Remote";
```

Property Value: [Word.EventSource](/en-us/javascript/api/word/word.eventsource) | "Local" | "Remote"

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

The event type. See Word.EventType for details.

```typescript
type: Word.EventType | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction";
```

Property Value: [Word.EventType](/en-us/javascript/api/word/word.eventtype) | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction"

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### uniqueLocalIds

Gets the unique IDs of the involved paragraphs. IDs are in standard 8-4-4-4-12 GUID format without curly braces and differ across sessions and coauthors.

```typescript
uniqueLocalIds: string[];
```

Property Value: string[]

Remarks: [API set: WordApi 1.6](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)