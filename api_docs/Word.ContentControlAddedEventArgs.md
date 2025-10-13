# Word.ContentControlAddedEventArgs interface

Package: [word](/en-us/javascript/api/word)

Provides information about the content control that raised contentControlAdded event.

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-onadded-event.yaml

// Registers the onAdded event handler on the document.
await Word.run(async (context) => {
  eventContext = context.document.onContentControlAdded.add(contentControlAdded);
  await context.sync();

  console.log("Added event handler for when content controls are added.");
});

...

async function contentControlAdded(event: Word.ContentControlAddedEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. IDs of content controls that were added:`, event.ids);
  });
}
```

## Properties

- [eventType](#eventtype): The event type. See Word.EventType for details.
- [ids](#ids): Gets the content control IDs.
- [source](#source): The source of the event. It can be local or remote (through coauthoring).

## Property Details

### eventType

The event type. See Word.EventType for details.

```typescript
eventType: Word.EventType | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction";
```

#### Property Value

[Word.EventType](/en-us/javascript/api/word/word.eventtype) | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction"

#### Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### ids

Gets the content control IDs.

```typescript
ids: number[];
```

#### Property Value

number[]

#### Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### source

The source of the event. It can be local or remote (through coauthoring).

```typescript
source: Word.EventSource | "Local" | "Remote";
```

#### Property Value

[Word.EventSource](/en-us/javascript/api/word/word.eventsource) | "Local" | "Remote"

#### Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)