# Word.ContentControlEnteredEventArgs interface

Package: [word](/en-us/javascript/api/word)

Provides information about the content control that raised contentControlEntered event.

## Remarks

[ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/content-control-onentered-event.yaml

await Word.run(async (context) => {
  const contentControls: Word.ContentControlCollection = context.document.contentControls;
  contentControls.load("items");
  await context.sync();

  // Register the onEntered event handler on each content control.
  if (contentControls.items.length === 0) {
    console.log("There aren't any content controls in this document so can't register event handlers.");
  } else {
    for (let i = 0; i < contentControls.items.length; i++) {
      eventContexts[i] = contentControls.items[i].onEntered.add(contentControlEntered);
      contentControls.items[i].track();
    }

    await context.sync();

    console.log("Added event handlers for when the cursor is placed in content controls.");
  }
});

...

async function contentControlEntered(event: Word.ContentControlEnteredEventArgs) {
  await Word.run(async (context) => {
    console.log(`${event.eventType} event detected. ID of content control that was entered: ${event.ids[0]}`);
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

#### Property Value

[Word.EventType](/en-us/javascript/api/word/word.eventtype) | "ContentControlDeleted" | "ContentControlSelectionChanged" | "ContentControlDataChanged" | "ContentControlAdded" | "CommentDeleted" | "CommentSelected" | "CommentDeselected" | "CommentChanged" | "CommentAdded" | "ContentControlEntered" | "ContentControlExited" | "ParagraphAdded" | "ParagraphChanged" | "ParagraphDeleted" | "AnnotationClicked" | "AnnotationHovered" | "AnnotationInserted" | "AnnotationRemoved" | "AnnotationPopupAction"

#### Remarks

[ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### ids

Gets the content control IDs.

```typescript
ids: number[];
```

#### Property Value

number[]

#### Remarks

[ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### source

The source of the event. It can be local or remote (through coauthoring).

```typescript
source: Word.EventSource | "Local" | "Remote";
```

#### Property Value

[Word.EventSource](/en-us/javascript/api/word/word.eventsource) | "Local" | "Remote"

#### Remarks

[ API set: WordApi 1.5 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)