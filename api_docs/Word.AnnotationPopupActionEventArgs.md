# Word.AnnotationPopupActionEventArgs interface

Package: [word](/en-us/javascript/api/word)

Represents action information that's passed back on annotation pop-up action event.

## Remarks

[API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Registers event handlers.
await Word.run(async (context) => {
  eventContexts[0] = context.document.onParagraphAdded.add(paragraphChanged);
  eventContexts[1] = context.document.onParagraphChanged.add(paragraphChanged);

  eventContexts[2] = context.document.onAnnotationClicked.add(onClickedHandler);
  eventContexts[3] = context.document.onAnnotationHovered.add(onHoveredHandler);
  eventContexts[4] = context.document.onAnnotationInserted.add(onInsertedHandler);
  eventContexts[5] = context.document.onAnnotationRemoved.add(onRemovedHandler);
  eventContexts[6] = context.document.onAnnotationPopupAction.add(onPopupActionHandler);

  await context.sync();

  console.log("Event handlers registered.");
});

...

async function onPopupActionHandler(args: Word.AnnotationPopupActionEventArgs) {
  await Word.run(async (context) => {
    let message = `AnnotationPopupAction: ID ${args.id} = `;
    if (args.action === "Accept") {
      message += `Accepted: ${args.critiqueSuggestion}`;
    } else {
      message += "Rejected";
    }

    console.log(message);
  });
}
```

## Properties

- action: Specifies the chosen action in the pop-up menu.
- critiqueSuggestion: Specifies the accepted suggestion (only populated when accepting a critique suggestion).
- id: Specifies the annotation ID for which the event was fired.

## Property Details

### action

Specifies the chosen action in the pop-up menu.

```typescript
action: string;
```

Property Value: string

Remarks: [API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### critiqueSuggestion

Specifies the accepted suggestion (only populated when accepting a critique suggestion).

```typescript
critiqueSuggestion: string;
```

Property Value: string

Remarks: [API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### id

Specifies the annotation ID for which the event was fired.

```typescript
id: string;
```

Property Value: string

Remarks: [API set: WordApi 1.8](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)