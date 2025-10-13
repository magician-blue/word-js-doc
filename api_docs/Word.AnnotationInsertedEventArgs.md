# Word.AnnotationInsertedEventArgs interface

Package: [word](/en-us/javascript/api/word)

Holds annotation information that is passed back on annotation added event.

## Remarks

[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

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

async function onInsertedHandler(args: Word.AnnotationInsertedEventArgs) {
  await Word.run(async (context) => {
    const annotations = [];
    for (let i = 0; i < args.ids.length; i++) {
      let annotation: Word.Annotation = context.document.getAnnotationById(args.ids[i]);
      annotation.load("id,critiqueAnnotation");

      annotations.push(annotation);
    }

    await context.sync();

    for (let annotation of annotations) {
      console.log(`AnnotationInserted: ID ${annotation.id}:`, annotation.critiqueAnnotation.critique);
    }
  });
}
```

## Properties

- ids: Specifies the annotation IDs for which the event was fired.

## Property Details

### ids

Specifies the annotation IDs for which the event was fired.

```typescript
ids: string[];
```

Property Value: string[]

Remarks: [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)