# Word.AnnotationState enum

Package: [word](/en-us/javascript/api/word)

Represents the state of the annotation.

## Remarks

[ API set: WordApi 1.7 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-annotations.yaml

// Accepts the first annotation found in the selected paragraph.
await Word.run(async (context) => {
  const paragraph: Word.Paragraph = context.document.getSelection().paragraphs.getFirst();
  const annotations: Word.AnnotationCollection = paragraph.getAnnotations();
  annotations.load("id,state,critiqueAnnotation");

  await context.sync();

  for (let i = 0; i < annotations.items.length; i++) {
    const annotation: Word.Annotation = annotations.items[i];

    if (annotation.state === Word.AnnotationState.created) {
      console.log(`Accepting ID ${annotation.id}...`);
      annotation.critiqueAnnotation.accept();

      await context.sync();
      break;
    }
  }
});
```

## Fields

- accepted = "Accepted"
  - Accepted.
  - [ API set: WordApi 1.7 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- created = "Created"
  - Created.
  - [ API set: WordApi 1.7 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- rejected = "Rejected"
  - Rejected.
  - [ API set: WordApi 1.7 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)