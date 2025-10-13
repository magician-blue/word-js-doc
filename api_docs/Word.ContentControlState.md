# Word.ContentControlState enum

Package: [word](/en-us/javascript/api/word)

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the state of the content control.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/99-preview-apis/insert-and-change-content-controls.yaml

// Sets the state of the first content control.
await Word.run(async (context) => {
  const state = ((document.getElementById("state-to-set") as HTMLSelectElement)
    .value as unknown) as Word.ContentControlState;
  let firstContentControl = context.document.contentControls.getFirstOrNullObject();
  await context.sync();

  if (firstContentControl.isNullObject) {
    console.warn("There are no content controls in this document.");
    return;
  }

  firstContentControl.setState(state);
  firstContentControl.load("id");
  await context.sync();

  console.log(`Set state of first content control with ID ${firstContentControl.id} to ${state}.`);
});
```

## Fields

- error = "Error"
  - Error state.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- warning = "Warning"
  - Warning state.
  - [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)