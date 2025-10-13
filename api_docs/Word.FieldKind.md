# Word.FieldKind enum

Package: [word](/en-us/javascript/api/word)

Represents the kind of field. Indicates how the field works in relation to updating.

## Remarks

[API set: WordApi 1.5]

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-fields.yaml

// Gets the first field in the document.
await Word.run(async (context) => {
  const field: Word.Field = context.document.body.fields.getFirstOrNullObject();
  field.load(["code", "result", "locked", "type", "data", "kind"]);

  await context.sync();

  if (field.isNullObject) {
    console.log("This document has no fields.");
  } else {
    console.log("Code of first field: " + field.code, "Result of first field: " + JSON.stringify(field.result), "Type of first field: " + field.type, "Is the first field locked? " + field.locked, "Kind of the first field: " + field.kind);
  }
});
```

## Fields

- cold = "Cold"
  - Represents that the field doesn't have a result.
  - [API set: WordApi 1.5]

- hot = "Hot"
  - Represents that the field is automatically updated each time it's displayed or each time the page is reformatted, but which can also be manually updated.
  - [API set: WordApi 1.5]

- none = "None"
  - Represents that the field is invalid. For example, a pair of field characters with nothing inside.
  - [API set: WordApi 1.5]

- warm = "Warm"
  - Represents that the field is automatically updated when the source changes or the field can be manually updated.
  - [API set: WordApi 1.5]