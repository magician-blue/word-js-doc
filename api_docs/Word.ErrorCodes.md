# Word.ErrorCodes enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

## Remarks

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    if (error.code === Word.ErrorCodes.itemNotFound) {
      console.warn("No checkbox content control is currently selected.");
    } else {
      console.error(error);
    }
  }
}
```

## Fields

- accessDenied = "AccessDenied"
- generalException = "GeneralException"
- invalidArgument = "InvalidArgument"
- itemNotFound = "ItemNotFound"
- notAllowed = "NotAllowed"
- notImplemented = "NotImplemented"
- searchDialogIsOpen = "SearchDialogIsOpen"
- searchStringInvalidOrTooLong = "SearchStringInvalidOrTooLong"