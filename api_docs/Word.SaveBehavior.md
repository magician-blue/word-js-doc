# Word.SaveBehavior enum

Package: [word](/en-us/javascript/api/word)

Specifies the save behavior for `Document.save`.

## Remarks

[ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/save-close.yaml

// If the document hasn't been saved before, prompts
// user with options for if or how they want to save.
await Word.run(async (context) => {
  context.document.save(Word.SaveBehavior.prompt);
  await context.sync();
});
```

## Fields

- prompt = "Prompt"
  - Displays the "Save As" dialog to the user if the document hasn't been saved. Won't take effect if the document was previously saved.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

- save = "Save"
  - Saves the document without prompting the user. If it's a new document, it will be saved with the default name or specified name in the default location.
  - [ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]