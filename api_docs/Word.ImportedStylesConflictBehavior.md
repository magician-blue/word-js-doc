# Word.ImportedStylesConflictBehavior enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Specifies how to handle any conflicts, that is, when imported styles have the same name as existing styles in the current document.

## Remarks

[API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Fields

- createNew = "CreateNew"
  - Rename conflicting imported styles so that both versions are kept in the current document. For example, if MyStyle already exists in the document, then the imported version could be added as MyStyle1.
  - [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- ignore = "Ignore"
  - Ignore conflicting imported styles and keep the existing version of those styles in the current document.
  - [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- overwrite = "Overwrite"
  - Overwrite the existing styles in the current document.
  - [API set: WordApiDesktop 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)