# Word.UnderlineType enum

- Package: [word](/en-us/javascript/api/word)

The supported styles for underline format.

## Remarks

[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Underline format text
// Run a batch operation against the Word object model.
await Word.run(async (context) => {

    // Create a range proxy object for the current selection.
    const selection = context.document.getSelection();

    // Queue a command to underline the current selection.
    selection.font.underline = Word.UnderlineType.single;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    await context.sync();
    console.log('The selection now has an underline style.');
});
```

## Fields

- dashLine = "DashLine"
  - A single dash underline.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- dashLineHeavy = "DashLineHeavy"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- dashLineLong = "DashLineLong"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- dashLineLongHeavy = "DashLineLongHeavy"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- dotDashLine = "DotDashLine"
  - An alternating dot-dash underline.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- dotDashLineHeavy = "DotDashLineHeavy"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- dotLine = "DotLine"
  - Warning: dotLine has been deprecated.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- dotted = "Dotted"
  - A dotted underline.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- dottedHeavy = "DottedHeavy"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- double = "Double"
  - A double underline.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- hidden = "Hidden"
  - Warning: hidden has been deprecated.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- mixed = "Mixed"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- none = "None"
  - No underline.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- single = "Single"
  - A single underline. This is the default value.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- thick = "Thick"
  - A single thick underline.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- twoDotDashLine = "TwoDotDashLine"
  - An alternating dot-dot-dash underline.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- twoDotDashLineHeavy = "TwoDotDashLineHeavy"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- wave = "Wave"
  - A single wavy underline.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- waveDouble = "WaveDouble"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- waveHeavy = "WaveHeavy"
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- word = "Word"
  - Only underline individual words.
  - [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)