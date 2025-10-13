# Word.StyleType enum

Package: [word](/en-us/javascript/api/word)

Represents the type of style.

## Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Applies the specified style to a paragraph.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to apply.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else if (style.type != Word.StyleType.paragraph) {
    console.log(`The '${styleName}' style isn't a paragraph style.`);
  } else {
    const body: Word.Body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Do you want to create a solution that extends the functionality of Word? You can use the Office Add-ins platform to extend Word clients running on the web, on a Windows desktop, or on a Mac.",
      "Start"
    );
    const paragraph: Word.Paragraph = body.paragraphs.getFirst();
    paragraph.style = style.nameLocal;
    console.log(`'${styleName}' style applied to first paragraph.`);
  }
});
```

## Fields

- character = "Character"
  - Represents that the style is a character style.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- list = "List"
  - Represents that the style is a list style. Currently supported on desktop.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- paragraph = "Paragraph"
  - Represents that the style is a paragraph style.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- table = "Table"
  - Represents that the style is a table style.
  - [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)