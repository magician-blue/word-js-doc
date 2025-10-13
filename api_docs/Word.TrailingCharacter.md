# Word.TrailingCharacter enum

Package: [word](/en-us/javascript/api/word)

Represents the character inserted after the list item mark.

## Remarks

[ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/manage-list-styles.yaml

// Gets the properties of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name-to-use") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to get properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load("type");
  await context.sync();

  if (style.isNullObject || style.type != Word.StyleType.list) {
    console.warn(`There's no existing style with the name '${styleName}'. Or this isn't a list style.`);
  } else {
    // Load objects to log properties and their values in the console.
    style.load();
    style.listTemplate.load();
    await context.sync();

    console.log(`Properties of the '${styleName}' style:`, style);

    const listLevels = style.listTemplate.listLevels;
    listLevels.load("items");
    await context.sync();

    console.log(`List levels of the '${styleName}' style:`, listLevels);
  }
});
```

## Fields

- trailingNone = "TrailingNone"
  - No character is inserted.
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- trailingSpace = "TrailingSpace"
  - A space is inserted. Default.
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- trailingTab = "TrailingTab"
  - A tab is inserted.
  - [ API set: WordApiDesktop 1.1 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)