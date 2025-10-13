# Word.ListBullet enum

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

## Remarks

[API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/20-lists/organize-list.yaml

// Inserts a list starting with the first paragraph then set numbering and bullet types of the list items.
await Word.run(async (context) => {
  const paragraphs: Word.ParagraphCollection = context.document.body.paragraphs;
  paragraphs.load("$none");

  await context.sync();

  // Use the first paragraph to start a new list.
  const list: Word.List = paragraphs.items[0].startNewList();
  list.load("$none");

  await context.sync();

  // To add new items to the list, use Start or End on the insertLocation parameter.
  list.insertParagraph("New list item at the start of the list", "Start");
  const paragraph: Word.Paragraph = list.insertParagraph("New list item at the end of the list (set to list level 5)", "End");

  // Set numbering for list level 1.
  list.setLevelNumbering(0, Word.ListNumbering.arabic);

  // Set bullet type for list level 5.
  list.setLevelBullet(4, Word.ListBullet.arrow);

  // Set list level for the last item in this list.
  paragraph.listItem.level = 4;

  list.load("levelTypes");

  await context.sync();
});
```

## Fields

- arrow = "Arrow"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- checkmark = "Checkmark"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- custom = "Custom"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- diamonds = "Diamonds"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- hollow = "Hollow"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- solid = "Solid"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- square = "Square"
  - [API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)