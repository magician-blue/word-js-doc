# Word.ContentControlOptions interface

Package: [word](/en-us/javascript/api/word)

Specifies the options that define which content controls are returned.

## Remarks

[ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

#### Examples
```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/10-content-controls/insert-and-change-checkbox-content-control.yaml

// Toggles the isChecked property of the first checkbox content control found in the selection.
await Word.run(async (context) => {
  const selectedRange: Word.Range = context.document.getSelection();
  let selectedContentControl = selectedRange
    .getContentControls({
      types: [Word.ContentControlType.checkBox]
    })
    .getFirstOrNullObject();
  selectedContentControl.load("id,checkboxContentControl/isChecked");

  await context.sync();

  if (selectedContentControl.isNullObject) {
    const parentContentControl: Word.ContentControl = selectedRange.parentContentControl;
    parentContentControl.load("id,type,checkboxContentControl/isChecked");
    await context.sync();

    if (parentContentControl.isNullObject || parentContentControl.type !== Word.ContentControlType.checkBox) {
      console.warn("No checkbox content control is currently selected.");
      return;
    } else {
      selectedContentControl = parentContentControl;
    }
  }

  const isCheckedBefore = selectedContentControl.checkboxContentControl.isChecked;
  console.log("isChecked state before:", `id: ${selectedContentControl.id} ... isChecked: ${isCheckedBefore}`);
  selectedContentControl.checkboxContentControl.isChecked = !isCheckedBefore;
  selectedContentControl.load("id,checkboxContentControl/isChecked");
  await context.sync();

  console.log(
    "isChecked state after:",
    `id: ${selectedContentControl.id} ... isChecked: ${selectedContentControl.checkboxContentControl.isChecked}`
  );
});
```

## Properties

- types: An array of content control types, item must be 'RichText', 'PlainText', 'CheckBox', 'DropDownList', or 'ComboBox'.

## Property Details

### types

An array of content control types, item must be 'RichText', 'PlainText', 'CheckBox', 'DropDownList', or 'ComboBox'.

```typescript
types: Word.ContentControlType[];
```

#### Property Value

[Word.ContentControlType](/en-us/javascript/api/word/word.contentcontroltype)[]

#### Remarks

[ [API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

Note: 'PlainText' support was added in WordApi 1.5. 'CheckBox' support was added in WordApi 1.7. 'DropDownList' and 'ComboBox' support was added in WordApi 1.9.