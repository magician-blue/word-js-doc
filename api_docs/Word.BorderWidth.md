# Word.BorderWidth enum

Package: [word](/en-us/javascript/api/word)

Represents the width of a style's border.

## Remarks

API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### Examples

```TypeScript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Updates border properties (e.g., type, width, color) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update border properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const borders: Word.BorderCollection = style.borders;
    borders.load("items");
    await context.sync();

    borders.outsideBorderType = Word.BorderType.dashed;
    borders.outsideBorderWidth = Word.BorderWidth.pt025;
    borders.outsideBorderColor = "green";
    console.log("Updated outside borders.");
  }
});
```

## Fields

- mixed = "Mixed"
  - Mixed width.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- none = "None"
  - None width.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt025 = "Pt025"
  - 0.25 point.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt050 = "Pt050"
  - 0.50 point.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt075 = "Pt075"
  - 0.75 point.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt100 = "Pt100"
  - 1.00 point. This is the default.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt150 = "Pt150"
  - 1.50 points.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt225 = "Pt225"
  - 2.25 points.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt300 = "Pt300"
  - 3.00 points.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt450 = "Pt450"
  - 4.50 points.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)
- pt600 = "Pt600"
  - 6.00 points.
  - API set: [WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)