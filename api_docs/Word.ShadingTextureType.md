# Word.ShadingTextureType enum

Package: [word](/en-us/javascript/api/word)

Represents the shading texture. To learn more about how to apply backgrounds like textures, see Add, change, or delete the background color in Word.

## Remarks

[API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

#### Examples

```typescript
// Link to full sample: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-styles.yaml

// Updates shading properties (e.g., texture, pattern colors) of the specified style.
await Word.run(async (context) => {
  const styleName = (document.getElementById("style-name") as HTMLInputElement).value;
  if (styleName == "") {
    console.warn("Enter a style name to update shading properties.");
    return;
  }

  const style: Word.Style = context.document.getStyles().getByNameOrNullObject(styleName);
  style.load();
  await context.sync();

  if (style.isNullObject) {
    console.warn(`There's no existing style with the name '${styleName}'.`);
  } else {
    const shading: Word.Shading = style.shading;
    shading.load();
    await context.sync();

    shading.backgroundPatternColor = "blue";
    shading.foregroundPatternColor = "yellow";
    shading.texture = Word.ShadingTextureType.darkTrellis;

    console.log("Updated shading.");
  }
});
```

## Fields

- darkDiagonalDown = "DarkDiagonalDown"  
  Represents dark diagonal-down texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- darkDiagonalUp = "DarkDiagonalUp"  
  Represents dark diagonal-up texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- darkGrid = "DarkGrid"  
  Represents dark horizontal-cross texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- darkHorizontal = "DarkHorizontal"  
  Represents dark horizontal texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- darkTrellis = "DarkTrellis"  
  Represents dark diagonal-cross texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- darkVertical = "DarkVertical"  
  Represents dark vertical texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- lightDiagonalDown = "LightDiagonalDown"  
  Represents light diagonal-down texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- lightDiagonalUp = "LightDiagonalUp"  
  Represents light diagonal-up texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- lightGrid = "LightGrid"  
  Represents light horizontal-cross texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- lightHorizontal = "LightHorizontal"  
  Represents light horizontal texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- lightTrellis = "LightTrellis"  
  Represents light diagonal-cross texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- lightVertical = "LightVertical"  
  Represents light vertical texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- none = "None"  
  Represents that there's no texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent10 = "Percent10"  
  Represents 10 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent12Pt5 = "Percent12Pt5"  
  Represents 12.5 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent15 = "Percent15"  
  Represents 15 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent20 = "Percent20"  
  Represents 20 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent25 = "Percent25"  
  Represents 25 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent30 = "Percent30"  
  Represents 30 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent35 = "Percent35"  
  Represents 35 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent37Pt5 = "Percent37Pt5"  
  Represents 37.5 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent40 = "Percent40"  
  Represents 40 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent45 = "Percent45"  
  Represents 45 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent5 = "Percent5"  
  Represents 5 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent50 = "Percent50"  
  Represents 50 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent55 = "Percent55"  
  Represents 55 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent60 = "Percent60"  
  Represents 60 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent62Pt5 = "Percent62Pt5"  
  Represents 62.5 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent65 = "Percent65"  
  Represents 65 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent70 = "Percent70"  
  Represents 70 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent75 = "Percent75"  
  Represents 75 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent80 = "Percent80"  
  Represents 80 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent85 = "Percent85"  
  Represents 85 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent87Pt5 = "Percent87Pt5"  
  Represents 87.5 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent90 = "Percent90"  
  Represents 90 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- percent95 = "Percent95"  
  Represents 95 percent texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

- solid = "Solid"  
  Represents solid texture.  
  [API set: WordApiDesktop 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

[Add, change, or delete the background color in Word]: https://support.microsoft.com/office/db481e61-7af6-4063-bbcd-b276054a5515