# Word.Interfaces.PageCollectionLoadOptions interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents the collection of page.

## Remarks

[API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- height  
  For EACH ITEM in the collection: Gets the height, in points, of the paper defined in the Page Setup dialog box.
- index  
  For EACH ITEM in the collection: Gets the index of the page. The page index is 1-based and independent of the user's custom page numbering.
- width  
  For EACH ITEM in the collection: Gets the width, in points, of the paper defined in the Page Setup dialog box.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- Signature: `$all?: boolean;`
- Property Value: boolean

### height

For EACH ITEM in the collection: Gets the height, in points, of the paper defined in the Page Setup dialog box.

- Signature: `height?: boolean;`
- Property Value: boolean
- Remarks: [API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### index

For EACH ITEM in the collection: Gets the index of the page. The page index is 1-based and independent of the user's custom page numbering.

- Signature: `index?: boolean;`
- Property Value: boolean
- Remarks: [API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

For EACH ITEM in the collection: Gets the width, in points, of the paper defined in the Page Setup dialog box.

- Signature: `width?: boolean;`
- Property Value: boolean
- Remarks: [API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)