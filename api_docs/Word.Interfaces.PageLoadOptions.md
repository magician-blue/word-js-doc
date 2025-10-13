# Word.Interfaces.PageLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a page in the document. Page objects manage the page layout and content.

## Remarks

[ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- height  
  Gets the height, in points, of the paper defined in the Page Setup dialog box.

- index  
  Gets the index of the page. The page index is 1-based and independent of the user's custom page numbering.

- width  
  Gets the width, in points, of the paper defined in the Page Setup dialog box.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

TypeScript: `$all?: boolean;`

Property Value: boolean

### height

Gets the height, in points, of the paper defined in the Page Setup dialog box.

TypeScript: `height?: boolean;`

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### index

Gets the index of the page. The page index is 1-based and independent of the user's custom page numbering.

TypeScript: `index?: boolean;`

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Gets the width, in points, of the paper defined in the Page Setup dialog box.

TypeScript: `width?: boolean;`

Property Value: boolean

Remarks: [ API set: WordApiDesktop 1.2 ](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)