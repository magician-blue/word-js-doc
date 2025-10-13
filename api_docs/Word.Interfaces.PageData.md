# Word.Interfaces.PageData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling page.toJSON().

## Properties

- breaks  
  Gets a BreakCollection object that represents the breaks on the page.
- height  
  Gets the height, in points, of the paper defined in the Page Setup dialog box.
- index  
  Gets the index of the page. The page index is 1-based and independent of the user's custom page numbering.
- width  
  Gets the width, in points, of the paper defined in the Page Setup dialog box.

## Property Details

### breaks

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a `BreakCollection` object that represents the breaks on the page.

```typescript
breaks?: Word.Interfaces.BreakData[];
```

Property Value: [Word.Interfaces.BreakData](/en-us/javascript/api/word/word.interfaces.breakdata)[]

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### height

Gets the height, in points, of the paper defined in the Page Setup dialog box.

```typescript
height?: number;
```

Property Value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### index

Gets the index of the page. The page index is 1-based and independent of the user's custom page numbering.

```typescript
index?: number;
```

Property Value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### width

Gets the width, in points, of the paper defined in the Page Setup dialog box.

```typescript
width?: number;
```

Property Value: number

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)