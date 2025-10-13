# Word.Interfaces.DropCapData interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

An interface describing the data returned by calling dropCap.toJSON().

## Properties

- distanceFromText — Gets the distance (in points) between the dropped capital letter and the paragraph text.
- fontName — Gets the name of the font for the dropped capital letter.
- linesToDrop — Gets the height (in lines) of the dropped capital letter.
- position — Gets the position of the dropped capital letter.

## Property Details

### distanceFromText

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the distance (in points) between the dropped capital letter and the paragraph text.

```typescript
distanceFromText?: number;
```

Property Value: number

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ] https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### fontName

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the name of the font for the dropped capital letter.

```typescript
fontName?: string;
```

Property Value: string

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ] https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### linesToDrop

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the height (in lines) of the dropped capital letter.

```typescript
linesToDrop?: number;
```

Property Value: number

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ] https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets

### position

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the position of the dropped capital letter.

```typescript
position?: Word.DropPosition | "None" | "Normal" | "Margin";
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.dropposition | "None" | "Normal" | "Margin"

Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ] https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets