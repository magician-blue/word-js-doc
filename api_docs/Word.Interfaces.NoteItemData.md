# Word.Interfaces.NoteItemData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `noteItem.toJSON()`.

## Properties

- body — Represents the body object of the note item. It's the portion of the text within the footnote or endnote.
- reference — Represents a footnote or endnote reference in the main document.
- type — Represents the note item type: footnote or endnote.

## Property Details

### body

Represents the body object of the note item. It's the portion of the text within the footnote or endnote.

```typescript
body?: Word.Interfaces.BodyData;
```

Property Value: [Word.Interfaces.BodyData](/en-us/javascript/api/word/word.interfaces.bodydata)

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### reference

Represents a footnote or endnote reference in the main document.

```typescript
reference?: Word.Interfaces.RangeData;
```

Property Value: [Word.Interfaces.RangeData](/en-us/javascript/api/word/word.interfaces.rangedata)

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Represents the note item type: footnote or endnote.

```typescript
type?: Word.NoteItemType | "Footnote" | "Endnote";
```

Property Value: [Word.NoteItemType](/en-us/javascript/api/word/word.noteitemtype) | "Footnote" | "Endnote"

Remarks  
[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)