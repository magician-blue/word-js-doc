# Word.Interfaces.NoteItemUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the NoteItem object, for use in `noteItem.set({ ... })`.

## Properties

- body: Represents the body object of the note item. It's the portion of the text within the footnote or endnote.
- reference: Represents a footnote or endnote reference in the main document.

## Property Details

### body

Represents the body object of the note item. It's the portion of the text within the footnote or endnote.

```typescript
body?: Word.Interfaces.BodyUpdateData;
```

Property Value: [Word.Interfaces.BodyUpdateData](/en-us/javascript/api/word/word.interfaces.bodyupdatedata)

Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### reference

Represents a footnote or endnote reference in the main document.

```typescript
reference?: Word.Interfaces.RangeUpdateData;
```

Property Value: [Word.Interfaces.RangeUpdateData](/en-us/javascript/api/word/word.interfaces.rangeupdatedata)

Remarks

[API set: WordApi 1.5](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)