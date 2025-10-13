# Word.Interfaces.SectionCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains the collection of the document's [Word.Section](/en-us/javascript/api/word/word.section) objects.

## Remarks

[ API set: WordApi 1.1 ]

## Properties

| Property | Description |
| --- | --- |
| $all | Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`). |
| body | For EACH ITEM in the collection: Gets the body object of the section. This doesn't include the header/footer and other section metadata. |
| pageSetup | For EACH ITEM in the collection: Returns a `PageSetup` object that's associated with the section. |
| protectedForForms | For EACH ITEM in the collection: Specifies if the section is protected for forms. |

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

---

### body

For EACH ITEM in the collection: Gets the body object of the section. This doesn't include the header/footer and other section metadata.

```typescript
body?: Word.Interfaces.BodyLoadOptions;
```

Property Value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks

[ API set: WordApi 1.1 ]

---

### pageSetup

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Returns a `PageSetup` object that's associated with the section.

```typescript
pageSetup?: Word.Interfaces.PageSetupLoadOptions;
```

Property Value: [Word.Interfaces.PageSetupLoadOptions](/en-us/javascript/api/word/word.interfaces.pagesetuploadoptions)

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]

---

### protectedForForms

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Specifies if the section is protected for forms.

```typescript
protectedForForms?: boolean;
```

Property Value: boolean

Remarks

[ API set: WordApi BETA (PREVIEW ONLY) ]