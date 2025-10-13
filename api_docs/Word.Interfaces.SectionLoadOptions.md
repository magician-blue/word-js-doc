# Word.Interfaces.SectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a section in a Word document.

## Remarks

[ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

## Properties

- [$all](#word-word-interfaces-sectionloadoptions-all-member) — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- [body](#word-word-interfaces-sectionloadoptions-body-member) — Gets the body object of the section. This doesn't include the header/footer and other section metadata.
- [pageSetup](#word-word-interfaces-sectionloadoptions-pagesetup-member) — Returns a PageSetup object that's associated with the section.
- [protectedForForms](#word-word-interfaces-sectionloadoptions-protectedforforms-member) — Specifies if the section is protected for forms.

## Property Details

<a id="word-word-interfaces-sectionloadoptions-all-member"></a>
### $all

Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

```typescript
$all?: boolean;
```

Property Value
- boolean

<a id="word-word-interfaces-sectionloadoptions-body-member"></a>
### body

Gets the body object of the section. This doesn't include the header/footer and other section metadata.

```typescript
body?: Word.Interfaces.BodyLoadOptions;
```

Property Value
- [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks  
[ [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

<a id="word-word-interfaces-sectionloadoptions-pagesetup-member"></a>
### pageSetup

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a PageSetup object that's associated with the section.

```typescript
pageSetup?: Word.Interfaces.PageSetupLoadOptions;
```

Property Value
- [Word.Interfaces.PageSetupLoadOptions](/en-us/javascript/api/word/word.interfaces.pagesetuploadoptions)

Remarks  
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

<a id="word-word-interfaces-sectionloadoptions-protectedforforms-member"></a>
### protectedForForms

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the section is protected for forms.

```typescript
protectedForForms?: boolean;
```

Property Value
- boolean

Remarks  
[ [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]