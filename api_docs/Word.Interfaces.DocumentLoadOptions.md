# Word.Interfaces.DocumentLoadOptions interface

Package: [word](/en-us/javascript/api/word)

The Document object is the top level object. A Document object contains one or more sections, content controls, and the body that contains the contents of the document.

## Remarks
[API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- `$all` — Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- `activeWindow` — Gets the active window for the document.
- `attachedTemplate` — Specifies a `Template` object that represents the template attached to the document.
- `autoHyphenation` — Specifies if automatic hyphenation is turned on for the document.
- `autoSaveOn` — Specifies if the edits in the document are automatically saved.
- `bibliography` — Returns a `Bibliography` object that represents the bibliography references contained within the document.
- `body` — Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
- `changeTrackingMode` — Specifies the ChangeTracking mode.
- `consecutiveHyphensLimit` — Specifies the maximum number of consecutive lines that can end with hyphens.
- `hyphenateCaps` — Specifies whether words in all capital letters can be hyphenated.
- `languageDetected` — Specifies whether Microsoft Word has detected the language of the document text.
- `pageSetup` — Returns a `PageSetup` object that's associated with the document.
- `properties` — Gets the properties of the document.
- `saved` — Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

## Property Details

### $all
Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property value: boolean

---

### activeWindow
Gets the active window for the document.

```typescript
activeWindow?: Word.Interfaces.WindowLoadOptions;
```

Property value: [Word.Interfaces.WindowLoadOptions](/en-us/javascript/api/word/word.interfaces.windowloadoptions)

Remarks  
[API set: WordApiDesktop 1.2](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### attachedTemplate
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `Template` object that represents the template attached to the document.

```typescript
attachedTemplate?: Word.Interfaces.TemplateLoadOptions;
```

Property value: [Word.Interfaces.TemplateLoadOptions](/en-us/javascript/api/word/word.interfaces.templateloadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### autoHyphenation
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if automatic hyphenation is turned on for the document.

```typescript
autoHyphenation?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### autoSaveOn
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the edits in the document are automatically saved.

```typescript
autoSaveOn?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bibliography
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Bibliography` object that represents the bibliography references contained within the document.

```typescript
bibliography?: Word.Interfaces.BibliographyLoadOptions;
```

Property value: [Word.Interfaces.BibliographyLoadOptions](/en-us/javascript/api/word/word.interfaces.bibliographyloadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### body
Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

```typescript
body?: Word.Interfaces.BodyLoadOptions;
```

Property value: [Word.Interfaces.BodyLoadOptions](/en-us/javascript/api/word/word.interfaces.bodyloadoptions)

Remarks  
[API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### changeTrackingMode
Specifies the ChangeTracking mode.

```typescript
changeTrackingMode?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### consecutiveHyphensLimit
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the maximum number of consecutive lines that can end with hyphens.

```typescript
consecutiveHyphensLimit?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hyphenateCaps
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether words in all capital letters can be hyphenated.

```typescript
hyphenateCaps?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### languageDetected
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word has detected the language of the document text.

```typescript
languageDetected?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pageSetup
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `PageSetup` object that's associated with the document.

```typescript
pageSetup?: Word.Interfaces.PageSetupLoadOptions;
```

Property value: [Word.Interfaces.PageSetupLoadOptions](/en-us/javascript/api/word/word.interfaces.pagesetuploadoptions)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### properties
Gets the properties of the document.

```typescript
properties?: Word.Interfaces.DocumentPropertiesLoadOptions;
```

Property value: [Word.Interfaces.DocumentPropertiesLoadOptions](/en-us/javascript/api/word/word.interfaces.documentpropertiesloadoptions)

Remarks  
[API set: WordApi 1.3](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### saved
Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

```typescript
saved?: boolean;
```

Property value: boolean

Remarks  
[API set: WordApi 1.1](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)