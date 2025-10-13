# Word.Interfaces.DocumentUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the Document object, for use in document.set({ ... }).

## Properties

- activeWindow: Gets the active window for the document.
- autoHyphenation: Specifies if automatic hyphenation is turned on for the document.
- autoSaveOn: Specifies if the edits in the document are automatically saved.
- bibliography: Returns a Bibliography object that represents the bibliography references contained within the document.
- body: Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
- changeTrackingMode: Specifies the ChangeTracking mode.
- consecutiveHyphensLimit: Specifies the maximum number of consecutive lines that can end with hyphens.
- hyphenateCaps: Specifies whether words in all capital letters can be hyphenated.
- languageDetected: Specifies whether Microsoft Word has detected the language of the document text.
- pageSetup: Returns a PageSetup object that's associated with the document.
- properties: Gets the properties of the document.

## Property Details

### activeWindow

Gets the active window for the document.

```typescript
activeWindow?: Word.Interfaces.WindowUpdateData;
```

#### Property Value
[Word.Interfaces.WindowUpdateData](/en-us/javascript/api/word/word.interfaces.windowupdatedata)

#### Remarks
[API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### autoHyphenation

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if automatic hyphenation is turned on for the document.

```typescript
autoHyphenation?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### autoSaveOn

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the edits in the document are automatically saved.

```typescript
autoSaveOn?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bibliography

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Bibliography object that represents the bibliography references contained within the document.

```typescript
bibliography?: Word.Interfaces.BibliographyUpdateData;
```

#### Property Value
[Word.Interfaces.BibliographyUpdateData](/en-us/javascript/api/word/word.interfaces.bibliographyupdatedata)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### body

Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

```typescript
body?: Word.Interfaces.BodyUpdateData;
```

#### Property Value
[Word.Interfaces.BodyUpdateData](/en-us/javascript/api/word/word.interfaces.bodyupdatedata)

#### Remarks
[API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### changeTrackingMode

Specifies the ChangeTracking mode.

```typescript
changeTrackingMode?: Word.ChangeTrackingMode | "Off" | "TrackAll" | "TrackMineOnly";
```

#### Property Value
[Word.ChangeTrackingMode](/en-us/javascript/api/word/word.changetrackingmode) | "Off" | "TrackAll" | "TrackMineOnly"

#### Remarks
[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### consecutiveHyphensLimit

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the maximum number of consecutive lines that can end with hyphens.

```typescript
consecutiveHyphensLimit?: number;
```

#### Property Value
number

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hyphenateCaps

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether words in all capital letters can be hyphenated.

```typescript
hyphenateCaps?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### languageDetected

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word has detected the language of the document text.

```typescript
languageDetected?: boolean;
```

#### Property Value
boolean

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pageSetup

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a PageSetup object that's associated with the document.

```typescript
pageSetup?: Word.Interfaces.PageSetupUpdateData;
```

#### Property Value
[Word.Interfaces.PageSetupUpdateData](/en-us/javascript/api/word/word.interfaces.pagesetupupdatedata)

#### Remarks
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### properties

Gets the properties of the document.

```typescript
properties?: Word.Interfaces.DocumentPropertiesUpdateData;
```

#### Property Value
[Word.Interfaces.DocumentPropertiesUpdateData](/en-us/javascript/api/word/word.interfaces.documentpropertiesupdatedata)

#### Remarks
[API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)