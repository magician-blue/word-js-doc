# Word.Interfaces.DocumentData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `document.toJSON()`.

## Properties

- `activeWindow` — Gets the active window for the document.
- `autoHyphenation` — Specifies if automatic hyphenation is turned on for the document.
- `autoSaveOn` — Specifies if the edits in the document are automatically saved.
- `bibliography` — Returns a `Bibliography` object that represents the bibliography references contained within the document.
- `body` — Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.
- `changeTrackingMode` — Specifies the ChangeTracking mode.
- `consecutiveHyphensLimit` — Specifies the maximum number of consecutive lines that can end with hyphens.
- `contentControls` — Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.
- `customXmlParts` — Gets the custom XML parts in the document.
- `documentLibraryVersions` — Returns a `DocumentLibraryVersionCollection` object that represents the collection of versions of a shared document that has versioning enabled and that's stored in a document library on a server.
- `frames` — Returns a `FrameCollection` object that represents all the frames in the document.
- `hyperlinks` — Returns a `HyperlinkCollection` object that represents all the hyperlinks in the document.
- `hyphenateCaps` — Specifies whether words in all capital letters can be hyphenated.
- `languageDetected` — Specifies whether Microsoft Word has detected the language of the document text.
- `pageSetup` — Returns a `PageSetup` object that's associated with the document.
- `properties` — Gets the properties of the document.
- `saved` — Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.
- `sections` — Gets the collection of section objects in the document.
- `settings` — Gets the add-in's settings in the document.
- `windows` — Gets the collection of `Word.Window` objects for the document.

## Property Details

### activeWindow

Gets the active window for the document.

```typescript
activeWindow?: Word.Interfaces.WindowData;
```

Property Value: [Word.Interfaces.WindowData](/en-us/javascript/api/word/word.interfaces.windowdata)

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### autoHyphenation

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if automatic hyphenation is turned on for the document.

```typescript
autoHyphenation?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### autoSaveOn

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the edits in the document are automatically saved.

```typescript
autoSaveOn?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### bibliography

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Bibliography` object that represents the bibliography references contained within the document.

```typescript
bibliography?: Word.Interfaces.BibliographyData;
```

Property Value: [Word.Interfaces.BibliographyData](/en-us/javascript/api/word/word.interfaces.bibliographydata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### body

Gets the body object of the main document. The body is the text that excludes headers, footers, footnotes, textboxes, etc.

```typescript
body?: Word.Interfaces.BodyData;
```

Property Value: [Word.Interfaces.BodyData](/en-us/javascript/api/word/word.interfaces.bodydata)

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### changeTrackingMode

Specifies the ChangeTracking mode.

```typescript
changeTrackingMode?: Word.ChangeTrackingMode | "Off" | "TrackAll" | "TrackMineOnly";
```

Property Value: [Word.ChangeTrackingMode](/en-us/javascript/api/word/word.changetrackingmode) | "Off" | "TrackAll" | "TrackMineOnly"

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### consecutiveHyphensLimit

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the maximum number of consecutive lines that can end with hyphens.

```typescript
consecutiveHyphensLimit?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### contentControls

Gets the collection of content control objects in the document. This includes content controls in the body of the document, headers, footers, textboxes, etc.

```typescript
contentControls?: Word.Interfaces.ContentControlData[];
```

Property Value: [Word.Interfaces.ContentControlData](/en-us/javascript/api/word/word.interfaces.contentcontroldata)[]

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### customXmlParts

Gets the custom XML parts in the document.

```typescript
customXmlParts?: Word.Interfaces.CustomXmlPartData[];
```

Property Value: [Word.Interfaces.CustomXmlPartData](/en-us/javascript/api/word/word.interfaces.customxmlpartdata)[]

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### documentLibraryVersions

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `DocumentLibraryVersionCollection` object that represents the collection of versions of a shared document that has versioning enabled and that's stored in a document library on a server.

```typescript
documentLibraryVersions?: Word.Interfaces.DocumentLibraryVersionData[];
```

Property Value: [Word.Interfaces.DocumentLibraryVersionData](/en-us/javascript/api/word/word.interfaces.documentlibraryversiondata)[]

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### frames

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `FrameCollection` object that represents all the frames in the document.

```typescript
frames?: Word.Interfaces.FrameData[];
```

Property Value: [Word.Interfaces.FrameData](/en-us/javascript/api/word/word.interfaces.framedata)[]

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hyperlinks

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `HyperlinkCollection` object that represents all the hyperlinks in the document.

```typescript
hyperlinks?: Word.Interfaces.HyperlinkData[];
```

Property Value: [Word.Interfaces.HyperlinkData](/en-us/javascript/api/word/word.interfaces.hyperlinkdata)[]

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### hyphenateCaps

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether words in all capital letters can be hyphenated.

```typescript
hyphenateCaps?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### languageDetected

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether Microsoft Word has detected the language of the document text.

```typescript
languageDetected?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### pageSetup

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `PageSetup` object that's associated with the document.

```typescript
pageSetup?: Word.Interfaces.PageSetupData;
```

Property Value: [Word.Interfaces.PageSetupData](/en-us/javascript/api/word/word.interfaces.pagesetupdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### properties

Gets the properties of the document.

```typescript
properties?: Word.Interfaces.DocumentPropertiesData;
```

Property Value: [Word.Interfaces.DocumentPropertiesData](/en-us/javascript/api/word/word.interfaces.documentpropertiesdata)

Remarks: [API set: WordApi 1.3](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### saved

Indicates whether the changes in the document have been saved. A value of true indicates that the document hasn't changed since it was saved.

```typescript
saved?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### sections

Gets the collection of section objects in the document.

```typescript
sections?: Word.Interfaces.SectionData[];
```

Property Value: [Word.Interfaces.SectionData](/en-us/javascript/api/word/word.interfaces.sectiondata)[]

Remarks: [API set: WordApi 1.1](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### settings

Gets the add-in's settings in the document.

```typescript
settings?: Word.Interfaces.SettingData[];
```

Property Value: [Word.Interfaces.SettingData](/en-us/javascript/api/word/word.interfaces.settingdata)[]

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### windows

Gets the collection of `Word.Window` objects for the document.

```typescript
windows?: Word.Interfaces.WindowData[];
```

Property Value: [Word.Interfaces.WindowData](/en-us/javascript/api/word/word.interfaces.windowdata)[]

Remarks: [API set: WordApiDesktop 1.2](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)