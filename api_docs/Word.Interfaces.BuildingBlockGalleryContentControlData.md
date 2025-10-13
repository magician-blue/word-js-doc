# Word.Interfaces.BuildingBlockGalleryContentControlData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface describing the data returned by calling `buildingBlockGalleryContentControl.toJSON()`.

## Properties

- [appearance](#appearance) — Specifies the appearance of the content control.
- [buildingBlockCategory](#buildingblockcategory) — Specifies the category for the building block content control.
- [buildingBlockType](#buildingblocktype) — Specifies a `BuildingBlockType` value that represents the type of building block for the building block content control.
- [color](#color) — Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.
- [id](#id) — Gets the identification for the content control.
- [isTemporary](#istemporary) — Specifies whether to remove the content control from the active document when the user edits the contents of the control.
- [level](#level) — Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.
- [lockContentControl](#lockcontentcontrol) — Specifies if the content control is locked (can't be deleted). `true` means that the user can't delete it from the active document, `false` means it can be deleted.
- [lockContents](#lockcontents) — Specifies if the contents of the content control are locked (not editable). `true` means the user can't edit the contents, `false` means the contents are editable.
- [placeholderText](#placeholdertext) — Returns a `BuildingBlock` object that represents the placeholder text for the content control.
- [range](#range) — Returns a `Range` object that represents the contents of the content control in the active document.
- [showingPlaceholderText](#showingplaceholdertext) — Gets if the placeholder text for the content control is being displayed.
- [tag](#tag) — Specifies a tag to identify the content control.
- [title](#title) — Specifies the title for the content control.
- [xmlMapping](#xmlmapping) — Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

## Property Details

### appearance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the appearance of the content control.

```typescript
appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property Value: [Word.ContentControlAppearance](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrolappearance) | "BoundingBox" | "Tags" | "Hidden"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### buildingBlockCategory

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the category for the building block content control.

```typescript
buildingBlockCategory?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### buildingBlockType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a `BuildingBlockType` value that represents the type of building block for the building block content control.

```typescript
buildingBlockType?: Word.BuildingBlockType | "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography";
```

Property Value: [Word.BuildingBlockType](https://learn.microsoft.com/en-us/javascript/api/word/word.buildingblocktype) | "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

```typescript
color?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the identification for the content control.

```typescript
id?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isTemporary

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

```typescript
isTemporary?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### level

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

```typescript
level?: Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell";
```

Property Value: [Word.ContentControlLevel](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrollevel) | "Inline" | "Paragraph" | "Row" | "Cell"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the content control is locked (can't be deleted). `true` means that the user can't delete it from the active document, `false` means it can be deleted.

```typescript
lockContentControl?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockContents

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the contents of the content control are locked (not editable). `true` means the user can't edit the contents, `false` means the contents are editable.

```typescript
lockContents?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### placeholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BuildingBlock` object that represents the placeholder text for the content control.

```typescript
placeholderText?: Word.Interfaces.BuildingBlockData;
```

Property Value: [Word.Interfaces.BuildingBlockData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.buildingblockdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Range` object that represents the contents of the content control in the active document.

```typescript
range?: Word.Interfaces.RangeData;
```

Property Value: [Word.Interfaces.RangeData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.rangedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### showingPlaceholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets if the placeholder text for the content control is being displayed.

```typescript
showingPlaceholderText?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tag

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a tag to identify the content control.

```typescript
tag?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### title

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the title for the content control.

```typescript
title?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
xmlMapping?: Word.Interfaces.XmlMappingData;
```

Property Value: [Word.Interfaces.XmlMappingData](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.xmlmappingdata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)