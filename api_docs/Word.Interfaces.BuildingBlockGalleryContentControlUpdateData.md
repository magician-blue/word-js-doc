# Word.Interfaces.BuildingBlockGalleryContentControlUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the BuildingBlockGalleryContentControl object, for use in buildingBlockGalleryContentControl.set({ ... }).

## Properties

- appearance — Specifies the appearance of the content control.
- buildingBlockCategory — Specifies the category for the building block content control.
- buildingBlockType — Specifies a BuildingBlockType value that represents the type of building block for the building block content control.
- color — Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.
- isTemporary — Specifies whether to remove the content control from the active document when the user edits the contents of the control.
- lockContentControl — Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.
- lockContents — Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.
- placeholderText — Returns a BuildingBlock object that represents the placeholder text for the content control.
- range — Returns a Range object that represents the contents of the content control in the active document.
- tag — Specifies a tag to identify the content control.
- title — Specifies the title for the content control.
- xmlMapping — Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

## Property Details

### appearance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the appearance of the content control.

```typescript
appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property value: [Word.ContentControlAppearance](/en-us/javascript/api/word/word.contentcontrolappearance) | "BoundingBox" | "Tags" | "Hidden"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### buildingBlockCategory

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the category for the building block content control.

```typescript
buildingBlockCategory?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### buildingBlockType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a BuildingBlockType value that represents the type of building block for the building block content control.

```typescript
buildingBlockType?: Word.BuildingBlockType | "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography";
```

Property value: [Word.BuildingBlockType](/en-us/javascript/api/word/word.buildingblocktype) | "QuickParts" | "CoverPage" | "Equations" | "Footers" | "Headers" | "PageNumber" | "Tables" | "Watermarks" | "AutoText" | "TextBox" | "PageNumberTop" | "PageNumberBottom" | "PageNumberPage" | "TableOfContents" | "CustomQuickParts" | "CustomCoverPage" | "CustomEquations" | "CustomFooters" | "CustomHeaders" | "CustomPageNumber" | "CustomTables" | "CustomWatermarks" | "CustomAutoText" | "CustomTextBox" | "CustomPageNumberTop" | "CustomPageNumberBottom" | "CustomPageNumberPage" | "CustomTableOfContents" | "Custom1" | "Custom2" | "Custom3" | "Custom4" | "Custom5" | "Bibliography" | "CustomBibliography"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

```typescript
color?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isTemporary

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

```typescript
isTemporary?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lockContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

```typescript
lockContentControl?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### lockContents

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

```typescript
lockContents?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### placeholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlock object that represents the placeholder text for the content control.

```typescript
placeholderText?: Word.Interfaces.BuildingBlockUpdateData;
```

Property value: [Word.Interfaces.BuildingBlockUpdateData](/en-us/javascript/api/word/word.interfaces.buildingblockupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a Range object that represents the contents of the content control in the active document.

```typescript
range?: Word.Interfaces.RangeUpdateData;
```

Property value: [Word.Interfaces.RangeUpdateData](/en-us/javascript/api/word/word.interfaces.rangeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### tag

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a tag to identify the content control.

```typescript
tag?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### title

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the title for the content control.

```typescript
title?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
xmlMapping?: Word.Interfaces.XmlMappingUpdateData;
```

Property value: [Word.Interfaces.XmlMappingUpdateData](/en-us/javascript/api/word/word.interfaces.xmlmappingupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)