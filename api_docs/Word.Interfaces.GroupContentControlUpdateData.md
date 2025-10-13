# Word.Interfaces.GroupContentControlUpdateData interface

Package: [word](/en-us/javascript/api/word)

An interface for updating data on the GroupContentControl object, for use in groupContentControl.set({ ... }).

## Properties

- appearance: Specifies the appearance of the content control.
- color: Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.
- isTemporary: Specifies whether to remove the content control from the active document when the user edits the contents of the control.
- lockContentControl: Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.
- lockContents: Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.
- placeholderText: Returns a BuildingBlock object that represents the placeholder text for the content control.
- range: Gets a Range object that represents the contents of the content control in the active document.
- tag: Specifies a tag to identify the content control.
- title: Specifies the title for the content control.
- xmlMapping: Gets an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

## Property Details

### appearance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the appearance of the content control.

TypeScript:
```typescript
appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property Value: [Word.ContentControlAppearance](/en-us/javascript/api/word/word.contentcontrolappearance) | "BoundingBox" | "Tags" | "Hidden"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

TypeScript:
```typescript
color?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isTemporary

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

TypeScript:
```typescript
isTemporary?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

TypeScript:
```typescript
lockContentControl?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### lockContents

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

TypeScript:
```typescript
lockContents?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### placeholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlock object that represents the placeholder text for the content control.

TypeScript:
```typescript
placeholderText?: Word.Interfaces.BuildingBlockUpdateData;
```

Property Value: [Word.Interfaces.BuildingBlockUpdateData](/en-us/javascript/api/word/word.interfaces.buildingblockupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a Range object that represents the contents of the content control in the active document.

TypeScript:
```typescript
range?: Word.Interfaces.RangeUpdateData;
```

Property Value: [Word.Interfaces.RangeUpdateData](/en-us/javascript/api/word/word.interfaces.rangeupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### tag

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a tag to identify the content control.

TypeScript:
```typescript
tag?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### title

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the title for the content control.

TypeScript:
```typescript
title?: string;
```

Property Value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

TypeScript:
```typescript
xmlMapping?: Word.Interfaces.XmlMappingUpdateData;
```

Property Value: [Word.Interfaces.XmlMappingUpdateData](/en-us/javascript/api/word/word.interfaces.xmlmappingupdatedata)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)