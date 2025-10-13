# Word.Interfaces.RepeatingSectionContentControlUpdateData interface

- Package: word

An interface for updating data on the RepeatingSectionContentControl object, for use in repeatingSectionContentControl.set({ ... }).

## Properties

- allowInsertDeleteSection: Specifies whether users can add or remove sections from this repeating section content control by using the user interface.
- appearance: Specifies the appearance of the content control.
- color: Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.
- isTemporary: Specifies whether to remove the content control from the active document when the user edits the contents of the control.
- lockContentControl: Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.
- lockContents: Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.
- placeholderText: Returns a BuildingBlock object that represents the placeholder text for the content control.
- range: Gets a Range object that represents the contents of the content control in the active document.
- repeatingSectionItemTitle: Specifies the name of the repeating section items used in the context menu associated with this repeating section content control.
- tag: Specifies a tag to identify the content control.
- title: Specifies the title for the content control.
- xmlapping: Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

## Property Details

### allowInsertDeleteSection

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether users can add or remove sections from this repeating section content control by using the user interface.

TypeScript:
```
allowInsertDeleteSection?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### appearance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the appearance of the content control.

TypeScript:
```
appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property Value: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden"

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

TypeScript:
```
color?: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### isTemporary

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

TypeScript:
```
isTemporary?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### lockContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the content control is locked (can't be deleted). true means that the user can't delete it from the active document, false means it can be deleted.

TypeScript:
```
lockContentControl?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### lockContents

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the contents of the content control are locked (not editable). true means the user can't edit the contents, false means the contents are editable.

TypeScript:
```
lockContents?: boolean;
```

Property Value: boolean

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### placeholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a BuildingBlock object that represents the placeholder text for the content control.

TypeScript:
```
placeholderText?: Word.Interfaces.BuildingBlockUpdateData;
```

Property Value: Word.Interfaces.BuildingBlockUpdateData

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a Range object that represents the contents of the content control in the active document.

TypeScript:
```
range?: Word.Interfaces.RangeUpdateData;
```

Property Value: Word.Interfaces.RangeUpdateData

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### repeatingSectionItemTitle

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the name of the repeating section items used in the context menu associated with this repeating section content control.

TypeScript:
```
repeatingSectionItemTitle?: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### tag

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a tag to identify the content control.

TypeScript:
```
tag?: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### title

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the title for the content control.

TypeScript:
```
title?: string;
```

Property Value: string

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]

---

### xmlapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an XmlMapping object that represents the mapping of the content control to XML data in the data store of the document.

TypeScript:
```
xmlapping?: Word.Interfaces.XmlMappingUpdateData;
```

Property Value: Word.Interfaces.XmlMappingUpdateData

Remarks: [ API set: WordApi BETA (PREVIEW ONLY) ]