# Word.Interfaces.PictureContentControlData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `pictureContentControl.toJSON()`.

## Properties

- appearance: Specifies the appearance of the content control.
- color: Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.
- id: Returns the identification for the content control.
- isTemporary: Specifies whether to remove the content control from the active document when the user edits the contents of the control.
- level: Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.
- lockContentControl: Specifies if the content control is locked (can't be deleted). `true` means that the user can't delete it from the active document, `false` means it can be deleted.
- lockContents: Specifies if the contents of the content control are locked (not editable). `true` means the user can't edit the contents, `false` means the contents are editable.
- placeholderText: Returns a `BuildingBlock` object that represents the placeholder text for the content control.
- range: Returns a `Range` object that represents the contents of the content control in the active document.
- showingPlaceholderText: Returns whether the placeholder text for the content control is being displayed.
- tag: Specifies a tag to identify the content control.
- title: Specifies the title for the content control.
- xmlMapping: Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

## Property Details

### appearance

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the appearance of the content control.

```typescript
appearance?: Word.ContentControlAppearance | "BoundingBox" | "Tags" | "Hidden";
```

Property value:
[Word.ContentControlAppearance](/en-us/javascript/api/word/word.contentcontrolappearance) | "BoundingBox" | "Tags" | "Hidden"

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### color

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the red-green-blue (RGB) value of the color of the content control. You can provide the value in the '#RRGGBB' format.

```typescript
color?: string;
```

Property value:
string

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### id

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the identification for the content control.

```typescript
id?: string;
```

Property value:
string

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### isTemporary

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether to remove the content control from the active document when the user edits the contents of the control.

```typescript
isTemporary?: boolean;
```

Property value:
boolean

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### level

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns the level of the content control—whether the content control surrounds text, paragraphs, table cells, or table rows; or if it is inline.

```typescript
level?: Word.ContentControlLevel | "Inline" | "Paragraph" | "Row" | "Cell";
```

Property value:
[Word.ContentControlLevel](/en-us/javascript/api/word/word.contentcontrollevel) | "Inline" | "Paragraph" | "Row" | "Cell"

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### lockContentControl

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the content control is locked (can't be deleted). `true` means that the user can't delete it from the active document, `false` means it can be deleted.

```typescript
lockContentControl?: boolean;
```

Property value:
boolean

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### lockContents

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies if the contents of the content control are locked (not editable). `true` means the user can't edit the contents, `false` means the contents are editable.

```typescript
lockContents?: boolean;
```

Property value:
boolean

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### placeholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `BuildingBlock` object that represents the placeholder text for the content control.

```typescript
placeholderText?: Word.Interfaces.BuildingBlockData;
```

Property value:
[Word.Interfaces.BuildingBlockData](/en-us/javascript/api/word/word.interfaces.buildingblockdata)

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### range

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `Range` object that represents the contents of the content control in the active document.

```typescript
range?: Word.Interfaces.RangeData;
```

Property value:
[Word.Interfaces.RangeData](/en-us/javascript/api/word/word.interfaces.rangedata)

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### showingPlaceholderText

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns whether the placeholder text for the content control is being displayed.

```typescript
showingPlaceholderText?: boolean;
```

Property value:
boolean

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### tag

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a tag to identify the content control.

```typescript
tag?: string;
```

Property value:
string

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### title

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the title for the content control.

```typescript
title?: string;
```

Property value:
string

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]

### xmlMapping

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns an `XmlMapping` object that represents the mapping of the content control to XML data in the data store of the document.

```typescript
xmlMapping?: Word.Interfaces.XmlMappingData;
```

Property value:
[Word.Interfaces.XmlMappingData](/en-us/javascript/api/word/word.interfaces.xmlmappingdata)

Remarks:
[ API set: WordApi BETA (PREVIEW ONLY) ]