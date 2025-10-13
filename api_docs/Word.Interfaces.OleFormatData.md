# Word.Interfaces.OleFormatData interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

An interface describing the data returned by calling `oleFormat.toJSON()`.

## Properties

- classType: Specifies the class type for the specified OLE object, picture, or field.
- iconIndex: Specifies the icon that is used when the `displayAsIcon` property is `true`.
- iconLabel: Specifies the text displayed below the icon for the OLE object.
- iconName: Specifies the program file in which the icon for the OLE object is stored.
- iconPath: Gets the path of the file in which the icon for the OLE object is stored.
- isDisplayedAsIcon: Gets whether the specified object is displayed as an icon.
- isFormattingPreservedOnUpdate: Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.
- label: Gets a string that's used to identify the portion of the source file that's being linked.
- progID: Gets the programmatic identifier (`ProgId`) for the specified OLE object.

## Property Details

### classType

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the class type for the specified OLE object, picture, or field.

```typescript
classType?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### iconIndex

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the icon that is used when the `displayAsIcon` property is `true`.

```typescript
iconIndex?: number;
```

Property value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### iconLabel

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the text displayed below the icon for the OLE object.

```typescript
iconLabel?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### iconName

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the program file in which the icon for the OLE object is stored.

```typescript
iconName?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### iconPath

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the path of the file in which the icon for the OLE object is stored.

```typescript
iconPath?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isDisplayedAsIcon

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets whether the specified object is displayed as an icon.

```typescript
isDisplayedAsIcon?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### isFormattingPreservedOnUpdate

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies whether formatting done in Microsoft Word to the linked OLE object is preserved.

```typescript
isFormattingPreservedOnUpdate?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### label

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a string that's used to identify the portion of the source file that's being linked.

```typescript
label?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### progID

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the programmatic identifier (`ProgId`) for the specified OLE object.

```typescript
progID?: string;
```

Property value: string

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)