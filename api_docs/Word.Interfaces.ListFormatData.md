# Word.Interfaces.ListFormatData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `listFormat.toJSON()`.

## Properties

- [isSingleList](#issinglelist)
  - Indicates whether the `ListFormat` object contains a single list.
- [isSingleListTemplate](#issinglelisttemplate)
  - Indicates whether the `ListFormat` object contains a single list template.
- [list](#list)
  - Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.
- [listLevelNumber](#listlevelnumber)
  - Specifies the list level number for the first paragraph for the `ListFormat` object.
- [listString](#liststring)
  - Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.
- [listTemplate](#listtemplate)
  - Gets the list template associated with the `ListFormat` object.
- [listType](#listtype)
  - Gets the type of the list for the `ListFormat` object.
- [listValue](#listvalue)
  - Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.

## Property Details

### isSingleList

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indicates whether the `ListFormat` object contains a single list.

```typescript
isSingleList?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### isSingleListTemplate

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indicates whether the `ListFormat` object contains a single list template.

```typescript
isSingleListTemplate?: boolean;
```

Property Value
- boolean

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### list

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.

```typescript
list?: Word.Interfaces.ListData;
```

Property Value
- [Word.Interfaces.ListData](/en-us/javascript/api/word/word.interfaces.listdata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listLevelNumber

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the list level number for the first paragraph for the `ListFormat` object.

```typescript
listLevelNumber?: number;
```

Property Value
- number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listString

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.

```typescript
listString?: string;
```

Property Value
- string

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listTemplate

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the list template associated with the `ListFormat` object.

```typescript
listTemplate?: Word.Interfaces.ListTemplateData;
```

Property Value
- [Word.Interfaces.ListTemplateData](/en-us/javascript/api/word/word.interfaces.listtemplatedata)

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listType

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the type of the list for the `ListFormat` object.

```typescript
listType?: Word.ListType | "ListNoNumbering" | "ListListNumOnly" | "ListBullet" | "ListSimpleNumbering" | "ListOutlineNumbering" | "ListMixedNumbering" | "ListPictureBullet";
```

Property Value
- [Word.ListType](/en-us/javascript/api/word/word.listtype) | "ListNoNumbering" | "ListListNumOnly" | "ListBullet" | "ListSimpleNumbering" | "ListOutlineNumbering" | "ListMixedNumbering" | "ListPictureBullet"

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### listValue

> Note  
> This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.

```typescript
listValue?: number;
```

Property Value
- number

Remarks  
[API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)