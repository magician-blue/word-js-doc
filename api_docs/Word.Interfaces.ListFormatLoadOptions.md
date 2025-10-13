# Word.Interfaces.ListFormatLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the list formatting characteristics of a range.

## Remarks
[ API set: WordApi BETA (PREVIEW ONLY) ]

## Properties
- $all  
  Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

- isSingleList  
  Indicates whether the `ListFormat` object contains a single list.

- isSingleListTemplate  
  Indicates whether the `ListFormat` object contains a single list template.

- list  
  Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.

- listLevelNumber  
  Specifies the list level number for the first paragraph for the `ListFormat` object.

- listString  
  Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.

- listTemplate  
  Gets the list template associated with the `ListFormat` object.

- listType  
  Gets the type of the list for the `ListFormat` object.

- listValue  
  Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.

## Property Details

### $all
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value  
boolean

### isSingleList
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indicates whether the `ListFormat` object contains a single list.

```typescript
isSingleList?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### isSingleListTemplate
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Indicates whether the `ListFormat` object contains a single list template.

```typescript
isSingleListTemplate?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### list
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Returns a `List` object that represents the first formatted list contained in the `ListFormat` object.

```typescript
list?: Word.Interfaces.ListLoadOptions;
```

Property Value  
[Word.Interfaces.ListLoadOptions](/en-us/javascript/api/word/word.interfaces.listloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### listLevelNumber
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the list level number for the first paragraph for the `ListFormat` object.

```typescript
listLevelNumber?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### listString
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the string representation of the list value of the first paragraph in the range for the `ListFormat` object.

```typescript
listString?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### listTemplate
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the list template associated with the `ListFormat` object.

```typescript
listTemplate?: Word.Interfaces.ListTemplateLoadOptions;
```

Property Value  
[Word.Interfaces.ListTemplateLoadOptions](/en-us/javascript/api/word/word.interfaces.listtemplateloadoptions)

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### listType
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the type of the list for the `ListFormat` object.

```typescript
listType?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]

### listValue
Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the numeric value of the the first paragraph in the range for the `ListFormat` object.

```typescript
listValue?: boolean;
```

Property Value  
boolean

Remarks  
[ API set: WordApi BETA (PREVIEW ONLY) ]