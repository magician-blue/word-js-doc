# Word.Interfaces.ReflectionFormatData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling reflectionFormat.toJSON().

## Properties

- blur
  - Specifies the degree of blur effect applied to the ReflectionFormat object as a value between 0.0 and 100.0.
- offset
  - Specifies the amount of separation, in points, of the reflected image from the shape.
- size
  - Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.
- transparency
  - Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).
- type
  - Specifies a ReflectionType value that represents the type and direction of the lighting for a shape reflection.

## Property Details

### blur

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of blur effect applied to the ReflectionFormat object as a value between 0.0 and 100.0.

```typescript
blur?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### offset

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the amount of separation, in points, of the reflected image from the shape.

```typescript
offset?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### size

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.

```typescript
size?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: number;
```

Property Value: number

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Specifies a ReflectionType value that represents the type and direction of the lighting for a shape reflection.

```typescript
type?: Word.ReflectionType | "Mixed" | "None" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9";
```

Property Value: [Word.ReflectionType](/en-us/javascript/api/word/word.reflectiontype) | "Mixed" | "None" | "Type1" | "Type2" | "Type3" | "Type4" | "Type5" | "Type6" | "Type7" | "Type8" | "Type9"

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)