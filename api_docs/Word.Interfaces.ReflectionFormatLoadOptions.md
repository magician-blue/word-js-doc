# Word.Interfaces.ReflectionFormatLoadOptions interface

Package: word

Note
This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Represents the reflection formatting for a shape in Word.

## Remarks

[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- $all — Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).
- blur — Specifies the degree of blur effect applied to the ReflectionFormat object as a value between 0.0 and 100.0.
- offset — Specifies the amount of separation, in points, of the reflected image from the shape.
- size — Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.
- transparency — Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).
- type — Specifies a ReflectionType value that represents the type and direction of the lighting for a shape reflection.

## Property Details

### $all

Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value
- boolean

### blur

Specifies the degree of blur effect applied to the `ReflectionFormat` object as a value between 0.0 and 100.0.

```typescript
blur?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### offset

Specifies the amount of separation, in points, of the reflected image from the shape.

```typescript
offset?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### size

Specifies the size of the reflection as a percentage of the reflected shape from 0 to 100.

```typescript
size?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### transparency

Specifies the degree of transparency for the reflection effect as a value between 0.0 (opaque) and 1.0 (clear).

```typescript
transparency?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### type

Specifies a `ReflectionType` value that represents the type and direction of the lighting for a shape reflection.

```typescript
type?: boolean;
```

Property Value
- boolean

Remarks
[API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)