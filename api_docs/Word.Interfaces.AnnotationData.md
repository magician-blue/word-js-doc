# Word.Interfaces.AnnotationData interface

Package: [word](/en-us/javascript/api/word)

An interface describing the data returned by calling `annotation.toJSON()`.

## Properties

- [id](#id) — Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.
- [state](#state) — Gets the state of the annotation.

## Property Details

### id

Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.

```typescript
id?: string;
```

#### Property Value

string

#### Remarks

[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

### state

Gets the state of the annotation.

```typescript
state?: Word.AnnotationState | "Created" | "Accepted" | "Rejected";
```

#### Property Value

[Word.AnnotationState](/en-us/javascript/api/word/word.annotationstate) | "Created" | "Accepted" | "Rejected"

#### Remarks

[API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)