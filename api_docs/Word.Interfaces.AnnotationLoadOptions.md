# Word.Interfaces.AnnotationLoadOptions interface

Package: https://learn.microsoft.com/en-us/javascript/api/word

Represents an annotation attached to a paragraph.

## Remarks
[API set: WordApi 1.7](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- $all  
  Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

- critiqueAnnotation  
  Gets the critique annotation object.

- id  
  Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.

- state  
  Gets the state of the annotation.

## Property Details

### $all
Specifying $all for the load options loads all the scalar properties (such as Range.address) but not the navigational properties (such as Range.format.fill.color).

TypeScript:
```
$all?: boolean;
```

Property Value: boolean

---

### critiqueAnnotation
Gets the critique annotation object.

TypeScript:
```
critiqueAnnotation?: Word.Interfaces.CritiqueAnnotationLoadOptions;
```

Property Value: https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.critiqueannotationloadoptions

Remarks: [API set: WordApi 1.7](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id
Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.

TypeScript:
```
id?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.7](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### state
Gets the state of the annotation.

TypeScript:
```
state?: boolean;
```

Property Value: boolean

Remarks: [API set: WordApi 1.7](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)