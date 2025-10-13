# Word.Interfaces.AnnotationCollectionLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Contains a collection of [Word.Annotation](/en-us/javascript/api/word/word.annotation) objects.

## Remarks

[ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

## Properties

- [$all](#word-word-interfaces-annotationcollectionloadoptions-all-member): Specifying $all for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- [critiqueAnnotation](#word-word-interfaces-annotationcollectionloadoptions-critiqueannotation-member): For EACH ITEM in the collection: Gets the critique annotation object.
- [id](#word-word-interfaces-annotationcollectionloadoptions-id-member): For EACH ITEM in the collection: Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.
- [state](#word-word-interfaces-annotationcollectionloadoptions-state-member): For EACH ITEM in the collection: Gets the state of the annotation.

## Property Details

<a id="word-word-interfaces-annotationcollectionloadoptions-all-member"></a>
### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property Value: boolean

<a id="word-word-interfaces-annotationcollectionloadoptions-critiqueannotation-member"></a>
### critiqueAnnotation

For EACH ITEM in the collection: Gets the critique annotation object.

```typescript
critiqueAnnotation?: Word.Interfaces.CritiqueAnnotationLoadOptions;
```

Property Value: [Word.Interfaces.CritiqueAnnotationLoadOptions](/en-us/javascript/api/word/word.interfaces.critiqueannotationloadoptions)

Remarks: [ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

<a id="word-word-interfaces-annotationcollectionloadoptions-id-member"></a>
### id

For EACH ITEM in the collection: Gets the unique identifier, which is meant to be used for easier tracking of Annotation objects.

```typescript
id?: boolean;
```

Property Value: boolean

Remarks: [ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]

<a id="word-word-interfaces-annotationcollectionloadoptions-state-member"></a>
### state

For EACH ITEM in the collection: Gets the state of the annotation.

```typescript
state?: boolean;
```

Property Value: boolean

Remarks: [ [API set: WordApi 1.7](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets) ]