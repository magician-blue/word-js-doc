# Word.Interfaces.CustomXmlPartScopedCollectionLoadOptions interface

Package: [word](https://learn.microsoft.com/en-us/javascript/api/word)

Contains the collection of [Word.CustomXmlPart](https://learn.microsoft.com/en-us/javascript/api/word/word.customxmlpart) objects with a specific namespace.

## Remarks
[API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties
- $all: Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- builtIn: For EACH ITEM in the collection: Gets a value that indicates whether the `CustomXmlPart` is built-in.
- documentElement: For EACH ITEM in the collection: Gets the root element of a bound region of data in the document. If the region is empty, the property returns `Nothing`.
- id: For EACH ITEM in the collection: Gets the ID of the custom XML part.
- namespaceUri: For EACH ITEM in the collection: Gets the namespace URI of the custom XML part.
- xml: For EACH ITEM in the collection: Gets the XML representation of the current `CustomXmlPart` object.

## Property Details

### $all
Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

Property value: boolean

---

### builtIn
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets a value that indicates whether the `CustomXmlPart` is built-in.

```typescript
builtIn?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### documentElement
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the root element of a bound region of data in the document. If the region is empty, the property returns `Nothing`.

```typescript
documentElement?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

Property value: [Word.Interfaces.CustomXmlNodeLoadOptions](https://learn.microsoft.com/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id
For EACH ITEM in the collection: Gets the ID of the custom XML part.

```typescript
id?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### namespaceUri
For EACH ITEM in the collection: Gets the namespace URI of the custom XML part.

```typescript
namespaceUri?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi 1.4](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### xml
Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

For EACH ITEM in the collection: Gets the XML representation of the current `CustomXmlPart` object.

```typescript
xml?: boolean;
```

Property value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)