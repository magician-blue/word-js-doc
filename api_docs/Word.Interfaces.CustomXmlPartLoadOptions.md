# Word.Interfaces.CustomXmlPartLoadOptions interface

Package: [word](/en-us/javascript/api/word)

Represents a custom XML part.

## Remarks

[API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

## Properties

- [$all](#all)
  - Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).
- [builtIn](#builtin)
  - Gets a value that indicates whether the `CustomXmlPart` is built-in.
- [documentElement](#documentelement)
  - Gets the root element of a bound region of data in the document. If the region is empty, the property returns `Nothing`.
- [id](#id)
  - Gets the ID of the custom XML part.
- [namespaceUri](#namespaceuri)
  - Gets the namespace URI of the custom XML part.
- [xml](#xml)
  - Gets the XML representation of the current `CustomXmlPart` object.

## Property Details

### $all

Specifying `$all` for the load options loads all the scalar properties (such as `Range.address`) but not the navigational properties (such as `Range.format.fill.color`).

```typescript
$all?: boolean;
```

- Property Value: boolean

---

### builtIn

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets a value that indicates whether the `CustomXmlPart` is built-in.

```typescript
builtIn?: boolean;
```

- Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### documentElement

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the root element of a bound region of data in the document. If the region is empty, the property returns `Nothing`.

```typescript
documentElement?: Word.Interfaces.CustomXmlNodeLoadOptions;
```

- Property Value: [Word.Interfaces.CustomXmlNodeLoadOptions](/en-us/javascript/api/word/word.interfaces.customxmlnodeloadoptions)

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### id

Gets the ID of the custom XML part.

```typescript
id?: boolean;
```

- Property Value: boolean

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### namespaceUri

Gets the namespace URI of the custom XML part.

```typescript
namespaceUri?: boolean;
```

- Property Value: boolean

Remarks: [API set: WordApi 1.4](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)

---

### xml

Note: This API is provided as a preview for developers and may change based on feedback that we receive. Do not use this API in a production environment.

Gets the XML representation of the current `CustomXmlPart` object.

```typescript
xml?: boolean;
```

- Property Value: boolean

Remarks: [API set: WordApi BETA (PREVIEW ONLY)](/en-us/javascript/api/requirement-sets/word/word-api-requirement-sets)